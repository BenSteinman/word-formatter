from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.text.run import Run
import re

# Paragraph Classifier

def classify_paragraph(paragraph : Paragraph) -> dict:
    """
    Classifies a paragraph and returns a metadata dictionary.

    Args:
        paragraph (Paragraph): A python-docx Paragraph object.

    Returns:
        dict: A dictionary containing:
            - "style" (str): The paragraph's style name, or "No Style" if none is assigned.
            - "full_text" (str): The raw text content of the paragraph.
            - "is_blank" (bool): True if the paragraph contains no visible text.
            - "is_list" (bool): True if the paragraph is part of a list.
            - "list_info" (dict): Contains "list_id" and "list_level", both None if not a list.
            - "heading_level" (int | None): Final resolved heading level (e.g. 1, 2), or None
                                            if not a heading or level could not be determined.
                                            Pass 2 in the formatter resolves None for heuristic headings.
            - "heading_source" (str | None): How the heading was detected — "style", "heuristic", or None.
            - "heuristic_score" (int): Raw score from score_heuristics(), always present.
            - "runs" (list): List of run metadata dicts for each run in the paragraph.
    """
    text = paragraph.text
    style = paragraph.style.name if paragraph.style else "No Style"
    
    heuristic_score = score_heuristics(paragraph)
    style_level = detect_header(paragraph)

    if style_level is not None:
        heading_level = style_level
        heading_source = "style"
    elif heuristic_score >= 4:
        heading_level = extract_section_level(text) or 1
        heading_source = "heuristic"
    else:
        heading_level = None
        heading_source = None

    meta = {
        "style": style,
        "full_text": text,
        "is_blank": is_blank_text(text),
        "is_list": is_list(paragraph),
        "list_info": get_list_info(paragraph),
        "heading_level": heading_level,
        "heading_source": heading_source,
        "heuristic_score": heuristic_score,
        "runs": classify_runs(paragraph)
    }

    return meta

def normalize_text(text : str) -> str:
    """
    Strips leading/trailing whitespace and removes non-breaking spaces.

    Args:
        text (str): Raw text from a paragraph.

    Returns:
        str: Cleaned text string, or empty string if input is falsy.
    """
    if not text:
        return ""
    return text.replace("\u00A0", "").strip()

def is_blank_text(text : str) -> bool:
    """
    Determines whether a paragraph's text is blank.

    Args:
        text (str): Raw text from a paragraph.

    Returns:
        bool: True if the text is empty or contains only whitespace/non-breaking spaces.
    """
    return normalize_text(text) == ""

def detect_header(paragraph : Paragraph) -> int | None:
    """
    Detects whether a paragraph is a heading and returns its level.

    Args:
        paragraph (Paragraph): A python-docx Paragraph object.

    Returns:
        int | None: The heading level as an integer (e.g. 1 for Heading 1),
                    or None if the paragraph is not a heading.
    """
    style = paragraph.style.name if paragraph.style else ""

    if style.startswith("Heading"):
        try:
            return int(style.replace("Heading","").strip())
        except:
            return None
        
    return None

def extract_section_level(text: str) -> int | None:
    """
    Infers a heading level from a section number prefix in the text.

    Counts the number of dot-separated numeric segments at the start
    of the string to determine depth.

    Examples:
        "4. Introduction"   -> 1
        "4.1 Background"    -> 2
        "4.1.2 Details"     -> 3
        "No number here"    -> None

    Args:
        text (str): The raw text of a paragraph.

    Returns:
        int | None: The inferred heading level, or None if no section
                    number prefix is found.
    """
    match = re.match(r'^\d+(\.\d+)*\.?\s.+', text.strip())
    if not match:
        return None
    numeric_prefix = match.group(0).split()[0]   # grab "4.1.2" or "4.1.2."
    numeric_prefix = numeric_prefix.rstrip('.')  # strip trailing dot -> "4.1.2"
    segments = numeric_prefix.split('.')         # ["4", "1", "2"]
    return len(segments)                         # 3

def score_heuristics(paragraph: Paragraph) -> int:
    """
    Scores a paragraph against heuristic signals that suggest it functions
    as a heading, even if it does not use a heading style.

    Args:
        paragraph (Paragraph): A python-docx Paragraph object.

    Returns:
        int: A score representing heading likelihood. Higher = more likely a heading.
    """
    score = 0
    text = paragraph.text

    # Short text (< 15 words) -> +1
    word_count = len(text.split())
    if word_count < 15:
        score += 1

    # Single run -> +1
    run_count = len(paragraph.runs)
    if run_count == 1:
        score += 1

    # No trailing punctuation -> +1
    stripped_text = text.strip()
    if stripped_text and stripped_text[-1] not in '.?!':
        score += 1

    # Bold -> +2
    runs = paragraph.runs
    if runs:
        is_bold = all(run.bold for run in runs)
        if is_bold:
            score += 2

    # ALL CAPS or Title Case -> +1
    expression = r'^([^a-z]+|[A-Z][a-z]*(\s[A-Z][a-z]*)*)$'
    if re.match(expression, text.strip()):
        score += 1

    # Starts with section number -> +2
    if extract_section_level(text) is not None:
        score += 2


    return score

def is_list(paragraph : Paragraph) -> bool:
    """
    Determines whether a paragraph is part of a list.

    Checks both the XML numbering properties (numPr) and the paragraph's
    style name as a fallback.

    Args:
        paragraph (Paragraph): A python-docx Paragraph object.

    Returns:
        bool: True if the paragraph is a list item, False otherwise.
    """
    p = paragraph._p

    if p.pPr is not None and p.pPr.numPr is not None:
        return True
    
    style = paragraph.style.name if paragraph.style else "No Style"
    return "List" in style

def get_list_info(paragraph : Paragraph) -> dict:
    """
    Retrieves numbering metadata for a list paragraph.

    Args:
        paragraph (Paragraph): A python-docx Paragraph object.

    Returns:
        dict: A dictionary containing:
            - "list_id" (int | None): The numbering definition ID, or None if not a list.
            - "list_level" (int | None): The indentation level of the list item (0-based),
                                         or None if not a list.
    """
    
    if not is_list(paragraph):
        return {
            "list_id" : None,
            "list_level" : None
        }
    
    p = paragraph._p
    numPr = p.pPr.numPr

    if numPr is None or numPr.numId is None or numPr.ilvl is None:
        return {
            "list_id" : None,
            "list_level" : None
        }

    return {
            "list_id" : numPr.numId.val,
            "list_level" : numPr.ilvl.val
    }

def classify_run(run: Run) -> dict:
    """
    Classifies a single run and returns its formatting metadata.

    Args:
        run (Run): A python-docx Run object.

    Returns:
        dict: A dictionary containing:
            - "text" (str): The run's text content.
            - "bold" (bool | None): Bold state, or None if inherited from style.
            - "italic" (bool | None): Italic state, or None if inherited from style.
            - "underline" (bool | None): Underline state, or None if inherited from style.
            - "style" (str | None): The run's style name, or None if not set.
            - "font" (str | None): The font name, or None if not set.
            - "size" (float | None): The font size in points, or None if not set.
    """
    meta = {
        "text": run.text,
        "bold": run.bold,
        "italic": run.italic,
        "underline": run.underline,
        "style": run.style.name if run.style else None,
        "font": run.font.name,
        "size": run.font.size.pt if run.font.size else None
    }

    return meta

def classify_runs(paragraph: Paragraph) -> list:
    """
    Classifies all runs in a paragraph.

    Args:
        paragraph (Paragraph): A python-docx Paragraph object.

    Returns:
        list: A list of run metadata dicts, one per run.
    """
    returnList = []
    
    for run in paragraph.runs:
        returnList.append(classify_run(run))

    return returnList

# Table Classifier

def classify_table(table: Table) -> dict:
    """
    Classifies a table and returns its structural metadata.

    Args:
        table (Table): A python-docx Table object.

    Returns:
        dict: A dictionary containing:
            - "rows" (int): Number of rows in the table.
            - "cols" (int): Number of columns in the table.
            - "cell_count" (int): Number of unique cells (accounts for merged cells).
    """
    rows = len(table.rows)
    cols = len(table.columns)
    
    seen = set()
    for row in table.rows:
        for cell in row.cells:
            seen.add(cell._tc)

    meta = {
        "rows" : rows,
        "cols" : cols,
        "cell_count" : len(seen)
    }

    return meta