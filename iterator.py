from docx.text.paragraph import Paragraph
from docx.table import Table

def iter_blocks(doc):
    """
    Iterates through the top-level blocks of a Word document in order.

    Parses the raw XML body of the document and yields each block as a
    dictionary containing its type and corresponding python-docx object.

    Args:
        doc: A python-docx Document object.

    Yields:
        dict: A dictionary containing:
            - "type" (str): The block type, either "paragraph" or "table".
            - "object": The corresponding python-docx object
                        (Paragraph or Table).
    """
    for child in doc.element.body:
        tag = child.tag.split("}")[-1]

        if tag == "p":
            yield {
                "type": "paragraph",
                "object": Paragraph(child, doc)
        }

        elif tag == "tbl":
            yield {
                "type": "table",
                "object": Table(child, doc)
            }
