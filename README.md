# Word Formatter

A CLI tool that applies consistent formatting to Word documents based on a configurable rules file.

## What it does

- Classifies each paragraph/table in a `.docx` (headings, lists, body text) using style info and heuristics
- Applies font, size, bold/italic/underline, heading, table, header/footer, and margin rules from `rules.toml`
- Auto-numbers headings and formats memo headers (To/From/Date/Re)
- Backs up the original file before writing output to an `Output/` folder next to it

## Usage

```bash
python main.py <path_to_docx>
```

## Configuration

Edit `rules.toml` to control formatting (fonts, colors, table styling, margins, header/footer text, etc.).

## Tests

Test scripts live in `Tests/` and rely on local sample `.docx` fixtures (not included in this repo).
