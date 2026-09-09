import sys
from pathlib import Path
from formatter import format_document

def main():
    """
    CLI entry point for the Word Formatter.

    Usage:
        python main.py <path_to_docx>
    """
    if len(sys.argv) < 2:
        print("Error: No file path provided.")
        print("Usage: python main.py <path_to_docx>")
        sys.exit(1)

    input_path = sys.argv[1]
    path = Path(input_path)

    if not path.exists():
        print(f"Error: File not found: {input_path}")
        sys.exit(1)

    if path.suffix.lower() != ".docx":
        print(f"Error: File must be a .docx file: {input_path}")
        sys.exit(1)

    format_document(str(path))

if __name__ == "__main__":
    main()
