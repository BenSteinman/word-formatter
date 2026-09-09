import shutil
from pathlib import Path
from datetime import datetime
import tomllib

def backup_document(input_path: str) -> str:
    """
    Creates a timestamped backup of a document in a Backups/ subfolder
    next to the original file.

    Args:
        input_path (str): Path to the original document.

    Returns:
        str: Path to the backup file.
    """

    original = Path(input_path)

    if not original.exists():
        raise FileNotFoundError(
            f"Document not found: {input_path}"
        )

    backup_dir = original.parent / "Backups"
    backup_dir.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    backup_name = (
        f"{original.stem}_{timestamp}{original.suffix}"
    )

    backup_path = backup_dir / backup_name

    shutil.copy2(original, backup_path)

    print(f"Backup saved to: {backup_path}")

    return str(backup_path)

def load_rules() -> dict:
    """
    Loads formatting rules from rules.toml.

    Returns:
        dict: Parsed TOML configuration.
    """

    rules_path = Path(__file__).parent / "rules.toml"

    if not rules_path.exists():
        raise FileNotFoundError(
            f"Rules file not found: {rules_path}"
        )

    with open(rules_path, "rb") as f:
        return tomllib.load(f)