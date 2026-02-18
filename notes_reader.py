import subprocess
from pathlib import Path


def read_notes_text(notes_path: Path) -> str:
    """
    Returns the full text of a notes file.

    Supports:
      - .txt
      - .docx
      - .pages (macOS only; uses AppleScript + Pages)
    """
    suffix = notes_path.suffix.lower()

    if suffix == ".txt":
        return notes_path.read_text(encoding="utf-8", errors="ignore")

    if suffix == ".docx":
        from docx import Document
        doc = Document(notes_path)
        return "\n".join(p.text for p in doc.paragraphs)

    if suffix == ".pages":
        # macOS only: use AppleScript to open in Pages and extract text
        applescript = f'''
        tell application "Pages"
            set docRef to open POSIX file "{notes_path}"
            set paraList to every paragraph of body text of docRef
            set AppleScript's text item delimiters to linefeed
            set outText to paraList as text
            close docRef saving no
            quit
        end tell
        return outText
        '''
        completed = subprocess.run(
            ["osascript", "-e", applescript],
            capture_output=True, text=True, check=True
        )
        return completed.stdout

    raise ValueError("Unsupported notes file. Use .pages, .docx, or .txt")