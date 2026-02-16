from pathlib import Path
import json
import pytesseract
from pdf2image import convert_from_path
import pytesseract

pytesseract.pytesseract.tesseract_cmd = "/opt/homebrew/bin/tesseract"

import re

def normalize_line(line: str) -> str:
    # normalize whitespace and common OCR artifacts
    line = line.replace("’", "'").replace("“", '"').replace("”", '"')
    line = re.sub(r"\s+", " ", line)
    return line.strip()


def is_junk_line(line: str) -> bool:
    if not line:
        return True

    # page numbers
    if re.fullmatch(r"\d+", line):
        return True

    # very short garbage
    if len(line) < 3:
        return True

    return False

def is_chord_line(line: str) -> bool:
    """
    Detect lines that are mostly chord symbols:
    G   D/F#   Em7   Cadd9
    """
    tokens = line.split()
    if not tokens:
        return False

    chord_pattern = re.compile(r"^[A-G][#b]?(m|maj|min|sus|dim|aug)?\d*(/[A-G][#b]?)?$")

    matches = sum(1 for t in tokens if chord_pattern.match(t))
    return matches / len(tokens) > 0.6

def is_metadata_line(line: str) -> bool:
    lower = line.lower()

    # common headers / metadata
    if lower.startswith(("key:", "tuning:", "tempo:", "capo:")):
        return True

    if "chords by" in lower:
        return True

    if lower in {"chords", "intro:", "outro:"}:
        return True

    # page indicators
    if re.search(r"page\s+\d+/\d+", lower):
        return True

    # instructions
    if lower in {"repeat chorus", "repeat bridge"}:
        return True

    return False

def is_mostly_non_lyric(line: str) -> bool:
    """
    Remove lines that contain very few actual words.
    """
    letters = sum(c.isalpha() for c in line)
    return letters < 5

def is_symbol_heavy(line: str) -> bool:
    symbols = sum(not c.isalnum() and not c.isspace() for c in line)
    return symbols > len(line) * 0.3

def extract_text_via_ocr(pdf_path: Path) -> list[str]:
    lines = []

    images = convert_from_path(
        pdf_path,
        dpi=300,
        poppler_path="/opt/homebrew/bin"
    )

    for page_num, image in enumerate(images, start=1):
        text = pytesseract.image_to_string(image)

        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue
            lines.append(line)

    return lines

def clean_ocr_lines(lines):
    cleaned = []

    for line in lines:
        line = normalize_line(line)

        if is_junk_line(line):
            continue

        if is_metadata_line(line):
            continue

        if is_chord_line(line):
            continue

        if is_symbol_heavy(line):
            continue

        if is_mostly_non_lyric(line):
            continue

        cleaned.append(line)

    return cleaned

def group_lines_into_sections(lines):
    sections = []
    current = None
    counters = {}

    header_pattern = re.compile(r"^(verse\s*\d+|chorus|bridge|outro|intro)\s*$", re.IGNORECASE)

    def start_section(label: str):
        nonlocal current

        label = label.strip()
        label_lower = label.lower()

        # Normalize plain headers like "VERSE 1" -> "Verse 1"
        label = " ".join(w.capitalize() for w in label_lower.split())

        if label_lower.startswith("verse"):
            match = re.search(r"\d+", label_lower)
            number = match.group() if match else "1"
            section_id = f"verse{number}"
            sec_type = "verse"
        elif "chorus" in label_lower:
            counters.setdefault("chorus", 0)
            counters["chorus"] += 1
            section_id = f"chorus{counters['chorus']}"
            sec_type = "chorus"
            label = "Chorus"
        elif "bridge" in label_lower:
            counters.setdefault("bridge", 0)
            counters["bridge"] += 1
            section_id = f"bridge{counters['bridge']}"
            sec_type = "bridge"
            label = "Bridge"
        else:
            base = re.sub(r"[^a-z0-9]+", "", label_lower)
            counters.setdefault(base, 0)
            counters[base] += 1
            section_id = f"{base}{counters[base]}"
            sec_type = base or "other"

        current = {
            "id": section_id,
            "label": label,
            "type": sec_type,
            "lines": []
        }

    for line in lines:
        # Header style 1: [Verse 1]
        if line.startswith("[") and line.endswith("]"):
            if current:
                sections.append(current)
            start_section(line.strip("[]"))
            continue

        # Header style 2: VERSE 1 / CHORUS
        if header_pattern.match(line.strip()):
            if current:
                sections.append(current)
            start_section(line.strip())
            continue

        # Normal lyric line
        if current:
            current["lines"].append(line)

    if current:
        sections.append(current)

    return sections

def chunk_lines(lines, size=4):
    return [lines[i:i+size] for i in range(0, len(lines), size)]

def build_song_json(title: str, sections, output_path: Path):
    json_sections = []

    for sec in sections:
        json_sections.append({
            "id": sec["id"],
            "label": sec["label"],
            "type": sec["type"],
            # New format: store raw lines exactly as they appear in the PDF/OCR.
            # SlideBuilder decides how to split them into slides.
            "lines": sec.get("lines", [])
        })

    song = {
        "schema_version": "1.0",
        "song": {
            "title": title,
            "author": "",
            "copyright": "",
            "ccli_number": "",
            "notes": ""
        },
        "structure": {
            "sections": json_sections
        },
        "chords": {
            "enabled": False,
            "sections": {}
        }
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(song, f, indent=2, ensure_ascii=False)

def import_song_from_pdf(pdf_path: Path, songs_folder: Path) -> Path:
    raw_lines = extract_text_via_ocr(pdf_path)
    cleaned = clean_ocr_lines(raw_lines)
    sections = group_lines_into_sections(cleaned)

    # Fallback: if no headers were detected, put everything into one section
    if not sections and cleaned:
        sections = [{
            "id": "lyrics",
            "label": "Lyrics",
            "type": "verse",
            "lines": cleaned
        }]

    title = pdf_path.stem.replace("_", " ").title()
    output_path = songs_folder / f"{pdf_path.stem.lower()}.json"

    build_song_json(title, sections, output_path)
    return output_path


if __name__ == "__main__":
    pdf = Path(
        "/Users/george/Documents/Spiritual/Church/WorshipSlides/SongPDFs/Come_Jesus_Come.pdf"
    )

    raw_lines = extract_text_via_ocr(pdf)
    print(f"RAW OCR LINES: {len(raw_lines)}")

    cleaned = clean_ocr_lines(raw_lines)
    print(f"CLEANED LINES: {len(cleaned)}")

    sections = group_lines_into_sections(cleaned)

    output = Path(
        "/Users/george/Documents/Spiritual/Church/WorshipSlides/songs/come_jesus_come.json"
    )

    build_song_json("Come Jesus Come", sections, output)
    print(f"Song JSON created: {output}")