import pdfplumber
from pathlib import Path
import json


def extract_lyrics_from_pdf(pdf_path: Path) -> list[str]:
    lines = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            print(f"\n--- PAGE {page_num} ---")

            words = page.extract_words(use_text_flow=True)
            print(f"Words found: {len(words)}")

            for w in words:
                text = w["text"].strip()
                if text:
                    lines.append(text)

    return lines


def looks_like_chords(line: str) -> bool:
    """
    Very simple heuristic:
    If a line is mostly chord symbols, skip it.
    """
    chord_chars = set("ABCDEFGabcdefg#bm7/() ")
    letters = [c for c in line if c.isalpha()]

    if not letters:
        return True

    ratio = sum(1 for c in line if c in chord_chars) / len(line)
    return ratio > 0.8


def build_song_json(title: str, lyric_lines: list[str], output_path: Path):
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
            "sections": [
                {
                    "id": "lyrics",
                    "label": "Lyrics",
                    "type": "verse",
                    "slides": [
                        {
                            "lines": lyric_lines
                        }
                    ]
                }
            ]
        },
        "chords": {
            "enabled": False,
            "sections": {}
        }
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(song, f, indent=2, ensure_ascii=False)


if __name__ == "__main__":
    pdf = Path("/Users/george/Documents/Spiritual/Church/WorshipSlides/SongPDFs/Come_Jesus_Come.pdf")
    output = Path("songs/imported_song.json")

    lyrics = extract_lyrics_from_pdf(pdf)
    build_song_json("Imported Song", lyrics, output)

    print(f"Imported {len(lyrics)} lyric lines")