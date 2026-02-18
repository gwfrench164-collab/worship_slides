import json
import os
import re
from pathlib import Path

# If you copied kjv.json into the repo root (same folder as this file),
# this will find it automatically:
DEFAULT_KJV_PATH = Path(__file__).resolve().parent / "kjv.json"

ALLOW_CHAPTER_ONLY = False

BOOK_NAMES = [
 "Genesis","Exodus","Leviticus","Numbers","Deuteronomy","Joshua","Judges","Ruth",
 "1 Samuel","2 Samuel","1 Kings","2 Kings","1 Chronicles","2 Chronicles",
 "Ezra","Nehemiah","Esther","Job","Psalms","Proverbs","Ecclesiastes","Song of Solomon",
 "Isaiah","Jeremiah","Lamentations","Ezekiel","Daniel","Hosea","Joel","Amos","Obadiah",
 "Jonah","Micah","Nahum","Habakkuk","Zephaniah","Haggai","Zechariah","Malachi",
 "Matthew","Mark","Luke","John","Acts","Romans","1 Corinthians","2 Corinthians",
 "Galatians","Ephesians","Philippians","Colossians","1 Thessalonians","2 Thessalonians",
 "1 Timothy","2 Timothy","Titus","Philemon","Hebrews","James","1 Peter","2 Peter",
 "1 John","2 John","3 John","Jude","Revelation",
 # Optional common abbreviations to make regex matching more forgiving
 "Mt","Matt","Mk","Mrk","Lk","Lu","Jn","Jo","Ac","Rom","Gal","Eph","Phil","Php","Col","Heb","Jas","Jam","Rev","Re"
]

BOOK_RE = "(" + "|".join(re.escape(b) for b in BOOK_NAMES) + ")"
REF_PATTERN = re.compile(
    rf"\b{BOOK_RE}\s*(\d{{1,3}})(?::\s*(\d{{1,3}}(?:[-,]\d{{1,3}})*(?:,\s*\d{{1,3}}(?:[-,]\d{{1,3}})*)?))?",
    re.IGNORECASE
)

def clean_text(txt: str) -> str:
    if not txt:
        return ""
    txt = txt.replace("¶", "").replace("‹", "").replace("›", "")
    txt = txt.replace("<", "").replace(">", "")
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()

def preprocess_text_for_refs(text: str) -> str:
    roman_map = {
        r"\bI\s+Samuel": "1 Samuel",
        r"\bII\s+Samuel": "2 Samuel",
        r"\bI\s+Kings?": "1 Kings",
        r"\bII\s+Kings?": "2 Kings",
        r"\bI\s+Chron": "1 Chronicles",
        r"\bII\s+Chron": "2 Chronicles",
        r"\bI\s+Cor": "1 Corinthians",
        r"\bII\s+Cor": "2 Corinthians",
        r"\bI\s+Thess": "1 Thessalonians",
        r"\bII\s+Thess": "2 Thessalonians",
        r"\bI\s+Tim": "1 Timothy",
        r"\bII\s+Tim": "2 Timothy",
        r"\bI\s+Pet": "1 Peter",
        r"\bII\s+Pet": "2 Peter",
        r"\bI\s+Jn": "1 John",
        r"\bII\s+Jn": "2 John",
        r"\bIII\s+Jn": "3 John",
    }
    for pattern, replacement in roman_map.items():
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)

    text = re.sub(r"\b11\s+Cor", "2 Corinthians", text, flags=re.IGNORECASE)

    abbrev_map = {
        r"\bGen\b": "Genesis",
        r"\bExo\b": "Exodus", r"\bEx\b": "Exodus",
        r"\bLev\b": "Leviticus",
        r"\bNum\b": "Numbers", r"\bNu\b": "Numbers",
        r"\bDeut\b": "Deuteronomy", r"\bDt\b": "Deuteronomy",
        r"\bJudg\b": "Judges", r"\bJdg\b": "Judges",
        r"\bPs\b": "Psalms", r"\bPsa\b": "Psalms",
        r"\bProv\b": "Proverbs", r"\bPr\b": "Proverbs",
        r"\bEccl\b": "Ecclesiastes", r"\bEcc\b": "Ecclesiastes",
        r"\bSong\b": "Song of Solomon",
        r"\bIsa\b": "Isaiah",
        r"\bJer\b": "Jeremiah",
        r"\bLam\b": "Lamentations",
        r"\bEzek\b": "Ezekiel", r"\bEze\b": "Ezekiel",
        r"\bDan\b": "Daniel",
        r"\bHos\b": "Hosea",
        r"\bObad\b": "Obadiah", r"\bOb\b": "Obadiah",
        r"\bJon\b": "Jonah",
        r"\bMic\b": "Micah",
        r"\bHab\b": "Habakkuk",
        r"\bZeph\b": "Zephaniah", r"\bZep\b": "Zephaniah",
        r"\bZech\b": "Zechariah", r"\bZec\b": "Zechariah",
        r"\bMal\b": "Malachi",
        r"\bMat\b": "Matthew", r"\bMatt\b": "Matthew",
        r"\bMk\b": "Mark", r"\bMrk\b": "Mark",
        r"\bLk\b": "Luke", r"\bLu\b": "Luke",
        r"\bJn\b": "John",
        r"\bAc\b": "Acts",
        r"\bRom\b": "Romans",
        r"\bGal\b": "Galatians",
        r"\bEph\b": "Ephesians",
        r"\bPhil\b": "Philippians", r"\bPhp\b": "Philippians",
        r"\bCol\b": "Colossians",
        r"\bHeb\b": "Hebrews",
        r"\bJas\b": "James", r"\bJam\b": "James",
        r"\bRev\b": "Revelation", r"\bRe\b": "Revelation",
    }
    for pattern, replacement in abbrev_map.items():
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)

    return text

def normalize_book_name(book: str):
    b = re.sub(r"\s+", " ", book).strip().lower()
    aliases = {
        "ps": "psalms", "psa": "psalms", "psalm": "psalms",
        "song": "song of solomon",
        "jn": "john", "jo": "john",
        "mt": "matthew", "mat": "matthew", "matt": "matthew",
        "mk": "mark", "mrk": "mark",
        "lk": "luke", "lu": "luke",
        "rom": "romans",
        "1 cor": "1 corinthians", "2 cor": "2 corinthians",
        "1 thess": "1 thessalonians", "2 thess": "2 thessalonians",
        "1 tim": "1 timothy", "2 tim": "2 timothy",
        "1 pet": "1 peter", "2 pet": "2 peter",
        "1 jn": "1 john", "2 jn": "2 john", "3 jn": "3 john",
        "rev": "revelation", "re": "revelation",
    }
    return aliases.get(b, b)

def extract_ordered_refs(doc_text: str) -> list[str]:
    doc_text = preprocess_text_for_refs(doc_text)

    ordered_refs = []
    seen = set()

    for m in REF_PATTERN.finditer(doc_text):
        book = m.group(1)
        chapter = m.group(2)
        verses = m.group(3)
        refstr = f"{book} {chapter}:{verses}" if verses else f"{book} {chapter}"

        key = refstr.lower()
        if key not in seen:
            seen.add(key)
            ordered_refs.append(refstr)

    return ordered_refs

def load_bible_json(json_path: str):
    if not os.path.exists(json_path):
        raise FileNotFoundError(f"Bible JSON not found: {json_path}")

    with open(json_path, encoding="utf-8") as fh:
        data = json.load(fh)

    entries = data if isinstance(data, list) else data.get("verses", [])
    bible = {}

    for entry in entries:
        if not isinstance(entry, dict):
            continue
        book = str(entry.get("book_name", "")).strip()
        if not book:
            continue
        try:
            chapter = int(entry.get("chapter"))
            verse = int(entry.get("verse"))
        except (TypeError, ValueError):
            continue

        text = clean_text(str(entry.get("text", "")).strip())
        key = (normalize_book_name(book), chapter, verse)
        bible[key] = text

    return bible

def fetch_verse_text(ref_str: str, bible_dict):
    m = re.match(rf"^\s*{BOOK_RE}\s+(\d{{1,3}})(?::(.+))?\s*$", ref_str, re.IGNORECASE)
    if not m:
        return ""

    book_raw = m.group(1)
    chapter = int(m.group(2))
    verses_part = m.group(3)
    book = normalize_book_name(book_raw)

    verses_texts = []

    if not verses_part:
        if not ALLOW_CHAPTER_ONLY:
            return ""
        v = 1
        while True:
            key = (book, chapter, v)
            if key in bible_dict:
                verses_texts.append(bible_dict[key])
                v += 1
            else:
                break
    else:
        for piece in re.split(r"\s*,\s*", verses_part):
            if "-" in piece:
                bounds = piece.split("-")
                if len(bounds) == 2:
                    try:
                        start_v = int(bounds[0]); end_v = int(bounds[1])
                    except ValueError:
                        continue
                    for vv in range(start_v, end_v + 1):
                        key = (book, chapter, vv)
                        if key in bible_dict:
                            verses_texts.append(bible_dict[key])
            else:
                try:
                    vv = int(piece)
                except ValueError:
                    continue
                key = (book, chapter, vv)
                if key in bible_dict:
                    verses_texts.append(bible_dict[key])

    return " ".join(verses_texts).strip()