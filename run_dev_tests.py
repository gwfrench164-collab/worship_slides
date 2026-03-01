#!/usr/bin/env python3
"""
Development test runner (no GUI).

Runs:
- Song deck generation using SlideBuilder
- Verse deck generation from sermon notes using bible_extractor + verse_slide_builder

Outputs:
- decks/ songs_test.pptx, verses_test.pptx
- qa/ qa_report.txt + qa_report.json (simple heuristics)

Usage (examples):
  python run_dev_tests.py --template templates/template_from_service.pptx --songs_dir /path/to/songs --notes_dir /path/to/notes --out_dir dev_out
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from slide_builder import SlideBuilder
from notes_reader import read_notes_text
from bible_extractor import extract_ordered_refs, load_bible_json, fetch_verse_text, DEFAULT_KJV_PATH
from verse_slide_builder import build_verse_deck
from qa_tools import analyze_pptx


def _collect_song_jsons(songs_dir: Path) -> list[Path]:
    return sorted([p for p in songs_dir.rglob("*.json") if p.is_file()])


def _collect_notes_files(notes_dir: Path) -> list[Path]:
    exts = {".docx", ".txt"}  # .pages requires macOS "osascript"
    return sorted([p for p in notes_dir.rglob("*") if p.is_file() and p.suffix.lower() in exts])


def _refs_and_texts_from_notes(notes_path: Path, bible_dict) -> list[tuple[str, str]]:
    text = read_notes_text(notes_path)
    refs = extract_ordered_refs(text)
    out: list[tuple[str, str]] = []
    for r in refs:
        verse_text = fetch_verse_text(r, bible_dict)
        out.append((r, verse_text))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True, help="Path to PPTX template that contains {{TITLE}}, {{LYRICS}}, {{VERSE TXT}}, {{VERSE REF}}")
    ap.add_argument("--songs_dir", required=True, help="Folder containing song .json files")
    ap.add_argument("--notes_dir", required=True, help="Folder containing sermon notes (.docx/.pages/.txt)")
    ap.add_argument("--out_dir", default="dev_out", help="Output folder")
    ap.add_argument("--max_notes", type=int, default=3, help="How many notes files to process for verses")
    ap.add_argument("--song_fit", default="normal", choices=["tight","normal","loose"], help="Song fit preset")
    ap.add_argument("--verse_fit", default="normal", choices=["tight","normal","loose"], help="Verse fit preset")
    args = ap.parse_args()

    template = Path(args.template).expanduser().resolve()
    songs_dir = Path(args.songs_dir).expanduser().resolve()
    notes_dir = Path(args.notes_dir).expanduser().resolve()
    out_dir = Path(args.out_dir).expanduser().resolve()
    decks_dir = out_dir / "decks"
    qa_dir = out_dir / "qa"
    decks_dir.mkdir(parents=True, exist_ok=True)
    qa_dir.mkdir(parents=True, exist_ok=True)

    # --- Songs ---
    song_files = _collect_song_jsons(songs_dir)
    if not song_files:
        raise SystemExit(f"No song .json files found under {songs_dir}")
    songs_out = decks_dir / "songs_test.pptx"
    sb = SlideBuilder(template, song_fit_preset=args.song_fit)
    sb.build_deck(song_files, songs_out)

    # --- Verses (from notes) ---
    bible = load_bible_json(str(DEFAULT_KJV_PATH))
    notes_files = _collect_notes_files(notes_dir)[: max(1, int(args.max_notes))]
    all_refs: list[tuple[str,str]] = []
    for nf in notes_files:
        all_refs.extend(_refs_and_texts_from_notes(nf, bible))

    verses_out = decks_dir / "verses_test.pptx"
    build_verse_deck(template, all_refs, verses_out, fit_preset=args.verse_fit)

    # --- QA ---
    report = {
        "songs": analyze_pptx(songs_out),
        "verses": analyze_pptx(verses_out),
    }
    (qa_dir / "qa_report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")

    # human-readable summary
    lines = []
    for k in ("songs","verses"):
        lines.append(f"== {k.upper()} ==")
        r = report[k]
        lines.append(f"slides: {r['slide_count']}")
        lines.append(f"flags: sparse={len(r['flags']['SPARSE'])} tail={len(r['flags'].get('TAIL', []))} orphan={len(r['flags'].get('ORPHAN_START', []))} crowded={len(r['flags']['CROWDED'])} tiny_text={len(r['flags']['TINY_TEXT'])}")
        if r["flags"]["SPARSE"]:
            lines.append("  SPARSE slides: " + ", ".join(map(str, r["flags"]["SPARSE"][:20])))
        if r['flags'].get('TAIL'):
            lines.append("  TAIL slides: " + ", ".join(map(str, r['flags']['TAIL'][:20])))
        if r['flags'].get('ORPHAN_START'):
            lines.append("  ORPHAN_START slides: " + ", ".join(map(str, r['flags']['ORPHAN_START'][:20])))
        if r["flags"]["CROWDED"]:
            lines.append("  CROWDED slides: " + ", ".join(map(str, r["flags"]["CROWDED"][:20])))
        if r["flags"]["TINY_TEXT"]:
            lines.append("  TINY_TEXT slides: " + ", ".join(map(str, r["flags"]["TINY_TEXT"][:20])))
        lines.append("")
    (qa_dir / "qa_report.txt").write_text("\n".join(lines), encoding="utf-8")

    print("Wrote:")
    print(" -", songs_out)
    print(" -", verses_out)
    print(" -", qa_dir / "qa_report.txt")
    print(" -", qa_dir / "qa_report.json")


if __name__ == "__main__":
    main()