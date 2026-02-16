import json
from pptx_utils import remove_slide
from pathlib import Path
from pptx_utils import (
    load_template,
    add_title_slide,
    add_lyrics_slide,
    add_title_slide_from_template,
    add_lyrics_slide_from_template,
)

TITLE_TEMPLATE_SLIDE = 0
LYRICS_TEMPLATE_SLIDE = 1


class SlideBuilder:
    def __init__(self, template_path: Path, max_lines: int = 4):
        self.template_path = template_path
        self.max_lines = max_lines

    def _chunk(self, lines: list[str]) -> list[list[str]]:
        size = max(1, int(self.max_lines))
        return [lines[i:i + size] for i in range(0, len(lines), size)]

    def _is_token_template(self, prs) -> bool:
        """
        Keynote-exported templates use normal slides with {{TITLE}} / {{LYRICS}} tokens.
        PowerPoint-authored templates typically use layouts named "Title" / "Bullets".
        """
        try:
            texts = []
            for i in range(min(5, len(prs.slides))):
                for shape in prs.slides[i].shapes:
                    if shape.has_text_frame:
                        texts.append(shape.text_frame.text)
            all_text = "\n".join(texts)
            return ("{{TITLE}}" in all_text) or ("{{LYRICS}}" in all_text)
        except Exception:
            return False

    def _section_lines(self, section: dict) -> list[str]:
        """Return lyric lines for a section (new or legacy format)."""
        # New format
        if isinstance(section.get("lines"), list):
            return [str(x).strip() for x in section.get("lines", []) if str(x).strip()]

        # Legacy format: flatten slides
        out: list[str] = []
        for s in section.get("slides", []):
            for line in s.get("lines", []):
                line = str(line).strip()
                if line:
                    out.append(line)
        return out

    def build_deck(self, song_files, output_path: Path):
        prs = load_template(self.template_path)

        use_tokens = self._is_token_template(prs)

        for song_file in song_files:
            with open(song_file, "r", encoding="utf-8") as f:
                song = json.load(f)

            title = song["song"]["title"]

            if use_tokens:
                add_title_slide_from_template(prs, TITLE_TEMPLATE_SLIDE, title)
            else:
                add_title_slide(prs, title)

            for section in song["structure"]["sections"]:
                lines = self._section_lines(section)
                if not lines:
                    continue

                # Chunk according to max lines per slide
                for chunk in self._chunk(lines):
                    if use_tokens:
                        add_lyrics_slide_from_template(prs, LYRICS_TEMPLATE_SLIDE, chunk)
                    else:
                        add_lyrics_slide(prs, chunk)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        # If using a token-based template, remove the original template slides
        if use_tokens:
            remove_slide(prs, 1)
            remove_slide(prs, 0)

        prs.save(output_path)