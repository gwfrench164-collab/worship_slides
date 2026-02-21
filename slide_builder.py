import json
import math
from pathlib import Path

from pptx_utils import (
    load_template,
    remove_slide,
    find_template_slide_index,
    add_title_slide_from_template,
    add_lyrics_slide_from_template,
    TOKEN_TITLE,
    TOKEN_LYRICS,
)

EMU_PER_INCH = 914400
PT_PER_INCH = 72


def _emu_to_pt(emu: int) -> float:
    return (emu / EMU_PER_INCH) * PT_PER_INCH


def _find_token_shape(slide, token: str):
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        try:
            if token in shape.text:
                return shape
        except Exception:
            continue
    return None


def _best_font_size_pts_from_shape(shape) -> float:
    """
    Keynote/PowerPoint exports often put the real font size on run[0].
    """
    try:
        tf = shape.text_frame
        if not tf.paragraphs:
            return 60.0
        p0 = tf.paragraphs[0]
        if p0.runs and p0.runs[0].font.size:
            return float(p0.runs[0].font.size.pt)
        if p0.font and p0.font.size:
            return float(p0.font.size.pt)
    except Exception:
        pass
    return 60.0


def _line_spacing_factor_from_shape(shape, font_size_pts: float) -> float:
    """
    Try to infer line spacing multiplier from the template paragraph.
    """
    try:
        p0 = shape.text_frame.paragraphs[0]
        ls = p0.line_spacing
        if ls is None:
            return 1.10
        if isinstance(ls, (float, int)):
            # if it's points, convert to multiplier
            if ls > 3:
                return float(ls) / max(font_size_pts, 1.0)
            return float(ls)
        if hasattr(ls, "pt"):
            return float(ls.pt) / max(font_size_pts, 1.0)
    except Exception:
        pass
    return 1.10


def _compute_template_capacity(prs, lyrics_tpl_idx: int):
    """
    Returns:
      chars_per_line_est (int),
      max_visual_lines_in_box (int)

    This adapts automatically if user changes font size or the textbox size in the template.
    """
    slide = prs.slides[lyrics_tpl_idx]
    shape = _find_token_shape(slide, TOKEN_LYRICS)
    if shape is None:
        # safe fallbacks
        return 34, 6

    width_pts = _emu_to_pt(int(shape.width))
    height_pts = _emu_to_pt(int(shape.height))

    font_size_pts = _best_font_size_pts_from_shape(shape)
    line_factor = _line_spacing_factor_from_shape(shape, font_size_pts)

    # Heuristic character width in points (works well for big worship fonts)
    avg_char_w = font_size_pts * 0.43
    chars_per_line = max(10, int(width_pts / max(avg_char_w, 1.0)))

    # Visual lines the box can hold
    line_height_pts = font_size_pts * line_factor
    max_visual_lines = max(1, int(height_pts / max(line_height_pts, 1.0)))

    # Guardrails (prevents extreme weirdness)
    chars_per_line = max(16, min(chars_per_line, 90))
    max_visual_lines = max(3, min(max_visual_lines, 10))

    # Safety margin so we don't hit the very bottom
    max_visual_lines = max(1, int(max_visual_lines * 0.90))

    return chars_per_line, max_visual_lines


def _estimate_visual_lines_for_lyric_line(text: str, chars_per_line: int) -> int:
    """
    Rough estimate of how many wrapped *visual* lines a single lyric line will occupy.
    We do NOT actually wrap here; PowerPoint will wrap naturally.
    This is just to avoid overcrowding.
    """
    t = " ".join((text or "").split())
    if not t:
        return 0
    # +0.15 accounts for punctuation/spacing differences vs pure length
    return max(1, math.ceil((len(t) * 1.15) / max(chars_per_line, 1)))


def _pack_lyric_lines_adaptive(raw_lines: list[str],
                              chars_per_line: int,
                              max_visual_lines: int,
                              max_lyric_lines_cap: int | None = None) -> list[list[str]]:
    """
    Groups ORIGINAL lyric lines into slide-sized chunks based on estimated visual lines.
    Keeps lyric-line identity intact (each original line remains its own paragraph).
    """
    slides: list[list[str]] = []
    cur: list[str] = []
    cur_visual = 0

    for line in raw_lines:
        line = (line or "").rstrip()
        if not line.strip():
            continue

        need = _estimate_visual_lines_for_lyric_line(line, chars_per_line)

        # If empty slide, always take at least one line (even if it "overflows" estimate)
        if not cur:
            cur = [line]
            cur_visual = need
            continue

        # Would adding this line overflow capacity?
        too_many_visual = (cur_visual + need) > max_visual_lines
        too_many_lyric = (max_lyric_lines_cap is not None and (len(cur) + 1) > max_lyric_lines_cap)

        if too_many_visual or too_many_lyric:
            slides.append(cur)
            cur = [line]
            cur_visual = need
        else:
            cur.append(line)
            cur_visual += need

    if cur:
        slides.append(cur)

    # Optional tiny polish: avoid last slide with only 1 line if we can steal from previous
    if len(slides) >= 2 and len(slides[-1]) == 1 and len(slides[-2]) >= 3:
        # move one line from end of previous to start of last
        moved = slides[-2].pop()
        slides[-1].insert(0, moved)

    return slides


class SlideBuilder:
    def __init__(self, template_path: Path, max_lines: int | None = None, hanging_indent_pt: float = 10.0):
        """
        Token-only templates.

        max_lines:
          - If None: fully adaptive (recommended).
          - If set (e.g., 3): acts as a HARD CAP on lyric lines per slide.
            Useful if someone wants a consistent look.

        hanging_indent_pt:
          subtle indent for wrapped continuation lines (PowerPoint does the wrap).
        """
        self.template_path = template_path
        self.max_lines = None if max_lines is None else max(1, int(max_lines))
        self.hanging_indent_pt = float(hanging_indent_pt)

    def _section_lines(self, section: dict) -> list[str]:
        # New format
        if isinstance(section.get("lines"), list):
            return [str(x).rstrip() for x in section.get("lines", []) if str(x).strip()]

        # Legacy format
        out: list[str] = []
        for s in section.get("slides", []):
            for line in s.get("lines", []):
                line = str(line).rstrip()
                if line.strip():
                    out.append(line)
        return out

    def build_deck(self, song_files, output_path: Path):
        prs = load_template(self.template_path)

        title_tpl_idx = find_template_slide_index(prs, [TOKEN_TITLE])
        lyrics_tpl_idx = find_template_slide_index(prs, [TOKEN_LYRICS])

        chars_per_line, max_visual_lines = _compute_template_capacity(prs, lyrics_tpl_idx)

        for song_file in song_files:
            with open(song_file, "r", encoding="utf-8") as f:
                song = json.load(f)

            title = song["song"]["title"]
            add_title_slide_from_template(prs, title_tpl_idx, title)

            for section in song["structure"]["sections"]:
                raw_lines = self._section_lines(section)
                if not raw_lines:
                    continue

                slide_chunks = _pack_lyric_lines_adaptive(
                    raw_lines=raw_lines,
                    chars_per_line=chars_per_line,
                    max_visual_lines=max_visual_lines,
                    max_lyric_lines_cap=self.max_lines,  # None = adaptive-only
                )

                for chunk in slide_chunks:
                    add_lyrics_slide_from_template(
                        prs,
                        lyrics_tpl_idx,
                        chunk,
                        hanging_indent_pt=self.hanging_indent_pt,
                    )

        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Remove the original template slides (remove higher index first)
        for idx in sorted({title_tpl_idx, lyrics_tpl_idx}, reverse=True):
            remove_slide(prs, idx)

        prs.save(output_path)