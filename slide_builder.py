import json
import math
import os
from pathlib import Path
from typing import Optional, Tuple, List

from pptx_utils import (
    load_template,
    find_template_slide_index,
    add_title_slide_from_template,
    add_lyrics_slide_from_template,
    add_debug_guides,
    TOKEN_TITLE,
    TOKEN_LYRICS,
)

# Pillow is used ONLY for measuring text width/height; PowerPoint still renders.
from PIL import ImageFont

from debug_tools import DebugSettings, DebugRecorder


# Optional: matplotlib does a good job resolving a font *family name* to a real file path on macOS/Windows/Linux.
try:
    from matplotlib.font_manager import FontProperties, findfont  # type: ignore
except Exception:  # pragma: no cover
    FontProperties = None
    findfont = None


# NOTE:
# We intentionally *do not* freeze debug settings at import time.
# Users may toggle env-vars between runs; we re-read env inside build_deck().
DEBUG_SETTINGS = DebugSettings.from_env()

EMU_PER_INCH = 914400
PT_PER_INCH = 72
MEASURE_DPI = 96  # Pillow measures in pixels; use a consistent DPI conversion
PX_PER_PT = MEASURE_DPI / PT_PER_INCH  # 96/72 = 1.333...


def _emu_to_pt(emu: int) -> float:
    return (emu / EMU_PER_INCH) * PT_PER_INCH


def _find_token_shape(slide, token: str):
    # Prefer shape name (user convention), then fallback to token-in-text
    for shape in slide.shapes:
        try:
            if getattr(shape, "name", None) == token:
                return shape
        except Exception:
            pass

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
    """Best-effort font size (pts) from the lyric placeholder."""
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
    """Return a multiplier for line spacing when possible."""
    try:
        p0 = shape.text_frame.paragraphs[0]
        ls = p0.line_spacing
        if ls is None:
            return 1.10

        if isinstance(ls, (float, int)):
            # If it looks like points, convert to multiplier
            if ls > 3:
                return float(ls) / max(font_size_pts, 1.0)
            return float(ls)

        if hasattr(ls, "pt"):
            return float(ls.pt) / max(font_size_pts, 1.0)
    except Exception:
        pass
    return 1.10


def _font_family_from_shape(shape) -> Optional[str]:
    """Best-effort: pull a font family name from the lyric placeholder."""
    if not getattr(shape, "has_text_frame", False):
        return None

    try:
        tf = shape.text_frame
        if not tf.paragraphs:
            return None
        p0 = tf.paragraphs[0]

        # Try paragraph-level font
        if p0.font and p0.font.name:
            return str(p0.font.name)

        # Try first run font (common in Keynote exports)
        if p0.runs and p0.runs[0].font and p0.runs[0].font.name:
            return str(p0.runs[0].font.name)

        # Try DrawingML default run properties (theme/layout)
        # python-pptx exposes the underlying lxml element as _txBody.
        el = tf._txBody  # noqa: SLF001
        latin = el.xpath(
            ".//a:lstStyle//a:lvl1pPr//a:defRPr//a:latin",
            namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"},
        )
        if latin and "typeface" in latin[0].attrib:
            return latin[0].attrib.get("typeface") or None
    except Exception:
        return None

    return None


def _resolve_font_path(font_family: str | None) -> Optional[str]:
    """Resolve a font *family name* to an installed font file path.

    - If the template provides a family name, try to resolve that.
    - If the family name is missing/empty, try to get a sensible system default.
    """
    family = (font_family or "").strip()

    # 1) Use matplotlib if available (best cross-platform resolver)
    if FontProperties is not None and findfont is not None:
        try:
            fp = FontProperties(family=family) if family else FontProperties()
            path = findfont(fp, fallback_to_default=True)
            if path and os.path.exists(path):
                return path
        except Exception:
            pass

    # 2) Fallback: macOS common font dirs (best-effort filename match)
    if family:
        needle = family.lower().replace(" ", "")
        for d in (
            "/System/Library/Fonts",
            "/System/Library/Fonts/Supplemental",
            "/Library/Fonts",
            os.path.expanduser("~/Library/Fonts"),
        ):
            if not os.path.isdir(d):
                continue
            for fn in os.listdir(d):
                if not fn.lower().endswith((".ttf", ".otf", ".ttc")):
                    continue
                if needle in fn.lower().replace(" ", ""):
                    p = os.path.join(d, fn)
                    if os.path.exists(p):
                        return p

    return None




def _try_load_font(font_path: str | None, size_px: int) -> ImageFont.FreeTypeFont:
    """Load a TrueType/OpenType font robustly on macOS.

    Pillow raises OSError('cannot open resource') when the font file can't be found/opened.
    This helper tries:
      1) the resolved font_path (from the template font name),
      2) a few common macOS fonts,
      3) matplotlib's default font (if available),
      4) Pillow's built-in default font (last resort).
    """
    candidates: List[str] = []
    if font_path:
        candidates.append(font_path)

    candidates.extend([
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Helvetica.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
    ])

    if FontProperties is not None and findfont is not None:
        try:
            dflt = findfont(FontProperties(), fallback_to_default=True)
            if dflt:
                candidates.append(dflt)
        except Exception:
            pass

    for p in candidates:
        try:
            if p and os.path.exists(p):
                return ImageFont.truetype(p, size_px)
        except OSError:
            continue

    return ImageFont.load_default()

def _build_measure_font(template_shape) -> Tuple[ImageFont.FreeTypeFont, float, float, float]:
    """
    Returns:
      (pillow_font, width_px, height_px, line_height_px)
    """
    width_pts = _emu_to_pt(int(template_shape.width))
    height_pts = _emu_to_pt(int(template_shape.height))

    font_size_pts = _best_font_size_pts_from_shape(template_shape)
    line_factor = _line_spacing_factor_from_shape(template_shape, font_size_pts)

    font_family = _font_family_from_shape(template_shape) or ""
    font_path = _resolve_font_path(font_family)  # may be None

    # Convert points to pixels for Pillow measurement
    size_px = max(8, int(round(font_size_pts * PX_PER_PT)))

    font = _try_load_font(font_path, size_px)

    try:
        ascent, descent = font.getmetrics()
    except Exception:
        # Pillow default font may not expose metrics; approximate.
        ascent = int(size_px * 0.8)
        descent = int(size_px * 0.2)
    raw_line_h = (ascent + descent)
    line_h = raw_line_h * max(line_factor, 1.0)

    return font, width_pts * PX_PER_PT, height_pts * PX_PER_PT, float(line_h)


_SMALL_WORDS = {
    "and", "or", "but", "the", "a", "an", "of", "to", "in", "on", "at", "for", "by", "with",
    "that", "this", "is", "be", "as", "if", "so", "yet", "nor", "from"
}


def _text_width_px(font: ImageFont.FreeTypeFont, text: str) -> float:
    # Pillow >= 8 has getlength; fallback to getbbox
    try:
        return float(font.getlength(text))
    except Exception:
        bbox = font.getbbox(text)
        return float(bbox[2] - bbox[0])


def _wrap_one_lyric_line_by_width(line: str, font: ImageFont.FreeTypeFont, max_width_px: float, *, dbg: DebugRecorder | None = None, ctx: dict | None = None) -> List[str]:
    """Wrap a single lyric line by *measured* text width."""
    line = (line or "").strip()
    if not line:
        return []

    # Safety margin: PowerPoint's internal padding/kerning can differ slightly.
    max_w = max_width_px * (dbg.settings.width_safety if (dbg and dbg.settings.enabled) else 0.97)

    words = line.split()
    if dbg and dbg.settings.enabled:
        dbg.log(f"[WRAP] line={line!r} max_width_px={max_width_px:.1f} max_w={max_w:.1f}")
    if ctx is not None:
        ctx.setdefault('wrap', []).append({'input': line, 'max_width_px': max_width_px, 'max_w': max_w, 'steps': []})

    out: List[str] = []
    i = 0

    while i < len(words):
        cur = words[i]
        j = i + 1

        while j < len(words):
            candidate = cur + " " + words[j]
            w = _text_width_px(font, candidate)
            if dbg and dbg.settings.enabled:
                dbg.log(f"[WRAP]   try={candidate!r} w={w:.1f}px ok={w <= max_w}")
            if ctx is not None:
                ctx['wrap'][-1]['steps'].append({'try': candidate, 'w': w, 'ok': w <= max_w})
            if w <= max_w:
                cur = candidate
                j += 1
            else:
                break

        # Avoid ending a line with tiny connector words when possible
        parts = cur.split(" ")
        if len(parts) >= 2 and parts[-1].lower() in _SMALL_WORDS and j < len(words):
            parts.pop()
            cur = " ".join(parts)
            j -= 1

        out.append(cur)
        i = j

    # Anti-orphan: avoid 1 short word on its own last line if we can rebalance
    if len(out) >= 2:
        last_words = out[-1].split()
        if len(last_words) == 1 and len(out[-2].split()) >= 3:
            prev_words = out[-2].split()
            moved = prev_words[-1]
            new_prev = " ".join(prev_words[:-1])
            new_last = moved + " " + out[-1]
            if _text_width_px(font, new_last) <= max_w and _text_width_px(font, new_prev) <= max_w:
                out[-2] = new_prev
                out[-1] = new_last

    return out


def _pack_lyrics_into_slides_by_height(
    lyric_lines: List[str],
    font: ImageFont.FreeTypeFont,
    box_width_px: float,
    box_height_px: float,
    line_height_px: float,
    lyric_gap_em: float = 0.35,
    *,
    dbg: DebugRecorder | None = None,
    ctx: dict | None = None,
) -> List[Tuple[List[str], List[bool]]]:
    """
    Returns list of slides.
    Each slide is (display_lines, lyric_starts):
      - display_lines: wrapped display lines (no blank lines)
      - lyric_starts: True for the first display line of an original lyric line
                      (used to apply paragraph spacing in PPTX without wasting a whole line).
    """
    slides: List[Tuple[List[str], List[bool]]] = []
    cur_lines: List[str] = []
    cur_flags: List[bool] = []
    used_h = 0.0

    gap_px = max(0.0, line_height_px * float(lyric_gap_em))

    for lyric in lyric_lines:
        wrapped = _wrap_one_lyric_line_by_width(lyric, font, box_width_px, dbg=dbg, ctx=ctx)
        if not wrapped:
            continue

        # Height needed if we add this lyric (including gap before it if not the first paragraph on slide)
        add_gap = gap_px if cur_lines else 0.0
        needed_h = add_gap + (len(wrapped) * line_height_px)

        if cur_lines and (used_h + needed_h) > (box_height_px * (dbg.settings.height_safety if (dbg and dbg.settings.enabled) else 0.98)):
            if dbg and dbg.settings.enabled:
                dbg.log(f"[PACK] break: used_h={used_h:.1f}px needed_h={needed_h:.1f}px box_h={box_height_px:.1f}px")
            if ctx is not None:
                ctx.setdefault('pack', []).append({'event': 'break', 'used_h': used_h, 'needed_h': needed_h, 'box_height_px': box_height_px})
            slides.append((cur_lines, cur_flags))
            cur_lines, cur_flags, used_h = [], [], 0.0
            add_gap = 0.0
            needed_h = len(wrapped) * line_height_px

        # If still too tall for an empty slide, hard-split the wrapped display lines
        if not cur_lines and needed_h > (box_height_px * (dbg.settings.height_safety if (dbg and dbg.settings.enabled) else 0.98)):
            # Split by how many lines fit vertically
            max_lines = max(1, int((box_height_px * (dbg.settings.height_safety if (dbg and dbg.settings.enabled) else 0.98)) // max(line_height_px, 1.0)))
            i = 0
            while i < len(wrapped):
                chunk = wrapped[i:i + max_lines]
                flags = [True] + [False] * (len(chunk) - 1)
                slides.append((chunk, flags))
                i += max_lines
            continue

        # Apply gap by marking the first line of this lyric as a new paragraph start
        for k, dl in enumerate(wrapped):
            cur_lines.append(dl)
            cur_flags.append(True if k == 0 else False)

        used_h += add_gap + (len(wrapped) * line_height_px)

    if cur_lines:
        slides.append((cur_lines, cur_flags))

    return slides


def _split_into_lyric_groups(display_lines: List[str], lyric_starts: List[bool]) -> List[Tuple[List[str], List[bool]]]:
    """Split a slide's (lines, flags) into groups representing original lyric lines."""
    groups: List[Tuple[List[str], List[bool]]] = []
    cur_l: List[str] = []
    cur_f: List[bool] = []
    for ln, fl in zip(display_lines, lyric_starts):
        if fl and cur_l:
            groups.append((cur_l, cur_f))
            cur_l, cur_f = [], []
        cur_l.append(ln)
        cur_f.append(fl)
    if cur_l:
        groups.append((cur_l, cur_f))
    return groups


def _join_lyric_groups(groups: List[Tuple[List[str], List[bool]]]) -> Tuple[List[str], List[bool]]:
    lines: List[str] = []
    flags: List[bool] = []
    for gl, gf in groups:
        lines.extend(gl)
        flags.extend(gf)
    return lines, flags


def _rebalance_single_lyric_slides(
    packed: List[Tuple[List[str], List[bool]]],
    *,
    min_lyrics_per_slide: int = 2,
    min_lyrics_left_on_prev: int = 2,
    dbg: DebugRecorder | None = None,
) -> List[Tuple[List[str], List[bool]]]:
    """Heuristic: avoid a slide that contains only one lyric group.

    Moves the *last* lyric group from the previous slide to the *front* of the
    current slide, while preserving wrapped continuations.
    """
    if len(packed) < 2:
        return packed

    out: List[Tuple[List[str], List[bool]]] = [(list(l), list(f)) for (l, f) in packed]

    for i in range(1, len(out)):
        cur_lines, cur_flags = out[i]
        prev_lines, prev_flags = out[i - 1]

        cur_groups = _split_into_lyric_groups(cur_lines, cur_flags)
        prev_groups = _split_into_lyric_groups(prev_lines, prev_flags)

        if len(cur_groups) >= min_lyrics_per_slide:
            continue
        if len(prev_groups) <= min_lyrics_left_on_prev:
            continue

        moved = prev_groups.pop(-1)
        cur_groups.insert(0, moved)

        out[i - 1] = _join_lyric_groups(prev_groups)
        out[i] = _join_lyric_groups(cur_groups)

        if dbg and dbg.settings.enabled:
            dbg.log(f"[REBALANCE] moved 1 lyric group from slide {i} to slide {i+1} (within section)")

    return out


class SlideBuilder:
    def __init__(
        self,
        template_path: Path,
        song_fit_preset: str = "normal",
        lyric_gap_em: float = 0.35,
    ):
        """
        Token-only templates. Adapts to template font + lyric textbox size.

        lyric_gap_em: vertical gap between lyric lines, as a fraction of line height.
                      (0.30-0.45 is a good range)
        """
        self.template_path = template_path
        self.song_fit_preset = song_fit_preset  # kept for backwards compatibility, currently unused
        self.lyric_gap_em = float(lyric_gap_em)

    def _section_lines(self, section: dict) -> List[str]:
        # New format
        if isinstance(section.get("lines"), list):
            return [str(x).rstrip() for x in section.get("lines", []) if str(x).strip()]

        # Legacy format: flatten slides
        out: List[str] = []
        for s in section.get("slides", []):
            for line in s.get("lines", []):
                line = str(line).rstrip()
                if line.strip():
                    out.append(line)
        return out

    def build_deck(self, song_files, output_path: Path):
        # Re-read debug settings every run so users can toggle without reinstalling.
        dbg_settings = DebugSettings.from_env()
        dbg = DebugRecorder(dbg_settings)
        if dbg_settings.enabled:
            dbg.start_run('songs', str(self.template_path), str(output_path))
        prs = load_template(self.template_path)

        title_tpl_idx = find_template_slide_index(prs, [TOKEN_TITLE])
        lyrics_tpl_idx = find_template_slide_index(prs, [TOKEN_LYRICS])

        lyrics_tpl_slide = prs.slides[lyrics_tpl_idx]
        lyrics_shape = _find_token_shape(lyrics_tpl_slide, TOKEN_LYRICS)
        if lyrics_shape is None:
            raise RuntimeError("Template lyrics slide missing {{LYRICS}} placeholder (shape name or token text).")

        family = ""
        resolved = None
        font, box_w_px, box_h_px, line_h_px = _build_measure_font(lyrics_shape)
        if dbg_settings.enabled:
            family = _font_family_from_shape(lyrics_shape) or ''
            resolved = _resolve_font_path(family) if family else None
            dbg.log(f"[FONT] family={family!r} resolved_path={resolved!r}")
            dbg.log(f"[GEOM] box_w_px={box_w_px:.1f} box_h_px={box_h_px:.1f} line_h_px={line_h_px:.1f} lyric_gap_em={self.lyric_gap_em}")
            # TextFrame margins (EMU); useful for understanding under/over-filling
            try:
                tf = lyrics_shape.text_frame
                ml = int(getattr(tf, 'margin_left', 0) or 0)
                mr = int(getattr(tf, 'margin_right', 0) or 0)
                mt = int(getattr(tf, 'margin_top', 0) or 0)
                mb = int(getattr(tf, 'margin_bottom', 0) or 0)
            except Exception:
                ml = mr = mt = mb = 0
            dbg.log(f"[MARGINS] left={ml} right={mr} top={mt} bottom={mb} (EMU)")
            usable_w_emu = max(0, int(lyrics_shape.width) - ml - mr)
            usable_h_emu = max(0, int(lyrics_shape.height) - mt - mb)
            usable_rect_emu = (int(lyrics_shape.left) + ml, int(lyrics_shape.top) + mt, usable_w_emu, usable_h_emu)


        # Convert lyric gap to points for PowerPoint paragraph spacing
        lyric_gap_pt = (line_h_px / PX_PER_PT) * self.lyric_gap_em

        for song_file in song_files:
            with open(song_file, "r", encoding="utf-8") as f:
                song = json.load(f)

            title = song["song"]["title"]
            add_title_slide_from_template(prs, title_tpl_idx, title)

            for section in song["structure"]["sections"]:
                raw_lines = self._section_lines(section)
                if not raw_lines:
                    continue

                ctx = {} if dbg_settings.enabled else None
                packed = _pack_lyrics_into_slides_by_height(
                    raw_lines,
                    font=font,
                    box_width_px=box_w_px,
                    box_height_px=box_h_px,
                    line_height_px=line_h_px,
                    lyric_gap_em=self.lyric_gap_em,
                    dbg=dbg if dbg_settings.enabled else None,
                    ctx=ctx,
                )

                # If packing creates a "lonely" slide with only one short lyric line,
                # rebalance by pulling one lyric group from the previous slide.
                packed = _rebalance_single_lyric_slides(
                    packed,
                    min_lyrics_per_slide=2,
                    min_lyrics_left_on_prev=2,
                    dbg=dbg if dbg_settings.enabled else None,
                )

                for display_lines, lyric_starts in packed:
                    slide = add_lyrics_slide_from_template(
                        prs,
                        lyrics_tpl_idx,
                        display_lines,
                        lyric_starts=lyric_starts,
                        lyric_gap_pt=lyric_gap_pt,
                    )
                    if dbg_settings.enabled and dbg_settings.draw_guides:
                        caption = f"{family} | {os.path.basename(resolved) if resolved else 'unresolved'} | fs={_best_font_size_pts_from_shape(lyrics_shape):.1f}pt"
                        add_debug_guides(slide, slide.shapes[0] if False else _find_token_shape(slide, TOKEN_LYRICS) or slide.shapes[0], usable_rect_emu=usable_rect_emu, caption=caption)

                    if dbg_settings.enabled:
                        dbg.add_slide_record({
                            'type': 'lyrics',
                            'lines': display_lines,
                            'lyric_starts': lyric_starts,
                            'geom': {'box_w_px': box_w_px, 'box_h_px': box_h_px, 'line_h_px': line_h_px},
                            'wrap_pack_ctx': ctx,
                        })
                        # Optional visual guides are added in pptx_utils when enabled
                    

        prs.save(output_path)
        if dbg_settings.enabled:
            dbg.flush()
