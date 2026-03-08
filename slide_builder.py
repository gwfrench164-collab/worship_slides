import json
import os
from pathlib import Path
from typing import Optional, Tuple, List

from pptx_utils import (
    load_template,
    find_template_slide_index,
    add_title_slide_from_template,
    add_lyrics_slide_from_template,
    add_debug_guides,
    remove_slide,
    TOKEN_TITLE,
    TOKEN_LYRICS,
)

from PIL import ImageFont

from debug_tools import DebugSettings, DebugRecorder

try:
    from matplotlib.font_manager import FontProperties, findfont  # type: ignore
except Exception:
    FontProperties = None
    findfont = None

DEBUG_SETTINGS = DebugSettings.from_env()

EMU_PER_INCH = 914400
PT_PER_INCH = 72
MEASURE_DPI = 96
PX_PER_PT = MEASURE_DPI / PT_PER_INCH


def _emu_to_pt(emu: int) -> float:
    return (emu / EMU_PER_INCH) * PT_PER_INCH


def _find_token_shape(slide, token: str):
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


def _slide_contains_token(slide, token_substring: str = "{{") -> bool:
    """Return True if any text on the slide contains token_substring."""
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            try:
                if token_substring in (shape.text or ""):
                    return True
            except Exception:
                pass
    return False


def remove_template_placeholder_slides(prs) -> int:
    """
    Remove any slides that still contain template tokens like {{TITLE}}, {{LYRICS}},
    {{VERSE REF}}, or {{VERSE TXT}}. Returns number removed.
    """
    to_remove = [i for i, s in enumerate(prs.slides) if _slide_contains_token(s, "{{")]
    for i in reversed(to_remove):
        remove_slide(prs, i)
    return len(to_remove)


def _best_font_size_pts_from_shape(shape) -> float:
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
    try:
        p0 = shape.text_frame.paragraphs[0]
        ls = p0.line_spacing
        if ls is None:
            return 1.10
        if isinstance(ls, (float, int)):
            if ls > 3:
                return float(ls) / max(font_size_pts, 1.0)
            return float(ls)
        if hasattr(ls, "pt"):
            return float(ls.pt) / max(font_size_pts, 1.0)
    except Exception:
        pass
    return 1.10


def _font_family_from_shape(shape) -> Optional[str]:
    if not getattr(shape, "has_text_frame", False):
        return None
    try:
        tf = shape.text_frame
        if not tf.paragraphs:
            return None
        p0 = tf.paragraphs[0]
        if p0.font and p0.font.name:
            return str(p0.font.name)
        if p0.runs and p0.runs[0].font and p0.runs[0].font.name:
            return str(p0.runs[0].font.name)
        el = tf._txBody
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
    family = (font_family or "").strip()
    if FontProperties is not None and findfont is not None:
        try:
            fp = FontProperties(family=family) if family else FontProperties()
            path = findfont(fp, fallback_to_default=True)
            if path and os.path.exists(path):
                return path
        except Exception:
            pass

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
    width_pts = _emu_to_pt(int(template_shape.width))
    height_pts = _emu_to_pt(int(template_shape.height))
    font_size_pts = _best_font_size_pts_from_shape(template_shape)
    line_factor = _line_spacing_factor_from_shape(template_shape, font_size_pts)
    font_family = _font_family_from_shape(template_shape) or ""
    font_path = _resolve_font_path(font_family)
    size_px = max(8, int(round(font_size_pts * PX_PER_PT)))
    font = _try_load_font(font_path, size_px)

    try:
        ascent, descent = font.getmetrics()
    except Exception:
        ascent = int(size_px * 0.8)
        descent = int(size_px * 0.2)

    raw_line_h = ascent + descent
    line_h = raw_line_h * max(line_factor, 1.0)
    return font, width_pts * PX_PER_PT, height_pts * PX_PER_PT, float(line_h)


_SMALL_WORDS = {
    "and", "or", "but", "the", "a", "an", "of", "to", "in", "on", "at", "for", "by", "with",
    "that", "this", "is", "be", "as", "if", "so", "yet", "nor", "from"
}


def _text_width_px(font: ImageFont.FreeTypeFont, text: str) -> float:
    try:
        return float(font.getlength(text))
    except Exception:
        bbox = font.getbbox(text)
        return float(bbox[2] - bbox[0])


def _wrap_one_lyric_line_by_width(line: str, font: ImageFont.FreeTypeFont, max_width_px: float, *, dbg: DebugRecorder | None = None, ctx: dict | None = None) -> List[str]:
    line = (line or "").strip()
    if not line:
        return []

    max_w = max_width_px * (dbg.settings.width_safety if (dbg and dbg.settings.enabled) else 0.97)
    words = line.split()

    out: List[str] = []
    i = 0
    while i < len(words):
        cur = words[i]
        j = i + 1
        while j < len(words):
            candidate = cur + " " + words[j]
            if _text_width_px(font, candidate) <= max_w:
                cur = candidate
                j += 1
            else:
                break

        parts = cur.split(" ")
        if len(parts) >= 2 and parts[-1].lower() in _SMALL_WORDS and j < len(words):
            parts.pop()
            cur = " ".join(parts)
            j -= 1

        out.append(cur)
        i = j

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
    slides: List[Tuple[List[str], List[bool]]] = []
    cur_lines: List[str] = []
    cur_flags: List[bool] = []
    used_h = 0.0
    gap_px = max(0.0, line_height_px * float(lyric_gap_em))

    for lyric in lyric_lines:
        wrapped = _wrap_one_lyric_line_by_width(lyric, font, box_width_px, dbg=dbg, ctx=ctx)
        if not wrapped:
            continue

        add_gap = gap_px if cur_lines else 0.0
        needed_h = add_gap + (len(wrapped) * line_height_px)

        if cur_lines and (used_h + needed_h) > (box_height_px * (dbg.settings.height_safety if (dbg and dbg.settings.enabled) else 0.98)):
            slides.append((cur_lines, cur_flags))
            cur_lines, cur_flags, used_h = [], [], 0.0
            add_gap = 0.0
            needed_h = len(wrapped) * line_height_px

        if not cur_lines and needed_h > (box_height_px * (dbg.settings.height_safety if (dbg and dbg.settings.enabled) else 0.98)):
            max_lines = max(1, int((box_height_px * (dbg.settings.height_safety if (dbg and dbg.settings.enabled) else 0.98)) // max(line_height_px, 1.0)))
            i = 0
            while i < len(wrapped):
                chunk = wrapped[i:i + max_lines]
                flags = [True] + [False] * (len(chunk) - 1)
                slides.append((chunk, flags))
                i += max_lines
            continue

        for k, dl in enumerate(wrapped):
            cur_lines.append(dl)
            cur_flags.append(True if k == 0 else False)

        used_h += add_gap + (len(wrapped) * line_height_px)

    if cur_lines:
        slides.append((cur_lines, cur_flags))
    return slides


def _split_into_lyric_groups(display_lines: List[str], lyric_starts: List[bool]) -> List[Tuple[List[str], List[bool]]]:
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
    min_prev_groups_to_borrow: int = 2,
    lonely_max_display_lines: int = 2,
    dbg: DebugRecorder | None = None,
) -> List[Tuple[List[str], List[bool]]]:
    if len(packed) < 2:
        return packed

    out: List[Tuple[List[str], List[bool]]] = [(list(l), list(f)) for (l, f) in packed]
    for i in range(1, len(out)):
        cur_lines, cur_flags = out[i]
        prev_lines, prev_flags = out[i - 1]
        cur_groups = _split_into_lyric_groups(cur_lines, cur_flags)
        prev_groups = _split_into_lyric_groups(prev_lines, prev_flags)

        if len(cur_groups) >= min_lyrics_per_slide or len(cur_lines) > lonely_max_display_lines or len(prev_groups) < min_prev_groups_to_borrow:
            continue

        moved = prev_groups[-1]
        new_prev_groups = prev_groups[:-1]
        new_cur_groups = [moved] + cur_groups
        if not new_prev_groups:
            continue

        prev_remaining_lines = sum(len(gl) for (gl, _gf) in new_prev_groups)
        if prev_remaining_lines <= lonely_max_display_lines and len(new_prev_groups) < min_lyrics_per_slide:
            continue

        out[i - 1] = _join_lyric_groups(new_prev_groups)
        out[i] = _join_lyric_groups(new_cur_groups)
    return out


def _estimate_slide_height_px(lines: List[str], lyric_starts: List[bool], line_height_px: float, gap_px: float) -> float:
    if not lines:
        return 0.0
    para_starts = sum(1 for f in (lyric_starts or []) if f) or 1
    return (len(lines) * float(line_height_px)) + (max(0, para_starts - 1) * float(gap_px))


def _rebalance_tail_slides(
    packed: List[Tuple[List[str], List[bool]]],
    *,
    box_height_px: float,
    line_height_px: float,
    gap_px: float,
    lonely_max_display_lines: int = 2,
    max_tail_display_lines: int = 2,
    min_prev_groups_to_borrow: int = 2,
    dbg: DebugRecorder | None = None,
) -> List[Tuple[List[str], List[bool]]]:
    if len(packed) < 2:
        return packed

    out: List[Tuple[List[str], List[bool]]] = [(list(l), list(f)) for (l, f) in packed]

    def _fits(lines: List[str], flags: List[bool]) -> bool:
        h = _estimate_slide_height_px(lines, flags, line_height_px=line_height_px, gap_px=gap_px)
        limit = box_height_px * (dbg.settings.height_safety if (dbg and dbg.settings.enabled) else 0.98)
        return h <= limit

    def _is_tail(lines: List[str]) -> bool:
        if not lines or len(lines) > max_tail_display_lines:
            return False
        chars = len(" ".join(lines).strip())
        return 0 < chars < 120

    def _is_lonely(lines: List[str], groups_count: int) -> bool:
        return len(lines) <= lonely_max_display_lines and groups_count < 2

    i = len(out) - 1
    while i > 0:
        cur_lines, cur_flags = out[i]
        prev_lines, prev_flags = out[i - 1]
        cur_groups = _split_into_lyric_groups(cur_lines, cur_flags)
        prev_groups = _split_into_lyric_groups(prev_lines, prev_flags)

        if not _is_tail(cur_lines):
            i -= 1
            continue

        merged_lines = prev_lines + cur_lines
        merged_flags = prev_flags + cur_flags
        if _fits(merged_lines, merged_flags):
            out[i - 1] = (merged_lines, merged_flags)
            out.pop(i)
            i = min(i, len(out) - 1)
            continue

        if len(prev_groups) >= min_prev_groups_to_borrow and len(prev_groups) > 1:
            moved = prev_groups[-1]
            new_prev_groups = prev_groups[:-1]
            new_cur_groups = [moved] + cur_groups
            if not new_prev_groups:
                i -= 1
                continue
            new_prev = _join_lyric_groups(new_prev_groups)
            new_cur = _join_lyric_groups(new_cur_groups)
            if not _is_lonely(new_prev[0], len(new_prev_groups)) and _fits(new_prev[0], new_prev[1]) and _fits(new_cur[0], new_cur[1]):
                out[i - 1] = new_prev
                out[i] = new_cur
        i -= 1

    return out


class SlideBuilder:
    def __init__(self, template_path: Path, song_fit_preset: str = "normal", lyric_gap_em: float = 0.35):
        self.template_path = template_path
        self.song_fit_preset = song_fit_preset
        self.lyric_gap_em = float(lyric_gap_em)

    def _section_slide_groups(self, section: dict) -> List[List[str]]:
        # Legacy authored slides: keep each slide boundary intact.
        if isinstance(section.get("slides"), list) and section.get("slides"):
            out: List[List[str]] = []
            for s in section.get("slides", []):
                lines = [str(line).rstrip() for line in s.get("lines", []) if str(line).strip()]
                if lines:
                    out.append(lines)
            return out

        # Newer schema: whole section is one authored unit.
        if isinstance(section.get("lines"), list):
            lines = [str(x).rstrip() for x in section.get("lines", []) if str(x).strip()]
            return [lines] if lines else []

        return []

    def build_deck(self, song_files, output_path: Path):
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
            try:
                tf = lyrics_shape.text_frame
                ml = int(getattr(tf, 'margin_left', 0) or 0)
                mr = int(getattr(tf, 'margin_right', 0) or 0)
                mt = int(getattr(tf, 'margin_top', 0) or 0)
                mb = int(getattr(tf, 'margin_bottom', 0) or 0)
            except Exception:
                ml = mr = mt = mb = 0
            usable_w_emu = max(0, int(lyrics_shape.width) - ml - mr)
            usable_h_emu = max(0, int(lyrics_shape.height) - mt - mb)
            usable_rect_emu = (int(lyrics_shape.left) + ml, int(lyrics_shape.top) + mt, usable_w_emu, usable_h_emu)

        lyric_gap_pt = (line_h_px / PX_PER_PT) * self.lyric_gap_em

        for song_file in song_files:
            try:
                with open(song_file, "r", encoding="utf-8", errors="replace") as f:
                    song = json.load(f)
            except Exception as e:
                if dbg_settings.enabled:
                    dbg.log(f"[SONG] Skipping unreadable song file: {song_file} ({e})")
                continue

            title = song["song"]["title"]
            add_title_slide_from_template(prs, title_tpl_idx, title)

            for section in song["structure"]["sections"]:
                slide_groups = self._section_slide_groups(section)
                if not slide_groups:
                    continue

                for raw_lines in slide_groups:
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

                    # Rebalance only within this authored slide group.
                    packed = _rebalance_single_lyric_slides(
                        packed,
                        min_lyrics_per_slide=2,
                        min_prev_groups_to_borrow=2,
                        dbg=dbg if dbg_settings.enabled else None,
                    )
                    packed = _rebalance_tail_slides(
                        packed,
                        box_height_px=box_h_px,
                        line_height_px=line_h_px,
                        gap_px=(line_h_px * float(self.lyric_gap_em)),
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
                            add_debug_guides(slide, _find_token_shape(slide, TOKEN_LYRICS) or slide.shapes[0], usable_rect_emu=usable_rect_emu, caption=caption)

                        if dbg_settings.enabled:
                            dbg.add_slide_record({
                                'type': 'lyrics',
                                'lines': display_lines,
                                'lyric_starts': lyric_starts,
                                'geom': {'box_w_px': box_w_px, 'box_h_px': box_h_px, 'line_h_px': line_h_px},
                                'wrap_pack_ctx': ctx,
                            })

        remove_template_placeholder_slides(prs)
        prs.save(output_path)
        if dbg_settings.enabled:
            dbg.flush()
