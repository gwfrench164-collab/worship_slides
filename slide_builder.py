import json
from pathlib import Path
from dataclasses import dataclass
import math

from PIL import ImageFont

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
DEFAULT_FONT_SIZE_PT = 60.0

def _emu_to_pt(emu: int) -> float:
    return (emu / EMU_PER_INCH) * PT_PER_INCH

def _pt_to_px(pt: float, dpi: float = 96.0) -> float:
    # 1pt = 1/72 inch
    return (pt / 72.0) * dpi

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
    try:
        tf = shape.text_frame
        if not tf.paragraphs:
            return DEFAULT_FONT_SIZE_PT
        p0 = tf.paragraphs[0]
        if p0.runs and p0.runs[0].font.size:
            return float(p0.runs[0].font.size.pt)
        if p0.font and p0.font.size:
            return float(p0.font.size.pt)
    except Exception:
        pass
    return DEFAULT_FONT_SIZE_PT

def _best_font_name_from_shape(shape) -> str:
    # Try run font name -> paragraph font name -> fall back
    try:
        tf = shape.text_frame
        if tf.paragraphs:
            p0 = tf.paragraphs[0]
            if p0.runs and p0.runs[0].font.name:
                return str(p0.runs[0].font.name)
            if p0.font and p0.font.name:
                return str(p0.font.name)
    except Exception:
        pass
    return "DejaVu Sans"

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

def _text_frame_margins_pt(shape) -> tuple[float, float, float, float]:
    # left, right, top, bottom in points
    try:
        tf = shape.text_frame
        ml = _emu_to_pt(int(getattr(tf, "margin_left", 0) or 0))
        mr = _emu_to_pt(int(getattr(tf, "margin_right", 0) or 0))
        mt = _emu_to_pt(int(getattr(tf, "margin_top", 0) or 0))
        mb = _emu_to_pt(int(getattr(tf, "margin_bottom", 0) or 0))
        return ml, mr, mt, mb
    except Exception:
        return 0.0, 0.0, 0.0, 0.0

@dataclass(frozen=True)
class FitSpec:
    width_pt: float
    height_pt: float
    font_size_pt: float
    font_name: str
    line_height_pt: float
    lyric_gap_pt: float
    # safety margins (empirically helps match PowerPoint wrapping)
    width_safety: float = 0.97
    height_safety: float = 0.98

class _PillowMeasurer:
    def __init__(self, font_name: str, font_size_pt: float, dpi: float = 96.0):
        self.dpi = dpi
        self.font_size_px = max(1, int(round(_pt_to_px(font_size_pt, dpi))))
        self.font = self._load_font(font_name, self.font_size_px)

    @staticmethod
    def _load_font(font_name: str, size_px: int):
        # Best effort: try name directly, else fall back to DejaVuSans which is commonly present
        # (Exact PowerPoint font matching would require shipping/pointing to the .ttf file.)
        try:
            return ImageFont.truetype(font_name, size_px)
        except Exception:
            pass
        try:
            return ImageFont.truetype("DejaVuSans.ttf", size_px)
        except Exception:
            return ImageFont.load_default()

    def text_width_px(self, s: str) -> float:
        s = s or ""
        try:
            # Pillow >= 8
            return float(self.font.getlength(s))
        except Exception:
            try:
                bbox = self.font.getbbox(s)
                return float(bbox[2] - bbox[0])
            except Exception:
                return float(len(s) * self.font_size_px * 0.55)

def _wrap_line_by_width(line: str, meas: _PillowMeasurer, max_width_px: float) -> list[str]:
    line = (line or "").strip()
    if not line:
        return []

    words = line.split()
    out: list[str] = []
    cur = ""

    for w in words:
        candidate = w if not cur else f"{cur} {w}"
        if meas.text_width_px(candidate) <= max_width_px:
            cur = candidate
            continue

        if cur:
            out.append(cur)
            cur = w
        else:
            # single word longer than width; hard-split characters
            chunk = ""
            for ch in w:
                cand2 = chunk + ch
                if meas.text_width_px(cand2) <= max_width_px or not chunk:
                    chunk = cand2
                else:
                    out.append(chunk)
                    chunk = ch
            if chunk:
                cur = chunk

    if cur:
        out.append(cur)

    # De-orphan last line if it's just 1 short word and can be balanced
    if len(out) >= 2:
        last_words = out[-1].split()
        if len(last_words) == 1 and len(last_words[0]) <= 4:
            prev_words = out[-2].split()
            if len(prev_words) >= 3:
                # try moving 1 word from prev to last
                moved = prev_words[-1]
                new_prev = " ".join(prev_words[:-1])
                new_last = f"{moved} {out[-1]}"
                if meas.text_width_px(new_prev) <= max_width_px and meas.text_width_px(new_last) <= max_width_px:
                    out[-2] = new_prev
                    out[-1] = new_last

    return out

def _compute_fit_spec(prs, lyrics_tpl_idx: int, preset: str = "normal") -> FitSpec:
    slide = prs.slides[lyrics_tpl_idx]
    shape = _find_token_shape(slide, TOKEN_LYRICS)
    if shape is None:
        # fallback
        font_size_pt = DEFAULT_FONT_SIZE_PT
        line_height_pt = font_size_pt * 1.10
        return FitSpec(width_pt=600, height_pt=300, font_size_pt=font_size_pt, font_name="DejaVu Sans",
                       line_height_pt=line_height_pt, lyric_gap_pt=font_size_pt * 0.35)

    width_pt = _emu_to_pt(int(shape.width))
    height_pt = _emu_to_pt(int(shape.height))

    ml, mr, mt, mb = _text_frame_margins_pt(shape)
    width_pt = max(10.0, width_pt - (ml + mr))
    height_pt = max(10.0, height_pt - (mt + mb))

    font_size_pt = _best_font_size_pts_from_shape(shape)
    font_name = _best_font_name_from_shape(shape)
    line_factor = _line_spacing_factor_from_shape(shape, font_size_pt)

    preset = (preset or "normal").lower().strip()
    if preset == "tight":
        line_factor = max(line_factor, 1.05)
        lyric_gap_pt = font_size_pt * 0.25
        width_safety = 0.965
        height_safety = 0.975
    elif preset == "loose":
        lyric_gap_pt = font_size_pt * 0.40
        width_safety = 0.98
        height_safety = 0.985
    else:
        lyric_gap_pt = font_size_pt * 0.33
        width_safety = 0.97
        height_safety = 0.98

    line_height_pt = max(font_size_pt * line_factor, font_size_pt * 1.05)

    return FitSpec(
        width_pt=width_pt,
        height_pt=height_pt,
        font_size_pt=font_size_pt,
        font_name=font_name,
        line_height_pt=line_height_pt,
        lyric_gap_pt=lyric_gap_pt,
        width_safety=width_safety,
        height_safety=height_safety,
    )

def _wrap_lyrics(lyric_lines: list[str], fit: FitSpec) -> list[list[dict]]:
    # Returns list of blocks; each block is a lyric line wrapped into display lines with metadata
    meas = _PillowMeasurer(fit.font_name, fit.font_size_pt)
    max_width_px = _pt_to_px(fit.width_pt * fit.width_safety, meas.dpi)

    blocks: list[list[dict]] = []
    for lyric in lyric_lines:
        wrapped = _wrap_line_by_width(lyric, meas, max_width_px)
        if not wrapped:
            continue
        block: list[dict] = []
        for i, t in enumerate(wrapped):
            block.append({"text": t, "is_lyric_start": i == 0})
        blocks.append(block)
    return blocks

def _pack_blocks_into_slides(blocks: list[list[dict]], fit: FitSpec) -> list[list[dict]]:
    max_height_pt = fit.height_pt * fit.height_safety
    line_h = fit.line_height_pt
    gap = fit.lyric_gap_pt

    slides: list[list[dict]] = []
    cur: list[dict] = []
    used_h = 0.0
    at_top = True

    def flush():
        nonlocal cur, used_h, at_top
        if cur:
            slides.append(cur)
        cur = []
        used_h = 0.0
        at_top = True

    for block in blocks:
        block_lines = block
        block_h = len(block_lines) * line_h
        add_gap = (0.0 if at_top else gap)

        if (used_h + add_gap + block_h) <= max_height_pt:
            if not at_top and gap > 0:
                # mark first line of the block as needing spacing (pptx_utils will translate to paragraph.space_before)
                block_lines = [dict(block_lines[0], space_before_pt=gap)] + block_lines[1:]
            cur.extend(block_lines)
            used_h += add_gap + block_h
            at_top = False
            continue

        # doesn't fit as a whole
        if not at_top:
            flush()
            # retry on empty slide
            add_gap = 0.0

        # if block still too tall for an empty slide, split by lines
        i = 0
        while i < len(block):
            remaining = len(block) - i
            # capacity lines on empty slide
            cap = int(max(1, math.floor(max_height_pt / line_h)))
            chunk = block[i:i + cap]
            # On continuation chunks (i>0), do NOT treat as new lyric start spacing at top
            if i == 0:
                # no space_before at very top
                pass
            else:
                # ensure no lyric-start spacing on first line of continuation
                chunk = [dict(chunk[0], is_lyric_start=False)] + chunk[1:]
            slides.append(chunk)
            i += cap

        # after splitting, start fresh
        cur = []
        used_h = 0.0
        at_top = True

    if cur:
        slides.append(cur)

    return slides

class SlideBuilder:
    def __init__(self, template_path: Path, song_fit_preset: str = "normal"):
        self.template_path = template_path
        self.song_fit_preset = song_fit_preset

    def _section_lines(self, section: dict) -> list[str]:
        if isinstance(section.get("lines"), list):
            return [str(x).rstrip() for x in section.get("lines", []) if str(x).strip()]
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

        fit = _compute_fit_spec(prs, lyrics_tpl_idx, preset=self.song_fit_preset)

        for song_file in song_files:
            with open(song_file, "r", encoding="utf-8") as f:
                song = json.load(f)

            title = song["song"]["title"]
            add_title_slide_from_template(prs, title_tpl_idx, title)

            for section in song["structure"]["sections"]:
                raw_lines = self._section_lines(section)
                if not raw_lines:
                    continue

                blocks = _wrap_lyrics(raw_lines, fit)
                slide_bodies = _pack_blocks_into_slides(blocks, fit)

                for body in slide_bodies:
                    # pptx_utils handles spacing using space_before_pt metadata
                    add_lyrics_slide_from_template(
                        prs,
                        lyrics_tpl_idx,
                        body,
                        lyric_gap_pt=fit.lyric_gap_pt,
                        hanging_indent_pt=0.0,
                    )

        output_path.parent.mkdir(parents=True, exist_ok=True)

        for idx in sorted({title_tpl_idx, lyrics_tpl_idx}, reverse=True):
            remove_slide(prs, idx)

        prs.save(output_path)
