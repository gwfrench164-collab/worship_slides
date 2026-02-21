from pathlib import Path
import re

from pptx_utils import (
    load_template,
    find_template_slide_index,
    add_scripture_slide_from_template,
    TOKEN_VERSE_REF,
    TOKEN_VERSE_TXT,
)

# --- constants ---
EMU_PER_PT = 12700  # 1 point = 12700 EMU

# Word Joiner: NOT whitespace; prevents splitting inside protected spans
_WORD_JOINER = "\u2060"

_BRACKET_SPAN_RE = re.compile(r"\[(.+?)\]")
_SENTENCE_BREAK_RE = re.compile(r"[.!?]\s+")
_SOFT_BREAK_RE = re.compile(r"[,;:]\s+")


def _emu_to_points(emu: int) -> float:
    return emu / EMU_PER_PT


def _find_shape_with_token(slide, token: str):
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False):
            try:
                if token in shape.text:
                    return shape
            except Exception:
                pass
    return None


def _best_font_size_pts(shape) -> float:
    """
    Keynote exports often store the real font size on the first run.
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


def _line_height_factor(shape, font_size_pts: float) -> float:
    """
    Convert line spacing to a multiplier. If unknown, assume a reasonable default.
    """
    try:
        p0 = shape.text_frame.paragraphs[0]
        ls = p0.line_spacing
        if ls is None:
            return 1.10
        if isinstance(ls, (float, int)):
            # If it's points (big), convert to multiplier
            if ls > 3:
                return float(ls) / max(font_size_pts, 1.0)
            return float(ls)
        if hasattr(ls, "pt"):
            return float(ls.pt) / max(font_size_pts, 1.0)
    except Exception:
        pass
    return 1.10


def estimate_max_chars_for_box(shape, preset: str = "normal") -> int:
    """
    Estimate how many characters can fit in the verse textbox without overflowing.
    This adapts automatically to template font size and textbox size.
    """
    width_pts = _emu_to_points(int(shape.width))
    height_pts = _emu_to_points(int(shape.height))

    font_size = _best_font_size_pts(shape)
    line_factor = _line_height_factor(shape, font_size)

    # Character width heuristic: wide worship fonts ~0.40–0.46 of font size
    avg_char_w = font_size * 0.43
    chars_per_line = max(10, int(width_pts / max(avg_char_w, 1.0)))

    line_height = font_size * line_factor
    lines_fit = max(1, int(height_pts / max(line_height, 1.0)))

    # safety factor (prevents “text above/below slide”)
    preset = (preset or "normal").lower().strip()
    if preset == "tight":
        safety = 0.78
    elif preset == "loose":
        safety = 0.92
    else:
        safety = 0.85

    max_chars = int(chars_per_line * lines_fit * safety)

    # guardrails so we don’t get silly values
    return max(120, min(max_chars, 1200))


def protect_bracket_spans(text: str) -> str:
    """
    Prevent slide-splitting from cutting inside [bracketed spans]
    by converting spaces inside brackets into WORD_JOINER.
    """
    def repl(m: re.Match) -> str:
        inner = " ".join(m.group(1).split())
        inner = inner.replace(" ", _WORD_JOINER)
        return f"[{inner}]"
    return _BRACKET_SPAN_RE.sub(repl, text)


def restore_bracket_spaces(text: str) -> str:
    return text.replace(_WORD_JOINER, " ")


def split_into_slide_chunks(text: str, max_chars: int) -> list[str]:
    """
    Split text into chunks <= max_chars, preferring sentence breaks, then soft breaks,
    then spaces. Never splits inside protected bracket spans because they contain WORD_JOINER.
    """
    text = " ".join((text or "").split())
    if not text:
        return []

    chunks = []
    remaining = text

    while len(remaining) > max_chars:
        window = remaining[:max_chars]

        # Prefer sentence end
        cut = None
        for m in _SENTENCE_BREAK_RE.finditer(window):
            cut = m.end() - 1  # keep punctuation
        if cut is None:
            # Prefer comma/semicolon/colon
            for m in _SOFT_BREAK_RE.finditer(window):
                cut = m.end() - 1

        if cut is None:
            # last space in window (WORD_JOINER is not whitespace, so protected spans won't break)
            cut = window.rfind(" ")

        # if we still can't find a good cut, hard cut
        if cut is None or cut < int(max_chars * 0.55):
            cut = max_chars

        part = remaining[:cut].strip()
        chunks.append(part)

        remaining = remaining[cut:].strip()

    if remaining:
        chunks.append(remaining)

    return chunks


def build_verse_deck(
    template_path: Path,
    refs_and_texts: list[tuple[str, str]],
    output_path: Path,
    fit_preset: str = "normal",  # tight / normal / loose
):
    prs = load_template(template_path)

    # Find the scripture template slide by tokens (user can move it anywhere)
    tpl_idx = find_template_slide_index(prs, [TOKEN_VERSE_REF, TOKEN_VERSE_TXT])
    tpl_slide = prs.slides[tpl_idx]

    verse_shape = _find_shape_with_token(tpl_slide, TOKEN_VERSE_TXT)
    if verse_shape is None:
        raise RuntimeError("Could not find verse text box containing {{VERSE TXT}} on template slide.")

    max_chars = estimate_max_chars_for_box(verse_shape, preset=fit_preset)

    for ref, verse_text in refs_and_texts:
        verse_text = " ".join((verse_text or "").split())
        if not verse_text:
            add_scripture_slide_from_template(prs, tpl_idx, ref, "(text unavailable)")
            continue

        # Protect bracket spans so we never cut inside them
        protected = protect_bracket_spans(verse_text)

        # Split into slide-sized chunks WITHOUT inserting manual newlines
        raw_chunks = split_into_slide_chunks(protected, max_chars=max_chars)

        for chunk in raw_chunks:
            chunk = restore_bracket_spaces(chunk)
            add_scripture_slide_from_template(prs, tpl_idx, ref, chunk)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)