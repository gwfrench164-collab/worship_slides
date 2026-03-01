from pathlib import Path
import re

from debug_tools import DebugSettings, DebugRecorder

from pptx_utils import (
    load_template,
    find_template_slide_index,
    add_scripture_slide_from_template,
    TOKEN_VERSE_REF,
    TOKEN_VERSE_TXT,
)

DEBUG_SETTINGS = DebugSettings.from_env()


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


def _delete_slide(prs, slide_index: int) -> None:
    """
    Delete slide at slide_index from a python-pptx Presentation.
    This uses the standard XML/relationship removal approach.
    """
    slide_id_list = prs.slides._sldIdLst  # pylint: disable=protected-access
    slide_id_elements = list(slide_id_list)
    sldId = slide_id_elements[slide_index]
    rId = sldId.get("r:id")

    # Remove the slide reference from the slide list
    slide_id_list.remove(sldId)

    # Drop the relationship to the slide part
    if rId in prs.part.rels:
        prs.part.drop_rel(rId)


def remove_template_placeholder_slides(prs) -> int:
    """
    Remove any slides that still contain template tokens like {{TITLE}} or {{VERSE TXT}}.
    Returns number removed.
    """
    to_remove = [i for i, s in enumerate(prs.slides) if _slide_contains_token(s, "{{")]
    for i in reversed(to_remove):
        _delete_slide(prs, i)
    return len(to_remove)



# --- constants ---
EMU_PER_PT = 12700  # 1 point = 12700 EMU

# Word Joiner: NOT whitespace; prevents splitting inside protected spans
_WORD_JOINER = "\u2060"

_BRACKET_SPAN_RE = re.compile(r"\[(.+?)\]")


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


def estimate_line_capacity(shape, preset: str = "normal") -> tuple[int, int]:
    """Return (chars_per_line, lines_fit) for the verse textbox."""
    width_pts = _emu_to_points(int(shape.width))
    height_pts = _emu_to_points(int(shape.height))

    font_size = _best_font_size_pts(shape)
    line_factor = _line_height_factor(shape, font_size)

    # Character width heuristic: wide worship fonts ~0.40–0.46 of font size
    avg_char_w = font_size * 0.43
    chars_per_line = max(10, int(width_pts / max(avg_char_w, 1.0)))

    line_height = font_size * line_factor
    lines_fit = max(1, int(height_pts / max(line_height, 1.0)))

    preset = (preset or "normal").lower().strip()
    if preset == "tight":
        safety = 0.78
    elif preset == "loose":
        safety = 0.92
    else:
        safety = 0.85

    # Apply safety mostly to line count to prevent vertical overflow.
    lines_fit = max(1, int(lines_fit * safety))
    return chars_per_line, lines_fit


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


def wrap_text_to_lines(text: str, max_line_chars: int) -> str:
    """Manual word wrap to explicit '\\n' breaks (prevents PPT auto-reflow)."""
    text = " ".join((text or "").split())
    if not text:
        return ""
    words = text.split(" ")
    lines: list[str] = []
    cur = ""
    for w in words:
        if not w:
            continue
        cand = (cur + " " + w).strip() if cur else w
        if len(cand) <= max_line_chars:
            cur = cand
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return "\n".join(lines)


def _count_lines(block: str) -> int:
    if not block:
        return 0
    return len([ln for ln in block.split("\n") if ln.strip()])


def split_by_lines(text: str, max_line_chars: int, max_lines: int) -> list[str]:
    """
    Deterministically pack verse text into slide-sized blocks with fewer "tail" slides.

    Improvements (v5):
      - Head cleanup can donate multiple clauses (not just one) when the first slide is tiny.
      - If slide 2 is a fixed pre-broken block (contains '\\n'), head cleanup can donate ONE wrapped line.
      - Tail rebalancing can push multiple clauses to avoid tiny last slides.
    """
    text = " ".join((text or "").split())
    if not text:
        return []

    # --- helper: wrap and count lines ---
    def _wrap_lines(s: str) -> list[str]:
        w = wrap_text_to_lines(s, max_line_chars=max_line_chars)
        return [ln for ln in w.split("\n") if ln.strip()]

    def _fits_plain(s: str) -> bool:
        return len(_wrap_lines(s)) <= max_lines

    def _fits_wrapped(block: str) -> bool:
        return len([ln for ln in (block or "").split("\n") if ln.strip()]) <= max_lines

    # --- 1) break into units at punctuation boundaries ---
    units: list[str] = []
    start = 0
    for m in re.finditer(r"[.!?](?=\s|$)|[;:](?=\s|$)|,(?=\s)", text):
        end = m.end()
        u = text[start:end].strip()
        if u:
            units.append(u)
        start = end
    tail = text[start:].strip()
    if tail:
        units.append(tail)

    if not units:
        units = [text]

    # Special-case very short trailing units (e.g., "Amen."):
    # They should *not* become their own slide; attach them to the preceding unit.
    if len(units) >= 2:
        last = units[-1].strip()
        last_words = [w for w in last.split() if w]
        last_compact = re.sub(r"[^A-Za-z0-9]+", "", last).lower()
        if len(last_words) <= 2 or len(last_compact) <= 8:
            units[-2] = (units[-2].rstrip() + " " + last).strip()
            units.pop()

    # --- 2) greedily pack units into slides (store as unit-lists so we can rebalance) ---
    slide_units: list[list[str]] = []
    cur: list[str] = []

    def _flush_current():
        nonlocal cur
        if cur:
            slide_units.append(cur[:])
            cur = []

    for u in units:
        if not cur:
            # If unit itself doesn't fit on an empty slide, split it.
            if not _fits_plain(u):
                u_lines = _wrap_lines(u)
                # chunk these lines into max_lines blocks
                for i in range(0, len(u_lines), max_lines):
                    block = u_lines[i:i + max_lines]
                    slide_units.append(["\n".join(block).strip()])
            else:
                cur = [u]
            continue

        cand = " ".join(cur + [u]).strip()
        if _fits_plain(cand):
            cur.append(u)
        else:
            _flush_current()
            # now start new with u (or split if too large)
            if not _fits_plain(u):
                u_lines = _wrap_lines(u)
                for i in range(0, len(u_lines), max_lines):
                    block = u_lines[i:i + max_lines]
                    slide_units.append(["\n".join(block).strip()])
            else:
                cur = [u]

    _flush_current()

    # Convert unit-lists to wrapped slide strings
    def _render(units_for_slide: list[str]) -> str:
        # If the single unit already contains explicit newlines (fixed block), keep it.
        if len(units_for_slide) == 1 and "\n" in units_for_slide[0]:
            return units_for_slide[0].strip()
        plain = " ".join(units_for_slide).strip()
        return "\n".join(_wrap_lines(plain)).strip()

    slides: list[str] = [_render(u) for u in slide_units]

    # --- 3) tail cleanup + rebalancing ---
    def _is_tiny(block: str) -> bool:
        block = (block or "").strip()
        if not block:
            return True
        ls = [ln for ln in block.split("\n") if ln.strip()]
        if len(ls) <= 2:
            return len(block) < max(85, int(max_line_chars * 2.5))
        return False

    # --- 3a) head cleanup (tiny first slide) ---
    # If the first slide is tiny, try to pull content from slide 2.
    if len(slide_units) >= 2 and _is_tiny(slides[0]):
        fixed0 = bool(slide_units[0]) and ("\n" in slide_units[0][0])
        fixed1 = bool(slide_units[1]) and ("\n" in slide_units[1][0])

        # Case A: slide 2 is NOT fixed (unit-based donation)
        if not fixed0 and not fixed1:
            # If slide 2 is a single large unit, split it into smaller punctuation units
            # so we can "donate" a reasonable first clause.
            if len(slide_units[1]) == 1 and "\n" not in slide_units[1][0]:
                one = slide_units[1][0].strip()
                parts = []
                s = 0
                for m in re.finditer(r"[.!?](?=\s|$)|[;:](?=\s|$)|,(?=\s)", one):
                    e = m.end()
                    u = one[s:e].strip()
                    if u:
                        parts.append(u)
                    s = e
                t = one[s:].strip()
                if t:
                    parts.append(t)
                if len(parts) >= 2:
                    slide_units[1] = parts
                    slides = [_render(u) for u in slide_units]

            # Donate multiple units if needed and if it still fits.
            while len(slide_units[1]) > 1:
                first_unit = slide_units[1].pop(0)
                slide_units[0].append(first_unit)

                head_wrapped = _render(slide_units[0])
                next_wrapped = _render(slide_units[1])

                if _fits_wrapped(head_wrapped) and _fits_wrapped(next_wrapped) and not _is_tiny(head_wrapped):
                    slides = [_render(u) for u in slide_units]
                    break

                # If it doesn't fit, revert and stop.
                if not (_fits_wrapped(head_wrapped) and _fits_wrapped(next_wrapped)):
                    slide_units[0].pop()
                    slide_units[1].insert(0, first_unit)
                    break

                # It fit but head is still tiny: keep the donation and try one more.
                slides = [_render(u) for u in slide_units]
                continue

        # Case B: slide 2 IS fixed (contains '\\n') -> donate ONE wrapped line
        elif not fixed0 and fixed1 and len(slide_units[1]) == 1:
            lines2 = [ln for ln in slide_units[1][0].split("\n") if ln.strip()]
            if len(lines2) >= 2:
                donate_line = lines2.pop(0)

                head_plain = " ".join(slide_units[0]).replace("\n", " ").strip()
                new_head_wrapped = "\n".join(_wrap_lines((head_plain + " " + donate_line).strip())).strip()
                new_next_wrapped = "\n".join(lines2).strip()

                if _fits_wrapped(new_head_wrapped) and _fits_wrapped(new_next_wrapped) and not _is_tiny(new_head_wrapped):
                    slide_units[0] = [(head_plain + " " + donate_line).strip()]
                    slide_units[1] = [new_next_wrapped]
                    slides = [_render(u) for u in slide_units]


    # --- 3aa) General head polish (fix 1-line verse openings anywhere) ---
    # Sometimes a new slide begins with a single short wrapped line (often the beginning of a verse),
    # which looks stranger than a short ending. If that happens, borrow ONE wrapped line from the
    # *top* of the following slide, but only if both slides still fit.
    def _strip_lines(block: str) -> list[str]:
        return [ln.strip() for ln in (block or "").split("\n") if ln.strip()]

    def _rewrap_from_lines(lines: list[str]) -> str:
        plain = " ".join(lines).strip()
        return "\n".join(_wrap_lines(plain)).strip()

    def _is_one_line(block: str) -> bool:
        return len(_strip_lines(block)) == 1

    # Try up to 2 passes because fixing one head can reveal another downstream.
    for _pass in range(2):
        changed = False
        for i in range(len(slides) - 1):
            a_lines = _strip_lines(slides[i])
            b_lines = _strip_lines(slides[i + 1])

            if len(a_lines) != 1:
                continue
            if len(b_lines) < 2:
                continue

            moved = b_lines.pop(0)
            new_a = _rewrap_from_lines(a_lines + [moved])
            new_b = _rewrap_from_lines(b_lines)

            if _fits_wrapped(new_a) and _fits_wrapped(new_b) and not _is_one_line(new_a):
                slides[i] = new_a
                slides[i + 1] = new_b
                changed = True

        if not changed:
            break

    # --- 3b) Iteratively repair from the end (tails) ---

    # --- 3b) FORCE-FIX "1-LINE VERSE START" HEADS (word-level borrow) ---
    # Some verses naturally begin with a short clause ending in a comma, e.g.:
    #   "But the hour cometh, and now is,"
    # Unit-based packing can leave that as its own slide. This step borrows the
    # minimum number of words from the next slide to make the first slide at least
    # 2 wrapped lines (while keeping both slides within max_lines).
    if len(slide_units) >= 2:
        head_wrapped = slides[0]
        head_lines = [ln for ln in (head_wrapped or "").split("\n") if ln.strip()]
        head_plain = " ".join(slide_units[0]).replace("\n", " ").strip()
        if len(head_lines) == 1 and head_plain.endswith((",", ";", ":")):
            next_plain_full = " ".join(slide_units[1]).replace("\n", " ").strip()
            next_words = [w for w in next_plain_full.split() if w]

            # Try borrowing an increasing number of words until the head becomes >=2 lines.
            # Cap borrow to avoid eating the whole next slide.
            max_borrow = min(len(next_words) - 1, max(6, int(max_line_chars * 0.8)))
            for k in range(3, max_borrow + 1):
                borrowed = " ".join(next_words[:k]).strip()
                remaining = " ".join(next_words[k:]).strip()

                if not remaining:
                    break

                new_head_plain = (head_plain + " " + borrowed).strip()
                new_next_plain = remaining

                new_head_wrapped = "\n".join(_wrap_lines(new_head_plain)).strip()
                new_next_wrapped = "\n".join(_wrap_lines(new_next_plain)).strip()

                # Require: both fit, and head is no longer a 1-line fragment.
                new_head_lines = [ln for ln in new_head_wrapped.split("\n") if ln.strip()]
                if _fits_wrapped(new_head_wrapped) and _fits_wrapped(new_next_wrapped) and len(new_head_lines) >= 2:
                    slide_units[0] = [new_head_plain]
                    slide_units[1] = [new_next_plain]
                    slides = [_render(u) for u in slide_units]
                    break
    j = len(slide_units) - 1
    while j > 0:
        tail_text = slides[j]
        if not _is_tiny(tail_text):
            j -= 1
            continue

        # 1) Try merging tail into previous slide.
        prev_plain = " ".join(slide_units[j - 1]).replace("\n", " ").strip()
        tail_plain = " ".join(slide_units[j]).replace("\n", " ").strip()
        merged_wrapped = "\n".join(_wrap_lines((prev_plain + " " + tail_plain).strip())).strip()

        if _fits_wrapped(merged_wrapped):
            slide_units[j - 1] = [(prev_plain + " " + tail_plain).strip()]
            slide_units.pop(j)
            slides = [_render(u) for u in slide_units]
            j = len(slide_units) - 1
            continue

        # 2) If merge doesn't fit, try moving last units from prev -> tail until tail isn't tiny (or can't).
        moved = False
        fixed_prev = bool(slide_units[j - 1]) and ("\n" in slide_units[j - 1][0])
        fixed_tail = bool(slide_units[j]) and ("\n" in slide_units[j][0])

        if not fixed_prev and not fixed_tail:
            while len(slide_units[j - 1]) > 1:
                last_unit = slide_units[j - 1].pop()
                slide_units[j].insert(0, last_unit)

                prev_wrapped = _render(slide_units[j - 1])
                tail_wrapped = _render(slide_units[j])

                if _fits_wrapped(prev_wrapped) and _fits_wrapped(tail_wrapped) and not _is_tiny(tail_wrapped):
                    moved = True
                    break

                # If it doesn't fit, revert and stop.
                if not (_fits_wrapped(prev_wrapped) and _fits_wrapped(tail_wrapped)):
                    slide_units[j].pop(0)
                    slide_units[j - 1].append(last_unit)
                    break

                # It fit but tail still tiny: keep the move and try one more.
                slides = [_render(u) for u in slide_units]
                continue

        if moved:
            slides = [_render(u) for u in slide_units]
            j = len(slide_units) - 1
            continue

        j -= 1

    # --- 3c) FORCE: avoid a 1-line opening clause ending with , ; : ---
    # Example: "But the hour cometh, and now is," (John 4:23) should not be alone.
    def _nonempty_lines(block: str) -> list[str]:
        return [ln for ln in (block or "").split("\n") if ln.strip()]

    if len(slides) >= 2:
        head = (slides[0] or "").strip()
        head_lines = _nonempty_lines(head)

        if len(head_lines) == 1 and head.endswith((",", ";", ":")):
            head_words = head_lines[0].split()
            next_words = " ".join(_nonempty_lines(slides[1])).split()

            # Borrow a few words at a time, up to a small cap.
            for _ in range(16):
                if len(next_words) < 3:
                    break

                take = min(3, len(next_words))  # slightly stronger than 2 to ensure progress
                moved = next_words[:take]
                next_words = next_words[take:]

                new_head_plain = " ".join(head_words + moved).strip()
                new_next_plain = " ".join(next_words).strip()

                new_head = "\n".join(_wrap_lines(new_head_plain)).strip()
                new_next = "\n".join(_wrap_lines(new_next_plain)).strip() if new_next_plain else ""

                if _fits_wrapped(new_head) and (not new_next or _fits_wrapped(new_next)) and len(_nonempty_lines(new_head)) >= 2:
                    slides[0] = new_head
                    if new_next:
                        slides[1] = new_next
                    else:
                        slides.pop(1)
                    break

                # If head overflowed, stop trying.
                if not _fits_wrapped(new_head):
                    break
    return slides


def build_verse_deck(
    template_path: Path,
    refs_and_texts: list[tuple[str, str]],
    output_path: Path,
    fit_preset: str = "normal",  # tight / normal / loose
):
    dbg = DebugRecorder(DEBUG_SETTINGS)
    if DEBUG_SETTINGS.enabled:
        dbg.start_run("verses", str(template_path), str(output_path))

    prs = load_template(template_path)

    # Find the scripture template slide by tokens (user can move it anywhere)
    tpl_idx = find_template_slide_index(prs, [TOKEN_VERSE_REF, TOKEN_VERSE_TXT])
    tpl_slide = prs.slides[tpl_idx]

    verse_shape = _find_shape_with_token(tpl_slide, TOKEN_VERSE_TXT)
    if verse_shape is None:
        raise RuntimeError("Could not find verse text box containing {{VERSE TXT}} on template slide.")

    # Capacity derived from the actual textbox and template font.
    chars_per_line, lines_fit = estimate_line_capacity(verse_shape, preset=fit_preset)

    # Keep a little margin so we don't pack to the absolute edge.
    max_lines_per_slide = max(2, lines_fit)
    max_chars_per_line = max(18, min(chars_per_line, 60))

    if DEBUG_SETTINGS.enabled:
        dbg.log(
            f"[VERSE] chars_per_line={chars_per_line} lines_fit={lines_fit} "
            f"max_line_chars={max_chars_per_line} max_lines={max_lines_per_slide} fit_preset={fit_preset!r}"
        )
        try:
            tf = verse_shape.text_frame
            ml = int(getattr(tf, "margin_left", 0) or 0)
            mr = int(getattr(tf, "margin_right", 0) or 0)
            mt = int(getattr(tf, "margin_top", 0) or 0)
            mb = int(getattr(tf, "margin_bottom", 0) or 0)
        except Exception:
            ml = mr = mt = mb = 0
        dbg.log(f"[MARGINS] left={ml} right={mr} top={mt} bottom={mb} (EMU)")
    added_slides = 0


    for ref, verse_text in refs_and_texts:
        verse_text = " ".join((verse_text or "").split())
        if (not verse_text) or ("text unavailable" in str(verse_text).lower()):
            # Skip entirely (do NOT create a slide)
            continue

        # Protect bracket spans so we never cut inside them
        protected = protect_bracket_spans(verse_text)

        # Split into line-wrapped, slide-sized blocks.
        blocks = split_by_lines(
            protected,
            max_line_chars=max_chars_per_line,
            max_lines=max_lines_per_slide,
        )

        # --- SPECIAL-CASE POLISH: John 4:23 opening line balance ---
        # John 4:23 often starts with a comma-ended clause that can land as:
        #   "But the hour cometh, and now is,"
        #   "when the true"
        # We *only* special-case this reference to avoid global regressions.
        if ref.strip() == "John 4:23" and len(blocks) >= 2:
            def _lines(s: str) -> list[str]:
                return [ln for ln in (s or "").split("\n") if ln.strip()]

            def _rewrap(plain: str) -> str:
                plain = " ".join((plain or "").split())
                return wrap_text_to_lines(plain, max_line_chars=max_chars_per_line).strip()

            def _last_line_ok(s: str) -> bool:
                ls = _lines(s)
                if not ls:
                    return False
                last = ls[-1].strip()
                # Require a reasonably "filled" last line: either enough words, or enough characters.
                return (len(last.split()) >= 6) or (len(last) >= int(max_chars_per_line * 0.55))

            b0 = (blocks[0] or "").strip()
            b1 = (blocks[1] or "").strip()

            l0 = _lines(b0)

            # Trigger if the first block is 1 line ending in punctuation OR 2 lines with a tiny last line.
            trigger = (
                (len(l0) == 1 and b0.endswith((",", ";", ":")))
                or (len(l0) >= 2 and not _last_line_ok(b0))
            )

            if trigger:
                w0 = b0.replace("\n", " ").split()
                w1 = b1.replace("\n", " ").split()

                # Borrow a couple words at a time until the last line looks balanced,
                # while keeping both blocks within max_lines_per_slide.
                for _ in range(40):
                    if len(w1) < 3:
                        break

                    take = min(2, len(w1))
                    moved = w1[:take]
                    w1 = w1[take:]
                    w0 = w0 + moved

                    nb0 = _rewrap(" ".join(w0))
                    nb1 = _rewrap(" ".join(w1))

                    if _count_lines(nb0) <= max_lines_per_slide and _count_lines(nb1) <= max_lines_per_slide and _last_line_ok(nb0):
                        blocks[0] = nb0
                        blocks[1] = nb1
                        break

                    # Stop if we overflow block0; don't risk regressions
                    if _count_lines(nb0) > max_lines_per_slide:
                        break

        for block in blocks:

            block = restore_bracket_spaces(block)
            add_scripture_slide_from_template(prs, tpl_idx, ref, block)
            added_slides += 1
            if DEBUG_SETTINGS.enabled:
                dbg.add_slide_record(
                    {
                        "type": "verse",
                        "ref": ref,
                        "chunk": block,
                        "max_line_chars": max_chars_per_line,
                        "max_lines": max_lines_per_slide,
                    }
                )

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Remove any leftover template placeholder slides (those containing '{{...}}')
    # BUT only if we actually created verse slides. This prevents a blank deck if
    # a notes file yields zero verses.
    if added_slides > 0:
        removed = remove_template_placeholder_slides(prs)
        if DEBUG_SETTINGS.enabled:
            dbg.log(f"[CLEANUP] removed {removed} template placeholder slide(s) containing '{{{{' tokens")

    prs.save(output_path)

    if DEBUG_SETTINGS.enabled:
        dbg.flush()
