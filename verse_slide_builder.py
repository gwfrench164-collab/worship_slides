from pathlib import Path
import textwrap

from pptx_utils import (
    load_template,
    find_template_slide_index,
    add_scripture_slide_from_template,
    TOKEN_VERSE_REF,
    TOKEN_VERSE_TXT,
)

import re

_NBSP = "\u00A0"
_BRACKET_SPAN_RE = re.compile(r"\[(.+?)\]")

def protect_bracket_spans_for_wrapping(text: str) -> str:
    """
    Prevent wrapping from splitting inside [bracketed spans] by converting
    spaces inside the brackets into NBSP (non-breaking spaces).
    """
    def repl(m: re.Match) -> str:
        inner = m.group(1)
        # normalize inner spacing, then make inner spaces non-breaking
        inner = " ".join(inner.split())
        inner = inner.replace(" ", _NBSP)
        return f"[{inner}]"

    return _BRACKET_SPAN_RE.sub(repl, text)

def wrap_verse(text: str, width: int = 55) -> list[str]:
    text = " ".join(text.split())
    if not text:
        return []

    # ✅ protect bracketed italics spans so they don't split across slides
    text = protect_bracket_spans_for_wrapping(text)

    return textwrap.wrap(text, width=width)


def chunks(lines: list[str], size: int) -> list[list[str]]:
    return [lines[i:i+size] for i in range(0, len(lines), size)]


def build_verse_deck(template_path: Path,
                     refs_and_texts: list[tuple[str, str]],
                     output_path: Path,
                     max_lines_per_slide: int = 4,
                     wrap_width: int = 55):
    """
    max_lines_per_slide applies to the VERSE TEXT only.
    The verse reference is placed in its own token box: {{VERSE REF}}.
    """
    prs = load_template(template_path)

    # Find the scripture template slide by tokens (user can move it anywhere)
    tpl_idx = find_template_slide_index(prs, [TOKEN_VERSE_REF, TOKEN_VERSE_TXT])

    for ref, verse_text in refs_and_texts:
        wrapped = wrap_verse(verse_text, width=wrap_width)

        if not wrapped:
            add_scripture_slide_from_template(prs, tpl_idx, ref, ["(text unavailable)"])
            continue

        for chunk in chunks(wrapped, max_lines_per_slide):
            # restore normal spaces (undo NBSP protection)
            chunk = [line.replace("\u00A0", " ") for line in chunk]

            add_scripture_slide_from_template(prs, tpl_idx, ref, chunk)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)