from pathlib import Path
import textwrap

from pptx_utils import (
    load_template,
    find_template_slide_index,
    add_scripture_slide_from_template,
    TOKEN_VERSE_REF,
    TOKEN_VERSE_TXT,
)


def wrap_verse(text: str, width: int = 55) -> list[str]:
    text = " ".join(text.split())
    if not text:
        return []
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
            add_scripture_slide_from_template(prs, tpl_idx, ref, chunk)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)