from copy import deepcopy
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.enum.text import PP_ALIGN
import re


def load_template(path):
    return Presentation(path)


# -------------------------
# Slide duplication (preserves Keynote/PowerPoint styling)
# -------------------------

def _copy_relationships(src_slide, dst_slide):
    src_part = src_slide.part
    dst_part = dst_slide.part

    for _rId, rel_obj in src_part.rels.items():
        if rel_obj.is_external:
            continue
        if rel_obj.reltype == RT.SLIDE_LAYOUT:
            continue
        dst_part.rels._add_relationship(rel_obj.reltype, rel_obj.target_part)


def duplicate_slide(prs, slide_index: int):
    src = prs.slides[slide_index]
    blank_layout = prs.slide_layouts[0]
    dst = prs.slides.add_slide(blank_layout)

    # remove default shapes
    for shape in list(dst.shapes):
        el = shape._element
        el.getparent().remove(el)

    # copy background
    if src._element.bg is not None and len(src._element.bg) > 0:
        dst._element.get_or_add_bg()
        dst._element.bg.clear()
        dst._element.bg.append(deepcopy(src._element.bg[0]))

    # copy shapes
    for shape in src.shapes:
        new_el = deepcopy(shape._element)
        dst.shapes._spTree.insert_element_before(new_el, "p:extLst")

    _copy_relationships(src, dst)
    return dst


def remove_slide(prs, index: int):
    slide_id_list = prs.slides._sldIdLst  # pylint: disable=protected-access
    slides = list(slide_id_list)
    slide_id_list.remove(slides[index])


# -------------------------
# Token finding + replacement
# -------------------------

def _slide_text_contains(slide, token: str) -> bool:
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        try:
            if token in shape.text:
                return True
        except Exception:
            pass
    return False


def find_template_slide_index(prs, required_tokens: list[str]) -> int:
    for i, slide in enumerate(prs.slides):
        if all(_slide_text_contains(slide, tok) for tok in required_tokens):
            return i
    raise RuntimeError(f"Template slide not found containing tokens: {required_tokens}")


def _get_best_font_source(paragraph):
    if paragraph.runs:
        return paragraph.runs[0].font
    return paragraph.font


def _copy_font_style(dst_font, src_font):
    dst_font.name = src_font.name
    dst_font.size = src_font.size
    dst_font.bold = src_font.bold
    dst_font.italic = src_font.italic

    if src_font.color is not None:
        try:
            if src_font.color.type == 1 and src_font.color.rgb is not None:
                dst_font.color.rgb = src_font.color.rgb
            elif src_font.color.type == 2 and src_font.color.theme_color is not None:
                dst_font.color.theme_color = src_font.color.theme_color
        except Exception:
            pass


def _force_alignment_like_template(p, template_alignment):
    if template_alignment is None or template_alignment == PP_ALIGN.LEFT:
        p.alignment = PP_ALIGN.LEFT
    else:
        p.alignment = template_alignment


def _replace_token_text(slide, token: str, new_text: str) -> bool:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        tf = shape.text_frame
        if token not in tf.text:
            continue

        p0 = tf.paragraphs[0]
        alignment = p0.alignment
        level = p0.level
        space_before = p0.space_before
        space_after = p0.space_after
        line_spacing = p0.line_spacing
        src_font = _get_best_font_source(p0)

        lines = (new_text or "").split("\n")
        tf.clear()

        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line

            _force_alignment_like_template(p, alignment)
            p.level = level
            p.space_before = space_before
            p.space_after = space_after
            p.line_spacing = line_spacing

            _copy_font_style(p.font, src_font)

        return True

    return False


_BRACKET_ITALIC_RE = re.compile(r"\[(.+?)\]")


def _replace_token_text_with_bracket_italics(slide, token: str, new_text: str) -> bool:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        tf = shape.text_frame
        if token not in tf.text:
            continue

        p0 = tf.paragraphs[0]
        alignment = p0.alignment
        level = p0.level
        space_before = p0.space_before
        space_after = p0.space_after
        line_spacing = p0.line_spacing
        src_font = _get_best_font_source(p0)

        lines = (new_text or "").split("\n")
        tf.clear()

        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()

            _force_alignment_like_template(p, alignment)
            p.level = level
            p.space_before = space_before
            p.space_after = space_after
            p.line_spacing = line_spacing

            p.text = ""
            pos = 0
            for m in _BRACKET_ITALIC_RE.finditer(line):
                before = line[pos:m.start()]
                if before:
                    r = p.add_run()
                    r.text = before
                    _copy_font_style(r.font, src_font)

                ital = m.group(1)
                if ital:
                    r = p.add_run()
                    r.text = ital
                    _copy_font_style(r.font, src_font)
                    r.font.italic = True

                pos = m.end()

            after = line[pos:]
            if after:
                r = p.add_run()
                r.text = after
                _copy_font_style(r.font, src_font)

        return True

    return False


# -------------------------
# Public token helpers
# -------------------------

TOKEN_TITLE = "{{TITLE}}"
TOKEN_LYRICS = "{{LYRICS}}"
TOKEN_VERSE_REF = "{{VERSE REF}}"
TOKEN_VERSE_TXT = "{{VERSE TXT}}"


def add_title_slide_from_template(prs, template_slide_index: int, title_text: str):
    slide = duplicate_slide(prs, template_slide_index)
    if not _replace_token_text(slide, TOKEN_TITLE, title_text):
        raise RuntimeError("Template title slide missing {{TITLE}} token.")
    return slide

from pptx.util import Pt

def add_lyrics_slide_from_template(prs, template_slide_index: int, lines: list[str], hanging_indent_pt: float = 0.0):
    slide = duplicate_slide(prs, template_slide_index)
    lyric_text = "\n".join(lines)

    # Replace token
    if not _replace_token_text(slide, TOKEN_LYRICS, lyric_text):
        raise RuntimeError("Template lyrics slide missing {{LYRICS}} token.")

    # Apply a subtle hanging indent so wrapped continuation lines tuck in slightly.
    # This preserves lyric-line identity (each lyric line is its own paragraph),
    # while making wrapped lines easier to read.
    if hanging_indent_pt and hanging_indent_pt > 0:
        shape = None
        for sh in slide.shapes:
            if getattr(sh, "has_text_frame", False):
                try:
                    if TOKEN_LYRICS in sh.text:
                        shape = sh
                        break
                except Exception:
                    pass
        # NOTE: after replacement, the token text is gone, so the above might fail.
        # Instead, find the first large text frame on the slide and indent it.
        if shape is None:
            for sh in slide.shapes:
                if getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip():
                    shape = sh
                    break

        if shape is not None:
            tf = shape.text_frame
            for p in tf.paragraphs:
                # Hanging indent: first line left, wrapped continuation indented
                p.left_indent = Pt(hanging_indent_pt)
                p.first_line_indent = Pt(-hanging_indent_pt)

    return slide

def add_scripture_slide_from_template(prs, template_slide_index: int, verse_ref: str, verse_text: str):
    slide = duplicate_slide(prs, template_slide_index)

    ok1 = _replace_token_text(slide, TOKEN_VERSE_REF, verse_ref)
    ok2 = _replace_token_text_with_bracket_italics(slide, TOKEN_VERSE_TXT, verse_text)

    if not ok1:
        raise RuntimeError("Scripture template slide missing {{VERSE REF}} token.")
    if not ok2:
        raise RuntimeError("Scripture template slide missing {{VERSE TXT}} token.")
    return slide