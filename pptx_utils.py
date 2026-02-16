from copy import deepcopy
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT


def load_template(path):
    return Presentation(path)


def _get_layout_by_name(prs, name):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    raise ValueError(f"Slide layout not found: {name}")


# -------------------------
# Layout/placeholder approach (PowerPoint-authored templates)
# -------------------------

def add_title_slide(prs, title_text):
    layout = _get_layout_by_name(prs, "Title")
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = title_text


def add_lyrics_slide(prs, lines):
    layout = _get_layout_by_name(prs, "Bullets")
    slide = prs.slides.add_slide(layout)

    body = None
    for placeholder in slide.placeholders:
        if placeholder.placeholder_format.type == PP_PLACEHOLDER_TYPE.BODY:
            body = placeholder
            break

    if body is None:
        raise RuntimeError("No BODY placeholder found on Bullets layout")

    tf = body.text_frame
    tf.clear()

    for i, line in enumerate(lines):
        if i == 0:
            tf.text = line
        else:
            tf.add_paragraph().text = line


# -------------------------
# Token-based approach (Keynote-exported templates)
# Uses slides that literally contain {{TITLE}} and {{LYRICS}} text boxes
# -------------------------

def _copy_relationships(src_slide, dst_slide):
    """
    Copy only the media relationships needed for images/backgrounds.

    IMPORTANT:
    python-pptx _add_relationship(reltype, target, is_external=False)
    does NOT accept an rId argument. Passing a string as the 3rd arg sets
    is_external=True and will crash at save time.
    """
    src_part = src_slide.part
    dst_part = dst_slide.part

    for rId, rel_obj in src_part.rels.items():
        # Skip external relationships entirely
        if rel_obj.is_external:
            continue

        # Skip slideLayout relationship (dst slide already has one)
        if rel_obj.reltype == RT.SLIDE_LAYOUT:
            continue

        # Copy images/media/etc. (internal parts)
        dst_part.rels._add_relationship(rel_obj.reltype, rel_obj.target_part)

def duplicate_slide(prs, slide_index: int):
    """
    Duplicate a slide (including shapes and background) so exported Keynote styling is preserved.
    """
    src = prs.slides[slide_index]
    blank_layout = prs.slide_layouts[0]
    dst = prs.slides.add_slide(blank_layout)

    # remove default shapes on destination slide
    for shape in list(dst.shapes):
        el = shape._element
        el.getparent().remove(el)

    # copy background if present
    if src._element.bg is not None and len(src._element.bg) > 0:
        dst._element.get_or_add_bg()
        dst._element.bg.clear()
        dst._element.bg.append(deepcopy(src._element.bg[0]))

    # copy shapes
    for shape in src.shapes:
        new_el = deepcopy(shape._element)
        dst.shapes._spTree.insert_element_before(new_el, "p:extLst")

    # copy image relationships, etc.
    _copy_relationships(src, dst)
    return dst

def remove_slide(prs, index: int):
    """
    Remove slide at given index (0-based).
    """
    slide_id_list = prs.slides._sldIdLst  # pylint: disable=protected-access
    slides = list(slide_id_list)
    slide_id_list.remove(slides[index])

def _replace_token_text(slide, token: str, new_text: str) -> bool:
    """
    Replace token text while preserving formatting from the token paragraph.
    Preserves: font, alignment, line spacing, space before/after, indent/level.
    """
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        tf = shape.text_frame
        if token not in tf.text:
            continue

        # --- Capture paragraph + font formatting from the token slide ---
        p0 = tf.paragraphs[0]
        alignment = p0.alignment
        level = p0.level
        space_before = p0.space_before
        space_after = p0.space_after
        line_spacing = p0.line_spacing

        # Prefer paragraph font (Keynote exports often set formatting here)
        pf = p0.font
        font_name = pf.name
        font_size = pf.size
        font_bold = pf.bold
        font_italic = pf.italic

        # Color can be tricky (theme vs RGB). Only keep RGB if set.
        font_color = None
        if pf.color is not None and pf.color.type == 1:
            font_color = pf.color.rgb

        # --- Rewrite text ---
        lines = new_text.split("\n")
        tf.clear()

        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = line

            # paragraph formatting
            p.alignment = alignment
            p.level = level
            p.space_before = space_before
            p.space_after = space_after
            p.line_spacing = line_spacing

            # paragraph font formatting
            p.font.name = font_name
            p.font.size = font_size
            p.font.bold = font_bold
            p.font.italic = font_italic
            if font_color is not None:
                p.font.color.rgb = font_color

        return True

    return False

def add_title_slide_from_template(prs, template_slide_index: int, title_text: str):
    slide = duplicate_slide(prs, template_slide_index)
    if not _replace_token_text(slide, "{{TITLE}}", title_text):
        raise RuntimeError("Template title slide missing {{TITLE}} token.")


def add_lyrics_slide_from_template(prs, template_slide_index: int, lines: list[str]):
    slide = duplicate_slide(prs, template_slide_index)
    lyric_text = "\n".join(lines)
    if not _replace_token_text(slide, "{{LYRICS}}", lyric_text):
        raise RuntimeError("Template lyrics slide missing {{LYRICS}} token.")
