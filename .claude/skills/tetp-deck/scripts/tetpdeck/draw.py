"""Low-level drawing helpers. The only module that touches python-pptx shapes
directly for pattern content; everything routes through tokens + geometry."""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from . import tokens as T
from .geometry import Rect


def _align(rtl: bool, default=PP_ALIGN.LEFT):
    if default == PP_ALIGN.CENTER:
        return default
    return PP_ALIGN.RIGHT if rtl else PP_ALIGN.LEFT


def box(slide, rect: Rect, fill: str | None = None, line: str | None = None,
        shape=MSO_SHAPE.RECTANGLE):
    sh = slide.shapes.add_shape(
        shape, Inches(rect.x), Inches(rect.y), Inches(rect.w), Inches(rect.h)
    )
    sh.shadow.inherit = False
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = RGBColor.from_string(line)
        sh.line.width = Pt(0.75)
    sh.text_frame.word_wrap = True
    return sh


def text(slide, rect: Rect, runs, *, size=T.BODY_SIZE, color=T.INK, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, rtl: bool = False,
         spacing: float = T.LINE_SPACING):
    """Add a borderless text box. `runs` is a string or list of paragraph dicts:
    {text, size?, color?, bold?, bullet?, level?, space_before?}."""
    tb = slide.shapes.add_textbox(
        Inches(rect.x), Inches(rect.y), Inches(rect.w), Inches(rect.h)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(runs, str):
        runs = [{"text": runs}]
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = para.get("align", _align(rtl, align))
        p.line_spacing = spacing
        if para.get("space_before"):
            p.space_before = Pt(para["space_before"])
        if para.get("level") is not None:
            p.level = para["level"]
        txt = para["text"]
        if para.get("bullet"):
            txt = ("– " if para.get("level") else "• ") + txt
        r = p.add_run()
        r.text = txt
        f = r.font
        f.name = T.FONT
        f.size = para.get("size", size)
        f.bold = para.get("bold", bold)
        f.color.rgb = RGBColor.from_string(para.get("color", color))
        if rtl:
            _set_rtl(p)
    return tb


def _set_rtl(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set("rtl", "1")


def accent_bar(slide, rect: Rect, color: str = T.COPPER):
    return box(slide, Rect(rect.x, rect.y, rect.w, T.ACCENT_BAR), fill=color)


def bullets_to_paras(items, *, size=T.BODY_SIZE, color=T.INK):
    """Normalize spec bullet items (strings or {text, level}) to paragraph dicts."""
    out = []
    for it in items:
        if isinstance(it, str):
            it = {"text": it}
        out.append({
            "text": it["text"],
            "level": it.get("level", 0),
            "bullet": True,
            "size": size,
            "color": color,
            "bold": it.get("bold", False),
            "space_before": 4,
        })
    return out
