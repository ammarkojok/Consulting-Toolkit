"""Core 8 pattern renderers. Each pattern draws content onto a layout canvas;
geometry comes from geometry.py, styling from tokens.py — patterns decide
arrangement only. Every renderer takes (slide, spec, canvas_rect, rtl)."""

from __future__ import annotations

from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

from . import tokens as T
from .draw import (accent_bar, box, bullets_to_paras, cell, chip, gradient_box,
                   header_bar, status_dot, status_legend, text)
from .geometry import Rect, columns, est_text_height, grid, hsplit, vsplit
from .charts import add_chart


def cards(slide, spec, cv: Rect, rtl: bool):
    items = spec["cards"]
    max_cols = len(items) if len(items) <= 4 else 3
    nrows = 1 if len(items) <= max_cols else 2
    # A single sparse row reads unfinished at full canvas height; cap and
    # let whitespace sit below, like a hand-built consulting slide.
    area = cv if nrows == 2 else Rect(cv.x, cv.y, cv.w, min(cv.h, 3.9))
    cells = grid(area, len(items), max_cols=max_cols, gap=T.GAP, rtl=rtl)
    numbered = spec.get("numbered", True)
    for i, (cell_rect, item) in enumerate(zip(cells, items), start=1):
        cell(slide, cell_rect)
        accent_bar(slide, cell_rect, item.get("accent", T.COPPER))
        inner = cell_rect.inset(T.PAD + 0.04)
        paras = []
        if numbered:
            paras.append({"text": f"{i:02d}", "size": T.SUBHEAD_SIZE, "bold": True,
                          "color": T.COPPER})
        paras.append({"text": item["heading"], "size": T.HEADING_SIZE, "bold": True,
                      "color": T.NAVY, "space_before": 2 if numbered else 0})
        if item.get("body"):
            body = item["body"] if isinstance(item["body"], list) else [item["body"]]
            for b in body:
                paras.append({"text": b if isinstance(b, str) else b["text"],
                              "size": T.BODY_SIZE, "space_before": 6})
        text(slide, Rect(inner.x, inner.y + 0.08, inner.w, inner.h - 0.08),
             paras, rtl=rtl)


def comparison(slide, spec, cv: Rect, rtl: bool):
    panels = [spec["left"], spec["right"]]
    schemes = ["navy", "copper"]
    # Size panels to content (plus headroom) rather than the full canvas.
    need = max(
        sum(est_text_height(b if isinstance(b, str) else b["text"],
                            cv.w / 2 - 0.5, T.BODY_SIZE.pt) + 0.06
            for b in p.get("bullets", []))
        for p in panels
    )
    area = Rect(cv.x, cv.y, cv.w, min(cv.h, T.HEADER_BAR + need + 0.65))
    for rect, panel, scheme in zip(columns(area, 2, T.GAP, rtl), panels, schemes):
        bar = header_bar(slide, rect, panel["heading"], scheme, rtl)
        body = Rect(rect.x, rect.y + bar.h + 0.04, rect.w, rect.h - bar.h - 0.04)
        cell(slide, body)
        text(slide, body.inset(T.PAD),
             bullets_to_paras(panel.get("bullets", [])), rtl=rtl)


def kpis(slide, spec, cv: Rect, rtl: bool):
    items = spec["kpis"]
    band_h = min(2.45, cv.h)
    band = Rect(cv.x, cv.y + (cv.h - band_h) / 2 if spec.get("centered", True) else cv.y,
                cv.w, band_h)
    col_w = (band.w - T.GAP * (len(items) - 1)) / len(items)
    for rect, item in zip(columns(band, len(items), T.GAP, rtl), items):
        cell(slide, rect)
        accent_bar(slide, rect, item.get("accent", T.NAVY))
        # Step the value size down when it would wrap the column.
        value = str(item["value"])
        vsize = T.KPI_VALUE_SIZE
        from pptx.util import Pt
        while len(value) * 0.0095 * vsize.pt > col_w - 2 * T.PAD and vsize.pt > 20:
            vsize = Pt(vsize.pt - 4)
        paras = [{"text": value, "size": vsize, "bold": True,
                  "color": T.NAVY, "align": PP_ALIGN.CENTER}]
        if item.get("delta"):
            paras.append({"text": item["delta"], "size": T.SMALL_SIZE,
                          "color": T.MUTED, "align": PP_ALIGN.CENTER})
        paras.append({"text": item["label"], "size": T.KPI_LABEL_SIZE,
                      "color": T.INK, "align": PP_ALIGN.CENTER, "space_before": 4})
        text(slide, rect.inset(T.PAD), paras, anchor=MSO_ANCHOR.MIDDLE, rtl=rtl)


def process(slide, spec, cv: Rect, rtl: bool):
    from pptx.enum.shapes import MSO_SHAPE
    steps = spec["steps"]
    chev_h = 0.55
    body_h = cv.h - chev_h - 0.12
    cols = columns(Rect(cv.x, cv.y, cv.w, cv.h), len(steps), T.GAP, rtl)
    schemes = ["navy", "copper", "indigo"]
    for i, (rect, step) in enumerate(zip(cols, steps)):
        shape = MSO_SHAPE.PENTAGON if i == 0 and not rtl else MSO_SHAPE.CHEVRON
        gradient_box(slide, Rect(rect.x, rect.y, rect.w, chev_h),
                     schemes[i % len(schemes)], shape=shape)
        text(slide, Rect(rect.x + 0.12, rect.y, rect.w - 0.24, chev_h),
             [{"text": step["heading"], "size": T.BODY_SIZE, "bold": True,
               "color": T.WHITE, "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE, rtl=rtl)
        if step.get("body"):
            body = step["body"] if isinstance(step["body"], list) else [step["body"]]
            text(slide, Rect(rect.x, rect.y + chev_h + 0.12, rect.w, body_h),
                 bullets_to_paras(body), rtl=rtl)


def timeline(slide, spec, cv: Rect, rtl: bool):
    miles = spec["milestones"]
    line_y = cv.y + cv.h * 0.42
    box(slide, Rect(cv.x + 0.1, line_y, cv.w - 0.2, 0.025), fill=T.GRAY)
    cols = columns(Rect(cv.x, cv.y, cv.w, cv.h), len(miles), 0.08, rtl)
    dot = 0.16
    for i, (rect, m) in enumerate(zip(cols, miles)):
        cx = rect.x + rect.w / 2
        from pptx.enum.shapes import MSO_SHAPE
        done = m.get("done", False)
        box(slide, Rect(cx - dot / 2, line_y - dot / 2 + 0.0125, dot, dot),
            fill=T.NAVY if done else T.WHITE,
            line=None if done else T.NAVY, shape=MSO_SHAPE.OVAL)
        chip_w = min(rect.w - 0.1, max(0.9, len(m["date"]) * 0.085))
        chip(slide, Rect(cx - chip_w / 2, line_y - 0.52, chip_w, 0.26),
             m["date"], fill=T.NAVY if m.get("done") else T.INDIGO, rtl=rtl)
        text(slide, Rect(rect.x + 0.05, line_y + 0.22, rect.w - 0.1, cv.h - (line_y - cv.y) - 0.3),
             [{"text": m["label"], "size": T.BODY_SIZE, "align": PP_ALIGN.CENTER}],
             rtl=rtl)


def chart_takeaways(slide, spec, cv: Rect, rtl: bool):
    if spec.get("takeaways"):
        chart_rect, side = hsplit(cv, [0.66, 0.34], T.GAP, rtl)
        bar = header_bar(slide, side, spec.get("takeaways_heading", "Key takeaways"),
                         "navy", rtl)
        body = Rect(side.x, side.y + bar.h + 0.04, side.w, side.h - bar.h - 0.04)
        cell(slide, body)
        text(slide, body.inset(T.PAD), bullets_to_paras(spec["takeaways"]), rtl=rtl)
    else:
        chart_rect = cv
    cell(slide, chart_rect)
    add_chart(slide, chart_rect.inset(0.18, 0.16), spec["chart"])


def table(slide, spec, cv: Rect, rtl: bool):
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    cols_spec = spec["columns"]
    rows_spec = spec["rows"]
    if rtl:
        cols_spec = list(reversed(cols_spec))
        rows_spec = [list(reversed(r)) for r in rows_spec]
    nrows, ncols = len(rows_spec) + 1, len(cols_spec)
    h = min(cv.h, 0.38 * nrows)
    gf = slide.shapes.add_table(nrows, ncols, Inches(cv.x), Inches(cv.y),
                                Inches(cv.w), Inches(h))
    tbl = gf.table
    tbl.first_row = False
    tbl.horz_banding = False
    for j, name in enumerate(cols_spec):
        cell = tbl.cell(0, j)
        cell.text = str(name)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor.from_string(T.NAVY)
        _style_cell(cell, bold=True, color=T.WHITE, rtl=rtl)
    for i, row in enumerate(rows_spec, start=1):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            # YAML parses bare Yes/No as booleans; render them back as words.
            if val is True:
                val = "Yes"
            elif val is False:
                val = "No"
            cell.text = "" if val is None else str(val)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(
                T.WHITE if i % 2 else "F7F7F7"
            )
            _style_cell(cell, bold=False, color=T.INK, rtl=rtl)


def _style_cell(cell, *, bold: bool, color: str, rtl: bool):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    cell.margin_left = cell.margin_right = Pt(6)
    cell.margin_top = cell.margin_bottom = Pt(3)
    for p in cell.text_frame.paragraphs:
        p.alignment = PP_ALIGN.RIGHT if rtl else PP_ALIGN.LEFT
        for r in p.runs:
            r.font.name = T.FONT
            r.font.size = T.BODY_SIZE
            r.font.bold = bold
            r.font.color.rgb = RGBColor.from_string(color)


def bullets_pattern(slide, spec, cv: Rect, rtl: bool):
    """Bullet hierarchy drawn on a canvas (used when a native body placeholder
    isn't the target). Optionally splits into two balanced columns."""
    items = spec["bullets"]
    paras = bullets_to_paras(items)
    if spec.get("two_column") and len(items) > 5:
        half = (len(paras) + 1) // 2
        for rect, chunk in zip(columns(cv, 2, T.SECTION_GAP, rtl),
                               (paras[:half], paras[half:])):
            text(slide, rect, chunk, rtl=rtl)
    else:
        text(slide, cv, paras, rtl=rtl)


def tracker(slide, spec, cv: Rect, rtl: bool):
    """Status tracker (reference style): category rail, item cells with status
    dots, progress-note column, legend."""
    legend_h = 0.3 if spec.get("legend", True) else 0.0
    area = Rect(cv.x, cv.y, cv.w, cv.h - legend_h - (0.08 if legend_h else 0))
    rail_w = 1.35 if spec.get("category") else 0.0
    grid_x = area.x + rail_w + (T.GAP if rail_w else 0)
    grid_w = area.w - rail_w - (T.GAP if rail_w else 0)
    item_w = grid_w * 0.34
    note_w = grid_w - item_w - T.GAP

    if rtl:
        item_x = area.x + area.w - rail_w - (T.GAP if rail_w else 0) - item_w
        note_x = item_x - T.GAP - note_w
        rail_x = area.x + area.w - rail_w
    else:
        item_x, note_x, rail_x = grid_x, grid_x + item_w + T.GAP, area.x

    bar_y = area.y
    header_bar(slide, Rect(item_x, bar_y, item_w, 0), spec.get("items_heading", "Output"), "navy", rtl)
    header_bar(slide, Rect(note_x, bar_y, note_w, 0), spec.get("notes_heading", "Progress Update"), "navy", rtl)

    rows_area = Rect(area.x, bar_y + T.HEADER_BAR + 0.06, area.w,
                     area.h - T.HEADER_BAR - 0.06)
    if rail_w:
        box(slide, Rect(rail_x, rows_area.y, rail_w, rows_area.h), fill=T.PANEL_TAN)
        text(slide, Rect(rail_x + 0.08, rows_area.y, rail_w - 0.16, rows_area.h),
             [{"text": spec["category"], "size": T.BODY_SIZE, "bold": True,
               "color": T.COPPER_DEEP, "align": PP_ALIGN.CENTER}],
             anchor=MSO_ANCHOR.MIDDLE, rtl=rtl)

    items = spec["rows"]
    row_h = (rows_area.h - T.GAP * (len(items) - 1)) / len(items)
    for i, it in enumerate(items):
        y = rows_area.y + i * (row_h + T.GAP)
        icell = Rect(item_x, y, item_w, row_h)
        cell(slide, icell, fill="F7F7FA", border=None)
        dot_x = icell.x + (0.2 if rtl else icell.w - 0.2)
        text(slide, Rect(icell.x + (0.35 if rtl else 0.1), icell.y,
                         icell.w - 0.45, row_h),
             [{"text": it["item"], "size": T.BODY_SIZE, "bold": True,
               "color": T.NAVY}], anchor=MSO_ANCHOR.MIDDLE, rtl=rtl)
        status_dot(slide, dot_x, y + row_h / 2, it.get("status", "not_started"))
        ncell = Rect(note_x, y, note_w, row_h)
        cell(slide, ncell)
        if it.get("note"):
            text(slide, ncell.inset(0.1, 0.06),
                 [{"text": "–  " + it["note"], "size": T.BODY_SIZE}],
                 anchor=MSO_ANCHOR.MIDDLE, rtl=rtl)

    if legend_h:
        order = ["not_started", "on_track", "completed", "risk", "delayed"]
        used = [s for s in order if any(r.get("status") == s for r in items)] or order[:3]
        status_legend(slide, Rect(grid_x, area.y + area.h + 0.08,
                                  min(grid_w, 1.35 * len(used)), legend_h),
                      used, rtl)


PATTERNS = {
    "cards": cards,
    "comparison": comparison,
    "kpis": kpis,
    "process": process,
    "timeline": timeline,
    "chart": chart_takeaways,
    "table": table,
    "bullets": bullets_pattern,
    "tracker": tracker,
}
