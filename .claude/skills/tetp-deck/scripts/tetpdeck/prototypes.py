"""Kearney-design prototypes, TETP-colored.

These rebuild the strongest slide designs from the Kearney HTML slide system as
native PPTX prototypes on the TETP canonical template. Composition is theirs;
every color/font comes from tetpdeck.tokens (never Kearney purple). Shapes that
authors should refill carry deliberate snake_case names — they are the slots.

Run scripts/build_library.py to regenerate assets/library-tetp.pptx.
"""

from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

from . import tokens as T
from .draw import box, cell, chip, gradient_box, text
from .geometry import Rect, canvas, columns, hsplit, rows, vsplit


def _slide(prs, layouts, layout_name, title, subtitle=None, source=None):
    slide = prs.slides.add_slide(layouts[layout_name])
    for ph in list(slide.placeholders):
        idx = ph.placeholder_format.idx
        if idx == 0:
            ph.text = title
        elif idx == 23 and subtitle:
            ph.text = subtitle
        elif idx == 20 and subtitle:
            ph.text = subtitle
        elif idx == 22 and source:
            ph.text = source
        else:
            ph._element.getparent().remove(ph._element)
    return slide


def exec_summary(prs, layouts):
    """Kearney 'summary-americas': one calm box, bold-lead bullets."""
    s = _slide(prs, layouts, "TETP EN Horizontal",
               "Executive summary: the program is ahead on reach and behind on depth",
               subtitle="Executive summary", source="Source: TETP (illustrative)")
    cv = canvas("horizontal")
    bullets = [
        "Activation is no longer the constraint: adoption is ahead of target across all academies.",
        "Depth is the gap: usage per learner trails the global benchmark by 15 percent.",
        "Usage is polarized: a strong power-user core coexists with a long light-usage tail.",
        "Costs are under control: run-rate savings are 12 percent ahead of plan.",
        "Management focus: move light users into recurring work through targeted enablement.",
    ]
    panel_h = min(cv.h - 0.3, 0.55 * len(bullets) + 0.6)
    panel = Rect(cv.x, cv.y + 0.1, cv.w, panel_h)
    cell(s, panel)
    tb = text(s, panel.inset(0.3, 0.25),
              [{"text": b, "bullet": True, "size": Pt(12), "space_before": 10}
               for b in bullets])
    tb.name = "summary_bullets"
    return s


def scqa_summary(prs, layouts):
    """Kearney SCQA panels: Situation / Complication / Answer."""
    s = _slide(prs, layouts, "TETP EN Vertical",
               "Situation, complication, answer", subtitle="Executive summary",
               source="Source: TETP (illustrative)")
    cv = canvas("vertical")
    area = Rect(cv.x, cv.y, cv.w, min(cv.h, 3.6))
    labels = ["Situation", "Complication", "Answer"]
    copies = [
        "The program has an active portfolio of initiatives across all academies.",
        "The initiatives do not share a common value logic or decision cadence.",
        "Create one program office and sequence three integrated moves this quarter.",
    ]
    for i, (rect, label, copy) in enumerate(
            zip(columns(area, 3, T.GAP), labels, copies)):
        cell(s, rect, fill=T.PANEL, border=None)
        bar = box(s, Rect(rect.x, rect.y, 0.045, rect.h),
                  fill=T.NAVY if i < 2 else T.COPPER)
        lbl = text(s, Rect(rect.x + 0.18, rect.y + 0.2, rect.w - 0.36, 0.3),
                   [{"text": label.upper(), "size": Pt(10), "bold": True,
                     "color": T.COPPER_DEEP if i == 2 else T.MUTED}])
        lbl.name = f"scqa_label_{i + 1}"
        body = text(s, Rect(rect.x + 0.18, rect.y + 0.55, rect.w - 0.36,
                            rect.h - 0.75),
                    [{"text": copy, "size": Pt(11.5)}])
        body.name = f"scqa_body_{i + 1}"
    return s


def divider(prs, layouts):
    """Kearney divider: oversized section number, display title, accent mark."""
    s = _slide(prs, layouts, "Title Only", "Section divider")
    cv = canvas("vertical")
    num = text(s, Rect(cv.x + 0.2, cv.y + 0.6, 2.6, 1.9),
               [{"text": "01", "size": Pt(96), "bold": True, "color": T.COPPER}])
    num.name = "divider_number"
    title = text(s, Rect(cv.x + 0.25, cv.y + 2.7, cv.w - 1.2, 1.4),
                 [{"text": "Context and value at stake", "size": Pt(30),
                   "bold": True, "color": T.NAVY}])
    title.name = "divider_title"
    box(s, Rect(cv.x + 0.28, cv.y + 4.15, 0.9, 0.05), fill=T.COPPER)
    sub = text(s, Rect(cv.x + 0.25, cv.y + 4.4, cv.w - 1.6, 0.9),
               [{"text": "Use divider slides to reset attention and signal the "
                         "next mode of discussion.", "size": Pt(12),
                 "color": T.MUTED}])
    sub.name = "divider_subtitle"
    return s


def kpi_hero(prs, layouts):
    """Kearney KPI slide: one hero number, two supporting KPIs, interpretation."""
    s = _slide(prs, layouts, "TETP EN Horizontal",
               "The first wave should target the few metrics that signal real change",
               subtitle="KPI evidence", source="Source: TETP (illustrative)")
    cv = canvas("horizontal")
    area = Rect(cv.x, cv.y + 0.05, cv.w, cv.h - 0.25)
    hero_rect, right = hsplit(area, [0.34, 0.66], T.GAP + 0.04)
    gradient_box(s, hero_rect, "navy")
    hl = text(s, Rect(hero_rect.x + 0.25, hero_rect.y + 0.35,
                      hero_rect.w - 0.5, 0.3),
              [{"text": "VALUE AT STAKE", "size": Pt(10), "bold": True,
                "color": T.WHITE}])
    hl.name = "hero_label"
    hv = text(s, Rect(hero_rect.x + 0.25, hero_rect.y + 0.8,
                      hero_rect.w - 0.5, 1.6),
              [{"text": "340", "size": Pt(72), "bold": True, "color": T.WHITE}])
    hv.name = "hero_value"
    hc = text(s, Rect(hero_rect.x + 0.25, hero_rect.y + 2.5,
                      hero_rect.w - 0.5, hero_rect.h - 2.7),
              [{"text": "bps margin upside across offer, service, and working "
                        "capital moves.", "size": Pt(12), "color": T.WHITE}])
    hc.name = "hero_caption"

    kpi_row, interp = vsplit(right, [0.62, 0.38], T.GAP)
    for i, rect in enumerate(columns(kpi_row, 2, T.GAP), start=1):
        cell(s, rect)
        box(s, Rect(rect.x, rect.y, rect.w, 0.045), fill=T.COPPER)
        v = text(s, Rect(rect.x + 0.2, rect.y + 0.25, rect.w - 0.4, 0.9),
                 [{"text": ["42%", "11"][i - 1], "size": Pt(40), "bold": True,
                   "color": T.NAVY}])
        v.name = f"kpi{i}_value"
        l = text(s, Rect(rect.x + 0.2, rect.y + 1.25, rect.w - 0.4, 0.9),
                 [{"text": ["of leakage tied to exceptions, overrides, and manual approvals",
                            "days average delay between decision and field execution"][i - 1],
                   "size": Pt(11)}])
        l.name = f"kpi{i}_label"
        d = text(s, Rect(rect.x + 0.2, rect.y + rect.h - 0.5, rect.w - 0.4, 0.35),
                 [{"text": ["Highest control opportunity", "Cadence gap"][i - 1],
                   "size": Pt(9.5), "bold": True, "color": T.COPPER_DEEP}])
        d.name = f"kpi{i}_delta"
    cell(s, interp, fill=T.PANEL, border=None)
    il = text(s, Rect(interp.x + 0.2, interp.y + 0.15, interp.w - 0.4, 0.3),
              [{"text": "EVIDENCE INTERPRETATION", "size": Pt(9.5), "bold": True,
                "color": T.MUTED}])
    il.name = "interp_label"
    ib = text(s, Rect(interp.x + 0.2, interp.y + 0.5, interp.w - 0.4,
                      interp.h - 0.65),
              [{"text": "Supporting numbers explain the mechanism behind the "
                        "headline rather than compete with it.", "size": Pt(11)}])
    ib.name = "interp_body"
    return s


def next_steps(prs, layouts):
    """Kearney next-steps visual: numbered action rows with owner/due meta."""
    s = _slide(prs, layouts, "TETP EN Horizontal",
               "Immediate actions before the next steering committee",
               subtitle="Next steps", source="Source: TETP (illustrative)")
    cv = canvas("horizontal")
    actions = [
        ("Choose priority workflows",
         "Pick the four repeated tasks where better context changes output quality.",
         "Program lead", "Friday"),
        ("Create reusable assets",
         "Package prompts, source files, examples, and quality checks per workflow.",
         "Practice leads", "Next week"),
        ("Name workflow owners",
         "Assign one person to maintain each asset and review adoption signals.",
         "Champion network", "Week 2"),
        ("Schedule the follow-up readout",
         "Re-run the analysis after four weeks and present deltas to the committee.",
         "Program office", "Week 4"),
    ]
    area = Rect(cv.x, cv.y + 0.05, cv.w, cv.h - 0.2)
    for i, (rect, (title, copy, owner, due)) in enumerate(
            zip(rows(area, len(actions), T.GAP), actions), start=1):
        cell(s, rect)
        d = 0.42
        box(s, Rect(rect.x + 0.18, rect.y + (rect.h - d) / 2, d, d),
            fill=T.NAVY, shape=MSO_SHAPE.OVAL)
        text(s, Rect(rect.x + 0.18, rect.y + (rect.h - d) / 2, d, d),
             [{"text": str(i), "size": Pt(16), "bold": True, "color": T.WHITE,
               "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
        tt = text(s, Rect(rect.x + 0.85, rect.y + 0.12, rect.w * 0.55, 0.35),
                  [{"text": title, "size": Pt(12.5), "bold": True,
                    "color": T.NAVY}])
        tt.name = f"action{i}_title"
        tc = text(s, Rect(rect.x + 0.85, rect.y + 0.48, rect.w * 0.55,
                          rect.h - 0.55),
                  [{"text": copy, "size": Pt(10.5)}])
        tc.name = f"action{i}_copy"
        meta_x = rect.x + rect.w * 0.72
        for j, (k, v) in enumerate((("OWNER", owner), ("DUE", due))):
            mx = meta_x + j * (rect.w * 0.14)
            text(s, Rect(mx, rect.y + 0.16, rect.w * 0.13, 0.25),
                 [{"text": k, "size": Pt(8.5), "bold": True, "color": T.MUTED}])
            mv = text(s, Rect(mx, rect.y + 0.42, rect.w * 0.13, rect.h - 0.5),
                      [{"text": v, "size": Pt(10.5), "bold": True}])
            mv.name = f"action{i}_{k.lower()}"
    return s


PROTOTYPES = {
    "exec_summary": (exec_summary,
                     "Opening executive summary: one calm box of bold-lead bullets"),
    "scqa_summary": (scqa_summary,
                     "Situation / Complication / Answer three-panel summary"),
    "divider": (divider, "Section divider: oversized number + display title"),
    "kpi_hero": (kpi_hero,
                 "One hero number with two supporting KPIs and interpretation"),
    "next_steps": (next_steps,
                   "Numbered action rows with owner and due-date meta"),
}
