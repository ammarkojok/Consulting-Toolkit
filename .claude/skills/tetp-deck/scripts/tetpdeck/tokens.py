"""TETP design tokens — the only place colors, fonts, and scales are defined.

Values come from the canonical template's theme ("Custom 10") and master chrome.
Patterns and charts must reference these names; hex literals anywhere else in
tetpdeck are a lint error waiting to happen.
"""

from pptx.util import Pt

# --- palette (theme "Custom 10" + reference-deck house palette) -------------
INK = "1E1E1E"        # dk1 — body text
WHITE = "FFFFFF"      # lt1
NAVY = "3A3B68"       # dk2 — primary brand, headings, emphasis
NAVY_SOFT = "5D5E92"  # gradient partner of NAVY (mined from reference decks)
COPPER = "C4996C"     # lt2 — secondary brand, accents
COPPER_DEEP = "A87F4E"  # gradient partner of COPPER
GRAY = "D2D2D2"       # accent1 — lines, muted
BORDER = "D9D9D9"     # thin cell borders
TAN = "C5A783"        # accent2
PANEL_TAN = "EAE1D6"  # light tan swimlane/rail panels
PANEL = "F4F4F8"      # light neutral panel behind cell groups
MAUVE = "B6AFB0"      # accent3
INDIGO = "66689B"     # accent4 — chart series, soft emphasis
DEEP = "505282"       # accent5 — chart series
TAN2 = "C5A078"       # accent6
CANVAS = "EFEFEF"     # master content-canvas gray (sampled from chrome)
MUTED = "6E6E6E"      # derived mid-gray for captions/sources
GOOD = "4F7B58"       # legacy status green (flat)
WARN = "B0532E"       # legacy status red-brown (flat)

# Status code mined from reference decks (tracker legend).
STATUS = {
    "not_started": "A6A6A6",
    "on_track": "00B050",
    "completed": "002060",
    "risk": "FFC000",
    "delayed": "C00000",
}
STATUS_LABELS = {
    "not_started": "Not Started", "on_track": "On Track",
    "completed": "Completed", "risk": "Risk of Delay", "delayed": "Delayed",
}

# Ordered series palette for charts and category accents.
SERIES = [NAVY, COPPER, INDIGO, TAN, DEEP, MAUVE]

# Gradient recipe for header bars/rails (reference decks): light -> brand, ~60deg.
GRADIENTS = {
    "navy": (NAVY_SOFT, NAVY),
    "copper": (COPPER, COPPER_DEEP),
    "indigo": (INDIGO, DEEP),
}
GRADIENT_ANGLE = 60

# --- typography (all Arial; scale mined from reference decks) ---------------
FONT = "Arial"

TITLE_SIZE = Pt(20)          # full-width titles (Horizontal/Wide)
TITLE_COL_SIZE = Pt(18)      # narrow sidebar titles (Vertical/Split/Title Only)
SUBHEAD_SIZE = Pt(12)
HEADING_SIZE = Pt(12)        # card/panel/column headings
BODY_SIZE = Pt(10.5)
SMALL_SIZE = Pt(9)           # captions, axis labels — never go below MIN_SIZE
SOURCE_SIZE = Pt(9)
KPI_VALUE_SIZE = Pt(32)
KPI_LABEL_SIZE = Pt(10.5)
MIN_SIZE = Pt(9)             # hard floor (~12px)

LINE_SPACING = 1.12

# --- spacing scale (inches) --------------------------------------------------
GAP = 0.14        # gutter between sibling blocks
PAD = 0.12        # inner padding of cards/panels
SECTION_GAP = 0.26
ACCENT_BAR = 0.045  # thickness of card accent top bars
HEADER_BAR = 0.32   # height of gradient header bars
