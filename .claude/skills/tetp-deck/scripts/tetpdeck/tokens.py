"""TETP design tokens — the only place colors, fonts, and scales are defined.

Values come from the canonical template's theme ("Custom 10") and master chrome.
Patterns and charts must reference these names; hex literals anywhere else in
tetpdeck are a lint error waiting to happen.
"""

from pptx.util import Pt

# --- palette (theme "Custom 10") -------------------------------------------
INK = "1E1E1E"        # dk1 — body text
WHITE = "FFFFFF"      # lt1
NAVY = "3A3B68"       # dk2 — primary brand, headings, emphasis
COPPER = "C4996C"     # lt2 — secondary brand, accents
GRAY = "D2D2D2"       # accent1 — lines, muted
TAN = "C5A783"        # accent2
MAUVE = "B6AFB0"      # accent3
INDIGO = "66689B"     # accent4 — chart series, soft emphasis
DEEP = "505282"       # accent5 — chart series
TAN2 = "C5A078"       # accent6
CANVAS = "EFEFEF"     # master content-canvas gray (sampled from chrome)
MUTED = "6E6E6E"      # derived mid-gray for captions/sources
GOOD = "4F7B58"       # status green (flat, restrained)
WARN = "B0532E"       # status red-brown (flat, restrained)

# Ordered series palette for charts and category accents.
SERIES = [NAVY, COPPER, INDIGO, TAN, DEEP, MAUVE]

# --- typography (all Arial, per theme major+minor fonts) --------------------
FONT = "Arial"

TITLE_SIZE = Pt(20)          # full-width titles (Horizontal/Wide)
TITLE_COL_SIZE = Pt(18)      # narrow sidebar titles (Vertical/Split/Title Only)
SUBHEAD_SIZE = Pt(12)
HEADING_SIZE = Pt(13)        # card/panel/column headings
BODY_SIZE = Pt(11)
SMALL_SIZE = Pt(9.5)         # captions, axis labels — never go below MIN_SIZE
SOURCE_SIZE = Pt(9.5)
KPI_VALUE_SIZE = Pt(36)
KPI_LABEL_SIZE = Pt(11)
MIN_SIZE = Pt(9)             # hard floor (~12px)

LINE_SPACING = 1.15

# --- spacing scale (inches) --------------------------------------------------
GAP = 0.17        # gutter between sibling blocks
PAD = 0.14        # inner padding of cards/panels
SECTION_GAP = 0.30
ACCENT_BAR = 0.045  # thickness of card accent top bars
