"""Geometry engine: canvases, grid splits, and RTL mirroring.

Every pattern computes its boxes through this module; no pattern hardcodes
left/right positions (see ADR-0003). All values are inches.
"""

from __future__ import annotations

import math
from dataclasses import dataclass

SLIDE_W = 13.333
SLIDE_H = 7.5


@dataclass(frozen=True)
class Rect:
    x: float
    y: float
    w: float
    h: float

    def mirrored(self) -> "Rect":
        return Rect(SLIDE_W - self.x - self.w, self.y, self.w, self.h)

    def inset(self, dx: float, dy: float | None = None) -> "Rect":
        dy = dx if dy is None else dy
        return Rect(self.x + dx, self.y + dy, self.w - 2 * dx, self.h - 2 * dy)


# Free content canvases per layout, measured from the canonical template (EN).
# AR equivalents are exact mirrors (verified against the AR master chrome).
CANVASES = {
    "vertical": Rect(3.75, 0.42, 9.16, 6.25),     # TETP EN Vertical / Title Only
    "horizontal": Rect(0.43, 1.51, 12.49, 5.45),  # TETP EN Horizontal body area
    "wide": Rect(0.42, 2.99, 12.49, 3.50),        # EN Wide gray band, below sub-title
    "split": Rect(6.92, 0.42, 6.16, 6.25),        # EN Split right (visual) half
}

# Text zone on EN Split's left half, below the wide title.
SPLIT_TEXT_ZONE = Rect(0.42, 1.60, 5.83, 5.20)


def canvas(name: str, rtl: bool = False) -> Rect:
    r = CANVASES[name]
    return r.mirrored() if rtl else r


def columns(rect: Rect, n: int, gap: float, rtl: bool = False) -> list[Rect]:
    """Split a rect into n equal columns. RTL reverses visual order so that
    the first content item lands where a right-to-left reader starts."""
    w = (rect.w - gap * (n - 1)) / n
    cols = [Rect(rect.x + i * (w + gap), rect.y, w, rect.h) for i in range(n)]
    return list(reversed(cols)) if rtl else cols


def rows(rect: Rect, n: int, gap: float) -> list[Rect]:
    h = (rect.h - gap * (n - 1)) / n
    return [Rect(rect.x, rect.y + i * (h + gap), rect.w, h) for i in range(n)]


def grid(rect: Rect, n: int, max_cols: int, gap: float, rtl: bool = False) -> list[Rect]:
    """Lay n cells into a row-major grid of at most max_cols columns."""
    ncols = min(n, max_cols)
    nrows = math.ceil(n / ncols)
    cells: list[Rect] = []
    for r in rows(rect, nrows, gap):
        cells.extend(columns(r, ncols, gap, rtl))
    return cells[:n] if not rtl else _rtl_rowmajor(cells, n, ncols)


def _rtl_rowmajor(cells: list[Rect], n: int, ncols: int) -> list[Rect]:
    # columns() already reversed each row; just truncate to n.
    return cells[:n]


def hsplit(rect: Rect, fractions: list[float], gap: float, rtl: bool = False) -> list[Rect]:
    """Split horizontally by fractions (must sum to ~1)."""
    avail = rect.w - gap * (len(fractions) - 1)
    out, x = [], rect.x
    for f in fractions:
        w = avail * f
        out.append(Rect(x, rect.y, w, rect.h))
        x += w + gap
    return list(reversed(out)) if rtl else out


def vsplit(rect: Rect, fractions: list[float], gap: float) -> list[Rect]:
    avail = rect.h - gap * (len(fractions) - 1)
    out, y = [], rect.y
    for f in fractions:
        h = avail * f
        out.append(Rect(rect.x, y, rect.w, h))
        y += h + gap
    return out


# --- overflow estimation ------------------------------------------------------

def est_text_height(text: str, width_in: float, font_pt: float,
                    spacing: float = 1.15) -> float:
    """Conservative estimate of rendered text height in inches."""
    char_w = 0.0075 * font_pt  # avg Arial char width, slightly conservative
    cpl = max(1, int(width_in / char_w))
    lines = 0
    for raw in text.split("\n"):
        lines += max(1, math.ceil(len(raw) / cpl))
    return lines * (font_pt / 72.0) * (spacing + 0.1)
