# CONTEXT.md

Domain glossary for the Consulting-Toolkit. Terms only — no implementation details.

## Terms

**Canonical Template**
The single `.pptx` file the deck renderer always opens as its base. Merged from the
TETP slide-master deck (EN + AR masters, chrome) and the distinct content layouts of
the TETP `.thmx` theme, deduplicated to one master per language on the navy/copper
("Custom 10") theme. No deck is ever built from any other template file.

**Layout**
A native PowerPoint slide layout inside the Canonical Template, identified by its
authored name (e.g. `EN Cover`, `EN Two Columns`, `EN Split`, `AR Vertical`).
Layouts own chrome and placeholders; decks instantiate layouts, never blank slides.

**Chrome**
The fixed furniture a Layout contributes to every slide built on it: logos,
"Restricted Access" marker, page number, source line, background bands. Chrome is
never redrawn or repositioned by deck authors.

**Manifest**
The machine-readable description of the Canonical Template: every Layout, its
placeholders (index, type, position, size, character budget), theme colors, and
fonts. The Manifest — not the template binary — is what an authoring LLM reads.

**Deck Spec**
The declarative document an authoring LLM writes to describe a deck: per slide, a
Layout (or Pattern) name plus content for its slots. Contains no coordinates, fonts,
or colors.

**Pattern**
A named, parametric slide arrangement defined by this toolkit (e.g. card row, KPI
banner, two-panel comparison) and rendered onto a content Layout's canvas by
deterministic code. Patterns expose content slots; their geometry, spacing, and
colors are fixed by the toolkit, not chosen per-deck.
