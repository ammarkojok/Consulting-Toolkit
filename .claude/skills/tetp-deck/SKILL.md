---
name: tetp-deck
description: Build TETP-branded PowerPoint decks (.pptx) from declarative YAML specs on the canonical TETP slide master. Use when the user asks for a TETP presentation, status deck, readout, proposal, or any PowerPoint on the TETP template, in English or Arabic.
---

# TETP Deck Builder

Generate finished, brand-safe `.pptx` decks on the canonical TETP template. You
never design slides freehand: you classify content, pick patterns, write a Deck
Spec, build, **look at the rendered slides**, and fix. The renderer guarantees
template fidelity; your job is content judgment.

## Workflow

1. **Classify the deck intent** (status readout, proposal, briefing, …) and draft
   the storyline: one action title per slide — a short sentence stating the
   insight, never a topic label.
2. **Pick a pattern per slide** using the decision table in
   [references/PATTERNS.md](references/PATTERNS.md). Read it before authoring.
   Consult `assets/manifest.json` only when you need placeholder geometry or
   budgets beyond what PATTERNS.md states.
3. **Write the Deck Spec** (YAML): see [schema/deck-spec.schema.json](schema/deck-spec.schema.json)
   and the worked example in [examples/example-deck.yaml](examples/example-deck.yaml).
   The spec contains content only — no coordinates, fonts, or hex colors.
4. **Build and render:**

   ```bash
   python3 scripts/build_deck.py SPEC.yaml --out OUT.pptx --render
   ```

   Spec errors (overflow, bad fields, missing AR layout) fail the build with
   line-level messages. Fix the spec, never the renderer, unless the renderer is
   actually wrong.
5. **Look at every preview PNG** in `OUT-preview/` and check: no clipped or
   overlapping text, no wrapped KPI values, balanced columns, readable chart
   labels, chrome intact (logo, Restricted Access, page number, source line).
   Fix the spec and rebuild until clean. Never deliver unrendered decks.
6. **Deliver** the `.pptx` (and `.ppttc` if think-cell charts were flagged),
   stating what was validated and any placeholder content the user must replace.

## Library prototypes (clone-and-fill)

For maximum design fidelity, prefer a library prototype over a parametric
pattern when one matches the content. Prototypes are fully designed slides
from real TETP decks, indexed in `assets/library.json` — each entry has a
`use` description; read that file to choose. The library grows only from
TETP-designed slides supplied by the user; do not invent prototypes or import
outside designs. The builder clones the prototype verbatim — every layer,
icon, and gradient survives — and replaces only the text slots you address by
shape name:

```bash
python3 scripts/list_library_slots.py                    # available prototypes
python3 scripts/list_library_slots.py delivery_tracker   # its slots
```

Before filling a prototype, open its annotated slot map: `references/slotmaps/<name>.png`
shows each slot's marker in position and `references/slotmaps/SLOTMAPS.md` maps
markers to shape names and original text. Regenerate with
`python3 scripts/build_slot_maps.py` after library changes (output is
gitignored; rebuild it in sessions that have the library files).

```yaml
- type: library
  prototype: delivery_tracker
  slots:
    "Title 1": "New action title"
    "Rectangle: Rounded Corners 1126": "New row item"
```

Slot values may be dicts — `{text, color: RRGGBB, size: pt, bold: true}` — to
override the run style a skeleton slot inherits (some skeletons carry white
text in empty runs; the render check exposes this as washed-out text).

Rules: fill or verify EVERY content slot — unfilled slots keep the prototype's
original text, and the render check exists to catch exactly that. Slots are
text-only for now (status-dot colors stay as in the prototype; pick a prototype
whose dots match, or note it for manual touch-up). `library.pptx` contains
Restricted Access material and is gitignored — sessions without it must fall
back to parametric patterns.

## Hard rules

- Decks are built only from `assets/TETP-canonical.pptx` via `build_deck.py`.
  Never open the template and place shapes manually; never use another template.
- Never redraw, recolor, or re-typeset TETP chrome: logos, Restricted Access
  marker, page numbers, sidebar bands.
- All styling comes from tokens (`scripts/tetpdeck/tokens.py`). If a design needs
  something tokens can't express, extend tokens/patterns deliberately — never
  inline one-off styling.
- Action titles on every content slide. Body copy left-aligned (right-aligned
  for Arabic). Minimum text size 9 pt.
- Charts are native PowerPoint charts, no gridlines, no chart titles — the slide
  title carries the conclusion. Set `thinkcell: true` on a chart only when the
  user wants a think-cell handoff (`.ppttc`); say explicitly that the pipeline
  cannot visually verify think-cell output (ADR-0002).
- `language: ar` mirrors geometry automatically but only Cover, Agenda,
  Vertical, Horizontal, and End Page exist in Arabic; the builder rejects the
  rest (ADR-0003). Arabic output is currently unvalidated — say so when
  delivering.
- Label placeholder or synthetic content plainly; never ship it as fact.
- Quote YAML values that look like booleans or contain `:`/`?` (e.g. `"Yes"`,
  `"On plan?"`).

## Maintenance

- Template changes go into `scripts/build_canonical_template.py`, then re-run it
  and `scripts/extract_manifest.py`, and re-render `examples/example-deck.yaml`
  to confirm nothing regressed.
- Project glossary lives in repo-root `CONTEXT.md`; decisions in `docs/adr/`.
  The thmx "Table" layout is banned (leftover Kearney chrome) — see ADR-0001.
