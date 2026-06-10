# Pattern selection and slide-type reference

## Decision table — content shape → slide type

| Your content is… | Use type | Notes |
|---|---|---|
| Deck opening | `cover` | Title ≤ 90 chars, subtitle, report_type, date |
| Section/chapter break | `agenda` | Auto-numbered chapter heading |
| 2–5 headline numbers | `kpis` | Value + label, optional delta |
| 2–6 parallel ideas/themes | `cards` | Heading + 1–3 sentence body each |
| Two things contrasted | `comparison` | Recommended option goes `left` (navy) |
| A sequence of stages | `process` | 2–6 steps, chevron flow |
| Events over time | `timeline` | 2–8 milestones, `done: true` fills the dot |
| One chart + its meaning | `chart` | Add 2–4 `takeaways`; they render as a sidebar |
| Structured records | `table` | ≤ 6 columns, ≤ 9 rows; quote "Yes"/"No" |
| Two labeled lists | `two_columns` | Native layout; heading + body per side |
| Narrative/bullet argument | `bullets` | Use `level: 1` for support lines; `two_column: true` if > 8 items |
| Deck closing | `end` | No content fields |

When nothing fits, compose with `cards` (the most flexible) or ask the user —
do not invent freehand layouts.

## Canvas choice (pattern slides)

Default canvases are sensible; override with `canvas:` only when needed.

| Canvas | Layout | Best for | Title style |
|---|---|---|---|
| `vertical` (default for cards/comparison/kpis/chart) | TETP EN Vertical | Taller content, sidebar title | ≤ 150 chars, narrow column |
| `horizontal` (default for process/timeline/table/bullets) | TETP EN Horizontal | Wide flows and tables | ≤ 120 chars, full width |
| `wide` | EN Wide | Full-bleed emphasis band | ≤ 90 chars |
| `split` | EN Split | Half text, half visual | ≤ 110 chars |

## Common slide fields

Every content slide takes: `title` (required), `subtitle`, `source`, `notes`.
`source` should cite real provenance or say "(illustrative)".

## Capacity budgets (lint-enforced)

- Card body: ~350 chars at 3 columns; fewer cards = more room.
- Comparison bullets: ≤ 700 chars per panel.
- KPI values: keep ≤ 10 chars; longer values auto-shrink.
- Table: header + 9 rows max on `horizontal`.
- Bullets: ~14 lines per column.

## Charts

```yaml
chart:
  kind: bar | hbar | stacked_bar | stacked_bar_100 | line | pie | doughnut | scatter
  categories: [Q1, Q2, Q3]
  series:
    - name: Actual
      values: [10, 20, 30]
      color: 3A3B68        # optional; defaults to the token series palette
  data_labels: true         # optional
  thinkcell: true           # optional; also emits a .ppttc handoff (ADR-0002)
```

Scatter series use `values: [[x, y], ...]`. Never request gridlines, 3D, dual
axes, or chart titles — the renderer won't draw them and the brand forbids them.

## Arabic decks

Set `deck.language: ar`. Allowed types: cover, agenda, end, and pattern slides
on `vertical`/`horizontal` canvases. Geometry mirrors automatically; paragraphs
are marked RTL. Arabic output is engine-supported but not yet validated — flag
this in delivery.
