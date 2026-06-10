# 0002 — Charts default to native PowerPoint; think-cell available per chart via .ppttc

## Status

Accepted (2026-06-10)

## Context

The TETP template embeds think-cell data objects, and think-cell is the house
charting tool for analysts. think-cell cannot run headlessly — it is a licensed
PowerPoint plugin — so a pipeline can neither generate finished think-cell charts
nor visually verify them in the render-QA loop. think-cell does offer an
automation format (`.ppttc`: JSON pairing a template's named chart frames with
data) that generates real think-cell charts when opened on a machine with the
plugin installed.

## Decision

The deck renderer builds native PowerPoint charts (python-pptx chart parts) styled
from theme tokens, so every deck comes out complete and visually QA-able. Any chart
in a Deck Spec may set `thinkcell: true`, which additionally emits a companion
`.ppttc` file so an analyst can regenerate that chart as a native think-cell object
in one step.

## Consequences

- Generated decks are always finished and verifiable; nothing ships that the QA
  loop couldn't see.
- Analysts keep a one-step path to real think-cell charts (waterfall, Mekko, Gantt)
  where the native chart vocabulary falls short.
- The `.ppttc` output is unverifiable by the pipeline by nature; it is an explicit
  handoff artifact, never the default.
