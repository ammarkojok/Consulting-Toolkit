# 0003 — The geometry engine is RTL-aware from day one; v1 ships EN patterns only

## Status

Accepted (2026-06-10)

## Context

The canonical template carries both English (LTR) and Arabic (RTL, mirrored chrome)
masters. Mirroring is geometric: every pattern's column order, alignment, and anchor
positions flip. Retrofitting RTL into a finished LTR-only pattern library means
reworking every pattern.

## Consequences

- Every pattern computes its geometry through the grid engine with a direction
  parameter; no pattern hardcodes left/right positions directly.
- v1 validates English output only — Arabic rendering exists in the engine but is
  unvalidated until Arabic copy testing is done deliberately.
- A future reader will find direction plumbed through patterns that only ever run
  LTR today; this is intentional, not dead code.
