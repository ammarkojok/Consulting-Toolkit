# 0001 — One canonical template file, merged from the .pptx and .thmx sources

## Status

Accepted (2026-06-10)

## Context

Two TETP brand artifacts exist: a slide-master `.pptx` (2 masters — EN and AR — with
5 layouts each: Cover, End Page, Agenda, Vertical, Horizontal) and a `.thmx` theme
(9 near-duplicate masters, 138 layouts of ~14 distinct names, including content
layouts the `.pptx` lacks: Empty, Wide, Split, Two Columns, Bios, Case Study,
Horizontal/Standard Content, Table; English only; 8 of 9 color schemes identical to
the `.pptx` theme, one divergent "Takamul-Color").

A renderer that accepts arbitrary template files inherits this duplication, and
PowerPoint's behavior with many same-named masters is a known source of
wrong-layout and wrong-theme errors. python-pptx also cannot open a `.thmx` as a
document base.

## Decision

Build one canonical `.pptx`: the slide-master `.pptx` as the base (keeping EN and AR
masters and chrome), with the distinct content layouts from the `.thmx` imported,
deduplicated to one master per language, on the navy/copper "Custom 10" theme.
The deck renderer only ever opens this file. The original uploads are kept as
source artifacts but are never rendered from.

## Consequences

- Every generated deck inherits exactly one master per language and one theme;
  the wrong-master class of errors is structurally impossible.
- The "Takamul-Color" scheme is dropped; reintroducing it later means adding a
  second canonical theme variant deliberately, not by accident.
- Layout changes (new or edited layouts) are made in the canonical file and
  re-manifested — the upstream `.pptx`/`.thmx` are not edited.
