#!/usr/bin/env python3
"""Build a TETP deck from a YAML Deck Spec.

Usage:
    python3 build_deck.py SPEC.yaml [--out OUT.pptx] [--render]

--render converts the built deck to per-slide PNGs (LibreOffice + pdftoppm)
next to the output file, in <deck>-preview/. ALWAYS render and look at every
PNG before delivering a deck.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from tetpdeck.builder import DeckBuilder, SpecError  # noqa: E402


def render_previews(pptx: Path) -> Path:
    out_dir = pptx.parent / f"{pptx.stem}-preview"
    out_dir.mkdir(exist_ok=True)
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf",
         "--outdir", str(out_dir), str(pptx)],
        check=True, capture_output=True,
    )
    pdf = out_dir / f"{pptx.stem}.pdf"
    subprocess.run(
        ["pdftoppm", "-png", "-r", "80", str(pdf), str(out_dir / "slide")],
        check=True,
    )
    return out_dir


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("spec", type=Path)
    ap.add_argument("--out", type=Path)
    ap.add_argument("--render", action="store_true")
    args = ap.parse_args()

    out = args.out or args.spec.with_suffix(".pptx")
    try:
        builder = DeckBuilder(args.spec)
        builder.build(out)
    except SpecError as e:
        print("SPEC ERRORS — deck not built:\n" + str(e), file=sys.stderr)
        return 1
    print(f"built {out}")

    if args.render:
        out_dir = render_previews(out)
        pages = sorted(out_dir.glob("slide-*.png"))
        print(f"rendered {len(pages)} previews in {out_dir}/")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
