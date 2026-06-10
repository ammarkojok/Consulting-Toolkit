#!/usr/bin/env python3
"""List the fillable text slots of a library prototype.

Usage:
    python3 list_library_slots.py                 # list prototypes
    python3 list_library_slots.py PROTOTYPE_NAME  # list its slots
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from tetpdeck.library import lib_index, list_slots  # noqa: E402

if len(sys.argv) < 2:
    idx = lib_index()
    if not idx:
        print("library is empty — add prototypes to assets/library.pptx "
              "and name them in assets/library.json")
    for name, num in sorted(idx.items(), key=lambda kv: kv[1]):
        print(f"{name}  (library slide {num})")
else:
    for shape, text in list_slots(sys.argv[1]):
        print(f"{shape!r}: {text[:90]}")
