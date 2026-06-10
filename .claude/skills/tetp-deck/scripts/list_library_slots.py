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
    for name, entry in sorted(idx.items()):
        if isinstance(entry, int):
            entry = {"file": "library.pptx", "slide": entry, "use": ""}
        print(f"{name}  ({entry['file']} slide {entry['slide']})  {entry.get('use', '')}")
else:
    for shape, text in list_slots(sys.argv[1]):
        print(f"{shape!r}: {text[:90]}")
