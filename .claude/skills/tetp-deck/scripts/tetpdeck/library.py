"""Slide library: clone-and-fill of fully designed prototype slides.

The parametric patterns reproduce design *recipes*; the library reproduces the
design itself. Prototype slides (hand-built, e.g. taken from real TETP decks)
live in assets/library.pptx with a name index in assets/library.json. A deck
spec uses them via:

    - type: library
      prototype: tracker_status
      slots:
        "Title 1": "New action title"
        "TextBox 12": ["line one", "line two"]

Slots are addressed by shape name (see list_library_slots.py). Cloning is
zip-level surgery on the built deck: the prototype slide part, its media and
embeddings are copied verbatim — every layer, group, gradient, and icon
survives — and only the addressed text runs are replaced. The slide's layout
relationship is remapped by name onto the canonical template's layout so the
deck keeps exactly one master per language (ADR-0001).
"""

from __future__ import annotations

import json
import posixpath
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent.parent
LIB_PPTX = SKILL / "assets" / "library.pptx"
LIB_INDEX = SKILL / "assets" / "library.json"

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
SLIDE_CT = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
SLIDE_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
LAYOUT_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
NOTES_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
FALLBACK_LAYOUT = "TETP EN Horizontal"


def lib_index() -> dict:
    return json.loads(LIB_INDEX.read_text(encoding="utf-8")) if LIB_INDEX.exists() else {}


def resolve_prototype(name: str) -> tuple[Path, int]:
    """Return (pptx_path, slide_number). Index entries are either a bare slide
    number (legacy; lives in library.pptx) or {"file": ..., "slide": ..., "use": ...}."""
    idx = lib_index()
    if name not in idx:
        raise KeyError(
            f"prototype {name!r} not in library index; available: {sorted(idx)}"
        )
    entry = idx[name]
    if isinstance(entry, int):
        return LIB_PPTX, entry
    path = SKILL / "assets" / entry["file"]
    if not path.exists():
        raise FileNotFoundError(
            f"prototype {name!r} needs {entry['file']}, which is not present in "
            f"this checkout (restricted libraries are not committed)"
        )
    return path, entry["slide"]


def list_slots(name: str) -> list[tuple[str, str]]:
    """Return (shape_name, current_text) for every text-bearing shape."""
    lib_path, num = resolve_prototype(name)
    out = []
    with zipfile.ZipFile(lib_path) as z:
        root = etree.fromstring(z.read(f"ppt/slides/slide{num}.xml"))
    for sp in root.iter(f"{{{NS['p']}}}sp"):
        nv = sp.find(f".//{{{NS['p']}}}cNvPr")
        texts = [t.text or "" for t in sp.iter(f"{{{NS['a']}}}t")]
        joined = " | ".join(t for t in texts if t.strip())
        if joined:
            out.append((nv.get("name"), joined))
    return out


def _replace_slot_text(root, shape_name: str, value) -> bool:
    """Replace the text of the named shape. A string collapses the shape to one
    paragraph (keeping its first run's formatting); a list maps to existing
    paragraphs in order (extra paragraphs are removed, extra values appended as
    clones of the last paragraph)."""
    for sp in root.iter(f"{{{NS['p']}}}sp"):
        nv = sp.find(f".//{{{NS['p']}}}cNvPr")
        if nv is None or nv.get("name") != shape_name:
            continue
        tx = sp.find(f".//{{{NS['p']}}}txBody")
        if tx is None:
            return False
        paras = tx.findall(f"{{{NS['a']}}}p")
        values = [value] if isinstance(value, str) else list(value)
        # Trim or grow paragraph list.
        while len(paras) > len(values):
            tx.remove(paras.pop())
        while len(paras) < len(values):
            import copy
            clone = copy.deepcopy(paras[-1])
            tx.append(clone)
            paras.append(clone)
        for p, val in zip(paras, values):
            runs = p.findall(f"{{{NS['a']}}}r")
            if not runs:
                # paragraph had no run (empty) — make one
                r = etree.SubElement(p, f"{{{NS['a']}}}r")
                t = etree.SubElement(r, f"{{{NS['a']}}}t")
                runs = [r]
            for extra in runs[1:]:
                p.remove(extra)
            t = runs[0].find(f"{{{NS['a']}}}t")
            if t is None:
                t = etree.SubElement(runs[0], f"{{{NS['a']}}}t")
            t.text = val
        return True
    return False


def apply_library_slides(deck_path: Path, items: list[dict]) -> None:
    """Insert cloned prototype slides into a built deck.

    items: [{"position": int (0-based slide order), "prototype": str,
             "slots": {shape_name: str|list}}]
    """
    if not items:
        return
    work = Path(tempfile.mkdtemp(prefix="tetp-lib-"))
    deck_dir = work / "deck"
    with zipfile.ZipFile(deck_path) as z:
        z.extractall(deck_dir)
    lib_dirs: dict[Path, Path] = {}

    def lib_dir_for(lib_path: Path) -> Path:
        if lib_path not in lib_dirs:
            d = work / f"lib{len(lib_dirs)}"
            with zipfile.ZipFile(lib_path) as z:
                z.extractall(d)
            lib_dirs[lib_path] = d
        return lib_dirs[lib_path]

    # Deck bookkeeping ---------------------------------------------------------
    ct = etree.parse(str(deck_dir / "[Content_Types].xml"))
    pres = etree.parse(str(deck_dir / "ppt" / "presentation.xml"))
    pres_rels = etree.parse(str(deck_dir / "ppt" / "_rels" / "presentation.xml.rels"))
    sld_lst = pres.getroot().find(f"{{{NS['p']}}}sldIdLst")

    existing = [int(m.group(1)) for f in (deck_dir / "ppt" / "slides").glob("slide*.xml")
                if (m := re.match(r"slide(\d+)\.xml$", f.name))]
    next_num = max(existing, default=0) + 1
    next_rid = max(int(r.get("Id")[3:]) for r in pres_rels.getroot()) + 1
    next_sid = max([int(e.get("id")) for e in sld_lst], default=255) + 1

    # Map of canonical layout name -> part path
    deck_layouts = {}
    for f in (deck_dir / "ppt" / "slideLayouts").glob("slideLayout*.xml"):
        m = re.search(r'<p:cSld name="([^"]*)"', f.read_text(encoding="utf-8"))
        if m:
            deck_layouts[m.group(1)] = f.name

    copied_assets: dict[tuple, str] = {}

    def copy_asset(lib_part: str, lib_dir: Path, tag: str) -> str:
        key = (str(lib_dir), lib_part)
        if key in copied_assets:
            return copied_assets[key]
        kind, name = lib_part.split("/")[1], lib_part.split("/")[-1]
        dst_rel = f"ppt/{kind}/lib{tag}_{name}"
        dst = deck_dir / dst_rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(lib_dir / lib_part, dst)
        ext = name.rsplit(".", 1)[-1].lower()
        defaults = {d.get("Extension", "").lower()
                    for d in ct.getroot().findall(f"{{{NS['ct']}}}Default")}
        mime = {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg",
                "emf": "image/x-emf", "svg": "image/svg+xml",
                "bin": "application/vnd.openxmlformats-officedocument.oleObject"}.get(ext)
        if mime and ext not in defaults:
            el = etree.SubElement(ct.getroot(), f"{{{NS['ct']}}}Default")
            el.set("Extension", ext)
            el.set("ContentType", mime)
        if kind == "tags":
            el = etree.SubElement(ct.getroot(), f"{{{NS['ct']}}}Override")
            el.set("PartName", f"/{dst_rel}")
            el.set("ContentType",
                   "application/vnd.openxmlformats-officedocument.presentationml.tags+xml")
        copied_assets[key] = dst_rel
        return dst_rel

    for item in items:
        lib_path, num = resolve_prototype(item["prototype"])
        lib_dir = lib_dir_for(lib_path)
        lib_tag = str(list(lib_dirs).index(lib_path))
        slide_xml = lib_dir / "ppt" / "slides" / f"slide{num}.xml"
        slide_rels = lib_dir / "ppt" / "slides" / "_rels" / f"slide{num}.xml.rels"

        root = etree.parse(str(slide_xml)).getroot()
        for shape_name, value in (item.get("slots") or {}).items():
            if not _replace_slot_text(root, shape_name, value):
                raise KeyError(
                    f"prototype {item['prototype']!r}: no shape named {shape_name!r}"
                )

        new_num = next_num
        next_num += 1
        new_part = deck_dir / "ppt" / "slides" / f"slide{new_num}.xml"
        new_part.parent.mkdir(parents=True, exist_ok=True)
        new_part.write_bytes(etree.tostring(root, xml_declaration=True,
                                            encoding="UTF-8", standalone=True))

        # Rewrite rels: layout remap by name, drop notes, copy assets.
        rels = etree.parse(str(slide_rels))
        rroot = rels.getroot()
        # find lib layout name
        for rel in list(rroot):
            rtype, target = rel.get("Type"), rel.get("Target")
            if rtype == LAYOUT_RELTYPE:
                lib_layout = posixpath.normpath(posixpath.join("ppt/slides", target))
                m = re.search(r'<p:cSld name="([^"]*)"',
                              (lib_dir / lib_layout).read_text(encoding="utf-8"))
                lname = m.group(1) if m else ""
                mapped = deck_layouts.get(lname) or deck_layouts.get(
                    "TETP EN " + lname.replace("EN ", "")) or deck_layouts[FALLBACK_LAYOUT]
                rel.set("Target", f"../slideLayouts/{mapped}")
            elif rtype == NOTES_RELTYPE:
                rroot.remove(rel)
            elif target.startswith("../"):
                lib_part = posixpath.normpath(posixpath.join("ppt/slides", target))
                rel.set("Target",
                        posixpath.relpath(copy_asset(lib_part, lib_dir, lib_tag),
                                          "ppt/slides"))
        rels_out = deck_dir / "ppt" / "slides" / "_rels" / f"slide{new_num}.xml.rels"
        rels_out.parent.mkdir(parents=True, exist_ok=True)
        rels.write(str(rels_out), xml_declaration=True, encoding="UTF-8",
                   standalone=True)

        ov = etree.SubElement(ct.getroot(), f"{{{NS['ct']}}}Override")
        ov.set("PartName", f"/ppt/slides/slide{new_num}.xml")
        ov.set("ContentType", SLIDE_CT)

        rid = f"rId{next_rid}"
        next_rid += 1
        rel_el = etree.SubElement(pres_rels.getroot(), f"{{{NS['rel']}}}Relationship")
        rel_el.set("Id", rid)
        rel_el.set("Type", SLIDE_RELTYPE)
        rel_el.set("Target", f"slides/slide{new_num}.xml")

        sld = etree.Element(f"{{{NS['p']}}}sldId")
        sld.set("id", str(next_sid))
        next_sid += 1
        sld.set(f"{{{NS['r']}}}id", rid)
        pos = min(item["position"], len(sld_lst))
        sld_lst.insert(pos, sld)

    ct.write(str(deck_dir / "[Content_Types].xml"), xml_declaration=True,
             encoding="UTF-8", standalone=True)
    pres.write(str(deck_dir / "ppt" / "presentation.xml"), xml_declaration=True,
               encoding="UTF-8", standalone=True)
    pres_rels.write(str(deck_dir / "ppt" / "_rels" / "presentation.xml.rels"),
                    xml_declaration=True, encoding="UTF-8", standalone=True)

    tmp_out = work / "out.pptx"
    with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(deck_dir.rglob("*")):
            if f.is_file():
                z.write(f, f.relative_to(deck_dir).as_posix())
    shutil.move(str(tmp_out), deck_path)
    shutil.rmtree(work)
