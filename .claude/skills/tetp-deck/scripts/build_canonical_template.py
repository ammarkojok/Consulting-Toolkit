#!/usr/bin/env python3
"""Build the canonical TETP template (see docs/adr/0001).

Takes the TETP slide-master .pptx (EN + AR masters, chrome) as the base and
imports the distinct content layouts from the TETP .thmx donor master, rewriting
relationships so every imported layout hangs off the base EN master. Media,
think-cell embeddings, and tags referenced by donor layouts are copied along,
renamed to avoid collisions.

Usage: python3 build_canonical_template.py [--out PATH]
"""

from __future__ import annotations

import argparse
import posixpath
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
BASE_PPTX = SKILL / "assets" / "source" / "TETP_Slide_Master_Templates.pptx"
THMX = SKILL / "assets" / "source" / "TETP.thmx"
DEFAULT_OUT = SKILL / "assets" / "TETP-canonical.pptx"

# Donor layouts from .thmx master 9 (theme identical to the base EN master).
# thmx "EN Empty" duplicates base "EN Vertical"; thmx "EN_Horizontal_Content" and
# "EN Baseline Layout" duplicate base "EN Horizontal"; generic "Title and Content"
# carries no TETP chrome. "Table" (slideLayout138) still carries un-rebranded
# Kearney chrome (KEARNEY footer logo, tan band) and must never ship. Those are
# deliberately not imported.
DONOR_LAYOUTS = {
    135: "EN Wide",
    128: "EN Split",
    129: "EN Two Columns",
    130: "EN Bios",
    132: "EN Case Study",
    134: "Title Only",
}

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
REL_NS = NS["rel"]
LAYOUT_CT = (
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
)
TAGS_CT = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"
MASTER_RELTYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)
LAYOUT_RELTYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)


def build(out_path: Path) -> None:
    work = Path(tempfile.mkdtemp(prefix="tetp-canonical-"))
    base_dir = work / "base"
    thmx_dir = work / "thmx"
    for src, dst in ((BASE_PPTX, base_dir), (THMX, thmx_dir)):
        with zipfile.ZipFile(src) as z:
            z.extractall(dst)

    layouts_dir = base_dir / "ppt" / "slideLayouts"
    next_layout_num = (
        max(int(p.stem.replace("slideLayout", "")) for p in layouts_dir.glob("slideLayout*.xml"))
        + 1
    )

    # --- base master1 (EN): rels + sldLayoutIdLst ---------------------------
    master_rels_path = base_dir / "ppt" / "slideMasters" / "_rels" / "slideMaster1.xml.rels"
    master_rels = etree.parse(str(master_rels_path))
    next_rid = (
        max(int(r.get("Id")[3:]) for r in master_rels.getroot()) + 1
    )

    master_path = base_dir / "ppt" / "slideMasters" / "slideMaster1.xml"
    master = etree.parse(str(master_path))
    id_lst = master.getroot().find("p:sldLayoutIdLst", NS)
    next_slid = max(int(e.get("id")) for e in id_lst) + 1

    # --- content types -------------------------------------------------------
    ct_path = base_dir / "[Content_Types].xml"
    ct = etree.parse(str(ct_path))
    ct_root = ct.getroot()
    defaults = {d.get("Extension").lower() for d in ct_root.findall("ct:Default", NS)}

    def ensure_default(ext: str, mime: str) -> None:
        if ext.lower() not in defaults:
            el = etree.SubElement(ct_root, f"{{{NS['ct']}}}Default")
            el.set("Extension", ext)
            el.set("ContentType", mime)
            defaults.add(ext.lower())

    def add_override(part: str, mime: str) -> None:
        el = etree.SubElement(ct_root, f"{{{NS['ct']}}}Override")
        el.set("PartName", part)
        el.set("ContentType", mime)

    copied_assets: dict[str, str] = {}  # donor part path -> base part path

    def copy_asset(donor_rel_target: str) -> str:
        """Copy a media/embedding/tag part from the thmx into the base package."""
        # Donor targets are relative to theme/slideLayouts/, e.g. "../media/image6.emf".
        donor_part = posixpath.normpath(posixpath.join("theme/slideLayouts", donor_rel_target))
        if donor_part in copied_assets:
            return copied_assets[donor_part]
        kind = donor_part.split("/")[1]  # media | embeddings | tags
        name = donor_part.split("/")[-1]
        dst_rel = f"ppt/{kind}/thmx_{name}"
        dst = base_dir / dst_rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(thmx_dir / donor_part, dst)
        ext = name.rsplit(".", 1)[-1].lower()
        if kind == "tags":
            add_override(f"/{dst_rel}", TAGS_CT)
        elif ext == "bin":
            ensure_default("bin", "application/vnd.openxmlformats-officedocument.oleObject")
        elif ext == "emf":
            ensure_default("emf", "image/x-emf")
        elif ext == "png":
            ensure_default("png", "image/png")
        elif ext in ("jpeg", "jpg"):
            ensure_default(ext, "image/jpeg")
        copied_assets[donor_part] = dst_rel
        return dst_rel

    for donor_num, name in sorted(DONOR_LAYOUTS.items()):
        donor_xml = thmx_dir / "theme" / "slideLayouts" / f"slideLayout{donor_num}.xml"
        donor_rels = (
            thmx_dir / "theme" / "slideLayouts" / "_rels" / f"slideLayout{donor_num}.xml.rels"
        )
        new_num = next_layout_num
        next_layout_num += 1
        new_part = f"ppt/slideLayouts/slideLayout{new_num}.xml"
        shutil.copyfile(donor_xml, base_dir / new_part)

        # Rewrite the donor's relationships: master -> base master1; assets copied in.
        rels = etree.parse(str(donor_rels))
        for rel in rels.getroot():
            rtype = rel.get("Type")
            target = rel.get("Target")
            if rtype == MASTER_RELTYPE:
                rel.set("Target", "../slideMasters/slideMaster1.xml")
            elif target.startswith("../"):
                new_target = copy_asset(target)
                rel.set("Target", posixpath.relpath(new_target, "ppt/slideLayouts"))
        rels_out = base_dir / "ppt" / "slideLayouts" / "_rels" / f"slideLayout{new_num}.xml.rels"
        rels.write(str(rels_out), xml_declaration=True, encoding="UTF-8", standalone=True)

        add_override(f"/{new_part}", LAYOUT_CT)

        # Attach to base EN master.
        rid = f"rId{next_rid}"
        next_rid += 1
        rel_el = etree.SubElement(master_rels.getroot(), f"{{{REL_NS}}}Relationship")
        rel_el.set("Id", rid)
        rel_el.set("Type", LAYOUT_RELTYPE)
        rel_el.set("Target", f"../slideLayouts/slideLayout{new_num}.xml")

        id_el = etree.SubElement(id_lst, f"{{{NS['p']}}}sldLayoutId")
        id_el.set("id", str(next_slid))
        id_el.set(f"{{{NS['r']}}}id", rid)
        next_slid += 1
        print(f"imported thmx slideLayout{donor_num} ('{name}') -> {new_part}")

    master_rels.write(str(master_rels_path), xml_declaration=True, encoding="UTF-8", standalone=True)
    master.write(str(master_path), xml_declaration=True, encoding="UTF-8", standalone=True)
    ct.write(str(ct_path), xml_declaration=True, encoding="UTF-8", standalone=True)

    # --- repack ---------------------------------------------------------------
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.exists():
        out_path.unlink()
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(base_dir.rglob("*")):
            if f.is_file():
                z.write(f, f.relative_to(base_dir).as_posix())
    shutil.rmtree(work)
    print(f"wrote {out_path}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", type=Path, default=DEFAULT_OUT)
    build(ap.parse_args().out)
