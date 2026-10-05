"""Atomically add the Product customUI part to an XLAM."""
from __future__ import annotations

import argparse
import io
import json
import posixpath
import shutil
import hashlib
import os
import re
import tempfile
import zipfile
import sys
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
from ooxml_webextensions import clean_members
from xlam_package_order import ordered_names
from generate_quick_format_icons import QUICK_FORMAT_IMAGES, render_icon

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
UI_REL = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
LEGACY_UI_REL = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
UI_CT = "application/xml"
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
IMAGE_RELS_PART = "customUI/_rels/customUI14.xml.rels"


def quick_format_members(ribbon: Path, xml: bytes) -> dict[str, bytes]:
    """Resolve only owned relative PNGs; no remote or caller-supplied image paths."""
    root = ET.fromstring(xml)
    image_ids = [element.attrib["image"] for element in root.iter() if "image" in element.attrib]
    if not image_ids:
        return {}
    manifest_path = ribbon.parent / "brand-images.json"
    brand_images = json.loads(manifest_path.read_text(encoding="utf-8"))["images"] if manifest_path.is_file() else {}
    if not set(QUICK_FORMAT_IMAGES).issubset(image_ids) or not set(image_ids).issubset(set(QUICK_FORMAT_IMAGES) | set(brand_images)):
        raise ValueError("custom Ribbon images must exactly match the quick-format allowlist")
    relationships = ET.Element(f"{{{REL_NS}}}Relationships")
    additions = {}
    source_root = ribbon.parent.resolve()
    # Dynamic menus share the package image IDs even before their first opening.
    for image_id in sorted(set(image_ids) | set(brand_images)):
        relative = QUICK_FORMAT_IMAGES[image_id] if image_id in QUICK_FORMAT_IMAGES else brand_images[image_id]["path"]
        source = (source_root / relative).resolve()
        if not source.is_relative_to(source_root) or not source.is_file():
            raise ValueError("quick-format image path is missing or outside the ribbon directory")
        content = source.read_bytes()
        if image_id in QUICK_FORMAT_IMAGES and content != render_icon(image_id):
            raise ValueError("quick-format image content mismatch")
        if image_id in brand_images and hashlib.sha256(content).hexdigest() != brand_images[image_id]["sha256"]:
            raise ValueError("brand image content mismatch")
        additions["customUI/" + relative] = content
        ET.SubElement(relationships, f"{{{REL_NS}}}Relationship", Id=image_id, Type=IMAGE_REL, Target=relative)
    ET.register_namespace("", REL_NS)
    additions[IMAGE_RELS_PART] = ET.tostring(relationships, encoding="utf-8", xml_declaration=True)
    return additions


def synchronize_workbook_codename(members: dict[str, bytes]) -> None:
    """Keep Excel's workbook metadata consistent with the renamed VBA host."""
    binary = members.get("xl/vbaProject.bin", b"")
    if "NxProductWorkbook".encode("utf-16le") not in binary:
        return
    xml = members["xl/workbook.xml"]
    pattern = rb'(<workbookPr\b[^>]*\bcodeName=")[^"]*(")'
    updated, count = re.subn(pattern, rb'\g<1>NxProductWorkbook\2', xml)
    if count != 1:
        raise ValueError("Product workbook codeName metadata must occur exactly once")
    ET.fromstring(updated)
    members["xl/workbook.xml"] = updated


def inject(xlam: Path, ribbon: Path, ribbon_sha: str | None = None) -> None:
    if not xlam.is_file() or not ribbon.is_file():
        raise FileNotFoundError("XLAM and ribbon inputs are required")
    if ribbon_sha and hashlib.sha256(ribbon.read_bytes()).hexdigest() != ribbon_sha.lower():
        raise ValueError("ribbon SHA mismatch")
    if ribbon.read_bytes()[:1] != b"<":
        raise ValueError("ribbon XML is not XML text")
    ribbon_bytes = ribbon.read_bytes()
    image_members = quick_format_members(ribbon, ribbon_bytes)
    target = "customUI/customUI14.xml"
    with zipfile.ZipFile(xlam, "r") as source:
        info = source.infolist()
        raw_names = [item.filename for item in info]
        if len(raw_names) != len(set(raw_names)):
            raise ValueError("duplicate ZIP entry rejected")
        names = set(raw_names)
        required = {"[Content_Types].xml", "_rels/.rels", "xl/vbaProject.bin"}
        missing = required - names
        if missing:
            raise ValueError(f"XLAM OOXML structure incomplete: {sorted(missing)}")
        if target in names:
            raise ValueError("existing customUI part rejected")
        if set(image_members) & names:
            raise ValueError("existing custom Ribbon image or relationship part rejected")
        members = {name: source.read(name) for name in names}
    members, _ = clean_members(members)
    synchronize_workbook_codename(members)
    content_types = ET.fromstring(members["[Content_Types].xml"])
    ct_matches = [e for e in content_types if e.attrib.get("PartName") == "/" + target]
    if len(ct_matches) > 1:
        raise ValueError("duplicate customUI content-type rejected")
    if not ct_matches:
        ET.register_namespace("", CT_NS)
        ET.SubElement(content_types, f"{{{CT_NS}}}Override", PartName="/" + target, ContentType=UI_CT)
    if image_members:
        png_types = [
            element for element in content_types
            if element.tag == f"{{{CT_NS}}}Default" and element.attrib.get("Extension", "").lower() == "png"
        ]
        if len(png_types) > 1 or any(element.attrib.get("ContentType") != "image/png" for element in png_types):
            raise ValueError("conflicting PNG content-type rejected")
        if not png_types:
            ET.SubElement(content_types, f"{{{CT_NS}}}Default", Extension="png", ContentType="image/png")
        for part in image_members:
            overrides = [element for element in content_types if element.attrib.get("PartName") == "/" + part]
            if overrides:
                raise ValueError("existing custom Ribbon image content-type override rejected")
    ET.register_namespace("", CT_NS)
    members["[Content_Types].xml"] = ET.tostring(content_types, encoding="utf-8", xml_declaration=True)
    rels = ET.fromstring(members["_rels/.rels"])
    rel_matches = [e for e in rels if e.attrib.get("Type") in {UI_REL, LEGACY_UI_REL}]
    if len(rel_matches) > 1:
        raise ValueError("duplicate customUI relationship rejected")
    if rel_matches:
        raise ValueError("existing customUI relationship rejected")
    ET.register_namespace("", REL_NS)
    ET.SubElement(rels, f"{{{REL_NS}}}Relationship", Id="rIdNxRibbon", Type=UI_REL, Target=target)
    members["_rels/.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
    members[target] = ribbon_bytes
    members.update(image_members)
    fd, tmp_name = tempfile.mkstemp(prefix=xlam.name + ".", suffix=".tmp", dir=xlam.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp_name, "w", compression=zipfile.ZIP_DEFLATED) as out:
            for name in ordered_names(members):
                out.writestr(name, members[name])
        with zipfile.ZipFile(tmp_name, "r") as check:
            if target not in check.namelist() or check.read(target) != ribbon_bytes:
                raise ValueError("ribbon injection verification failed")
            if any(check.read(name) != expected for name, expected in image_members.items()):
                raise ValueError("custom Ribbon image injection verification failed")
        Path(tmp_name).replace(xlam)
    finally:
        Path(tmp_name).unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlam", type=Path, required=True)
    parser.add_argument("--ribbon", type=Path, required=True)
    parser.add_argument("--ribbon-sha", default=None)
    args = parser.parse_args()
    inject(args.xlam, args.ribbon, args.ribbon_sha)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
