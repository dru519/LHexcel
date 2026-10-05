from __future__ import annotations

"""Read-only OOXML/XLSM package inventory used by the independence gate.

The extractor deliberately works on ZIP bytes only.  It never resolves a
reference path and does not attempt to execute VBA.  Missing or encrypted
parts are reported as ``INCOMPLETE`` rather than being treated as evidence of
low similarity.
"""

import argparse
import base64
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _xml(data: bytes) -> ET.Element | None:
    try:
        return ET.fromstring(data)
    except (ET.ParseError, ValueError):
        return None


def _text(value: str | None) -> str:
    return (value or "").strip()


def _relationship_parts(names: list[str], blobs: dict[str, bytes]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for name in names:
        if not name.endswith(".rels"):
            continue
        root = _xml(blobs[name])
        if root is None:
            continue
        source = name[:-5].rstrip("/") or "/"
        if source.endswith("/_rels"):
            source = source[:-6]
        elif "/_rels/" in source:
            source = source.replace("/_rels/", "/")
        for rel in root.findall("{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
            rows.append({
                "part": source,
                "id": _text(rel.attrib.get("Id")),
                "type": _text(rel.attrib.get("Type")),
                "target": _text(rel.attrib.get("Target")),
                "target_mode": _text(rel.attrib.get("TargetMode")),
            })
    return sorted(rows, key=lambda row: (row["part"], row["id"], row["target"]))


def _sheets(blobs: dict[str, bytes], relationships: list[dict[str, str]]) -> list[dict[str, object]]:
    workbook = _xml(blobs.get("xl/workbook.xml", b""))
    if workbook is None:
        return []
    rel_targets = {row["id"]: row["target"] for row in relationships if row["part"] == "xl/workbook"}
    result: list[dict[str, object]] = []
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        rid = sheet.attrib.get("{" + NS["rel"] + "}id", "")
        target = rel_targets.get(rid, "")
        if target and not target.startswith("/"):
            target = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
        result.append({"name": sheet.attrib.get("name", ""), "sheet_id": sheet.attrib.get("sheetId", ""), "relationship_id": rid, "part": target})
    return result


def _cell_inventory(blobs: dict[str, bytes], sheets: list[dict[str, object]]) -> dict[str, dict[str, object]]:
    result: dict[str, dict[str, object]] = {}
    shared_root = _xml(blobs.get("xl/sharedStrings.xml", b""))
    shared = ["".join(node.itertext()) for node in shared_root.findall("main:si", NS)] if shared_root is not None else []
    for sheet in sheets:
        part = str(sheet["part"])
        root = _xml(blobs.get(part, b""))
        cells = []
        styles: dict[str, int] = {}
        if root is not None:
            for cell in root.findall(".//main:c", NS):
                ref = cell.attrib.get("r", "")
                style = cell.attrib.get("s", "0")
                styles[style] = styles.get(style, 0) + 1
                value = cell.findtext("main:v", default="", namespaces=NS)
                inline = cell.find("main:is", NS)
                if inline is not None:
                    value = "".join(inline.itertext())
                elif cell.attrib.get("t") == "s" and value.isdigit() and int(value) < len(shared):
                    value = shared[int(value)]
                if ref or value:
                    cells.append({"ref": ref, "type": cell.attrib.get("t", ""), "value": value or ""})
        result[str(sheet["name"])] = {"part": part, "cell_count": len(cells), "occupied_cells": cells, "style_counts": styles}
    return result


def _styles(blobs: dict[str, bytes]) -> dict[str, int]:
    root = _xml(blobs.get("xl/styles.xml", b""))
    if root is None:
        return {}
    return {name: len(root.findall("main:" + name, NS)[0]) if root.findall("main:" + name, NS) else 0 for name in ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs", "tableStyles")}


def _ribbon(blobs: dict[str, bytes]) -> dict[str, object]:
    labels: list[str] = []
    callbacks: list[dict[str, str]] = []
    for name, data in blobs.items():
        if not (name.lower().startswith("customui/") or name.lower().endswith("customui14.xml")):
            continue
        root = _xml(data)
        if root is None:
            continue
        for element in root.iter():
            for key, value in element.attrib.items():
                if key.casefold() in {"label", "screentip", "supertip", "description"}:
                    labels.append(value)
                if key.casefold() in {"onaction", "getlabel", "getenabled", "getvisible", "getimage", "getsize", "getinsertbeforemso", "getshowimage"}:
                    callbacks.append({"part": name, "element": element.tag.rsplit("}", 1)[-1], "attribute": key, "value": value})
    return {"parts": sorted(name for name in blobs if name.lower().startswith("customui/")), "labels": sorted(set(labels)), "callbacks": sorted(callbacks, key=lambda row: tuple(row.values()))}


def extract_zip(package: Path) -> dict[str, object]:
    package = Path(package)
    try:
        with zipfile.ZipFile(package, "r") as archive:
            infos = archive.infolist()
            blobs = {info.filename: archive.read(info) for info in infos if not info.is_dir()}
    except (OSError, zipfile.BadZipFile, RuntimeError, KeyError) as exc:
        return {"status": "INCOMPLETE", "error": f"package extraction failed: {exc}", "package": str(package)}
    names = [info.filename for info in infos]
    duplicate_names = sorted({name for name in names if names.count(name) > 1})
    relationships = _relationship_parts(names, blobs)
    sheets = _sheets(blobs, relationships)
    assets = []
    for name, data in sorted(blobs.items()):
        lowered = name.lower()
        if any(token in lowered for token in ("/media/", "/embeddings/", "/activex/", "oleobject", ".bin")):
            assets.append({"path": name, "sha256": _sha(data), "size": len(data)})
    custom_xml = [{"path": name, "sha256": _sha(data), "size": len(data)} for name, data in sorted(blobs.items()) if name.lower().startswith("customxml/")]
    defined_names = []
    workbook = _xml(blobs.get("xl/workbook.xml", b""))
    if workbook is not None:
        defined_names = [{"name": node.attrib.get("name", ""), "local_sheet_id": node.attrib.get("localSheetId", ""), "formula": "".join(node.itertext()).strip()} for node in workbook.findall("main:definedNames/main:definedName", NS)]
    inventory = _cell_inventory(blobs, sheets)
    strings = sorted({str(cell["value"]) for sheet in inventory.values() for cell in sheet["occupied_cells"] if str(cell["value"])})
    report: dict[str, object] = {
        "status": "PASS",
        "package": package.name,
        "package_sha256": _sha(package.read_bytes()),
        "zip_entry_count": len(names),
        "zip_entries": [{"path": name, "directory": info.is_dir(), "size": len(blobs.get(name, b"")), "sha256": _sha(blobs[name]) if name in blobs else ""} for info in sorted(infos, key=lambda item: item.filename)],
        "duplicate_zip_entries": duplicate_names,
        "relationships": relationships,
        "sheet_topology": sheets,
        "defined_names": defined_names,
        "cell_inventory": inventory,
        "strings": strings,
        "styles": _styles(blobs),
        "custom_xml": custom_xml,
        "assets": assets,
        "ribbon": _ribbon(blobs),
        "vba_project": ({"sha256": _sha(blobs["xl/vbaProject.bin"]), "size": len(blobs["xl/vbaProject.bin"])} if "xl/vbaProject.bin" in blobs else None),
    }
    if duplicate_names:
        report["status"] = "INCOMPLETE"
        report["error"] = "duplicate ZIP entry names"
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("package", type=Path)
    parser.add_argument("--json", type=Path, required=True)
    args = parser.parse_args(argv)
    report = extract_zip(args.package)
    args.json.parent.mkdir(parents=True, exist_ok=True)
    args.json.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["status"] == "PASS" else 20


if __name__ == "__main__":
    raise SystemExit(main())
