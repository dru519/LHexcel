#!/usr/bin/env python3
"""Generate the shared A-type Full Functions and Navigator catalog."""
from __future__ import annotations

import argparse
import json
import os
import re
import tempfile
from collections import defaultdict
from pathlib import Path
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[1]
FEATURES = ROOT / "contracts" / "feature-contract.json"
COMMANDS = ROOT / "contracts" / "command-contract.json"
RIBBON = ROOT / "contracts" / "ribbon-shell-contract.json"
NAVIGATION = ROOT / "src" / "resources" / "navigation.ko-KR.json"
VBA = ROOT / "src" / "vba" / "ui" / "NxGeneratedNavigationCatalog.bas"
CSHARP = ROOT / "src" / "dotnet" / "NxHost" / "NxGeneratedNavigatorCatalog.cs"
PROFILE_IDS = ["internal-xlam", "enhanced-dll"]
NAVIGATION_FIELDS = {
    "ribbon_group_id", "object_scope", "purpose_tags", "launch_surface",
    "mutation_scope", "selection_mode", "feedback_mode", "profile_availability",
}
ITEM_TYPES = {"feature", "command"}
ID_PATTERN = re.compile(r"^NX-[A-Z0-9-]+$")
DISPLAY_TEXT = {
    "excel_ready": "선택 없이 사용 가능",
    "active_data_region": "현재 데이터 표", "active_window": "현재 Excel 창",
    "active_workbook": "현재 통합문서", "active_worksheet": "현재 시트",
    "range_or_shape": "선택 셀 또는 그림", "range_selection": "선택한 셀",
    "selection_optional": "선택 없이 사용 가능", "selection_or_clipboard": "선택 셀 또는 복사한 내용",
    "shape_selection": "선택한 그림", "fast": "바로 실행", "guarded": "확인 후 실행", "planned": "계획 확인",
}


def load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def xml(value: object) -> str:
    return escape(str(value), {'"': "&quot;", "'": "&apos;"})


def vba(value: object) -> str:
    return str(value).replace('"', '""')


def csharp(value: object) -> str:
    return str(value).replace("\\", "\\\\").replace('"', '\\"').replace("\r", "\\r").replace("\n", "\\n")


def safe_id(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_]", "_", value)


def categories(shell: dict) -> list[dict[str, object]]:
    source = shell.get("navigation_categories")
    if not isinstance(source, list):
        raise ValueError("A-type navigation category source rejected")
    rows = [
        {
            "id": str(row["id"]),
            "label": str(row["label"]),
            "image_mso": str(row["image_mso"]),
            "order": int(row["order"]),
        }
        for row in source
    ]
    ids = [row["id"] for row in rows]
    if (
        not ids
        or len(set(ids)) != len(ids)
        or [row["order"] for row in rows] != sorted({row["order"] for row in rows})
        or any(not row["image_mso"] for row in rows)
    ):
        raise ValueError("navigation categories must be unique ordered groups")
    return rows


def validate_navigation(item_id: str, navigation: object, category_ids: set[str]) -> dict:
    navigation_keys = set(navigation) if isinstance(navigation, dict) else set()
    allowed_fields = NAVIGATION_FIELDS | ({"transport"} if item_id == "NX-HANGUL-TABLE-SEND" else set())
    if navigation_keys != allowed_fields:
        raise ValueError(f"navigation fields rejected: {item_id}")
    if item_id == "NX-HANGUL-TABLE-SEND" and navigation.get("transport") != "embedded-HWPX":
        raise ValueError(f"Hangul navigation transport rejected: {item_id}")
    if navigation["ribbon_group_id"] not in category_ids:
        raise ValueError(f"navigation category rejected: {item_id}")
    for field in ("object_scope", "purpose_tags"):
        values = navigation[field]
        if not isinstance(values, list) or not values or any(not isinstance(value, str) or not value for value in values):
            raise ValueError(f"navigation {field} rejected: {item_id}")
    if navigation["profile_availability"] != PROFILE_IDS:
        raise ValueError(f"navigation profiles rejected: {item_id}")
    return dict(navigation)


def search_text(item_id: str, label: str, category_label: str) -> str:
    """Build the approved literal navigator index without synonym expansion."""
    tokens = [item_id, label, category_label]
    return " | ".join(dict.fromkeys(token.strip() for token in tokens if token and token.strip()))


def feature_input_context(feature: dict, navigation: dict) -> str:
    scope = set(str(value) for value in navigation["object_scope"])
    mode = str(navigation["selection_mode"])
    if mode == "required":
        if "shape" in scope and "selection" not in scope:
            return "shape_selection"
        if "shape" in scope:
            return "range_or_shape"
        return "range_selection"
    if mode == "optional":
        return "selection_optional"
    if "worksheet" in scope:
        return "active_worksheet"
    if "workbook" in scope or "filesystem" in scope:
        return "active_workbook"
    if "window" in scope:
        return "active_window"
    return "excel_ready"


def build_document(feature_contract: dict, command_contract: dict, shell: dict) -> dict:
    category_rows = categories(shell)
    category_by_id = {str(row["id"]): row for row in category_rows}
    category_ids = set(category_by_id)
    command_groups = {str(row["id"]): row for row in command_contract.get("groups", [])}
    if len(command_groups) != 14 or command_contract.get("normalized_command_count") != len(command_contract.get("commands", [])):
        raise ValueError("command group catalog rejected")
    if any(str(row.get("navigation_mode", "nested")) not in {"nested", "flat"} for row in command_groups.values()):
        raise ValueError("command group navigation mode rejected")

    items: list[dict[str, object]] = []
    for feature in feature_contract.get("features", []):
        if not feature.get("catalog_visible"):
            continue
        item_id = str(feature["id"])
        navigation = validate_navigation(item_id, feature.get("navigation"), category_ids)
        label = str(feature["ribbon"].get("label") or item_id)
        description = str(feature.get("description_ko") or " > ".join(str(value) for value in feature.get("catalog_path", [])) or label)
        route_key = f"feature:{item_id}"
        items.append({
            "id": item_id,
            "item_type": "feature",
            "route_key": route_key,
            "label_ko": label,
            "description_ko": description,
            "search_text": search_text(
                item_id,
                label,
                str(category_by_id[navigation["ribbon_group_id"]]["label"]),
            ),
            "category_id": navigation["ribbon_group_id"],
            "category_label": category_by_id[navigation["ribbon_group_id"]]["label"],
            "category_order": category_by_id[navigation["ribbon_group_id"]]["order"],
            "group_id": "NX-NAV-GRP-DATA-COMPARE" if feature.get("catalog_path", [])[1:2] == ["데이터 비교"] else "NX-NAV-GRP-FEATURES",
            "group_label": "데이터 비교" if feature.get("catalog_path", [])[1:2] == ["데이터 비교"] else "기능",
            "group_mode": "nested" if feature.get("catalog_path", [])[1:2] == ["데이터 비교"] else "flat",
            "group_order": 0,
            "sort_order": int(feature["ribbon"].get("order", 0)),
            "image_mso": str(feature["ribbon"].get("image_mso") or "ControlProperties"),
            "input_context": feature_input_context(feature, navigation),
            "selection_mode": str(navigation["selection_mode"]),
            "launch_surface": str(feature["launch_surface"]),
            "mutation_scope": str(navigation["mutation_scope"]),
            "feedback_mode": str(navigation["feedback_mode"]),
            "execution_grade": str(feature["base_execution_grade"]),
            "object_scope": ",".join(str(value) for value in navigation["object_scope"]),
            "navigation": navigation,
        })

    for command in command_contract.get("commands", []):
        if not command.get("taskpane_visible"):
            continue
        item_id = str(command["id"])
        navigation = validate_navigation(item_id, command.get("navigation"), category_ids)
        group_id = str(command["group_id"])
        if group_id not in command_groups:
            raise ValueError(f"unknown command group: {item_id}")
        label = str(command["label_ko"])
        description = str(command["description_ko"])
        route_key = f"command:{item_id}"
        items.append({
            "id": item_id,
            "item_type": "command",
            "route_key": route_key,
            "label_ko": label,
            "description_ko": description,
            "search_text": search_text(
                item_id,
                label,
                str(category_by_id[navigation["ribbon_group_id"]]["label"]),
            ),
            "category_id": navigation["ribbon_group_id"],
            "category_label": category_by_id[navigation["ribbon_group_id"]]["label"],
            "category_order": category_by_id[navigation["ribbon_group_id"]]["order"],
            "group_id": group_id,
            "group_label": str(command_groups[group_id]["label_ko"]),
            "group_mode": str(command.get("navigation_mode", command_groups[group_id].get("navigation_mode", "nested"))),
            "group_order": int(command["group_order"]),
            "sort_order": int(command["sort_order"]),
            "image_mso": str(command["image_mso"]),
            "input_context": str(command["input_context"]),
            "selection_mode": str(navigation["selection_mode"]),
            "launch_surface": str(navigation["launch_surface"]),
            "mutation_scope": str(navigation["mutation_scope"]),
            "feedback_mode": str(navigation["feedback_mode"]),
            "execution_grade": "guarded" if command["mutates_document"] else "fast",
            "object_scope": ",".join(str(value) for value in navigation["object_scope"]),
            "navigation": navigation,
        })

    items.sort(key=lambda row: (row["category_order"],
                               0 if row["category_id"] == "NX-GRP-SHEET" else row["group_order"],
                               row["sort_order"], row["id"]))
    ids = [str(row["id"]) for row in items]
    if any(ID_PATTERN.fullmatch(item_id) is None for item_id in ids) or len(ids) != len(set(ids)):
        raise ValueError("navigation item identity rejected")
    feature_count = sum(row["item_type"] == "feature" for row in items)
    command_count = sum(row["item_type"] == "command" for row in items)
    expected_command_count = sum(
        bool(command.get("taskpane_visible"))
        for command in command_contract.get("commands", [])
    )
    expected_feature_count = sum(bool(feature.get("catalog_visible")) for feature in feature_contract.get("features", []))
    if feature_count != expected_feature_count or command_count != expected_command_count or len(items) != feature_count + expected_command_count:
        raise ValueError("navigation inventory count rejected")
    return {
        "schema_version": 2,
        "surface": "a-type-shared-catalog",
        "categories": category_rows,
        "feature_count": feature_count,
        "command_count": command_count,
        "item_count": len(items),
        "items": items,
    }


def render_json(document: dict) -> str:
    return json.dumps(document, ensure_ascii=False, indent=2) + "\n"


def category_menu_rows(rows: list[dict]) -> list[tuple[dict, list[dict] | None]]:
    """The one grouping/order rule shared by Ribbon menus and Full Functions."""
    result = []
    emitted_groups: set[str] = set()
    for item in rows:
        if (item["item_type"] == "feature" and item.get("group_mode") != "nested") or item.get("group_mode") == "flat":
            result.append((item, None))
            continue
        group_id = str(item["group_id"])
        if group_id not in emitted_groups:
            emitted_groups.add(group_id)
            result.append((item, [row for row in rows if row["group_id"] == group_id and row.get("group_mode") != "flat"]))
    return result


VBA_ROUTE_FIELDS = [
    "route_key", "item_type", "id", "label_ko", "description_ko",
    "category_id", "category_label", "image_mso", "input_context",
    "selection_mode", "launch_surface", "mutation_scope", "feedback_mode",
    "execution_grade", "search_text", "object_scope",
]
VBA_ITEMS_PER_CHUNK = 32


def render_vba_menu_button(item: dict) -> str:
    return (
        '    xml = xml & "<button id=""nav_{id}"" label=""{label}"" imageMso=""{image}"" '
        'tag=""nx1|{kind}|{target}"" enabled=""" & NxRouteAvailabilityRibbonValue("{route}") & '
        '""" supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("{route}")) & '
        '""" onAction=""NxRibbonExecute""/>"'
    ).format(
        id=safe_id(str(item["id"])),
        label=vba(xml(item["label_ko"])),
        image=vba(xml(item["image_mso"])),
        kind=item["item_type"],
        target=item["id"],
        route=item["route_key"],
    )


def render_vba(document: dict) -> str:
    by_category: dict[str, list[dict]] = defaultdict(list)
    for item in document["items"]:
        by_category[str(item["category_id"])].append(item)
    item_chunks = [
        document["items"][offset:offset + VBA_ITEMS_PER_CHUNK]
        for offset in range(0, len(document["items"]), VBA_ITEMS_PER_CHUNK)
    ]
    lines = [
        'Attribute VB_Name = "NxGeneratedNavigationCatalog"',
        "Option Explicit",
        "",
        "Private mNavigationItems As Collection",
        "Private mNavigationRouteIndex As Object",
        "",
        "Public Function NxGeneratedNavigationItems() As Collection",
        "    Dim items As Collection",
        "    If mNavigationItems Is Nothing Then",
        "        Set items = New Collection",
    ]
    for chunk_index in range(len(item_chunks)):
        lines.append(f"        NxGeneratedNavigationAddItems{chunk_index + 1:02d} items")
    lines.extend([
        "        Set mNavigationItems = items",
        "    End If",
        "    Set NxGeneratedNavigationItems = mNavigationItems",
        "End Function",
        "",
    ])
    for chunk_index, chunk in enumerate(item_chunks, start=1):
        lines.append(f"Private Sub NxGeneratedNavigationAddItems{chunk_index:02d}(ByRef items As Collection)")
        for item in chunk:
            values = ", ".join(f'"{vba(item[field])}"' for field in VBA_ROUTE_FIELDS)
            lines.append(f"    items.Add Array({values})")
        lines.extend(["End Sub", ""])
    lines.extend([
        "Private Function NxGeneratedNavigationRouteIndex() As Object",
        "    Dim item As Variant",
        "    Dim routeIndex As Object",
        "    If mNavigationRouteIndex Is Nothing Then",
        '        Set routeIndex = CreateObject("Scripting.Dictionary")',
        "        routeIndex.CompareMode = vbBinaryCompare",
        "        For Each item In NxGeneratedNavigationItems()",
        "            routeIndex.Add CStr(item(0)), item",
        "        Next item",
        "        Set mNavigationRouteIndex = routeIndex",
        "    End If",
        "    Set NxGeneratedNavigationRouteIndex = mNavigationRouteIndex",
        "End Function",
        "",
        "Public Function NxGeneratedNavigationRouteExists(ByVal routeKey As String) As Boolean",
        "    routeKey = NxCanonicalRouteKey(routeKey)",
        "    NxGeneratedNavigationRouteExists = NxGeneratedNavigationRouteIndex().Exists(routeKey)",
        "End Function",
        "",
        "Public Function NxGeneratedNavigationRouteField(ByVal routeKey As String, ByVal fieldName As String) As String",
        "    Dim item As Variant",
        "    Dim routeIndex As Object",
        "    Dim fieldIndex As Long",
        "    routeKey = NxCanonicalRouteKey(routeKey)",
        "    Select Case fieldName",
    ])
    for index, field in enumerate(VBA_ROUTE_FIELDS):
        lines.append(f'        Case "{field}": fieldIndex = {index}')
    lines.extend([
        '        Case Else: NxRaiseContractError "Unknown navigation route field"',
        "    End Select",
        "    Set routeIndex = NxGeneratedNavigationRouteIndex()",
        '    If Not routeIndex.Exists(routeKey) Then NxRaiseContractError "Unknown navigation route key"',
        "    item = routeIndex.Item(routeKey)",
        "    NxGeneratedNavigationRouteField = CStr(item(fieldIndex))",
        "End Function",
        "",
        "Public Function NxGeneratedNavigationMenuXml() As String",
        "    Dim xml As String",
        '    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"',
    ])
    active_categories = [row for row in document["categories"] if by_category.get(str(row["id"]))]
    for category in active_categories:
        lines.append(f"    xml = xml & NxGeneratedNavigationCategory_{safe_id(str(category['id']))}()")
    lines.extend(['    NxGeneratedNavigationMenuXml = xml & "</menu>"', "End Function", ""])
    for category in active_categories:
        category_id = str(category["id"])
        function_name = f"NxGeneratedNavigationCategory_{safe_id(category_id)}"
        lines.extend([
            f"Private Function {function_name}() As String",
            "    Dim xml As String",
            f'    xml = "<menu id=""nav_cat_{safe_id(category_id)}"" label=""{vba(xml(category["label"]))}"" imageMso=""{vba(xml(category["image_mso"]))}"">"',
        ])
        rows = by_category[category_id]
        for item, children in category_menu_rows(rows):
            if children is None:
                lines.append(render_vba_menu_button(item))
                continue
            group_id = str(item["group_id"])
            lines.append(f'    xml = xml & "<menu id=""nav_grp_{safe_id(group_id)}"" label=""{vba(xml(item["group_label"]))}"" imageMso=""{vba(xml(item["image_mso"]))}"">"')
            for grouped in children:
                lines.append(render_vba_menu_button(grouped))
            lines.append('    xml = xml & "</menu>"')
        lines.extend([f'    {function_name} = xml & "</menu>"', "End Function", ""])
    lines.extend([
        'Public Function NxNavigationDisplayText(ByVal token As String) As String',
        '    Select Case token',
    ])
    for token, caption in DISPLAY_TEXT.items():
        lines.append(f'        Case "{token}": NxNavigationDisplayText = "{vba(caption)}"')
    lines.extend(['        Case Else: NxNavigationDisplayText = "입력 조건 확인"', '    End Select', 'End Function', ''])
    brand = json.loads((ROOT / "src/ribbon/brand-images.json").read_text(encoding="utf-8"))
    lines.extend(["", "Public Function NxGeneratedBrandImage(ByVal tag As String) As String", "    Select Case tag"])
    for tag, key in sorted(brand["tags"].items()):
        lines.append(f'        Case "{vba(tag)}": NxGeneratedBrandImage = "NxBrand_{key.replace("-", "_")}_16"')
    lines.extend(["    End Select", "End Function", ""])
    return "\n".join(lines)


def render_csharp(document: dict) -> str:
    lines = [
        "// Generated by tools/generate_navigation_surface.py. Do not edit.",
        "namespace LH.NxHost {",
        "    internal static class NxGeneratedNavigatorCatalog {",
        "        internal static readonly NavigatorCategory[] Categories = new NavigatorCategory[] {",
    ]
    for category in document["categories"]:
        lines.append(f'            new NavigatorCategory("{csharp(category["id"])}", "{csharp(category["label"])}"),')
    lines.extend([
        "        };",
        "        internal static readonly NavigatorItem[] Items = new NavigatorItem[] {",
    ])
    for item in document["items"]:
        lines.append(
            '            new NavigatorItem("{kind}", "{id}", "{label}", "{description}", "{category}", "{search}", '
            '"{input_context}", "{selection_mode}", "{launch_surface}", "{mutation_scope}", "{feedback_mode}", "{execution_grade}"),'.format(
                kind=csharp(item["item_type"]), id=csharp(item["id"]), label=csharp(item["label_ko"]),
                description=csharp(item["description_ko"]), category=csharp(item["category_id"]), search=csharp(item["search_text"]),
                input_context=csharp(item["input_context"]), selection_mode=csharp(item["selection_mode"]),
                launch_surface=csharp(item["launch_surface"]), mutation_scope=csharp(item["mutation_scope"]),
                feedback_mode=csharp(item["feedback_mode"]), execution_grade=csharp(item["execution_grade"]),
            )
        )
    lines.extend(["        };", "        internal static string DisplayText(string token) {", "            switch (token) {"])
    for token, caption in DISPLAY_TEXT.items():
        lines.append(f'                case "{token}": return "{csharp(caption)}";')
    lines.extend(['                default: return "입력 조건 확인";', "            }", "        }", "    }", "}", ""])
    return "\n".join(lines)


def atomic_write(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def outputs() -> dict[Path, str]:
    document = build_document(load(FEATURES), load(COMMANDS), load(RIBBON))
    return {NAVIGATION: render_json(document), VBA: render_vba(document), CSHARP: render_csharp(document)}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    args = parser.parse_args()
    rendered = outputs()
    if args.check:
        stale = [path for path, content in rendered.items() if not path.is_file() or path.read_bytes() != content.encode("utf-8")]
        if stale:
            raise SystemExit("stale navigation surface: " + ", ".join(str(path) for path in stale))
        return 0
    for path, content in rendered.items():
        atomic_write(path, content)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
