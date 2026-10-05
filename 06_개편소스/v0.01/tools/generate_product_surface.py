#!/usr/bin/env python3
"""Generate the Home and product RibbonX surfaces from closed contracts."""
from __future__ import annotations

import argparse
import json
import os
import tempfile
from copy import deepcopy
from pathlib import Path
from xml.sax.saxutils import escape

try:  # CLI and importlib-based contract tests use different module roots.
    from generate_navigation_surface import build_document, category_menu_rows
    from generate_quick_format_icons import QUICK_FORMAT_IMAGES
except ModuleNotFoundError:
    from tools.generate_navigation_surface import build_document, category_menu_rows
    from tools.generate_quick_format_icons import QUICK_FORMAT_IMAGES

ROOT = Path(__file__).resolve().parents[1]
BRAND_IMAGES = json.loads((ROOT / "src/ribbon/brand-images.json").read_text(encoding="utf-8"))
FEATURES = ROOT / "contracts" / "feature-contract.json"
COMMANDS = ROOT / "contracts" / "command-contract.json"
EXECUTION_GRADES = ROOT / "contracts" / "execution-grade-contract.json"
SHELL = ROOT / "contracts" / "ribbon-shell-contract.json"
UI_SURFACES = ROOT / "contracts" / "ui-surface-contract.json"
XML = ROOT / "src" / "ribbon" / "customUI14.xml"
GENERATED = ROOT / "src" / "vba" / "ui" / "NxGeneratedProductRegistry.bas"
LAUNCH = ROOT / "src" / "vba" / "frame" / "NxGeneratedLaunchRegistry.bas"
LIMITS_POLICY = ROOT / "src" / "vba" / "core" / "CNxLimitsPolicy.cls"
LIMITS_IDS_BEGIN = "' BEGIN GENERATED RELEASE FEATURE IDS (generate_product_surface.py)"
LIMITS_IDS_END = "' END GENERATED RELEASE FEATURE IDS"

PRODUCT_GROUP_IDS = [
    "NX-GRP-MYEXCEL",
    "NX-GRP-SAVE",
    "NX-GRP-PRINT",
    "NX-GRP-COPY",
    "NX-GRP-INSERT",
    "NX-GRP-FILE",
    "NX-GRP-DATA",
    "NX-GRP-TEMPLATE",
    "NX-GRP-WORKSHEET-TOOLS",
    "NX-GRP-STYLE",
    "NX-GRP-UTIL",
    "NX-GRP-INFO",
]
RIBBON_HIDDEN_COMMAND_GROUP_IDS = {"NX-CMD-GRP-FORMULA-REFERENCE"}
ENTRY_IDS = {
    "NX-ENTRY-START",
    "NX-ENTRY-AI",
    "NX-ENTRY-TEMPLATE",
    "NX-ENTRY-ALL",
    "NX-ENTRY-FAVORITES",
    "NX-ENTRY-MANAGEMENT",
    "NX-ENTRY-FOCUS-SETTINGS",
    "NX-ENTRY-FOCUS-RESET",
    "NX-MGMT-INSTALL",
    "NX-MGMT-REMOVE",
    "NX-MGMT-SAVED-FOLDER",
    "NX-MGMT-INSTALL-FOLDER",
    "NX-MGMT-QAT",
    "NX-MGMT-EXCEL-ADDINS",
    "NX-MGMT-FILE-INFO",
    "NX-MGMT-SHORTCUTS",
    "NX-ENTRY-HANGUL-SETTINGS",
}
DYNAMIC_CALLBACKS = {
    "NxRibbonGetManagementContent",
    "NxRibbonGetFavoritesContent",
    "NxRibbonGetFocusContent",
    "NxRibbonGetAllFunctionsContent",
}
FORBIDDEN_LABELS = {"추천 기능", "최근 실행", "다시 불러오기", "실행·보안 설정", "정보·문제 해결"}
HANGUL_TABLE_SEND_REQUIREMENT = {
    "feature_id": "NX-HANGUL-TABLE-SEND",
    "name": "아래한글 표 전송",
    "reference": "v3.4",
    "status": "implemented",
    "transport": "embedded-HWPX",
}
OWNER_REGISTRARS = {
    "ai": "NxAiRegisterFeatures",
    "template": "NxTemplateRegisterFeatures",
    "data": "NxDataRegisterFeatures",
    "hangul": "NxHangulRegisterFeatures",
    "draw": "NxDrawingRegisterFeatures",
    "file": "NxFileRegisterFeatures",
    "calculator": "NxCalculatorRegisterFeatures",
    "symbols": "NxSymbolsRegisterFeatures",
    "navigator": "NxNavigatorRegisterFeatures",
}
UI_SURFACE_CLASSES = {"direct", "compact", "medium", "guided"}
# Mirror CNxFeatureDefinition runtime checks. A value it rejects breaks every Ribbon click.
RUNTIME_REUSE_MODES = {"direct", "adapt", "spec-only", "clean-room"}
RUNTIME_EXECUTION_GRADES = {"fast", "guarded", "planned"}
NAVIGATION_FIELDS = {
    "ribbon_group_id", "object_scope", "purpose_tags", "launch_surface",
    "mutation_scope", "selection_mode", "feedback_mode", "profile_availability",
}


def load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def release_contract(catalog: dict) -> dict:
    features = catalog.get("features", [])
    feature_ids = [feature.get("id") for feature in features]
    if not feature_ids or len(feature_ids) != len(set(feature_ids)):
        raise ValueError("feature catalog must contain unique feature IDs")
    lineage = catalog.get("lineage")
    if not isinstance(lineage, dict):
        raise ValueError("feature lineage is required")
    release_ids = lineage.get("release_feature_ids")
    if not isinstance(release_ids, list) or len(release_ids) != len(set(release_ids)):
        raise ValueError("release feature IDs must be a unique list")
    if set(release_ids) != set(feature_ids):
        raise ValueError("release features must exactly cover the catalog")
    if lineage.get("hangul_table_send_requirement") != HANGUL_TABLE_SEND_REQUIREMENT:
        raise ValueError("direct Hangul table-send requirement rejected")
    by_id = {feature["id"]: feature for feature in features}
    scoped = deepcopy(catalog)
    scoped["features"] = [deepcopy(by_id[feature_id]) for feature_id in release_ids]
    return scoped


def validate_ui_surfaces(catalog: dict, ui_contract: dict) -> None:
    """Reject ambiguous or oversized launch surfaces before Ribbon generation."""
    if ui_contract.get("schema_version") != 1:
        raise ValueError("UI surface contract version rejected")
    if set(ui_contract.get("classes", {})) != UI_SURFACE_CLASSES:
        raise ValueError("UI surface classes rejected")

    features = catalog.get("features", [])
    feature_ids = {feature.get("id") for feature in features}
    excluded = set(ui_contract.get("excluded_feature_ids", []))
    guided_allowlist = set(ui_contract.get("guided_feature_ids", []))
    if not excluded <= feature_ids or not guided_allowlist <= feature_ids:
        raise ValueError("UI surface feature allowlist contains an unknown feature")
    if excluded != {"NX-AI-IMAGE"}:
        raise ValueError("UI surface exclusions rejected")

    required_fields = set(catalog.get("required_fields", []))
    if "ui_surface" not in required_fields:
        raise ValueError("UI surface must be a required feature field")

    guided_actual: set[str] = set()
    for feature in features:
        feature_id = feature["id"]
        surface = feature.get("ui_surface")
        if feature_id in excluded:
            if surface is not None:
                raise ValueError(f"excluded feature has a UI surface: {feature_id}")
            continue
        if surface not in UI_SURFACE_CLASSES:
            raise ValueError(f"unknown UI surface for {feature_id}: {surface}")
        if surface == "direct":
            if feature.get("launch_surface") != "direct" or feature.get("dialog_template") not in (None, ""):
                raise ValueError(f"direct UI surface cannot open a dialog: {feature_id}")
        if surface == "guided":
            guided_actual.add(feature_id)

    if guided_actual != guided_allowlist:
        raise ValueError("guided UI surface allowlist mismatch")


def surfaces(shell: dict) -> dict:
    value = shell.get("surfaces")
    if not isinstance(value, dict) or set(value) != {"home", "product"}:
        raise ValueError("shell must declare only canonical home and product surfaces")
    return value


def feature_control(feature: dict, placement: dict) -> dict:
    ribbon = feature["ribbon"]
    feature_id = feature["id"]
    control = {
        "kind": "button",
        "id": ribbon.get("control_id") or feature_id,
        "target": feature_id,
        "role": ribbon.get("placement", "small"),
        "label": ribbon.get("label") or feature_id,
        "tag": f"nx1|feature|{feature_id}",
        "keytip": ribbon.get("keytip") or feature_id[-1],
        "size": ribbon.get("size", "small"),
        "callback": ribbon.get("callback", "NxRibbonExecute"),
        "control_type": ribbon.get("control_type", 2),
        "image_mso": ribbon.get("image_mso"),
        "screentip": ribbon.get("label") or feature_id,
        "supertip": feature.get("description_ko") or ribbon.get("label") or feature_id,
    }
    control.update({key: deepcopy(value) for key, value in placement.items() if key != "feature_id"})
    if "keytip" not in placement:
        control["keytip"] = None
    return control


def command_control(command: dict) -> dict:
    command_id = str(command["id"])
    return {
        "kind": "button",
        "id": f"NX-RIBBON-{command_id}",
        "target": command_id,
        "label": command["label_ko"],
        "tag": f"nx1|command|{command_id}",
        "callback": "NxRibbonExecute",
        "image_mso": command["image_mso"],
        "screentip": command["label_ko"],
        "supertip": command["description_ko"],
        "control_type": 2,
    }


def expand_groups(contract: dict, command_contract: dict, groups: list[dict], navigation_items: list[dict]) -> list[dict]:
    by_id = {feature["id"]: feature for feature in contract["features"]}
    commands_by_id = {str(command["id"]): command for command in command_contract["commands"]}
    commands_by_group: dict[str, list[dict]] = {}
    for command in command_contract["commands"]:
        if not command.get("taskpane_visible"):
            continue
        commands_by_group.setdefault(str(command["group_id"]), []).append(command)
    for commands in commands_by_group.values():
        commands.sort(key=lambda command: (int(command["sort_order"]), str(command["id"])))

    def expand(control: dict) -> dict:
        if "catalog_category_id" in control:
            result = deepcopy(control)
            category_id = result.pop("catalog_category_id")
            control_ids = result.pop("catalog_control_ids", {})
            rows = [row for row in navigation_items if row["category_id"] == category_id]
            if not set(control_ids) <= {row["id"] for row in rows}:
                raise ValueError(f"unknown catalog control override: {category_id}")
            if not rows:
                raise ValueError(f"empty catalog category: {category_id}")

            def catalog_button(row: dict) -> dict:
                if row["item_type"] == "feature":
                    button = feature_control(by_id[row["id"]], {})
                else:
                    button = command_control(commands_by_id[row["id"]])
                button["id"] = control_ids.get(row["id"], result["id"] + "-" + row["id"]) if row["item_type"] == "feature" else control_ids.get(row["id"], "NX-RIBBON-" + row["id"])
                button.pop("keytip", None)
                button["size"] = "small"
                return button

            result["kind"] = "menu"
            menu = []
            for row, children in category_menu_rows(rows):
                if children is None:
                    menu.append(catalog_button(row))
                else:
                    menu.append({"kind": "menu", "id": result["id"] + "-" + row["group_id"],
                                 "label": row["group_label"], "image_mso": row["image_mso"],
                                 "menu": [catalog_button(child) for child in children]})
            # Administrative and built-in Excel entries are explicit supplements.
            result["menu"] = menu + list(result.get("menu", []))
        elif "inline_command_group_id" in control:
            group_id = str(control["inline_command_group_id"])
            commands = commands_by_group.get(group_id, [])
            if not commands:
                raise ValueError(f"empty inline command group: {group_id}")
            result = deepcopy(control)
            result.pop("inline_command_group_id", None)
            result.pop("command_group_id", None)
            excluded = {str(value) for value in result.pop("inline_exclude_command_ids", [])}
            existing = list(result.get("menu", []))
            result["kind"] = "menu"
            result["menu"] = existing + [
                command_control(command)
                for command in commands
                if str(command["id"]) not in excluded
            ]
        elif "feature_id" in control:
            feature_id = control["feature_id"]
            if feature_id not in by_id:
                raise ValueError(f"unknown shell feature: {feature_id}")
            result = feature_control(by_id[feature_id], control)
        elif "command_id" in control:
            command_id = str(control["command_id"])
            if command_id not in commands_by_id:
                raise ValueError(f"unknown shell command: {command_id}")
            result = command_control(commands_by_id[command_id])
            result.update({
                key: deepcopy(value)
                for key, value in control.items()
                if key not in {"command_id", "kind"}
            })
            result["kind"] = "button"
        elif "command_group_id" in control:
            group_id = str(control["command_group_id"])
            commands = commands_by_group.get(group_id, [])
            if not commands:
                raise ValueError(f"empty command group: {group_id}")
            result = deepcopy(control)
            result.pop("command_group_id", None)
            excluded = set(result.pop("exclude_command_ids", []))
            result["kind"] = "menu"
            result["menu"] = [command_control(command) for command in commands
                              if command["id"] not in excluded] + list(result.get("menu", []))
        else:
            result = deepcopy(control)
        if "menu" in result:
            result["menu"] = [expand(item) for item in result["menu"]]
        if "box" in result:
            result["box"] = [expand(item) for item in result["box"]]
        if "toggle" in result:
            result["toggle"] = expand(result["toggle"])
        return result

    expanded = deepcopy(groups)
    for group in expanded:
        group["controls"] = [expand(control) for control in group.get("controls", [])]
    return expanded


def expand_shell(contract: dict, shell: dict, command_contract: dict) -> dict:
    expanded = deepcopy(shell)
    navigation_items = build_document(contract, command_contract, shell)["items"]
    for surface in surfaces(expanded).values():
        surface["groups"] = expand_groups(contract, command_contract, surface.get("groups", []), navigation_items)
    return expanded


def iter_controls(items: list[dict]):
    for item in items:
        yield item
        yield from iter_controls(item.get("menu", []))
        yield from iter_controls(item.get("box", []))
        if item.get("toggle"):
            yield from iter_controls([item["toggle"]])


def surface_controls(surface: dict) -> list[dict]:
    return list(iter_controls([control for group in surface.get("groups", []) for control in group.get("controls", [])]))


def validate(contract: dict, shell: dict, command_contract: dict) -> None:
    features = contract.get("features", [])
    feature_ids = [feature["id"] for feature in features]
    feature_set = set(feature_ids)
    if command_contract.get("schema_version") != 2:
        raise ValueError("command contract schema rejected")
    commands = command_contract.get("commands", [])
    command_ids = [str(command.get("id")) for command in commands]
    command_set = set(command_ids)
    visible_command_set = {
        str(command["id"])
        for command in commands
        if command.get("taskpane_visible")
    }
    if len(command_ids) != command_contract.get("normalized_command_count") or len(command_set) != len(command_ids):
        raise ValueError("command inventory rejected")
    command_group_ids = [str(group.get("id")) for group in command_contract.get("groups", [])]
    if len(command_group_ids) != 14 or len(set(command_group_ids)) != 14:
        raise ValueError("command group inventory rejected")
    if len(feature_ids) != len(feature_set):
        raise ValueError("feature contract must contain unique feature IDs")
    required = {
        "id",
        "owner",
        "base_execution_grade",
        "planned_cell_threshold",
        "privacy",
        "ribbon",
        "ui_surface",
        "navigation",
    }
    if not required.issubset(contract.get("required_fields", [])):
        raise ValueError("feature contract required fields rejected")
    planned_cell_limit = load(EXECUTION_GRADES).get("planned_cell_threshold")
    if not isinstance(planned_cell_limit, int) or isinstance(planned_cell_limit, bool) or planned_cell_limit <= 0:
        raise ValueError("execution grade planned cell threshold rejected")
    for feature in features:
        dialog = [bool(feature.get(key)) for key in ("dialog_id", "dialog_template", "dialog_variant")]
        resolver = bool(feature.get("state_resolver"))
        launch_valid = {
            "direct": not any(dialog) and not resolver,
            "dialog": all(dialog) and not resolver,
            "workbench": all(dialog) and not resolver,
            "hybrid": all(dialog) and resolver,
        }.get(feature.get("launch_surface"), False)
        if not launch_valid:
            raise ValueError(f"feature launch metadata rejected: {feature['id']}")
        if feature.get("reuse_mode") not in RUNTIME_REUSE_MODES:
            raise ValueError(f"feature reuse mode rejected: {feature['id']}")
        if feature.get("base_execution_grade") not in RUNTIME_EXECUTION_GRADES:
            raise ValueError(f"feature execution grade rejected: {feature['id']}")
        planned_cell_threshold = feature.get("planned_cell_threshold")
        if (
            not isinstance(planned_cell_threshold, int)
            or isinstance(planned_cell_threshold, bool)
            or planned_cell_threshold <= 0
            or planned_cell_threshold > planned_cell_limit
        ):
            raise ValueError(f"feature planned cell threshold rejected: {feature['id']}")
        navigation = feature.get("navigation")
        navigation_keys = set(navigation) if isinstance(navigation, dict) else set()
        allowed_navigation_keys = NAVIGATION_FIELDS
        if feature.get("id") == "NX-HANGUL-TABLE-SEND":
            allowed_navigation_keys = NAVIGATION_FIELDS | {"transport"}
        if navigation_keys != allowed_navigation_keys:
            raise ValueError(f"feature navigation fields rejected: {feature['id']}")
        if feature.get("id") == "NX-HANGUL-TABLE-SEND" and navigation.get("transport") != "embedded-HWPX":
            raise ValueError("Hangul table transfer transport must be HWPX")
        if feature["id"] == "NX-UTIL-DOCUMENT-NAVIGATOR":
            if (navigation["ribbon_group_id"], feature["ribbon"]["workflow_group_id"]) != (
                "NX-GRP-SHEET", "NX-GRP-WORKSHEET-TOOLS"
            ):
                raise ValueError("document navigator navigation and physical Ribbon group mismatch")
        elif navigation["ribbon_group_id"] != feature["ribbon"]["workflow_group_id"]:
            raise ValueError(f"feature navigation Ribbon group mismatch: {feature['id']}")
        for field in ("object_scope", "purpose_tags"):
            values = navigation[field]
            if not isinstance(values, list) or not values or any(not isinstance(value, str) or not value for value in values):
                raise ValueError(f"feature navigation {field} rejected: {feature['id']}")
        if navigation["launch_surface"] != feature["launch_surface"]:
            raise ValueError(f"feature navigation launch surface mismatch: {feature['id']}")
        if navigation["selection_mode"] not in {"none", "optional", "required"}:
            raise ValueError(f"feature navigation selection mode rejected: {feature['id']}")
        if navigation["feedback_mode"] not in {"inline", "completion"}:
            raise ValueError(f"feature navigation feedback mode rejected: {feature['id']}")
        if navigation["profile_availability"] != ["internal-xlam", "enhanced-dll"]:
            raise ValueError(f"feature navigation profile availability rejected: {feature['id']}")
    if shell.get("schema_version") != 3 or "groups" in shell or "tab" in shell:
        raise ValueError("legacy Ribbon shell shape is rejected")

    raw_surfaces = surfaces(shell)
    def command_group_refs(items: list[dict]) -> list[str]:
        refs: list[str] = []
        for item in items:
            if item.get("command_group_id"):
                refs.append(str(item["command_group_id"]))
            if item.get("inline_command_group_id"):
                refs.append(str(item["inline_command_group_id"]))
            refs.extend(command_group_refs(item.get("menu", [])))
            refs.extend(command_group_refs(item.get("box", [])))
            if item.get("toggle"):
                refs.extend(command_group_refs([item["toggle"]]))
        return refs

    raw_command_groups = command_group_refs(
        [control for group in raw_surfaces["product"].get("groups", []) for control in group.get("controls", [])]
    )
    if raw_command_groups:
        raise ValueError("category menus must use the shared navigation catalog")
    catalog_categories = [control["catalog_category_id"] for control in surface_controls(raw_surfaces["product"]) if "catalog_category_id" in control]
    expected_catalog_categories = {row["id"] for row in shell["navigation_categories"]} - {"NX-GRP-FORMULA", "NX-GRP-TEMPLATE"}
    template_group = next(group for group in raw_surfaces["product"]["groups"] if group["id"] == "NX-GRP-TEMPLATE")
    template_controls = template_group["controls"][0].get("box", [])
    if len(template_controls) != 2 or template_controls[0].get("kind") != "button" or template_controls[0].get("feature_id") != "NX-TPL-LIST" or template_controls[1].get("id") != "NX-FUNCTION-MENU":
        raise ValueError("Template category must directly open its single manager")
    if len(catalog_categories) != len(set(catalog_categories)) or set(catalog_categories) != expected_catalog_categories:
        raise ValueError("Ribbon catalog category coverage rejected")
    if any(control.get("command_group_id") for control in surface_controls(raw_surfaces["home"])):
        raise ValueError("Home surface must not duplicate the command inventory")
    home = raw_surfaces["home"]
    product = raw_surfaces["product"]
    if home.get("idMso") != "TabHome" or [group.get("id") for group in home.get("groups", [])] != ["NX-HOME-GRP-MYEXCEL"]:
        raise ValueError("Home surface identity rejected")
    if home["groups"][0].get("insert_before_mso") != "GroupClipboard":
        raise ValueError("Home MyExcel group must precede GroupClipboard")
    if product.get("id") != "NX-TAB" or [group.get("id") for group in product.get("groups", [])] != PRODUCT_GROUP_IDS:
        raise ValueError("product group order rejected")
    expected_group_icons = {
        "NX-GRP-MYEXCEL": "ControlProperties",
        "NX-GRP-SAVE": "FileSave",
        "NX-GRP-PRINT": "PrintPreviewAndPrint",
        "NX-GRP-COPY": "Copy",
        "NX-GRP-INSERT": "PictureInsertFromFile",
        "NX-GRP-FILE": "FileOpen",
        "NX-GRP-TEMPLATE": "FileNew",
        "NX-GRP-DATA": "TableInsert",
        "NX-GRP-WORKSHEET-TOOLS": "ViewNormalViewExcel",
        "NX-GRP-STYLE": "TableAutoFormat",
        "NX-GRP-UTIL": "AddInManager",
        "NX-GRP-INFO": "Info",
    }
    if {group["id"]: group.get("image_mso") for group in product["groups"]} != expected_group_icons:
        raise ValueError("product collapsed-group icon contract rejected")
    if home["groups"][0].get("image_mso") != expected_group_icons["NX-GRP-MYEXCEL"]:
        raise ValueError("Home collapsed-group icon contract rejected")

    expanded = expand_shell(contract, shell, command_contract)
    for surface_name, surface in surfaces(expanded).items():
        seen_ids: set[str] = set()
        for control in surface_controls(surface):
            kind = control.get("kind")
            if kind not in {"button", "toggle_button", "split_button", "builtin_button", "menu", "dynamic_menu", "box", "inline_command_group"}:
                raise ValueError(f"unknown Ribbon control kind on {surface_name}: {kind}")
            control_id = control.get("id")
            if kind != "builtin_button":
                if not control_id:
                    raise ValueError(f"custom Ribbon control ID missing on {surface_name}")
                if control_id in seen_ids:
                    raise ValueError(f"duplicate Ribbon control ID on {surface_name}: {control_id}")
                seen_ids.add(control_id)

            if control.get("control_type") == 1:
                if kind not in {"menu", "dynamic_menu"} or control.get("size") != "large":
                    raise ValueError("type 1 controls must be large list menus")
            if control.get("control_type") == 3:
                if kind != "button" or control.get("show_label") is not False:
                    raise ValueError("type 3 controls must be icon-only buttons")
                if not all(control.get(field) for field in ("image_mso", "screentip", "supertip")):
                    raise ValueError("type 3 controls require image and tooltips")
            if control.get("image") is not None:
                if control["image"] not in QUICK_FORMAT_IMAGES or control.get("control_type") != 3:
                    raise ValueError("custom Ribbon image is not an approved quick-format button")

            if kind == "dynamic_menu":
                if control.get("get_content") not in DYNAMIC_CALLBACKS:
                    raise ValueError("dynamic list callback rejected")
                if any(key in control for key in ("menu", "box", "target", "tag", "callback")):
                    raise ValueError("dynamic list must not expose a static target")
            elif kind == "menu":
                if not control.get("menu"):
                    raise ValueError("static list menu must contain living actions")
                if any(key in control for key in ("target", "tag", "callback", "get_content")):
                    raise ValueError("static menu is a container, not an executable target")
            elif kind == "box":
                if control.get("box_style") not in {"horizontal", "vertical"} or not control.get("box"):
                    raise ValueError("Ribbon boxes must be non-empty horizontal or vertical containers")
                if any(key in control for key in ("target", "tag", "callback", "get_content")):
                    raise ValueError("Ribbon box is a container, not an executable target")
            elif kind == "builtin_button":
                if control.get("element", "button") not in {"button", "toggleButton"}:
                    raise ValueError("unsupported native control element")
                if not control.get("id_mso") or any(key in control for key in ("id", "target", "tag", "callback")):
                    raise ValueError("built-in buttons must use only idMso execution")
            elif kind == "split_button":
                toggle = control.get("toggle")
                menu_id = control.get("menu_id")
                if not isinstance(toggle, dict) or toggle.get("kind") != "toggle_button":
                    raise ValueError("split button must contain one toggle button")
                if not menu_id or menu_id in seen_ids or not control.get("menu"):
                    raise ValueError("split button menu contract rejected")
                seen_ids.add(menu_id)
                if any(key in control for key in ("target", "tag", "callback", "get_content")):
                    raise ValueError("split button is a container, not an executable target")
            elif kind == "toggle_button":
                target = control.get("target")
                tag = control.get("tag")
                if control.get("callback") != "NxRibbonToggleFocus" or control.get("get_pressed") != "NxRibbonGetFocusPressed":
                    raise ValueError("focus toggle callback contract rejected")
                parts = str(tag).split("|")
                if target != "NX-DATA-FOCUS-CELL" or parts != ["nx1", "feature", target]:
                    raise ValueError("focus toggle target contract rejected")
            else:
                target = control.get("target")
                tag = control.get("tag")
                if not target or not tag or control.get("callback") != "NxRibbonExecute":
                    raise ValueError("custom button routing is incomplete")
                parts = tag.split("|")
                if len(parts) != 3 or parts[0] != "nx1" or parts[2] != target:
                    raise ValueError("custom button tag/target mismatch")
                if parts[1] == "feature" and target not in feature_set:
                    raise ValueError(f"unknown feature target: {target}")
                if parts[1] == "entry" and target not in ENTRY_IDS:
                    raise ValueError(f"unknown entry target: {target}")
                if parts[1] == "command" and target not in command_set:
                    raise ValueError(f"unknown command target: {target}")
                if parts[1] not in {"feature", "entry", "command"}:
                    raise ValueError("static category routing is rejected")

        def validate_keytips(items: list[dict]) -> None:
            keytips = [str(item["keytip"]) for item in items if item.get("keytip")]
            if len(keytips) != len(set(keytips)):
                raise ValueError(f"duplicate sibling keytip on {surface_name}")
            for item in items:
                validate_keytips(item.get("menu", []))
                validate_keytips(item.get("box", []))
                if item.get("toggle"):
                    validate_keytips([item["toggle"]])

        for group in surface.get("groups", []):
            validate_keytips(group.get("controls", []))

    expected_labels_by_surface = {
        "home": ["내엑셀", "즐겨찾기", "포커스셀", "전체기능"],
        "product": ["내엑셀", "즐겨찾기", "포커스셀", "전체기능"],
    }
    expected_kinds = ["dynamic_menu", "dynamic_menu", "dynamic_menu", "dynamic_menu"]
    for surface_name in ("home", "product"):
        identity = surfaces(expanded)[surface_name]["groups"][0]["controls"]
        if [item.get("label") for item in identity] != expected_labels_by_surface[surface_name] or [item.get("kind") for item in identity] != expected_kinds:
            raise ValueError("identity group must repeat the approved four controls")
    home_identity = surfaces(expanded)["home"]["groups"][0]["controls"]
    product_identity = surfaces(expanded)["product"]["groups"][0]["controls"]
    for index, (home_control, product_control) in enumerate(zip(home_identity, product_identity, strict=True)):
        for field in ("kind", "label", "image_mso", "get_content", "keytip", "control_type"):
            if home_control.get(field) != product_control.get(field):
                raise ValueError(f"home/product identity control {index} must share {field}")

    expected_groups = list(
        zip(
            PRODUCT_GROUP_IDS,
            ["내엑셀", "저장", "인쇄", "복붙", "삽입", "파일관리", "데이터", " ", " ", "스타일", "추가기능", "정보진단"],
        )
    )
    product_groups = surfaces(expanded)["product"]["groups"]
    if [(group.get("id"), group.get("label")) for group in product_groups] != expected_groups:
        raise ValueError("approved product group labels rejected")

    def group_targets(group_id: str) -> set[str]:
        group = next(item for item in product_groups if item["id"] == group_id)
        return {
            str(control.get("target"))
            for control in iter_controls(group.get("controls", []))
            if control.get("target")
        }

    if "NX-UTIL-SYMBOLS" in group_targets("NX-GRP-INSERT"):
        raise ValueError("symbols must not appear in Insert")
    if "NX-DATA-FOCUS-CELL" in group_targets("NX-GRP-DATA"):
        raise ValueError("focus cell must not appear in Data")
    if group_targets("NX-GRP-STYLE") & {"NX-DRAW-INSERT-PICTURE", "NX-DRAW-FIT-PICTURE"}:
        raise ValueError("picture commands must not be duplicated in Style")
    required_utilities = {
        "NX-UTIL-CALCULATOR",
        "NX-UTIL-SYMBOLS",
        "NX-UTIL-NAVIGATOR",
        "NX-MGMT-SHORTCUTS",
        "NX-ENTRY-AI",
    }
    if not required_utilities <= group_targets("NX-GRP-UTIL"):
        raise ValueError("utility group is missing an approved tool")

    all_function_lists = [
        control
        for surface in surfaces(expanded).values()
        for control in surface_controls(surface)
        if control.get("label") == "전체기능"
    ]
    if len(all_function_lists) != 2 or any(
        control.get("get_content") != "NxRibbonGetAllFunctionsContent" for control in all_function_lists
    ):
        raise ValueError("all-functions dynamic lists must be live on both surfaces")
    focus_lists = [
        control
        for surface in surfaces(expanded).values()
        for control in surface_controls(surface)
        if control.get("id") in {"NX-HOME-FOCUS", "NX-PROD-FOCUS"}
    ]
    if len(focus_lists) != 2 or any(
        control.get("kind") != "dynamic_menu" or control.get("get_content") != "NxRibbonGetFocusContent"
        for control in focus_lists
    ):
        raise ValueError("focus dynamic menus must be live on both surfaces")

    save = next(group for group in surfaces(expanded)["product"]["groups"] if group["id"] == "NX-GRP-SAVE")
    save_menu = next((item for item in save["controls"] if item.get("id") == "NX-SAVE-MENU"), None)
    if (
        save_menu is None
        or save_menu.get("kind") != "menu"
        or save_menu.get("size") != "large"
        or save_menu.get("label") != "저장"
        or save_menu.get("image_mso") != "FileSave"
        or len(save_menu.get("menu", [])) != 7
    ):
        raise ValueError("integrated Save menu rejected")
    if any(item.get("id") == "NX-SAVE-PDF-MENU" for item in save["controls"]):
        raise ValueError("standalone PDF menu rejected")
    save_box = next(item for item in save["controls"] if item["id"] == "NX-SAVE-VERTICAL")
    if [item.get("target") for item in save_box["box"]] != [
        "NX-FILE-MANNER-SAVE",
        "NX-FILE-SHEET-COPY-SAVE",
        "NX-FILE-RANGE-COPY-SAVE",
    ]:
        raise ValueError("save vertical order rejected")
    if [item.get("image_mso") for item in save_box["box"]] != [
        "FileSaveAs",
        "FileSaveAsExcelXlsx",
        "FileSaveAsOtherFormats",
    ] or any(item.get("control_type") != 2 or item.get("show_label") is False for item in save_box["box"]):
        raise ValueError("save vertical icon-title contract rejected")

    product_controls = surface_controls(surfaces(expanded)["product"])
    quick_boxes = [
        item for item in product_controls if item.get("id") == "NX-STYLE-QUICK-VERTICAL"
    ]
    if len(quick_boxes) != 1 or quick_boxes[0].get("box_style") != "vertical":
        raise ValueError("quick-format controls require one vertical container")
    quick_rows = quick_boxes[0].get("box", [])
    if len(quick_rows) != 3 or any(
        row.get("kind") != "box" or row.get("box_style") != "horizontal" or len(row.get("box", [])) != 6
        for row in quick_rows
    ):
        raise ValueError("quick-format controls require three rows of six buttons")
    quick_buttons = [button for row in quick_rows for button in row["box"][:5]]
    alignment_ids = ["AlignLeft", "AlignCenter", "AlignRight"]
    for row, mso in zip(quick_rows, alignment_ids, strict=True):
        button = row["box"][5]
        if button.get("kind") != "builtin_button" or button.get("id_mso") != mso or button.get("show_label") is not False:
            raise ValueError("alignment column must use icon-only native Excel commands")
    images = list(QUICK_FORMAT_IMAGES)
    expected_quick_images = images[:4] + [images[12]] + images[4:8] + [images[13]] + images[8:12] + [images[14]]
    number_routes = {
        "nx-quick-number-percent": "NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-PERCENTAGE",
        "nx-quick-number-comma": "NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-NUMBER",
        "nx-quick-number-decimal": "NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-DECIMAL",
    }
    if [button.get("image") for button in quick_buttons] != expected_quick_images or any(
        button.get("target") != number_routes.get(image, "NX-CMD-STYLE-" + image.removeprefix("nx-quick-").upper())
        or button.get("control_type") != 3 or button.get("show_label") is not False
        for button, image in zip(quick_buttons, expected_quick_images, strict=True)
    ):
        raise ValueError("quick-format button image, route or order rejected")
    style_group = next(group for group in surfaces(expanded)["product"]["groups"] if group["id"] == "NX-GRP-STYLE")
    if quick_boxes[0] not in style_group["controls"]:
        raise ValueError("quick-format controls must stay in the product Style group")
    custom_images = [
        item.get("image") for surface in surfaces(expanded).values()
        for item in surface_controls(surface) if item.get("image")
    ]
    if custom_images != expected_quick_images:
        raise ValueError("custom Ribbon images must occur only on the 15 quick-format buttons")
    static_feature_targets = [
        item.get("target")
        for item in product_controls
        if item.get("tag", "").startswith("nx1|feature|")
    ]
    # Focus is rendered at runtime by NxRibbonGetFocusContent so that Excel
    # gives it the same compact title-arrow geometry as adjacent dynamic menus.
    static_feature_targets.append("NX-DATA-FOCUS-CELL")
    static_features = set(static_feature_targets)
    expected_static = {feature["id"] for feature in features if feature["owner"] != "ai" and feature.get("catalog_visible")}
    if static_features != expected_static:
        raise ValueError("all non-AI features must be reachable from the product Ribbon")
    # Representative buttons intentionally repeat the same route in their category menu.
    if any(target.startswith("NX-AI-") for target in static_features):
        raise ValueError("AI presets must remain internal to the workbench")
    if sum(item.get("target") == "NX-ENTRY-AI" for item in product_controls) != 2:
        raise ValueError("AI workbench must appear in the category menu and representative button")

    def raw_excluded_command_ids(items: list[dict]) -> set[str]:
        excluded: set[str] = set()
        for item in items:
            excluded.update(str(value) for value in item.get("inline_exclude_command_ids", []))
            excluded.update(raw_excluded_command_ids(item.get("menu", [])))
            excluded.update(raw_excluded_command_ids(item.get("box", [])))
            if item.get("toggle"):
                excluded.update(raw_excluded_command_ids([item["toggle"]]))
        return excluded

    raw_product_controls = [
        control
        for group in raw_surfaces["product"].get("groups", [])
        for control in group.get("controls", [])
    ]
    direct_command_ids = {
        str(item["command_id"])
        for item in iter_controls(raw_product_controls)
        if item.get("command_id")
    }
    excluded_command_ids = raw_excluded_command_ids(raw_product_controls) - direct_command_ids
    hidden_command_ids = {
        str(command["id"])
        for command in commands
        if str(command["group_id"]) in RIBBON_HIDDEN_COMMAND_GROUP_IDS
    }
    expected_ribbon_commands = (visible_command_set - hidden_command_ids - excluded_command_ids) | direct_command_ids
    ribbon_command_targets = [
        str(item.get("target"))
        for item in product_controls
        if str(item.get("tag", "")).startswith("nx1|command|")
    ]
    if set(ribbon_command_targets) != expected_ribbon_commands:
        raise ValueError("Ribbon command inventory rejected")
    # Route duplicates are allowed; XML control identities remain unique above.

    labels = {str(item.get("label")) for surface in surfaces(expanded).values() for item in surface_controls(surface)}
    if labels & FORBIDDEN_LABELS:
        raise ValueError("forbidden recommendation/recent/support surface detected")
    for feature in features:
        ribbon = feature.get("ribbon")
        if not isinstance(ribbon, dict) or ribbon.get("launch_mode") != "feature":
            raise ValueError("feature launch metadata rejected")


def attrs(control: dict, *, include_label: bool = True) -> list[str]:
    if control.get("kind") == "builtin_button":
        result = [f'idMso="{escape(str(control["id_mso"]))}"']
    else:
        result = [f'id="{escape(str(control["id"]))}"']
    if include_label and control.get("label") is not None and control.get("show_label", True):
        result.append(f'label="{escape(str(control["label"]))}"')
    if control.get("tag"):
        result.append(f'tag="{escape(str(control["tag"]))}"')
        if str(control["tag"]).startswith(("nx1|feature|", "nx1|command|")):
            result.append('getEnabled="NxRibbonGetEnabled"')
    if control.get("keytip"):
        result.append(f'keytip="{escape(str(control["keytip"]))}"')
    brand_key = BRAND_IMAGES["tags"].get(control.get("tag")) or BRAND_IMAGES["labels"].get(control.get("label"))
    if control.get("image"):
        result.append(f'image="{escape(str(control["image"]))}"')
    elif brand_key:
        image_size = 32 if control.get("size") == "large" or control.get("control_type") == 1 else 16
        result.append(f'image="NxBrand_{brand_key.replace("-", "_")}_{image_size}"')
    elif control.get("image_mso"):
        result.append(f'imageMso="{escape(str(control["image_mso"]))}"')
    if control.get("screentip"):
        result.append(f'screentip="{escape(str(control["screentip"]))}"')
    if control.get("supertip"):
        result.append(f'supertip="{escape(str(control["supertip"]))}"')
    if control.get("size") == "large" or control.get("control_type") == 1:
        result.append('size="large"')
    return result


def render_control(control: dict, indent: str) -> list[str]:
    kind = control["kind"]
    if kind == "builtin_button":
        values = attrs(control)
        if control.get("show_label") is False:
            values.append('showLabel="false"')
        element = control.get("element", "button")
        return [f'{indent}<{element} {" ".join(values)} />']
    if kind == "dynamic_menu":
        values = attrs(control)
        values.append(f'getContent="{escape(str(control["get_content"]))}"')
        return [f'{indent}<dynamicMenu {" ".join(values)} />']
    if kind == "box":
        lines = [f'{indent}<box {" ".join(attrs(control, include_label=False))} boxStyle="{escape(control["box_style"])}">']
        for child in control["box"]:
            lines.extend(render_control(child, indent + "  "))
        return lines + [f"{indent}</box>"]
    if kind == "menu":
        values = attrs(control)
        if control.get("show_label") is False:
            values.append(f'label="{escape(str(control["label"]))}"')
            values.append('showLabel="false"')
        if control.get("show_image") is False:
            values.append('showImage="false"')
        lines = [f'{indent}<menu {" ".join(values)}>']
        for child in control["menu"]:
            lines.extend(render_control(child, indent + "  "))
        return lines + [f"{indent}</menu>"]
    if kind == "split_button":
        lines = [f'{indent}<splitButton {" ".join(attrs(control, include_label=False))}>']
        lines.extend(render_control(control["toggle"], indent + "  "))
        # Preserve a name for assistive technology on any remaining split
        # buttons. Focus Cell uses a standard menu so its title-arrow spacing,
        # accessibility, and collapsed-group behavior match neighboring menus.
        lines.append(
            f'{indent}  <menu id="{escape(str(control["menu_id"]))}" '
            f'label="{escape(str(control.get("label", control["menu_id"]))) }">'
        )
        for child in control["menu"]:
            lines.extend(render_control(child, indent + "    "))
        lines += [f"{indent}  </menu>", f"{indent}</splitButton>"]
        return lines
    if kind == "toggle_button":
        values = attrs(control)
        values.append(f'onAction="{escape(str(control["callback"]))}"')
        values.append(f'getPressed="{escape(str(control["get_pressed"]))}"')
        return [f'{indent}<toggleButton {" ".join(values)} />']
    values = attrs(control, include_label=control.get("control_type") != 3)
    values.append(f'onAction="{escape(str(control["callback"]))}"')
    if control.get("show_label") is False or control.get("control_type") == 3:
        if control.get("label") is not None:
            values.append(f'label="{escape(str(control["label"]))}"')
        values.append('showLabel="false"')
    return [f'{indent}<button {" ".join(values)} />']


def render_xml(shell: dict) -> str:
    lines = [
        '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="NxRibbonOnLoad">',
        "  <ribbon>",
        "    <tabs>",
    ]
    for surface_name in ("home", "product"):
        surface = surfaces(shell)[surface_name]
        tab_attrs = f'idMso="{escape(surface["idMso"])}"' if surface.get("idMso") else f'id="{escape(surface["id"])}" label="{escape(surface["label"])}"'
        lines.append(f"      <tab {tab_attrs}>")
        for group in surface["groups"]:
            group_attrs = [f'id="{escape(group["id"])}"', f'label="{escape(group["label"])}"']
            if group.get("image_mso"):
                group_attrs.append(f'imageMso="{escape(str(group["image_mso"]))}"')
            if group.get("insert_before_mso"):
                group_attrs.append(f'insertBeforeMso="{escape(str(group["insert_before_mso"]))}"')
            lines.append(f'        <group {" ".join(group_attrs)}>')
            for control in group["controls"]:
                lines.extend(render_control(control, "          "))
            lines.append("        </group>")
        lines.append("      </tab>")
    lines += ["    </tabs>", "  </ribbon>", "</customUI>", ""]
    return "\n".join(lines)


def render_vba(contract: dict, shell: dict) -> str:
    product = surfaces(shell)["product"]
    lines = [
        'Attribute VB_Name = "NxGeneratedProductRegistry"',
        "Option Explicit",
        "' GENERATED FILE - edit contracts/feature-contract.json and ribbon-shell-contract.json",
    ]
    lines.extend("' Feature: " + feature["id"] for feature in contract["features"])
    lines += [
        "Public Function NxCreateGeneratedProductRegistry() As CNxFeatureRegistry",
        "    Dim registry As CNxFeatureRegistry",
        "    Set registry = CreateBaseRegistry()",
    ]
    owners = list(dict.fromkeys(feature["owner"] for feature in contract["features"]))
    for owner in owners:
        registrar = OWNER_REGISTRARS.get(owner)
        if not registrar:
            raise ValueError(f"unknown feature owner: {owner}")
        lines.append(f"    {registrar} registry")
    lines += [
        "    Set NxCreateGeneratedProductRegistry = registry",
        "End Function",
        "",
        "Public Function NxProductSurfaceGroupCount() As Long",
        f"    NxProductSurfaceGroupCount = {len(product['groups'])}",
        "End Function",
        "",
        "Public Function NxProductSurfaceFeatureCount() As Long",
        f"    NxProductSurfaceFeatureCount = {len(contract['features'])}",
        "End Function",
        "",
        "Public Function NxGeneratedReleaseFeatureIds() As Variant",
    ]
    ids = [feature["id"] for feature in contract["features"]]
    for offset in range(0, len(ids), 8):
        chunk = ", ".join(f'"{feature_id}"' for feature_id in ids[offset:offset + 8])
        prefix = "    NxGeneratedReleaseFeatureIds = Array(" if offset == 0 else "        "
        suffix = ")" if offset + 8 >= len(ids) else ", _"
        lines.append(prefix + chunk + suffix)
    lines += ["End Function", ""]
    return "\n".join(lines)


def render_limits_release_ids(feature_ids: list[str], newline: str) -> str:
    lines = ["    ids = Array( _"]
    for offset in range(0, len(feature_ids), 5):
        chunk = ", ".join(f'"{feature_id}"' for feature_id in feature_ids[offset:offset + 5])
        suffix = ")" if offset + 5 >= len(feature_ids) else ", _"
        lines.append("        " + chunk + suffix)
    return newline.join(lines)


def project_limits_release_ids(source: str, feature_ids: list[str]) -> str:
    if source.count(LIMITS_IDS_BEGIN) != 1 or source.count(LIMITS_IDS_END) != 1:
        raise ValueError("limits release feature ID markers must occur exactly once")
    if "\r\n" in source:
        source = source.replace("\r\n", "\n").replace("\n", "\r\n")
    before, marked = source.split(LIMITS_IDS_BEGIN, 1)
    _, after = marked.split(LIMITS_IDS_END, 1)
    newline = "\r\n" if "\r\n" in source else "\n"
    generated = render_limits_release_ids(feature_ids, newline)
    return before + LIMITS_IDS_BEGIN + newline + generated + newline + "    " + LIMITS_IDS_END + after


def render_launch(contract: dict) -> str:
    def value(raw: object) -> str:
        if raw is None or raw == "":
            return "vbNullString"
        return '"' + str(raw).replace('"', '""') + '"'

    lines = [
        'Attribute VB_Name = "NxGeneratedLaunchRegistry"',
        "Option Explicit",
        "' GENERATED FILE - edit contracts/feature-contract.json",
        "Public Sub NxConfigureGeneratedLaunch(ByVal definition As CNxFeatureDefinition, ByVal featureId As String)",
        '    If definition Is Nothing Then NxRaiseContractError "Feature definition is required"',
        '    If Len(featureId) = 0 Then NxRaiseContractError "Feature ID is required"',
        "    Select Case featureId",
    ]
    for feature in contract["features"]:
        lines.extend(
            [
                f'        Case "{feature["id"]}"',
                "            definition.ConfigureLaunch "
                + ", ".join(
                    value(feature.get(key))
                    for key in (
                        "launch_surface",
                        "dialog_id",
                        "dialog_template",
                        "dialog_variant",
                        "state_resolver",
                        "source_asset_id",
                        "reuse_mode",
                    )
                ),
            ]
        )
    lines += [
        '        Case Else: NxRaiseContractError "Unknown generated launch feature"',
        "    End Select",
        "End Sub",
        "",
        "Public Function NxGeneratedFeatureLabel(ByVal featureId As String) As String",
        "    Select Case featureId",
    ]
    for feature in contract["features"]:
        lines.extend(
            [
                f'        Case "{feature["id"]}"',
                f'            NxGeneratedFeatureLabel = {value(feature["ribbon"]["label"])}',
            ]
        )
    lines += [
        '        Case Else: NxRaiseContractError "Unknown generated feature label"',
        "    End Select",
        "End Function",
        "",
        "Public Function NxGeneratedFeatureImageMso(ByVal featureId As String) As String",
        "    Select Case featureId",
    ]
    for feature in contract["features"]:
        lines.extend(
            [
                f'        Case "{feature["id"]}"',
                f'            NxGeneratedFeatureImageMso = {value(feature["ribbon"].get("image_mso"))}',
            ]
        )
    lines += [
        '        Case Else: NxRaiseContractError "Unknown generated feature image"',
        "    End Select",
        "End Function",
        "",
    ]
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


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    if args.write == args.check:
        parser.error("choose exactly one of --write or --check")
    catalog = load(FEATURES)
    validate_ui_surfaces(catalog, load(UI_SURFACES))
    contract = release_contract(catalog)
    shell = load(SHELL)
    command_contract = load(COMMANDS)
    validate(contract, shell, command_contract)
    expanded = expand_shell(contract, shell, command_contract)
    expected = {
        XML: render_xml(expanded),
        GENERATED: render_vba(contract, expanded),
        LAUNCH: render_launch(contract),
        LIMITS_POLICY: project_limits_release_ids(
            LIMITS_POLICY.read_bytes().decode("utf-8"),
            [feature["id"] for feature in contract["features"]],
        ),
    }
    if args.check:
        for path, content in expected.items():
            if not path.exists() or path.read_bytes() != content.encode("utf-8"):
                raise SystemExit(f"stale generated output: {path}")
        return 0
    for path, content in expected.items():
        atomic_write(path, content)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
