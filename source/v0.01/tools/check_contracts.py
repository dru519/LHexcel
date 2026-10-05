"""Fail-closed checks for the v0.01 clean-room source boundary."""

from __future__ import annotations

import argparse
import fnmatch
import hashlib
import importlib.util
import json
import re
import sys
from pathlib import Path
from typing import Any


REQUIRED_RULE_FIELDS = ("pattern", "origin", "spec_ids", "license")
APPROVED_NON_PROJECT_ORIGINS = {
    "provenance/approved-lineage-policy.json": "user-approved-v3x-lineage",
}
# 도구가 만드는 부산물. 프로젝트 자체 테스트를 한 번 돌리는 것만으로 게이트가
# 실패하지 않도록 .git 과 같은 범주로 열거에서 제외한다.
IGNORED_NAMES = {"__pycache__", ".git", ".pytest_cache"}
IGNORED_SUFFIXES = {".pyc"}
PROVENANCE_EXCLUDED_ROOTS = {"out", "tmp", "__pycache__"}

# 파일 관리자(Finder, 탐색기)가 만드는 부산물. 제품 payload가 아니고 Git 추적 대상도
# 아니므로 .git 과 같은 범주로 clean-room 열거에서 제외한다. 이것을 소스로 세면
# 폴더를 한 번 열어보는 것만으로 게이트가 실패한다. 대소문자 변형까지 막는다.
OS_ARTIFACT_NAMES = {".ds_store", "thumbs.db", "desktop.ini"}
OS_ARTIFACT_PREFIXES = ("._",)  # macOS AppleDouble sidecar
R46_CONTRACT_NAMES = (
    "bridge-contract.json",
    "build-profile-contract.json",
    "command-contract.json",
)
V34_COMMAND_PROVENANCE = "provenance/v34-command-catalog.json"
R62_COMMAND_REMOVAL = "tests/fixtures/r62-basic-command-removal.json"
# Original r68 commands have no donor IDs; keep their approved identity closed.
R68_NATIVE_COMMANDS = {
    "NX-CMD-STYLE-" + suffix: "NX_STYLE_" + suffix.replace("-", "_")
    for suffix in (
        "FONT-SIZE-9", "FONT-SIZE-10", "FONT-SIZE-12", "FONT-SIZE-15",
        "FONT-COLOR-BLACK", "FONT-COLOR-RED", "FONT-COLOR-BLUE", "FONT-COLOR-GREEN",
        "FILL-GRAY", "FILL-LIGHT-RED", "FILL-LIGHT-YELLOW", "FILL-LIGHT-GREEN",
    )
}
NATIVE_COMMAND_IDS = set(R68_NATIVE_COMMANDS) | {"NX-CMD-DATA-DATE-CONVERT"}
R74_NATIVE_COMMANDS = {"NX-CMD-NUMBER-EMPHASIS": ("NxCmdNumberFormat", "NX_NUMBER_EMPHASIS")}
R74_NATIVE_COMMANDS.update({"NX-CMD-VIEW-PRESET-" + str(i): ("NxCmdView", "NX_VIEW_PRESET_" + str(i)) for i in range(1, 6)})
R74_NATIVE_COMMANDS["NX-CMD-VIEW-PRESET-SETTINGS"] = ("NxCmdView", "NX_VIEW_PRESET_SETTINGS")
NATIVE_COMMAND_IDS |= R74_NATIVE_COMMANDS.keys()
R94_NATIVE_COMMANDS = {
    "NX-CMD-FUNCTION-ROUND": ("NxCmdFormula", "NX_FUNCTION_ROUND"),
    "NX-CMD-FUNCTION-IFERROR": ("NxCmdFormula", "NX_FUNCTION_IFERROR"),
}
NATIVE_COMMAND_IDS |= R94_NATIVE_COMMANDS.keys()
EXPECTED_BRIDGE_OPERATIONS = {
    "PING",
    "SHOW_NAVIGATOR",
    "SHOW_DOCUMENT_NAVIGATOR",
    "HIDE_DOCUMENT_NAVIGATOR",
    "HIDE_NAVIGATOR",
    "FOCUS_START",
    "FOCUS_UPDATE",
    "FOCUS_STOP",
}
EXPECTED_FORBIDDEN_BRIDGE_OPERATIONS = {
    "APPLICATION_RUN",
    "SHELL_EXECUTE",
    "NETWORK_REQUEST",
}
EXPECTED_BUILD_PROFILE_IDS = {"enhanced-dll", "internal-xlam"}
EXPECTED_COMMAND_FIELDS = [
    "id", "label_ko", "aliases_ko", "description_ko", "image_mso", "search_terms_ko",
    "category_id", "group_id", "group_order", "sort_order", "input_context",
    "active_when", "mutates_document", "recovery_policy", "ribbon_visible",
    "taskpane_visible", "favorite_eligible", "shortcut_eligible", "handler",
    "args", "legacy_ids", "legacy_entrypoints", "navigation",
]
EXPECTED_COMMAND_HANDLERS = [
    "NxCmdClipboard", "NxCmdPrint", "NxCmdStyle", "NxCmdNumberFormat",
    "NxCmdSelection", "NxCmdRowsColumns", "NxCmdAlignment", "NxCmdInsertDelete",
    "NxCmdFilterSort", "NxCmdFormula", "NxCmdSheet", "NxCmdView", "NxCmdInfo", "NxCmdPivot",
]
COMMAND_ID = re.compile(r"^NX-CMD-[A-Z0-9-]+$")
OPERATION_KEY = re.compile(r"^[A-Z0-9_]+$")


def _load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"cannot read JSON: {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"JSON object required: {path}")
    return value


def _is_os_artifact(path: Path) -> bool:
    for part in path.parts:
        lowered = part.casefold()
        if lowered in OS_ARTIFACT_NAMES or lowered.startswith(OS_ARTIFACT_PREFIXES):
            return True
    return False


def _source_files(root: Path) -> list[Path]:
    files: list[Path] = []
    for path in root.rglob("*"):
        if (
            not path.is_file()
            or any(part in IGNORED_NAMES for part in path.parts)
            or path.suffix in IGNORED_SUFFIXES
            or _is_os_artifact(path)
        ):
            continue
        relative_path = path.relative_to(root)
        if any(part.casefold() in PROVENANCE_EXCLUDED_ROOTS for part in relative_path.parts):
            continue
        relative = relative_path.as_posix()
        # origins.json declares the rules that cover the tree. It cannot
        # cover itself without a recursive policy dependency.
        if relative not in {"provenance/origins.json", V34_COMMAND_PROVENANCE}:
            files.append(path)
    return sorted(files, key=lambda item: item.relative_to(root).as_posix())


def require_exact_ids(rows: list[dict[str, Any]], expected: set[str], label: str) -> None:
    actual = {row.get("id") for row in rows}
    if actual != expected:
        raise ValueError(f"{label} ids mismatch: {sorted(str(item) for item in actual)}")


def _validate_command_execution(root: Path, command_rows: list[dict[str, Any]]) -> list[str]:
    errors: list[str] = []
    ui_root = root / "src" / "vba" / "ui"
    generated_path = ui_root / "NxGeneratedCommandRegistry.bas"
    registry_path = ui_root / "NxCommandRegistry.bas"
    handlers_path = ui_root / "NxCommandHandlers.bas"
    required_paths = (generated_path, registry_path, handlers_path)
    missing = [path.name for path in required_paths if not path.is_file()]
    if missing:
        return [f"command execution source missing: {', '.join(missing)}"]

    generated = generated_path.read_text(encoding="utf-8-sig")
    # Compare the complete deterministic output, not only regex-recognized rows:
    # extra malformed Add statements, sealed flags and count functions must fail too.
    try:
        spec = importlib.util.spec_from_file_location(
            "_nx_current_command_renderer", root / "tools" / "generate_command_surface.py")
        if spec is None or spec.loader is None:
            raise ValueError("command renderer unavailable")
        renderer = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(renderer)
        expected_generated = renderer.render_vba(command_rows)
    except (OSError, ImportError, AttributeError, KeyError, IndexError, TypeError, ValueError) as exc:
        errors.append(f"generated command registry renderer rejected: {exc}")
    else:
        if generated != expected_generated:
            errors.append("generated command registry does not exactly bind current catalog content")
    expected = {
        str(row["id"]): (str(row["group_id"]), str(row["handler"]), str(row["args"][0]))
        for row in command_rows
        if isinstance(row.get("args"), list) and len(row["args"]) == 1
    }
    rows = re.findall(
        r'^\s*registry\.Add "(NX-CMD-[A-Z0-9-]+)", "(?:""|[^"])*", "([A-Z0-9-]+)", "(NX-CMD-GRP-[A-Z0-9-]+)", "(NxCmd[A-Za-z0-9]+)", "([A-Z0-9_]+)", (?:True|False), (?:True|False), "[a-z_]+"$',
        generated,
        re.M,
    )
    actual = {command_id: (group_id, handler, operation) for command_id, _, group_id, handler, operation in rows}
    if len(rows) != len(command_rows) or len(actual) != len(rows) or actual != expected:
        errors.append("generated command registry does not exactly bind current catalog operations")

    expected_handlers = {row["handler"] for row in command_rows}
    handlers_source = handlers_path.read_text(encoding="utf-8-sig")
    implemented = re.findall(r"^Public Sub (NxCmd[A-Za-z0-9]+)\b", handlers_source, re.M)
    dispatches = re.findall(
        r'^\s*Case "(NxCmd[A-Za-z0-9]+)": (NxCmd[A-Za-z0-9]+) definition\.ArgumentKey$',
        registry_path.read_text(encoding="utf-8-sig"), re.M,
    )
    if (len(implemented) != len(expected_handlers) or set(implemented) != expected_handlers
            or len(dispatches) != len(expected_handlers)
            or set(dispatches) != {(handler, handler) for handler in expected_handlers}):
        errors.append("command dispatch handler set differs from current catalog")

    sources = "\n".join(path.read_text(encoding="utf-8-sig") for path in required_paths)
    for forbidden in ("Application.Run", "CallByName", "ExecuteMso", "NxBuiltInCommandExecute", "placeholder", "not implemented"):
        if forbidden.casefold() in sources.casefold():
            errors.append(f"command execution contains forbidden dispatch token: {forbidden}")

    for handler in sorted(expected_handlers):
        blocks = list(re.finditer(rf"^Public Sub {re.escape(handler)}\b.*?^End Sub$", sources, re.M | re.S))
        if len(blocks) != 1:
            errors.append(f"command handler implementation count rejected: {handler}")
            continue
        accepted = [
            operation
            for case in re.findall(r"^\s*Case\s+(.+?)(?::|$)", blocks[0].group(0), re.M)
            for operation in re.findall(r'"([A-Z0-9_]+)"', case)
        ]
        expected_operations = {operation for _, mapped_handler, operation in expected.values() if mapped_handler == handler}
        if set(accepted) != expected_operations or len(accepted) != len(set(accepted)):
            errors.append(f"command handler operation coverage rejected: {handler}")
    return errors


def _historical_v34_catalog(root: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Read historical identity only; never invoke importer outputs/main or write files."""
    importer = root / "tools" / "import_v34_command_catalog.py"
    donor = root.parent / "v3.4" / "candidate" / "src" / "mdShortcutCatalog.bas"
    if not importer.is_file() or not donor.is_file():
        raise ValueError("historical command lineage reader or donor missing")
    try:
        spec = importlib.util.spec_from_file_location("_nx_historical_v34_catalog", importer)
        if spec is None or spec.loader is None:
            raise ValueError("historical command lineage reader rejected")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        raw_rows = module.parse_donor(donor.read_text(encoding="utf-8-sig"))
        return raw_rows, module.normalize_rows(raw_rows)
    except (OSError, ValueError, ImportError, AttributeError) as exc:
        raise ValueError(f"historical command lineage rejected: {exc}") from exc


def _validate_current_command_lineage(root: Path, commands: list[dict[str, Any]]) -> list[str]:
    try:
        removal = _load_json(root / R62_COMMAND_REMOVAL)
        _, historical_rows = _historical_v34_catalog(root)
    except ValueError as exc:
        return [str(exc)]
    historical = {row["id"]: row for row in historical_rows}
    removed = removal.get("removed_commands")
    preserved = removal.get("preserved_number_format_ids")
    if (not isinstance(removed, list) or len(removed) != 44
            or not all(isinstance(row, dict) and isinstance(row.get("id"), str) for row in removed)
            or not isinstance(preserved, list) or len(preserved) != 28
            or not all(isinstance(item, str) for item in preserved)):
        return ["approved removal fixture structure rejected"]
    removed_ids = {row["id"] for row in removed}
    if (len(removed_ids) != len(removed) or not removed_ids <= historical.keys()
            or len(set(preserved)) != len(preserved)):
        return ["approved removal fixture identities rejected"]
    for row in removed:
        donor_row = historical[row["id"]]
        if donor_row["args"] != [row.get("arg")] or donor_row["handler"] != row.get("handler"):
            return ["approved removal fixture lineage binding rejected"]
    extra = _load_json(root / "tests/fixtures/r74-command-removal.json")["removed_commands"]
    if len(extra) != 36 or len({row["id"] for row in extra}) != 36:
        return ["r74 removal inventory rejected"]
    for row in extra:
        donor = historical.get(row["id"])
        if donor is None or donor["args"] != [row.get("arg")] or donor["handler"] != row.get("handler"):
            return ["r74 removal lineage binding rejected"]
    removed_ids |= {row["id"] for row in extra}
    errors = []
    current = {row["id"]: row for row in commands}
    if current.keys() != (historical.keys() - removed_ids) | NATIVE_COMMAND_IDS or len(current) != len(commands):
        errors.append("current command ids differ from historical catalog minus approved removal plus approved native commands")
    expected_number = {row["id"] for row in historical_rows if row["handler"] == "NxCmdNumberFormat"}
    current_number = {row["id"] for row in commands if row["handler"] == "NxCmdNumberFormat"}
    if set(preserved) != expected_number or current_number != (expected_number - removed_ids) | {"NX-CMD-NUMBER-EMPHASIS"}:
        errors.append("preserved number-format command set differs from historical catalog")
    for command_id, (handler, arg) in R74_NATIVE_COMMANDS.items():
        row = current.get(command_id, {})
        if row.get("handler") != handler or row.get("args") != [arg]:
            errors.append("r74 native command binding rejected: " + command_id)
    for command_id, (handler, arg) in R94_NATIVE_COMMANDS.items():
        row = current.get(command_id, {})
        if (row.get("handler") != handler or row.get("args") != [arg]
                or row.get("legacy_ids") != [] or row.get("legacy_entrypoints") != []
                or row.get("group_id") != "NX-CMD-GRP-FORMULA-REFERENCE"
                or row.get("mutates_document") is not True
                or row.get("recovery_policy") != "command_scoped_guard"
                or row.get("input_context") != "range_selection"):
            errors.append("r94 native command binding rejected: " + command_id)
    for command_id in current.keys() & historical.keys():
        for field in ("handler", "args", "legacy_ids", "legacy_entrypoints"):
            if current[command_id].get(field) != historical[command_id][field]:
                errors.append(f"command lineage binding changed: {command_id}.{field}")
    for command_id in current.keys() & R68_NATIVE_COMMANDS.keys():
        expected = {
            "handler": "NxCmdStyle", "args": [R68_NATIVE_COMMANDS[command_id]],
            "legacy_ids": [], "legacy_entrypoints": [], "group_id": "NX-CMD-GRP-STYLE",
            "mutates_document": True, "recovery_policy": "command_scoped_guard",
            "input_context": "range_selection",
        }
        for field, value in expected.items():
            if current[command_id].get(field) != value:
                errors.append(f"native command binding changed: {command_id}.{field}")
    date_command = current.get("NX-CMD-DATA-DATE-CONVERT", {})
    for field, value in {
        "handler": "NxCmdSelection", "args": ["NX_DATA_DATE_CONVERT"],
        "legacy_ids": [], "legacy_entrypoints": [], "group_id": "NX-CMD-GRP-SELECTION-MOVE",
        "mutates_document": False, "recovery_policy": "none", "input_context": "range_selection",
    }.items():
        if date_command.get(field) != value:
            errors.append(f"date conversion launch binding changed: {field}")
    return errors


def _validate_v34_command_provenance(root: Path) -> list[str]:
    # The clean-room checker is also exercised against minimal policy fixtures.
    # Require the v3.4 catalog provenance only when the r46 command surface is
    # actually present; production still fails closed because it ships that
    # contract.
    if not (root / "contracts" / "command-contract.json").is_file():
        return []
    path = root / V34_COMMAND_PROVENANCE
    donor = root.parent / "v3.4" / "candidate" / "src" / "mdShortcutCatalog.bas"
    if not path.is_file():
        return ["v3.4 command provenance missing"]
    if not donor.is_file():
        return ["v3.4 command provenance donor missing"]
    try:
        provenance = _load_json(path)
        raw_rows, historical_rows = _historical_v34_catalog(root)
    except ValueError as exc:
        return [str(exc)]
    source = provenance.get("source")
    if not isinstance(source, dict):
        return ["v3.4 command provenance source rejected"]
    if (
        provenance.get("schema_version") != 1
        or source.get("path") != "../v3.4/candidate/src/mdShortcutCatalog.bas"
        or source.get("sha256") != hashlib.sha256(donor.read_bytes()).hexdigest()
        or source.get("reuse_basis") != "user-owned-v3.4-lineage"
        # These are historical donor counts, deliberately not active contract counts.
        or provenance.get("selected_raw_count") != len(raw_rows)
        or provenance.get("normalized_command_count") != len(historical_rows)
        or provenance.get("duplicate_groups_merged") != len(raw_rows) - len(historical_rows)
        or provenance.get("third_party_assets") != []
    ):
        return ["v3.4 command provenance integrity mismatch"]
    return []


def _validate_r46_contracts(root: Path) -> list[str]:
    errors: list[str] = []
    paths = [root / "contracts" / name for name in R46_CONTRACT_NAMES]
    if not any(path.exists() for path in paths):
        return errors
    missing = [path.name for path in paths if not path.is_file()]
    if missing:
        return [f"r46 contract missing: {', '.join(missing)}"]
    try:
        bridge = _load_json(root / "contracts" / "bridge-contract.json")
        profiles = _load_json(root / "contracts" / "build-profile-contract.json")
        commands = _load_json(root / "contracts" / "command-contract.json")
    except ValueError as exc:
        return [str(exc)]

    if bridge.get("schema_version") != 3 or bridge.get("interface_version") != "nx-bridge-v1":
        errors.append("bridge contract version mismatch")
    if bridge.get("product_title") != "내엑셀 v0.01_r86" or bridge.get("excel_bitness") != ["x86", "x64"]:
        errors.append("bridge product identity mismatch")
    handshake = bridge.get("handshake")
    expected_handshake_fields = [
        "product_title",
        "interface_version",
        "bridge_sha256",
        "feature_sha256",
        "command_sha256",
        "excel_bitness",
    ]
    if (
        not isinstance(handshake, dict)
        or handshake.get("required_fields") != expected_handshake_fields
        or handshake.get("required_hashes") != ["bridge_sha256", "feature_sha256", "command_sha256"]
        or handshake.get("identity_rule") != "exact"
        or handshake.get("product_manifest_hash") is not False
    ):
        errors.append("bridge handshake contract mismatch")
    allowed = bridge.get("allowed_operations")
    forbidden = bridge.get("forbidden_operations")
    if not isinstance(allowed, list) or set(allowed) != EXPECTED_BRIDGE_OPERATIONS:
        errors.append("bridge allowed operations mismatch")
    if not isinstance(forbidden, list) or set(forbidden) != EXPECTED_FORBIDDEN_BRIDGE_OPERATIONS:
        errors.append("bridge forbidden operations mismatch")
    if isinstance(allowed, list) and isinstance(forbidden, list) and set(allowed) & set(forbidden):
        errors.append("bridge operations cannot be both allowed and forbidden")

    profile_rows = profiles.get("profiles")
    if profiles.get("schema_version") != 1 or not isinstance(profile_rows, list):
        errors.append("build profile contract structure mismatch")
        profile_rows = []
    typed_profiles = [row for row in profile_rows if isinstance(row, dict)]
    if len(typed_profiles) != len(profile_rows):
        errors.append("build profiles must be objects")
    try:
        require_exact_ids(typed_profiles, EXPECTED_BUILD_PROFILE_IDS, "build profile")
    except ValueError as exc:
        errors.append(str(exc))
    profile_by_id = {row.get("id"): row for row in typed_profiles}
    internal = profile_by_id.get("internal-xlam", {})
    enhanced = profile_by_id.get("enhanced-dll", {})
    if profiles.get("default_profile") != "enhanced-dll" or profiles.get("fallback_profile") != "internal-xlam":
        errors.append("build profile default/fallback mismatch")
    if internal.get("allowed_sidecars") != [] or internal.get("bridge_enabled") is not False:
        errors.append("internal-xlam must be a sidecar-free bridge-disabled profile")
    if enhanced.get("allowed_sidecars") != ["NxHost32.dll", "NxHost64.dll", "NxCore32.dll", "NxCore64.dll"] or enhanced.get("bridge_enabled") is not True:
        errors.append("enhanced-dll sidecar contract mismatch")
    required_profile_fields = profiles.get("required_fields")
    if not isinstance(required_profile_fields, list):
        errors.append("build profile required_fields must be a list")
        required_profile_fields = []
    for profile_id, row in profile_by_id.items():
        for field in required_profile_fields:
            if field not in row:
                errors.append(f"build profile {profile_id} missing {field}")

    command_rows = commands.get("commands")
    required_fields = commands.get("required_fields")
    allowed_handlers = commands.get("allowed_handlers")
    if allowed_handlers != EXPECTED_COMMAND_HANDLERS:
        errors.append("command handler allowlist mismatch")
        allowed_handlers = []
    if (
        commands.get("schema_version") != 2
        or required_fields != EXPECTED_COMMAND_FIELDS
        or not isinstance(commands.get("groups"), list)
        or not commands["groups"]
        or not isinstance(command_rows, list)
        or not command_rows
    ):
        errors.append("command contract structure mismatch")
        command_rows = []
    group_rows = commands.get("groups") if isinstance(commands.get("groups"), list) else []
    group_by_id = {row.get("id"): row for row in group_rows if isinstance(row, dict)}
    if (len(group_by_id) != len(group_rows)
            or set(group_by_id) != {row.get("group_id") for row in command_rows if isinstance(row, dict)}
            or {row.get("handler") for row in group_by_id.values()} != set(allowed_handlers)):
        errors.append("command group handler mapping mismatch")
    if type(commands.get("normalized_command_count")) is not int or commands["normalized_command_count"] != len(command_rows):
        errors.append("command normalized count differs from current rows")
    ids: list[str] = []
    legacy_ids: list[str] = []
    operations_by_handler: dict[str, set[str]] = {handler: set() for handler in allowed_handlers}
    for index, row in enumerate(command_rows):
        if not isinstance(row, dict):
            errors.append(f"command {index} must be an object")
            continue
        missing_fields = [field for field in EXPECTED_COMMAND_FIELDS if field not in row]
        if missing_fields:
            errors.append(f"command {index} missing: {', '.join(missing_fields)}")
        row_legacy = row.get("legacy_ids")
        if (not isinstance(row_legacy, list)
                or (not row_legacy and row.get("id") not in NATIVE_COMMAND_IDS)
                or not all(isinstance(item, str) and item for item in row_legacy)):
            errors.append(f"command {index} legacy ids rejected")
        else:
            legacy_ids.extend(row_legacy)
        command_id = row.get("id")
        if not isinstance(command_id, str) or COMMAND_ID.fullmatch(command_id) is None:
            errors.append(f"command {index} id rejected")
        else:
            ids.append(command_id)
        if not isinstance(row.get("mutates_document"), bool):
            errors.append(f"command {index} mutates_document must be boolean")
        image_mso = row.get("image_mso")
        if not isinstance(image_mso, str) or not re.fullmatch(r"[A-Za-z][A-Za-z0-9]+", image_mso):
            errors.append(f"command {index} image_mso rejected")
        elif image_mso in {"MacroPlay", "ViewMacros"}:
            errors.append(f"command {index} generic image_mso rejected")
        if not isinstance(row.get("shortcut_eligible"), bool):
            errors.append(f"command {index} shortcut_eligible must be boolean")
        if row.get("handler") not in allowed_handlers:
            errors.append(f"command {index} handler rejected")
            continue
        group = group_by_id.get(row.get("group_id"))
        # r59 presentation reclassification does not change the typed executor.
        regrouped = {
            "NX-CMD-STYLE-RB-EDIT-MEMO-ADD-LHEXCELFORMULA": ("NX-CMD-GRP-FORMULA-REFERENCE", "NxCmdStyle"),
        }
        moved_executor = regrouped.get(command_id) == (row.get("group_id"), row["handler"])
        if not isinstance(group, dict) or (group.get("handler") != row["handler"] and not moved_executor):
            errors.append(f"command {index} group handler rejected")
        args = row.get("args")
        if not isinstance(args, list) or len(args) != 1 or not isinstance(args[0], str) or OPERATION_KEY.fullmatch(args[0]) is None:
            errors.append(f"command {index} operation rejected")
        else:
            operation = args[0]
            if operation in operations_by_handler[row["handler"]]:
                errors.append(f"command {index} operation is duplicated for handler")
            operations_by_handler[row["handler"]].add(operation)
    if len(ids) != len(set(ids)):
        errors.append("command ids must be unique")
    native_count = sum(isinstance(row, dict) and row.get("id") in NATIVE_COMMAND_IDS for row in command_rows)
    if (len(legacy_ids) != len(set(legacy_ids)) or type(commands.get("raw_command_count")) is not int
            or commands["raw_command_count"] != len(legacy_ids) + native_count):
        errors.append("command raw count differs from unique current legacy ids plus approved native commands")
    for group_id, group in group_by_id.items():
        group_raw_count = sum(len(row.get("legacy_ids", [])) + (row.get("id") in NATIVE_COMMAND_IDS) for row in command_rows
                              if isinstance(row, dict) and row.get("group_id") == group_id
                              and isinstance(row.get("legacy_ids"), list))
        if type(group.get("raw_count")) is not int or group["raw_count"] != group_raw_count:
            errors.append(f"command group raw count differs from current rows: {group_id}")
    source_path = root / "src" / "resources" / "commands.ko-KR.json"
    if not source_path.is_file():
        errors.append("command source missing")
    else:
        try:
            if _load_json(source_path) != commands:
                errors.append("command source and contract differ")
        except ValueError as exc:
            errors.append(str(exc))
    typed_rows = [row for row in command_rows if isinstance(row, dict)
                  and all(field in row for field in ("id", "group_id", "handler", "args"))]
    if len(typed_rows) != len(command_rows):
        errors.append("command execution binding rows are incomplete")
    errors.extend(_validate_current_command_lineage(root, typed_rows))
    errors.extend(_validate_command_execution(root, typed_rows))
    return errors


def validate(root: Path) -> list[str]:
    errors: list[str] = []
    origins_path = root / "provenance" / "origins.json"
    forbidden_path = root / "provenance" / "forbidden-inputs.json"
    allowlist_path = root / "provenance" / "independence-allowlist.json"
    required_metadata = (origins_path, forbidden_path, allowlist_path)
    if any(not path.is_file() for path in required_metadata):
        missing = [str(path.relative_to(root)) for path in required_metadata if not path.is_file()]
        return [f"provenance missing: {', '.join(missing)}"]

    try:
        origins = _load_json(origins_path)
        forbidden = _load_json(forbidden_path)
        allowlist = _load_json(allowlist_path)
    except ValueError as exc:
        return [str(exc)]

    rules = origins.get("rules")
    if not isinstance(rules, list) or not rules:
        errors.append("provenance rules must be a non-empty list")
        rules = []
    normalized_rules: list[dict[str, Any]] = []
    for index, rule in enumerate(rules):
        if not isinstance(rule, dict):
            errors.append(f"provenance rule {index} must be an object")
            continue
        missing = [field for field in REQUIRED_RULE_FIELDS if not rule.get(field)]
        if missing:
            errors.append(f"provenance rule {index} missing: {', '.join(missing)}")
        pattern = rule.get("pattern")
        expected_origin = APPROVED_NON_PROJECT_ORIGINS.get(pattern, "project-original")
        if rule.get("origin") != expected_origin:
            errors.append(f"provenance rule {index} origin must be {expected_origin}")
        if rule.get("license") != "Proprietary":
            errors.append(f"provenance rule {index} license must be Proprietary")
        if not isinstance(rule.get("spec_ids"), list) or not all(isinstance(item, str) and item for item in rule.get("spec_ids", [])):
            errors.append(f"provenance rule {index} spec_ids must be a non-empty string list")
        normalized_rules.append(rule)

    configured_exclusions = origins.get("excluded_paths", [])
    if configured_exclusions not in (None, []):
        errors.append("provenance exclusions are not allowed; only origins.json is intrinsically excluded")
    files = _source_files(root)
    path_tokens = forbidden.get("path_tokens", [])
    if not isinstance(path_tokens, list) or not path_tokens or not all(isinstance(item, str) and item for item in path_tokens):
        errors.append("forbidden path_tokens must be a non-empty string list")
        path_tokens = []
    allowlist_entries = allowlist.get("entries", [])
    if not isinstance(allowlist_entries, list) or not allowlist_entries:
        errors.append("independence allowlist entries must be a non-empty list")
    else:
        for index, entry in enumerate(allowlist_entries):
            if not isinstance(entry, dict):
                errors.append(f"independence allowlist entry {index} must be an object")
                continue
            required = ("axis", "kind", "value", "max_count", "rationale")
            missing = [field for field in required if entry.get(field) in (None, "")]
            if missing:
                errors.append(f"independence allowlist entry {index} missing: {', '.join(missing)}")
            if entry.get("axis") not in {"code", "workbook", "ui_wording", "assets"}:
                errors.append(f"independence allowlist entry {index} has invalid axis")
            if not isinstance(entry.get("max_count"), int) or entry.get("max_count", 0) < 1:
                errors.append(f"independence allowlist entry {index} max_count must be a positive integer")
    for path in files:
        relative = path.relative_to(root).as_posix()
        lowered = relative.casefold()
        for token in path_tokens:
            if token.casefold() in lowered:
                errors.append(f"forbidden input path token {token!r}: {relative}")
        matches = [rule.get("pattern") for rule in normalized_rules if isinstance(rule.get("pattern"), str) and fnmatch.fnmatchcase(relative, rule["pattern"])]
        if len(matches) != 1:
            errors.append(f"{relative} must match exactly one provenance glob; matched {len(matches)}: {matches}")
    errors.extend(_validate_v34_command_provenance(root))
    errors.extend(_validate_r46_contracts(root))
    return errors


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    if not root.is_dir():
        print(f"root does not exist: {root}", file=sys.stderr)
        return 2
    errors = validate(root)
    if errors:
        print(f"FAIL clean-room contracts ({len(errors)} error(s))")
        for error in errors:
            print(f"- {error}")
        return 1
    print("PASS clean-room contracts: provenance and forbidden-input boundaries are valid")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
