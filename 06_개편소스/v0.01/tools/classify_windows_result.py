#!/usr/bin/env python3
"""Classify a Windows verification lane without changing anything in it."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

VERDICTS = {"PASS", "DIAGNOSTIC", "RUNNING", "NO_RESULT", "CONTRACT_INVALID"}
_UUID = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$", re.I)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _status(path: Path) -> tuple[dict[str, str], list[str], str | None]:
    values: dict[str, str] = {}
    returns: list[str] = []
    try:
        for raw in path.read_text(encoding="utf-8-sig").splitlines():
            if not raw.strip() or "=" not in raw:
                continue
            key, value = raw.split("=", 1)
            key, value = key.strip(), value.strip()
            if key == "return_zip":
                returns.append(value)
            else:
                values[key] = value
    except (OSError, UnicodeError) as exc:
        return {}, [], str(exc)
    return values, returns, None


def _json_member(archive: zipfile.ZipFile, name: str) -> dict[str, Any] | None:
    try:
        value = json.loads(archive.read(name).decode("utf-8-sig"))
    except (KeyError, OSError, UnicodeError, json.JSONDecodeError):
        return None
    return value if isinstance(value, dict) else None


def classify_lane(lane: str | Path) -> dict[str, Any]:
    """Return a stable JSON-serialisable verdict for *lane* (read-only)."""
    root = Path(lane)
    status_path = root / "status" / "run-status.txt"
    return_dir = root / "return"
    result: dict[str, Any] = {
        "verdict": "NO_RESULT",
        "suite": None,
        "run_id": None,
        "exit": None,
        "primary_exit": None,
        "zip": None,
        "return_zip": None,
        "zip_sha256": None,
        "green_compile": {"location": None, "classification": "NO_EVIDENCE"},
        "green_compile_location": None,
        "green_compile_classification": "NO_EVIDENCE",
    }
    if not root.is_dir():
        result["verdict"] = "CONTRACT_INVALID"
        result["error"] = "lane is not a directory"
        return result

    values: dict[str, str] = {}
    returns: list[str] = []
    status_error = None
    if status_path.is_file():
        values, returns, status_error = _status(status_path)
        result["suite"] = values.get("package_label")
        result["run_id"] = values.get("run_id")
        if values.get("launcher_exit", "").lstrip("-").isdigit():
            result["exit"] = int(values["launcher_exit"])
        if values.get("primary_launcher_exit", "").lstrip("-").isdigit():
            result["primary_exit"] = int(values["primary_launcher_exit"])

    pass_zip, diagnostic_zip = return_dir / "PASS.zip", return_dir / "DIAGNOSTIC.zip"
    present = [p for p in (pass_zip, diagnostic_zip) if p.is_file()]
    running = (root / "status" / ".running").exists()
    finished = "wrapper_finished" in values
    errors: list[str] = []
    if status_error:
        errors.append("status unreadable")
    if values.get("run_id") and not _UUID.fullmatch(values["run_id"]):
        errors.append("invalid run_id")
    if len(present) > 1:
        errors.append("both PASS.zip and DIAGNOSTIC.zip exist")
    declared = [name for name in returns if name in {"PASS.zip", "DIAGNOSTIC.zip"}]
    if len(set(declared)) > 1:
        errors.append("status declares conflicting return ZIPs")
    if present:
        archive_name = present[0].name
        if not status_path.is_file():
            errors.append("return ZIP exists without run-status.txt")
        if not finished or not values.get("wrapper_finished"):
            errors.append("return ZIP requires wrapper_finished status")
        if result["exit"] is None:
            errors.append("return ZIP requires integer launcher_exit status")
        if "primary_launcher_exit" in values and result["primary_exit"] is None:
            errors.append("primary_launcher_exit must be an integer when present")
        if (
            result["primary_exit"] is not None
            and result["exit"] != result["primary_exit"]
            and result["exit"] != 25
        ):
            errors.append("launcher_exit may differ from primary only for cleanup exit 25")
        if returns != [archive_name]:
            errors.append("status must declare exactly one matching return_zip")
        if running:
            errors.append("return ZIP cannot coexist with .running sentinel")
        if archive_name == "PASS.zip" and result["exit"] != 0:
            errors.append("PASS.zip requires launcher_exit=0")
        if archive_name == "DIAGNOSTIC.zip" and (result["exit"] is None or result["exit"] == 0):
            errors.append("DIAGNOSTIC.zip launcher_exit must be nonzero")

    if errors:
        result["verdict"] = "CONTRACT_INVALID"
        result["errors"] = errors
        return result
    if not present:
        result["verdict"] = "RUNNING" if running or (status_path.is_file() and not finished) else "NO_RESULT"
        return result
    archive_path = present[0]
    result["zip"] = {"name": archive_path.name, "sha256": _sha256(archive_path)}
    result["return_zip"] = archive_path.name
    result["zip_sha256"] = result["zip"]["sha256"]
    # ProductRibbon returns have a manifest/evidence contract rather than the
    # generic results/*.json contract used by the other suites.
    if values.get("package_label") == "ProductRibbon":
        try:
            from inspect_windows_return import inspect_return
        except ImportError:  # pragma: no cover - direct package execution
            from .inspect_windows_return import inspect_return
        inspected = inspect_return(return_dir, "ProductRibbon")
        result["suite"] = "ProductRibbon"
        if result["run_id"] and inspected.get("run_id") and result["run_id"] != inspected["run_id"]:
            result["verdict"] = "CONTRACT_INVALID"
            result["errors"] = ["status and ProductRibbon manifest run_id disagree"]
            return result
        result["run_id"] = result["run_id"] or inspected.get("run_id")
        expected_diagnostic_exit = (
            result["primary_exit"] if result["primary_exit"] is not None else result["exit"]
        )
        if archive_path.name == "DIAGNOSTIC.zip" and inspected.get("exit_code") != expected_diagnostic_exit:
            result["verdict"] = "CONTRACT_INVALID"
            result["errors"] = ["status and ProductRibbon diagnostic exit_code disagree"]
            return result
        if inspected.get("verdict") == "PASS":
            result["verdict"] = "PASS"
        elif archive_path.name == "DIAGNOSTIC.zip" and inspected.get("diagnostic_valid") is True:
            result["verdict"] = "DIAGNOSTIC"
        else:
            result["verdict"] = "CONTRACT_INVALID"
            result["errors"] = inspected.get("errors", [])
        return result
    try:
        with zipfile.ZipFile(archive_path) as archive:
            names = set(archive.namelist())
            diagnostic = _json_member(archive, "diagnostic.json")
            result_names = sorted(n for n in names if n.startswith("results/") and n.count("/") == 1 and n.endswith(".json"))
            preferred = f"results/{values.get('package_label')}.json" if values.get("package_label") else ""
            if preferred in result_names:
                result_names.remove(preferred)
                result_names.insert(0, preferred)
            result_doc = next((_json_member(archive, n) for n in result_names), None)
            if diagnostic is not None:
                diagnostic_run_id = diagnostic.get("run_id")
                if result["run_id"] and diagnostic_run_id and result["run_id"] != diagnostic_run_id:
                    errors.append("status and diagnostic run_id disagree")
                result["run_id"] = result["run_id"] or diagnostic_run_id
                diagnostic_exit = diagnostic.get("exit_code")
                if not isinstance(diagnostic_exit, int) or isinstance(diagnostic_exit, bool) or diagnostic_exit == 0:
                    errors.append("diagnostic exit_code must be a nonzero integer")
                elif (result["primary_exit"] if result["primary_exit"] is not None else result["exit"]) != diagnostic_exit:
                    errors.append("status and diagnostic exit_code disagree")
            if result_doc:
                result_run_id = result_doc.get("run_id")
                if result["run_id"] and result_run_id and result["run_id"] != result_run_id:
                    errors.append("status and result run_id disagree")
                result["suite"] = result["suite"] or result_doc.get("suite")
                compile_doc = result_doc.get("compile")
                if isinstance(compile_doc, dict):
                    location = {k: compile_doc[k] for k in ("active_module", "line", "caret_column", "source_line_sha256") if k in compile_doc}
                    if location:
                        classification = "COMPILE_DIALOG" if compile_doc.get("dialog_text") else "COMPILE_OK"
                        result["green_compile"] = {"location": location, "classification": classification}
                        result["green_compile_location"] = location
                        result["green_compile_classification"] = classification
            if not result["run_id"] or not _UUID.fullmatch(str(result["run_id"])):
                errors.append("missing or invalid run_id in result")
            if result["suite"] is None:
                errors.append("missing suite in result")
            if archive_path.name == "PASS.zip":
                if diagnostic is not None or not result_doc:
                    errors.append("PASS.zip has invalid result contract")
                if result_doc:
                    if result_doc.get("status") != "PASS" or result_doc.get("classification") != "PASS":
                        errors.append("PASS.zip suite result is not PASS/PASS")
                    run = result_doc.get("run")
                    if not isinstance(run, dict):
                        errors.append("PASS.zip suite result lacks run summary")
                    else:
                        failed, total = run.get("failed"), run.get("total")
                        passed, skipped = run.get("passed"), run.get("skipped", 0)
                        if failed != 0 or not all(isinstance(v, int) and not isinstance(v, bool) for v in (total, passed, skipped)) or total != passed + skipped:
                            errors.append("PASS.zip run summary is not zero-failure and complete")
                if result["exit"] != 0:
                    errors.append("PASS.zip launcher_exit must be zero")
            elif diagnostic is None:
                errors.append("DIAGNOSTIC.zip lacks diagnostic.json")
    except (OSError, zipfile.BadZipFile) as exc:
        errors.append(f"invalid ZIP: {exc}")
    if errors:
        result["verdict"] = "CONTRACT_INVALID"
        result["errors"] = errors
    else:
        result["verdict"] = "PASS" if archive_path.name == "PASS.zip" else "DIAGNOSTIC"
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("lane", type=Path)
    args = parser.parse_args(argv)
    print(json.dumps(classify_lane(args.lane), ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
