from __future__ import annotations

import argparse
import hashlib
import json
import re
import uuid
from pathlib import Path


BASE_FILES = (
    "src/vba/core/NxTypes.bas", "src/vba/core/NxConstants.bas", "src/vba/core/NxFormulaInterop.bas", "src/vba/core/NxUserErrors.bas", "src/vba/core/INxFeatureCommand.cls",
    "src/vba/core/CNxExecutionContext.cls", "src/vba/core/CNxPlan.cls", "src/vba/core/CNxResult.cls",
    "src/vba/core/CNxStateGuard.cls", "src/vba/core/CNxSideEffect.cls", "src/vba/core/CNxCellRollbackJournal.cls",
    "src/vba/core/NxContextFactory.bas", "src/vba/core/NxCommandRunner.bas", "src/vba/core/CNxApproval.cls",
    "src/vba/core/CNxSideEffectDispatcher.cls", "src/vba/core/CNxLimitsPolicy.cls",
    "src/vba/core/CNxRangeSelectionSession.cls", "src/vba/core/NxRangePicker.bas",
    "tests/vba/NxTestHarness.bas", "tests/vba/T_Core.bas", "tests/vba/CFakeFeatureCommand.cls",
)


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def root_relative(root: Path, value: str) -> tuple[Path, str]:
    candidate = (root / value).resolve()
    relative = candidate.relative_to(root).as_posix()
    suffix = candidate.suffix.lower()
    is_form_json = relative.lower().endswith(".form.json")
    if suffix not in {".bas", ".cls", ".vba"} and not is_form_json:
        raise ValueError("only BAS, CLS, form JSON, and form VBA inputs are allowed")
    if not candidate.is_file():
        raise ValueError(f"module is missing: {relative}")
    return candidate, relative


def dependency_record(root: Path, raw: str) -> dict[str, str]:
    parts = raw.split("|")
    if len(parts) != 3:
        raise ValueError("dependency must be path|suite|owner")
    path, suite, owner = parts
    if not suite or not re.fullmatch(r"[A-Za-z][A-Za-z0-9_-]*", suite):
        raise ValueError("dependency suite rejected")
    if not owner or not re.fullmatch(r"[a-z][a-z0-9_-]*", owner):
        raise ValueError("dependency owner rejected")
    try:
        full = (root / path).resolve()
        relative = full.relative_to(root).as_posix()
    except ValueError as exc:
        raise ValueError("dependency manifest path rejected") from exc
    if not re.fullmatch(r"tests/manifests/[A-Za-z][A-Za-z0-9_-]*\.json", relative):
        raise ValueError("dependency manifest path rejected")
    if not full.is_file():
        raise ValueError("dependency manifest is missing")
    document = json.loads(full.read_text(encoding="utf-8"))
    if set(document) != {
        "schema_version", "suite", "owner_phase", "base_allowlist_digest", "dependencies", "modules",
    }:
        raise ValueError("dependency manifest property set rejected")
    if document.get("schema_version") != 2 or type(document.get("schema_version")) is not int:
        raise ValueError("dependency schema rejected")
    if document.get("suite") != suite or document.get("owner_phase") != owner:
        raise ValueError("dependency suite or owner rejected")
    return {"path": relative, "sha256": digest(full), "suite": suite, "owner_phase": owner}


def base_files_from_builder(builder: Path) -> tuple[str, ...]:
    source = builder.read_text(encoding="utf-8-sig")
    match = re.search(r"\$greenSources\s*=\s*@\((.*?)\n\)", source, re.S)
    if match is None:
        raise ValueError("greenSources block is missing")
    paths = tuple(re.findall(r"Join-Path \$dataRoot '([^']+)'", match.group(1)))
    if paths != BASE_FILES:
        raise ValueError("base allowlist must remain the exact 21-file contract")
    return paths


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True)
    parser.add_argument("--suite", required=True)
    parser.add_argument("--owner", required=True)
    parser.add_argument("--base-allowlist-from", required=True)
    parser.add_argument("--dependency", action="append", default=[])
    parser.add_argument("--output", required=True)
    parser.add_argument("modules", nargs="+")
    args = parser.parse_args()

    root = Path(args.root).resolve()
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_-]*", args.suite):
        raise ValueError("suite rejected")
    if not re.fullmatch(r"[a-z][a-z0-9_-]*", args.owner):
        raise ValueError("owner rejected")
    base_paths = base_files_from_builder((root / args.base_allowlist_from).resolve())
    base_records = [(path, digest(root / path)) for path in base_paths]
    canonical = "".join(f"{path}\n{sha}\n" for path, sha in base_records)
    base_digest = hashlib.sha256(canonical.encode("utf-8")).hexdigest()

    seen: set[str] = set()
    base_lower = {path.lower() for path in base_paths}
    modules: list[dict[str, str]] = []
    for raw in args.modules:
        full, relative = root_relative(root, raw)
        if full.suffix.lower() == ".vba" or relative.lower().endswith(".form.json"):
            expected = Path("src/vba/frame/ui") if args.owner == "frame" else Path(f"src/vba/features/{args.owner}/ui")
            if Path(relative).parent != expected:
                raise ValueError("UserForm input crosses the exact owner UI path")
        key = relative.lower()
        if key in seen:
            raise ValueError("duplicate extension module")
        if key in base_lower:
            raise ValueError("base shadowing is forbidden")
        if args.owner.lower() not in {part.lower() for part in Path(relative).parts}:
            raise ValueError("extension module crosses owner phase")
        seen.add(key)
        modules.append({"path": relative, "sha256": digest(full), "owner_phase": args.owner})

    dependencies = [dependency_record(root, raw) for raw in args.dependency]
    dependency_keys = [(row["path"].lower(), row["suite"].lower(), row["owner_phase"]) for row in dependencies]
    if len(dependency_keys) != len(set(dependency_keys)):
        raise ValueError("duplicate dependency rejected")

    document = {
        "schema_version": 2,
        "suite": args.suite,
        "owner_phase": args.owner,
        "base_allowlist_digest": base_digest,
        "dependencies": dependencies,
        "modules": modules,
    }
    try:
        output = (root / args.output).resolve()
        output.relative_to(root)
    except ValueError as exc:
        raise ValueError("output path escapes repository root") from exc
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.name}.{uuid.uuid4().hex}.tmp")
    temporary.write_text(
        json.dumps(document, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    temporary.replace(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
