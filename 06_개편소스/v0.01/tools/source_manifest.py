"""Build and freeze a provenance manifest without reading reference inputs.

The command is intentionally fail-closed: a source file must be tracked by
Git, covered by exactly one provenance rule, and have all required metadata.
The live v0.01 tree is not frozen by this module until its caller has made a
clean source commit; tests use a temporary synthetic Git repository instead.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


class FreezeError(RuntimeError):
    """Raised when a source cannot be safely represented by a freeze."""


REQUIRED_ROW_FIELDS = ("path", "origin", "spec_ids", "sha256", "created_commit", "license")
DEFAULT_SOURCE_DIRS = ("src", "contracts", "build", "tools", "provenance")
IGNORED_NAMES = {"__pycache__"}
IGNORED_SUFFIXES = {".pyc"}
SELF_PATHS = {"provenance/origins.json"}
REQUIRED_BUNDLE_FILES = ("Product.xlam",)
SHA64 = re.compile(r"^[0-9a-f]{64}$")
CONTROL_BUNDLE_PATHS = {"manifest.json", "SHA256SUMS", *REQUIRED_BUNDLE_FILES}


def _relative(path: Path, root: Path) -> str:
    return path.resolve().relative_to(root.resolve()).as_posix()


def _safe_directory(root: Path) -> str:
    for candidate in (root.resolve(), *root.resolve().parents):
        if (candidate / ".git").exists():
            value = candidate.as_posix()
            if re.match(r"^[A-Za-z]:/Mac/", value):
                value = "//Mac/" + value.split("/Mac/", 1)[1]
            return value
    return root.resolve().as_posix()


def _git(root: Path, *args: str) -> str:
    try:
        result = subprocess.run(
            ["git", "-c", f"safe.directory={_safe_directory(root)}", "-C", str(root), *args],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
    except OSError as exc:
        raise FreezeError(f"git unavailable: {exc}") from exc
    if result.returncode != 0:
        raise FreezeError(f"git command failed: {' '.join(args)}: {result.stderr.strip()}")
    return result.stdout.strip()


def _git_path(root: Path, relative: str) -> tuple[Path, str]:
    """Resolve a source-root path to the enclosing repository path."""
    repository = Path(_git(root, "rev-parse", "--show-toplevel")).resolve()
    absolute = (root / relative).resolve()
    try:
        repository_relative = absolute.relative_to(repository).as_posix()
    except ValueError as exc:
        raise FreezeError(f"source path escapes Git repository: {absolute}") from exc
    return repository, repository_relative


def _source_files(root: Path, source_dirs: Sequence[str]) -> list[Path]:
    paths: list[Path] = []
    seen: set[str] = set()
    for directory in source_dirs:
        base = root / directory
        if not base.is_dir():
            raise FreezeError(f"source directory missing: {directory}")
        for path in base.rglob("*"):
            if not path.is_file() or any(part in IGNORED_NAMES for part in path.parts) or path.suffix in IGNORED_SUFFIXES:
                continue
            relative = _relative(path, root)
            if relative not in SELF_PATHS:
                if relative in seen:
                    raise FreezeError(f"duplicate source path: {relative}")
                seen.add(relative)
                paths.append(path)
    return sorted(paths, key=lambda path: _relative(path, root))


def _rule_pattern(rule: Mapping[str, Any]) -> str:
    pattern = rule.get("pattern", rule.get("glob"))
    if not isinstance(pattern, str) or not pattern.strip():
        raise FreezeError("provenance rule pattern must be a non-empty string")
    return pattern


def _matches(relative: str, pattern: str) -> bool:
    # pathlib.PurePath.match handles the useful ** semantics and is portable;
    # fnmatch is retained for simple patterns such as README.md.
    from fnmatch import fnmatchcase
    from pathlib import PurePosixPath

    return fnmatchcase(relative, pattern) or PurePosixPath(relative).match(pattern)


def _metadata(rule: Mapping[str, Any], relative: str) -> tuple[str, list[str], str]:
    origin = rule.get("origin")
    spec_ids = rule.get("spec_ids")
    license_name = rule.get("license")
    if not isinstance(origin, str) or not origin.strip():
        raise FreezeError(f"{relative}: origin is required")
    if not isinstance(spec_ids, list) or not spec_ids or not all(isinstance(item, str) and item.strip() for item in spec_ids):
        raise FreezeError(f"{relative}: spec_ids must be a non-empty string list")
    if not isinstance(license_name, str) or not license_name.strip():
        raise FreezeError(f"{relative}: license is required")
    return origin, list(spec_ids), license_name


def _commit_for(root: Path, relative: str, commits: Mapping[str, str] | None) -> str:
    if commits and relative in commits:
        commit = commits[relative]
        if not isinstance(commit, str) or not commit.strip():
            raise FreezeError(f"{relative}: created_commit is empty")
        return commit
    repository, repository_relative = _git_path(root, relative)
    value = _git(repository, "log", "-1", "--format=%H", "--", repository_relative)
    if not value:
        raise FreezeError(f"{relative}: SOURCE_UNTRACKED (no created commit)")
    return value


def build_manifest(
    root: Path,
    source_dirs: Sequence[str] = DEFAULT_SOURCE_DIRS,
    rules: Sequence[Mapping[str, Any]] | None = None,
    commits: Mapping[str, str] | None = None,
) -> dict[str, Any]:
    """Return a deterministic per-file provenance manifest.

    ``commits`` is an explicit fixture seam for isolated tests. Production
    callers omit it so Git history is consulted and untracked files fail.
    """
    root = root.resolve()
    if not root.is_dir():
        raise FreezeError(f"root does not exist: {root}")
    if rules is None:
        origins_path = root / "provenance" / "origins.json"
        try:
            origins = json.loads(origins_path.read_text(encoding="utf-8"))
            rules = origins.get("rules", [])
        except (OSError, json.JSONDecodeError) as exc:
            raise FreezeError(f"cannot read provenance origins: {exc}") from exc
    if not rules:
        raise FreezeError("provenance rules must be non-empty")

    files = _source_files(root, source_dirs)
    rows: list[dict[str, Any]] = []
    for path in files:
        relative = _relative(path, root)
        matched = [rule for rule in rules if isinstance(rule, Mapping) and _matches(relative, _rule_pattern(rule))]
        if len(matched) != 1:
            raise FreezeError(f"{relative}: provenance rule must match exactly one (matched {len(matched)})")
        origin, spec_ids, license_name = _metadata(matched[0], relative)
        rows.append(
            {
                "path": relative,
                "origin": origin,
                "spec_ids": spec_ids,
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                "created_commit": _commit_for(root, relative, commits),
                "license": license_name,
            }
        )
    digest_input = "\n".join(f"{row['path']}\t{row['sha256']}" for row in rows).encode("utf-8")
    return {
        "version": 1,
        "policy": "clean-room",
        "root": root.as_posix(),
        "source_dirs": list(source_dirs),
        "source_tree_digest": hashlib.sha256(digest_input).hexdigest(),
        "files": rows,
    }


def require_clean_source(paths: Iterable[str | Path], root: Path | None = None) -> None:
    """Reject dirty, untracked, or missing source paths using Git status.

    ``git status`` omits ignored paths, so a status-only check would let an
    ignored source file through as if it were committed.  Tracking is checked
    explicitly before the cleanliness check to keep the freeze fail-closed.
    """
    path_list = [Path(path) for path in paths]
    if not path_list:
        raise FreezeError("no source paths supplied")
    source_root = (root or Path.cwd()).resolve()
    repository = Path(_git(source_root, "rev-parse", "--show-toplevel")).resolve()
    for path in path_list:
        absolute = path.resolve() if path.is_absolute() else (source_root / path).resolve()
        if not absolute.is_file():
            raise FreezeError(f"SOURCE_MISSING: {absolute}")
        relative = _relative(absolute, source_root)
        try:
            repository_relative = absolute.relative_to(repository).as_posix()
        except ValueError as exc:
            raise FreezeError(f"SOURCE_UNTRACKED_OR_DIRTY: {relative} (outside repository)") from exc
        try:
            _git(repository, "ls-files", "--error-unmatch", "--", repository_relative)
        except FreezeError as exc:
            raise FreezeError(f"SOURCE_UNTRACKED_OR_DIRTY: {relative} (not tracked)") from exc
        status = _git(repository, "status", "--porcelain=v1", "--untracked-files=all", "--", repository_relative)
        if status:
            raise FreezeError(f"SOURCE_UNTRACKED_OR_DIRTY: {relative}: {status.splitlines()[0]}")


def _write_json_atomic(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=path.parent, delete=False) as handle:
        json.dump(value, handle, ensure_ascii=False, indent=2, sort_keys=False)
        handle.write("\n")
        temporary = Path(handle.name)
    os.replace(temporary, path)


def _bundle_records(bundle_root: Path, candidate: Path) -> list[dict[str, Any]]:
    """Return the single Product XLAM record and reject external runtime payloads."""
    if candidate.name != "Product.xlam":
        return [{"path": candidate.name, "sha256": hashlib.sha256(candidate.read_bytes()).hexdigest()}]
    if not candidate.is_file():
        raise FreezeError("bundle member missing or unsafe: Product.xlam")
    forbidden = [
        path for path in bundle_root.rglob("*")
        if path.is_file() and (path.suffix.lower() in {".exe", ".dll", ".pyd"} or path.name == "Product.xlam.sidecar-manifest.json")
    ]
    if forbidden:
        raise FreezeError("single XLAM bundle contains forbidden external runtime payload")
    return [{"path": "Product.xlam", "sha256": hashlib.sha256(candidate.read_bytes()).hexdigest()}]


def _bundle_digest(records: Sequence[Mapping[str, Any]]) -> str:
    payload = "".join(f"{row['path']}\n{row['sha256']}\n" for row in sorted(records, key=lambda row: str(row["path"]).encode("utf-8"))).encode()
    return hashlib.sha256(payload).hexdigest()


def freeze(root: Path, candidate: Path, source_dirs: Sequence[str] = DEFAULT_SOURCE_DIRS) -> dict[str, Any]:
    """Freeze a clean Git source and candidate artifact into ``out/frozen``."""
    root = root.resolve()
    candidate = candidate.resolve()
    if not candidate.is_file():
        raise FreezeError(f"candidate missing: {candidate}")
    files = _source_files(root, source_dirs)
    require_clean_source(files, root=root)
    manifest = build_manifest(root, source_dirs=source_dirs)
    source_commit = _git(root, "rev-parse", "HEAD")
    artifact_sha = hashlib.sha256(candidate.read_bytes()).hexdigest()
    records = _bundle_records(candidate.parent, candidate)
    bundle_digest = _bundle_digest(records)
    identity = {"source_commit": source_commit, "artifact_sha256": artifact_sha, "product_bundle_sha256": bundle_digest}
    output_root = root / "out" / "frozen"
    preserved = output_root / source_commit / bundle_digest
    current = output_root / "current"
    if os.path.lexists(preserved):
        raise FreezeError(f"frozen product bundle already exists: {preserved}")
    preserved.mkdir(parents=True, exist_ok=True)
    for record in records:
        source = candidate.parent / record["path"]
        destination = preserved / record["path"]
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
    manifest["frozen_source_commit"] = source_commit
    manifest["artifact_sha256"] = artifact_sha
    manifest["product_bundle_sha256"] = bundle_digest
    manifest["product_files"] = records
    manifest_path = preserved / "source-manifest.json"
    _write_json_atomic(manifest_path, manifest)
    freeze_manifest = {
        "version": 1,
        "frozen_source_commit": source_commit,
        "source_tree_digest": manifest["source_tree_digest"],
        "artifact_sha256": artifact_sha,
        "artifact": candidate.name,
        "product_files": records,
        "product_bundle_sha256": bundle_digest,
    }
    _write_json_atomic(preserved / "freeze-manifest.json", freeze_manifest)
    current_stage = Path(tempfile.mkdtemp(prefix=".current-", dir=output_root))
    for record in records:
        source = preserved / record["path"]
        destination = current_stage / record["path"]
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)
    shutil.copy2(manifest_path, current_stage / "source-manifest.json")
    shutil.copy2(preserved / "freeze-manifest.json", current_stage / "freeze-manifest.json")
    if current.exists():
        shutil.rmtree(current)
    os.replace(current_stage, current)
    return {**identity, "source_tree_digest": manifest["source_tree_digest"], "preserved": preserved.as_posix(), "product_files": records}


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    command = sub.add_parser("freeze")
    command.add_argument("--root", type=Path, required=True)
    command.add_argument("--candidate", type=Path, required=True)
    command.add_argument("--source-dir", action="append", dest="source_dirs")
    args = parser.parse_args(argv)
    try:
        result = freeze(args.root, args.candidate, tuple(args.source_dirs or DEFAULT_SOURCE_DIRS))
    except FreezeError as exc:
        print(f"FAIL source freeze: {exc}")
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
