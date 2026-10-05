#!/usr/bin/env python3
"""Build a deterministic three-way source inventory for the r53 integration."""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import subprocess
from collections import Counter
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_BASE = "91b6456a088f6b1827c1966b15fdeb1d70628a9f"
DEFAULT_R52 = "641869480d2e4c4194321435aa2a52777ba4a36b"
PROTECTED = "src/vba/features/data/NxDataUi.bas"
R52_PRESERVED_DELETIONS = {
    "tests/vba/product/T_ProductExhaustive.bas",
}
EXCLUDED_DIRECTORIES = {
    ".git",
    ".pytest_cache",
    "__pycache__",
    "out",
    "reports",
    "tmp",
}


def run_git(repo: Path, *arguments: str) -> str:
    completed = subprocess.run(
        [
            "git",
            "-c",
            f"safe.directory={repo.as_posix()}",
            "-c",
            "core.quotepath=false",
            "-C",
            str(repo),
            *arguments,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="strict",
        check=False,
    )
    if completed.returncode:
        raise RuntimeError(completed.stderr.strip() or completed.stdout.strip())
    return completed.stdout


def find_repo_root(start: Path) -> Path:
    for candidate in (start, *start.parents):
        if (candidate / ".git").exists():
            return candidate
    raise RuntimeError(f"Git worktree root was not found above: {start}")


def sha256(data: bytes | None) -> str:
    return "" if data is None else hashlib.sha256(data).hexdigest()


def git_blob(data: bytes | None) -> str:
    if data is None:
        return ""
    header = f"blob {len(data)}\0".encode("ascii")
    return hashlib.sha1(header + data).hexdigest()


def included(relative: str) -> bool:
    path = Path(relative)
    return (
        path.name != ".DS_Store"
        and not any(part in EXCLUDED_DIRECTORIES for part in path.parts)
    )


def collect(root: Path) -> dict[str, bytes]:
    rows: dict[str, bytes] = {}
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        relative = path.relative_to(root).as_posix()
        if included(relative):
            rows[relative] = path.read_bytes()
    return rows


def base_blobs(repo: Path, commit: str, prefix: str) -> dict[str, str]:
    output = run_git(repo, "ls-tree", "-r", commit, "--", prefix)
    rows: dict[str, str] = {}
    prefix_with_slash = prefix.rstrip("/") + "/"
    for line in output.splitlines():
        metadata, full_path = line.split("\t", 1)
        relative = full_path.removeprefix(prefix_with_slash)
        if included(relative):
            rows[relative] = metadata.split()[2]
    return rows


def classify(base: str, canonical: bytes | None, r52: bytes | None) -> str:
    canonical_blob = git_blob(canonical)
    r52_blob = git_blob(r52)
    if not base:
        if canonical is not None and r52 is not None:
            return "SAME_NEW" if canonical == r52 else "OVERLAP_NEW_DIFFERENT"
        if canonical is not None:
            return "CANONICAL_NEW"
        if r52 is not None:
            return "R52_NEW"
        raise ValueError("empty three-way row")
    if canonical is None and r52 is None:
        return "BOTH_DELETED"
    if canonical is None:
        return "CANONICAL_DELETED" if r52_blob == base else "DELETE_OVERLAP_DIFFERENT"
    if r52 is None:
        return "R52_DELETED" if canonical_blob == base else "DELETE_OVERLAP_DIFFERENT"
    if canonical_blob == base and r52_blob == base:
        return "UNCHANGED"
    if canonical == r52:
        return "SAME_CHANGE"
    if canonical_blob == base:
        return "R52_ONLY"
    if r52_blob == base:
        return "CANONICAL_ONLY"
    return "OVERLAP_DIFFERENT"


def default_decision(classification: str) -> tuple[str, str]:
    if classification in {"CANONICAL_ONLY", "CANONICAL_NEW"}:
        return "KEEP_CANONICAL", "기준 프로젝트에서만 변경됨"
    if classification in {"R52_ONLY", "R52_NEW"}:
        return "KEEP_R52", "r52 기준선에서만 변경됨"
    if classification in {"SAME_CHANGE", "SAME_NEW"}:
        return "KEEP_SAME", "양쪽 내용이 동일함"
    if classification == "BOTH_DELETED":
        return "KEEP_DELETED", "양쪽에서 삭제됨"
    return "REVIEW", "양쪽 변경 또는 삭제를 수동 판정해야 함"


def applied_decision(relative: str, classification: str) -> tuple[str, str]:
    if relative == PROTECTED:
        return "KEEP_CANONICAL", "사용자 보호 파일의 기준 프로젝트 버전을 채택함"
    if relative in R52_PRESERVED_DELETIONS:
        return "KEEP_R52", "r52 실제 Excel 검증에 필요한 테스트 소유 fixture를 보존함"
    if classification in {"CANONICAL_ONLY", "CANONICAL_NEW", "CANONICAL_DELETED", "R52_DELETED"}:
        return "KEEP_CANONICAL", "기준 프로젝트 변경 또는 삭제를 r53에 반영함"
    if classification in {"R52_ONLY", "R52_NEW", "DELETE_OVERLAP_DIFFERENT"}:
        return "KEEP_R52", "r52 네이티브 구현 또는 검증 자산을 r53에 보존함"
    if classification in {"OVERLAP_DIFFERENT", "OVERLAP_NEW_DIFFERENT"}:
        return "MERGED", "공통 부모 기준 3방향 통합 후 계약에 맞게 충돌을 해소함"
    return default_decision(classification)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--canonical-root", required=True, type=Path)
    parser.add_argument("--r52-root", required=True, type=Path)
    parser.add_argument("--r53-root", default=ROOT, type=Path)
    parser.add_argument("--base-commit", default=DEFAULT_BASE)
    parser.add_argument("--r52-commit", default=DEFAULT_R52)
    parser.add_argument("--integration-applied", action="store_true")
    parser.add_argument("--output-json", default=ROOT / "reports" / "r53-integration-inventory.json", type=Path)
    parser.add_argument("--output-csv", default=ROOT / "reports" / "r53-integration-decisions.csv", type=Path)
    args = parser.parse_args()

    canonical_root = args.canonical_root.resolve()
    r52_root = args.r52_root.resolve()
    r53_root = args.r53_root.resolve()
    for label, path in (("canonical", canonical_root), ("r52", r52_root), ("r53", r53_root)):
        if not path.is_dir():
            parser.error(f"{label} source root is missing: {path}")

    repo = find_repo_root(r53_root)
    resolved_repo = Path(run_git(repo, "rev-parse", "--show-toplevel").strip()).resolve()
    if resolved_repo != repo:
        parser.error(f"unexpected Git root: {resolved_repo}")
    prefix = r53_root.relative_to(repo).as_posix()
    base = base_blobs(repo, args.base_commit, prefix)
    canonical = collect(canonical_root)
    r52 = collect(r52_root)
    r53 = collect(r53_root)
    paths = sorted(set(base) | set(canonical) | set(r52))

    rows: list[dict[str, str]] = []
    decisions: list[dict[str, str]] = []
    for relative in paths:
        classification = classify(base.get(relative, ""), canonical.get(relative), r52.get(relative))
        if classification == "UNCHANGED":
            continue
        if args.integration_applied:
            decision, reason = applied_decision(relative, classification)
        else:
            decision, reason = default_decision(classification)
        rows.append(
            {
                "path": relative,
                "classification": classification,
                "base_git_blob": base.get(relative, ""),
                "canonical_sha256": sha256(canonical.get(relative)),
                "r52_sha256": sha256(r52.get(relative)),
                "r53_sha256": sha256(r53.get(relative)),
            }
        )
        decisions.append(
            {
                "path": relative,
                "classification": classification,
                "decision": decision,
                "reason": reason,
            }
        )

    protected_canonical = canonical.get(PROTECTED)
    protected_r52 = r52.get(PROTECTED)
    if protected_canonical is None or protected_r52 is None:
        parser.error(f"protected file is missing: {PROTECTED}")

    counts = dict(sorted(Counter(row["classification"] for row in rows).items()))
    report = {
        "schema_version": 1,
        "base_commit": args.base_commit,
        "r52_commit": args.r52_commit,
        "canonical_source": str(canonical_root),
        "r52_source": str(r52_root),
        "r53_source": str(r53_root),
        "integration_applied": args.integration_applied,
        "excluded_directories": sorted(EXCLUDED_DIRECTORIES),
        "protected": {
            "path": PROTECTED,
            "canonical_sha256": sha256(protected_canonical),
            "r52_sha256": sha256(protected_r52),
            "r53_sha256": sha256(r53.get(PROTECTED)),
        },
        "counts": counts,
        "rows": rows,
    }
    args.output_json.parent.mkdir(parents=True, exist_ok=True)
    args.output_json.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, sort_keys=False) + "\n",
        encoding="utf-8",
        newline="\n",
    )
    args.output_csv.parent.mkdir(parents=True, exist_ok=True)
    with args.output_csv.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=("path", "classification", "decision", "reason"),
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(decisions)

    print(json.dumps({"rows": len(rows), "counts": counts}, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
