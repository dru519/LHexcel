#!/usr/bin/env python3
"""Synchronize NxHost runtime identity and raw contract hashes."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CONTRACTS = {
    "bridge": ROOT / "contracts/bridge-contract.json",
    "feature": ROOT / "contracts/feature-contract.json",
    "command": ROOT / "contracts/command-contract.json",
}
VBA = ROOT / "src/vba/bridge/NxHostBridge.bas"
CS = ROOT / "src/dotnet/NxHost/BridgeService.cs"


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def replace_once(source: str, pattern: str, replacement: str, label: str) -> str:
    updated, count = re.subn(pattern, replacement, source, count=1, flags=re.MULTILINE)
    if count != 1:
        raise ValueError(f"{label} replacement count rejected: {count}")
    return updated


def expected_sources() -> dict[Path, str]:
    bridge = json.loads(CONTRACTS["bridge"].read_text(encoding="utf-8"))
    title = bridge["product_title"]
    hashes = {name: digest(path) for name, path in CONTRACTS.items()}

    vba = VBA.read_text(encoding="utf-8-sig")
    vba = replace_once(
        vba,
        r'Private Const NX_HOST_PRODUCT_TITLE As String = "[^"]+"',
        f'Private Const NX_HOST_PRODUCT_TITLE As String = "{title}"',
        "VBA product title",
    )
    for name, value in hashes.items():
        vba = replace_once(
            vba,
            rf'Private Const {name}_sha256 As String = "[0-9a-f]{{64}}"',
            f'Private Const {name}_sha256 As String = "{value}"',
            f"VBA {name} hash",
        )

    cs = CS.read_text(encoding="utf-8-sig")
    cs = replace_once(
        cs,
        r'public const string ProductTitle = "[^"]+";',
        f'public const string ProductTitle = "{title}";',
        "C# product title",
    )
    for name, member in (("bridge", "BridgeHash"), ("feature", "FeatureHash"), ("command", "CommandHash")):
        cs = replace_once(
            cs,
            rf'public const string {member} = "[0-9a-f]{{64}}";',
            f'public const string {member} = "{hashes[name]}";',
            f"C# {name} hash",
        )
    return {VBA: vba, CS: cs}


def atomic_write(path: Path, content: str) -> None:
    fd, temporary = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(content)
        os.replace(temporary, path)
    except Exception:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass
        raise


def main() -> int:
    parser = argparse.ArgumentParser()
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    args = parser.parse_args()

    expected = expected_sources()
    stale = [path for path, content in expected.items() if path.read_text(encoding="utf-8-sig") != content]
    if args.check:
        if stale:
            for path in stale:
                print(f"stale NxHost handshake source: {path.relative_to(ROOT)}")
            return 1
        return 0
    for path in stale:
        atomic_write(path, expected[path])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
