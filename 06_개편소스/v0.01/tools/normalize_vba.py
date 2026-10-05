from __future__ import annotations

"""Deterministic VBA export normalisation; no reference files are opened."""

import argparse
import hashlib
import json
import re
from pathlib import Path


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = []
    for line in text.split("\n"):
        line = line.rstrip()
        if re.match(r"^(Attribute\s+(VB_Name|VB_GlobalNameSpace|VB_Creatable|VB_PredeclaredId|VB_Exposed|VB_Description)\s*=|VERSION\s+\d|Begin\s+VB\.)", line, re.I):
            continue
        lines.append(line)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines) + "\n"


def normalize_vba(text: str) -> str:
    """Compatibility alias used by the independence comparator."""
    return normalize_text(text)


def normalize_file(source: Path, destination: Path | None = None) -> dict[str, object]:
    data = normalize_text(source.read_text(encoding="utf-8-sig"))
    if destination is not None:
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_text(data, encoding="utf-8", newline="\n")
    return {"source": source.name, "sha256": hashlib.sha256(data.encode()).hexdigest(), "line_count": len(data.splitlines()), "text": data}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    result = normalize_file(args.source, args.output)
    print(json.dumps({k: v for k, v in result.items() if k != "text"}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
