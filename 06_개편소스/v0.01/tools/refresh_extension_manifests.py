from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


ORDER = ("Frame", "AI", "G005", "Template", "Data", "Draw", "File", "Calculator", "Symbols")


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    builder = root / "tools" / "build_extension_manifest.py"
    for suite in ORDER:
        manifest_path = root / "tests" / "manifests" / f"{suite}.json"
        document = json.loads(manifest_path.read_text(encoding="utf-8"))
        args = [
            sys.executable,
            str(builder),
            "--root",
            str(root),
            "--suite",
            document["suite"],
            "--owner",
            document["owner_phase"],
            "--base-allowlist-from",
            "tests/windows/Build-TestXlam.ps1",
            "--output",
            str(manifest_path.relative_to(root)),
        ]
        for dependency in document["dependencies"]:
            args.extend(
                [
                    "--dependency",
                    "|".join(
                        (
                            dependency["path"],
                            dependency["suite"],
                            dependency["owner_phase"],
                        )
                    ),
                ]
            )
        args.extend(module["path"] for module in document["modules"])
        subprocess.run(args, check=True, cwd=root)
    print("refreshed extension manifests:", ", ".join(ORDER))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
