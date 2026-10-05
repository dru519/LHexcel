"""PowerShell/VBA가 호출하는 JSON 표준 출력 CLI."""

from __future__ import annotations

import argparse
import json
import sys

from .builder import build_hwpx
from .cleanup import cleanup_expired
from .validator import validate_hwpx


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="naeexcel_hwpx")
    sub = parser.add_subparsers(dest="command", required=True)
    build = sub.add_parser("build")
    build.add_argument("--input", required=True)
    build.add_argument("--output", required=True)
    validate = sub.add_parser("validate")
    validate.add_argument("--input", required=True)
    cleanup = sub.add_parser("cleanup")
    cleanup.add_argument("--directory", required=True)
    cleanup.add_argument("--hours", type=float, default=24.0)
    args = parser.parse_args(argv)
    try:
        if args.command == "build":
            result = build_hwpx(args.input, args.output)
        elif args.command == "validate":
            result = validate_hwpx(args.input)
        else:
            result = cleanup_expired(args.directory, args.hours)
    except Exception as exc:
        result = {"status": "FAIL", "error": type(exc).__name__, "message": str(exc)}
    print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    return 0 if result.get("status") == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
