"""Inventory real execution routes; never infer timings from test PASS status."""
import argparse
import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def collect(root):
    features = json.loads((root / "contracts/feature-contract.json").read_text(encoding="utf-8"))
    commands = json.loads((root / "contracts/command-contract.json").read_text(encoding="utf-8"))
    routes = json.loads((root / "contracts/exhaustive-native-contract.json").read_text(encoding="utf-8"))
    definitions = {("feature", f["id"]): f for f in features["features"]}
    definitions.update({("command", c["id"]): c for c in commands["commands"]})
    result = []
    for route in routes["cases"]:
        definition = definitions[(route["kind"], route["route_id"])]
        navigation = definition["navigation"]
        for profile in ("internal-xlam", "enhanced-dll"):
            result.append({
                "release": routes["release_id"], "kind": route["kind"],
                "route_id": route["route_id"],
                "label": definition["ribbon"]["label"] if route["kind"] == "feature" else definition["label_ko"],
                "group": navigation["ribbon_group_id"], "profile": profile,
                "available": profile in navigation["profile_availability"],
                "surface": navigation["launch_surface"],
                "popup_ms": "NOT_RUN", "preview_ms": "NOT_RUN", "execution_ms": "NOT_RUN",
                "input_cells": "", "samples": 0, "fixture_included": "", "evidence": "",
            })
    return result


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    rows = collect(ROOT)
    # New artifact per run; do not replace measured historical reports.
    with args.output.open("x", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)
    print(json.dumps({"rows": len(rows), "routes": len({r["route_id"] for r in rows}),
                      "timing_status": "NOT_RUN", "output": str(args.output)}))


if __name__ == "__main__":
    main()
