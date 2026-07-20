"""Sync exact SBA dimensions into flats.yaml from flat_dimensions.yaml."""

from __future__ import annotations

import sys
from pathlib import Path

import yaml


def sync(root: Path) -> int:
    dim_path = root / "flat_dimensions.yaml"
    flats_path = root / "flats.yaml"

    with dim_path.open(encoding="utf-8") as handle:
        dimensions = yaml.safe_load(handle)["dimensions"]

    with flats_path.open(encoding="utf-8") as handle:
        flats_data = yaml.safe_load(handle)

    updated = 0
    for flat_no, sqft in dimensions.items():
        if flat_no not in flats_data["flats"]:
            print(f"Warning: {flat_no} in dimensions but not in flats.yaml", file=sys.stderr)
            continue
        if flats_data["flats"][flat_no]["sqft"] != sqft:
            flats_data["flats"][flat_no]["sqft"] = int(sqft)
            updated += 1

    for flat_no in flats_data["flats"]:
        if flat_no not in dimensions:
            print(f"Warning: {flat_no} missing from flat_dimensions.yaml", file=sys.stderr)

    with flats_path.open("w", encoding="utf-8") as handle:
        yaml.safe_dump(flats_data, handle, default_flow_style=False, allow_unicode=True, sort_keys=False, width=100)

    example_path = root / "flats.example.yaml"
    with example_path.open("w", encoding="utf-8") as handle:
        yaml.safe_dump(flats_data, handle, default_flow_style=False, allow_unicode=True, sort_keys=False, width=100)

    return updated


if __name__ == "__main__":
    root = Path(__file__).resolve().parents[1]
    count = sync(root)
    print(f"Synced {count} sqft value(s) from flat_dimensions.yaml → flats.yaml")
