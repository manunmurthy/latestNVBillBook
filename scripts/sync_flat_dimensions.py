"""Sync exact SBA dimensions into the block flat files."""

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

    flat_files = [flats_path]
    for block_file in flats_data.get("block_files", {}).values():
        block_path = Path(block_file)
        if not block_path.is_absolute():
            block_path = root / block_path
        flat_files.append(block_path)

    flat_records: dict[str, tuple[Path, dict]] = {}
    for flat_file in flat_files:
        with flat_file.open(encoding="utf-8") as handle:
            data = yaml.safe_load(handle) or {}
        entries = data.get("flats", {})
        for flat_no, entry in entries.items():
            if flat_no in flat_records:
                raise ValueError(f"Duplicate flat {flat_no} in registry files")
            flat_records[flat_no] = (flat_file, entry)

    updated = 0
    for flat_no, sqft in dimensions.items():
        if flat_no not in flat_records:
            print(f"Warning: {flat_no} not found in flat registry files", file=sys.stderr)
            continue
        _, entry = flat_records[flat_no]
        if entry["sqft"] != sqft:
            entry["sqft"] = int(sqft)
            updated += 1

    for flat_no in flat_records:
        if flat_no not in dimensions:
            print(f"Warning: {flat_no} missing from flat_dimensions.yaml", file=sys.stderr)

    for flat_file in flat_files[1:]:
        with flat_file.open(encoding="utf-8") as handle:
            data = yaml.safe_load(handle) or {}
        data["flats"] = {
            flat_no: entry
            for flat_no, (record_file, entry) in flat_records.items()
            if record_file == flat_file
        }
        with flat_file.open("w", encoding="utf-8") as handle:
            yaml.safe_dump(data, handle, default_flow_style=False, allow_unicode=True, sort_keys=False, width=100)

    return updated


if __name__ == "__main__":
    root = Path(__file__).resolve().parents[1]
    count = sync(root)
    print(f"Synced {count} sqft value(s) from flat_dimensions.yaml → flats.yaml")
