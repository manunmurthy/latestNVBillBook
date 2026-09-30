"""Tests for the split flat registry loader."""

from pathlib import Path

import yaml

from nv_billbook.flats_registry import FlatsRegistry


def test_load_combines_flat_block_files(tmp_path: Path) -> None:
    (tmp_path / "flats_a.yaml").write_text(
        yaml.safe_dump({"flats": {"A001": {"block": "A", "sqft": 1000}}}),
        encoding="utf-8",
    )
    (tmp_path / "flats_b.yaml").write_text(
        yaml.safe_dump({"flats": {"B001": {"block": "B", "sqft": 1200}}}),
        encoding="utf-8",
    )
    (tmp_path / "flats.yaml").write_text(
        yaml.safe_dump(
            {
                "meta": {"maintenance_rate_per_sqft": 2.5},
                "block_files": {"A": "flats_a.yaml", "B": "flats_b.yaml"},
                "unmapped_payers": {"entries": []},
            }
        ),
        encoding="utf-8",
    )

    registry = FlatsRegistry.load(tmp_path / "flats.yaml")

    assert list(registry.flats) == ["A001", "B001"]
