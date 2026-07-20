"""Load and validate application configuration."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass
class Config:
    society_name: str
    short_name: str
    input_dir: Path
    output_dir: Path
    columns: dict[str, str]
    date_format: str
    categories: dict[str, list[str]] = field(default_factory=dict)
    bank: str = "hdfc"
    flats_registry_path: Path | None = None
    flat_dimensions_path: Path = Path("flat_dimensions.yaml")

    @classmethod
    def load(cls, path: Path) -> "Config":
        with path.open(encoding="utf-8") as handle:
            raw: dict[str, Any] = yaml.safe_load(handle)

        society = raw.get("society", {})
        paths = raw.get("paths", {})
        flats_registry = raw.get("flats_registry")
        flat_dimensions = raw.get("flat_dimensions", "flat_dimensions.yaml")
        return cls(
            society_name=society.get("name", "Nava Vaibhva"),
            short_name=society.get("short_name", "Nava Vaibhva"),
            input_dir=Path(paths.get("input_dir", "data/input")),
            output_dir=Path(paths.get("output_dir", "data/output")),
            columns=raw.get("columns", {}),
            date_format=raw.get("date_format", "%d/%m/%Y"),
            categories=raw.get("categories", {}),
            bank=raw.get("bank", "hdfc"),
            flats_registry_path=Path(flats_registry) if flats_registry else None,
            flat_dimensions_path=Path(flat_dimensions),
        )
