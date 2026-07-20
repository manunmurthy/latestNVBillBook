"""Flat registry — SBA dimensions and bank-statement payer names."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


def _normalize_name(value: str) -> str:
    return re.sub(r"\s+", " ", value.upper().strip())


@dataclass
class FlatInfo:
    flat_no: str
    block: str
    sqft: int
    payer_names: list[str] = field(default_factory=list)
    notes: str = ""

    @property
    def maintenance_amount(self) -> float:
        return 0.0  # set by registry using rate


@dataclass
class FlatsRegistry:
    meta: dict[str, Any]
    flats: dict[str, FlatInfo]
    unmapped_payers: list[dict[str, Any]]
    _payer_index: dict[str, str] = field(default_factory=dict, repr=False)

    @classmethod
    def load(cls, path: Path, dimensions_path: Path | None = None) -> "FlatsRegistry":
        with path.open(encoding="utf-8") as handle:
            raw: dict[str, Any] = yaml.safe_load(handle)

        sqft_map: dict[str, int] = {}
        if dimensions_path and dimensions_path.exists():
            with dimensions_path.open(encoding="utf-8") as handle:
                dim_raw = yaml.safe_load(handle)
            sqft_map = {
                flat_no.upper(): int(sqft)
                for flat_no, sqft in dim_raw.get("dimensions", {}).items()
            }

        flats: dict[str, FlatInfo] = {}
        for flat_no, entry in raw.get("flats", {}).items():
            flat_key = flat_no.upper()
            sqft = sqft_map.get(flat_key, int(entry.get("sqft", 0)))
            flats[flat_key] = FlatInfo(
                flat_no=flat_key,
                block=entry.get("block", flat_no[0]),
                sqft=sqft,
                payer_names=entry.get("payer_names", []) or [],
                notes=entry.get("notes", "") or "",
            )

        registry = cls(
            meta=raw.get("meta", {}),
            flats=flats,
            unmapped_payers=raw.get("unmapped_payers", {}).get("entries", []),
        )
        registry._build_payer_index()
        return registry

    def _build_payer_index(self) -> None:
        index: dict[str, str] = {}
        for flat_no, info in self.flats.items():
            for name in info.payer_names:
                normalized = _normalize_name(name)
                if normalized:
                    index[normalized] = flat_no
        self._payer_index = index

    def get(self, flat_no: str | None) -> FlatInfo | None:
        if not flat_no:
            return None
        return self.flats.get(flat_no.upper().replace("-", "").replace(" ", ""))

    def lookup_by_payer(self, payer_name: str) -> FlatInfo | None:
        normalized = _normalize_name(payer_name)
        if not normalized:
            return None

        flat_no = self._payer_index.get(normalized)
        if flat_no:
            return self.flats[flat_no]

        for key, candidate_flat in self._payer_index.items():
            if normalized in key or key in normalized:
                return self.flats[candidate_flat]

        return None

    def resolve_flat(self, flat_no: str | None, payer_name: str) -> tuple[str | None, FlatInfo | None]:
        info = self.get(flat_no)
        if info:
            return info.flat_no, info

        info = self.lookup_by_payer(payer_name)
        if info:
            return info.flat_no, info

        return flat_no, None

    def maintenance_amount(self, flat_no: str) -> float | None:
        info = self.get(flat_no)
        if not info or not info.sqft:
            return None
        rate = self.meta.get("maintenance_rate_per_sqft", 2.5)
        return round(info.sqft * rate, 2)
