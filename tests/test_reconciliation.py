"""Tests for flat-wise maintenance reconciliation."""

import pandas as pd

from nv_billbook.flats_registry import FlatInfo, FlatsRegistry
from nv_billbook.reconciliation import (
    RECONCILIATION_COLUMNS,
    build_maintenance_reconciliation,
)


def test_reconciliation_has_one_row_per_flat_and_exact_statuses() -> None:
    registry = FlatsRegistry(
        meta={"maintenance_rate_per_sqft": 2.5, "water_bill_per_month": 100},
        flats={
            "A001": FlatInfo("A001", "A", 1000),
            "A002": FlatInfo("A002", "A", 1000),
            "A003": FlatInfo("A003", "A", 1000),
            "A004": FlatInfo("A004", "A", 1000),
        },
        unmapped_payers=[],
    )
    credits = pd.DataFrame(
        [
            {"flat_no": "A001", "credit": 2000},
            {"flat_no": "A001", "credit": 600},
            {"flat_no": "A002", "credit": 2700},
            {"flat_no": "A003", "credit": 2500},
            {"flat_no": None, "credit": 9999},
        ]
    )

    result = build_maintenance_reconciliation(credits, registry)

    assert list(result.columns) == RECONCILIATION_COLUMNS
    assert list(result["Flat"]) == ["A001", "A002", "A003", "A004"]
    assert list(result["Expected Maintenance"]) == [2500.0] * 4
    assert list(result["Water Bill"]) == [100.0] * 4
    assert list(result["Expected"]) == [2600.0] * 4
    assert list(result["Paid"]) == [2600.0, 2700.0, 2500.0, 0.0]
    assert list(result["Amount Due"]) == [0.0, 0.0, 100.0, 2600.0]
    assert list(result["Extra Paid"]) == [0.0, 100.0, 0.0, 0.0]
    assert list(result["Status"]) == ["Paid", "Excess", "Short", "Pending"]


def test_command_line_water_bill_overrides_registry_value() -> None:
    registry = FlatsRegistry(
        meta={"maintenance_rate_per_sqft": 2.5, "water_bill_per_month": 100},
        flats={"A001": FlatInfo("A001", "A", 1000)},
        unmapped_payers=[],
    )

    result = build_maintenance_reconciliation(
        pd.DataFrame(), registry, water_bill_per_month=200
    )

    assert result.loc[0, "Water Bill"] == 200.0
    assert result.loc[0, "Expected"] == 2700.0
