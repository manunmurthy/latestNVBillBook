"""Build the flat-wise monthly maintenance reconciliation."""

from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP

import pandas as pd

from nv_billbook.flats_registry import FlatsRegistry


RECONCILIATION_COLUMNS = [
    "Flat",
    "Sqft",
    "Expected Maintenance",
    "Water Bill",
    "Expected",
    "Paid",
    "Amount Due",
    "Extra Paid",
    "Status",
]
MONEY_QUANTUM = Decimal("0.01")


def _money(value: object) -> Decimal:
    """Return a monetary value rounded to paise without using a tolerance."""
    if value is None or pd.isna(value):
        value = 0
    return Decimal(str(value)).quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP)


def build_maintenance_reconciliation(
    credits: pd.DataFrame,
    registry: FlatsRegistry,
    water_bill_per_month: object | None = None,
) -> pd.DataFrame:
    """Return exactly one reconciliation row for every registered flat.

    Only credits resolved to a flat in the registry are included in the paid
    total. Payments for unknown or unassigned flats remain on Credits for review.
    """
    rate = _money(registry.meta.get("maintenance_rate_per_sqft", 2.5))
    water_bill = _money(
        registry.meta.get("water_bill_per_month", 0)
        if water_bill_per_month is None
        else water_bill_per_month
    )
    paid_by_flat = {flat_no: Decimal("0.00") for flat_no in registry.flats}

    if not credits.empty:
        for _, transaction in credits.iterrows():
            flat_no = transaction.get("flat_no")
            info = registry.get(flat_no) if isinstance(flat_no, str) else None
            if info:
                paid_by_flat[info.flat_no] += _money(transaction.get("credit"))

    rows: list[dict[str, object]] = []
    for flat_no, info in registry.flats.items():
        expected_maintenance = _money(Decimal(info.sqft) * rate)
        expected = expected_maintenance + water_bill
        paid = paid_by_flat[flat_no].quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP)

        if paid == Decimal("0.00"):
            status = "Pending"
        elif paid == expected:
            status = "Paid"
        elif paid > expected:
            status = "Excess"
        else:
            status = "Short"

        rows.append(
            {
                "Flat": info.flat_no,
                "Sqft": info.sqft,
                "Expected Maintenance": float(expected_maintenance),
                "Water Bill": float(water_bill),
                "Expected": float(expected),
                "Paid": float(paid),
                "Amount Due": float(max(expected - paid, Decimal("0.00"))),
                "Extra Paid": float(max(paid - expected, Decimal("0.00"))),
                "Status": status,
            }
        )

    return pd.DataFrame(rows, columns=RECONCILIATION_COLUMNS)
