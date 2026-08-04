"""Build flat-wise collection history and multi-month reconciliations."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd

from nv_billbook.flats_registry import FlatsRegistry


HISTORY_COLUMNS = ["Flat", "Date", "Month", "Payer", "Payment Type", "Amount"]
MONTHLY_RECONCILIATION_COLUMNS = [
    "Flat",
    "Month",
    "Sqft",
    "Expected Maintenance",
    "Water Bill",
    "Expected",
    "Paid",
    "Amount Due",
    "Extra Paid",
    "Status",
]
PERIOD_SUMMARY_COLUMNS = [
    "Flat",
    "Total Expected",
    "Total Paid",
    "Total Amount Due",
    "Total Extra Paid",
    "Overall Status",
]
MONEY_QUANTUM = Decimal("0.01")


class MissingWaterBillError(ValueError):
    """Raised when a statement month has no configured water bill."""


def _money(value: object) -> Decimal:
    if value is None or pd.isna(value):
        value = 0
    return Decimal(str(value)).quantize(MONEY_QUANTUM, rounding=ROUND_HALF_UP)


def _month_label(month_key: str) -> str:
    return pd.Period(month_key, freq="M").strftime("%b %Y")


def month_keys_for_period(period_start: date, period_end: date) -> list[str]:
    """Return inclusive YYYY-MM keys for a statement period."""
    return [str(month) for month in pd.period_range(period_start, period_end, freq="M")]


def month_keys_from_transactions(transactions: pd.DataFrame) -> list[str]:
    """Return the distinct calendar months represented by transaction dates."""
    if "date" not in transactions.columns:
        return []
    dates = pd.to_datetime(transactions["date"], errors="coerce").dropna()
    return sorted(dates.dt.strftime("%Y-%m").unique().tolist())


def build_flat_transaction_history(credits: pd.DataFrame, registry: FlatsRegistry) -> pd.DataFrame:
    """Return every mapped incoming payment, sorted by flat and transaction date."""
    rows: list[dict[str, object]] = []
    for _, transaction in credits.iterrows():
        flat_no = transaction.get("flat_no")
        info = registry.get(flat_no) if isinstance(flat_no, str) else None
        if not info:
            continue

        transaction_date = pd.to_datetime(transaction.get("date"), errors="coerce")
        if pd.isna(transaction_date):
            continue

        rows.append(
            {
                "Flat": info.flat_no,
                "Date": transaction_date.strftime("%d/%m/%Y"),
                "Month": transaction_date.strftime("%b %Y"),
                "Payer": transaction.get("counterparty") or "",
                "Payment Type": transaction.get("category") or "",
                "Amount": float(_money(transaction.get("credit"))),
                "_sort_date": transaction_date,
            }
        )

    history = pd.DataFrame(rows)
    if history.empty:
        return pd.DataFrame(columns=HISTORY_COLUMNS)
    history = history.sort_values(["Flat", "_sort_date"], kind="stable")
    return history.drop(columns="_sort_date")[HISTORY_COLUMNS].reset_index(drop=True)


def build_monthly_reconciliation(
    credits: pd.DataFrame,
    registry: FlatsRegistry,
    month_keys: Sequence[str],
    water_bills_by_month: Mapping[str, object],
) -> pd.DataFrame:
    """Reconcile every registered flat for every statement month.

    A missing water bill is an error: the report must never silently use an
    incorrect expected amount.
    """
    missing_months = [month for month in month_keys if month not in water_bills_by_month]
    if missing_months:
        raise MissingWaterBillError(
            "Missing water bill configuration for: " + ", ".join(missing_months)
        )

    paid_by_flat_month = {
        (flat_no, month): Decimal("0.00")
        for flat_no in registry.flats
        for month in month_keys
    }
    for _, transaction in credits.iterrows():
        flat_no = transaction.get("flat_no")
        info = registry.get(flat_no) if isinstance(flat_no, str) else None
        transaction_date = pd.to_datetime(transaction.get("date"), errors="coerce")
        if not info or pd.isna(transaction_date):
            continue
        month_key = transaction_date.strftime("%Y-%m")
        if month_key in month_keys:
            paid_by_flat_month[(info.flat_no, month_key)] += _money(transaction.get("credit"))

    maintenance_rate = _money(registry.meta.get("maintenance_rate_per_sqft", 2.5))
    rows: list[dict[str, object]] = []
    for flat_no, info in registry.flats.items():
        maintenance = _money(Decimal(info.sqft) * maintenance_rate)
        for month_key in month_keys:
            water_bill = _money(water_bills_by_month[month_key])
            expected = maintenance + water_bill
            paid = paid_by_flat_month[(flat_no, month_key)]
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
                    "Month": _month_label(month_key),
                    "Sqft": info.sqft,
                    "Expected Maintenance": float(maintenance),
                    "Water Bill": float(water_bill),
                    "Expected": float(expected),
                    "Paid": float(paid),
                    "Amount Due": float(max(expected - paid, Decimal("0.00"))),
                    "Extra Paid": float(max(paid - expected, Decimal("0.00"))),
                    "Status": status,
                }
            )

    return pd.DataFrame(rows, columns=MONTHLY_RECONCILIATION_COLUMNS)


def build_period_summary(monthly_reconciliation: pd.DataFrame) -> pd.DataFrame:
    """Summarize expected and paid amounts for the complete statement period."""
    if monthly_reconciliation.empty:
        return pd.DataFrame(columns=PERIOD_SUMMARY_COLUMNS)

    rows: list[dict[str, object]] = []
    for flat_no, rows_for_flat in monthly_reconciliation.groupby("Flat", sort=False):
        expected = sum((_money(value) for value in rows_for_flat["Expected"]), Decimal("0.00"))
        paid = sum((_money(value) for value in rows_for_flat["Paid"]), Decimal("0.00"))
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
                "Flat": flat_no,
                "Total Expected": float(expected),
                "Total Paid": float(paid),
                "Total Amount Due": float(max(expected - paid, Decimal("0.00"))),
                "Total Extra Paid": float(max(paid - expected, Decimal("0.00"))),
                "Overall Status": status,
            }
        )

    return pd.DataFrame(rows, columns=PERIOD_SUMMARY_COLUMNS)
