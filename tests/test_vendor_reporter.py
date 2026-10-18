"""Tests for vendor payment history reports."""

from datetime import date

import pandas as pd

from nv_billbook.vendor_reporter import build_vendor_summary, filter_vendors


def _debits() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "date": date(2026, 4, 1),
                "counterparty": "Gowrishankar Fuel Station",
                "purpose": "Fuel",
                "category": "Fuel / DG",
                "reference": "1",
                "debit": 1000.0,
                "balance": 9000.0,
                "source_file": "April.xls",
            },
            {
                "date": date(2026, 5, 1),
                "counterparty": " gowrishankar fuel station ",
                "purpose": "Fuel",
                "category": "Fuel / DG",
                "reference": "2",
                "debit": 1500.0,
                "balance": 7500.0,
                "source_file": "May.pdf",
            },
        ]
    )


def test_vendor_summary_groups_case_and_spacing_variants() -> None:
    summary = build_vendor_summary(_debits())

    assert len(summary) == 1
    assert summary.loc[0, "Total Paid (₹)"] == 2500.0
    assert summary.loc[0, "Payments"] == 2


def test_vendor_filter_matches_partial_name() -> None:
    filtered = filter_vendors(_debits(), "gowrishankar")

    assert len(filtered) == 2
