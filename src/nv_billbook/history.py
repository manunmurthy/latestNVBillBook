"""CLI for a standalone three-month flat collection history report."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Mapping
from pathlib import Path

from nv_billbook.classifier import classify_transactions, split_by_type
from nv_billbook.collection_history import (
    build_flat_transaction_history,
    build_monthly_reconciliation,
    build_period_summary,
    month_keys_from_transactions,
)
from nv_billbook.config import Config
from nv_billbook.history_reporter import write_collection_history_report
from nv_billbook.main import _load_flats_registry
from nv_billbook.parser import parse_hdfc_statement


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create a flat-wise history report from one three-month HDFC statement."
    )
    parser.add_argument("--input", required=True, help="Three-month HDFC statement file (.xls or .xlsx)")
    parser.add_argument("--config", default="config.yaml", help="Path to config file")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    config_path = Path(args.config)
    if not config_path.exists():
        print(f"Config not found: {config_path}", file=sys.stderr)
        sys.exit(1)
    if not input_path.is_file():
        print(f"Input statement not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    config = Config.load(config_path)
    registry = _load_flats_registry(config)
    if not registry:
        print("A flats registry is required to create a collection history report.", file=sys.stderr)
        sys.exit(1)

    parsed = parse_hdfc_statement(input_path, config)
    month_keys = month_keys_from_transactions(parsed.transactions)
    if len(month_keys) != 3:
        print(
            "This command requires transactions from exactly three calendar months; "
            f"found: {', '.join(month_keys) or 'none'}.",
            file=sys.stderr,
        )
        sys.exit(1)

    water_bills = registry.meta.get("water_bills_by_month")
    if not isinstance(water_bills, Mapping):
        print(
            "Add meta.water_bills_by_month to flats.yaml before creating this report.",
            file=sys.stderr,
        )
        sys.exit(1)

    classified = classify_transactions(parsed.transactions, config, registry)
    credits = split_by_type(classified)["credits"]
    try:
        transaction_history = build_flat_transaction_history(credits, registry)
        monthly_reconciliation = build_monthly_reconciliation(
            credits, registry, month_keys, water_bills
        )
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    period_summary = build_period_summary(monthly_reconciliation)
    output_path = config.output_dir / (
        f"{month_keys[0]}_to_{month_keys[-1]}_collection_history.xlsx"
    )
    write_collection_history_report(
        output_path,
        transaction_history,
        monthly_reconciliation,
        period_summary,
    )
    print(f"Collection history report written: {output_path}")


if __name__ == "__main__":
    main()
