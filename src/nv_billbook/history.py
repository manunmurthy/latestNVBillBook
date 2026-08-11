"""CLI for a standalone multi-month flat collection history report."""

from __future__ import annotations

import argparse
import sys
from collections.abc import Mapping
from pathlib import Path

import pandas as pd

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
from nv_billbook.parser import find_statement_files


SUPPORTED_PERIOD_LENGTHS = {3, 6, 12}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create a flat-wise history report from one 3-, 6-, or 12-month HDFC statement."
    )
    parser.add_argument(
        "--input",
        required=True,
        help="HDFC statement file or folder covering 3, 6, or 12 calendar months",
    )
    parser.add_argument(
        "--flat",
        help="Optional flat number to generate a report for one flat only (for example A001)",
    )
    parser.add_argument("--config", default="config.yaml", help="Path to config file")
    return parser.parse_args()


def _load_statements(input_path: Path, config: Config) -> pd.DataFrame:
    if input_path.is_file():
        parsed = parse_hdfc_statement(input_path, config)
        return parsed.transactions

    if input_path.is_dir():
        files = sorted(find_statement_files(input_path))
        if not files:
            raise FileNotFoundError(f"No statement files found in folder: {input_path}")

        frames: list[pd.DataFrame] = []
        for file_path in files:
            parsed = parse_hdfc_statement(file_path, config)
            frames.append(parsed.transactions)

        combined = pd.concat(frames, ignore_index=True)
        combined = combined.sort_values(
            by=["date", "source_file", "date_str"],
            kind="stable",
            ignore_index=True,
        )
        return combined

    raise FileNotFoundError(f"Input statement not found: {input_path}")


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    config_path = Path(args.config)
    if not config_path.exists():
        print(f"Config not found: {config_path}", file=sys.stderr)
        sys.exit(1)
    config = Config.load(config_path)
    registry = _load_flats_registry(config)
    if not registry:
        print("A flats registry is required to create a collection history report.", file=sys.stderr)
        sys.exit(1)

    try:
        transactions = _load_statements(input_path, config)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    month_keys = month_keys_from_transactions(transactions)
    if input_path.is_dir():
        statement_files = find_statement_files(input_path)
        print(f"Loaded {len(statement_files)} statement file(s) from folder: {input_path}")
        print(f"Combining {len(month_keys)} month(s): {', '.join(month_keys)}")
    if input_path.is_file() and len(month_keys) not in SUPPORTED_PERIOD_LENGTHS:
        print(
            "This command requires transactions from 3, 6, or 12 calendar months; "
            f"found: {', '.join(month_keys) or 'none'}.",
            file=sys.stderr,
        )
        sys.exit(1)
    if not month_keys:
        print("No transaction months found in the provided statements.", file=sys.stderr)
        sys.exit(1)

    water_bills = registry.meta.get("water_bills_by_month")
    if not isinstance(water_bills, Mapping):
        print(
            "Add meta.water_bills_by_month to flats.yaml before creating this report.",
            file=sys.stderr,
        )
        sys.exit(1)

    flat_filter: str | None = None
    if args.flat:
        flat_filter = args.flat.strip().upper().replace("-", "").replace(" ", "")
        if flat_filter not in registry.flats:
            print(f"Flat not found in registry: {args.flat}", file=sys.stderr)
            sys.exit(1)

    classified = classify_transactions(transactions, config, registry)
    credits = split_by_type(classified)["credits"]
    try:
        transaction_history = build_flat_transaction_history(credits, registry)
        monthly_reconciliation = build_monthly_reconciliation(
            credits, registry, month_keys, water_bills
        )
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    if flat_filter:
        transaction_history = transaction_history[transaction_history["Flat"] == flat_filter]
        monthly_reconciliation = monthly_reconciliation[
            monthly_reconciliation["Flat"] == flat_filter
        ]

    period_summary = build_period_summary(monthly_reconciliation)
    filename = f"{month_keys[0]}_to_{month_keys[-1]}_collection_history.xlsx"
    if flat_filter:
        filename = f"{month_keys[0]}_to_{month_keys[-1]}_{flat_filter}_collection_history.xlsx"
    output_path = config.output_dir / filename
    write_collection_history_report(
        output_path,
        transaction_history,
        monthly_reconciliation,
        period_summary,
    )
    if flat_filter:
        print(f"Collection history report written for {flat_filter}: {output_path}")
    else:
        print(f"Collection history report written: {output_path}")


if __name__ == "__main__":
    main()
