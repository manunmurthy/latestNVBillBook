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
from nv_billbook.parser import (
    INPUT_FORMATS,
    find_statement_files,
    is_supported_extension,
    normalize_input_format,
    parse_hdfc_statement,
)


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
        help=(
            "Optional comma-separated flat numbers to include (for example "
            "A301,A302,A303)"
        ),
    )
    parser.add_argument("--config", default="config.yaml", help="Path to config file")
    parser.add_argument(
        "--input-format",
        choices=INPUT_FORMATS,
        default="auto",
        help="Read Excel/CSV, PDF, both, or detect automatically (default: auto)",
    )
    return parser.parse_args()


def _parse_flat_filter(value: str, registry) -> list[str]:
    """Normalize and validate a comma-separated flat selection."""
    flat_numbers = [
        item.strip().upper().replace("-", "").replace(" ", "")
        for item in value.split(",")
        if item.strip()
    ]
    if not flat_numbers:
        raise ValueError("At least one flat number is required when using --flat.")

    # Preserve the user's order while avoiding duplicate worksheets.
    selected = list(dict.fromkeys(flat_numbers))
    missing = [flat_no for flat_no in selected if flat_no not in registry.flats]
    if missing:
        raise ValueError(f"Flat not found in registry: {', '.join(missing)}")
    return selected


def _deduplicate_transactions(transactions: pd.DataFrame) -> pd.DataFrame:
    """Remove duplicate rows when the same statement is supplied as PDF and Excel."""
    if transactions.empty:
        return transactions
    dedupe = transactions.copy()
    dedupe["_narration_key"] = (
        dedupe["narration"].fillna("").astype(str).str.upper().str.replace(r"\s+", "", regex=True)
    )
    key_columns = [
        "date",
        "_narration_key",
        "reference",
        "debit",
        "credit",
        "balance",
    ]
    return dedupe.drop_duplicates(subset=key_columns, keep="first").drop(
        columns=["_narration_key"]
    )


def _load_statements(
    input_path: Path,
    config: Config,
    input_format: str = "auto",
) -> pd.DataFrame:
    input_format = normalize_input_format(input_format)
    if input_path.is_file():
        if not is_supported_extension(input_path, input_format):
            raise ValueError(
                f"{input_path.name} is not allowed with --input-format {input_format}."
            )
        parsed = parse_hdfc_statement(input_path, config)
        return parsed.transactions

    if input_path.is_dir():
        files = sorted(find_statement_files(input_path, input_format))
        if not files:
            raise FileNotFoundError(f"No statement files found in folder: {input_path}")

        frames: list[pd.DataFrame] = []
        for file_path in files:
            try:
                parsed = parse_hdfc_statement(file_path, config)
            except Exception as exc:
                print(f"Skipping {file_path.name}: {exc}", file=sys.stderr)
                continue
            frames.append(parsed.transactions)

        if not frames:
            raise ValueError(f"No valid statement files found in folder: {input_path}")

        combined = pd.concat(frames, ignore_index=True)
        combined = _deduplicate_transactions(combined)
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
        transactions = _load_statements(input_path, config, args.input_format)
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    month_keys = month_keys_from_transactions(transactions)
    if input_path.is_dir():
        statement_files = find_statement_files(input_path, args.input_format)
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

    flat_filter: list[str] | None = None
    if args.flat:
        try:
            flat_filter = _parse_flat_filter(args.flat, registry)
        except ValueError as exc:
            print(str(exc), file=sys.stderr)
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
        transaction_history = transaction_history[
            transaction_history["Flat"].isin(flat_filter)
        ]
        monthly_reconciliation = monthly_reconciliation[
            monthly_reconciliation["Flat"].isin(flat_filter)
        ]

    period_summary = build_period_summary(monthly_reconciliation)
    filename = f"{month_keys[0]}_to_{month_keys[-1]}_collection_history.xlsx"
    if flat_filter:
        if len(flat_filter) == 1:
            selection_name = flat_filter[0]
        else:
            selection_name = "selected"
        filename = (
            f"{month_keys[0]}_to_{month_keys[-1]}_"
            f"{selection_name}_collection_history.xlsx"
        )
    output_path = config.output_dir / filename
    write_collection_history_report(
        output_path,
        transaction_history,
        monthly_reconciliation,
        period_summary,
    )
    if flat_filter:
        print(
            "Collection history report written for "
            f"{', '.join(flat_filter)}: {output_path}"
        )
    else:
        print(f"Collection history report written: {output_path}")


if __name__ == "__main__":
    main()
