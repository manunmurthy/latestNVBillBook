"""CLI entry point for generating monthly expense reports."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from nv_billbook.classifier import classify_transactions, split_by_type
from nv_billbook.config import Config
from nv_billbook.parser import (
    find_statement_files,
    infer_report_month,
    parse_hdfc_statement,
)
from nv_billbook.reporter import write_monthly_report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate Nava Vaibhva monthly expense Excel reports from HDFC bank statements."
    )
    parser.add_argument(
        "--month",
        help="Report month in YYYY-MM format (e.g. 2026-06)",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Process all statement files in the input directory",
    )
    parser.add_argument(
        "--input",
        help="Statement file or folder (overrides config input_dir)",
    )
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to config file (default: config.yaml)",
    )
    return parser.parse_args()


def _resolve_input_paths(config: Config, input_arg: str | None) -> list[Path]:
    if input_arg:
        path = Path(input_arg)
        if path.is_file():
            return [path]
        if path.is_dir():
            return find_statement_files(path)
        raise FileNotFoundError(f"Input path not found: {path}")

    if not config.input_dir.exists():
        config.input_dir.mkdir(parents=True, exist_ok=True)
    return find_statement_files(config.input_dir)


def _process_statement(path: Path, config: Config, month_filter: str | None) -> Path | None:
    print(f"Processing: {path.name}")
    parsed = parse_hdfc_statement(path, config)
    report_month = infer_report_month(parsed.meta, parsed.transactions)

    if month_filter and report_month != month_filter:
        print(f"  Skipped (statement month {report_month} != requested {month_filter})")
        return None

    classified = classify_transactions(parsed.transactions, config)
    split = split_by_type(classified)
    output_path = write_monthly_report(parsed, split, config, report_month)

    credits = split["credits"]
    debits = split["debits"]
    print(
        f"  Month: {report_month} | "
        f"Credits: {len(credits)} (₹{credits['credit'].sum():,.2f}) | "
        f"Debits: {len(debits)} (₹{debits['debit'].sum():,.2f})"
    )
    print(f"  Report written: {output_path}")
    return output_path


def main() -> None:
    args = parse_args()
    config_path = Path(args.config)

    if not config_path.exists():
        print(
            f"Config not found: {config_path}\n"
            "Copy config.example.yaml to config.yaml and adjust if needed.",
            file=sys.stderr,
        )
        sys.exit(1)

    if not args.month and not args.all and not args.input:
        print("Specify --month YYYY-MM, --all, or --input <file/folder>", file=sys.stderr)
        sys.exit(1)

    config = Config.load(config_path)
    month_filter = args.month if not args.all else None

    try:
        input_paths = _resolve_input_paths(config, args.input)
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    if not input_paths:
        print("No statement files found.", file=sys.stderr)
        sys.exit(1)

    generated: list[Path] = []
    for path in input_paths:
        try:
            output = _process_statement(path, config, month_filter)
            if output:
                generated.append(output)
        except Exception as exc:
            print(f"Failed to process {path.name}: {exc}", file=sys.stderr)

    if not generated:
        print("No reports generated.", file=sys.stderr)
        sys.exit(1)

    print(f"\nDone. {len(generated)} report(s) created in {config.output_dir}/")


if __name__ == "__main__":
    main()
