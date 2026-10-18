"""CLI entry point for generating monthly expense reports."""

from __future__ import annotations

import argparse
import sys
from dataclasses import replace
from pathlib import Path

import pandas as pd

from nv_billbook.classifier import classify_transactions, split_by_type
from nv_billbook.config import Config
from nv_billbook.flats_registry import FlatsRegistry
from nv_billbook.parser import (
    INPUT_FORMATS,
    ParsedStatement,
    find_statement_files,
    filter_transactions_by_months,
    infer_report_month,
    is_supported_extension,
    normalize_input_format,
    parse_month_list,
    parse_hdfc_statement,
    deduplicate_transactions,
)
from nv_billbook.reconciliation import build_maintenance_reconciliation
from nv_billbook.reporter import write_monthly_report
from nv_billbook.vendor_reporter import write_vendor_report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate Nava Vaibhva monthly expense Excel reports from HDFC bank statements."
    )
    month_group = parser.add_mutually_exclusive_group()
    month_group.add_argument(
        "--month",
        help="Report month in YYYY-MM format (e.g. 2026-06)",
    )
    month_group.add_argument(
        "--months",
        help="Comma-separated months to extract from a longer statement (e.g. 2026-04,2026-07)",
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
    parser.add_argument(
        "--water-bill",
        type=float,
        help="Fixed per-flat monthly water bill for this report (overrides flats.yaml)",
    )
    parser.add_argument(
        "--input-format",
        choices=INPUT_FORMATS,
        default="auto",
        help="Read Excel/CSV, PDF, both, or detect automatically (default: auto)",
    )
    parser.add_argument(
        "--combine",
        action="store_true",
        help="Combine all selected statement files into one consolidated workbook",
    )
    parser.add_argument(
        "--vendor-report",
        action="store_true",
        help="Create a vendor summary and vendor payment history workbook",
    )
    parser.add_argument(
        "--vendor",
        help="Optional vendor/payee filter; accepts a comma-separated list",
    )
    return parser.parse_args()


def _resolve_input_paths(
    config: Config,
    input_arg: str | None,
    input_format: str = "auto",
) -> list[Path]:
    input_format = normalize_input_format(input_format)
    if input_arg:
        path = Path(input_arg)
        if path.is_file():
            if not is_supported_extension(path, input_format):
                raise ValueError(
                    f"{path.name} is not allowed with --input-format {input_format}."
                )
            return [path]
        if path.is_dir():
            return find_statement_files(path, input_format)
        raise FileNotFoundError(f"Input path not found: {path}")

    if not config.input_dir.exists():
        config.input_dir.mkdir(parents=True, exist_ok=True)
    return find_statement_files(config.input_dir, input_format)


def _load_flats_registry(config: Config) -> FlatsRegistry | None:
    if not config.flats_registry_path:
        return None
    path = config.flats_registry_path
    if not path.is_absolute():
        path = Path.cwd() / path
    if not path.exists():
        print(f"Warning: flats registry not found at {path} — continuing without it")
        return None

    dim_path = config.flat_dimensions_path
    if not dim_path.is_absolute():
        dim_path = Path.cwd() / dim_path

    registry = FlatsRegistry.load(path, dimensions_path=dim_path if dim_path.exists() else None)
    print(f"Loaded flats registry: {len(registry.flats)} flats")
    return registry


def _generate_reports(
    parsed: ParsedStatement,
    config: Config,
    month_filter: list[str] | None,
    flats_registry: FlatsRegistry | None,
    water_bill_per_month: float | None,
    combined: bool = False,
) -> list[Path]:
    if combined:
        available_months = sorted(
            {
                value.strftime("%Y-%m")
                for value in parsed.transactions["date"].dropna()
            }
        )
        selected_months = month_filter or available_months
        missing_months = [month for month in selected_months if month not in available_months]
        if missing_months:
            raise ValueError(
                "Requested month(s) not found in the combined statements: "
                + ", ".join(missing_months)
            )
        filtered_transactions = filter_transactions_by_months(
            parsed.transactions,
            selected_months,
        )
        if filtered_transactions.empty:
            return []
        dates = filtered_transactions["date"].dropna()
        filtered_meta = replace(
            parsed.meta,
            period_start=min(dates) if not dates.empty else parsed.meta.period_start,
            period_end=max(dates) if not dates.empty else parsed.meta.period_end,
        )
        report_month = f"{selected_months[0]}_to_{selected_months[-1]}"
        parsed_items = [(report_month, filtered_meta, filtered_transactions)]
    elif not args.vendor_report:
        selected_months = month_filter or [infer_report_month(parsed.meta, parsed.transactions)]
        parsed_items = []
        for report_month in selected_months:
            filtered_transactions = filter_transactions_by_months(
                parsed.transactions,
                [report_month],
            )
            if filtered_transactions.empty:
                print(f"  Skipped (no transactions found for {report_month})")
                continue
            dates = filtered_transactions["date"].dropna()
            filtered_meta = parsed.meta
            if not dates.empty:
                filtered_meta = replace(
                    parsed.meta,
                    period_start=min(dates),
                    period_end=max(dates),
                )
            parsed_items.append((report_month, filtered_meta, filtered_transactions))

    generated: list[Path] = []
    for report_month, filtered_meta, filtered_transactions in parsed_items:
        filtered_parsed = ParsedStatement(filtered_meta, filtered_transactions)
        classified = classify_transactions(filtered_transactions, config, flats_registry)
        split = split_by_type(classified)
        output_path = write_monthly_report(
            filtered_parsed,
            split,
            config,
            report_month,
            flats_registry=flats_registry,
            water_bill_per_month=water_bill_per_month,
            period_months=len(selected_months) if combined else 1,
            report_label=(
                f"{selected_months[0]} to {selected_months[-1]}"
                if combined
                else None
            ),
            combined_report=combined,
        )

        credits = split["credits"]
        debits = split["debits"]
        print(
            f"  Month: {report_month} | "
            f"Credits: {len(credits)} (₹{credits['credit'].sum():,.2f}) | "
            f"Debits: {len(debits)} (₹{debits['debit'].sum():,.2f})"
        )
        if flats_registry:
            identified_flats = {
                info.flat_no
                for flat_no in credits.get("flat_no", [])
                if isinstance(flat_no, str)
                for info in [flats_registry.get(flat_no)]
                if info
            }
            reconciliation = build_maintenance_reconciliation(
                credits,
                flats_registry,
                water_bill_per_month=water_bill_per_month,
                period_months=len(selected_months) if combined else 1,
            )
            status_counts = reconciliation["Status"].value_counts()
            print(
                f"  Flat identification: {len(identified_flats)}/{len(flats_registry.flats)} "
                "flats have at least one identified credit"
            )
            print(
                "  Maintenance status: "
                + " | ".join(
                    f"{status} {int(status_counts.get(status, 0))}"
                    for status in ("Paid", "Short", "Excess", "Pending")
                )
            )
        print(f"  Report written: {output_path}")
        generated.append(output_path)
    return generated


def _process_statement(
    path: Path,
    config: Config,
    month_filter: list[str] | None,
    flats_registry: FlatsRegistry | None,
    water_bill_per_month: float | None,
) -> list[Path]:
    print(f"Processing: {path.name}")
    return _generate_reports(
        parse_hdfc_statement(path, config),
        config,
        month_filter,
        flats_registry,
        water_bill_per_month,
    )


def _combine_statements(paths: list[Path], config: Config) -> ParsedStatement:
    """Parse and combine multiple statement files into one normalized statement."""
    parsed_statements = []
    for path in paths:
        try:
            parsed_statements.append(parse_hdfc_statement(path, config))
        except Exception as exc:
            print(f"Skipping {path.name}: {exc}", file=sys.stderr)
    if not parsed_statements:
        raise ValueError("No valid statement files found to combine.")
    transactions = deduplicate_transactions(
        pd.concat(
            [parsed.transactions for parsed in parsed_statements],
            ignore_index=True,
        )
    ).sort_values(
        by=["date", "source_file", "date_str"],
        kind="stable",
        ignore_index=True,
    )
    dates = transactions["date"].dropna()
    first_meta = parsed_statements[0].meta
    meta = replace(
        first_meta,
        source_file=Path("Combined statements"),
        period_start=min(dates) if not dates.empty else first_meta.period_start,
        period_end=max(dates) if not dates.empty else first_meta.period_end,
    )
    return ParsedStatement(meta=meta, transactions=transactions)


def _generate_vendor_report(
    paths: list[Path],
    config: Config,
    month_filter: list[str] | None,
    vendor_filter: str | None,
) -> Path:
    parsed = _combine_statements(paths, config)
    available_months = sorted(
        {value.strftime("%Y-%m") for value in parsed.transactions["date"].dropna()}
    )
    selected_months = month_filter or available_months
    missing_months = [month for month in selected_months if month not in available_months]
    if missing_months:
        raise ValueError(
            "Requested month(s) not found in the statements: " + ", ".join(missing_months)
        )
    transactions = filter_transactions_by_months(parsed.transactions, selected_months)
    classified = classify_transactions(transactions, config)
    debits = split_by_type(classified)["debits"]
    suffix = "filtered" if vendor_filter else "all_vendors"
    output_path = config.output_dir / (
        f"{selected_months[0]}_to_{selected_months[-1]}_{suffix}_payment_history.xlsx"
    )
    return write_vendor_report(
        output_path,
        debits,
        f"{selected_months[0]} to {selected_months[-1]}",
        vendor_filter=vendor_filter,
    )


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

    if not args.month and not args.months and not args.all and not args.input:
        print("Specify --month/--months, --all, or --input <file/folder>", file=sys.stderr)
        sys.exit(1)

    config = Config.load(config_path)
    flats_registry = _load_flats_registry(config)
    try:
        month_filter = parse_month_list(args.months or args.month)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    try:
        input_paths = _resolve_input_paths(config, args.input, args.input_format)
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    if not input_paths:
        print("No statement files found.", file=sys.stderr)
        sys.exit(1)

    generated: list[Path] = []
    if args.vendor_report:
        try:
            vendor_output = _generate_vendor_report(
                input_paths,
                config,
                month_filter,
                args.vendor,
            )
            generated.append(vendor_output)
            print(f"Vendor report written: {vendor_output}")
        except Exception as exc:
            print(f"Failed to generate vendor report: {exc}", file=sys.stderr)

    if args.combine:
        try:
            combined = _combine_statements(input_paths, config)
            generated.extend(
                _generate_reports(
                    combined,
                    config,
                    month_filter,
                    flats_registry,
                    args.water_bill,
                    combined=True,
                )
            )
        except Exception as exc:
            print(f"Failed to combine statements: {exc}", file=sys.stderr)
    else:
        for path in input_paths:
            try:
                outputs = _process_statement(
                    path,
                    config,
                    month_filter,
                    flats_registry,
                    args.water_bill,
                )
                generated.extend(outputs)
            except Exception as exc:
                print(f"Failed to process {path.name}: {exc}", file=sys.stderr)

    if not generated:
        print("No reports generated.", file=sys.stderr)
        sys.exit(1)

    print(f"\nDone. {len(generated)} report(s) created in {config.output_dir}/")


if __name__ == "__main__":
    main()
