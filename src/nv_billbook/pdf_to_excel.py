"""Convert HDFC PDF statements into a simple Excel transaction sheet."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
import pdfplumber


DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2,4}")
TRANSACTION_RE = re.compile(
    r"^(?P<date>\d{2}/\d{2}/\d{2,4})\s+"
    r"(?P<body>.*?)\s+"
    r"(?P<ref>\d{10,})\s+"
    r"(?P<value_date>\d{2}/\d{2}/\d{2,4})\s+"
    r"(?P<amount>[\d,]+\.\d{2})\s+"
    r"(?P<closing>[\d,]+\.\d{2})(?:\s+.*)?$"
)
HEADER_MARKERS = {"date", "narration", "chq./ref.no.", "valuedt", "withdrawalamt.", "depositamt."}


def _clean_text(value: object) -> str:
    text = "" if value is None else str(value).strip()
    return "" if text.lower() == "nan" else text


def _to_number(value: str | None) -> float | None:
    if not value:
        return None
    return float(value.replace(",", ""))


def _normalize_lines(text: str) -> list[str]:
    lines = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        line = re.sub(r"\s+", " ", line)
        lines.append(line)
    return lines


def _extract_transaction_lines(lines: list[str]) -> list[str]:
    transactions: list[str] = []
    current: list[str] = []

    for line in lines:
        lower = line.lower()
        if lower.startswith("page no."):
            continue
        if lower.startswith("from :") or lower.startswith("accountbranch :"):
            continue
        if HEADER_MARKERS.issubset({part.lower() for part in line.split()}):
            continue

        if DATE_RE.match(line):
            if current:
                transactions.append(" ".join(current))
            current = [line]
        else:
            if current:
                current.append(line)

    if current:
        transactions.append(" ".join(current))
    return transactions


def extract_hdfc_pdf_transactions(pdf_path: Path) -> pd.DataFrame:
    rows: list[dict[str, object]] = []
    previous_closing: float | None = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = _normalize_lines(text)
            txns = _extract_transaction_lines(lines)
            for txn in txns:
                match = TRANSACTION_RE.match(txn)
                if not match:
                    continue
                groups = match.groupdict()
                narration = groups["body"].strip()
                amount = _to_number(groups["amount"])
                closing = _to_number(groups["closing"])
                txn_type = ""
                if previous_closing is not None and closing is not None and amount is not None:
                    if closing > previous_closing:
                        txn_type = "Credit"
                    elif closing < previous_closing:
                        txn_type = "Debit"
                previous_closing = closing if closing is not None else previous_closing
                rows.append(
                    {
                        "Date": groups["date"],
                        "Narration": narration,
                        "Reference": groups["ref"],
                        "Value Date": groups["value_date"],
                        "Amount": amount,
                        "Transaction Type": txn_type,
                        "Closing Balance": closing,
                    }
                )

    if not rows:
        raise ValueError(f"No transactions found in {pdf_path.name}")

    df = pd.DataFrame(rows)
    return df


def convert_pdf_to_excel(pdf_path: Path, output_path: Path | None = None) -> Path:
    df = extract_hdfc_pdf_transactions(pdf_path)
    if output_path is None:
        output_path = pdf_path.with_suffix(".xlsx")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Transactions", index=False)
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert an HDFC PDF statement into Excel.")
    parser.add_argument("--input", required=True, help="Path to the PDF statement")
    parser.add_argument("--output", help="Optional output .xlsx path")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    try:
        written = convert_pdf_to_excel(input_path, Path(args.output) if args.output else None)
    except Exception as exc:
        print(f"Failed to convert PDF: {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Excel file written: {written}")


if __name__ == "__main__":
    main()
