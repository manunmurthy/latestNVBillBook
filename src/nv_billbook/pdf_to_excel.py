"""Convert HDFC PDF statements into a simple Excel transaction sheet."""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber


DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2,4}")
TRANSACTION_RE = re.compile(
    r"^(?P<date>\d{2}/\d{2}/\d{2,4})\s+"
    r"(?P<body>.*?)\s+"
    r"(?P<ref>\d{10,})\s+"
    r"(?P<value_date>\d{2}/\d{2}/\d{2,4})\s+"
    r"(?P<withdrawal>[\d,]+\.\d{2})\s+"
    r"(?P<deposit>[\d,]+\.\d{2})\s+"
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


def _first_transaction_type(page) -> str:
    """Read the first row's debit/credit side when no prior balance exists."""
    try:
        tables = page.extract_tables()
    except Exception:
        return ""
    for table in tables:
        if not table:
            continue
        header = [_clean_text(value).lower().replace(" ", "") for value in table[0]]
        if "withdrawalamt." not in header or "depositamt." not in header:
            continue
        withdrawal_index = header.index("withdrawalamt.")
        deposit_index = header.index("depositamt.")
        for row in table[1:]:
            if not row or not _clean_text(row[0]):
                continue
            withdrawal = _clean_text(row[withdrawal_index]) if withdrawal_index < len(row) else ""
            deposit = _clean_text(row[deposit_index]) if deposit_index < len(row) else ""
            if withdrawal and not deposit:
                return "Debit"
            if deposit and not withdrawal:
                return "Credit"
    return ""


def _header_key(value: object) -> str:
    return re.sub(r"[^a-z]", "", _clean_text(value).lower())


def _extract_table_rows(page) -> list[dict[str, object]]:
    """Extract transaction rows from the PDF table without column interleaving."""
    rows: list[dict[str, object]] = []
    try:
        tables = page.extract_tables()
    except Exception:
        return rows

    for table in tables:
        if not table:
            continue
        header_index = None
        indexes: dict[str, int] = {}
        for row_index, header_row in enumerate(table[:3]):
            keys = [_header_key(value) for value in header_row]
            candidates = {
                "date": next((i for i, key in enumerate(keys) if key == "date"), None),
                "narration": next((i for i, key in enumerate(keys) if key == "narration"), None),
                "reference": next(
                    (i for i, key in enumerate(keys) if "chqrefno" in key or key == "reference"),
                    None,
                ),
                "value_date": next((i for i, key in enumerate(keys) if key == "valuedate"), None),
                "withdrawal": next((i for i, key in enumerate(keys) if "withdrawal" in key), None),
                "deposit": next((i for i, key in enumerate(keys) if "deposit" in key), None),
                "closing": next((i for i, key in enumerate(keys) if "closingbalance" in key), None),
            }
            if all(value is not None for value in candidates.values()):
                header_index = row_index
                indexes = {key: int(value) for key, value in candidates.items()}
                break
        if header_index is None:
            # HDFC prints headers only on the first page. Subsequent tables
            # keep the same seven-column layout and begin directly with data.
            first_row = table[0]
            if len(first_row) >= 7 and DATE_RE.match(_clean_text(first_row[0])):
                header_index = -1
                indexes = {
                    "date": 0,
                    "narration": 1,
                    "reference": 2,
                    "value_date": 3,
                    "withdrawal": 4,
                    "deposit": 5,
                    "closing": 6,
                }
            else:
                continue

        for row in table[header_index + 1 :]:
            if not row or indexes["date"] >= len(row):
                continue
            date_text = _clean_text(row[indexes["date"]])
            if not DATE_RE.match(date_text):
                continue
            narration = re.sub(r"\s+", " ", _clean_text(row[indexes["narration"]]))
            withdrawal = _to_number(_clean_text(row[indexes["withdrawal"]])) or 0.0
            deposit = _to_number(_clean_text(row[indexes["deposit"]])) or 0.0
            closing = _to_number(_clean_text(row[indexes["closing"]]))
            if deposit > 0 and withdrawal == 0:
                transaction_type = "Credit"
                amount = deposit
            elif withdrawal > 0 and deposit == 0:
                transaction_type = "Debit"
                amount = withdrawal
            else:
                transaction_type = ""
                amount = deposit or withdrawal
            rows.append(
                {
                    "Date": date_text,
                    "Narration": narration,
                    "Reference": _clean_text(row[indexes["reference"]]),
                    "Value Date": _clean_text(row[indexes["value_date"]]),
                    "Withdrawal Amount": withdrawal,
                    "Deposit Amount": deposit,
                    "Amount": amount,
                    "Transaction Type": transaction_type,
                    "Closing Balance": closing,
                }
            )

        if rows:
            return rows
    return rows


def extract_hdfc_pdf_metadata(pdf_path: Path) -> dict[str, object]:
    """Extract statement period and account number from an HDFC PDF header."""
    period_start = period_end = None
    account_no = None
    with pdfplumber.open(pdf_path) as pdf:
        header_text = "\n".join((page.extract_text() or "") for page in pdf.pages[:2])

    period_match = re.search(
        r"From\s*:\s*(\d{2}/\d{2}/\d{4})\s+To\s*:\s*(\d{2}/\d{2}/\d{4})",
        header_text,
        re.IGNORECASE,
    )
    if period_match:
        period_start = datetime.strptime(period_match.group(1), "%d/%m/%Y").date()
        period_end = datetime.strptime(period_match.group(2), "%d/%m/%Y").date()

    account_match = re.search(r"AccountNo\s*:\s*(\d+)", header_text, re.IGNORECASE)
    if account_match:
        account_no = account_match.group(1)

    return {
        "period_start": period_start,
        "period_end": period_end,
        "account_no": account_no,
    }


def extract_hdfc_pdf_transactions(
    pdf_path: Path,
    max_pages: int | None = None,
) -> pd.DataFrame:
    """Extract transactions, optionally limiting extraction to the first pages."""
    if max_pages is not None and max_pages < 1:
        raise ValueError("max_pages must be at least 1")

    rows: list[dict[str, object]] = []

    with pdfplumber.open(pdf_path) as pdf:
        pages = pdf.pages[:max_pages] if max_pages else pdf.pages
        for page in pages:
            table_rows = _extract_table_rows(page)
            if table_rows:
                rows.extend(table_rows)
                continue

            text = page.extract_text() or ""
            lines = _normalize_lines(text)
            txns = _extract_transaction_lines(lines)
            for txn_index, txn in enumerate(txns):
                match = TRANSACTION_RE.match(txn)
                if not match:
                    continue
                groups = match.groupdict()
                narration = groups["body"].strip()
                withdrawal = _to_number(groups["withdrawal"]) or 0.0
                deposit = _to_number(groups["deposit"]) or 0.0
                closing = _to_number(groups["closing"])
                if deposit > 0 and withdrawal == 0:
                    txn_type = "Credit"
                    amount = deposit
                elif withdrawal > 0 and deposit == 0:
                    txn_type = "Debit"
                    amount = withdrawal
                else:
                    txn_type = ""
                    amount = deposit or withdrawal
                rows.append(
                    {
                        "Date": groups["date"],
                        "Narration": narration,
                        "Reference": groups["ref"],
                        "Value Date": groups["value_date"],
                        "Withdrawal Amount": withdrawal,
                        "Deposit Amount": deposit,
                        "Amount": amount,
                        "Transaction Type": txn_type,
                        "Closing Balance": closing,
                    }
                )

    if not rows:
        raise ValueError(f"No transactions found in {pdf_path.name}")

    df = pd.DataFrame(rows)
    return df


def convert_pdf_to_excel(
    pdf_path: Path,
    output_path: Path | None = None,
    max_pages: int | None = None,
) -> Path:
    df = extract_hdfc_pdf_transactions(pdf_path, max_pages=max_pages)
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
    parser.add_argument(
        "--max-pages",
        type=int,
        help="Optional limit for debugging; read only the first N PDF pages",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    try:
        written = convert_pdf_to_excel(
            input_path,
            Path(args.output) if args.output else None,
            max_pages=args.max_pages,
        )
    except Exception as exc:
        print(f"Failed to convert PDF: {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Excel file written: {written}")


if __name__ == "__main__":
    main()
