"""Parse HDFC bank statement exports (XLS/XLSX/CSV)."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from nv_billbook.config import Config

TRANSACTION_DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2}$")
STATEMENT_PERIOD_RE = re.compile(
    r"Statement From\s*:\s*(\d{2}/\d{2}/\d{4})\s+To\s*:\s*(\d{2}/\d{2}/\d{4})",
    re.IGNORECASE,
)
HEADER_MARKERS = ("date", "narration", "withdrawal amt.", "deposit amt.")


@dataclass
class StatementMeta:
    source_file: Path
    period_start: date | None
    period_end: date | None
    account_no: str | None = None


@dataclass
class ParsedStatement:
    meta: StatementMeta
    transactions: pd.DataFrame


def _read_raw_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, header=None, dtype=str)
    if suffix == ".xls":
        return pd.read_excel(path, header=None, engine="xlrd", dtype=str)
    if suffix in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, header=None, engine="openpyxl", dtype=str)
    raise ValueError(f"Unsupported file type: {path.suffix}")


def _find_header_row(raw: pd.DataFrame) -> int:
    for row_idx in range(len(raw)):
        row_values = [
            str(value).strip().lower()
            for value in raw.iloc[row_idx].tolist()
            if pd.notna(value) and str(value).strip()
        ]
        if not row_values:
            continue
        joined = " | ".join(row_values)
        if all(marker in joined for marker in HEADER_MARKERS):
            return row_idx
    raise ValueError("Could not find transaction header row in statement")


def _extract_metadata(raw: pd.DataFrame, path: Path) -> StatementMeta:
    text_blob = "\n".join(
        str(value)
        for value in raw.fillna("").astype(str).values.flatten()
        if str(value).strip()
    )

    period_start = period_end = None
    match = STATEMENT_PERIOD_RE.search(text_blob)
    if match:
        period_start = datetime.strptime(match.group(1), "%d/%m/%Y").date()
        period_end = datetime.strptime(match.group(2), "%d/%m/%Y").date()

    account_no = None
    account_match = re.search(r"Account No\s*:(\d+)", text_blob, re.IGNORECASE)
    if account_match:
        account_no = account_match.group(1)

    return StatementMeta(
        source_file=path,
        period_start=period_start,
        period_end=period_end,
        account_no=account_no,
    )


def _to_number(value: object) -> float | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    cleaned = text.replace(",", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def _parse_date(value: object, fmt: str) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    if not text:
        return None
    for pattern in (fmt, "%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return None


def parse_hdfc_statement(path: Path, config: Config) -> ParsedStatement:
    """Parse a single HDFC statement file into normalized transactions."""
    raw = _read_raw_table(path)
    meta = _extract_metadata(raw, path)
    header_row = _find_header_row(raw)

    header_values = [
        str(value).strip() if pd.notna(value) else ""
        for value in raw.iloc[header_row].tolist()
    ]
    data = raw.iloc[header_row + 1 :].copy()
    data.columns = header_values
    data = data.loc[:, [column for column in data.columns if column]]

    col = config.columns
    date_col = col.get("date", "Date")
    desc_col = col.get("description", "Narration")
    debit_col = col.get("debit", "Withdrawal Amt.")
    credit_col = col.get("credit", "Deposit Amt.")
    balance_col = col.get("balance", "Closing Balance")
    ref_col = col.get("reference", "Chq./Ref.No.")
    value_date_col = col.get("value_date", "Value Dt")

    rows: list[dict[str, object]] = []
    for _, record in data.iterrows():
        date_text = str(record.get(date_col, "")).strip()
        if not TRANSACTION_DATE_RE.match(date_text):
            continue

        debit = _to_number(record.get(debit_col))
        credit = _to_number(record.get(credit_col))
        if debit is None and credit is None:
            continue

        rows.append(
            {
                "date": _parse_date(date_text, config.date_format),
                "date_str": date_text,
                "narration": str(record.get(desc_col, "")).strip(),
                "reference": str(record.get(ref_col, "")).strip()
                if ref_col in record.index
                else "",
                "value_date": _parse_date(record.get(value_date_col), config.date_format)
                if value_date_col in record.index
                else None,
                "debit": debit or 0.0,
                "credit": credit or 0.0,
                "balance": _to_number(record.get(balance_col)),
            }
        )

    transactions = pd.DataFrame(rows)
    if transactions.empty:
        raise ValueError(f"No transactions found in {path.name}")

    transactions["amount"] = transactions["credit"] - transactions["debit"]
    transactions["source_file"] = path.name
    return ParsedStatement(meta=meta, transactions=transactions)


def infer_report_month(meta: StatementMeta, transactions: pd.DataFrame) -> str:
    """Return YYYY-MM for the report filename."""
    if meta.period_start:
        return meta.period_start.strftime("%Y-%m")
    first_date = transactions["date"].dropna().min()
    if isinstance(first_date, date):
        return first_date.strftime("%Y-%m")
    raise ValueError("Could not determine report month from statement")


def find_statement_files(input_dir: Path) -> list[Path]:
    patterns = ("*.xls", "*.xlsx", "*.xlsm", "*.csv")
    files: list[Path] = []
    for pattern in patterns:
        files.extend(sorted(input_dir.glob(pattern)))
    return files
