"""Write standalone multi-month flat collection history workbooks."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


TITLE_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
HEADER_FONT = Font(bold=True)
MONEY_HEADERS = {
    "Amount",
    "Expected Maintenance",
    "Water Bill",
    "Expected",
    "Paid",
    "Amount Due",
    "Extra Paid",
    "Total Expected",
    "Total Paid",
    "Total Amount Due",
    "Total Extra Paid",
}


def _write_table(
    writer: pd.ExcelWriter,
    sheet_name: str,
    frame: pd.DataFrame,
    start_row: int,
) -> int:
    """Write and style one table, returning the next available row index."""
    frame.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
    worksheet = writer.sheets[sheet_name]
    header_row = start_row + 1
    for cell in worksheet[header_row]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

    for row in worksheet.iter_rows(
        min_row=header_row + 1,
        max_row=header_row + len(frame),
        max_col=len(frame.columns),
    ):
        for cell in row:
            header = worksheet.cell(row=header_row, column=cell.column).value
            if header in MONEY_HEADERS and isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0.00"

    return start_row + len(frame) + 3


def _set_column_widths(worksheet) -> None:
    for column_cells in worksheet.iter_cols():
        values = [str(cell.value) for cell in column_cells if cell.value is not None]
        if values:
            worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = min(
                max(len(value) for value in values) + 2, 30
            )


def write_collection_history_report(
    output_path: Path,
    transaction_history: pd.DataFrame,
    monthly_reconciliation: pd.DataFrame,
    period_summary: pd.DataFrame,
) -> Path:
    """Write one worksheet per flat with its summary, monthly rows, and payments."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    flat_order = monthly_reconciliation["Flat"].drop_duplicates().tolist()

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for flat_no in flat_order:
            sheet_name = str(flat_no)
            flat_summary = period_summary[period_summary["Flat"] == flat_no]
            flat_monthly = monthly_reconciliation[monthly_reconciliation["Flat"] == flat_no]
            flat_transactions = transaction_history[transaction_history["Flat"] == flat_no]

            next_row = _write_table(writer, sheet_name, flat_summary, start_row=2)
            worksheet = writer.sheets[sheet_name]
            worksheet["A1"] = f"Flat {flat_no} — Collection History"
            worksheet["A1"].font = TITLE_FONT
            worksheet["A1"].fill = TITLE_FILL

            worksheet.cell(row=next_row, column=1).value = "Monthly Reconciliation"
            worksheet.cell(row=next_row, column=1).font = HEADER_FONT
            next_row = _write_table(writer, sheet_name, flat_monthly, start_row=next_row)

            worksheet.cell(row=next_row, column=1).value = "Payment Transactions"
            worksheet.cell(row=next_row, column=1).font = HEADER_FONT
            _write_table(writer, sheet_name, flat_transactions, start_row=next_row)

            worksheet.freeze_panes = "A3"
            _set_column_widths(worksheet)

    return output_path
