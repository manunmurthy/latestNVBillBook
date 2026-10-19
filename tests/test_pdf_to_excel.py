"""Regression tests for HDFC PDF transaction column parsing."""

from nv_billbook.pdf_to_excel import TRANSACTION_RE, _extract_table_rows


def test_pdf_transaction_regex_keeps_withdrawal_deposit_and_closing_balance_separate() -> None:
    row = (
        "12/02/2024 UPI-MANU N-MANUNMURTHY@OKAXIS-HDFC0001036-404343829586-UPI "
        "404343829586 12/02/2024 0.00 2,920.00 976,055.50"
    )

    match = TRANSACTION_RE.match(row)

    assert match is not None
    assert match.group("withdrawal") == "0.00"
    assert match.group("deposit") == "2,920.00"
    assert match.group("closing") == "976,055.50"


class _TablePage:
    def extract_tables(self):
        return [
            [
                [
                    "12/02/2024",
                    "UPI-MANU N-MANUNMURTHY@OKAXIS-HDFC0001036-404343829586-UPI",
                    "404343829586",
                    "12/02/2024",
                    "0.00",
                    "2,920.00",
                    "9,76,055.50",
                ]
            ]
        ]


def test_pdf_table_parser_reads_headerless_continuation_pages() -> None:
    rows = _extract_table_rows(_TablePage())

    assert len(rows) == 1
    assert rows[0]["Transaction Type"] == "Credit"
    assert rows[0]["Amount"] == 2920.0
    assert rows[0]["Narration"].startswith("UPI-MANU N")
