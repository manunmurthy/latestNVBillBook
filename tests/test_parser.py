"""Tests for HDFC statement parsing."""

from pathlib import Path

import pytest

from nv_billbook.classifier import classify_transactions, split_by_type
from nv_billbook.config import Config
from nv_billbook.parser import (
    find_statement_files,
    is_supported_extension,
    normalize_input_format,
    parse_hdfc_statement,
)

JUNE_STATEMENT = Path(
    "/Users/I071733/Downloads/NVBankStatements/Acct_Statement_June2026.xls"
)


@pytest.fixture
def config() -> Config:
    root = Path(__file__).resolve().parents[1]
    return Config.load(root / "config.yaml")


@pytest.mark.skipif(not JUNE_STATEMENT.exists(), reason="Sample statement not available")
def test_parse_june_statement(config: Config) -> None:
    parsed = parse_hdfc_statement(JUNE_STATEMENT, config)
    assert len(parsed.transactions) == 218
    assert parsed.meta.period_start is not None
    assert parsed.meta.period_start.month == 6


@pytest.mark.skipif(not JUNE_STATEMENT.exists(), reason="Sample statement not available")
def test_classify_june_statement(config: Config) -> None:
    parsed = parse_hdfc_statement(JUNE_STATEMENT, config)
    classified = classify_transactions(parsed.transactions, config)
    split = split_by_type(classified)

    assert len(split["credits"]) == 194
    assert len(split["debits"]) >= 20
    assert split["credits"]["credit"].sum() > 0
    assert split["debits"]["debit"].sum() > 0


def test_input_format_selection() -> None:
    assert normalize_input_format(None) == "auto"
    assert normalize_input_format("PDF") == "pdf"
    assert is_supported_extension(Path("statement.pdf"), "both")
    assert not is_supported_extension(Path("statement.pdf"), "excel")
    assert is_supported_extension(Path("statement.xls"), "excel")


def test_find_statement_files_supports_pdf_and_excel_modes(tmp_path) -> None:
    for name in ("a.xls", "b.xlsx", "c.pdf", "notes.txt"):
        (tmp_path / name).write_text("", encoding="utf-8")

    assert [path.name for path in find_statement_files(tmp_path, "pdf")] == ["c.pdf"]
    assert [path.name for path in find_statement_files(tmp_path, "excel")] == [
        "a.xls",
        "b.xlsx",
    ]
    assert [path.name for path in find_statement_files(tmp_path, "both")] == [
        "a.xls",
        "b.xlsx",
        "c.pdf",
    ]
