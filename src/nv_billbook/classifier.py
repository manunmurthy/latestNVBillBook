"""Classify bank transactions into credits, debits, and other."""

from __future__ import annotations

import re
from typing import Iterable

import pandas as pd

from nv_billbook.config import Config
from nv_billbook.flats_registry import FlatsRegistry

FLAT_PATTERNS = [
    re.compile(r"\b([ABC])[- ]?(\d{3})\b", re.IGNORECASE),
    re.compile(r"\b([ABC])(\d{3})\b", re.IGNORECASE),
    re.compile(r"\b([ABC]\d{3})\b", re.IGNORECASE),
    re.compile(r"([ABC])(\d{3})(?=[A-Z\s\-]|$)", re.IGNORECASE),
    re.compile(r"-([ABC]\d{3})\s", re.IGNORECASE),
]

INCOME_TYPE_RULES: list[tuple[str, Iterable[str]]] = [
    ("Maintenance", ("maintenance", "maint", "mnt", "main")),
    ("Water Bill", ("waterbill", "water bill", "water")),
    ("AMC", ("amc",)),
    ("Association / Legal", ("association", "legal", "one time")),
    ("Advance / Other Income", ("advance", "billbook", "coconut", "refund")),
]

DEBIT_CATEGORY_RULES: list[tuple[str, Iterable[str]]] = [
    ("Utilities (BESCOM/Water)", ("bescom", "waterbill", "water bill", "watertanker", "tanker")),
    ("STP / Plumbing", ("stp", "plumb", "sewage")),
    ("Electrical", ("electrical", "electric", "motor")),
    ("Solar", ("solar", "plus solar", "plussolar")),
    ("Gardening", ("garden", "gardner", "gardener")),
    ("Fuel / DG", ("fuel", "dg", "gowrishankar")),
    ("Garbage / Waste", ("garbage", "waste", "tractor")),
    ("Security / Housekeeping", ("security", "housekeeping", "house keeping")),
    ("Telecom (Airtel/etc)", ("airtel", "si hg", "telecom")),
    ("Bank / Cash Charges", ("chgs", "charges", "gst", "cash wdl")),
    ("Salary / STP Vendor", ("stp", "anjaney", "salary")),
    ("Repairs & Maintenance", ("repair", "maintenance")),
    ("General Expense", ("tpt-", "neft dr", "self -", "chq paid")),
]


def _normalize_flat(block: str, number: str) -> str:
    return f"{block.upper()}{number}"


def extract_flat_no(text: str) -> str | None:
    upper = text.upper()
    for pattern in FLAT_PATTERNS:
        match = pattern.search(upper)
        if not match:
            continue
        if len(match.groups()) == 2:
            return _normalize_flat(match.group(1), match.group(2))
        token = match.group(1).upper()
        return token
    return None


def extract_income_type(text: str) -> str:
    lowered = text.lower()
    for label, keywords in INCOME_TYPE_RULES:
        if any(keyword in lowered for keyword in keywords):
            return label
    return "Maintenance / General"


def extract_payer_name(text: str) -> str:
    narration = text.strip()

    if narration.upper().startswith("UPI-"):
        body = narration[4:]
        parts = body.split("-")
        if parts:
            return parts[0].strip()

    tpt_match = re.search(r"-TPT-(.+?)-([A-Z .]+)$", narration, re.IGNORECASE)
    if tpt_match:
        return tpt_match.group(2).strip()

    neft_match = re.search(r"NEFT CR-[^-]+-(.+?)-FLAT MAINTENANCE", narration, re.IGNORECASE)
    if neft_match:
        return neft_match.group(1).strip()

    neft_match = re.search(r"NEFT CR-[^-]+-(.+?)-NAVA VAIBHAVA", narration, re.IGNORECASE)
    if neft_match:
        return neft_match.group(1).strip()

    neft_fallback = re.search(r"NEFT CR-[^-]+-(.+?)-", narration, re.IGNORECASE)
    if neft_fallback:
        return neft_fallback.group(1).strip()

    imps_match = re.search(r"IMPS-\d+-(.+?)-", narration, re.IGNORECASE)
    if imps_match:
        return imps_match.group(1).strip()

    return ""


def extract_debit_payee(text: str) -> str:
    narration = text.strip()
    neft_match = re.search(r"NEFT DR-[^-]+-(.+?)-NETBANK", narration, re.IGNORECASE)
    if neft_match:
        return neft_match.group(1).strip()

    tpt_match = re.search(r"-TPT-(.+?)-([A-Z .]+)$", narration, re.IGNORECASE)
    if tpt_match:
        return tpt_match.group(2).strip()

    if narration.upper().startswith("SI "):
        return narration.split()[1] if len(narration.split()) > 1 else "Standing Instruction"

    if "CHQ PAID" in narration.upper():
        return "Cheque Payment"

    if "CHGS" in narration.upper():
        return "HDFC Bank"

    return ""


def extract_debit_purpose(text: str) -> str:
    narration = text.strip()
    neft_match = re.search(r"NETBANK, MUM-HDFCH\d+-(.+)$", narration, re.IGNORECASE)
    if neft_match:
        return neft_match.group(1).strip()

    tpt_match = re.search(r"-TPT-(.+?)-[A-Z .]+$", narration, re.IGNORECASE)
    if tpt_match:
        return tpt_match.group(1).strip()

    if narration.upper().startswith("SI "):
        return narration

    return narration


def categorize_debit(text: str, config: Config) -> str:
    lowered = text.lower()
    for category, keywords in config.categories.items():
        if any(keyword.lower() in lowered for keyword in keywords):
            return category

    for label, keywords in DEBIT_CATEGORY_RULES:
        if any(keyword in lowered for keyword in keywords):
            return label

    return "Uncategorized"


def is_other_transaction(row: pd.Series) -> bool:
    """Flag bank charges, standing instructions, or internal movements as Other."""
    narration = str(row.get("narration", "")).lower()
    other_markers = (
        "chgs",
        "charges incl gst",
        "cash wdl chgs",
        "rev-",
        "reversal",
    )
    return any(marker in narration for marker in other_markers)


def classify_transactions(
    transactions: pd.DataFrame,
    config: Config,
    flats_registry: FlatsRegistry | None = None,
) -> pd.DataFrame:
    """Add classification columns and split type (credit/debit/other)."""
    classified = transactions.copy()

    types: list[str] = []
    categories: list[str] = []
    flat_nos: list[str | None] = []
    sqft_values: list[int | None] = []
    counterparty: list[str] = []
    purpose: list[str] = []

    for _, row in classified.iterrows():
        narration = str(row["narration"])
        if row["credit"] > 0 and row["debit"] == 0:
            txn_type = "credit"
            categories.append(extract_income_type(narration))
            payer = extract_payer_name(narration)
            flat_no = extract_flat_no(narration)
            counterparty.append(payer)
            purpose.append(categories[-1])
        elif row["debit"] > 0 and row["credit"] == 0:
            if is_other_transaction(row):
                txn_type = "other"
            else:
                txn_type = "debit"
            categories.append(categorize_debit(narration, config))
            payer = extract_debit_payee(narration)
            flat_no = extract_flat_no(narration)
            counterparty.append(payer)
            purpose.append(extract_debit_purpose(narration))
        else:
            txn_type = "other"
            categories.append("Uncategorized")
            flat_no = extract_flat_no(narration)
            payer = ""
            counterparty.append("")
            purpose.append(narration)

        sqft_value = None
        if flats_registry:
            resolved_flat, flat_info = flats_registry.resolve_flat(flat_no, payer)
            if resolved_flat and not flat_no:
                flat_no = resolved_flat
            if flat_info:
                sqft_value = flat_info.sqft
            elif flat_no:
                info = flats_registry.get(flat_no)
                if info:
                    sqft_value = info.sqft
        elif flat_no:
            sqft_value = None

        types.append(txn_type)
        flat_nos.append(flat_no)
        sqft_values.append(sqft_value)

    classified["txn_type"] = types
    classified["category"] = categories
    classified["flat_no"] = flat_nos
    classified["sqft"] = sqft_values
    classified["counterparty"] = counterparty
    classified["purpose"] = purpose
    return classified


def split_by_type(classified: pd.DataFrame) -> dict[str, pd.DataFrame]:
    return {
        "credits": classified[classified["txn_type"] == "credit"].copy(),
        "debits": classified[classified["txn_type"] == "debit"].copy(),
        "other": classified[classified["txn_type"] == "other"].copy(),
        "all": classified.copy(),
    }
