# Nava Vaibhva — Monthly Expense Reports

Python tool to read **HDFC bank statements** and generate monthly Excel reports for **Nava Vaibhva Resi Welfare Association**.

Download the society account statement from HDFC net banking each month, run one command, and get a ready-to-review Excel workbook with **Credits**, **Debits**, and **Other** transactions separated.

---

## Latest changes

The following is implemented and working:

| Feature | Details |
|---------|---------|
| **HDFC `.xls` parser** | Auto-detects the transaction header row in HDFC exports — no manual editing |
| **Credit / debit / other split** | Deposits → Credits, vendor payments → Debits, bank charges → Other |
| **Smart extraction** | Flat no (e.g. `A003`, `B303`), payer/payee name, and purpose from UPI/TPT/NEFT narrations |
| **Expense categories** | Auto-tags debits (STP, gardening, electrical, solar, water, fuel, garbage, telecom, etc.) |
| **Excel report** | One workbook per month with 5 sheets (Summary, Credits, Debits, Other, All Transactions) |
| **Flexible input** | Point to a single file, a folder, or use `data/input/` |

**Verified with:** `Acct_Statement_June2026.xls` — 218 transactions processed successfully.

---

## What you get each month

Output file: **`data/output/YYYY-MM_summary.xlsx`**

| Sheet | What it contains |
|-------|------------------|
| **Summary** | Opening/closing balance, total credits & debits, income/expense breakdown by category |
| **Credits** | All deposits — flat no, payer name, type (maintenance / water / AMC) |
| **Debits** | All payments — payee, purpose, expense category |
| **Other** | Bank charges and unmatched entries |
| **All Transactions** | Complete transaction list |

---

## One-time setup (do this once)

### 1. Clone / open the project

```bash
cd ~/Documents/GitHub/latestNVBillBook
```

### 2. Create a virtual environment and install dependencies

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Config file

`config.yaml` is already set up for HDFC. If you ever need to recreate it:

```bash
cp config.example.yaml config.yaml
```

> `config.yaml` is gitignored — it may contain account details, so it stays on your machine only.

---

## How to run every month

Follow these steps at the start of each month for the **previous month's** statement.

### Step 1 — Download the HDFC statement

1. Log in to **HDFC Net Banking**
2. Go to **Accounts → Account Statement**
3. Select the **Nava Vaibhava Resi Welfare Association** account
4. Choose the date range for the month (e.g. 01/07/2026 to 31/07/2026)
5. Download as **Excel (.xls)** — not PDF (PDF is not supported yet)
6. Save the file, e.g.:

   ```
   ~/Downloads/NVBankStatements/Acct_Statement_July2026.xls
   ```

### Step 2 — Activate the environment

```bash
cd ~/Documents/GitHub/latestNVBillBook
source .venv/bin/activate
```

### Step 3 — Run the report

**Option A — Point directly to the downloaded file (recommended):**

```bash
PYTHONPATH=src python3 -m nv_billbook.main \
  --input "~/Downloads/NVBankStatements/Acct_Statement_July2026.xls"
```

**Option B — Copy to `data/input/` and run by month:**

```bash
cp ~/Downloads/NVBankStatements/Acct_Statement_July2026.xls data/input/
PYTHONPATH=src python3 -m nv_billbook.main --month 2026-07
```

**Option C — Process every `.xls` file in your statements folder:**

```bash
PYTHONPATH=src python3 -m nv_billbook.main \
  --input "~/Downloads/NVBankStatements/" --all
```

### Step 4 — Open the report

```bash
open data/output/2026-07_summary.xlsx      # macOS
# xdg-open data/output/2026-07_summary.xlsx   # Linux
```

### Step 5 — Review

1. **Summary** — check opening/closing balance matches the bank statement
2. **Credits** — verify maintenance collections; note any flat with missing payments
3. **Debits** — confirm all vendor payments are categorised correctly
4. **Other** — review bank charges (usually 1–2 entries)
5. Use the data to update your main maintenance workbook if needed

---

## Quick reference — CLI options

| Command | When to use |
|---------|-------------|
| `--input /path/to/file.xls` | Process one specific statement |
| `--input /path/to/folder/ --all` | Process all statements in a folder |
| `--month 2026-07` | Process only files matching July 2026 (uses `data/input/`) |
| `--config path/to/config.yaml` | Use a different config file |

Example output when it succeeds:

```
Processing: Acct_Statement_June2026.xls
  Month: 2026-06 | Credits: 194 (₹605,760.50) | Debits: 23 (₹681,310.46)
  Report written: data/output/2026-06_summary.xlsx

Done. 1 report(s) created in data/output/
```

---

## Monthly checklist

```
[ ] Download HDFC .xls statement for the month
[ ] Activate venv:  source .venv/bin/activate
[ ] Run script with --input pointing to the file
[ ] Open data/output/YYYY-MM_summary.xlsx
[ ] Verify Summary totals vs bank statement
[ ] Review Credits sheet (flat-wise collections)
[ ] Review Debits sheet (expense categories)
[ ] Archive statement + report for records
```

---

## Example — June 2026

| Metric | Value |
|--------|-------|
| Credit transactions | 194 |
| Debit transactions | 23 |
| Other (bank charges) | 1 |
| Total credits | ₹6,05,760.50 |
| Total debits | ₹6,81,310.46 |
| Opening balance | ₹21,62,369.72 |
| Closing balance | ₹20,86,481.51 |

Typical credit narrations: `A-003 JUNE 26 MAIN`, `B303JUNMAINTENANCE`, `C307 WATERBILL`

Typical debit narrations: STP vendor, gardening, electrical, solar, water tanker, fuel, garbage, Airtel SI

---

## HDFC statement format

HDFC `.xls` exports have a header block, then a transaction table:

| Date | Narration | Chq./Ref.No. | Value Dt | Withdrawal Amt. | Deposit Amt. | Closing Balance |
|------|-----------|--------------|----------|-----------------|--------------|-----------------|

The parser finds this table automatically — you do not need to edit the file.

---

## Project structure

```
latestNVBillBook/
├── README.md
├── requirements.txt
├── config.yaml              # Local HDFC settings (gitignored)
├── config.example.yaml      # Template — safe to commit
├── data/
│   ├── input/               # Optional: drop statement files here
│   └── output/              # Generated Excel reports appear here
├── src/nv_billbook/
│   ├── main.py              # CLI entry point
│   ├── config.py            # Load config.yaml
│   ├── parser.py            # HDFC XLS parser
│   ├── classifier.py        # Credit / debit / flat extraction
│   └── reporter.py          # Excel workbook builder
└── tests/
    └── test_parser.py
```

---

## Customising expense categories

Edit `config.yaml` to add or change keyword rules:

```yaml
categories:
  Utilities (BESCOM/Water):
    - "bescom"
    - "watertanker"
    - "waterbill"
  Gardening:
    - "garden"
    - "gardner"
  Garbage / Waste:
    - "garbage"
    - "waste"
    - "tractor"
```

Re-run the script after saving — no code changes needed.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `Config not found` | Run `cp config.example.yaml config.yaml` |
| `No statement files found` | Check the `--input` path; use the full path to the `.xls` file |
| `Could not find transaction header row` | Make sure the file is HDFC Excel export, not PDF |
| `ModuleNotFoundError` | Activate venv and run `pip install -r requirements.txt` |
| Wrong month in output filename | The tool reads the date range from inside the statement; verify the download covers the correct month |

---

## Roadmap

- [x] HDFC `.xls` parser with auto header detection
- [x] Credit / debit / other classification
- [x] Flat no & payer name extraction from UPI/TPT narrations
- [x] Monthly Excel report (Summary, Credits, Debits, Other, All)
- [ ] PDF statement support
- [ ] Flat-wise collection pivot (like legacy NV Maintenance workbook)
- [ ] Multi-month comparison sheet

---

## Notes

- Keep raw bank statements for audit — they live in `~/Downloads/NVBankStatements/` or `data/input/`
- Generated reports in `data/output/` are gitignored
- Do not commit `config.yaml` if it contains account numbers
- For new UPI/narration patterns, share 2–3 examples and the parser can be improved
