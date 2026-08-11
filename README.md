# Nava Vaibhva — Monthly Expense Reports

Python tool to read **HDFC bank statements** and generate monthly Excel reports for **Nava Vaibhva Resi Welfare Association**.

Download the society account statement from HDFC net banking each month, run one command, and get a ready-to-review Excel workbook with transaction details and a flat-wise maintenance reconciliation.

---

## Latest changes

The following is implemented and working:

| Feature | Details |
|---------|---------|
| **HDFC statement parser** | Auto-detects the transaction header row in HDFC exports — no manual editing |
| **Credit / debit / other split** | Deposits → Credits, vendor payments → Debits, bank charges → Other |
| **Smart extraction** | Flat no (e.g. `A003`, `B303`), payer/payee name, and purpose from UPI/TPT/NEFT narrations |
| **Expense categories** | Auto-tags debits (STP, gardening, electrical, solar, water, fuel, garbage, telecom, etc.) |
| **Maintenance reconciliation** | Calculates expected maintenance, water bill, paid amount, dues, excess, and status for every registered flat |
| **Excel report** | One workbook per month with 6 sheets, including Maintenance Reconciliation |
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
| **Maintenance Reconciliation** | One row per registered flat: expected charges, payments received, amount due or extra paid, and status |

---

## Multi-month collection history

For one HDFC statement that covers 3, 6, or 12 calendar months, create a separate consolidated workbook without changing the monthly report:

```bash
PYTHONPATH=src python3 -m nv_billbook.history \
  --input "/full/path/to/financial_year_HDFC_statement.xls"
```

The output is `data/output/YYYY-MM_to_YYYY-MM_collection_history.xlsx` with one worksheet per registered flat. For example, the first tab is `A001` and contains only A001 data:

If you want only one flat, pass `--flat A001` and the workbook will contain just that flat:

```bash
PYTHONPATH=src python3 -m nv_billbook.history \
  --input "/full/path/to/financial_year_HDFC_statement.xls" \
  --flat A001
```

That creates `data/output/YYYY-MM_to_YYYY-MM_A001_collection_history.xlsx`.

If your 3, 6, or 12 months are saved as separate statement files, point `--input` to the folder instead of a single file. The command will read every statement in that folder, combine them, and calculate the history across the full period:

```bash
PYTHONPATH=src python3 -m nv_billbook.history \
  --input "/full/path/to/folder-with-statements"
```

You can still combine this with `--flat A001` if you only want one flat.

| Section in each flat worksheet | What it contains |
|---|---|
| **Period Summary** | Expected, paid, due, extra paid, and overall status for the complete statement period |
| **Monthly Reconciliation** | Paid vs expected for each statement month, including due, extra paid, and status |
| **Payment Transactions** | Every mapped incoming payment for that flat, with date, month, payer, type, and amount |

Before running it, configure the water bill for every month in the statement under `meta.water_bills_by_month` in `flats.yaml`. A financial-year report needs values from April through the following March.

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
  --input "~/Downloads/NVBankStatements/Acct_Statement_July2026.xls" \
  --water-bill 287
```

`--water-bill` is the fixed per-flat water charge for this report. It overrides the fallback value in `flats.yaml`, so you can enter the current amount without editing any files.

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
2. **Maintenance Reconciliation** — check Pending, Short, and Excess flats; confirm the water bill used for the month
3. **Credits** — verify maintenance collections and investigate rows without a Flat No
4. **Debits** — confirm all vendor payments are categorised correctly
5. **Other** — review bank charges (usually 1–2 entries)

---

## Quick reference — CLI options

| Command | When to use |
|---------|-------------|
| `--input /path/to/file.xls` | Process one specific statement (`.xls`, `.xlsx`, `.xlsm`, or `.csv`) |
| `--input /path/to/folder/ --all` | Process all statements in a folder (`.xls`, `.xlsx`, `.xlsm`, or `.csv`) |
| `--month 2026-07` | Process only files matching July 2026 (uses `data/input/`) |
| `--config path/to/config.yaml` | Use a different config file |
| `--water-bill 287` | Set the fixed per-flat water bill for this report; overrides `flats.yaml` |

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
[ ] Review Maintenance Reconciliation — Pending, Short, and Excess flats
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

HDFC exports have a header block, then a transaction table:

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
├── flats.yaml               # Flat registry — owner names & payer aliases
├── flats.example.yaml       # Backup template of flat registry
├── data/
│   ├── input/               # Optional: drop statement files here
│   └── output/              # Generated Excel reports appear here
├── src/nv_billbook/
│   ├── main.py              # CLI entry point
│   ├── config.py            # Load config.yaml
│   ├── flats_registry.py    # Flat lookup & payer alias matching
│   ├── parser.py            # HDFC XLS parser
│   ├── classifier.py        # Credit / debit / flat extraction
│   ├── reconciliation.py    # Flat-wise maintenance reconciliation logic
│   ├── reporter.py          # Excel workbook builder
│   ├── collection_history.py # Multi-month history and reconciliation logic
│   ├── history_reporter.py  # Multi-month history workbook builder
│   └── history.py           # Multi-month history CLI entry point
└── tests/
    ├── test_parser.py
    ├── test_reconciliation.py
    └── test_collection_history.py
```

---

## Flat registry (`flats.yaml`)

Reference file for all **200 flats**. Built from two sources only:

1. **SBA dimension table** (FLAT NO / sqft) — your committee flat-size reference
2. **HDFC bank statement** — payer names seen in credit transactions (June 2026 as starting point)

No data is pulled from the old maintenance workbook.

### What the tool uses it for

- Attach **SBA (sqft)** to each credit when a flat number is known
- Match **UPI payer names** to flats when the narration has no flat number
- Track **`payer_names`** per flat as you confirm them from bank statements
- Provide the maintenance rate and fallback water bill for reconciliation

### Reconciliation settings

Store the standard values in `meta`:

```yaml
meta:
  maintenance_rate_per_sqft: 2.5
  water_bill_per_month: 287  # fallback when --water-bill is not supplied
  water_bills_by_month:      # required for every month in the history report
    "2026-04": 221
    "2026-05": 242
    # Add every month in the selected statement period.
    "2027-03": 287
```

For a month with a different water charge, use the command-line option instead of editing the file:

```bash
PYTHONPATH=src python3 -m nv_billbook.main \
  --input "/full/path/to/HDFC_statement.xls" \
  --water-bill 287
```

The reconciliation sheet calculates `Expected Maintenance = Sqft × maintenance rate`, then adds the water bill. It sums every mapped credit for the flat and assigns one of four statuses: **Paid**, **Excess**, **Short**, or **Pending**.

The multi-month history report selects the water bill using the transaction month. It supports 3-, 6-, and 12-month statements and stops with a clear error if any statement month is missing from `water_bills_by_month`.

### Structure (per flat)

```yaml
flats:
  A003:
    block: A
    sqft: 1095                         # from dimension table
    payer_names:                       # from bank statement only
      - "NETHRAVATHI  S"
    notes: ""
```

Flats with no payments in June 2026 have an empty `payer_names: []` — fill these in over time.

### Unmapped payers

Payers in bank credits **without** a flat number in the narration are listed at the bottom under `unmapped_payers`. Each month:

1. Check the **Credits** sheet for empty **Flat No** rows
2. Find the payer in `unmapped_payers`
3. Once you know the flat, add the name under that flat's `payer_names`:

```yaml
  A320:
    sqft: 1130
    payer_names:
      - "MANU N"
```

4. Update `meta.last_updated` and re-run the report

### SBA dimensions (`flat_dimensions.yaml`)

Exact sqft per flat — sourced from **`OnlyFlatDimensions.xlsx`**. Never rounded or estimated.

```yaml
dimensions:
  A003: 1095
  A320: 1228
```

After editing, sync into `flats.yaml`:

```bash
python3 scripts/sync_flat_dimensions.py
```

If any sqft looks wrong, correct it in `flat_dimensions.yaml` (not in code). The tool reads the exact integer you provide.

---

1. Download the new HDFC statement and run the report
2. Review new payers in **Credits** (empty Flat No column)
3. Add confirmed payer → flat mappings to `flats.yaml`
4. Re-run — **Flat No** and **SBA (sqft)** fill in automatically

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
- [x] Flat registry (`flats.yaml`) — 200 flats with SBA sqft + bank payer names
- [x] Payer-to-flat matching from bank statement names only
- [x] Flat-wise maintenance reconciliation worksheet
- [x] Multi-month (3, 6, or 12) flat transaction history and reconciliation report
- [ ] PDF statement support
- [ ] Flat-wise collection pivot (like legacy NV Maintenance workbook)
- [ ] Multi-month comparison sheet

---

## Notes

- Keep raw bank statements for audit — they live in `~/Downloads/NVBankStatements/` or `data/input/`
- Generated reports in `data/output/` are gitignored
- Do not commit `config.yaml` if it contains account numbers
- For new UPI/narration patterns, share 2–3 examples and the parser can be improved
