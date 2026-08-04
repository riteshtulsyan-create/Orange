---
name: fetch-bank-transactions
description: Fetch bank payment and receipt vouchers from Tally Prime for all loaded companies and produce two Excel reports — a flat workbook with static summary sheets and a second workbook with fully interactive native Excel PivotTables. Use this skill whenever the user says anything like "fetch bank transactions", "run bank transactions report", "generate bank transactions pivot", "pull payments from Tally", "bank voucher report", "update the bank transactions file", "refresh the bank data", or "show me payments and receipts". Trigger it even if the user only mentions one of the two outputs (flat file or pivot) — the skill always produces both together.
---

# fetch-bank-transactions

Runs the two-script pipeline that pulls every bank payment and receipt voucher from Tally Prime (all loaded companies, 1-Apr-2026 to 4-Aug-2026), classifies each row, writes a flat Excel workbook, then builds a second workbook with native Excel PivotTables on top.

The date range defaults to the constants in the script but can be overridden per-run without editing anything — see "Environment overrides" below. To extend the report to "till date", set `TO_DATE` to today.

## What the pipeline does

**Step 1 — `fetch-bank-transactions.mjs` (Node.js)**

- Connects to Tally Prime at `localhost:9000` (port comes from `.env` / `TALLY_PORT`)
- Fetches Payment and Receipt vouchers for every loaded company
- Classifies each row across 6 computed columns:
  - `Final Voucher` — Payment / Receipt
  - `Voucher Classification` — Bank / Cash / CHQ.Ret
  - `Month` — e.g. Apr-26
  - `TYPE` — granular expense/asset category (FINANCE COST, EMPLOYEE TRAVEL COST, Debtors, etc.)
  - `Category` — OTH / Related / Branch
  - `BS/PL` — BS / PL / Loan
- Classification uses a 3-priority chain: hardcoded overrides → v3 reference file → pattern-based rules
- Produces `outputs/Bank_Transactions_Apr_Aug2026.xlsx` with sheets:
  - Summary (per-company payment & receipt totals)
  - ALL / BANK / CASH / RETUNN (static pivot-style summaries by TYPE × Month)
  - Related Partry / Bracnch Mis (party-level breakdowns)
  - All Transactions (flat ledger, all 4900+ rows)
  - One sheet per company
- At the end of its run it **automatically calls** `create_pivot_report.py`

**Step 2 — `create_pivot_report.py` (Python)**

- Reads the flat workbook produced by Step 1
- Opens it in Excel via COM automation (xlwings)
- Adds **native interactive PivotTables** on 8 sheets (ALL, Type, Current Month, BANK, CASH, RETUNN, Related Partry, Bracnch Mis)
  - Main pivots (ALL / BANK / CASH / RETUNN): Final Voucher → BS/PL → TYPE in rows; Month in columns; Company as page filter; TYPEs sorted highest to lowest. BANK/CASH/RETUNN are pre-filtered to that Voucher Classification.
  - `Type`: same row/column layout as ALL, but with Company, Category, Voucher Classification and Party all available as page filters and none pre-set — the general-purpose slice-and-dice sheet.
  - `Current Month`: drill-down for the latest month in the data. Rows go one level deeper (Final Voucher → BS/PL → TYPE → Party) and columns are individual dates; Month is a page filter pre-set to the latest month.
  - Party pivots (Related Partry / Bracnch Mis): Co Code → Final Voucher → Party in rows; Month in columns; Category pre-filtered to Related / Branch; sorted highest to lowest
  - Months always appear in chronological order (Apr-26 → Mar-27)
- Saves the result as `outputs/Bank_Transactions_Apr_Aug2026_Pivot.xlsx`

## How to run

The entire pipeline is a single command. Run it from the `PL and BS` project directory (so `.env` is found):

```powershell
cd "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS"
node ".claude/skills/fetch-bank-transactions/scripts/fetch-bank-transactions.mjs"
```

Both output files are written to the `outputs/` subfolder of `PL and BS`.

To scope the fetch to a specific fiscal-year company set (see `Files/TALLY Company_List.xlsx` and CLAUDE.md "Company Period Filtering"), pass `--period`:

```powershell
node ".claude/skills/fetch-bank-transactions/scripts/fetch-bank-transactions.mjs" --period 2526
```

If the companies required for that period aren't currently open in Tally, the script stops and lists exactly which ones to open/close before re-running — it never silently runs on a partial or wrong company set.

**Prerequisites that must be true before running:**
- Tally Prime is open and listening on the host/port in `.env` (currently `localhost:9000`)
- The companies whose data you want are loaded/open in Tally (matching the `--period` code, if given)
- Node.js with the `xlsx` and `fast-xml-parser` packages installed in `PL and BS` (`npm install xlsx fast-xml-parser`), Python, `openpyxl`, and `xlwings` are installed
- Microsoft Excel is installed (xlwings drives it via COM)

## Environment overrides

| Variable | Default | Purpose |
|---|---|---|
| `TALLY_HOST` | from `.env` (`localhost`) | Tally host |
| `TALLY_PORT` | from `.env` (`9000`) | Tally HTTP port |
| `FROM_DATE` | `1-Apr-2026` | Period start (Tally date format, e.g. `1-Apr-2026`) |
| `TO_DATE` | `4-Aug-2026` | Period end — set to today for a "till date" report |
| `OUTPUT_FILE` | `outputs/Bank_Transactions_Apr_Aug2026.xlsx` | Flat output path |
| `REFERENCE_FILE` | `outputs/Bank_Transactions_Apr_Jun_2026_v3.xlsx` | Manual classification reference (v3) |
| `SOURCE_FILE` | same as OUTPUT_FILE | Input for pivot builder |
| `OUTPUT_FILE` (pivot script) | `outputs/Bank_Transactions_Apr_Aug2026_Pivot.xlsx` | Pivot output path |
| `EXPENSE_FILE` | `Files/EXPENSES REFERESH/Expense Sheet 2026-07-14 17-38-44.xlsx` | Ledger→MAIN GROUP cross-check source |
| `COMPANY_LIST_FILE` | `Files/TALLY Company_List.xlsx` | Company-code → company list source |
| `COMPANY_CODE` (or `--period`) | `26-27` | Which fiscal-year company set to fetch (`26-27`/`25-26`/`24-25`) |

⚠️ **`OUTPUT_FILE` is read by both scripts.** The Node script launches the Python pivot builder with the
same environment, so setting `OUTPUT_FILE` to redirect the flat file also redirects the pivot file to
that same path — the pivot build would overwrite the flat workbook. If you override `OUTPUT_FILE`, set
`SOURCE_FILE` to the same value and pick a distinct pivot path, or just run the two scripts separately.

Example — a run through today without editing the scripts:

```powershell
$env:TO_DATE = "4-Aug-2026"
node ".claude/skills/fetch-bank-transactions/scripts/fetch-bank-transactions.mjs"
Remove-Item Env:\TO_DATE
```

## Output files

| File | Contents |
|---|---|
| `outputs/Bank_Transactions_Apr_Aug2026.xlsx` | Flat data + static summaries — good for lookup and cross-checking |
| `outputs/Bank_Transactions_Apr_Aug2026_Pivot.xlsx` | Native Excel PivotTables — use the Company / Voucher Classification / Category filters interactively |

Filenames follow the period, so changing `TO_DATE` into a new month means changing the two filename
defaults as well (or passing them explicitly). Earlier periods' files are left in place, not overwritten.

## Classification logic (key rules)

- `BANK CHARGES*` sub-groups → `FINANCE COST`
- `CREDITOR FOR EMPLOYEE (EXP)` → `EMPLOYEE TRAVEL COST`
- `CREDIT CARD` → `EMPLOYEE TRAVEL COST`
- `CREDITOR FOR OTHER*` → pattern-matched by party name:
  - Clearing/Forwarding/Logistics → `CLEARING & FORWERDING EXP`
  - Transport/Roadways → `TRANSPORTATION EXP`
  - Electricity → `ELECTRICITY EXPENSES`
  - Rent → `RENT EXPENSES`
  - Hotel/Travel → `EMPLOYEE TRAVEL COST`
  - Repair/Maintenance → `REPAIR & MAINTENANCE COST`
  - Default → `EXPENSES`
- Hardcoded overrides (data-quality fixes): `MITHILI PROFFESIONAL'S-Cross Creation` → `Debtors / BS`
- If a party appears in the v3 reference file (`Bank_Transactions_Apr_Jun_2026_v3.xlsx` → "All Transactions" sheet), its manually classified TYPE takes priority over pattern rules

## Troubleshooting

- **No data / connection refused** — Tally Prime is not running, or is listening on a different port than `.env` says. Check what it is actually bound to (`Get-NetTCPConnection -State Listen | Where-Object { $_.OwningProcess -eq (Get-Process tally).Id }`) and update `TALLY_PORT` in `.env`. Also check that companies are open in Tally.
- **Pivot script fails with COM error** — Excel may not be installed, or another Excel instance is blocking the file. Close any open Excel windows and re-run.
- **`ModuleNotFoundError: No module named 'xlwings'`** — install it: `python -m pip install xlwings` (pulls in `pywin32`). Step 1 still succeeds without it; only the pivot workbook is skipped.
- **`OLE error 0x800a01a8` mid-way through "Building PivotTables"** — a transient Excel COM fault (VBA error 424), not a data or code problem. It has been observed failing on one pivot and then succeeding on an identical re-run with the same data. Just re-run `create_pivot_report.py`; the flat workbook does not need rebuilding.
- **Node script exits 0 even though the pivot build failed** — known wart: the pivot step's error is caught and printed but does not set a non-zero exit code. Always check the tail of the output for `Done.` rather than trusting the exit status.
- **`Amount (₹)` field not found** — Occurs only if the flat file is empty (0 data rows). Check Tally connection and date range.
- **Reference file not found** — The v3 manual classification file is optional. Without it, all classification falls back to pattern rules. Place it at `outputs/Bank_Transactions_Apr_Jun_2026_v3.xlsx` to re-enable it.

## Modifying date range or output filenames

Prefer the `FROM_DATE` / `TO_DATE` environment variables above for a one-off run. To change the
persistent defaults, edit the constants near the top of `scripts/fetch-bank-transactions.mjs`:

```javascript
const FROM_DATE   = process.env.FROM_DATE || '1-Apr-2026';  // default start date
const TO_DATE     = process.env.TO_DATE   || '4-Aug-2026';  // default end date
const OUTPUT_FILE = ...                                     // change output filename here
```

And in `scripts/create_pivot_report.py`:

```python
SOURCE_FILE = ...   # must match OUTPUT_FILE above
OUTPUT_FILE = ...   # pivot output filename
```
