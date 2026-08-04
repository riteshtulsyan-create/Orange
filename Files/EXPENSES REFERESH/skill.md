# Skill: Refresh Expense Sheet

## When to Use
Use this skill when the user wants to:
- Refresh/re-fetch the expense sheet from Tally
- Fetch expense data from Tally for a period
- Regenerate the Direct/Indirect expense workbook and its analysis sheets
- Check what changed in the expense sheet since the last refresh
- Phrases like "refresh the expense sheet", "re-fetch expenses from Tally",
  "update the expense workbook for [month/quarter]", "check the gap from Tally"

## What It Does
End to end, in one command:
1. Fetches all Direct + Indirect Expense vouchers from Tally for the 5
   operating entities (Otec-Surat, Otec-Noida, Ent-Surat, Ent-Noida, Colorix)
   for the given date range, and writes them to the `RAW DATA` sheet. Which
   underlying Tally company FILE to use for each entity is resolved per
   period (`2627`/`2526`/`2425`) from `Files/TALLY Company_List.xlsx` — see
   CLAUDE.md's "Company Period Filtering" section, the project-wide standing
   rule this skill follows. Never hardcode a company name/list here; always
   resolve it via `generate_expense_sheet.companies_for_period(period)`.
2. Classifies every DIRECT row into a component category (Head/Ink/Machine/
   Spare/Fabric/General-Other) based on the ledger's TYPES/LEDGER NAME.
3. Builds every analysis sheet on top, as live Excel PivotTables sharing one
   PivotCache (so they stay refreshable and always reconcile with each other):
   - **Summary** — EXP TYPE > MAIN GROUP × MONTH, sorted biggest-expense-first
     within each EXP TYPE, with an **AUDIT FLAG** column highlighting:
     top-5 cost drivers, any main group where one vendor is >70% of the total,
     and main groups literally named "Direct Expenses"/"Indirect Expenses"
     (recomputed fresh from whatever the new data shows - not tied to any past
     quarter's specific findings).
   - **Direct** — MAIN GROUP > TYPES × MONTH (Direct expenses only)
   - **Indirect Exp** — MAIN GROUP > SUB-GROUP × MONTH (Indirect expenses only)
   - **Direct by Component** — three pivots: Main Group × Month; a
     Component Category → Main Group → Vendor drill-down; and a
     Main Group × Component Category cross-tab
   - **One sheet per Indirect main group** (e.g. Employee Cost, Finance Cost,
     Advertisement, ...) — Vendor × Month, with a **Remarks** column drafted
     from the ledger's own narrations (clustered into one short line, not a
     "+N other narrations" count)
   - **Index** — hyperlinks to every sheet, a back-link on each sheet, and a
     **Must Review** column that rolls up every AUDIT FLAG from Summary onto
     the sheet that shows that main group's detail (highlighted in the same
     amber), so you can see at a glance which sheets need attention and jump
     straight there
4. Reorders tabs (Summary, Direct, Direct by Component, Indirect Exp,
   Advertisement, other indirect sheets alphabetically, RAW DATA last; Index
   inserted first) and autofits every row height.
5. Backs up the previous workbook (to `backups/`) before overwriting it.

Separately (not part of the standard refresh - run on request, e.g. "who last
touched the Tally data" / "before we print"): `scripts/last_entry_check.py`
adds/refreshes a **Last Entry Log** sheet (right after Index) showing, per
company, the most recently created/edited voucher (by Tally's own ALTERID
edit-sequence counter, independent of the voucher's own accounting date) with
its date, type, number, and the Tally username who entered/altered it. This
Tally setup does not record clock-time for entries, only the day - the sheet
says so explicitly, don't imply more precision than that.

## Scripts
- `scripts/fetch_tally_expenses.py` — Tally XML API client (ledger/group master
  walk → EXP TYPE/MAIN GROUP/SUB-GROUP/SUB-SUB GROUP/TYPES classification,
  voucher fetch, employee-name and department lookup)
- `scripts/generate_expense_sheet.py` — orchestrates the Tally fetch and writes
  the `RAW DATA` sheet only. Owns `companies_for_period(period_code)`, the
  single source of truth this whole folder uses to resolve which 5 Tally
  company files to fetch for `2627`/`2526`/`2425` (reads `Files/TALLY
  Company_List.xlsx`, falling back to reading it via COM if it's open/locked).
  Never hardcode a company name/list elsewhere in this folder - call this.
- `scripts/last_entry_check.py` — `fetch_all_last_entries(period="2627")` +
  `add_last_entry_sheet(workbook_path, rows)`; see the Last Entry Log
  description above
- `scripts/remark_logic.py` — clusters a vendor's raw narrations into one short
  Remarks line
- `scripts/build_analysis_sheets.py` — builds every pivot/Index/Summary-flag
  sheet described above on top of an existing `RAW DATA` sheet. Idempotent —
  safe to re-run against an already-built workbook (e.g. to re-scope the month
  range without re-fetching): it clears every sheet except RAW DATA before
  rebuilding. Builds on a local scratch copy and copies the result back over
  the target path with a plain filesystem copy - see "OneDrive/SharePoint
  lock" below for why.
- `scripts/refresh_expense_sheet.py` — the one command that runs fetch + build,
  for a STANDARD refresh (one date range, the 5 current companies)
- `scripts/check_gap.py` — compares the live workbook's RAW DATA against the
  newest file in `backups/` (or an explicitly given one) and reports what
  genuinely changed: new transactions, removed transactions, and amount
  corrections - filtering out cosmetic noise (ledger renames, voucher-type
  reclassification from GST-return finalization) that would otherwise look
  like false adds/removes in a naive diff.
- `scripts/refresh_with_history.py` — a ONE-OFF script (2026-07-16), currently
  **on hold** (`FETCH_HISTORY = False`). See "Historical FY25-26 data" below.

## How to Run

### Step 1: Ask for date range and period (if not provided)
Ask the user: "What date range? (e.g. Apr-Jun 2026 = 20260401 to 20260630)"
and, if not already stated or obvious from the dates, which period code
applies (`2627`/`2526`/`2425` — see CLAUDE.md). Defaults to `2627` if the
user doesn't say and the dates are clearly current-year.

### Step 2: Confirm Tally prerequisites
- Tally must be running with the companies that `2627`/`2526`/`2425` (per
  `Files/TALLY Company_List.xlsx`) resolve to actually open — companies not
  open are skipped, not an error, but check the printed "Companies required
  for period X" list against Tally's open-company list before trusting a run
  with fewer than 5 rows-added companies.
- Connection settings come from the project's `.env` (`TALLY_HOST`/`TALLY_PORT`,
  canonically `localhost:9010` per the project CLAUDE.md — never hardcode a
  different port).
- **Excel must be closed** on this machine before running — the script drives
  Excel via COM and needs exclusive access to the workbook file. (`Files/TALLY
  Company_List.xlsx` itself is fine to leave open — the period lookup falls
  back to reading it via COM if it's locked.)

### Step 3: Run the refresh
```
cd "C:\Users\orang\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS\Files\EXPENSES REFERESH"
python scripts/refresh_expense_sheet.py --from 20260401 --to 20260630 --period 2627
```
This backs up the previous version to `backups/` first, then overwrites
`Expense Sheet 2026-07-14 17-38-44.xlsx` in place. To write somewhere else
instead, add `--out "path\to\file.xlsx"`. `--period` defaults to `2627` if
omitted.

### Step 4: Check the gap against the previous version
```
python scripts/check_gap.py
```
Reports genuinely new/removed/changed transactions and old-vs-new totals by
EXP TYPE and by company, so you can tell the user exactly what moved since
the last refresh instead of just "it's updated."

### Step 5: Verify the save actually happened, then share the path
After `build_analysis_sheets.py` runs, confirm the write really landed (don't
just trust the script's own "SAVED" print - see "OneDrive/SharePoint lock"
below):
```powershell
Get-Item "...\Expense Sheet 2026-07-14 17-38-44.xlsx" | Select-Object LastWriteTime, Length
```
The timestamp should be within the last few minutes. Then report the file
path, the gap-check findings, and that a backup of the prior version was kept.

## Output Format (23 columns, RAW DATA sheet)
Company | EXP TYPE | MAIN GROUP | SUB-GROUP | SUB-SUB GROUP | TYPES | DATE |
LEDGER NAME | VCH. TYPE | VCH NO. | NET AMOUNT | MONTH | LOCATION | NARRATION |
SECTIONS | REMARK-1 | REMARK-2 | REMARK-3 (Specified period) | EMPLOYEE NAME |
DEPARTMENT | Financial Quarter | Company Name | COMPONENT CATEGORY

No cell is left blank: LEDGER NAME and the NARRATION/SECTIONS/REMARK-1/2/3
fields fall back to the row's TYPES (expense ledger name) when nothing else
can be derived. EMPLOYEE NAME shows "GROUP EMPLOYEE" when a narration
references more than one employee. DEPARTMENT is looked up from a fixed
Employee → Department table in `fetch_tally_expenses.py` (there is no
Department field in Tally's own ledger/voucher masters for this company) —
employees not in that table get a blank DEPARTMENT, same as before.

## Narration Parsing
SECTIONS, REMARK-1, REMARK-2, REMARK-3, EMPLOYEE NAME are parsed from the Tally
voucher narration using pipe (`|`) as delimiter:
`Main narration | SECTIONS | REMARK-1 | REMARK-2 | REMARK-3 | EMPLOYEE NAME`

## OneDrive/SharePoint ReadOnly lock (important - read before debugging a "nothing saved" issue)
The workbook lives in a SharePoint-backed OneDrive folder. Opening it directly
via Excel COM can hit a transient co-authoring/sync share-violation; when that
happens Excel silently falls back to opening the file **ReadOnly**, and
`Workbook.Save()` then no-ops without ever raising an error - the script prints
"SAVED" but nothing was actually written to disk. This bit us once (2026-07-16):
several rebuilds looked successful but the file's timestamp never moved.
Fix already in place: `build_analysis_sheets.build()` copies the workbook to a
local temp folder first, does all the Excel COM work there (never touches the
live OneDrive path while Excel has it open), then copies the result back with
a plain filesystem copy. If you ever see the LastWriteTime not move after a
run, that's the symptom to look for.

## Historical FY25-26 data for Ent-Surat / Ent-Noida - on hold
The two "historical" Tally company files -
`ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2024-26)` and
`ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA - FY25-26` - cannot currently
supply voucher-level detail via any API available here (confirmed twice, via
this project's own TDL fetch and independently via the Tally MCP connector's
trial-balance tool): Trial Balance proves real FY25-26 activity exists in both
(1,012 ledgers / Rs.14.8Cr and 540 ledgers / Rs.3.8Cr respectively), but the
voucher-level Collection query returns almost nothing for them. The user has
asked NOT to query these two companies for now. `refresh_with_history.py` has
`FETCH_HISTORY = False` to hard-disable that fetch step - do not flip it back
on without confirming a working access path first (most likely: a manual Day
Book export from the Tally GUI - Gateway of Tally → Day Book → set period →
Alt+E → Export → XML/Excel - for the two companies, which could then be parsed
and merged into RAW DATA the same way as the API-fetched rows).

## Known limitations / things worth knowing before re-running
- **AUDIT FLAG / Must Review reasons regenerate from scratch each run.** They
  will not mention specific things found in a past review (e.g. "training
  expense missing from PL_BS_GL_Detail") — only the three structural rules
  above. If you want that kind of deep-dive finding, ask for it as a separate
  analysis after refreshing.
- **Cross-checking against `PL_BS_GL_Detail.xlsx` is a separate, manual step**
  — this skill does not read that file. Ask for a reconciliation explicitly
  once both files are current.
- Remarks are plain text sitting next to a live pivot. If you refresh again
  with a wider date range, vendor rows can reorder and old remarks won't
  follow automatically — this skill regenerates them fresh every run, so that
  isn't an issue as long as you always refresh through this skill rather than
  hand-editing the workbook in between.
