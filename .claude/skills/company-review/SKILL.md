---
name: company-review
description: Build or refresh the quarterly "Company Review" — two Excel workbooks (Company_Review_Board_Pack and Company_Action_Plan) plus an HTML board dashboard, synthesized from PL_BS_GL_Detail.xlsx, GP REPORT.xlsx, the expense workbooks, Working_Capital_Review, CFS PL MOM.xlsx, and the Machine Sales Reports. Use this skill whenever the user asks to build, refresh, or update the Company Review / Board Pack / Action Plan / board dashboard for a new quarter, or says the Tally-derived source files have changed and the review needs to catch up.
---

# Company Review (Quarterly)

Produces three deliverables from six-to-eight Tally-derived source files:

1. `Files\Company_Review_Board_Pack_Q1_FY26-27.xlsx` — the newer, more complete workbook (3-entity consolidated view: Orange O Tec / Enterprises / Colorix, plus Machine Sales and BS-trend/NWC sheets).
2. `Files\Company_Action_Plan_Q1_FY26-27.xlsx` — an older, 5-company Surat/Noida breakdown layout, kept in parallel because it was the first deliverable the user relied on.
3. An HTML board dashboard, published as a Claude Artifact.

Both workbooks and the dashboard must always show **the same figures** — this skill's core discipline is keeping three independently-built artifacts in sync every time a source number changes.

## Non-negotiable formatting rules for this project

- **Tabular, labeled-column layout only.** No narrative bullets with numbers embedded inline, no `[Ref: ...]` citation tags. Every takeaway/action table uses explicit columns (e.g. Segment / Now / Before / Change / Takeaway). This was explicitly requested and reverted twice when violated.
- **Green = favourable, red = unfavourable**, via `CellIsRule` conditional formatting (see `green_red_cf` helper in the scripts). For expenses, a *decrease* is favourable — don't flip the polarity by copy-pasting the P&L convention.
- **The word "CFO" must never appear anywhere** in either workbook or the dashboard — an explicit standing instruction for this deliverable (unlike some older sibling skills in this project, which predate the instruction and still use the word).
- **All figures inside each workbook are formulas referencing a "Source Data - ..." sheet**, never duplicated as literals in the Executive Summary / Action Plan / Expense Details sheets. The Source Data sheets themselves hold plain literals (a documented, explicit snapshot — not live links to the external source files; say so in "Read Me First").
- Sheets are interlinked with clickable internal hyperlinks (`add_link` helper) — every sheet has a way back to Executive Summary.

## Step-by-step process to refresh for a new quarter

### 1. Confirm Tally is reachable and figure out what changed
Ask the user (or check directly) whether the P&L/BS workbook and the other source files have been
refreshed since the last run. Note: `fetch-pandl` no longer writes a fixed `PL_BS_GL_Detail.xlsx`
— every refresh produces a new dated `PL_BS_GL_Detail_<timestamp>.xlsx` in `Files\`, with older
ones moved to `Files\Old Version of PL BS\` (see that skill's SKILL.md, "File naming and version
history"). "Refreshed since the last run" means check the timestamp embedded in the newest
`PL_BS_GL_Detail_*.xlsx` filename still sitting in `Files\`, not a file-modified-date check on a
fixed name. If the user reports "Tally reconnected" or shares a screenshot from a parallel session
claiming a fix — **do not take that claim at face value**. Always independently re-verify (step 2)
before updating anything.

### 2. Extract fresh, genuinely-recalculated figures from the latest PL_BS_GL_Detail_*.xlsx
This file is a formula-only workbook maintained by a different process (the `fetch-pandl` skill or a parallel session) — if it was saved by a non-Excel writer, `openpyxl(data_only=True)` will return `None` everywhere, and even if Excel saved it, you still want a **forced full recalculation**, not stale cached values.

```powershell
cd "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS"
python ".claude\skills\company-review\scripts\read_pl_bs_recalc.py"
```

Before running:
- `tasklist //FI "IMAGENAME eq EXCEL.EXE"` — if Excel is already open **on this exact file** (check `tasklist //V` for the window title), **do not run this script**. Ask the user first; a second COM instance touching a file that's interactively open is exactly the collision pattern that has caused data loss before (OneDrive AutoSave silently overwriting a freshly regenerated file with a stale in-memory copy from an orphaned Excel process).
- If no Excel process, or a title-less/orphaned one, proceed.

The script opens with `CorruptLoad=1` (a plain `Workbooks.Open()` has failed with `0x800A03EC` on this file before; repair-mode load fixes it, but takes 60-90+ seconds — don't assume a timeout at 60s means failure, check the output file), forces `excel.CalculateFullRebuild()`, reads PL Summary / BS Summary / Ratio Analysis / MOM PL and BS / MOM BS GL Detail via bulk `Range.Value2` reads (never per-cell loops — that pitfall looked like a hang once and took an `iter_rows`-style bulk read to fix), and always closes with `wb.Close(SaveChanges=False)` + `excel.Quit()`. Output lands next to the script as `pl_bs_gl_detail_values.json`.

After it finishes: `tasklist //FI "IMAGENAME eq EXCEL.EXE"` again to confirm no orphaned process was left behind.

**Key figures to pull from the JSON and cross-check against whatever claim prompted the refresh:**
- `PL Summary` → Total Turnover / Cost of Sales / Gross Profit / Other Income / Operating Expenses / EBITDA / Finance Cost / PBT / PAT (Grand Total = last element of each row).
- `BS Summary` → `Reconciliation Gap (~0 expected)` — this is the row that matters most; if it isn't exactly (or near-exactly) zero across every column, the Balance Sheet does not actually reconcile yet, regardless of what anyone claims.
- `MOM BS GL Detail (TOTAL rows)` → use only rows where the label starts with `TOTAL:` — the same account name (e.g. "Sundry Creditors") also appears at ledger and sub-grouping level; only the `TOTAL:` row is the true group total.

### 3. Cross-check the other source files
- `GP REPORT.xlsx` — Q1 FY25-26 comparator Net Sales / Gross Profit by segment (Ink / Spare Parts / Heads).
- The two expense workbooks — Direct/Indirect totals and top categories, current quarter vs same quarter prior year.
- `Working_Capital_Review_Jul2026.xlsx` and/or `CFS PL MOM.xlsx` — Working Capital "At a Glance" figures (Inventory/Debtors/Creditors/Net WC) and the BS sheet for the Mar-Apr-May-Jun trend.
- `MACHINE SALES REPORT 25-26..xlsx` / `MACHINE SALES REPORT 26-27.xlsx` — Machine Sale revenue/units from the "Summary" pivot sheet, current quarter vs prior year.
- **Watch column mapping carefully** in `CFS PL MOM.xlsx`'s FINAL PL sheet — the true consolidated "CFS" total is a specific column (was column 32/AF), not the first similarly-named column (which may be a single-entity "O Tec" column). Misreading this produced a false "huge unexplained discrepancy" once; always check the header row's column labels before trusting a column index carried over from a prior run.

### 4. Update the literals in both build scripts
Both `build_board_pack.py` and `build_action_plan.py` hold the same figures as hardcoded literal tuples (`pl_rows`, `bs_rows`, `entity_pbt` / `company_pbt`, `ratio_rows`, `gap_rows` / `bs_trend`, `wc_bs_trend`, `gp_overall`/`gp_segment`, `exp_totals`/`exp_categories`, machine-sales rows). Update **both files identically** — they must never drift apart. Also update any prose that references specific numbers (Read Me First notes, BS-trend status commentary, Action Plan item text) — these are hand-authored and do not recompute themselves.

If an issue that was previously flagged as an open Action Plan item is now genuinely resolved (verified via step 2, not just claimed), **remove that item from the `action_plan` list** in both scripts rather than leaving a stale item pointing at a fixed problem.

### 5. Rebuild both workbooks

```powershell
python ".claude\skills\company-review\scripts\build_board_pack.py"
python ".claude\skills\company-review\scripts\build_action_plan.py"
```

Both print row markers on success — a mismatch or exception usually means a row-count assumption elsewhere in the script (merge ranges, print areas) needs adjusting.

### 6. Verify the rebuilt workbooks — read-only, never via Excel COM save
```python
import openpyxl
wb = openpyxl.load_workbook(r"...Files\Company_Review_Board_Pack_Q1_FY26-27.xlsx", data_only=False)
```
Read back the Source Data sheets and Action Plan list to confirm the new figures landed and any resolved item is actually gone. Never open these output files via `win32com` to "check" them — that risks the exact AutoSave-clobbering incident referenced in the `working-capital-review` skill. `openpyxl` read-only reads the saved formula text/literals directly and cannot be clobbered.

### 7. Update the HTML dashboard to match
`board_report_template.html` (copy to a scratch working file, edit, then publish via the Artifact tool) must mirror every number that changed in step 4: KPI tiles and their chips, the PBT-by-entity table, Key Ratios table, Balance Sheet trend card and its reconciliation-gap footnote, and the Action Plan cards (`ap-item p-high/p-medium/p-low`) — remove a card if its item was resolved, exactly as done in the workbooks.

Republish to the **same Artifact URL** each time (pass the same `file_path` within one conversation; pass `url` explicitly if resuming in a different conversation) so the user's saved link keeps working.

### 8. Report back — and never rubber-stamp a "fixed" claim
When a parallel session or the user reports a fix (e.g. "Balance Sheet gap resolved," "reconciled to the rupee"), always re-run step 2 against the actual current file before updating anything or giving a go-ahead to print/distribute. Two things to check every time:
- Does the Reconciliation Gap row actually read (near-)zero now, independently recalculated?
- Did the file's modified timestamp actually change since the last time you checked it, and is there an open Excel window on it right now (check `tasklist //V`) that means a further edit could still be in flight?

Report the specific before/after numbers, not just "confirmed" — this project's standing expectation is transparency about what was actually verified versus what was merely claimed.

## Known caveats to keep carrying forward

- Source Data sheets are a **snapshot**, not live links — if any of the seven-plus source files changes, someone must re-run this whole process; the workbook will not pick it up on its own.
- The Mar-26 point in the Balance Sheet trend comes from `CFS PL MOM.xlsx` (O Tec + Enterprises only, excludes Colorix) while Apr/May/Jun-26 come from `PL_BS_GL_Detail.xlsx` (all five companies) — a genuine scope change mid-trend, always disclosed rather than smoothed over.
- A true Q1 FY25-26 Balance Sheet, and several ratios needing full current-asset/liability classification for FY25-26, have not been cleanly available — marked "n/a" rather than estimated. Revisit each quarter in case the data becomes available.
