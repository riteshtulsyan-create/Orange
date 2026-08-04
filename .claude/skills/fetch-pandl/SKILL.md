---
name: fetch-pandl
description: Fetch P&L / Balance Sheet data from Tally Prime for all 5 Orange O Tec / Colorix companies and rebuild the consolidated PL_BS_GL_Detail.xlsx workbook (Schedule + Summary + Ratio Analysis + Month-on-Month + MOM Ratio Analysis + Cross Check sheets). Use this skill whenever the user asks to fetch PnL, refresh the PL and BS from Tally, update the PL BS file, rebuild the GL detail workbook, cross-check/verify the workbook, or asks about ratios (monthly or quarterly) — and by standing default, run its full procedure (live refresh, both quarter and monthly, all 5 companies) before delivering any final report or figures from this workbook, unless the user has explicitly scoped the request narrower for that ask.
---

# Fetch P&L / BS from Tally Prime and build the consolidated workbook

This skill produces **`Files/PL_BS_GL_Detail_<timestamp>.xlsx`** (see "File naming and version
history" below — there is no fixed filename) — a 12-sheet workbook consolidating Profit & Loss
and Balance Sheet data for 5 companies for Q1 FY26-27 (01-Apr-2026 to 30-Jun-2026), with full
GL-level drill-down, a management summary, ratio analysis (both quarterly and monthly — see
"Ratio Analysis structure" below), a month-on-month (MOM) breakdown, and a live formula-based
Cross Check sheet. Every refresh also produces a companion **`Profit_Change_Reconciliation_<same
timestamp>.xlsx`** showing what changed since the previous refresh and why (see Step E below). It
also documents `fetch_pandl.py`, a separate, simpler single-report utility kept for ad-hoc use.

## File naming and version history (2026-07-25 onward, permanent requirement)

**There is no fixed output filename.** Every refresh writes a brand-new, timestamped file rather
than overwriting the previous one in place — this was an explicit user request, so they can always
tell exactly when a given version was generated and can go back to any prior version if needed.

- **Main workbook**: `Files\PL_BS_GL_Detail_<YYYY-MM-DD_HHMM>.xlsx` (e.g.
  `PL_BS_GL_Detail_2026-07-25_1430.xlsx`) — the timestamp is when the rebuild ran, generated once
  and passed as `--timestamp` to both build scripts so they write to the exact same file (see "How
  to run").
- **Change report**: `Files\Profit_Change_Reconciliation_<same timestamp>.xlsx` — always shares
  its main workbook's timestamp, so the pair is easy to match up.
- **"The current workbook"** always means whichever `PL_BS_GL_Detail_*.xlsx` is newest directly
  inside `Files\` (not inside the archive folder below). `find_latest_pl_bs_file()` in
  `build_gl_detail_sidebyside.py` (and the equivalent glob in
  `company-review/scripts/read_pl_bs_recalc.py`) does this lookup — never hardcode the old fixed
  filename anywhere new.
- **Archive folder**: `Files\Old Version of PL BS\`. Once a refresh's new pair of dated files is
  built, verified (Step C), and the change report generated (Step E), the *previous* pair (the
  ones the new refresh just superseded) gets moved here — see Step F. `Files\` itself should
  therefore normally hold only the single newest dated pair; history lives in the archive folder,
  named exactly as it was when current, so you can always tell when any given version was made.
  On 2026-07-25 this folder was also seeded with every pre-existing manual backup that predated
  this mechanism (`PL_BS_GL_Detail Old.xlsx`, `PL_BS_GL_Detail v1 20 July 26.xlsx`,
  `PL_BS_GL_Detail.BACKUP_before_edit.xlsx`, and the genuine 22-Jul pre-refresh version renamed to
  `PL_BS_GL_Detail_2026-07-22_pre-refresh.xlsx`), at the user's request, so all historical versions
  live in one place. These older ones don't follow the `_<YYYY-MM-DD_HHMM>` pattern exactly and
  that's fine — nothing programmatically scans inside the archive folder, so their names don't need
  to be machine-parseable, only human-readable.
- The tiny `PL_BS_GL_Detail.grouping_snapshot.json` sidecar (Excel column/row grouping handoff
  between the two build scripts, see "Tab order and manual Excel grouping" below) deliberately
  keeps a **fixed** name regardless of timestamp — it's internal handoff state for a single build,
  not a version to preserve, and giving it a timestamp too would just leave one orphaned file per
  refresh forever.

## Period structure — Q1 kept, Jul-26 and Apr-Jul Total appended (2026-08-03 onward)

The MOM sheets carry **six** P&L period columns and **four** BS period columns:

| Statement | Columns |
|---|---|
| P&L | `Apr-26` · `May-26` · `Jun-26` · **`Q1 Total`** · `Jul-26` · **`Apr-Jul Total`** |
| BS | `Apr-26` · `May-26` · `Jun-26` · `Jul-26` |

`Q1 Total` (Apr-Jun) is deliberately preserved so earlier reports stay comparable — Jul-26 and the
cumulative Apr-Jul Total were **appended**, not folded into it (user decision 2026-08-03). The
Balance Sheet gets no cumulative column because it is point-in-time: the Jul-26 column already
*is* the as-at-31-Jul position.

Both cumulative P&L columns are **real fetches**, not derived sums — `qtr` reads
`gl_data/pl_<KEY>.csv`, `ytd` reads `gl_data/pl_ytd_<KEY>.csv`. Keeping `ytd` as its own pull is
what lets Sum(Apr:Jul) be checked against it as a genuine reconciliation rather than a tautology.

**Adding a further month (Aug-26, and eventually a Q2 Total) is now a small change**: add the
period to `PERIODS` in `fetch_gl_data.py`, add it to `PL_PERIODS`/`BS_PERIODS` and `MONTH_DAYS` in
`build_mom_gl_detail.py`, and give `load_period_rows` its branch. Everything downstream —column
geometry, ratio blocks, summaries— derives from `len(periods)` and `period_cols(idx)`, so no
column arithmetic needs touching.

**Off-ledger adjustments are as-at BALANCES, not flows.** A month's own P&L column therefore gets
the *increment* over the prior cumulative balance (`ytd - qtr` for July), never the full balance —
that is what keeps Apr+May+Jun equal to Q1 Total *and* Apr+May+Jun+Jul equal to Apr-Jul Total at
the same time. Handing July the whole 31-Jul balance would double-count the quarter's share.

**Closed 2026-08-04**: `Cross Check` now carries the Apr-Jul leg as well - columns
`Jul-26` / `Sum (Apr:Jul)` / `MOM Apr-Jul Total` / `Var: Sum(Apr:Jul) vs Apr-Jul Total`, with the
same yellow >Rs.10 highlighting as the other two variance columns. Because `ytd` is fetched in its
own right rather than derived as Q1+Jul, this is a genuine reconciliation and not a tautology. All
three variance columns read 0.00 across 228 cells as of the 2026-08-04_1401 build.

## "% of Sales" columns and the Valuation bridge — on every P&L sheet (2026-08-04 onward)

**`% of Sales`** columns are generated by the build scripts, never added by hand. Three on the MOM
sheets (`Q1 Total` / `Jul-26` / `Apr-Jul Total`, appended right of the last period block via
`add_pct_of_sales_columns()` in `build_mom_gl_detail.py`) and one Grand-Total column on the
quarter sheets (`add_grand_total_pct_column()` in `build_gl_detail_sidebyside.py`).
Coverage as of 2026-08-04 is complete: `PL GL Detail` and `PL Summary` carry one column each;
`MOM PL GL Detail`, `MOM PL and BS` and `MOM PL and BS - By Company` carry three each (the
By-Company sheet computes its percentages off the Grand Total block only, to avoid adding 24
columns to an already very wide sheet). The Valuation bridge is likewise on all five P&L sheets. Only the three
reporting periods get one — a column per month would double the sheet width for little gain.
Every cell is `=IFERROR(<GT>row/<GT>$turnover_row,"")`, so it can never disagree with the figure
beside it. **These were originally added by hand in Excel and were silently destroyed on the next
rebuild** — both schedules are regenerated from scratch every run. If a percentage column is ever
wanted somewhere new, add it to these helpers, never to the saved workbook.

**The Valuation-related Provisions bridge now exists on the MOM sheets too**
(`write_mom_valuation_bridge()`), not just `PL GL Detail` / `PL Summary`. Same selection rule —
whatever Tally files directly under `Direct Expenses` with no finer Sub Grouping — so the
month-on-month view can finally show how much of each period's profit rests on those items.

## Balance Sheet related-party / branch memo blocks (2026-08-04 onward)

The Balance Sheet carries two disclosure blocks below the Reconciliation Gap, mirroring the P&L's
Branch / Related Party split: **BRANCH BALANCES** and **RELATED PARTY BALANCES**. They are memo
only — the same ledgers stay counted in Total Assets / Total Liabilities, so the Gap is unaffected
(verified: adding them left all 48 gap cells unchanged).

**Detection requires BOTH tests — a sub-group-name rule alone misses about a third of them.**
Measured on live data 2026-08-04: 16 ledgers sit under sub-groups containing "RELATED"
(`BALANCE WITH RELATED PARTY`, `CREDITORS FOR RELATED`, `BALANCE WITH RELATED PARTY(Debtors)`),
but another 11 sit in ordinary groups — `Unsecured Loans`, `Loans & Advances (Asset)`,
`Sundry Debtors`, `Current Assets`, `Branch / Divisions` — and are identifiable only because the
LEDGER NAME carries a group company's name (e.g. `ORANGE O TEC PRIVATE LIMITED-USL-SURAT` filed
under plain `Unsecured Loans`). `classify_bs_memo()` applies both tests. Branch is kept separate
because the intra-group Surat↔Noida pair is ~₹58cr each way and would otherwise swamp the
genuine related-party balances.

## Adjustment constants go stale — re-read them every refresh

`--computed-lines` values are a snapshot, not settled facts; Tally's own computed lines move as
entries are posted. Between 2026-08-03 and 2026-08-04 Orange O Tec's Unadjusted Forex Gain/Loss
fell 77,750.00 → 55,000.00 for the quarter and the 31-Jul line was **cleared entirely**. The stale
constants produced Reconciliation Gaps of exactly the drift — 22,750.00 on the quarter sheets and
2,48,118.00 on the MOM Balance Sheet. **That is the signature to recognise: when a Gap equals the
difference between a constant and its current live value, the constant is stale, not the data.**
Re-run `--computed-lines` and update the dicts as part of every refresh, before blaming the fetch.

## `Related Party` sheet — the inter-company reconciliation (2026-08-04 onward)

A standalone sheet (`write_related_party_sheet()` in `build_mom_gl_detail.py`) gathering every
inter-company item, because a genuine group elimination **nets to zero** and that is worth being
able to see in one place. Five blocks:

1. **P&L related-party flows** per period, with a NET row.
2. **Balance Sheet related-party balances** per month-end, per company, with a NET row.
3. **Branch balances**, same treatment, kept separate.
4. **Counterparty matrix** — rows are the company whose books hold the balance, columns the
   counterparty it is booked against, plus a PAIR CHECK (A-vs-B + B-vs-A = 0) that names the two
   companies whose records disagree.
5. **Ledger detail** — every inter-company ledger with its mapped counterparty, by month.

Detail cells carry values; every NET / PAIR CHECK row is a live formula, amber above ₹10.

**The matrix covers RELATED only, never Branch.** Surat and Noida are the same legal entity in two
Tally files, and a branch ledger named after its *own* company (e.g. `ORANGE O TEC PVT.
LTD.(SURAT BRANCH)` sitting in the Surat books) maps to itself, so one leg of the pair drops out
and the check reports a spurious multi-crore break. Branch is tested by its own NET row instead.

`counterparty_of()` (in `build_gl_detail_sidebyside.py`) maps a ledger to the group company it
represents. **Order matters**: `ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA` contains all of
ORANGE O TEC / ENTERPRISES / NOIDA, so Colorix is tested first, then Enterprises with its own
NOIDA qualifier, then plain Orange O Tec. SURAT means head office and needs no rule — it falls
through to the right default.

**What it found on day one** (2026-08-04_1504 build): the P&L eliminates cleanly (₹9.97cr of
related-party sales against the same purchases, netting to ₹0.80 across Apr-Jul), but the Balance
Sheet does not. About **₹40.26 lakh is a counterparty MISCLASSIFICATION** rather than missing
money — booked against Enterprises in one set of books and Enterprises-Noida in the other, so the
pair rows are equal and opposite and cancel across the two. On top of that sit genuine residuals:
**₹4.48 lakh (May)**, **₹16,886 (Jul)** and **₹11 lakh of branch** in May/Jun/Jul.

**New sheets must be added to `desired_order` in `build_mom_gl_detail.py main()`** — anything not
listed there is silently dropped by the `wb._sheets` reassignment, with no error.

## Ratio Analysis structure — always both quarterly AND monthly (permanent requirement)

Two ratio sheets always exist, built from the **same shared ratio definitions**
(`ratio_definitions()` in `build_gl_detail_sidebyside.py`, imported by `build_mom_gl_detail.py`
as `b.ratio_definitions()`) so the two views can never drift out of sync with each other:

- **`Ratio Analysis`** — quarterly (Q1 Total), one column per company/family, built by
  `write_ratio_sheet()` in `build_gl_detail_sidebyside.py`, sourced from `PL GL Detail` /
  `BS GL Detail`. 91-day quarter used for day-count ratios (Debtor Days, Creditor Days).
- **`MOM Ratio Analysis`** — monthly (Apr-26 / May-26 / Jun-26, each computed independently
  from that month's own column — never a slice or average of the quarter), one period-block of
  8 company columns per month, built by `write_mom_ratio_sheet()` in `build_mom_gl_detail.py`,
  sourced from `MOM PL GL Detail` / `MOM BS GL Detail`. Each month's own actual day-count is used
  for Debtor/Creditor Days (30 for Apr, 31 for May, 30 for Jun — see `MONTH_DAYS`), not
  annualised. Q1 Total is deliberately **not** repeated here since `Ratio Analysis` already
  covers it — this sheet is Apr/May/Jun only.

**If a new ratio is ever added, add it to `ratio_definitions()` once** — both sheets will pick it
up automatically on the next rebuild. Never hand-add a ratio to only one of the two sheets. If a
new ratio needs a raw figure not already in `PL_COMPONENTS`/`BS_COMPONENTS`, add it there once too
— it becomes a new Component Figures row on both sheets automatically.

### Ratio Workings — every ratio sheet is auditable, not a black box (permanent requirement)

Both ratio sheets are laid out in **three stacked sections, top to bottom**, so a human can trace
any number by eye without opening another sheet (this was an explicit user request — "bring the
figures in the sheet below and calculate the formula and link the final output at the above"):

1. **RATIO SUMMARY** (top) — the headline numbers. Every cell is a same-sheet link down to its
   row in section 3 (e.g. `=I61`) — **never** a formula computed directly, so the top number and
   the workings below it can never show different figures.
2. **RATIO WORKINGS — COMPONENT FIGURES** (middle) — every raw P&L/BS figure any ratio needs
   (Gross Profit, Total Turnover, PAT, Total Assets, Sundry Debtors, ...), pulled **once** per
   company column directly from the GL Detail schedule (`='PL GL Detail'!E23`-style plain link —
   no arithmetic). One row per component, listed in `PL_COMPONENTS` / `BS_COMPONENTS`.
3. **RATIO WORKINGS — CALCULATIONS** (bottom) — the actual ratio formulas (still `IFERROR`-wrapped),
   built **only** from same-sheet references into section 2 (e.g. `=IFERROR(C33/C30*100,0)`) —
   never a cross-sheet reference. This is the row the Summary section links to.

Implemented via shared helpers in `build_gl_detail_sidebyside.py`, reused by both scripts:
`ratio_sheet_layout()` (precomputes row numbers for all three sections before any cell is
written, since Summary needs to know the future Calculations row to build its link),
`write_ratio_summary_section()`, and `write_ratio_workings_section()`. The monthly sheet calls
these once per period block (Apr/May/Jun), passing that period's own day-count and source columns
— see the `columns` list construction in `write_mom_ratio_sheet()`.

**Do not shortcut this** by writing a ratio's full cross-sheet formula directly into the Summary
row again (the pre-this-structure design) — it defeats the entire point (nothing to visually
cross-check against) and reintroduces the risk in the next paragraph.

**Column-reference bug to watch for** (root cause of a real bug hit once, keep watching for it):
`MOM Ratio Analysis` has its *own* column layout (label column + 3 period-blocks of 8, via
`ratio_period_cols()`/`RATIO_LABEL_COLS=1`) which is completely different from the *schedule*
sheets' column layout (`MOM PL GL Detail`/`MOM BS GL Detail`, via the module-level
`period_cols()`/`LABEL_COLS=4`). Any formula that writes to a cell on the ratio sheet but reads
from a cell on the schedule sheet needs **both** column mappings at once — `ratio_period_cols(pidx)`
for where to place the value on this sheet, `period_cols(pidx)` for which schedule column to
reference in the Component Figures formula. Using one mapping for both (an easy mistake — they're
both 8-key dicts with identical key names) produces formulas that silently read the *wrong
company's* column and compute nonsense (caught once this way: April Current Ratio came out as
-10.27 instead of a sane ~1.4). Always verify a freshly-added monthly/period-block sheet's
formulas via the `formulas` engine before trusting the numbers — see item 0 in "Verification
checklist" below. With the Ratio Workings structure, this is easier to catch: cross-check that
each period's Component Figures row (e.g. row 31, Total Turnover) matches that same month's
figure on the `MOM PL GL Detail` schedule sheet directly, for at least one company column.

## Tab order and manual Excel grouping — always preserved across rebuilds (permanent requirement)

**Sheet tab order** is fixed and reapplied on every rebuild (set at the end of
`build_mom_gl_detail.py main()`, since that script runs last): MOM sheets first (`MOM PL GL
Detail`, `MOM BS GL Detail`, `MOM PL and BS`, `MOM PL and BS - By Company`, `MOM Ratio Analysis`),
then `Ratio Analysis` (quarterly), then the base sheets (`PL Summary`, `BS Summary`, `PL GL
Detail`, `BS GL Detail`, `Notes`), then `Cross Check` last as a diagnostic. Do not reorder tabs by
hand in Excel expecting it to stick — change `desired_order` in `build_mom_gl_detail.py` instead.

**Manual column/row grouping** (Excel's Data > Group, which collapses columns/rows and marks them
`hidden` in the file) is a display preference the user sets up by hand and expects to survive
every future rebuild, even though both scripts regenerate their sheets from scratch each run.
This is handled by `snapshot_column_row_hidden()` / `restore_column_row_hidden()` in
`build_gl_detail_sidebyside.py`, plus a JSON sidecar (`Files/PL_BS_GL_Detail.grouping_snapshot.json`)
because `build_gl_detail_sidebyside.py`'s own rebuild (`openpyxl.Workbook()` from scratch) wipes
sheets it doesn't touch (the MOM sheets + Cross Check) before `build_mom_gl_detail.py` even runs —
an in-memory-only snapshot inside one script's `main()` cannot survive that. The pattern:
`build_gl_detail_sidebyside.py` snapshots the file as it stood before its own overwrite, restores
what it can (the 6 sheets it rebuilds), and always also persists the full snapshot to the JSON
sidecar; `build_mom_gl_detail.py` loads that sidecar (not an in-memory snapshot of its own,
which would already be too late) and restores hidden state onto the MOM sheets + Cross Check it
just recreated. **Never remove this wiring** when touching either script's `main()` — a rebuild
that silently resets a user's collapsed columns will look fine in code review and only surface as
"why did my grouping disappear" days later.

## Standing procedure — CFO mode (default behavior, not a one-off checklist)

**By default, always run the full procedure below (Steps A-D) before submitting any final
report, figure, or conclusion drawn from this workbook** — not just when explicitly asked to
refresh. This mirrors how a CFO operates: numbers get re-verified against the live source before
they go out the door, every time, not only when someone happens to ask. The one exception: if
the user's own message for that turn explicitly scopes the ask narrower (e.g. "just explain what
this row means," "don't refresh, use what's there," "quick question, no need to pull fresh
data"), skip straight to answering from the current file — but say so, so the scope-down is
visible rather than silent.

**A. Refresh** — one command, all 5 companies, all periods (~50 seconds):
```powershell
python ".claude\skills\fetch-pandl\scripts\fetch_gl_data.py" --all
```
This writes both the cumulative files (`gl_data/`) and the monthly detail (`gl_data_mom_detail/`).
**Prefer this over the MCP call sequence documented further below** — the tally-prime MCP server
is not connected in every session, and when it is missing that procedure is simply not runnable.
`fetch_gl_data.py` talks to Tally's XML gateway directly and needs no MCP. See "Refreshing the
CSVs without MCP" below for its validation story and the two traps it encodes.

**B. Rebuild** — generate one shared timestamp and pass it to both scripts so they write to the
same new dated file (see "File naming and version history" above and "How to run" below):
```powershell
$ts = Get-Date -Format "yyyy-MM-dd_HHmm"
python ".claude\skills\fetch-pandl\scripts\build_gl_detail_sidebyside.py" --timestamp $ts
python ".claude\skills\fetch-pandl\scripts\build_mom_gl_detail.py" --timestamp $ts
```
This does **not** touch or overwrite the previous refresh's dated file — both now sit side by side
in `Files\` until Step F archives the older one.

**C. Verify via the Cross Check sheet** — open the rebuilt workbook's `Cross Check` sheet (built
automatically by `build_mom_gl_detail.py`) and confirm, via the real formula-calculation engine
(never openpyxl's cached values — see "Known bugs found & fixed" item 3 below):
  - **"Var: MOM Q1 Total vs PL/BS Summary" must be ~0 everywhere** (consolidated, all 8
    company/family columns, P&L and BS). This is a same-source-data check — both pull from the
    identical live quarter CSVs, so any non-zero value here is a real bug, not a data-vintage
    difference. Investigate and fix before reporting anything.
  - **"Var: Sum(Apr:Jun) vs MOM Q1 Total" should also be ~0** once Step A has refreshed the
    monthly CSVs too (both checks draw from the same live pull now, so nothing should be stale).
    If it is not ~0 after a full refresh, that is a new, real gap — investigate rather than
    reporting it as an already-known/accepted caveat.

**D. Report** — only after C passes (or with any residual gap explicitly explained, root-caused,
and quantified) deliver the final figures/report to the user.

**E. Generate the change-reconciliation report** — run `build_gl_reconciliation_report.py` with no
arguments; it auto-discovers the two newest `PL_BS_GL_Detail_*.xlsx` files still in `Files\` (the
one Step B just built, and whichever it superseded) and writes a matching dated
`Profit_Change_Reconciliation_<timestamp>.xlsx`:
```powershell
python ".claude\skills\fetch-pandl\scripts\build_gl_reconciliation_report.py"
```
This bridges Turnover/Gross Profit/PAT old-vs-new per company, with the specific GL ledger(s)
driving each swing auto-identified (not hand-picked), plus a full GL-level
Added/Removed/Changed/Renamed audit for every company, for both P&L and Balance Sheet. This is a
**standing, permanent deliverable — run it every refresh**, not an ad-hoc report; the user asked
for this after the 2026-07-25 refresh showed an unexplained ₹1.81cr profit swing and wants it
available every time going forward. If fewer than 2 dated files exist yet (very first-ever run),
the script prints a message and exits cleanly without producing a report. Mention the headline PAT
movement (and its top driver, if material) in the report to the user whenever this file is
produced, the same way Cross Check gets mentioned.

**F. Archive the superseded version** — only after E has produced its report (so nothing needed
for the diff goes missing), move the *previous* refresh's dated pair — the older
`PL_BS_GL_Detail_*.xlsx` and its matching `Profit_Change_Reconciliation_*.xlsx` (NOT the ones Step
B/E just created) — into `Files\Old Version of PL BS\`, creating that folder if it doesn't exist
yet:
```powershell
New-Item -ItemType Directory -Force "Files\Old Version of PL BS" | Out-Null
Move-Item "Files\PL_BS_GL_Detail_<old-timestamp>.xlsx" "Files\Old Version of PL BS\" -Force
Move-Item "Files\Profit_Change_Reconciliation_<old-timestamp>.xlsx" "Files\Old Version of PL BS\" -Force
```
After this, `Files\` should hold exactly one current dated pair (this refresh's), with every
earlier pair sitting in the archive folder under its own original timestamped name — never delete
an archived version, and never rename it once archived.


## The 5 companies (exact Tally names — use these, not shortened versions)

| Short key | Exact Tally company name |
|---|---|
| `ORANGE_O_TEC_PRIVATE_LIMITED` | `ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)` |
| `PRIVATE_LIMITED_NOIDA` | `ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)` |
| `ENTERPRISES_PVT_LTD` | `ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)` |
| `ENTERPRISES_NOIDA` | `ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27` |
| `COLORIX` | `COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)` |

These pair into two "families" for the workbook's Total columns: Orange O Tec (CO1+CO2) and
Enterprises (CO3+CO4), with Colorix (CO5) standing alone. Grand Total = both families + Colorix.

**⚠️ Critical: always use the FULL exact name above with `set-company`, never a shortened
version.** A shortened name (e.g. `"COLORIX DIGITAL PRINTING SOLUTIONS LLP"` without the
`- (from 1-Apr-20)` suffix) can silently resolve to the wrong company in Tally's MCP server —
this happened once and returned an entire unrelated company's chart of accounts (1000+ ledgers)
without any error, because there are similarly-named company variants open at the same time
(e.g. multiple FY-year variants of Enterprises). Always double check `list-master` (collection
`company`) if in doubt.

## Directory layout (all inside this skill folder — NOT a session scratchpad)

```
.claude/skills/fetch-pandl/
  scripts/
    build_gl_detail_sidebyside.py   <- builds the main workbook (PL/BS Summary, GL Detail, Ratios, Notes)
    build_mom_gl_detail.py          <- builds the 4 MOM sheets on top of the same workbook
    build_gl_reconciliation_report.py <- Step E: old-vs-new change reconciliation report (see below)
    fetch_pandl.py                  <- separate, simpler single-report fetch (see below)
    archive/                        <- superseded early-iteration scripts, do not run
  gl_data/                          <- one pl_<KEY>.csv + bs_<KEY>.csv per company, Q1 totals
  gl_data_mom_detail/               <- pl_apr/may/jun_<KEY>.csv + bs_apr/may_<KEY>.csv per company
```

`Files\Profit_Change_Reconciliation_<timestamp>.xlsx` (Step E's output) lives in `Files\` alongside
the main workbook, not inside the skill folder — it's a user-facing deliverable like the main
workbook itself, not internal working state. Archived prior versions of both live in
`Files\Old Version of PL BS\` (Step F) — see "File naming and version history" above.

**This was previously a bug**: `gl_data/` and `gl_data_mom_detail/` used to live under a
session-specific Temp scratchpad path (tied to one particular conversation's session ID), so a
future session with a different ID could not find the data at all. Fixed 2026-07-18 — both
directories now live permanently inside the skill folder itself and both scripts read from
`Path(__file__).parent.parent / "gl_data"` (or `gl_data_mom_detail`), so they work from any
session, forever, as long as the files stay here.

## CSV format

Every CSV has exactly 4 columns: `primary_group,group_name,ledger_name,closing_balance`.
- `primary_group` comes from the `chart-of-accounts` MCP tool (the Tally group hierarchy root,
  e.g. `Sundry Debtors`, `Purchase Accounts`) — this is what the build script uses to decide
  which section/bucket a row belongs to.
- `group_name` comes from the `profit-loss`/`balance-sheet` MCP tool's own report output (the
  immediate parent group as Tally displays it, which can be more specific, e.g.
  `RELATED PARTY PURCHASE`) — used for finer classification (Related Party detection, Finance
  Cost detection) and shown in the workbook's "Sub Group" column.
- Sign convention throughout: **assets and expenses are negative, liabilities and income are
  positive** (Tally native convention). A balanced Balance Sheet always sums to exactly 0 across
  every row (Liabilities + Assets, since Assets carry a negative sign).

## Refreshing the CSVs without MCP — `fetch_gl_data.py` (preferred, 2026-08-03 onward)

`scripts/fetch_gl_data.py` reproduces the whole MCP fetch over Tally's plain XML gateway. Built
and validated 2026-08-03 against the then-current known-good CSVs; read its module docstring
before changing anything. Modes:

```powershell
python ".claude\skills\fetch-pandl\scripts\fetch_gl_data.py" --all              # full refresh
python ".claude\skills\fetch-pandl\scripts\fetch_gl_data.py" --validate         # diff vs existing CSVs
python ".claude\skills\fetch-pandl\scripts\fetch_gl_data.py" --computed-lines   # adjustment constants
python ".claude\skills\fetch-pandl\scripts\fetch_gl_data.py" --stock            # closing stock vs STOCK_MOM
```

**Two hard-won rules baked into it:**

1. **`ISLEDGERWISE=Yes` is what produces ledger-level rows.** `EXPLODEFLAG`,
   `SVEXPANDALLLEVELS`, `SVDETAILED` and a `REPORT ISMODIFY` were each tried and each silently
   returned only the ~32 group totals. Trial Balance with `ISLEDGERWISE` is also *fast*: ~0.3s per
   company/period.
2. **Never request balances via a `Ledger` COLLECTION.** `<TYPE>Ledger</TYPE>` with a
   `ClosingBalance` NATIVEMETHOD makes Tally walk every voucher from the company's book start. One
   such call for Colorix — the smallest company — ran **over 20 minutes at 100% CPU without
   returning**, and since Tally serialises XML on one thread it blocked every later request until
   the process was restarted. Master data (names/parents, no balances) is cheap and fine.

**Semantics it reconciles** (each was a real bug caught by `--validate`):

- `primary_group` = nearest **reserved** ancestor group, by walking the Parent chain. Tally's own
  `$PrimaryGroup` returns empty over XML.
- Trial Balance reports the **debit already negative**, so a ledger's balance is `dr + cr`, not
  `cr - dr`.
- `DSPACCNAME`/`DSPACCINFO` are **flat siblings**, so they must be paired positionally. Pairing
  via `find()` on the shared parent hands every ledger the document's first balance.
- The Trial Balance's `Profit & Loss A/c` is only the **opening balance brought forward**; the
  Balance Sheet view adds the period's profit. The row is emitted **even when Tally has no such
  ledger** — otherwise a company with no brought-forward balance silently loses its entire period
  profit from the Balance Sheet.
- Tally's computed BS lines (`Sales Bills to Make`, `Purchase Bills to Come`, `Unadjusted Forex
  Gain/Loss`) are **excluded** from the CSVs — they are not ledgers, carry no parent group, and
  `load_company_rows()` already injects them into the right bucket on both statements.

## How to refresh one company's data from Tally (step by step)

1. **`set-company`** with the exact full name from the table above.
2. **`set-period`** with `fromDate`/`toDate` for the period you need (e.g.
   `2026-04-01`/`2026-06-30` for the quarter). **Do this even though `profit-loss` and
   `balance-sheet` also accept `fromDate`/`toDate` parameters** — without an explicit
   `set-period` call first, those tools can return data for the wrong/default period (this
   caused a real false alarm once: a `balance-sheet` call without `set-period` returned a full
   company-history dump instead of the intended quarter).
3. **`chart-of-accounts`** — returns a `tableID` (cached ~15 min).
4. **`profit-loss`** and/or **`balance-sheet`** with the same `fromDate`/`toDate` — each returns
   its own `tableID`.
5. **`query-database`** with a `LEFT JOIN` of the report table to the chart-of-accounts table on
   `ledger_name`, `WHERE closing_balance <> 0`, `outputFormat: "CSV"`:
   ```sql
   SELECT c.primary_group, p.group_name, p.ledger_name, p.closing_balance
   FROM <profit_loss_or_balance_sheet_tableID> p
   LEFT JOIN <chart_of_accounts_tableID> c ON p.ledger_name = c.ledger_name
   WHERE p.closing_balance <> 0
   ORDER BY p.ledger_name
   ```
6. Save the CSV output to `gl_data/pl_<KEY>.csv` or `gl_data/bs_<KEY>.csv` (or the
   `gl_data_mom_detail/` equivalents for a specific month — see naming convention below).

### Known Tally API data-quality quirks (both handled automatically by `load_csv()`)

1. **Duplicate ledger rows**: for some date ranges, the same ledger (e.g. `Profit & Loss A/c`)
   appears twice under two different `group_name` values with the same balance.
   `load_csv()` keeps only the first occurrence of any `ledger_name`.
2. **Native `Stock-in-Hand` rows**: some companies/periods return `Opening Stock`/`Closing
   Stock` rows (blank `primary_group`) that duplicate this project's own stock figures (see
   below). `load_csv()` skips any row with `group_name == "Stock-in-Hand"`.

If you hand-write or hand-edit a CSV outside the MCP flow, keep both guards in mind — or better,
just don't worry about it, since `load_csv()` filters both cases on every load regardless of
source.

## MOM (month-on-month) file naming convention

Inside `gl_data_mom_detail/`:
- `pl_apr_<KEY>.csv`, `pl_may_<KEY>.csv` — that month's own P&L flow only (fromDate/toDate
  scoped to the month).
- `pl_jun_<KEY>.csv` — June's own P&L flow only (**not** cumulative; the Q1 Total column reuses
  `gl_data/pl_<KEY>.csv` directly, it does not sum apr+may+jun from these files).
- `bs_apr_<KEY>.csv`, `bs_may_<KEY>.csv` — Balance Sheet **as at** that month's end (fromDate
  should be 2026-04-01, i.e. from FY start, since BS is cumulative/point-in-time, not a flow).
  June's BS column reuses `gl_data/bs_<KEY>.csv` directly (same reasoning as above).

## Stock, and the manually-confirmed adjustments

### Stock must also land in retained earnings, not just in assets (fixed 2026-08-03)

Because this project replaces Tally's inventory figures with its own, it injects `Closing Stock
(Inventory)` as an **asset** — but Tally's real Balance Sheet also carries the matching stock
movement *inside* retained earnings (its `Profit & Loss A/c` line includes `+closing − opening`).
Injecting only the asset leaves the Balance Sheet short by exactly that movement, and every
company/month was out by precisely `closing_stock − opening_stock` until
`apply_stock_to_pl_ac()` (in `build_gl_detail_sidebyside.py`) was added to both load paths. The
opening leg is always the **FY-start** opening, never the prior month's close, because the
Balance Sheet is cumulative. After this fix all 20 company-month Balance Sheets reconcile to 0.00.

⚠️ **July closing stock is still PROVISIONAL for two companies.** `fetch_gl_data.py --stock`
reproduces `STOCK_MOM` exactly for Orange O Tec NOIDA, Enterprises NOIDA and Colorix across every
confirmed month, but disagrees for **Orange O Tec Pvt Ltd** and **Enterprises Pvt Ltd** in *every*
month including confirmed ones (Apr gap ₹6.67cr and ₹61.9 lakh respectively, and the gap varies
month to month, so it cannot be derived). Those two use a different stock basis than "sum of
StockItem ClosingValue" — their July figures need the same source that produced the existing
values, and should not be overwritten from the API until that basis is identified.

Closing/Opening Stock is **not a ledger** in Tally — it comes from the stock/inventory
valuation report, not the P&L or BS ledger pull, so it never appears in the CSVs above. Both
build scripts inject it explicitly from hardcoded `STOCK` / `STOCK_MOM` dicts in
`build_gl_detail_sidebyside.py` (opening/closing per company for the quarter) and
`build_mom_gl_detail.py` (`STOCK_MOM`, per company per month). **If stock figures change, update
these dicts directly — there is no CSV for this.**

Three more items are **not ledgers or group masters either** — they never appear via any
chart-of-accounts/ledger/group API call, only on Tally's own report *screens* (Group Summary
drill-downs). All three were found by comparing this workbook's Balance Sheet Reconciliation Gap
row (should always be ~0) against Tally's native reports until the residual was traced to a
specific screen, then hardcoded as constants in `build_gl_detail_sidebyside.py`:

| Constant | Company | Quarter (30-Jun) Amount | What it is | Confirmed via | Statements affected |
|---|---|---|---|---|---|
| `SALES_BILLS_TO_MAKE` | Orange O Tec Pvt Ltd (all 3 periods); Enterprises Pvt Ltd (May only, ₹3,51,050) | ₹3,54,79,998.84 | Goods dispatched, sales invoice not yet raised (real receivable, computed from pending Delivery Notes) | BS → Current Assets → Sundry Debtors → Group Summary | Both P&L (adds income) and BS (adds Sundry Debtors) |
| `PURCHASE_BILLS_TO_COME` | Enterprises Pvt Ltd (quarter/June only - no Apr/May adjustment needed for this one) | ₹6,78,615 | Mirror image: goods/services received from a related party, purchase invoice not yet booked | BS → Current Liabilities → Sundry Creditors → Group Summary **and independently** confirmed on the Purchase Group Summary screen (same figure, reducing Purchase Accounts) | Both P&L (reduces Cost of Sales, raises PAT) and BS (reduces Sundry Creditors) |
| `UNADJUSTED_FOREX_GAIN_LOSS` | Colorix | ₹45,600 | Standalone forex adjustment line Tally shows directly in Sources of Funds (Liabilities side) | BS → Sources of Funds (own line) | BS only |
| `UNADJUSTED_FOREX_GAIN_LOSS_ASSET` | Orange O Tec Pvt Ltd | (May only, ₹12,032) | Same concept as above, but for a company where Tally shows it on the Assets side instead - a separate, smaller item from that company's own Sales Bills to Make | BS → Assets (own line) | BS only |

### Setting these constants — read them off Tally, don't back-solve (2026-08-03 onward)

`fetch_gl_data.py --computed-lines` prints Tally's own computed Balance Sheet lines per company
per period. **This supersedes the screenshot / back-solve procedure below for getting the
numbers** — every figure previously confirmed by drill-down came back from it to the paisa
(Orange O Tec apr/may/qtr, Enterprises may, Enterprises `PURCHASE_BILLS_TO_COME` qtr, Colorix
apr). The drill-down procedure still matters for deciding **which bucket and which side** an item
belongs to; the API gives the amount, not the treatment.

⚠️ **The report prints SIGNED values — never take the magnitude.** These items change direction
between periods: Colorix's Unadjusted Forex Gain/Loss is **+44,266.20 in April but −12,987.45 in
July** (a gain, then a loss). Storing July as a positive magnitude put it on the wrong side and
left a 25,974.90 Reconciliation Gap — exactly twice the figure, which is the signature of a
sign error. Conventions: `UNADJUSTED_FOREX_GAIN_LOSS` stores the **sign as shown**; the other
three dicts store the **positive magnitude** (the build code applies the negation itself).

**One constant was retired this way.** `UNADJUSTED_FOREX_GAIN_LOSS["COLORIX"]` held 45,600.00 for
`qtr`/`may`. Tally reports no computed line at all for those periods — the 45,600.00 is a **real
ledger**, `EXCHANGE RATE DIFF EXP ACCOUNTS` (Direct Expenses), which the old MCP queries silently
omitted and the constant had been back-solved to plug. `fetch_gl_data.py` picks the ledger up, so
the constant was double-counting; it was removed, and the 2026-08-03 reconciliation report shows
Colorix Gross Profit moving by exactly +45,600 as a result.

**If a new residual gap appears for any company**: don't guess the treatment. Ask for (or take)
a screenshot of Tally's Balance Sheet drill-down (click into the residual bucket → Group
Summary) — that's what resolved all of the above. Then check whether the identical figure *also*
appears on the P&L side (Purchase Group Summary / Sales Group Summary) before deciding whether it
needs a one-sided or two-sided adjustment — getting this wrong is easy: an earlier pass on this
same item briefly had the P&L sign backwards (added as a further expense instead of an expense
*reduction*), caught only by cross-checking PL Summary's profit figure against the BS's own
"Profit & Loss A/c" ledger value (see next section).

**These four constants are per-period, not single numbers** (`{"qtr": ..., "apr": ..., "may": ...}`
- see the dicts in `build_gl_detail_sidebyside.py`, just above `load_company_rows`). This matters:
these balances genuinely grow/shrink month to month (pending dispatches get invoiced, new ones
accrue), so the 30-Jun figure **cannot** be reused for 30-Apr/31-May without breaking that month's
own BS Reconciliation Gap (hit for real 2026-07-18 - Apr/May gaps of ~₹1.7cr/₹1.6cr, entirely
explained by this, only caught once the MOM BS Reconciliation Gap row was actually checked per
month rather than assumed fine from the Cross Check sheet alone - see "Verification checklist"
item 1). `"qtr"` feeds `load_company_rows()` (used by the quarter sheets and by June's MOM BS,
which reuses `load_company_rows` directly); `"apr"`/`"may"` feed `load_period_rows()` in
`build_mom_gl_detail.py`, which applies them to Sundry Debtors/Sundry Creditors/Unadjusted Forex
directly (not via `load_company_rows`, since the CSV source differs). June's own MOM **P&L** flow
still uses the quarter-end `"qtr"` constant as an approximation (no confirmed monthly P&L split
exists for these items) - this is a deliberate, disclosed limitation, unrelated to the BS-side fix.
**When a month's gap appears, don't guess the split** - it can usually be back-solved exactly from
the observed Reconciliation Gap itself (the gap size IS the missing adjustment, if these known
items are the only unexplained source for that company - true here since the quarter was already
fully reconciled with no residual), then cross-verified against an actual Tally drill-down
screenshot once available. That backsolve-then-verify approach is what surfaced
`UNADJUSTED_FOREX_GAIN_LOSS_ASSET` above - the back-solved total for Orange O Tec Pvt Ltd in May
was correct in aggregate but turned out to bundle two real, separately-named Tally items together,
only visible once the actual screen was checked.

**A back-solved figure can zero the Gap while sitting in the wrong bucket entirely - always
verify against a screen before trusting the label, not just the number.** Hit for real
2026-07-18: Enterprises Pvt Ltd's May gap (₹3,51,050) was initially back-solved and placed under
`PURCHASE_BILLS_TO_COME` (a Sundry Creditors/liability reduction) purely because that was the
company's known adjustment type from the quarter - and it DID zero the Gap. It was still wrong:
`Gap = Total Liabilities + Total Assets` sums *both* buckets together, so reducing a liability by
X has the exact same effect on the Gap as increasing an asset's magnitude by X - **arithmetic
correctness on the Gap total is not proof the item is in the right place.** The real discrepancy,
found by comparing Tally's own `trial-balance` API output group-by-group against our CSV (every
other group matched to the rupee), was entirely inside **Sundry Debtors**, not Sundry Creditors -
a `SALES_BILLS_TO_MAKE`-type item, confirmed on-screen at exactly ₹3,51,050 via Balance Sheet →
Sundry Debtors → Group Summary. Corrected by moving it from `PURCHASE_BILLS_TO_COME` to
`SALES_BILLS_TO_MAKE["ENTERPRISES_PVT_LTD"]["may"]`. **Practical takeaway**: once a month's Gap
hits zero, that only proves the *total* is right - if you want the *workbook's own line items* to
be individually correct (not just the bottom line), pull a Tally `trial-balance` for the full
company/period and diff it group-by-group against the CSV (see "How to root-cause a Balance Sheet
gap precisely" below) rather than stopping at "the Gap is zero, good enough."

## Valuation-related provisions bridge (permanent, 2026-07-31 onward)

At the bottom of both `PL GL Detail` and `PL Summary`, after `PAT (Net of Related Party
Transactions)`, a standalone section highlights whichever ledgers Tally files directly under
`Direct Expenses` with no finer Sub Grouping of their own - in practice this has meant
`PROVISION FOR IMPORT-FOC(PL)` and `INVENTORY OVERHEAD COST`. These are the single most volatile,
judgment-heavy line items quarter to quarter (they are consistently the top PAT drivers in every
`Profit_Change_Reconciliation` report), so they get called out as a standalone bridge rather than
being left buried inside the ordinary Cost of Sales detail above:

- **Valuation-related Provisions** - the item(s) themselves, plus their subtotal.
- **Profit Before Valuation** = `PAT (Net of Related Party Transactions)` **minus** this subtotal
  - i.e. what profit would be if these specific items had never been posted.
- **Profit After Valuation** = `PAT (Net of Related Party Transactions)` directly - i.e. the true,
  final reported figure, unchanged by this bridge (it's a lens on PAT, not a new PAT).

Implemented in `write_pl_sheet()` (`build_gl_detail_sidebyside.py`) via `write_bucket_block()` on
a filtered slice of the same section data already loaded for Cost of Sales (`sections["COST OF
SALES"]["Direct Expenses"]` rows where `sub_grp == "Direct Expenses"` literally) - genuine live
formulas, not literals, and it never breaks if the specific ledger names change next quarter
(any ledger Tally puts directly under `Direct Expenses` with no finer sub-group is picked up
automatically). `PL Summary`'s version links straight to `PL GL Detail`'s rows via
`row_map["PROFIT BEFORE VALUATION"]` / `row_map["PROFIT AFTER VALUATION"]`, plus the existing
`section_detail["COST OF SALES"]["Direct Expenses"]["Direct Expenses"]` lookup for the subtotal.

**History worth knowing**: an earlier one-off patch (from a prior session, never captured in any
script) had added a "Profit Before Valuation / Profit After Valuation" block directly onto a
saved `.xlsx` with a *different*, broken formula: it ADDED these items back on top of PBT (which
already included them once) and then immediately SUBTRACTED them again, so "Profit After
Valuation" always collapsed back to exactly PBT regardless of the items' actual values - a
double-count that looked like a reconciliation but reconciled nothing. It survived silently
across rebuilds only because `PL GL Detail` is a fresh `openpyxl.Workbook()`/`create_sheet()`
each run and nothing in the build scripts ever wrote or cleared those specific rows - found and
replaced 2026-07-31 while investigating an unrelated live-refresh request. If a similar-looking
block ever reappears with numbers that don't move no matter what changes upstream, that's the
signature of this bug - trace the formulas (not the cached values) before trusting it.

### How to root-cause a Balance Sheet gap precisely (not just zero it)

When a Reconciliation Gap appears and a screenshot-by-screenshot guessing game isn't converging,
use the Tally `trial-balance` MCP tool directly instead - it returns true ledger-level closing
balances for the whole company at once, which can be grouped and diffed against the CSV
mechanically:

1. `set-company` + `set-period` (same fromDate/toDate convention as the BS pull - `2026-04-01` to
   the period end, since BS is cumulative from FY start) + `chart-of-accounts`.
2. `trial-balance` with the same fromDate/toDate, then `query-database`:
   ```sql
   SELECT c.primary_group, SUM(t.closing_balance) AS total_closing
   FROM <trial_balance_tableID> t
   LEFT JOIN <chart_of_accounts_tableID> c ON t.ledger_name = c.ledger_name
   WHERE t.closing_balance <> 0
   GROUP BY c.primary_group ORDER BY c.primary_group
   ```
3. Compute the SAME group-by-group sums from the CSV, but **through the actual `load_csv()`
   function**, not a naive `csv.DictReader` sum - `load_csv()` applies the duplicate-ledger dedup
   and Stock-in-Hand filter (see "Known Tally API data-quality quirks" below); skipping it produces
   a false-positive ~₹2.89cr "discrepancy" on the blank/`""` group that is actually just an
   unfiltered duplicate "Profit & Loss A/c" row plus the native Stock-in-Hand row we intentionally
   replace with our own synthetic value.
4. Every group should match to the rupee. Whichever one doesn't (or is present natively on Tally's
   BS screen - e.g. "Suspense A/c" - but has zero ledger balance via the API) is where the
   off-ledger item lives - drill into that specific group's Group Summary screen on that exact
   date to find and confirm the figure directly, the same way the existing four constants were
   found.
5. **Tally's own Group Summary screen, when given a date RANGE via F2:Period rather than a single
   as-at date, can show period *turnover* (Debit/Credit transaction volume) instead of the closing
   balance** - a real trap hit this session (a "Sundry Creditors" Group Summary screenshot with a
   `1-May-26 to 31-May-26` header showed unrelated multi-crore Debit/Credit figures that were
   turnover, not balances, sending the investigation down a wrong path before the `trial-balance`
   API comparison caught it). When a screen's numbers look far too large relative to the expected
   adjustment, check whether it's showing a period range instead of a single as-at date before
   trusting it.

**If you add a new per-period constant that changes the number of Balance Sheet buckets**: the
`TOTAL ASSETS`/`TOTAL LIABILITIES` formulas on `MOM PL and BS`/`MOM BS GL Detail` are written as a
literal `SUM(...)` of specific bucket TOTAL rows at build time (not a dynamic range) - adding the
new `primary_group` to `BS_ASSET_ORDER`/`BS_LIABILITY_ORDER` is necessary but the row numbers
below it also shift, so any hardcoded row number in your own verification script (not the
workbook itself, which is always formula-driven) will silently point at the wrong row after a
rebuild - always locate a row by searching for its label text, never reuse a row number from an
earlier read (this bit us once already, see "Verification checklist" item 0).

## Known bugs found & fixed (2026-07-16/17 session) — read before trusting a "gap = 0" claim

Two silent bugs were found and fixed this session. Both are now fixed in the scripts, but the
**patterns** they represent can recur with new data, so document them here permanently.

1. **Case-sensitivity silent data loss.** `classify_bs_row()` / `buckets_for()` in
   `build_gl_detail_sidebyside.py` match a CSV row's `primary_group` against
   `BS_LIABILITY_ORDER` / `BS_ASSET_ORDER` with an **exact string match**. A CSV once had
   `Cash-in-Hand` (capital H) while the script's list had `Cash-in-hand` (lowercase h) — the
   mismatch silently dropped Cash + Petty Cash (₹5,43,776 total) from the Balance Sheet with
   **no error at all**. Fixed by adding `CANONICAL_PRIMARY_GROUP` (a lowercase-keyed lookup dict
   built from both order lists) and normalizing every row's `primary_group` through it inside
   `load_csv()`. **If you ever hand-edit a CSV's `primary_group` column, don't worry about exact
   casing** — normalization handles it — but if a brand-new `primary_group` value appears in
   Tally that isn't in either order list yet, it will still be silently dropped (or
   misclassified as an asset by `classify_bs_row()`'s fallback) since normalization only maps
   *known* variants. Add it to `BS_LIABILITY_ORDER`/`BS_ASSET_ORDER` explicitly when that happens.
2. **MOM formula corruption from a discarded return value.** `write_mom_pl_schedule()` /
   `write_mom_bs_schedule()` in `build_mom_gl_detail.py` used to discard the bucket-rows return
   value (`_`) instead of capturing a flat `{sub_group: row}` dict, and the 4 summary-writer
   functions were then fed the wrong (nested `{bucket: {sub_group: row}}`) dict. This produced
   formulas with a raw Python dict repr embedded as the row reference — e.g.
   `='MOM PL GL Detail'!L{'Sales Accounts': 25, ...}` — which is invalid and un-openable as a
   real calculation, even though openpyxl would happily write it and even read it back as a
   harmless-looking string. Fixed by capturing a proper `section_buckets` dict and threading it
   through all 4 summary-writer functions (see function signatures in
   `build_mom_gl_detail.py`) — only the Indirect Expense drill-down line still correctly uses the
   nested `section_detail` dict, since that one line genuinely needs the sub-group breakdown.
3. **The verification-methodology lesson (most important one):** `openpyxl` does **not**
   evaluate formulas. Reading a freshly-saved workbook's formula-cell values — even with
   `data_only=True` — is unreliable, because openpyxl only returns whatever value was last
   cached by Excel itself; it never computes anything. **A Python re-summation of the source
   CSVs is also not a valid substitute** — it doesn't replicate the actual row-construction /
   bucket-matching logic in the scripts, so it can report "0 gap" even when the real workbook
   formulas are broken (this is exactly what happened here: an early "confirmed 0.00" claim was
   wrong because it was based on CSV re-summation, not the real formulas). **The only trustworthy
   way to verify a formula-heavy version of this workbook is to actually calculate it**, e.g. with
   the Python `formulas` library:
   ```python
   import formulas
   xl = formulas.ExcelModel().loads(r"Files\PL_BS_GL_Detail_<timestamp>.xlsx").finish()  # use this run's actual dated filename
   sol = xl.calculate()
   ```
   ⚠️ **The solution-key format is a trap that manufactures false passes.** Keys look like
   `'[PL_BS_GL_Detail_2026-08-03_1843.xlsx]MOM PL AND BS'!H7` — the **sheet name is uppercased but
   the FILENAME keeps its original case**. Uppercasing the whole key (the obvious thing to do)
   makes every lookup return `None`, and a checker that then reports "no non-zero cells found"
   prints a confident **0.00 that was computed over an empty set**. This happened for real on
   2026-08-03 and is the same class of false "confirmed 0.00" as item 3 below. **Always assert
   that a non-zero number of cells actually resolved** before reporting a pass — the corrected run
   resolved 48 gap cells and 216 Cross Check cells, which is what makes its 0.00 meaningful.

   This is slow (tens of seconds) and noisy (prints a progress bar), but it is a real calculation
   engine — it will throw a genuine `FormulaError` on a malformed formula (which is how bug #2
   above was actually caught) and its computed values can be trusted. Excel COM automation via
   PowerShell was also tried as an alternative and was unreliable (`Unable to get the Open
   property of the Workbooks class` even with no other Excel process running) — prefer `formulas`
   over COM for this project.

## How to run

```powershell
cd "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS"
$ts = Get-Date -Format "yyyy-MM-dd_HHmm"
python ".claude\skills\fetch-pandl\scripts\build_gl_detail_sidebyside.py" --timestamp $ts
python ".claude\skills\fetch-pandl\scripts\build_mom_gl_detail.py" --timestamp $ts
```

**Always pass the same `--timestamp` to both scripts** — they must write to the exact same dated
file (`PL_BS_GL_Detail_<timestamp>.xlsx`). If `--timestamp` is omitted, `build_gl_detail_sidebyside.py`
generates its own from `now()` and `build_mom_gl_detail.py` falls back to whichever
`PL_BS_GL_Detail_*.xlsx` is newest in `Files\` — fine for a quick standalone run of just one
script, but always pass an explicit shared timestamp when running the full pipeline so there is no
window where the two could pick different filenames. Always run `build_gl_detail_sidebyside.py`
**first** — `build_mom_gl_detail.py` imports it as a module
(`import build_gl_detail_sidebyside as b`) and reuses its `load_company_rows()`, constants, and
classification functions directly, so the two scripts must stay in sync. Never run anything in
`scripts/archive/` — those are superseded early iterations (e.g. `build_mom_sheets.py` writes
static/pre-computed MOM sheets and would silently destroy the current formula-linked ones if run
after this pipeline).

Since every run writes a brand-new dated filename rather than overwriting, the "file open in
Excel" lock error from before basically can't happen anymore for a fresh timestamp — but if you
re-run with an explicit `--timestamp` that happens to match a file already open in Excel, both
scripts still raise a clear `RuntimeError` telling you to close it and retry rather than silently
failing.

## Verification checklist (run after every rebuild, not just once)

0. **Do not trust openpyxl's cached formula values, and do not trust a CSV re-summation, as
   proof the Reconciliation Gap is 0** — see "Known bugs found & fixed" above. Before telling
   the user a file is "clear to print," actually calculate the workbook (Python `formulas`
   library) and locate the "RECONCILIATION GAP" row **by searching for the label text**, not by
   reusing a row number from an earlier read — row numbers shift whenever ledger rows are
   added/removed.
1. **Balance Sheet Reconciliation Gap row** (bottom of BS GL Detail / BS Summary, and **each MOM
   BS period column - Apr, May, AND Jun individually, not just the quarter**) should read ~0 for
   every company. If not: something is missing from that company's `gl_data/bs_<KEY>.csv` /
   `gl_data_mom_detail/bs_<period>_<KEY>.csv`, or the Stock/adjustment dicts are stale for that
   specific period. **Do not skip the per-month check** - the Cross Check sheet (item 3 below)
   only verifies MOM-vs-quarter *consistency*, it does NOT verify that any individual month's own
   BS internally balances to zero; a real ~₹1.7cr/₹1.6cr gap sat undetected in Apr/May for exactly
   this reason until checked directly (2026-07-18) - the per-period adjustment constants
   (`SALES_BILLS_TO_MAKE` etc., see "Stock, and the 3(+1) manually-confirmed adjustments" above)
   were simply missing for those two months.
2. **Cross-check PL profit against the BS's own "Profit & Loss A/c" ledger line** for each
   company — the two should be equal, *or* differ by exactly that company's ledger's carried-
   forward Opening Balance (visible on Tally's own BS screen as a separate "Opening Balance" /
   "Current Period" split). Any other unexplained difference means a manual adjustment (like the
   3 above) is only applied to one statement and needs the other side added too.
3. **MOM: Apr + May + Jun should equal the Q1 Total**, and **MOM Q1 Total should equal PL/BS
   Summary**, ledger-line by ledger-line, for both P&L and BS, for every company. This is now a
   **permanent, live sheet in the workbook itself — `Cross Check`** (built by
   `write_cross_check_sheet()` in `build_mom_gl_detail.py`, added automatically by every
   `build_mom_gl_detail.py` run) — do not hand-write an ad-hoc script for this anymore. Read its
   two variance columns per line item via the `formulas` engine (per item 0 above); non-zero
   cells beyond Rs.10 are also highlighted yellow if opened in Excel. See "Standing procedure —
   CFO mode" above for what each of the two checks means and what counts as a real bug vs an
   expected gap.
4. Confirm the file actually saved (`File saved: ...` printed by both scripts) and, if possible,
   spot-check a couple of headline figures (Turnover, PAT, Total Assets) against a fresh Tally
   screenshot for at least one company.
5. **Ratio sheets sanity check** — on both `Ratio Analysis` and `MOM Ratio Analysis`, spot-check
   a handful of cells (e.g. PAT Margin %, Current Ratio) via the `formulas` engine and eyeball
   whether they're in a sane range (Current Ratio positive and roughly 0.5-3, margins in the
   +/-50% range, not four-digit nonsense). A wrong-column bug here won't throw an Excel error —
   `IFERROR` swallows it into a plausible-looking but wrong number — so a range check is the only
   thing that catches it. See "Ratio Analysis structure" above for the specific bug pattern to
   watch for. Also confirm the **RATIO SUMMARY row equals its RATIO WORKINGS — CALCULATIONS row**
   for the same cell (e.g. `Ratio Analysis!I9` should equal `Ratio Analysis!I61`) — they should
   always be identical by construction (Summary is a same-sheet link to Calculations), so any
   difference means the link formula itself is broken, not just a wrong input.
6. **Tab order and grouping** — confirm `wb.sheetnames` matches the fixed order in "Tab order and
   manual Excel grouping" above, and that any columns/rows the user had collapsed before the
   rebuild are still `hidden` after it (`openpyxl`: `ws.column_dimensions[c].hidden` /
   `ws.row_dimensions[r].hidden`). Only needed after a run that could plausibly have touched
   sheet creation/deletion order or either script's `main()` — not every single rebuild.

## `MOM PL and BS 26-27.xlsx` — the standalone month-on-month deliverable

Built by **`scripts/build_mom_pl_bs_26_27.py`** (added 2026-08-04). Run it after Step A/B; it takes
no arguments:

```powershell
python ".claude\skills\fetch-pandl\scripts\build_mom_pl_bs_26_27.py"
```

Before this existed the file had **no builder at all** — it was hand-assembled once (01-Aug-2026)
out of the `SAMPLE_New_Sheets_Preview_v4.xlsx` exploration, so "refresh it" was not a runnable
operation and the only options were to hand-patch it again or write the script. Seven sheets:
`Index`, `PL`/`BS` × `Main Group` / `Sub Group` / `Full GL MOM`, and `Cross Check`.

It is a **presentation layer only** — no fetching, no figures of its own. It reads the same
`gl_data/` + `gl_data_mom_detail/` CSVs through the same `build_multi_period_tree()` that
`build_mom_gl_detail.py` uses, so it can never disagree with `PL_BS_GL_Detail_<ts>.xlsx` about a
number; if it does, that is a bug here, not a difference of opinion. Period blocks come straight
from `PL_PERIODS`/`BS_PERIODS`, so extending the year there extends this workbook with no edit.
The previous version is auto-archived into `Files\Old Version of PL BS\` on every run
(`--no-archive` to suppress).

Only `Full GL MOM` holds literals; `Sub Group` links to it and `Main Group` links to `Sub Group`,
so the three levels cannot drift apart.

### ⚠️ The sign trap — read before touching any headline row

`build_multi_period_tree()` already applies `b.display_sign(bucket)`, so what this script receives
is a **display magnitude**: income and liabilities positive, and expenses and assets *also*
positive (shown as positives under a "LESS:" heading). That is fine for every `SUM`-based subtotal,
because a subtotal only ever adds leaves sharing one bucket — but the **headline waterfall rows
must re-apply the sign themselves**. Summing bucket totals blindly produced, on the first run of
this script, a PAT of ₹1,940cr against ₹62cr of turnover, and a Balance Sheet "gap" of exactly
2 × total liabilities. Hence `EBITDA_ADD` / `EBITDA_LESS` and the explicit `(row, "+"|"-")` operand
lists in `formula_row()`, mirroring `write_multi_computed()` in `build_mom_gl_detail.py`. The
Reconciliation Gap is **Liabilities − Assets**, never a sum. If a new headline row is added, give
it explicit operators and re-verify with the `formulas` engine before believing the number.

## `fetch_pandl.py` — separate, simpler utility

A standalone script, unrelated to the pipeline above, for a quick single P&L pull without the
multi-company/multi-period/MOM machinery:

```powershell
python "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS\.claude\skills\fetch-pandl\scripts\fetch_pandl.py"
```

Arguments (all optional): `--from-date DD-MM-YYYY` (default `01-04-2025`), `--to-date
DD-MM-YYYY` (default `31-03-2026`), `--output <path>` (default `Files/PandL_Output.xlsx`),
`--port PORT` (default from `.env`), `--period 2627|2526|2425` (filter companies by fiscal year
via `Files/TALLY Company_List.xlsx` — see below). Use this only when explicitly asked for a
simple/quick P&L export — for anything about "the PL BS file" or "the consolidated workbook",
use the main pipeline above instead.

**Period filtering (`--period`):** reads `Files/TALLY Company_List.xlsx`, looks up the column
for the requested fiscal year (`26-27` / `25-26` / `24-25`), and only fetches data for
companies marked `1` in that column. Rows marked "Not To USE" are always excluded. Without
`--period`, all open Tally companies are fetched. This applies to both `fetch_pandl.py` and
`fetch_tb_masters.py`. See CLAUDE.md "Company Period Filtering" section for the exact company
lists per period code.

## Outstanding / deferred (not part of this skill's current scope)

- **Previous Year figures** for the two Enterprises units (Tally company codes 100022 / 102426,
  period 1-Apr-2025 to 31-Mar-2026) — requested once, explicitly put on hold by the user, not
  built into either script.
- **Cash Flow / Fund Flow statements** — would need the 31-Mar-2026 opening Balance Sheet,
  mentioned as a possible follow-up but not started.
