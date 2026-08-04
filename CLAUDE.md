# Project: PL and BS — Claude Instructions

## Tally Prime Connection

**Always connect to Tally on `localhost`** — Tally Prime runs on this machine. The port is
currently **9000** (Tally's default); it was 9010 previously.

- The host matters more than the port. Do NOT use the old server IP `192.168.33.254` — that is a
  network server and may not have all companies open. `localhost:9000` (this machine) and
  `192.168.33.254:9000` (the network server) share a port number but are **not** the same target.
- The `.env` file at the project root stores the canonical connection settings (`TALLY_HOST`,
  `TALLY_PORT`). It is the single source of truth — if Tally's port is changed in
  F1 > Settings > Connectivity, update `.env` to match and nothing else needs editing.
- All scripts must load `.env` at startup and use those values as defaults.
- If a script reports "Cannot connect", check what Tally is actually listening on before assuming
  it is down:
  `Get-NetTCPConnection -State Listen | Where-Object { $_.OwningProcess -eq (Get-Process tally).Id }`

**Always check WHICH companies are open before trusting a refresh — a wrong one fails silently.**
On 2026-08-04 Tally had `ORANGE O TEC PRIVATE LIMITED-NOIDA - (from 1-Apr-24)` (code 100001, the
FY24-25 file) open instead of `...-NOIDA-(from 1-Apr-25)` (code 100000, the FY26-27 one). It still
returned Apr–Jul-26 figures and nothing errored — the numbers were simply the wrong entity's, and
would have flowed straight into Total Orange O Tec and Grand Total. The tell was that its 30-Jun
balances did not tie to the previous good pull (Branch/Divisions ₹8.34cr vs ₹8.04cr). Run
`python .claude\skills\fetch-pandl\scripts\fetch_gl_data.py --probe` to list what is open and
diff it against the period table below before fetching.

## Scripts

Each fetch utility lives in its own skill folder under `.claude\skills\`:

| Script | Skill | Purpose |
|--------|-------|---------|
| `fetch_pandl.py` | `fetch-pandl` | Fetch a simple Profit & Loss report from Tally → `Files/PandL_Output.xlsx` |
| `fetch_tb_masters.py` | `fetch-tb-masters` | Fetch Chart of Accounts masters for all open Tally companies → `Files/TB_Masters_Report.xlsx` |
| `fetch-bank-transactions.mjs` | `fetch-bank-transactions` | Fetch bank payment/receipt vouchers from Tally → flat + pivot Excel reports |
| `build_mom_pl_bs_26_27.py` | `fetch-pandl` | Build `Files/MOM PL and BS 26-27.xlsx` — the standalone month-on-month P&L/BS deliverable (Main / Sub / Full-GL levels for both statements). Presentation layer over the same CSVs; run after a refresh. See the skill's own section for the sign trap. |

`fetch_tb_masters.py` used to live inside `fetch-pandl\scripts\` — it was moved into its own
`fetch-tb-masters` skill folder; use the path below, not the old one.

**Exception — stock / inventory reporting**, which does NOT use the project-level `.claude\skills\`
or `Files\`:

| Script | Skill | Purpose |
|--------|-------|---------|
| `fetch_stock_summary.py` | `Stock Report:fetch-stock-summary` | Fetch Stock Summary (closing inventory) with the full stock-group hierarchy → `Inventory\Stock_Summary_<company>_<timestamp>.xlsx` |
| *(extract not yet built)* | `Stock Report:inventory-valuation` | **Inventory Valuation check** — rebuild closing stock on FIFO / weighted-avg / last-purchase from voucher data, compare to Tally's carrying value, report the gap per legal entity → `Inventory\Stock_Valuation_Check_<timestamp>.xlsx` |
| *(extract not yet built)* | `Stock Report:inventory-ageing` | **Inventory Ageing** — split closing stock into 6 age buckets by receipt date (FIFO), add last-movement / months-of-cover analysis, drive a slow-moving provision → `Inventory\Stock_Ageing_Analysis_<timestamp>.xlsx` |

```
Inventory\                              <- ALL stock workbooks (top level, easy to find)
  Stock_Summary_<company>_<timestamp>.xlsx        <- current
  Stock_Valuation_Check_<timestamp>.xlsx          <- current
  SAMPLE LAYOUT - Stock_Valuation_Check.xlsx      <- approved design sample (illustrative figures)
  SAMPLE LAYOUT - Stock_Ageing_Analysis.xlsx      <- ageing design sample, awaiting approval
  Old Versions\                                   <- superseded, auto-archived on each run
Stock Report\.claude\skills\fetch-stock-summary\   <- SKILL.md + scripts\fetch_stock_summary.py
Stock Report\.claude\skills\inventory-valuation\   <- SKILL.md + scripts\build_sample_layout.py
Stock Report\.claude\skills\inventory-ageing\      <- SKILL.md + scripts\build_sample_layout.py
```

`inventory-valuation` reports at **legal-entity** level — the 5 Tally files collapse to 3 entities
(Orange O Tec = HO + Noida, Enterprises = HO + Noida, Colorix standalone), so HO↔Noida stock
transfers eliminate *inside* the entity, not at group consolidation. Its live extract script is
**not yet written**; the SKILL.md is the approved specification.

Rules for stock reports:

- Output always goes to **`Inventory\`** — never `Files\`, never inside `Stock Report\`.
- The skill stays where it is, in `Stock Report\.claude\skills\` (it registers as the
  directory-scoped skill `Stock Report:fetch-stock-summary`). Keep skill and output separate.
- Period is **always 1-Apr of the current FY through today** — the script's default, so normally
  just pass `--company`.

## Running Scripts

Use the matching skill (say "fetch pnl from tally", "fetch tb masters", "fetch bank transactions"), or run directly — always from this project directory so `.env` is found:

```bash
python "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS\.claude\skills\fetch-pandl\scripts\fetch_pandl.py"
python "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS\.claude\skills\fetch-tb-masters\scripts\fetch_tb_masters.py"
python "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS\Stock Report\.claude\skills\fetch-stock-summary\scripts\fetch_stock_summary.py" --company "<exact Tally company name>"
```

(The stock script resolves `.env` and its output folder from its own location, so it runs from any
working directory.)

## Quarterly Company Review

The `company-review` skill (`.claude\skills\company-review\`) builds/refreshes the quarterly Company Review: `Files\Company_Review_Board_Pack_Q1_FY26-27.xlsx`, `Files\Company_Action_Plan_Q1_FY26-27.xlsx`, and an HTML board dashboard published as a Claude Artifact. See that skill's `SKILL.md` for the full step-by-step process, including the mandatory independent-recalculation check against the latest `PL_BS_GL_Detail_<timestamp>.xlsx` (see below — there is no longer a fixed filename) before trusting any "fixed"/"reconciled" claim.

## Company Period Filtering — `Files/TALLY Company_List.xlsx`

When the user specifies a period code (**2627**, **2526**, or **2425**), look up the matching column in `Files/TALLY Company_List.xlsx` and **only fetch/run for the companies marked with `1` in that column**. This rule applies to ALL reports, scripts, and activities in this project.

| Period Code | Fiscal Year | Companies (as of 2026-07-25) |
|-------------|-------------|------------------------------|
| **2627** | FY 2026-27 | COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20) · ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27 · ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27) · ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27) · ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25) |
| **2526** | FY 2025-26 | COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20) · ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA - FY25-26 · ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2024-26) · ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27) · ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25) |
| **2425** | FY 2024-25 | COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20) · ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA - FY25-26 · ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2024-26) · ORANGE O TEC PRIVATE LIMITED (01-04-24TO31-03-25) · ORANGE O TEC PRIVATE LIMITED-NOIDA - (from 1-Apr-24) |

The source of truth is always the file — re-read it when asked, as the user may update it. Rows marked "Not To USE" (Group Company entries) are always excluded regardless of column markings.

All fetch scripts support `--period 2627` (or `2526`/`2425`) to apply this filter automatically. Without `--period`, scripts run against all open companies (default behaviour).

## Standing instruction: CFO mode — always work from live Tally data

**By default, before delivering any final report, figure, or conclusion drawn from the P&L/BS
workbook, run the full `fetch-pandl` skill procedure**: live refresh from Tally (quarter *and*
monthly, all 5 companies) → rebuild → verify the `Cross Check` sheet shows no real gaps (via the
`formulas` engine, never cached values) → generate the change-reconciliation report → only then
report. See the `fetch-pandl` skill's "Standing procedure — CFO mode" section for the exact steps
and what counts as a real bug vs an expected/explained gap.

**File naming (2026-07-25 onward): there is no fixed `PL_BS_GL_Detail.xlsx` filename anymore.**
Every refresh produces a new, dated `Files\PL_BS_GL_Detail_<YYYY-MM-DD_HHMM>.xlsx`, and a matching
dated `Files\Profit_Change_Reconciliation_<same timestamp>.xlsx` showing what changed vs the
previous refresh. The immediately-superseded previous version of each moves into
`Files\Old Version of PL BS\` once the new one is confirmed good. "The current workbook" always
means whichever `PL_BS_GL_Detail_*.xlsx` is newest directly inside `Files\` (not inside the
archive subfolder). See the `fetch-pandl` skill's "File naming and version history" section for
the full mechanics.

This is the default for every session, not a one-off — do it again next time even if it was just
done, unless the user's own message for that turn explicitly narrows the scope (e.g. "no need to
refresh," "just explain this row," "use what's already in the file"). If scoping down, say so
explicitly rather than silently skipping the refresh.
