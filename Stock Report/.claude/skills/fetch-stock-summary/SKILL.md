---
name: fetch-stock-summary
description: Fetch the Stock Summary (closing inventory) from Tally Prime for a company and rebuild the complete stock-group hierarchy — every group, sub-group and item name — into an Excel workbook. Use whenever the user asks for a stock report, inventory report, stock summary, closing stock, stock item list, stock group breakdown, or says "share the stock report for <company>".
---

# Fetch Stock Summary

Produces `Inventory\Stock_Summary_<company-slug>_<YYYY-MM-DD_HHMM>.xlsx` — the Tally Stock Summary
screen with **every** nesting level expanded, plus a flat pivot-ready item list.

## Folder layout — output is kept SEPARATE from the skill

```
PL and BS\
  Inventory\                                  <- ALL stock workbooks go here (top level, visible)
      Stock_Summary_<company>_<timestamp>.xlsx    <- current
      Old Versions\                               <- superseded, auto-archived
  Stock Report\.claude\skills\fetch-stock-summary\
      SKILL.md                                <- this file
      scripts\fetch_stock_summary.py          <- the fetch script
```

Rules:

- Stock workbooks go in **`Inventory\`** at the project root — never in `Files\`, and never
  inside `Stock Report\`. The user needs to find them without digging into a hidden `.claude` path.
- On each run the previous workbook for that company is moved to `Inventory\Old Versions\`
  automatically, so `Inventory\` holds exactly one current file per company.
- The script derives `Inventory\` from wherever `.env` lives, so it can be run from any working
  directory and still writes to the same place.

## Period rule (standing)

**Always run from 1-Apr of the current financial year through today**, unless the user names a
different period for that specific ask. This is the script's default when `--from`/`--to` are
omitted, so normally just pass `--company`.

## Run

```powershell
Set-Location "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS"
python "Stock Report\.claude\skills\fetch-stock-summary\scripts\fetch_stock_summary.py" --company "<exact Tally company name>"
```

`.env` (TALLY_HOST / TALLY_PORT) is located by walking up from the script, so any working
directory works — the `Set-Location` above is just for a short command line.

Options:

| Flag | Purpose |
|------|---------|
| `--company` | **Required.** Exact Tally company name. Validate first via the `list-master` MCP tool with `collection: company`. |
| `--from` / `--to` | `YYYY-MM-DD`. Defaults: 1-Apr of current FY → today. |
| `--expect-total` | Positive grand total to reconcile against (e.g. read off the user's Tally screen). Writes an OK/MISMATCH row into `Cross Check`. |
| `--output` | Override the output path. |
| `--host` / `--port` | Override `.env`. |

## Output sheets

**Default output is two sheets only.** Do not add others back unless asked.

| Sheet | Content |
|-------|---------|
| `Item List (Flat)` | One row per item. Columns: **Main Category · Sub Category · Product Group · Sub Group** (the four group levels), Full Group Path, Item Name, Quantity, Unit, Rate, Value, Depth, Has Qty, Zero Value. Autofiltered, pivot-ready. |
| `Cross Check` | Reconciliation and data-quality flags — see below. |

Pass `--all-sheets` to additionally emit `Summary` (collapsible group overview with bar chart and
exceptions block), `Hierarchy` (indented tree mirroring the Tally screen) and `Group Summary`.
Those builders are intact in the script — they are simply dropped before saving by default.

### The four level columns

The stock tree is four groups deep at most, so there are exactly four level columns:

| Column | Holds | Example |
|--------|-------|---------|
| `Main Category` | Top-level stock group | MACHINERY & MACHINERY PARTS |
| `Sub Category` | Second level | MACHINARY PARTS |
| `Product Group` | Third level | OTHER SPARE PARTS |
| `Sub Group` | Fourth level | DESIGN DEPARTMENT TOOLS |

Cells are **left blank** where an item does not go that deep — an empty cell means the level
genuinely does not exist for that item. Do not fill them by repeating the parent; that would imply
sub-groups that are not in Tally. If a company ever nests deeper than four, the extra levels are
folded into `Sub Group` and `Full Group Path` still carries the complete chain.

## Verification — do this before reporting figures

1. Confirm `Cross Check` shows **OK** for "Sum of top-level group roll-ups" (must equal Grand Total).
2. If the user supplied a Tally screenshot, pass its Grand Total via `--expect-total` and confirm **OK**.
3. Spot-check two or three level-2 group totals against the screenshot before quoting anything.

## Known behaviours — not bugs

- **Sign flip.** Tally returns stock closing values on the credit side (negative). The script detects
  a negative net total and multiplies everything by -1 so the workbook reads like the screen. The
  `Cross Check` sheet records whether the flip was applied.
- **Zero-value items.** Tally hides these on screen; the workbook keeps them (flagged `Zero Value = Yes`
  on the flat sheet) so the full item master is visible. Filter them out for a screen-identical view.
- **Blank group quantities.** A group row shows no quantity/rate in three cases, each verified
  against the Tally screen — do NOT "fix" these by summing:
  1. **Mixed units** under the group (e.g. `PAPER ROLL` = ROLL + MTR). A cross-UOM total is meaningless.
  2. **Net quantity of zero** (e.g. `MACHINERY HEAD`, all PCS, nets to nil). Tally prints blank, not `0.000`.
  3. **Top-level group rows.** Tally does not roll a quantity up to the primary group row even when
     the unit is unambiguous — `PRINTING INK` is all KGS and `SOFTWARE` all PCS, and both print blank.
     The script blanks level-1 rows to match.

  Every level-2 quantity and rate the script produces was reconciled against the Tally screen
  (FLOTEC INK 18,520.000 KGS @ 329.61, CHEMICALS 155.400 KGS @ 628.62, SUZHOU 9,000.000 MTR @ 5.46,
  MACHINE PROCUREMENT 1.000 PCS @ 6,905.00, 444-027 SOFT DONGLE 2.000 PCS @ 10,000.00).

## Genuine exceptions to raise with the user

Report these — they are accounting issues, not script issues:

- **Items with negative closing value / quantity** — negative stock means issues booked without
  a corresponding receipt.
- **Items with no stock group assigned** — master-data gap.
- **Items with a value but zero quantity** — stranded value with nothing on hand.
