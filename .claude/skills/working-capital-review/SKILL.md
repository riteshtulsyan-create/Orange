---
name: working-capital-review
description: Build or refresh the "Working Capital Review" CFO-style Excel workbook from the CFS 08-07-26.xlsx working-capital tracker. Use this skill whenever the user asks to update the working capital review, refresh the CFO summary, regenerate the working capital workbook, or says they've updated the CFS file and want the review file updated to match.
---

# Working Capital Review

Regenerates `Files\Working_Capital_Review_Jul2026.xlsx` from `Files\CFS 08-07-26.xlsx` — a segment-level (Machines, Ink, Head, Spare Parts, Others) working-capital snapshot comparing a current period against a prior period, pulled from Tally.

## Output structure

The generated workbook has:
- **Source Data 8-Jul** / **Source Data 12-Jun** — formatted, formula-safe copies of the two source sheets
- **Summary** sheet:
  - Headline (Net WC change, Rs. Cr and %)
  - At a Glance table (Inventory / Debtors / Creditors / Net WC, before vs after)
  - Segment-wise Detail table (same four metrics, broken out by segment)
  - **Key Takeaways: Net Working Capital Movement** — Segment / Now / Before / Change / Takeaway
  - **Key Takeaways: Receivables Ageing Risk** — Segment / Aged >180d / Total / % Overdue / Risk
  - **Action Plan** — Segment / Action / Now / Target (quarter-end) / Basis

Every numeric cell in the Summary sheet is a live formula pointing at the Source Data sheets — nothing is hardcoded. Only the "Takeaway" / "Risk" / "Basis" text columns are plain prose (see caveat below).

This tabular format — one labeled column per figure, no numbers embedded in narrative bullets — is a deliberate, user-confirmed preference for this project. Don't revert to bullet-point takeaways with inline numbers or `[Ref: ...]` citation tags; both were tried and rejected in favor of this structure.

## How to run

```powershell
cd "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS"
python ".claude\skills\working-capital-review\scripts\make_excel_report.py"
```

Optional args: `--base <path>` (default: `Files\CFS 08-07-26.xlsx`), `--output <path>` (default: `Files\Working_Capital_Review_Jul2026.xlsx`).

## Before running: check the source file hasn't changed shape

The base CFS file has already changed structure once (it used to be one `Sheet1` with two side-by-side date blocks; it's now two separate sheets, `" Jul 26"` and `"12 June"`, each with its own row numbering). If the script errors or produces obviously wrong numbers, open `CFS 08-07-26.xlsx` read-only with openpyxl and diff the sheet names / row layout against the constants at the top of `make_excel_report.py` (`CUR_INV`, `CUR_CUST`, `CUR_CRED`, `CUR_NET`, `CUR_PIVOT_CUST`, `CUR_PIVOT_VEND`, `PRI_INV`, etc.) — remap them if the layout moved.

## After running: sanity-check the hand-written commentary

The numeric columns (Now/Before/Change/%) always self-update correctly. The **prose** in `tableA_rows`, `tableB_rows`, and `tableC_rows` (near the bottom of the script) is hand-authored and does not. Claims like "the single biggest driver" or "improving on net" can flip to a different segment every time the source data is revised — this has already happened twice in this project (Ink and Others traded places as "biggest driver"; Machines flipped from "improving" to "slightly increasing" between two updates).

After regenerating, read the computed Now/Before/Change values back (see verification method below) and re-check each takeaway/risk/basis sentence still matches reality before telling the user it's done. Edit the row tuples directly in the script if not.

## Verifying output — do NOT use Excel COM automation on this file

A prior session used `win32com` to open the output workbook, recalculate, and read values back for verification. This caused a real incident: an orphaned, invisible Excel process from an earlier verification call still had a stale (older) copy of the workbook loaded in memory, and OneDrive's AutoSave silently wrote that stale copy back over a freshly regenerated file — clobbering real work.

Instead, verify by reading the saved `.xlsx` back with `openpyxl` in read-only mode (`data_only=False` to check formula text, or compute the expected numbers directly from the source file in a separate one-off script) — this never opens Excel and can't be clobbered by AutoSave. If you must confirm computed values (not just formula text), replicate the arithmetic in Python from the source workbook rather than letting Excel evaluate it.

If you do need Excel open for any reason, check `tasklist | grep -i excel` first and be aware that a blank `MainWindowTitle` usually means an orphaned automation instance (safe to close) while a populated title is very likely the user's own open window (do not close without asking) — but a single Excel process can also hold multiple workbook windows, so a visible title for one file doesn't guarantee no other file is open in the same process.

## Reporting back to the user

State the new headline figure (Net WC change, Rs. Cr and %) and flag anything that changed enough to have required editing the commentary text (e.g. "the biggest driver changed from X to Y since last time"). If a source-structure change forced you to drop or rebuild a table (this happened once with an aggregate ageing table that was removed from the source), say so explicitly rather than silently omitting it.
