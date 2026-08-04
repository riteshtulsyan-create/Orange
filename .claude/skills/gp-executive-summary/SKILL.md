---
name: gp-executive-summary
description: Add or refresh the "Executive Summary" sheet inside Files\GP REPORT.xlsx — a tabular, management-level summary of Gross Profit performance (overall, segment, volume/quantity, key drivers, data gaps, recommended actions). Use this skill whenever the user asks to summarize the GP report, refresh the GP executive summary, or says they've updated GP REPORT.xlsx and want the summary regenerated.
---

# GP Executive Summary

Regenerates the **Executive Summary** sheet (inserted as the first tab) inside `Files\GP REPORT.xlsx`, built from the workbook's own GP REPORT / INK GP REPORT / SPARE PARTS GP REPORT / HEAD GP REPORT sheets.

## Output structure

The generated "Executive Summary" sheet has:
- **1. Overall Performance** — Net Sales / Gross Profit / GP% for Q1 FY25-26 vs Q1 FY26-27
- **2. Segment-wise Performance** — Ink / Spare Parts / Heads / Total, sales and GP%
- **3. Volume / Quantity Trends** — Q1 quantity + realization-per-unit by segment, plus the Apr/May/Jun 2026 monthly quantity trend
- **4. Key Drivers & Anomalies** — the specific sub-category moves that explain the segment numbers (e.g. Pigment Ink margin collapse, the vanished Head "OTHER" category, rising discounts)
- **5. Data Gaps & Caveats** — known holes/quirks in the source report (Machine Sale has no revenue captured, YTD columns aren't like-for-like, Spare Parts quantities mix incompatible units)
- **6. Watch Items / Recommended Actions** — Priority / Area / Action / Rationale

Every section is a plain labeled-column table (no narrative bullets with embedded numbers) — this is a deliberate, user-confirmed formatting preference for this project.

The word "CFO" must never appear in the generated sheet — write it as a plain management/executive summary.

## How to run

```powershell
cd "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS"
python ".claude\skills\gp-executive-summary\scripts\build_exec_summary.py"
```

Optional arg: `--source <path>` (default: `Files\GP REPORT.xlsx`) — the file is read AND written in place; the old "Executive Summary" sheet (if present) is deleted and rebuilt fresh each run.

## Important: the numbers are hand-computed, not live formulas

Unlike the `working-capital-review` skill, the figures in this sheet are **hardcoded Python literals** computed by hand from the source workbook, not Excel formulas referencing the GP REPORT sheets. If `GP REPORT.xlsx`'s underlying data changes (new month added, figures revised, category renamed), this script's numbers go stale silently.

Before re-running for a refresh: re-derive every figure from the current `GP REPORT.xlsx` (open the `GP REPORT`, `INK GP REPORT`, `SPARE PARTS GP REPORT`, and `HEAD GP REPORT` sheets read-only with openpyxl, `data_only=True`) and update the literals in `build_exec_summary.py` accordingly — especially:
- Section 1 & 2 totals (rows 8/22/44/62 style "Q1 25-26" vs "Q1 26-27" columns — NOT the "YTD" columns, which are on different time bases: "25-26 YTD" is a full 12 months, "26-27 YTD" is Q1-to-date only)
- Section 3 quantity/realization figures (sale-quantity columns alongside net-sale columns, same Q1-vs-Q1 basis)
- Section 4's named anomalies — re-check which sub-categories are actually the biggest movers each period; the specific line items called out (e.g. Pigment Ink, Head "OTHER") can change from period to period

## Reporting back to the user

State the headline (Net Sales / GP / GP% change) and flag anything that required updating the hand-written commentary or named anomalies since the source data changed.
