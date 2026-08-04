---
name: fetch-tb-masters
description: Fetch Trial Balance Masters (Chart of Accounts) from Tally Prime for all open companies and export to Excel. Use this skill whenever the user asks to fetch TB masters, get trial balance masters, fetch chart of accounts, get ledger masters, or run the TB masters script from Tally.
---

# Fetch Trial Balance Masters from Tally Prime

Runs a Python script that fetches the Chart of Accounts (ledger masters + group hierarchy) from Tally Prime for all open companies and writes a formatted Excel report with audit trail and GL name variation detection.

## Script location

`C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS\.claude\skills\fetch-tb-masters\scripts\fetch_tb_masters.py`

## Output

`Files/TB_Masters_Report.xlsx` with three sheets:
- **Masters** — 18 columns per ledger (company, group hierarchy, schedule mapping, audit fields)
- **Name Variations** — ledgers whose names differ across companies (dot/space/spelling)
- **Summary** — ledger count, unmapped count, and variation count per company

## How to run

Always run from the project directory so `.env` is found:

```powershell
cd "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS"
python "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS\.claude\skills\fetch-tb-masters\scripts\fetch_tb_masters.py"
```

Arguments (all optional):
- `--output <path>` — output Excel path (default: `Files/TB_Masters_Report.xlsx`)
- `--host HOST` — Tally host (default: `localhost`)
- `--port PORT` — Tally port (default: `9000`, from `.env`)
- `--period 2627|2526|2425` — filter companies by fiscal year via `Files/TALLY Company_List.xlsx`

**Period filtering**: when `--period` is passed, the script reads `Files/TALLY Company_List.xlsx`, looks up the matching column (`26-27` / `25-26` / `24-25`), and only fetches masters for companies marked `1` in that column. Rows marked "Not To USE" are always excluded. Without `--period`, all open companies are fetched (default).

```powershell
# FY 2026-27 companies only
python "...\fetch_tb_masters.py" --period 2627
```

## Steps to follow when this skill triggers

1. Run the script via PowerShell from the project directory.
2. Report the output file path, total ledger count, and how many variation groups were flagged.
3. If Tally is unreachable, tell the user: "Make sure Tally Prime is open with at least one company loaded and the XML/HTTP server is enabled on the port set in `.env` (currently 9000)."
