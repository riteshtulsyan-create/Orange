"""
refresh_with_history.py

ONE-OFF historical backfill refresh (2026-07-16, put on hold 2026-07-16).

Ent-Surat and Ent-Noida were "split" by Tally into separate per-financial-year
company files. This script was written to backfill FY25-26 history from the
older split files:

  - "ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2024-26)"           (Ent-Surat)
  - "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA - FY25-26" (Ent-Noida)

DO NOT FETCH FROM THESE TWO COMPANIES. Confirmed via two independent tools
(this script's own voucher fetch, and the Tally MCP trial-balance tool) that
Trial Balance shows real FY25-26 activity in both (1,012 / 540 ledgers,
Rs.14.8Cr / Rs.3.8Cr movement), but the voucher-level Collection API returns
essentially nothing for them (1 voucher total for the Ent-Surat file; all
vouchers dated Apr-Jul 2026, not FY25-26, for the Ent-Noida file). This is a
Tally-side limitation on these two archived/split company files, not a bug in
this script - the user has asked to leave these alone until a manual export
from the Tally GUI (Day Book -> Alt+E -> XML/Excel) can supply the FY25-26
voucher detail instead. HISTORY_SOURCES and the Step 2 fetch are disabled
below; re-enable only if the user confirms the underlying Tally access issue
is resolved.

Ordinary future refreshes should keep using refresh_expense_sheet.py against
the standard 5-company COMPANIES list in generate_expense_sheet.py, which
never included these two companies.

Usage:
    python scripts/refresh_with_history.py
"""

import os
import sys
import shutil
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generate_expense_sheet
import build_analysis_sheets
from fetch_tally_expenses import get_open_companies

FOLDER = Path(__file__).resolve().parent.parent
TARGET_FILE = FOLDER / "Expense Sheet 2026-07-14 17-38-44.xlsx"

CURRENT_FROM, CURRENT_TO = "20260401", "20260630"
HISTORY_FROM, HISTORY_TO = "20250401", "20260331"

# On hold 2026-07-16 - do not fetch from these two companies (see module
# docstring). Kept here only as a record of which files would supply FY25-26
# history if/when a working access path is found. This duplicates 2 of the 5
# entries generate_expense_sheet.companies_for_period("2526") would now
# return (confirmed against Files/TALLY Company_List.xlsx on 2026-07-21) -
# switch to filtering that instead of hand-maintaining this list once the
# underlying Tally API access issue is resolved.
HISTORY_SOURCES = [
    {
        "tally_name":   "ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2024-26)",
        "key":          "OOTEPL_FY2426",
        "company_col":  "Ent-Surat",
        "location":     "SURAT",
        "company_name": "Ent",
    },
    {
        "tally_name":   "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA - FY25-26",
        "key":          "OOTEPL_NOIDA_FY2526",
        "company_col":  "Ent-Noida",
        "location":     "NOIDA",
        "company_name": "Ent",
    },
]

FETCH_HISTORY = False  # set True only once a working access path is confirmed


def main():
    if TARGET_FILE.exists():
        backup_dir = TARGET_FILE.parent / "backups"
        backup_dir.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        backup_path = backup_dir / f"{TARGET_FILE.stem}_backup_{stamp}{TARGET_FILE.suffix}"
        shutil.copy2(TARGET_FILE, backup_path)
        print(f"Backed up existing file to: {backup_path}")

    generate_expense_sheet._check_tally()
    open_companies = get_open_companies()
    open_names_lower = {n.lower() for n in open_companies}
    print(f"Open companies in Tally ({len(open_companies)}):")
    for n in open_companies:
        print(f"  - {n}")

    print("\n" + "=" * 70)
    print(f"STEP 1/3: Current period ({CURRENT_FROM} to {CURRENT_TO}) - 5 companies")
    print("=" * 70)
    current_rows = generate_expense_sheet.fetch_all_rows(
        CURRENT_FROM, CURRENT_TO,
        companies=generate_expense_sheet.companies_for_period("2627"),
        open_names_lower=open_names_lower,
    )
    print(f"Current-period rows: {len(current_rows)}")

    print("\n" + "=" * 70)
    if FETCH_HISTORY:
        print(f"STEP 2/3: Historical backfill ({HISTORY_FROM} to {HISTORY_TO}) - Ent-Surat/Ent-Noida older split files")
        print("=" * 70)
        history_rows = generate_expense_sheet.fetch_all_rows(
            HISTORY_FROM, HISTORY_TO,
            companies=HISTORY_SOURCES,
            open_names_lower=open_names_lower,
        )
        print(f"Historical-period rows: {len(history_rows)}")
    else:
        print("STEP 2/3: SKIPPED - historical backfill on hold, see module docstring")
        print("=" * 70)
        history_rows = []

    all_rows = current_rows + history_rows
    print(f"\nTotal combined rows: {len(all_rows)}")
    if not all_rows:
        print("No rows fetched at all - aborting without touching the workbook.")
        sys.exit(1)

    generate_expense_sheet.write_excel(all_rows, str(TARGET_FILE))

    build_from, build_to = (HISTORY_FROM, CURRENT_TO) if FETCH_HISTORY else (CURRENT_FROM, CURRENT_TO)
    print("\n" + "=" * 70)
    print(f"STEP 3/3: Building analysis sheets (scoped to {build_from} .. {build_to})")
    print("=" * 70)
    build_analysis_sheets.build(str(TARGET_FILE), build_from, build_to)

    print("\n" + "=" * 70)
    print(f"DONE. Refreshed workbook: {TARGET_FILE}")
    print("=" * 70)


if __name__ == "__main__":
    main()
