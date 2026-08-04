"""
refresh_expense_sheet.py

One-command refresh of the Expense Sheet workbook: re-fetches Direct + Indirect
expense vouchers from Tally for the given period, rebuilds the RAW DATA sheet,
then rebuilds every analysis sheet on top of it (Summary, Direct, Indirect Exp,
Direct by Component, one sheet per indirect main group, Index) so everything
stays linked and refreshes together.

A timestamped backup of the previous workbook is kept before overwriting.

Company selection for the period (2627/2526/2425) is resolved from
Files/TALLY Company_List.xlsx - see generate_expense_sheet.companies_for_period().

Usage (run from the EXPENSES REFERESH folder, or anywhere - paths are absolute):
    python scripts/refresh_expense_sheet.py --from 20260401 --to 20260630 --period 2627
"""

import os
import sys
import shutil
import argparse
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import generate_expense_sheet
import build_analysis_sheets

FOLDER = Path(__file__).resolve().parent.parent
TARGET_FILE = FOLDER / "Expense Sheet 2026-07-14 17-38-44.xlsx"


def refresh(from_date: str, to_date: str, target_file: Path = TARGET_FILE, period: str = "2627"):
    target_file = Path(target_file)

    if target_file.exists():
        backup_dir = target_file.parent / "backups"
        backup_dir.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        backup_path = backup_dir / f"{target_file.stem}_backup_{stamp}{target_file.suffix}"
        shutil.copy2(target_file, backup_path)
        print(f"Backed up existing file to: {backup_path}")

    print("=" * 70)
    print("STEP 1/2: Fetching Direct + Indirect expenses from Tally...")
    print("=" * 70)
    generate_expense_sheet.run(from_date, to_date, str(target_file), period=period)

    print("=" * 70)
    print("STEP 2/2: Building analysis sheets (Summary, Direct, Indirect Exp,")
    print("          Direct by Component, indirect vendor sheets, Index)...")
    print("=" * 70)
    build_analysis_sheets.build(str(target_file), from_date, to_date)

    print("=" * 70)
    print(f"DONE. Refreshed workbook: {target_file}")
    print("=" * 70)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Refresh the Expense Sheet end to end from Tally")
    parser.add_argument("--from", dest="from_date", required=True, help="YYYYMMDD (e.g. 20260401)")
    parser.add_argument("--to", dest="to_date", required=True, help="YYYYMMDD (e.g. 20260630)")
    parser.add_argument("--out", dest="out_file", default=None,
                         help="Target workbook path (default: overwrite the existing Expense Sheet in this folder)")
    parser.add_argument("--period", dest="period", default="2627",
                         help="Period code 2627/2526/2425 - selects companies via "
                              "Files/TALLY Company_List.xlsx (default: 2627)")
    args = parser.parse_args()
    refresh(args.from_date, args.to_date, Path(args.out_file) if args.out_file else TARGET_FILE,
            period=args.period)
