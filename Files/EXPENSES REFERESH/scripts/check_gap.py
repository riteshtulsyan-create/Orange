"""
check_gap.py

Compares the current RAW DATA sheet against the most recent backup (or an
explicitly given "before" file) and reports what genuinely changed - new
transactions, transactions no longer present, and amount corrections -
filtering out cosmetic noise (ledger renames, voucher-type reclassification
during GST-return finalization) that would otherwise show up as false
add/remove pairs in a naive row-by-row diff.

Usage (run from the EXPENSES REFERESH folder):
    python scripts/check_gap.py
        (compares the live workbook against the newest file in backups/)
    python scripts/check_gap.py --before "backups\\Expense Sheet ..._backup_....xlsx"
        (compares against a specific backup instead of the newest one)
"""

import os
import re
import sys
import argparse
from pathlib import Path
from collections import Counter, defaultdict

import openpyxl

FOLDER = Path(__file__).resolve().parent.parent
TARGET_FILE = FOLDER / "Expense Sheet 2026-07-14 17-38-44.xlsx"
BACKUP_DIR = FOLDER / "backups"


def _latest_backup() -> Path:
    backups = sorted(BACKUP_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime)
    if not backups:
        raise FileNotFoundError(f"No backups found in {BACKUP_DIR}")
    return backups[-1]


def _load_rows(path):
    wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    ws = wb["RAW DATA"]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    headers = rows[0]
    idx = {h: i for i, h in enumerate(headers)}
    data = [r for r in rows[1:] if r[0] is not None]
    return idx, data


def _norm_ledger(name):
    return re.sub(r"[^A-Z0-9]", "", str(name).upper())


def check_gap(before_path, after_path=TARGET_FILE):
    idx, old_rows = _load_rows(before_path)
    _, new_rows = _load_rows(after_path)

    def key(r):
        amt = r[idx["NET AMOUNT"]]
        amt = round(amt, 2) if isinstance(amt, (int, float)) else amt
        return (r[idx["Company"]], str(r[idx["DATE"]]), _norm_ledger(r[idx["LEDGER NAME"]]), amt)

    old_c = Counter(key(r) for r in old_rows)
    new_c = Counter(key(r) for r in new_rows)
    added = new_c - old_c
    removed = old_c - new_c

    old_by_key = defaultdict(list)
    for r in old_rows:
        old_by_key[key(r)].append(r)
    new_by_key = defaultdict(list)
    for r in new_rows:
        new_by_key[key(r)].append(r)

    print(f"Before: {before_path}")
    print(f"After:  {after_path}")
    print(f"Rows: {len(old_rows)} -> {len(new_rows)}\n")

    added_total = sum(new_by_key[k][0][idx["NET AMOUNT"]] * c for k, c in added.items())
    removed_total = sum(old_by_key[k][0][idx["NET AMOUNT"]] * c for k, c in removed.items())
    print(f"Genuinely new rows: {sum(added.values())}   (total {added_total:,.2f})")
    print(f"Genuinely removed rows: {sum(removed.values())}   (total {removed_total:,.2f})")
    print(f"Net genuine change: {added_total - removed_total:,.2f}\n")

    if added:
        print("--- New transactions ---")
        for k, c in sorted(added.items(), key=lambda kv: -abs(new_by_key[kv[0]][0][idx["NET AMOUNT"]])):
            r = new_by_key[k][0]
            print(f"  {r[idx['Company']]:12s} {str(r[idx['DATE']])[:10]} {r[idx['EXP TYPE']]:9s} "
                  f"{r[idx['MAIN GROUP']]:25s} {str(r[idx['LEDGER NAME']])[:35]:35s} "
                  f"{r[idx['NET AMOUNT']]:>14,.2f}  VCH:{r[idx['VCH NO.']]}")

    if removed:
        print("\n--- Removed / no-longer-present transactions ---")
        for k, c in sorted(removed.items(), key=lambda kv: -abs(old_by_key[kv[0]][0][idx["NET AMOUNT"]])):
            r = old_by_key[k][0]
            print(f"  {r[idx['Company']]:12s} {str(r[idx['DATE']])[:10]} {r[idx['EXP TYPE']]:9s} "
                  f"{r[idx['MAIN GROUP']]:25s} {str(r[idx['LEDGER NAME']])[:35]:35s} "
                  f"{r[idx['NET AMOUNT']]:>14,.2f}  VCH:{r[idx['VCH NO.']]}")

    print("\n--- Old total vs new total NET AMOUNT by EXP TYPE ---")
    old_by_type = defaultdict(float)
    new_by_type = defaultdict(float)
    for r in old_rows:
        old_by_type[r[idx["EXP TYPE"]]] += r[idx["NET AMOUNT"]] or 0
    for r in new_rows:
        new_by_type[r[idx["EXP TYPE"]]] += r[idx["NET AMOUNT"]] or 0
    for t in sorted(set(list(old_by_type) + list(new_by_type))):
        print(f"  {t:10s}  old={old_by_type[t]:>15,.2f}  new={new_by_type[t]:>15,.2f}  "
              f"diff={new_by_type[t] - old_by_type[t]:>15,.2f}")

    print("\n--- By company: old vs new total ---")
    old_by_co = defaultdict(float)
    new_by_co = defaultdict(float)
    for r in old_rows:
        old_by_co[r[idx["Company"]]] += r[idx["NET AMOUNT"]] or 0
    for r in new_rows:
        new_by_co[r[idx["Company"]]] += r[idx["NET AMOUNT"]] or 0
    for c in sorted(set(list(old_by_co) + list(new_by_co))):
        print(f"  {c:12s}  old={old_by_co[c]:>15,.2f}  new={new_by_co[c]:>15,.2f}  "
              f"diff={new_by_co[c] - old_by_co[c]:>15,.2f}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Show what genuinely changed since the last refresh")
    parser.add_argument("--before", default=None, help="Path to the 'before' file (default: newest file in backups/)")
    parser.add_argument("--after", default=str(TARGET_FILE), help="Path to the 'after' file (default: the live workbook)")
    args = parser.parse_args()
    before = Path(args.before) if args.before else _latest_backup()
    check_gap(before, Path(args.after))
