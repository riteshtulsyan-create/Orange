# -*- coding: utf-8 -*-
import win32com.client as win32
import pythoncom
import json

import glob
import os
from pathlib import Path

# fetch-pandl no longer writes a fixed "PL_BS_GL_Detail.xlsx" - every refresh produces a new
# dated "PL_BS_GL_Detail_<timestamp>.xlsx" (see fetch-pandl/SKILL.md "File naming and version
# history"), so always pick up whichever one is newest in Files/.
# FILES_DIR is derived from this file's location so it survives a change of user profile.
FILES_DIR = str(Path(__file__).resolve().parents[4] / "Files")
_candidates = sorted(glob.glob(os.path.join(FILES_DIR, "PL_BS_GL_Detail_*.xlsx")))
if not _candidates:
    raise FileNotFoundError(f"No PL_BS_GL_Detail_*.xlsx found in {FILES_DIR} - run the fetch-pandl "
                             f"refresh first.")
SRC = _candidates[-1]
OUT_JSON = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pl_bs_gl_detail_values.json")

pythoncom.CoInitialize()
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = None
result = {}
try:
    wb = excel.Workbooks.Open(SRC, ReadOnly=True, UpdateLinks=0, IgnoreReadOnlyRecommended=True, CorruptLoad=1)
    excel.CalculateFullRebuild()

    def get_row(sheet, row, last_col=10):
        ws = wb.Sheets(sheet)
        rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, last_col))
        return list(rng.Value2)

    pl_rows = {
        "Total Turnover": 8, "Total Cost of Sales": 13, "Gross Profit": 18,
        "Total Other Income": 21, "Total Operating Expenses": 64, "EBITDA": 65,
        "Total Finance Cost": 68, "PBT": 70, "PAT": 72,
    }
    result["PL Summary"] = {k: get_row("PL Summary", r) for k, r in pl_rows.items()}

    bs_ws = wb.Sheets("BS Summary")
    bs_last_row = bs_ws.UsedRange.Rows.Count
    bs_labels = {}
    for r in range(1, bs_last_row + 1):
        v = bs_ws.Cells(r, 1).Value2
        if v:
            bs_labels[str(v).strip()] = r
    result["BS Summary labels"] = bs_labels
    for label in ["Total Liabilities", "Total Assets", "Sundry Debtors", "Sundry Creditors", "Bank OD (Overdraft)", "Secured Loans", "TOTAL", "Reconciliation"]:
        for lbl, r in bs_labels.items():
            if label.upper() in lbl.upper():
                result.setdefault("BS Summary", {})[lbl] = get_row("BS Summary", r)

    ratio_ws = wb.Sheets("Ratio Analysis")
    ratio_last = ratio_ws.UsedRange.Rows.Count
    ratio_vals = {}
    for r in range(1, ratio_last + 1):
        lbl = ratio_ws.Cells(r, 1).Value2
        if lbl:
            ratio_vals[str(lbl).strip()] = get_row("Ratio Analysis", r, last_col=10)
    result["Ratio Analysis"] = ratio_vals

    mom_ws = wb.Sheets("MOM PL and BS")
    mom_last = mom_ws.UsedRange.Rows.Count
    mom_vals = {}
    for r in range(1, mom_last + 1):
        lbl = mom_ws.Cells(r, 1).Value2
        if lbl and any(k in str(lbl).upper() for k in ["TURNOVER", "COST OF SALES", "GROSS PROFIT", "EBITDA", "PBT", "TOTAL LIABIL", "TOTAL ASSET"]):
            mom_vals[str(lbl).strip()] = get_row("MOM PL and BS", r, last_col=6)
    result["MOM PL and BS"] = mom_vals

    momgl_ws = wb.Sheets("MOM BS GL Detail")
    momgl_last = momgl_ws.UsedRange.Rows.Count
    momgl_vals = {}
    for r in range(1, momgl_last + 1):
        d = momgl_ws.Cells(r, 4).Value2
        if d and str(d).upper().startswith("TOTAL:"):
            momgl_vals[str(d).strip()] = get_row("MOM BS GL Detail", r, last_col=28)
    result["MOM BS GL Detail (TOTAL rows)"] = momgl_vals

    plc_ws = wb.Sheets("PL Summary")
    plc_last = plc_ws.UsedRange.Rows.Count
    entity_rows = {}
    for r in range(1, plc_last + 1):
        lbl = plc_ws.Cells(r, 1).Value2
        if lbl and str(lbl).strip() in ("Total Turnover", "Gross Profit", "Total Operating Expenses", "EBITDA", "Total Finance Cost", "PBT (Profit Before Tax)", "PAT (Net Profit)"):
            entity_rows[str(lbl).strip()] = r
    result["PL Summary row map"] = entity_rows

    with open(OUT_JSON, "w") as f:
        json.dump(result, f, indent=2, default=str)
    print("Saved to:", OUT_JSON)
finally:
    if wb is not None:
        wb.Close(SaveChanges=False)
    excel.Quit()
    del excel
