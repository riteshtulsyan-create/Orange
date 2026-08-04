"""
last_entry_check.py

Finds the most-recently-entered voucher in Tally for each of the 5 companies
(highest ALTERID - Tally's own internal edit/creation sequence counter, which
tracks recency of data-entry action regardless of the voucher's own
accounting DATE) and writes/refreshes a "Last Entry Log" sheet in the Expense
Sheet workbook: Company | Last Entry Date | Voucher Type | Voucher No. |
Entered By | Altered By | Altered On.

IMPORTANT LIMITATION (checked directly against this Tally setup on
2026-07-21): Tally exposes ENTEREDBY/ALTEREDBY as real fields, and ALTEREDON
as a DATE, but VOUCHERTIME is present in the schema and empty for every
voucher across all 5 companies - this Tally setup does not record a
clock-time for when an entry was made, only the day. So "Last Entry Date" is
a date, not a date+time - there is no wall-clock timestamp to show even for
the very last entry made minutes ago.

Usage (as a module):
    from last_entry_check import fetch_all_last_entries, add_last_entry_sheet
    rows = fetch_all_last_entries()
    add_last_entry_sheet(workbook_path, rows)

Usage (CLI - fetch and print only, no workbook write):
    python scripts/last_entry_check.py
"""

import os
import sys
import shutil
import tempfile
import subprocess
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fetch_tally_expenses import TALLY_URL, _clean_xml, _post_tally
from generate_expense_sheet import companies_for_period

import requests
import win32com.client as win32

_LAST_ENTRY_XML = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>LastEntryCheck</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="LastEntryCheck" ISMODIFY="No">
            <TYPE>Voucher</TYPE>
            <FETCH>DATE,ALTERID,ENTEREDBY,ALTEREDBY,ALTEREDON,VOUCHERTYPENAME,VOUCHERNUMBER</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""


def _alterid(v):
    t = (v.findtext("ALTERID") or "0").strip()
    try:
        return int(t)
    except ValueError:
        return 0


def _fmt_date(d):
    d = (d or "").strip()
    if len(d) == 8 and d.isdigit():
        return f"{d[6:8]}-{d[4:6]}-{d[0:4]}"
    return d


def fetch_last_entry(tally_name: str) -> dict:
    """Returns the highest-ALTERID (most recently entered/edited) voucher's
    details for one company, across its full history (no date filter)."""
    xml_body = _LAST_ENTRY_XML.format(company=tally_name)
    root = _post_tally(xml_body, f"LastEntry/{tally_name[:30]}")
    vouchers = root.findall(".//VOUCHER")
    if not vouchers:
        return None
    top = max(vouchers, key=_alterid)
    return {
        "date":         _fmt_date(top.findtext("DATE")),
        "vch_type":     (top.findtext("VOUCHERTYPENAME") or "").strip(),
        "vch_no":       (top.findtext("VOUCHERNUMBER") or "").strip(),
        "entered_by":   (top.findtext("ENTEREDBY") or "").strip(),
        "altered_by":   (top.findtext("ALTEREDBY") or "").strip(),
        "altered_on":   _fmt_date(top.findtext("ALTEREDON")),
        "alter_id":     _alterid(top),
    }


def fetch_all_last_entries(open_names_lower: set = None, period: str = "2627") -> list[dict]:
    rows = []
    for cfg in companies_for_period(period):
        if open_names_lower is not None and cfg["tally_name"].lower() not in open_names_lower:
            continue
        info = fetch_last_entry(cfg["tally_name"])
        if info is None:
            continue
        info["company_col"] = cfg["company_col"]
        rows.append(info)
    return rows


def add_last_entry_sheet(workbook_path: str, rows: list[dict]):
    """Adds/refreshes a 'Last Entry Log' sheet in the given workbook, placed
    right after Index, and adds a link to it from Index if not already
    listed. Works on a local scratch copy first (same reasoning as
    build_analysis_sheets.build) to avoid a OneDrive/SharePoint ReadOnly
    lock silently no-op'ing the save."""
    local_dir = tempfile.mkdtemp(prefix="last_entry_")
    local_path = os.path.join(local_dir, os.path.basename(workbook_path))
    shutil.copy2(workbook_path, local_path)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(local_path)

    try:
        sheet_names = [s.Name for s in wb.Worksheets]
        if "Last Entry Log" in sheet_names:
            wb.Worksheets("Last Entry Log").Delete()

        if "Index" in sheet_names:
            ws = wb.Worksheets.Add(After=wb.Worksheets("Index"))
        else:
            ws = wb.Worksheets.Add(Before=wb.Worksheets(1))
        ws.Name = "Last Entry Log"

        headers = ["Company", "Last Entry Date", "Voucher Type", "Voucher No.",
                   "Entered By", "Altered By", "Altered On"]
        for col, h in enumerate(headers, start=1):
            c = ws.Cells(1, col)
            c.Value = h
            c.Font.Bold = True

        ws.Columns(2).NumberFormat = "@"
        ws.Columns(7).NumberFormat = "@"
        for r, row in enumerate(rows, start=2):
            ws.Cells(r, 1).Value = row["company_col"]
            ws.Cells(r, 2).Value = row["date"]
            ws.Cells(r, 3).Value = row["vch_type"]
            ws.Cells(r, 4).Value = row["vch_no"]
            ws.Cells(r, 5).Value = row["entered_by"]
            ws.Cells(r, 6).Value = row["altered_by"]
            ws.Cells(r, 7).Value = row["altered_on"]

        note_row = len(rows) + 3
        ws.Cells(note_row, 1).Value = (
            "Note: Tally does not record a clock-time for voucher entries in this "
            "setup (only the day) - \"Last Entry Date\" is the day that company's "
            "most recently created/edited voucher was touched, not a timestamp."
        )
        ws.Cells(note_row, 1).Font.Italic = True

        for col, width in enumerate([16, 15, 22, 20, 14, 14, 14], start=1):
            ws.Columns(col).ColumnWidth = width
        ws.UsedRange.Rows.AutoFit()

        if "Index" in sheet_names:
            ws_idx = wb.Worksheets("Index")
            last_row = ws_idx.Cells(ws_idx.Rows.Count, 1).End(-4162).Row
            existing = [ws_idx.Cells(r, 1).Value for r in range(2, last_row + 1)]
            if "Last Entry Log" not in existing:
                new_row = last_row + 1
                ws_idx.Cells(new_row, 1).Value = "Last Entry Log"
                ws_idx.Hyperlinks.Add(Anchor=ws_idx.Cells(new_row, 1), Address="",
                                       SubAddress="'Last Entry Log'!A1", TextToDisplay="Last Entry Log")
                ws_idx.Cells(new_row, 2).Value = (
                    "Most recently entered/edited voucher per company, with date and username"
                )
            cell = ws.Cells(1, 10)
            cell.Value = "<< Back to Index"
            ws.Hyperlinks.Add(Anchor=cell, Address="", SubAddress="Index!A1", TextToDisplay="<< Back to Index")

        wb.Save()
    finally:
        try:
            excel.Quit()
        except Exception:
            pass

    shutil.copy2(local_path, workbook_path)
    shutil.rmtree(local_dir, ignore_errors=True)


if __name__ == "__main__":
    for row in fetch_all_last_entries():
        print(row)
