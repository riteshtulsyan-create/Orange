"""
build_analysis_sheets.py

Takes a workbook containing only a "RAW DATA" sheet (as produced by
generate_expense_sheet.py) and builds every analysis sheet on top of it via
Excel COM, using live PivotTables sharing one PivotCache so everything stays
refreshable and reconciles automatically:

  - Summary             EXP TYPE > MAIN GROUP  x MONTH  (whole dataset)
  - Direct              MAIN GROUP > TYPES     x MONTH  (DIRECT only)
  - Indirect Exp        MAIN GROUP > SUB-GROUP x MONTH  (INDIRECT only)
  - Direct by Component 3 pivots: Main Group x Month; Category>MainGroup>Vendor
                        drill-down; Main Group x Category cross-tab (DIRECT only)
  - One sheet per INDIRECT main group present in the data (incl. Advertisement):
    Vendor x Month + a Remarks column drafted from ledger narrations
  - Index               hyperlinks to every sheet, plus a back-link on each sheet

Then: reorders tabs, autofits row heights everywhere, and sorts + highlights
the Summary sheet (biggest expense first per EXP TYPE, with an AUDIT FLAG
column applying three fixed rules recomputed fresh from whatever the data
shows - not the specific findings from any one quarter):
  1. Top-5 main groups by absolute spend
  2. Any main group where one vendor/ledger is >70% of that group's total
  3. Main groups literally named "Direct Expenses" / "Indirect Expenses"
     (a structural catch-all-naming issue, not a one-off)

Usage (as a module):
    from build_analysis_sheets import build
    build(r"C:\\path\\to\\Expense Sheet.xlsx", from_date="20260401", to_date="20260630")

Usage (CLI, for testing against an existing RAW DATA-only file):
    python build_analysis_sheets.py --file "Expense Sheet.xlsx" --from 20260401 --to 20260630
"""

import os
import re
import sys
import shutil
import tempfile
import argparse
import subprocess
import pickle
from collections import defaultdict, Counter
from datetime import datetime

import openpyxl
import win32com.client as win32

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from remark_logic import draft_remark

# ---------------------------------------------------------------------------
# Excel COM constants
# ---------------------------------------------------------------------------
xlDatabase   = 1
xlRowField   = 1
xlColumnField = 2
xlPageField  = 3
xlDataField  = 4
xlSum        = -4157
xlAscending  = 1
xlUp         = -4162

HIGHLIGHT_COLOR = 255 + (230 << 8) + (179 << 16)  # BGR-packed soft amber


def month_range_labels(from_date: str, to_date: str) -> list[str]:
    """YYYYMMDD, YYYYMMDD -> ['APR-26', 'MAY-26', 'JUN-26', ...] chronological,
    matching the RAW DATA MONTH column's own =UPPER(TEXT(date,"MMM-YY")) formula."""
    d0 = datetime.strptime(from_date, "%Y%m%d")
    d1 = datetime.strptime(to_date, "%Y%m%d")
    labels = []
    y, m = d0.year, d0.month
    while (y, m) <= (d1.year, d1.month):
        labels.append(datetime(y, m, 1).strftime("%b-%y").upper())
        m += 1
        if m > 12:
            m = 1
            y += 1
    return labels


def _running_excel_pids():
    out = subprocess.run(
        ["powershell.exe", "-NoProfile", "-Command",
         "(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Id"],
        capture_output=True, text=True,
    ).stdout
    return set(int(x) for x in out.split())


def sanitize_sheet_name(name: str, existing: set) -> str:
    bad = set('\\/*?:[]')
    cleaned = "".join(ch for ch in name if ch not in bad).strip()
    cleaned = cleaned[:31]
    base = cleaned
    n = 2
    while cleaned in existing:
        suffix = f" {n}"
        cleaned = base[: 31 - len(suffix)] + suffix
        n += 1
    return cleaned


# ---------------------------------------------------------------------------
# Step 0: read RAW DATA with openpyxl to precompute everything COM doesn't
# need to (and shouldn't) recompute cell-by-cell: narrations for remarks,
# vendor concentration and main-group totals for the Summary highlight rules.
# ---------------------------------------------------------------------------

def analyze_raw_data(workbook_path: str, month_labels: list[str]):
    wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
    ws = wb["RAW DATA"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    idx = {h: i for i, h in enumerate(headers)}
    month_set = set(month_labels)

    exp_i, mg_i, sg_i, ledger_i = idx["EXP TYPE"], idx["MAIN GROUP"], idx["SUB-GROUP"], idx["LEDGER NAME"]
    net_i, date_i, narr_i, rem1_i = idx["NET AMOUNT"], idx["DATE"], idx["NARRATION"], idx["REMARK-1"]

    narrations = defaultdict(lambda: defaultdict(list))     # main_group -> vendor -> [narration,...]
    fallback   = defaultdict(lambda: defaultdict(list))     # main_group -> vendor -> [remark1,...]
    mg_total   = defaultdict(float)                          # (exp_type, main_group) -> total
    vendor_total = defaultdict(lambda: defaultdict(float))   # (exp_type, main_group) -> vendor -> total
    indirect_groups = set()

    for row in ws.iter_rows(min_row=2, values_only=True):
        # Derive the month from DATE rather than trusting the cached MONTH
        # formula result - a plain openpyxl re-save (by anything, anywhere in
        # the chain) discards cached formula values, which would otherwise
        # silently make every row look month-less.
        d = row[date_i]
        m = d.strftime("%b-%y").upper() if isinstance(d, datetime) else None
        if m not in month_set:
            continue
        exp_type, mg, vendor, net = row[exp_i], row[mg_i], row[ledger_i] or "(Blank)", row[net_i] or 0
        mg_total[(exp_type, mg)] += net
        vendor_total[(exp_type, mg)][vendor] += net
        if exp_type == "INDIRECT":
            indirect_groups.add(mg)
            narr = row[narr_i]
            rem1 = row[rem1_i]
            if narr and str(narr).strip():
                narrations[mg][vendor].append(str(narr).strip())
            elif rem1 and str(rem1).strip():
                fallback[mg][vendor].append(str(rem1).strip())

    return {
        "narrations": narrations,
        "fallback": fallback,
        "mg_total": mg_total,
        "vendor_total": vendor_total,
        "indirect_groups": sorted(indirect_groups),
    }


def make_remark_getter(analysis):
    def get_remark(mgroup, vendor):
        texts = analysis["narrations"].get(mgroup, {}).get(vendor, [])
        fb_texts = analysis["fallback"].get(mgroup, {}).get(vendor, [])
        fallback_str = fb_texts[0] if fb_texts else "Refer ledger/voucher for details"
        return draft_remark(texts, fallback_str)
    return get_remark


# ---------------------------------------------------------------------------
# COM helpers shared by every pivot builder
# ---------------------------------------------------------------------------

def pivot_geometry(pt):
    """Locate a pivot's row-label column/rows and Grand Total column using its
    own Row/TableRange properties - robust regardless of whether Excel shows
    the real field name or a generic 'Row Labels' caption (only one row field)."""
    row_range = pt.RowRange
    table_range = pt.TableRange1
    header_row = row_range.Row
    label_col = row_range.Column + row_range.Columns.Count - 1  # innermost row field
    last_row = table_range.Row + table_range.Rows.Count - 1
    grand_col = table_range.Column + table_range.Columns.Count - 1
    return header_row, label_col, last_row, grand_col


def write_remarks_column(ws, pt, mgroup, get_remark):
    header_row, label_col, last_row, grand_col = pivot_geometry(pt)
    remarks_col = grand_col + 1
    ws.Cells(header_row, remarks_col).Value = "REMARKS"
    ws.Cells(header_row, remarks_col).Font.Bold = True
    for r in range(header_row + 1, last_row):
        label = ws.Cells(r, label_col).Value
        if label is None or label == "":
            continue
        remark = get_remark(mgroup, label)
        cell = ws.Cells(r, remarks_col)
        cell.NumberFormat = "@"  # force text - some narrations look like dates
        cell.Value = remark
        cell.WrapText = True
        cell.VerticalAlignment = -4160  # xlTop
    ws.Columns(remarks_col).ColumnWidth = 60
    return remarks_col


def hide_out_of_range_months(pt, month_labels, field_name="MONTH"):
    """Defensive only: hides any MONTH item outside the requested range (e.g. a
    voucher with a missing/blank date) and orders the requested months
    chronologically left-to-right."""
    pf = pt.PivotFields(field_name)
    try:
        item_names = [pf.PivotItems(i + 1).Name for i in range(pf.PivotItems().Count)]
    except Exception:
        item_names = []
    for name in item_names:
        if name not in month_labels:
            try:
                pf.PivotItems(name).Visible = False
            except Exception:
                pass
    for pos, name in enumerate(month_labels, start=1):
        try:
            pf.PivotItems(name).Position = pos
        except Exception:
            pass


def set_tabular_layout(pt, row_fields):
    """Put each row field in its own column (e.g. EXP TYPE | MAIN GROUP side by
    side) instead of the default Compact form, which stacks every row field
    into one indented column. NOTE: PivotTable.RowAxisLayout(0) does NOT do
    this (verified empirically - it's a silent no-op here); the per-field
    PivotField.LayoutForm property is what actually works."""
    for rf in row_fields:
        try:
            pt.PivotFields(rf).LayoutForm = 0  # xlTabularRow
        except Exception as e:
            print(f"  (LayoutForm failed for {rf}: {e})")


def build(workbook_path: str, from_date: str, to_date: str):
    month_labels = month_range_labels(from_date, to_date)
    print("Month scope:", month_labels)

    analysis = analyze_raw_data(workbook_path, month_labels)
    get_remark = make_remark_getter(analysis)
    print("Indirect main groups found:", analysis["indirect_groups"])

    # Work on a local, non-cloud-synced copy. Opening the live OneDrive/
    # SharePoint path directly can hit a transient co-authoring/sync share
    # violation - Excel then silently falls back to ReadOnly and Save()
    # no-ops without ever raising an error, so nothing gets written back.
    local_dir = tempfile.mkdtemp(prefix="expense_build_")
    local_path = os.path.join(local_dir, os.path.basename(workbook_path))
    shutil.copy2(workbook_path, local_path)

    pids_before = _running_excel_pids()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(local_path)
    pids_after = _running_excel_pids()
    new_pids = pids_after - pids_before
    print("Excel PID(s):", new_pids)

    if wb.ReadOnly:
        raise RuntimeError(f"Local working copy {local_path} opened ReadOnly - unexpected.")

    try:
        _clear_existing_analysis_sheets(wb)
        _run_build(wb, month_labels, analysis, get_remark)
        wb.Save()
        print("Saved local working copy:", local_path)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        sys.stdout.flush()
        for pid in new_pids:
            subprocess.run(["powershell.exe", "-NoProfile", "-Command",
                             f"Stop-Process -Id {pid} -Force -ErrorAction SilentlyContinue"])

    shutil.copy2(local_path, workbook_path)
    shutil.rmtree(local_dir, ignore_errors=True)
    print("SAVED:", workbook_path, "(mtime:", os.path.getmtime(workbook_path), ")")


def _clear_existing_analysis_sheets(wb):
    """Make re-running the build idempotent: drop every sheet except RAW DATA
    so add_sheet() never collides with a sheet name left over from a prior
    run (e.g. re-scoping the month range on an already-built workbook)."""
    for ws in list(wb.Worksheets):
        if ws.Name != "RAW DATA":
            ws.Delete()


def _run_build(wb, month_labels, analysis, get_remark):
    ws_raw = wb.Worksheets("RAW DATA")
    ws_raw.Activate()
    ws_raw.Range("A1").Select()  # clear any multi-area cell selection left over
                                  # from a prior save - Worksheets.Add() fails
                                  # with "won't work on multiple selections"
                                  # otherwise
    last_row = ws_raw.Cells(ws_raw.Rows.Count, 1).End(xlUp).Row
    last_col = ws_raw.Cells(1, ws_raw.Columns.Count).End(-4159).Column  # xlToLeft
    src_range = ws_raw.Range(ws_raw.Cells(1, 1), ws_raw.Cells(last_row, last_col))
    print("Source range:", src_range.Address)

    shared_cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=src_range)

    pivot_counter = [0]
    def next_pt_name():
        pivot_counter[0] += 1
        return f"PT_{pivot_counter[0]}"

    def add_sheet(name):
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = name
        return ws

    def new_pivot(ws, dest_cell):
        return shared_cache.CreatePivotTable(TableDestination=dest_cell, TableName=next_pt_name())

    def finish_numberformat(pt):
        try:
            pt.DataBodyRange.NumberFormat = "#,##0.00"
        except Exception as e:
            print("  (number format failed:", e, ")")

    # -----------------------------------------------------------------
    # Summary: EXP TYPE > MAIN GROUP  x  MONTH  (whole dataset)
    # -----------------------------------------------------------------
    ws_summary = add_sheet("Summary")
    pt = new_pivot(ws_summary, ws_summary.Range("A1"))
    pt.ManualUpdate = True
    pt.PivotFields("EXP TYPE").Orientation = xlRowField
    pt.PivotFields("EXP TYPE").Position = 1
    pt.PivotFields("MAIN GROUP").Orientation = xlRowField
    pt.PivotFields("MAIN GROUP").Position = 2
    pt.PivotFields("MONTH").Orientation = xlColumnField
    pt.AddDataField(pt.PivotFields("NET AMOUNT"), "Sum of NET AMOUNT", xlSum)
    pt.ManualUpdate = False
    set_tabular_layout(pt, ["EXP TYPE", "MAIN GROUP"])
    hide_out_of_range_months(pt, month_labels)
    finish_numberformat(pt)
    print("Summary pivot built.")

    # -----------------------------------------------------------------
    # Direct: MAIN GROUP > TYPES  x  MONTH  (DIRECT only)
    # -----------------------------------------------------------------
    ws_direct = add_sheet("Direct")
    pt = new_pivot(ws_direct, ws_direct.Range("A1"))
    pt.ManualUpdate = True
    pt.PivotFields("EXP TYPE").Orientation = xlPageField
    pt.PivotFields("EXP TYPE").CurrentPage = "DIRECT"
    pt.PivotFields("MAIN GROUP").Orientation = xlRowField
    pt.PivotFields("MAIN GROUP").Position = 1
    pt.PivotFields("TYPES").Orientation = xlRowField
    pt.PivotFields("TYPES").Position = 2
    pt.PivotFields("MONTH").Orientation = xlColumnField
    pt.AddDataField(pt.PivotFields("NET AMOUNT"), "Sum of NET AMOUNT", xlSum)
    pt.ManualUpdate = False
    set_tabular_layout(pt, ["MAIN GROUP", "TYPES"])
    hide_out_of_range_months(pt, month_labels)
    pt.PivotFields("MAIN GROUP").AutoSort(xlAscending, "Sum of NET AMOUNT")
    finish_numberformat(pt)
    print("Direct pivot built.")

    # -----------------------------------------------------------------
    # Indirect Exp: MAIN GROUP > SUB-GROUP  x  MONTH  (INDIRECT only)
    # -----------------------------------------------------------------
    ws_ind = add_sheet("Indirect Exp")
    pt = new_pivot(ws_ind, ws_ind.Range("A1"))
    pt.ManualUpdate = True
    pt.PivotFields("EXP TYPE").Orientation = xlPageField
    pt.PivotFields("EXP TYPE").CurrentPage = "INDIRECT"
    pt.PivotFields("MAIN GROUP").Orientation = xlRowField
    pt.PivotFields("MAIN GROUP").Position = 1
    pt.PivotFields("SUB-GROUP").Orientation = xlRowField
    pt.PivotFields("SUB-GROUP").Position = 2
    pt.PivotFields("MONTH").Orientation = xlColumnField
    pt.AddDataField(pt.PivotFields("NET AMOUNT"), "Sum of NET AMOUNT", xlSum)
    pt.ManualUpdate = False
    set_tabular_layout(pt, ["MAIN GROUP", "SUB-GROUP"])
    hide_out_of_range_months(pt, month_labels)
    pt.PivotFields("MAIN GROUP").AutoSort(xlAscending, "Sum of NET AMOUNT")
    finish_numberformat(pt)
    print("Indirect Exp pivot built.")

    # -----------------------------------------------------------------
    # One sheet per INDIRECT main group (incl. Advertisement)
    # -----------------------------------------------------------------
    existing_names = {s.Name for s in wb.Sheets}
    group_sheet_map = {}
    for mg in analysis["indirect_groups"]:
        display_name = "Advertisement" if mg == "ADVERTISING PROMOTION EXP" else mg.title()
        sname = sanitize_sheet_name(display_name, existing_names)
        existing_names.add(sname)
        group_sheet_map[mg] = sname

        ws = add_sheet(sname)
        pt = new_pivot(ws, ws.Range("A1"))
        pt.ManualUpdate = True
        pt.PivotFields("EXP TYPE").Orientation = xlPageField
        pt.PivotFields("EXP TYPE").CurrentPage = "INDIRECT"
        pt.PivotFields("MAIN GROUP").Orientation = xlPageField
        pt.PivotFields("MAIN GROUP").CurrentPage = mg
        pt.PivotFields("LEDGER NAME").Orientation = xlRowField
        pt.PivotFields("MONTH").Orientation = xlColumnField
        pt.AddDataField(pt.PivotFields("NET AMOUNT"), "Sum of NET AMOUNT", xlSum)
        pt.ManualUpdate = False
        hide_out_of_range_months(pt, month_labels)
        pt.PivotFields("LEDGER NAME").AutoSort(xlAscending, "Sum of NET AMOUNT")
        finish_numberformat(pt)

        write_remarks_column(ws, pt, mg, get_remark)
        ws.Columns(1).ColumnWidth = 38
        print("Built indirect sheet:", sname, "(", mg, ")")

    # -----------------------------------------------------------------
    # Direct by Component: 3 pivots
    # -----------------------------------------------------------------
    ws_dc = add_sheet("Direct by Component")

    def build_direct_pivot(dest_row, row_fields, col_field="MONTH"):
        pt = new_pivot(ws_dc, ws_dc.Cells(dest_row, 1))
        pt.ManualUpdate = True
        pt.PivotFields("EXP TYPE").Orientation = xlPageField
        pt.PivotFields("EXP TYPE").CurrentPage = "DIRECT"
        for i, rf in enumerate(row_fields, start=1):
            pt.PivotFields(rf).Orientation = xlRowField
            try:
                pt.PivotFields(rf).Position = i
            except Exception:
                pass
        pt.PivotFields(col_field).Orientation = xlColumnField
        pt.AddDataField(pt.PivotFields("NET AMOUNT"), "Sum of NET AMOUNT", xlSum)
        pt.ManualUpdate = False
        if len(row_fields) > 1:
            set_tabular_layout(pt, row_fields)
        if col_field == "MONTH":
            hide_out_of_range_months(pt, month_labels, col_field)
        pt.PivotFields(row_fields[0]).AutoSort(xlAscending, "Sum of NET AMOUNT")
        finish_numberformat(pt)
        return pt

    def next_free_row(ws):
        used = ws.UsedRange
        if used.Cells.Count == 1 and used.Cells(1, 1).Value is None:
            return 1
        return used.Row + used.Rows.Count + 2

    pt1 = build_direct_pivot(1, ["MAIN GROUP"])
    print("Direct-by-Component Pivot1 (Main Group x Month) built.")

    r = next_free_row(ws_dc)
    pt2 = build_direct_pivot(r, ["COMPONENT CATEGORY", "MAIN GROUP", "LEDGER NAME"])
    print("Direct-by-Component Pivot2 (Category>MainGroup>Vendor) built.")

    r = next_free_row(ws_dc)
    pt3 = build_direct_pivot(r, ["MAIN GROUP"], col_field="COMPONENT CATEGORY")
    pt3.PivotFields("MONTH").Orientation = xlPageField
    try:
        item_names = [pt3.PivotFields("MONTH").PivotItems(i + 1).Name
                       for i in range(pt3.PivotFields("MONTH").PivotItems().Count)]
        for name in item_names:
            if name not in month_labels:
                pt3.PivotFields("MONTH").PivotItems(name).Visible = False
    except Exception as e:
        print("  (pt3 month filter failed:", e, ")")
    print("Direct-by-Component Pivot3 (MainGroup x Category cross-tab) built.")

    ws_dc.Columns(1).ColumnWidth = 30
    ws_dc.Columns(2).ColumnWidth = 26
    ws_dc.Columns(3).ColumnWidth = 30

    # -----------------------------------------------------------------
    # Reorder tabs (before building Index, so Index lists sheets in the same
    # order the tabs actually appear in): Summary, Direct, Direct by Component,
    # Indirect Exp, Advertisement, [other indirect sheets alphabetically],
    # RAW DATA last. Index itself is inserted at position 1 afterwards.
    # -----------------------------------------------------------------
    other_indirect = sorted(
        sname for mg, sname in group_sheet_map.items() if mg != "ADVERTISING PROMOTION EXP"
    )
    target_order = ["Summary", "Direct", "Direct by Component", "Indirect Exp"]
    if "ADVERTISING PROMOTION EXP" in group_sheet_map:
        target_order.append(group_sheet_map["ADVERTISING PROMOTION EXP"])
    target_order += other_indirect + ["RAW DATA"]

    for i in range(1, len(target_order)):
        wb.Worksheets(target_order[i]).Move(After=wb.Worksheets(target_order[i - 1]))
    print("Reordered sheets:", [s.Name for s in wb.Worksheets])

    # -----------------------------------------------------------------
    # Sort Summary ascending (biggest expense first per EXP TYPE) and apply
    # rule-based AUDIT FLAG highlighting. Done before the Index is built so
    # the Index can point at exactly which sheets carry a flagged main group.
    # -----------------------------------------------------------------
    flagged_map = _sort_and_highlight_summary(ws_summary, analysis)

    # Roll each flagged (exp_type, main_group) up to the sheet that shows its
    # detail, so the Index can say "this is the sheet with the flagged item".
    sheet_flags = defaultdict(list)
    for (exp_type, mg), reason in flagged_map.items():
        if exp_type == "DIRECT":
            for sname in ("Direct", "Direct by Component"):
                sheet_flags[sname].append(f"{mg}: {reason}")
        else:
            sname = group_sheet_map.get(mg, "Indirect Exp")
            sheet_flags[sname].append(f"{mg}: {reason}")

    # -----------------------------------------------------------------
    # Index sheet - inserted at position 1, listing every other sheet in
    # its now-final tab order
    # -----------------------------------------------------------------
    ws_idx = wb.Worksheets.Add(Before=wb.Worksheets(1))
    ws_idx.Name = "Index"
    ws_idx.Cells(1, 1).Value = "Sheet"
    ws_idx.Cells(1, 1).Font.Bold = True
    ws_idx.Cells(1, 2).Value = "Description"
    ws_idx.Cells(1, 2).Font.Bold = True
    ws_idx.Cells(1, 3).Value = "Must Review"
    ws_idx.Cells(1, 3).Font.Bold = True

    descriptions = {
        "Summary": "Overall Direct vs Indirect expense summary",
        "Direct": "Direct expense pivot by Main Group / Types",
        "Indirect Exp": "Indirect expense pivot by Main Group / Sub-Group",
        "Direct by Component": "Direct expense broken down by component category (Head/Ink/Machine/Spare/Fabric) and vendor",
        "RAW DATA": "Raw transaction-level data - the source for every pivot in this workbook",
    }
    for mg, sname in group_sheet_map.items():
        descriptions.setdefault(sname, f"Indirect expense: {mg} - vendor-wise spend with remarks")

    row = 2
    for ws in wb.Worksheets:
        if ws.Name == "Index":
            continue
        ws_idx.Cells(row, 1).Value = ws.Name
        ws_idx.Hyperlinks.Add(Anchor=ws_idx.Cells(row, 1), Address="",
                               SubAddress=f"'{ws.Name}'!A1", TextToDisplay=ws.Name)
        ws_idx.Cells(row, 2).Value = descriptions.get(ws.Name, "")
        reasons = sheet_flags.get(ws.Name)
        if reasons:
            cell3 = ws_idx.Cells(row, 3)
            cell3.Value = "; ".join(reasons)
            cell3.WrapText = True
            cell3.VerticalAlignment = -4160
            ws_idx.Range(ws_idx.Cells(row, 1), ws_idx.Cells(row, 3)).Interior.Color = HIGHLIGHT_COLOR
        row += 1
    ws_idx.Columns(1).ColumnWidth = 32
    ws_idx.Columns(2).ColumnWidth = 70
    ws_idx.Columns(3).ColumnWidth = 60

    for ws in wb.Worksheets:
        if ws.Name in ("Index", "RAW DATA"):
            continue
        cell = ws.Cells(1, 10)
        cell.Value = "<< Back to Index"
        ws.Hyperlinks.Add(Anchor=cell, Address="", SubAddress="Index!A1", TextToDisplay="<< Back to Index")

    # -----------------------------------------------------------------
    # Autofit row heights everywhere
    # -----------------------------------------------------------------
    for ws in wb.Worksheets:
        try:
            ws.UsedRange.Rows.AutoFit()
        except Exception as e:
            print("  (autofit failed for", ws.Name, e, ")")

    ws_idx.Activate()


def _sort_and_highlight_summary(ws_summary, analysis):
    pt = ws_summary.PivotTables(1)
    pt.PivotFields("MAIN GROUP").AutoSort(xlAscending, "Sum of NET AMOUNT")
    print("Summary sorted ascending (biggest expense first) within each EXP TYPE.")

    mg_total = analysis["mg_total"]          # (exp_type, main_group) -> total
    vendor_total = analysis["vendor_total"]  # (exp_type, main_group) -> vendor -> total

    grand_total_abs = sum(abs(v) for v in mg_total.values())
    top5_keys = {k for k, _ in sorted(mg_total.items(), key=lambda kv: -abs(kv[1]))[:5]}

    def flag_reason(exp_type, mg):
        reasons = []
        total = mg_total.get((exp_type, mg), 0)
        if (exp_type, mg) in top5_keys and grand_total_abs:
            pct = abs(total) / grand_total_abs * 100
            reasons.append(f"Top-5 cost driver - {pct:.0f}% of total spend")
        vendors = vendor_total.get((exp_type, mg), {})
        if vendors and total:
            top_vendor, top_amt = min(vendors.items(), key=lambda kv: kv[1]) \
                if total < 0 else max(vendors.items(), key=lambda kv: kv[1])
            share = abs(top_amt) / abs(total) * 100 if total else 0
            if share > 70:
                reasons.append(f"{share:.0f}% concentrated in one ledger ({top_vendor})")
        if mg in ("Direct Expenses", "Indirect Expenses"):
            reasons.append("Catch-all bucket name - review the underlying postings")
        return "; ".join(reasons)

    header_row, mg_col, last_row, grand_col = pivot_geometry(pt)
    flag_col = grand_col + 1
    ws_summary.Cells(header_row, flag_col).Value = "AUDIT FLAG"
    ws_summary.Cells(header_row, flag_col).Font.Bold = True

    exp_type_col = mg_col - 1
    cur_exp_type = None
    flagged = 0
    flagged_map = {}  # (exp_type, main_group) -> reason, for the Index sheet to consume
    for r in range(header_row + 1, last_row):
        et = ws_summary.Cells(r, exp_type_col).Value
        if et:
            cur_exp_type = et
        mg = ws_summary.Cells(r, mg_col).Value
        if mg is None:
            continue
        reason = flag_reason(cur_exp_type, mg)
        if reason:
            ws_summary.Range(ws_summary.Cells(r, 1), ws_summary.Cells(r, grand_col)).Interior.Color = HIGHLIGHT_COLOR
            cell = ws_summary.Cells(r, flag_col)
            cell.Value = reason
            cell.WrapText = True
            cell.VerticalAlignment = -4160
            flagged += 1
            flagged_map[(cur_exp_type, mg)] = reason
    ws_summary.Columns(flag_col).ColumnWidth = 55
    ws_summary.UsedRange.Rows.AutoFit()
    print(f"Summary AUDIT FLAG applied to {flagged} rows.")
    return flagged_map


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Build analysis sheets on top of a RAW DATA-only workbook")
    parser.add_argument("--file", required=True, help="Path to the .xlsx (must contain a RAW DATA sheet)")
    parser.add_argument("--from", dest="from_date", required=True, help="YYYYMMDD")
    parser.add_argument("--to", dest="to_date", required=True, help="YYYYMMDD")
    args = parser.parse_args()
    build(os.path.abspath(args.file), args.from_date, args.to_date)
    print("Done.")
