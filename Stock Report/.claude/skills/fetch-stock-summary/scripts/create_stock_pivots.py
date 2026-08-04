#!/usr/bin/env python3
"""
create_stock_pivots.py
Adds native Excel PivotTables to a stock summary workbook produced by
fetch_stock_summary.py, in place.

The layout reproduces the pivots the user built by hand on
Stock_Summary_..._2026-07-30_1659.xlsx, plus one sheet per Main Group:

    Summary        rows: Main Category
    Group          rows: Main Category > Sub Category
    Sub Group      rows: Main Category > Sub Category > Product Group
    Machines       rows: Sub Category > Product Group > Item Name,
                   Main Category page field locked to MACHINERY & MACHINERY PARTS
    Printing Ink   as Machines, locked to PRINTING INK
    Paper Roll     as Machines, locked to PAPER ROLL
    Software       as Machines, locked to SOFTWARE

Every pivot measures Sum of Quantity and Sum of Value, and carries a Unit page
field so a single unit of measure can be isolated -- Sum of Quantity otherwise
adds KGS + PCS + MTR + ROLL together and is not a meaningful figure.

Requires Excel (COM automation via xlwings), same approach as this project's
fetch-bank-transactions skill.

Usage:
    python create_stock_pivots.py <workbook.xlsx>
"""

import sys
from pathlib import Path

import openpyxl

# VBA / COM integer constants
XL_ROW, XL_COL, XL_PAGE, XL_DATA = 1, 2, 3, 4
XL_DB = 1
XL_SUM = -4157
XL_TAB = 1          # xlTabularRow  (side-by-side row fields)
XL_COMPACT = 0      # xlCompactRow  ("Row Labels" in one column)

DATA_SHEET = "Item List (Flat)"
VALUE_FMT = "#,##0.00"
QTY_FMT = "#,##0.000"

# Main Groups that get their own sheet: (Tally group name, sheet tab name).
# Only groups carrying value are built -- see pick_group_sheets().
GROUP_SHEETS = [
    ("MACHINERY & MACHINERY PARTS", "Machines"),
    ("PRINTING INK", "Printing Ink"),
    ("PAPER ROLL", "Paper Roll"),
    ("SOFTWARE", "Software"),
]


def pick_group_sheets(xlsx: Path):
    """
    Return [(main_group, sheet_name)] for Main Groups that actually carry value,
    in descending value order. A group with nil value gets no sheet -- an empty
    pivot is noise, and Cross Check already records it.
    """
    wb = openpyxl.load_workbook(xlsx, read_only=True)
    ws = wb[DATA_SHEET]
    rows = ws.iter_rows(min_row=2, values_only=True)
    totals, units = {}, set()
    for r in rows:
        if r[0] is None:
            continue
        totals[r[0]] = totals.get(r[0], 0.0) + (r[9] or 0.0)
        if r[7]:
            units.add(r[7])
    wb.close()

    known = dict(GROUP_SHEETS)
    out = []
    for name, val in sorted(totals.items(), key=lambda kv: -kv[1]):
        if round(val, 2) == 0:
            continue
        tab = known.get(name)
        if tab is None:                       # a Main Group we have not named yet
            tab = name.title()[:31]
        out.append((name, tab, val))
    return out, sorted(units)


def add_pivot(wb, cache_factory, sheet_name, before_sheet,
              row_fields, page_fields, locked=None, compact=False, title=None):
    """Create one sheet with one PivotTable on it. Returns the xlwings sheet."""
    # Rebuild from scratch so the script is idempotent -- re-running on a
    # workbook that already has these sheets replaces them rather than failing.
    for existing in list(wb.sheets):
        if existing.name == sheet_name:
            existing.delete()
    ws = wb.sheets.add(sheet_name, before=before_sheet)

    # Always anchor the pivot at A1. Anchoring lower to leave room for a title
    # collides with the page-field block Excel inserts above the table, which
    # leaves the sheet holding a #VALUE! error instead of a pivot. The title
    # goes to one side instead.
    pt = cache_factory().CreatePivotTable(ws.range("A1").api,
                                          f"PT_{abs(hash(sheet_name)) % 100000}")
    pt.RowAxisLayout(XL_COMPACT if compact else XL_TAB)
    pt.RowGrand = True
    pt.ColumnGrand = True

    for pos, fname in enumerate(row_fields, 1):
        f = pt.PivotFields(fname)
        f.Orientation = XL_ROW
        f.Position = pos

    for fname in page_fields:
        f = pt.PivotFields(fname)
        f.Orientation = XL_PAGE
        if locked and fname == locked[0]:
            # Lock this sheet to one Main Group, the way the hand-built
            # 'Machines' sheet is locked to MACHINERY & MACHINERY PARTS.
            f.CurrentPage = locked[1]

    q = pt.AddDataField(pt.PivotFields("Quantity"), "Sum of Quantity", XL_SUM)
    q.NumberFormat = QTY_FMT
    v = pt.AddDataField(pt.PivotFields("Value"), "Sum of Value", XL_SUM)
    v.NumberFormat = VALUE_FMT

    ws.autofit(axis="columns")
    if title:
        cell = ws.range("G1")
        cell.value = title
        cell.api.Font.Bold = True
        cell.api.Font.Size = 13
    return ws


SHEET_NOTES = {
    "Index": "This sheet - what is in the workbook and where to find it.",
    "Summary": "Closing value and quantity by Main Category. Start here.",
    "Group": "Main Category > Sub Category. Filter by Product Group, Sub Group or Item.",
    "Sub Group": "Main Category > Sub Category > Product Group. One level deeper.",
    "Machines": "Every item under MACHINERY & MACHINERY PARTS.",
    "Printing Ink": "Every item under PRINTING INK.",
    "Paper Roll": "Every item under PAPER ROLL.",
    "Software": "Every item under SOFTWARE.",
    "Item List (Flat)": "Source data - one row per stock item. All pivots read from here.",
    "Cross Check": "Reconciliation and data-quality flags. Read before quoting any figure.",
    "Hierarchy": "Indented tree mirroring the Tally Stock Summary screen.",
    "Group Summary": "Every stock group with its level, path and rolled-up value.",
}


def read_meta(xlsx: Path):
    """Pull company / period / totals / exception counts off the Cross Check sheet."""
    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    meta, flags = {}, []
    if "Cross Check" in wb.sheetnames:
        for r in wb["Cross Check"].iter_rows(min_row=2, values_only=True):
            if r[0] is None:
                break
            meta[str(r[0])] = r[1]
            if r[2] in ("REVIEW", "MISMATCH") and r[1] not in (0, None):
                flags.append((str(r[0]), r[1], r[2]))
    wb.close()
    return meta, flags


def build_index(wb, xw, xlsx, groups, meta, flags):
    """
    Front sheet: a plain navigation list, one clickable link per sheet.

    Deliberately carries no company name, period, figures or exception counts.
    This workbook is produced for every company in Tally, so the index has to
    stay identical across all of them -- anything company-specific belongs on
    'Cross Check', which already holds company, period, generated timestamp,
    totals and every data-quality flag.
    """
    for existing in list(wb.sheets):
        if existing.name == "Index":
            existing.delete()
    ws = wb.sheets.add("Index", before=wb.sheets[0])

    ws.range("A1").value = "INDEX"
    ws.range("A1").api.Font.Size = 16
    ws.range("A1").api.Font.Bold = True

    r = 3
    for c, h in enumerate(["#", "Sheet", "What it shows", "Go to"], start=1):
        ws.range((r, c)).value = h
    ws.range((r, 1), (r, 4)).api.Font.Bold = True
    r += 1

    for i, sh in enumerate([s.name for s in wb.sheets], start=1):
        ws.range((r, 1)).value = i
        ws.range((r, 2)).value = sh
        ws.range((r, 3)).value = SHEET_NOTES.get(sh, "")
        if sh != "Index":
            ws.range((r, 4)).api.Hyperlinks.Add(
                Anchor=ws.range((r, 4)).api, Address="",
                SubAddress=f"'{sh}'!A1", TextToDisplay="open")
        r += 1

    ws.range("A:A").column_width = 5
    ws.range("B:B").column_width = 22
    ws.range("C:C").column_width = 66
    ws.range("D:D").column_width = 10
    ws.range("A1").select()
    return ws


def main():
    if len(sys.argv) < 2:
        print("usage: python create_stock_pivots.py <workbook.xlsx>", file=sys.stderr)
        return 2
    xlsx = Path(sys.argv[1]).resolve()
    if not xlsx.is_file():
        print(f"ERROR: not found: {xlsx}", file=sys.stderr)
        return 1

    groups, units = pick_group_sheets(xlsx)
    # Read everything openpyxl needs BEFORE Excel opens the file -- once Excel
    # has it open the lock is exclusive and openpyxl gets a PermissionError.
    meta, flags = read_meta(xlsx)
    print(f"Main Groups with value: {len(groups)}  |  distinct units: {len(units)}")

    try:
        import xlwings as xw
    except ImportError:
        print("ERROR: xlwings is not installed - cannot build native PivotTables.\n"
              "       The data workbook is still valid; install xlwings and re-run.",
              file=sys.stderr)
        return 1

    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(str(xlsx))
        src = wb.sheets[DATA_SHEET]
        last_row = src.range("A1").end("down").row
        n_cols = src.range("A1").end("right").column
        src_ref = f"'{DATA_SHEET}'!R1C1:R{last_row}C{n_cols}"
        print(f"Pivot source: {src_ref}")

        # A fresh cache per pivot -- shared caches make Excel couple the pivots'
        # grouping/refresh state, which is the same reason the bank-transactions
        # skill builds one per table.
        def cache():
            return wb.api.PivotCaches().Create(XL_DB, src_ref)

        anchor = wb.sheets[DATA_SHEET]

        add_pivot(wb, cache, "Summary", anchor,
                  row_fields=["Main Category"],
                  page_fields=["Unit"],
                  compact=True, title="Final Summary")

        add_pivot(wb, cache, "Group", anchor,
                  row_fields=["Main Category", "Sub Category"],
                  page_fields=["Product Group", "Sub Group", "Item Name", "Unit"])

        add_pivot(wb, cache, "Sub Group", anchor,
                  row_fields=["Main Category", "Sub Category", "Product Group"],
                  page_fields=["Sub Group", "Item Name", "Unit"])

        for main_group, tab, val in groups:
            add_pivot(wb, cache, tab, anchor,
                      row_fields=["Sub Category", "Product Group", "Item Name"],
                      page_fields=["Main Category", "Sub Group", "Unit"],
                      locked=("Main Category", main_group))
            print(f"  {tab:<16} <- {main_group:<32} {val:>16,.2f}")

        build_index(wb, xw, xlsx, groups, meta, flags)
        print(f"  {'Index':<16} <- front sheet, {len(wb.sheets)} sheets listed")

        wb.sheets["Index"].activate()
        wb.save()
        wb.close()
        print(f"\nPivots + Index written into: {xlsx}")
        return 0

    except Exception as e:
        print(f"ERROR building PivotTables: {e}\n"
              f"       The data workbook itself is unaffected and still usable.\n"
              f"       Most common cause: the workbook is already open in Excel.",
              file=sys.stderr)
        return 1
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


if __name__ == "__main__":
    sys.exit(main())
