#!/usr/bin/env python3
"""
create_all_pivots.py
Adds the Index + PivotTable sheets to the combined all-companies stock workbook
produced by fetch_stock_all_companies.py.

Six sheets total, as agreed with the user 2026-07-31:
    Index            navigation links only
    Summary          Company > Main Category
    Group            Company > Main Category > Sub Category
    Sub Group        Company > Main Category > Sub Category > Product Group
    Item List (Flat) source data
    Cross Check      reconciliation

Every pivot carries Company, Unit and the mapped Main Group as page fields, so
any company can be isolated and quantity can be restricted to a single unit of
measure (Sum of Quantity otherwise adds KGS + PCS + MTR together).

Usage:
    python create_all_pivots.py <workbook.xlsx>
"""

import sys
from pathlib import Path

import openpyxl

XL_ROW, XL_COL, XL_PAGE, XL_DATA = 1, 2, 3, 4
XL_DB = 1
XL_SUM = -4157
XL_TAB = 1
XL_COMPACT = 0

DATA_SHEET = "Item List (Flat)"
MOM_SHEET = "MOM Data"
VALUE_FMT = "#,##0.00"
QTY_FMT = "#,##0.000"

SHEET_NOTES = {
    "Index": "This sheet - what is in the workbook and where to find it.",
    "Summary": "Main Category down the side, Main Company across the top. "
               "Value only. Start here.",
    "Group": "Main Company > Main Category > Sub Category.",
    "Sub Group": "Main Company > Main Category > Sub Category > Product Group. "
                 "Filter Company or Location for HO vs Noida.",
    "Month on Month": "Closing stock value at each month end, by entity and by "
                      "Main Category. Each month is its own live fetch from Tally.",
    "MOM Pivot": "Month-on-month as a PIVOT - drag Main Company / Company / "
                 "Location / Sub Category to reshape it however you need.",
    "MOM Data": "Source data for MOM Pivot: one row per period x company x "
                "category x sub-category, including the as-on-date snapshot.",
    "Item List (Flat)": "Source data - every item of every company, one row each.",
    "Cross Check": "Per-company reconciliation, exceptions and name variations.",
}


def build_index(wb, stamp=None):
    """Plain navigation list - no company-specific content (see Cross Check)."""
    for existing in list(wb.sheets):
        if existing.name == "Index":
            existing.delete()
    ws = wb.sheets.add("Index", before=wb.sheets[0])

    ws.range("A1").value = "INDEX"
    ws.range("A1").api.Font.Size = 16
    ws.range("A1").api.Font.Bold = True
    if stamp:
        ws.range("A2").value = stamp
        ws.range("A2").api.Font.Italic = True
        ws.range("A2").api.Font.Size = 9

    r = 4
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


def add_pivot(wb, cache_factory, sheet_name, before_sheet,
              row_fields, page_fields, col_fields=(), with_qty=True, compact=False,
              stamp=None, value_field="Value", qty_field="Quantity"):
    for existing in list(wb.sheets):
        if existing.name == sheet_name:
            existing.delete()
    ws = wb.sheets.add(sheet_name, before=before_sheet)

    # Anchor at A1 -- anchoring lower collides with the page-field block Excel
    # inserts above the table and leaves a #VALUE! error instead of a pivot.
    pt = cache_factory().CreatePivotTable(ws.range("A1").api,
                                          f"PT_{abs(hash(sheet_name)) % 100000}")
    pt.RowAxisLayout(XL_COMPACT if compact else XL_TAB)
    pt.RowGrand = True
    pt.ColumnGrand = True

    for pos, fname in enumerate(row_fields, 1):
        f = pt.PivotFields(fname)
        f.Orientation = XL_ROW
        f.Position = pos
    for pos, fname in enumerate(col_fields, 1):
        f = pt.PivotFields(fname)
        f.Orientation = XL_COL
        f.Position = pos
    for fname in page_fields:
        pt.PivotFields(fname).Orientation = XL_PAGE

    if with_qty:
        q = pt.AddDataField(pt.PivotFields(qty_field), f"Sum of {qty_field}", XL_SUM)
        q.NumberFormat = QTY_FMT
    v = pt.AddDataField(pt.PivotFields(value_field), f"Sum of {value_field}", XL_SUM)
    v.NumberFormat = VALUE_FMT

    ws.autofit(axis="columns")
    if stamp:
        # Column J: clear of every pivot built here (widest is Summary at
        # column E), so the stamp can never collide with the table.
        cell = ws.range("J1")
        cell.value = stamp
        cell.api.Font.Italic = True
        cell.api.Font.Size = 9
    return ws


def main():
    if len(sys.argv) < 2:
        print("usage: python create_all_pivots.py <workbook.xlsx>", file=sys.stderr)
        return 2
    xlsx = Path(sys.argv[1]).resolve()
    if not xlsx.is_file():
        print(f"ERROR: not found: {xlsx}", file=sys.stderr)
        return 1

    # Read anything openpyxl needs BEFORE Excel takes an exclusive lock.
    wbk = openpyxl.load_workbook(xlsx, read_only=True)
    ds = wbk[DATA_SHEET]
    n_rows, n_cols = ds.max_row, ds.max_column
    # Row 1 is the refresh stamp, row 2 the header -- carry the stamp onto
    # every pivot sheet so each one states when it was last refreshed.
    stamp = ds["A1"].value if isinstance(ds["A1"].value, str) else None
    wbk.close()
    print(f"Source: {n_rows - 2:,} data rows x {n_cols} columns (header on row 2)")

    try:
        import xlwings as xw
    except ImportError:
        print("ERROR: xlwings not installed - cannot build PivotTables.", file=sys.stderr)
        return 1

    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(str(xlsx))

        src_ref = f"'{DATA_SHEET}'!R2C1:R{n_rows}C{n_cols}"

        def cache():
            return wb.api.PivotCaches().Create(XL_DB, src_ref)

        anchor = wb.sheets[DATA_SHEET]

        # Layouts below match what the user built by hand on the 0605 workbook
        # (2026-07-31). Summary is a cross-tab: Main Company across the top,
        # value only -- quantity is deliberately not summed here because it
        # mixes KGS/PCS/MTR. Group and Sub Group keep Company as a page field,
        # not a row field, so the rows read Main Company > groups.
        add_pivot(wb, cache, "Summary", anchor,
                  row_fields=["Main Category"],
                  col_fields=["Main Company"],
                  page_fields=["Company", "Location", "Unit"],
                  with_qty=False, compact=True, stamp=stamp)
        print("  Summary          rows: Main Category   cols: Main Company   (value only)")

        add_pivot(wb, cache, "Group", anchor,
                  row_fields=["Main Company", "Main Category", "Sub Category"],
                  page_fields=["Company", "Location", "Main Group (Mapped)",
                               "Product Group", "Sub Group", "Item Name", "Unit"],
                  stamp=stamp)
        print("  Group            Main Company > Main Category > Sub Category")

        add_pivot(wb, cache, "Sub Group", anchor,
                  row_fields=["Main Company", "Main Category", "Sub Category",
                              "Product Group"],
                  page_fields=["Company", "Location", "Main Group (Mapped)",
                               "Sub Group", "Sub Group 2", "Item Name", "Unit"],
                  stamp=stamp)
        print("  Sub Group        Main Company > Main Category > Sub Category > Product Group")

        # Month-on-month pivot, off its own long-form data sheet. Rows =
        # Main Category, columns = Period -- the layout the user was choosing
        # between -- but as a pivot they can reshape at will.
        if MOM_SHEET in [s.name for s in wb.sheets]:
            mom_ws = wb.sheets[MOM_SHEET]
            m_rows = mom_ws.range("A2").end("down").row
            m_cols = mom_ws.range("A2").end("right").column
            mom_ref = f"'{MOM_SHEET}'!R2C1:R{m_rows}C{m_cols}"
            print(f"MOM source: {mom_ref}")

            def mom_cache():
                return wb.api.PivotCaches().Create(XL_DB, mom_ref)

            add_pivot(wb, mom_cache, "MOM Pivot", wb.sheets[MOM_SHEET],
                      row_fields=["Main Category"],
                      col_fields=["Period"],
                      page_fields=["Period Type", "Main Company", "Company",
                                   "Location", "Main Group (Mapped)", "Sub Category"],
                      with_qty=False, compact=True, stamp=stamp)
            print("  MOM Pivot        rows: Main Category   cols: Period")

        build_index(wb, stamp)
        print(f"  Index            {len(wb.sheets)} sheets listed")

        wb.sheets["Index"].activate()
        wb.save()
        wb.close()
        print(f"\nPivots + Index written into: {xlsx}")
        return 0

    except Exception as e:
        print(f"ERROR building PivotTables: {e}\n"
              f"       The data workbook itself is unaffected and still usable.",
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
