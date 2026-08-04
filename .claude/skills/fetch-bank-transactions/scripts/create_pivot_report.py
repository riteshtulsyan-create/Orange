#!/usr/bin/env python3
"""
create_pivot_report.py
Reads the flat Bank Transactions xlsx (from fetch-bank-transactions.mjs) and
produces a new Excel file with native Excel PivotTables on each summary sheet.

Usage:
  python create_pivot_report.py [source.xlsx] [output.xlsx]

Env vars (override defaults):
  SOURCE_FILE   path to flat input xlsx
  OUTPUT_FILE   path for pivot output xlsx
"""

import os, sys
import openpyxl
import xlwings as xw

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
OUTPUTS_DIR = os.path.join(SCRIPT_DIR, '..', '..', '..', '..', 'outputs')

SOURCE_FILE = os.environ.get('SOURCE_FILE',
    os.path.join(OUTPUTS_DIR, 'Bank_Transactions_Apr_Aug2026.xlsx'))
OUTPUT_FILE = os.environ.get('OUTPUT_FILE',
    os.path.join(OUTPUTS_DIR, 'Bank_Transactions_Apr_Aug2026_Pivot.xlsx'))

MONTH_ORDER = ['Apr-26','May-26','Jun-26','Jul-26','Aug-26','Sep-26',
               'Oct-26','Nov-26','Dec-26','Jan-27','Feb-27','Mar-27']
SKIP_SHEETS = {'All Transactions','ALL','BANK','CASH','RETUNN',
               'Related Partry','Bracnch Mis','Summary'}

# VBA / COM integer constants
XL_ROW   = 1;  XL_COL  = 2;  XL_PAGE  = 3
XL_DB    = 1;  XL_SUM  = -4157;  XL_DESC = 2;  XL_ASC = 1
XL_TAB   = 1   # xlTabularRow (non-compact PivotTable layout)


# ── helpers ──────────────────────────────────────────────────────────────────

def col_letter(n):
    """1-based column index → Excel letter (A, B, …, Z, AA, …)"""
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def sort_months(field):
    """Set chronological order on the Month PivotField."""
    for i, m in enumerate(MONTH_ORDER, 1):
        try:
            field.PivotItems(m).Position = i
        except Exception:
            pass


def add_main_pivot(pc, ws, pname, vc_filter=None):
    """
    Create Final Voucher × BS/PL × TYPE in rows, Month in columns.
    Optional vc_filter restricts Voucher Classification page field.
    """
    pt = pc.CreatePivotTable(ws['A1'].api, pname)
    pt.RowAxisLayout(XL_TAB)       # tabular (not compact) layout
    pt.ColumnGrand = True
    pt.RowGrand    = True

    # Row fields: Final Voucher → BS/PL → TYPE
    for pos, fn in enumerate(['Final Voucher', 'BS/PL', 'TYPE'], 1):
        f = pt.PivotFields(fn)
        f.Orientation = XL_ROW
        f.Position    = pos
        f.Subtotals   = (False,) * 12
    # Show subtotal on BS/PL
    pt.PivotFields('BS/PL').Subtotals = (True,) + (False,) * 11

    # Column field: Month (chronological order)
    fm = pt.PivotFields('Month')
    fm.Orientation = XL_COL
    fm.Position    = 1
    sort_months(fm)

    # Value: Sum of Amount
    pt.AddDataField(pt.PivotFields('Amount (₹)'), 'Sum of Amount', XL_SUM)

    # Page filters
    pt.PivotFields('Company').Orientation = XL_PAGE
    if vc_filter:
        fvc = pt.PivotFields('Voucher Classification')
        fvc.Orientation = XL_PAGE
        try:
            fvc.CurrentPage = vc_filter
        except Exception:
            pass   # filter value may not exist in this period

    # Sort TYPE descending by value
    try:
        pt.PivotFields('TYPE').AutoSort(XL_DESC, 'Sum of Amount')
    except Exception:
        pass

    # Number format on values
    try:
        pt.DataBodyRange.NumberFormat = '#,##0.00'
    except Exception:
        pass

    ws.autofit()


def add_type_pivot(pc, ws, pname):
    """
    Same layout as add_main_pivot (ALL) but with extra page filters:
    Category, Voucher Classification, and Party — none pre-filtered.
    """
    pt = pc.CreatePivotTable(ws['A1'].api, pname)
    pt.RowAxisLayout(XL_TAB)
    pt.ColumnGrand = True
    pt.RowGrand    = True

    for pos, fn in enumerate(['Final Voucher', 'BS/PL', 'TYPE'], 1):
        f = pt.PivotFields(fn)
        f.Orientation = XL_ROW
        f.Position    = pos
        f.Subtotals   = (False,) * 12
    pt.PivotFields('BS/PL').Subtotals = (True,) + (False,) * 11

    fm = pt.PivotFields('Month')
    fm.Orientation = XL_COL
    fm.Position    = 1
    sort_months(fm)

    pt.AddDataField(pt.PivotFields('Amount (₹)'), 'Sum of Amount', XL_SUM)

    for fn in ['Company', 'Category', 'Voucher Classification', 'Party']:
        pt.PivotFields(fn).Orientation = XL_PAGE

    try:
        pt.PivotFields('TYPE').AutoSort(XL_DESC, 'Sum of Amount')
    except Exception:
        pass

    try:
        pt.DataBodyRange.NumberFormat = '#,##0.00'
    except Exception:
        pass

    ws.autofit()


def add_current_month_pivot(pc, ws, pname, latest_month):
    """
    Drill-down pivot for the current/latest month:
    Rows: Final Voucher → BS/PL → TYPE → Party (4 levels)
    Columns: Date (individual dates)
    Page: Category, Voucher Classification, Company, Month (set to latest_month)
    """
    pt = pc.CreatePivotTable(ws['A1'].api, pname)
    pt.RowAxisLayout(XL_TAB)
    pt.ColumnGrand = True
    pt.RowGrand    = True

    for pos, fn in enumerate(['Final Voucher', 'BS/PL', 'TYPE', 'Party'], 1):
        f = pt.PivotFields(fn)
        f.Orientation = XL_ROW
        f.Position    = pos
        f.Subtotals   = (False,) * 12
    pt.PivotFields('TYPE').Subtotals = (True,) + (False,) * 11

    fd = pt.PivotFields('Date')
    fd.Orientation = XL_COL
    fd.Position    = 1

    pt.AddDataField(pt.PivotFields('Amount (₹)'), 'Sum of Amount', XL_SUM)

    for fn in ['Category', 'Voucher Classification', 'Company']:
        pt.PivotFields(fn).Orientation = XL_PAGE

    fm = pt.PivotFields('Month')
    fm.Orientation = XL_PAGE
    if latest_month:
        try:
            fm.CurrentPage = latest_month
        except Exception:
            pass

    try:
        pt.DataBodyRange.NumberFormat = '#,##0.00'
    except Exception:
        pass

    ws.autofit()


def add_party_pivot(pc, ws, pname, cat_filter):
    """
    Create Co Code -> Final Voucher -> Party in rows, Month in columns
    (chronological order). Category page field is fixed to cat_filter
    ('Related' or 'Branch'); Voucher Classification is a page filter
    left on '(All)' so Bank/Cash/CHQ.Ret can be toggled interactively.
    """
    pt = pc.CreatePivotTable(ws['A1'].api, pname)
    pt.RowAxisLayout(XL_TAB)
    pt.ColumnGrand = True
    pt.RowGrand    = True

    for pos, fn in enumerate(['Co Code', 'Final Voucher', 'Party'], 1):
        f = pt.PivotFields(fn)
        f.Orientation = XL_ROW
        f.Position    = pos
        f.Subtotals   = (False,) * 12

    fm = pt.PivotFields('Month')
    fm.Orientation = XL_COL
    fm.Position    = 1
    sort_months(fm)

    pt.AddDataField(pt.PivotFields('Amount (₹)'), 'Sum of Amount', XL_SUM)

    fc = pt.PivotFields('Category')
    fc.Orientation = XL_PAGE
    try:
        fc.CurrentPage = cat_filter
    except Exception:
        pass

    pt.PivotFields('Voucher Classification').Orientation = XL_PAGE

    try:
        pt.PivotFields('Party').AutoSort(XL_DESC, 'Sum of Amount')
    except Exception:
        pass

    try:
        pt.DataBodyRange.NumberFormat = '#,##0.00'
    except Exception:
        pass

    ws.autofit()


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    src = sys.argv[1] if len(sys.argv) > 1 else SOURCE_FILE
    out = sys.argv[2] if len(sys.argv) > 2 else OUTPUT_FILE

    print(f'=== Pivot Report Builder ===')
    print(f'Source : {src}')
    print(f'Output : {out}\n')

    # ── 1. Read source data ────────────────────────────────────────────────
    print('Reading source...')
    wb_r = openpyxl.load_workbook(src, read_only=True, data_only=True)

    all_trans = [list(r) for r in wb_r['All Transactions'].values]
    summary   = ([list(r) for r in wb_r['Summary'].values]
                 if 'Summary' in wb_r.sheetnames else [])
    co_data   = {n: [list(r) for r in wb_r[n].values]
                 for n in wb_r.sheetnames if n not in SKIP_SHEETS}
    wb_r.close()

    if not all_trans:
        sys.exit('ERROR: "All Transactions" sheet not found in source.')
    headers = all_trans[0]
    data    = all_trans[1:]
    print(f'  {len(data)} rows  |  {len(headers)} columns')

    # Determine latest month present in the data (for Current Month pivot page filter)
    try:
        month_col = list(headers).index('Month')
    except ValueError:
        month_col = 18  # fallback: position in HEADERS
    months_present = sorted(
        {str(row[month_col]) for row in data if row[month_col] and str(row[month_col]) in MONTH_ORDER},
        key=lambda m: MONTH_ORDER.index(m),
    )
    latest_month = months_present[-1] if months_present else None
    print(f'  Latest month detected: {latest_month}')

    # ── 2. Write base file via openpyxl (fast, no COM overhead) ───────────
    print('Writing base file...')
    wb_w = openpyxl.Workbook()

    ws_at = wb_w.active
    ws_at.title = 'All Transactions'
    ws_at.append(headers)
    for row in data:
        ws_at.append(row)
    ws_at.freeze_panes = 'A2'

    ws_s = wb_w.create_sheet('Summary')
    for row in summary:
        ws_s.append(row)

    for co_name, co_rows in co_data.items():
        ws_co = wb_w.create_sheet(co_name)
        for row in co_rows:
            ws_co.append(row)
        ws_co.freeze_panes = 'A2'

    wb_w.save(out)
    wb_w.close()
    print(f'  Base saved: {out}')

    # ── 3. Open in Excel COM and add PivotTables ───────────────────────────
    print('\nOpening in Excel (hidden)...')
    app = xw.App(visible=False, add_book=False)
    app.screen_updating = False
    app.display_alerts  = False

    try:
        wb = app.books.open(out)
        ws_at  = wb.sheets['All Transactions']
        ws_sum = wb.sheets['Summary']

        # Format All Transactions as an Excel Table
        last_col = col_letter(len(headers))
        last_row = len(data) + 1
        tbl = ws_at.api.ListObjects.Add(
            1,                                            # xlSrcRange
            ws_at.range(f'A1:{last_col}{last_row}').api,
            None,
            1,                                            # xlYes (has headers)
        )
        tbl.Name        = 'TransData'
        tbl.TableStyle  = 'TableStyleMedium9'
        ws_at.autofit()

        # Build a fresh PivotCache for each PivotTable (avoids shared-cache
        # field-state conflicts where Excel marks Amount as already claimed).
        src_ref = f"'All Transactions'!$A$1:${last_col}${last_row}"
        def make_pc():
            return wb.api.PivotCaches().Create(XL_DB, src_ref)

        # ── Add pivot sheets in order after Summary, before All Transactions ──
        # Desired final order:
        #   Summary | ALL | BANK | CASH | RETUNN | Related Partry | Bracnch Mis
        #   | All Transactions | [company sheets]
        #
        # Strategy: insert each sheet after the previous one, starting after Summary.
        # Then move All Transactions to after Bracnch Mis at the end.

        prev = ws_sum

        def next_sheet(name):
            nonlocal prev
            ws = wb.sheets.add(name, after=prev)
            prev = ws
            return ws

        print('\nBuilding PivotTables:')

        ws = next_sheet('ALL')
        add_main_pivot(make_pc(), ws, 'PivotALL', None)
        print('  OK ALL')

        ws = next_sheet('Type')
        add_type_pivot(make_pc(), ws, 'PivotType')
        print('  OK Type')

        ws = next_sheet('Current Month')
        add_current_month_pivot(make_pc(), ws, 'PivotCurrMon', latest_month)
        print(f'  OK Current Month ({latest_month})')

        for name, pname, vc in [
            ('BANK',   'PivotBANK',   'Bank'),
            ('CASH',   'PivotCASH',   'Cash'),
            ('RETUNN', 'PivotRETUNN', 'CHQ.Ret'),
        ]:
            ws = next_sheet(name)
            add_main_pivot(make_pc(), ws, pname, vc)
            print(f'  OK {name}')

        for name, pname, cat in [
            ('Related Partry', 'PivotRelated', 'Related'),
            ('Bracnch Mis',    'PivotBranch',  'Branch'),
        ]:
            ws = next_sheet(name)
            add_party_pivot(make_pc(), ws, pname, cat)
            print(f'  OK {name}')

        ws_sum.autofit()

        # Move 'All Transactions' to after 'Bracnch Mis' (= prev)
        # Current order: [All Transactions, Summary, ALL, BANK, CASH, RETUNN,
        #                  Related Partry, Bracnch Mis, co1, co2, ...]
        # After move:     [Summary, ALL, BANK, CASH, RETUNN, Related Partry,
        #                  Bracnch Mis, All Transactions, co1, co2, ...]
        ws_at.api.Move(After=prev.api)

        print(f'\nSaving: {out}')
        wb.api.Save()   # Save() not SaveAs() — file is already open at this path
        print('Done.')

    finally:
        try:
            wb.close()
        except Exception:
            pass
        app.quit()


if __name__ == '__main__':
    main()
