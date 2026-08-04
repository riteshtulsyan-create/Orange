"""Build the item-wise GP Segment MoM workbook.

This is the accurate build. Segment comes from each STOCK ITEM's own group, so
it works for all five companies - including Enterprises and Colorix, whose
sales ledgers carry no segment tag at all.

Per item, per period:
    cost rate = (opening value + purchase value) / (opening qty + purchase qty)
    COGS      = sales qty * cost rate
    GP        = sales value - COGS

Falls back to a year-to-date weighted average, then the closing rate, when a
period has no opening stock and no purchases to strike a rate from.

Inputs:
    Files\\gp_voucher_cache.json   (fetch_gp_vouchers.py)  - required
    Files\\gp_segment_cache.json   (fetch_gp_data.py)      - stock group tree
                                                             and ledger sales
                                                             for reconciliation

Run:
    python .claude\\skills\\gp-segment-mom\\scripts\\build_gp_itemwise.py
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parents[3]
sys.path.insert(0, str(HERE))
from segment_map import (  # noqa: E402
    COMPANY_ENTITY, SEGMENTS, segment_for_item, segment_for_ledger,
)

TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, size=10, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="2E5C8A")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
BOLD = Font(bold=True, size=10)
NORMAL = Font(size=10)
NOTE_FONT = Font(size=9, italic=True, color="606060")
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '#,##0;[Red](#,##0)'
QTY = '#,##0.0;[Red](#,##0.0)'
PCT = '0.0%'
METRICS = ["Sales", "COGS", "Gross Profit", "GP %", "Qty"]


# ---------------------------------------------------------------------------
# per-item gross profit
# ---------------------------------------------------------------------------

def cost_rate_for(item_key, period_items, ytd_items, ytd_vouchers, period_vch):
    """Weighted-average cost per unit, with documented fallbacks.

    Returns (rate, basis) where basis records which rule supplied it, so the
    workbook can report how much of COGS rests on a fallback.
    """
    pi = period_items.get(item_key) or {}
    pv = period_vch.get(item_key) or {}
    open_qty = pi.get("open_qty", 0.0)
    open_val = pi.get("open_val", 0.0)
    purch_qty = pv.get("purch_qty", 0.0)
    purch_val = pv.get("purch_val", 0.0)

    denom = open_qty + purch_qty
    if denom > 1e-9 and (open_val + purch_val) > 0:
        return (open_val + purch_val) / denom, "period"

    yi = ytd_items.get(item_key) or {}
    yv = ytd_vouchers.get(item_key) or {}
    denom = yi.get("open_qty", 0.0) + yv.get("purch_qty", 0.0)
    numer = yi.get("open_val", 0.0) + yv.get("purch_val", 0.0)
    if denom > 1e-9 and numer > 0:
        return numer / denom, "ytd"

    close_qty, close_val = pi.get("close_qty", 0.0), pi.get("close_val", 0.0)
    if close_qty > 1e-9 and close_val > 0:
        return close_val / close_qty, "closing"

    return 0.0, "none"


def discount_by_segment(sdata, company, period, sales_mix):
    """Segment-tagged discount and rate-difference ledgers for one company.

    These sit in ledgers, never on an inventory line, so the item-wise pass
    cannot see them - but they are real reductions of revenue and the existing
    GP REPORT applies them. Untagged discounts are spread on the company's own
    item sales mix.
    """
    if not sdata:
        return {s: 0.0 for s in SEGMENTS}
    rows = sdata["companies"].get(company, {}).get("ledgers", {}).get(period, [])
    out = {s: 0.0 for s in SEGMENTS}
    untagged = 0.0
    for row in rows:
        if row.get("bucket") != "discount":
            continue
        segment = segment_for_ledger(row["ledger"])
        if segment:
            out[segment] += row["amount"]
        else:
            untagged += row["amount"]
    if untagged:
        base = {s: max(sales_mix.get(s, 0.0), 0.0) for s in SEGMENTS}
        denom = sum(base.values())
        for s in SEGMENTS:
            out[s] += (untagged * base[s] / denom) if denom > 0 else (
                untagged if s == "OTHER" else 0.0)
    return out


def company_segments(cdata, company, period, parents, sdata=None):
    """{segment: {Sales, COGS, Gross Profit, Qty}} for one company/period."""
    pdata = cdata.get("periods", {}).get(period, {})
    period_items = pdata.get("items", {})
    period_vch = pdata.get("vouchers", {})
    ytd_items = cdata.get("ytd_items", {})
    ytd_vch = cdata.get("ytd_vouchers", {})

    out = {s: defaultdict(float) for s in SEGMENTS}
    by_group = defaultdict(lambda: defaultdict(float))
    no_rate_sales = 0.0
    basis_count = defaultdict(float)

    for item_key, v in period_vch.items():
        sales_val = v.get("sales_val", 0.0)
        sales_qty = v.get("sales_qty", 0.0)
        if abs(sales_val) < 1e-9 and abs(sales_qty) < 1e-9:
            continue

        info = period_items.get(item_key) or ytd_items.get(item_key) or {}
        group = info.get("group") or "(No stock group assigned)"
        segment, _explicit, _used = segment_for_item(group, parents, company)

        rate, basis = cost_rate_for(item_key, period_items, ytd_items,
                                    ytd_vch, period_vch)
        cogs = sales_qty * rate
        if basis == "none" and abs(sales_val) > 1e-9:
            no_rate_sales += sales_val
        basis_count[basis] += abs(sales_val)

        out[segment]["Sales"] += sales_val
        out[segment]["COGS"] += cogs
        out[segment]["Gross Profit"] += sales_val - cogs
        out[segment]["Qty"] += sales_qty

        g = by_group[group]
        g["Sales"] += sales_val
        g["COGS"] += cogs
        g["Gross Profit"] += sales_val - cogs
        g["Qty"] += sales_qty
        g["_segment"] = segment

    disc = discount_by_segment(sdata, company, period,
                               {s: out[s]["Sales"] for s in SEGMENTS})
    for s in SEGMENTS:
        if disc[s]:
            out[s]["Sales"] += disc[s]
            out[s]["Gross Profit"] += disc[s]
            out[s]["Discount"] = out[s].get("Discount", 0.0) + disc[s]

    return out, by_group, no_rate_sales, basis_count


def aggregate(vdata, parents_by_company, period, company_filter=None, sdata=None):
    out = {s: defaultdict(float) for s in SEGMENTS}
    no_rate = 0.0
    basis = defaultdict(float)
    for company, cdata in vdata["companies"].items():
        if company_filter and not company_filter(company, cdata):
            continue
        segs, _g, nr, bc = company_segments(
            cdata, company, period, parents_by_company.get(company, {}), sdata)
        for s in SEGMENTS:
            for k, val in segs[s].items():
                out[s][k] += val
        no_rate += nr
        for k, val in bc.items():
            basis[k] += val
    for s in SEGMENTS:
        sales = out[s]["Sales"]
        out[s]["GP %"] = (out[s]["Gross Profit"] / sales) if sales else 0.0
    return out, no_rate, basis


def quarter_months(vdata, quarter_key):
    """Months that make up a quarter column."""
    q = next((p for p in vdata["periods"] if p["key"] == quarter_key), None)
    if not q:
        return []
    return [p["key"] for p in vdata["periods"]
            if p["kind"] == "month" and q["from"] <= p["from"] <= q["to"]]


def aggregate_column(vdata, parents, pk, company_filter=None, sdata=None):
    """Figures for one report column.

    Quarter columns are the SUM of their months rather than an independent
    quarterly fetch. Each period is costed on its own weighted-average rate,
    so a quarterly average does not equal the sum of three monthly averages -
    about a 2% gap on gross profit. Summing keeps the columns internally
    consistent, which is what a month-on-month pack needs; the independent
    quarterly fetch is retained purely as a cross-check.
    """
    months = quarter_months(vdata, pk)
    if not months:
        return aggregate(vdata, parents, pk, company_filter, sdata)[0]
    out = {s: defaultdict(float) for s in SEGMENTS}
    for mk in months:
        agg, _, _ = aggregate(vdata, parents, mk, company_filter, sdata)
        for s in SEGMENTS:
            for k, val in agg[s].items():
                if k != "GP %":
                    out[s][k] += val
    for s in SEGMENTS:
        sales = out[s]["Sales"]
        out[s]["GP %"] = (out[s]["Gross Profit"] / sales) if sales else 0.0
    return out


def rows_for(vdata, parents, period_keys, company_filter=None, sdata=None):
    rows = {s: {} for s in SEGMENTS}
    total = {}
    for pk in period_keys:
        agg = aggregate_column(vdata, parents, pk, company_filter, sdata)
        for s in SEGMENTS:
            rows[s][pk] = dict(agg[s])
        t_sales = sum(agg[s]["Sales"] for s in SEGMENTS)
        t_cogs = sum(agg[s]["COGS"] for s in SEGMENTS)
        total[pk] = {"Sales": t_sales, "COGS": t_cogs,
                     "Gross Profit": t_sales - t_cogs,
                     "GP %": ((t_sales - t_cogs) / t_sales) if t_sales else 0.0,
                     "Qty": sum(agg[s]["Qty"] for s in SEGMENTS)}
    rows["TOTAL"] = total
    return rows



# ---------------------------------------------------------------------------
# board reconciliation against PL_BS_GL_Detail
# ---------------------------------------------------------------------------

def read_pl_workbook(pl_path):
    """Pull the Q1 group trading account out of a PL_BS_GL_Detail workbook.

    Reads the 'MOM PL and BS' sheet, whose sixth column is the Q1 total.
    """
    import openpyxl as _ox
    wb = _ox.load_workbook(pl_path, data_only=True, read_only=True)
    if "MOM PL and BS" not in wb.sheetnames:
        return None
    ws = wb["MOM PL and BS"]
    want = {"Sales Accounts", "Direct Incomes", "Total Turnover",
            "Purchase Accounts", "Direct Expenses",
            "Stock Valuation (Inventory)", "Total Cost of Sales",
            "Branch Sales", "Branch Purchase", "GROSS PROFIT"}
    out = {}
    for row in ws.iter_rows(values_only=True):
        label = str(row[0] or row[1] or "").strip()
        if label in want and len(row) > 5 and isinstance(row[5], (int, float)):
            out[label] = float(row[5])
    return out or None


def write_reconciliation(wb, pl_figs, pl_name, item_sales, item_cogs,
                         item_gp, branch_sales, period_label):
    """Bridge the segment report to the statutory P&L, line by line.

    Both reports are right; they answer different questions. The segment
    report matches cost to what was actually sold. The P&L charges purchases
    as incurred and adjusts by stock movement. Everything between them is
    either a classification difference or a timing difference.
    """
    ws = wb.create_sheet("Board Reconciliation")
    r = _title(ws, "Board Reconciliation - Gross Profit, " + period_label, 4,
               "GP Segment (item-wise) vs " + pl_name +
               ". Both are correct; they answer different questions.")

    pl_sales = pl_figs.get("Sales Accounts", 0.0)
    pl_dirinc = pl_figs.get("Direct Incomes", 0.0)
    pl_direxp = abs(pl_figs.get("Direct Expenses", 0.0))
    pl_stockval = pl_figs.get("Stock Valuation (Inventory)", 0.0)
    pl_gp = pl_figs.get("GROSS PROFIT", 0.0)

    external = item_sales - branch_sales
    non_inventory = pl_sales - external
    # timing is the balancing figure, and is labelled as such
    timing = pl_gp - (item_gp - branch_sales + non_inventory + pl_dirinc
                      - pl_direxp + pl_stockval)

    state = {"r": r}

    def band(text):
        i = state["r"]
        c = ws.cell(row=i, column=1, value=text)
        c.font = BOLD
        for k in range(1, 4):
            ws.cell(row=i, column=k).fill = SUB_FILL
            ws.cell(row=i, column=k).border = BORDER
        state["r"] = i + 1

    def line(label, amount, note="", total=False, formula=None):
        i = state["r"]
        c = ws.cell(row=i, column=1, value=label)
        c.font = BOLD if total else NORMAL
        c.alignment = Alignment(wrap_text=True, vertical="top")
        v = ws.cell(row=i, column=2)
        v.value = formula if formula else amount
        v.number_format = MONEY
        v.font = BOLD if total else NORMAL
        v.border = BORDER
        if total:
            c.fill = TOTAL_FILL
            v.fill = TOTAL_FILL
        n = ws.cell(row=i, column=3, value=note)
        n.font = NOTE_FONT
        n.alignment = Alignment(wrap_text=True, vertical="top")
        state["r"] = i + 1
        return i

    for idx, h in enumerate(["Line", "Amount (Rs)", "What it is"], start=1):
        c = ws.cell(row=state["r"], column=idx, value=h)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
    state["r"] += 1

    band("A.  REVENUE BRIDGE")
    a1 = line("Segment report sales (incl. inter-branch, after discounts)",
              item_sales, "Total row of 'MoM GP - Consolidated'.")
    a2 = line("Less: inter-branch sales", -branch_sales,
              "The P&L shows branch sales and branch purchases separately and "
              "they net to nil. The segment report includes them.")
    a3 = line("External sales on inventory lines", None,
              "What the group actually sold to third parties, item by item.",
              total=True, formula="=B{0}+B{1}".format(a1, a2))
    a4 = line("Add: non-inventory revenue", non_inventory,
              "Software, services, high-seas margin, freight recovered. Real "
              "revenue that never appears on an inventory line, so the "
              "item-wise pass cannot see it.")
    a5 = line("Sub-total: P&L Sales Accounts", None, "", total=True,
              formula="=B{0}+B{1}".format(a3, a4))
    a6 = line("Add: Direct Incomes", pl_dirinc, "")
    line("Sub-total: P&L Total Turnover", None, "", total=True,
         formula="=B{0}+B{1}".format(a5, a6))
    state["r"] += 1

    band("B.  COST BRIDGE")
    b1 = line("Segment report COGS (matched cost of goods sold)", item_cogs,
              "Quantity sold x weighted-average cost. Only the cost of what "
              "left the shelf.")
    b2 = line("Add: direct costs charged to cost of sales in the P&L",
              pl_direxp,
              "Freight, customs duty, clearing and forwarding, packing, "
              "labour. Real landed costs, but not matched to any one item.")
    b3 = line("Add/(less): stock valuation adjustment", -pl_stockval,
              "Provisions and revaluations taken through cost of sales.")
    b4 = line("Add/(less): purchases-vs-consumption timing", -timing,
              "THE BIG ONE. The P&L charges purchases as incurred and adjusts "
              "by stock movement; the segment report charges only the cost of "
              "goods sold. Also carries the cost of inter-branch transfers.")
    line("Sub-total: P&L Total Cost of Sales", None, "", total=True,
         formula="=B{0}+B{1}+B{2}+B{3}".format(b1, b2, b3, b4))
    state["r"] += 1

    band("C.  GROSS PROFIT BRIDGE")
    c1 = line("Gross profit per GP Segment report", item_gp, "", total=True)
    c2 = line("Less: inter-branch sales", -branch_sales, "Per A above.")
    c3 = line("Add: non-inventory revenue", non_inventory, "Per A above.")
    c4 = line("Add: Direct Incomes", pl_dirinc, "Per A above.")
    c5 = line("Less: direct costs not matched to items", -pl_direxp,
              "Per B above.")
    c6 = line("Add/(less): stock valuation adjustment", pl_stockval,
              "Per B above.")
    c7 = line("Add/(less): purchases-vs-consumption timing", timing,
              "Per B above.")
    c8 = line("Gross profit per P&L (derived from the bridge)", None, "", total=True,
              formula="=B{0}+B{1}+B{2}+B{3}+B{4}+B{5}+B{6}".format(
                  c1, c2, c3, c4, c5, c6, c7))
    tie = line("P&L gross profit as reported", pl_gp,
               "Straight from the P&L workbook.", total=True)

    i = state["r"]
    ws.cell(row=i, column=1, value="Difference").font = BOLD
    d = ws.cell(row=i, column=2, value="=B{0}-B{1}".format(c8, tie))
    d.number_format = MONEY
    d.font = BOLD
    d.fill = OK_FILL
    ws.cell(row=i, column=3,
            value="Zero means the two reports fully reconcile.").font = NOTE_FONT
    state["r"] = i + 2

    band("WHICH NUMBER SHOULD THE BOARD USE?")
    for label, text in [
        ("Segment GP%", "Use for pricing, product mix and salesforce "
                        "performance. It is the true margin earned on what "
                        "was sold, matched item by item."),
        ("P&L gross profit", "Use for statutory reporting and the cash-cycle "
                             "discussion. It absorbs landed costs and stock "
                             "movement, so it swings when inventory is built "
                             "up or run down."),
        ("Why they differ", "Mostly timing, not error. In this period the "
                            "group bought more than it sold, so the P&L "
                            "carries cost the segment report has not yet "
                            "released to the income statement."),
    ]:
        i = state["r"]
        ws.cell(row=i, column=1, value=label).font = BOLD
        c = ws.cell(row=i, column=2, value=text)
        c.font = NORMAL
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3)
        ws.row_dimensions[i].height = 32
        state["r"] = i + 1

    ws.column_dimensions["A"].width = 56
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 78
    return ws


def write_index(wb, meta_rows, sheet_notes, headline):
    """Front page: what each sheet is, with links, plus the headline figures."""
    ws = wb.create_sheet("Index", 0)
    r = _title(ws, "Gross Profit by Segment - Index", 4)
    for label, text in meta_rows:
        ws.cell(row=r, column=1, value=label).font = BOLD
        c = ws.cell(row=r, column=2, value=text)
        c.font = NORMAL
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        r += 1
    r += 1

    ws.cell(row=r, column=1, value="HEADLINE").font = BOLD
    for k in range(1, 4):
        ws.cell(row=r, column=k).fill = SUB_FILL
        ws.cell(row=r, column=k).border = BORDER
    r += 1
    for label, formula, note in headline:
        ws.cell(row=r, column=1, value=label).font = NORMAL
        c = ws.cell(row=r, column=2, value=formula)
        c.number_format = PCT if "%" in label else MONEY
        c.font = BOLD
        c.border = BORDER
        n = ws.cell(row=r, column=3, value=note)
        n.font = NOTE_FONT
        n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    r += 1

    for idx, h in enumerate(["#", "Sheet", "What it shows"], start=1):
        c = ws.cell(row=r, column=idx, value=h)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
    r += 1
    for n, (sheet, note) in enumerate(sheet_notes, start=1):
        ws.cell(row=r, column=1, value=n).font = NORMAL
        c = ws.cell(row=r, column=2, value=sheet)
        c.hyperlink = "#'" + sheet + "'!A1"
        c.font = Font(size=10, color="0563C1", underline="single")
        c.border = BORDER
        d = ws.cell(row=r, column=3, value=note)
        d.font = NORMAL
        d.alignment = Alignment(wrap_text=True, vertical="top")
        d.border = BORDER
        r += 1

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 92
    return ws


# ---------------------------------------------------------------------------
# sheets
# ---------------------------------------------------------------------------

def _title(ws, text, ncols, subtitle=None):
    ws.cell(row=1, column=1, value=text).font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(ncols, 2))
    for c in range(1, max(ncols, 2) + 1):
        ws.cell(row=1, column=c).fill = TITLE_FILL
    ws.row_dimensions[1].height = 22
    if subtitle:
        ws.cell(row=2, column=1, value=subtitle).font = NOTE_FONT
        return 4
    return 3


def write_matrix(ws, title, subtitle, period_keys, row_groups, link_rows=None):
    """Write a segment x period matrix.

    link_rows, when given, maps (block_label, row_label) -> a list of
    '<sheet>'!<row> references. Numeric cells are then written as a SUM across
    those rows instead of a literal, so the figure is visibly derived from the
    sheet below it rather than independently recomputed.

    Returns {(block_label, row_label): row_number} so a later sheet can link
    back to this one.
    """
    ncols = 1 + len(period_keys) * len(METRICS)
    r = _title(ws, title, ncols, subtitle)
    col = 2
    for pk in period_keys:
        ws.cell(row=r, column=col, value=pk).font = HDR_FONT
        ws.merge_cells(start_row=r, start_column=col,
                       end_row=r, end_column=col + len(METRICS) - 1)
        for i in range(len(METRICS)):
            c = ws.cell(row=r, column=col + i)
            c.fill, c.border = HDR_FILL, BORDER
            c.alignment = Alignment(horizontal="center")
        col += len(METRICS)
    ws.cell(row=r, column=1).fill = HDR_FILL
    ws.cell(row=r, column=1).border = BORDER

    r += 1
    h = ws.cell(row=r, column=1, value="Segment")
    h.font, h.fill, h.border = HDR_FONT, HDR_FILL, BORDER
    col = 2
    for _ in period_keys:
        for m in METRICS:
            c = ws.cell(row=r, column=col, value=m)
            c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
            c.alignment = Alignment(horizontal="center")
            col += 1
    ws.freeze_panes = ws.cell(row=r + 1, column=2)
    r += 1

    row_index = {}
    for label, rows in row_groups:
        if label:
            ws.cell(row=r, column=1, value=label).font = BOLD
            for c in range(1, ncols + 1):
                ws.cell(row=r, column=c).fill = SUB_FILL
                ws.cell(row=r, column=c).border = BORDER
            r += 1
        for name, per_period in rows.items():
            is_total = name.upper().startswith("TOTAL")
            c1 = ws.cell(row=r, column=1, value=name)
            c1.font = BOLD if is_total else NORMAL
            c1.border = BORDER
            if is_total:
                c1.fill = TOTAL_FILL
            col = 2
            for pk in period_keys:
                fig = per_period.get(pk, {})
                for m in METRICS:
                    cell = ws.cell(row=r, column=col)
                    if m == "GP %":
                        gp = get_column_letter(col - 1)
                        sa = get_column_letter(col - 3)
                        cell.value = f"=IF({sa}{r}=0,0,{gp}{r}/{sa}{r})"
                        cell.number_format = PCT
                    else:
                        refs = (link_rows or {}).get((label, name))
                        if refs:
                            col_letter = get_column_letter(col)
                            cell.value = "=" + "+".join(
                                f"{sheet}!{col_letter}{rw}" for sheet, rw in refs)
                        else:
                            cell.value = fig.get(m, 0.0)
                        cell.number_format = QTY if m == "Qty" else MONEY
                    cell.font = BOLD if is_total else NORMAL
                    cell.border = BORDER
                    if is_total:
                        cell.fill = TOTAL_FILL
                    col += 1
            row_index[(label, name)] = r
            r += 1
        r += 1
    ws.column_dimensions["A"].width = 36
    for c in range(2, ncols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 14
    return row_index


def build(vpath: Path, spath: Path, out_path: Path, pl_path: Path = None):
    vdata = json.loads(vpath.read_text(encoding="utf-8"))
    sdata = json.loads(spath.read_text(encoding="utf-8")) if spath.exists() else None

    parents = {c: (sdata["companies"].get(c, {}).get("parents", {}) if sdata else {})
               for c in vdata["companies"]}

    periods = vdata["periods"]
    months = [p["key"] for p in periods if p["kind"] == "month"]
    quarters = [p["key"] for p in periods if p["kind"] == "quarter"]
    period_keys = []
    for m in months:
        if m.startswith("JUL") and quarters:
            period_keys.extend(q for q in quarters if q not in period_keys)
        period_keys.append(m)
    period_keys.extend(q for q in quarters if q not in period_keys)
    detail_period = quarters[0] if quarters else period_keys[0]

    wb = Workbook()

    # --- Basis --------------------------------------------------------------
    ws = wb.active
    ws.title = "Basis & Method"
    r = _title(ws, "Gross Profit by Segment - Item-wise, Month on Month", 4)
    for label, text in [
        ("Source", "Tally Prime, live. Every sales and purchase voucher INVENTORY "
                   "LINE, read item by item."),
        ("Generated", vdata.get("generated", "")),
        ("Period", f"{vdata.get('fy_start','')} to {vdata.get('through','')}"),
        ("", ""),
        ("Segment", "Taken from each STOCK ITEM's own stock group - not from the "
                    "ledger name. This is why it works for all five companies, "
                    "including Enterprises and Colorix whose sales ledgers carry "
                    "no segment tag."),
        ("Method", "cost rate = (opening value + purchase value) / (opening qty + "
                   "purchase qty);  COGS = sales qty x cost rate;  GP = sales - COGS"),
        ("Fallbacks", "Where a period has no opening stock and no purchases to "
                      "strike a rate from, a year-to-date weighted average is used, "
                      "then the closing rate. The Cross Check sheet reports how much "
                      "of sales rests on each basis."),
        ("", ""),
        ("Excluded deliberately", "Proforma invoices (voucher type parent 'Sales "
                                  "Order'), delivery challans ('Delivery Note'), "
                                  "stock journals, material in/out, rejections and "
                                  "orders. These move goods but are not sales - "
                                  "counting them would have overstated Orange O Tec "
                                  "HO April revenue by about 8.8 crore."),
        ("Included", "Sales and Purchase vouchers, net of Credit Notes (sales "
                     "returns) and Debit Notes (purchase returns). Cancelled and "
                     "optional vouchers are dropped."),
        ("", ""),
        ("Note on voucher types", "Tally lets the default voucher types be renamed. "
                                  "In Orange O Tec HO the base Sales type is renamed "
                                  "'Sales Accounts-HSS' and Purchase is 'GST "
                                  "PURCHASE', and each is its own parent. Types are "
                                  "therefore resolved by walking the parent chain to "
                                  "its root and classifying the root by pattern."),
        ("Quantity", "Summed in each item's own billing unit. Units differ between "
                     "items (NOS, PCS, KGS, MTR), so a segment quantity total mixes "
                     "units and is indicative only - use it for trend, not level."),
    ]:
        ws.cell(row=r, column=1, value=label).font = BOLD
        c = ws.cell(row=r, column=2, value=text)
        c.font = NORMAL
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 115

    # --- Matrices -----------------------------------------------------------
    # Entity sheet is built FIRST so the consolidated sheet can be written as
    # live cross-sheet sums of it rather than as independently recomputed
    # literals. Consolidated therefore always equals the sum of the entities.
    entities = sorted({c["entity"] for c in vdata["companies"].values()})
    entity_sheet = "'MoM GP - by Entity'"
    entity_blocks = [(f"{e}   (HO + branch)",
                      rows_for(vdata, parents, period_keys,
                               lambda n, c, e=e: c["entity"] == e, sdata))
                     for e in entities]
    entity_rows = write_matrix(wb.create_sheet("MoM GP - by Entity"),
                               "Gross Profit by Segment - by Legal Entity",
                               "Five Tally files collapse to three legal "
                               "entities. Feeds the Consolidated sheet.",
                               period_keys, entity_blocks)

    link_rows = {}
    for seg in list(SEGMENTS) + ["TOTAL"]:
        refs = [(entity_sheet, entity_rows[(blk, seg)])
                for blk, _ in entity_blocks if (blk, seg) in entity_rows]
        if refs:
            link_rows[("", seg)] = refs

    write_matrix(wb.create_sheet("MoM GP - Consolidated"),
                 "Gross Profit by Segment - Group Consolidated",
                 "Every figure is a live sum of the 'MoM GP - by Entity' sheet.",
                 period_keys,
                 [("", rows_for(vdata, parents, period_keys, sdata=sdata))],
                 link_rows=link_rows)

    write_matrix(wb.create_sheet("MoM GP - by Company"),
                 "Gross Profit by Segment - by Tally Company", "",
                 period_keys,
                 [(f"{n}   [{COMPANY_ENTITY.get(n, ('?', '?'))[0]}]",
                   rows_for(vdata, parents, period_keys,
                            lambda x, c, t=n: x == t, sdata))
                  for n in vdata["companies"]])

    # --- Segment detail by stock group -------------------------------------
    ws = wb.create_sheet("Segment Detail")
    r = _title(ws, "Segment Detail - by Stock Group", 8,
               f"{detail_period}, item-wise.")
    for i, h in enumerate(["Segment", "Stock Group", "Sales", "COGS",
                           "Gross Profit", "GP %", "Qty"], start=1):
        c = ws.cell(row=r, column=i, value=h)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
    ws.freeze_panes = ws.cell(row=r + 1, column=1)
    r += 1
    combined = defaultdict(lambda: defaultdict(float))
    seg_of = {}
    for company, cdata in vdata["companies"].items():
        _s, by_group, _nr, _b = company_segments(
            cdata, company, detail_period, parents.get(company, {}), sdata)
        for grp, v in by_group.items():
            seg_of[grp] = v["_segment"]
            for k in ("Sales", "COGS", "Gross Profit", "Qty"):
                combined[grp][k] += v[k]
    for grp, v in sorted(combined.items(),
                         key=lambda kv: (SEGMENTS.index(seg_of.get(kv[0], "OTHER")),
                                         -kv[1]["Sales"])):
        ws.cell(row=r, column=1, value=seg_of.get(grp, "OTHER")).font = NORMAL
        ws.cell(row=r, column=2, value=grp).font = NORMAL
        for i, k in enumerate(["Sales", "COGS", "Gross Profit"], start=3):
            c = ws.cell(row=r, column=i, value=v[k])
            c.number_format = MONEY
            c.font = NORMAL
        c = ws.cell(row=r, column=6, value=f"=IF(C{r}=0,0,E{r}/C{r})")
        c.number_format = PCT
        c.font = NORMAL
        c = ws.cell(row=r, column=7, value=v["Qty"])
        c.number_format = QTY
        c.font = NORMAL
        r += 1
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 46
    for c in "CDEFG":
        ws.column_dimensions[c].width = 16

    # --- Cross check --------------------------------------------------------
    ws = wb.create_sheet("Cross Check")
    r = _title(ws, "Cross Check", 5)
    for i, h in enumerate(["Check", "Value", "Expected", "Status", "Note"], start=1):
        c = ws.cell(row=r, column=i, value=h)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
    cursor = [r + 1]

    def check(label, value, expected, note, tol=1.0, pct_fmt=False, warn=False):
        row = cursor[0]
        ok = abs(value - expected) <= max(tol, abs(expected) * 0.02)
        ws.cell(row=row, column=1, value=label).font = NORMAL
        for i, v in ((2, value), (3, expected)):
            c = ws.cell(row=row, column=i, value=v)
            c.number_format = PCT if pct_fmt else MONEY
            c.font = NORMAL
        st = ws.cell(row=row, column=4,
                     value="OK" if ok else ("REVIEW" if warn else "CHECK"))
        st.fill = OK_FILL if ok else (WARN_FILL if warn else BAD_FILL)
        st.font = BOLD
        ws.cell(row=row, column=5, value=note).font = NOTE_FONT
        cursor[0] += 1

    # months must sum to Q1
    if quarters:
        q = quarters[0]
        qmonths = quarter_months(vdata, q)
        for metric in ("Sales", "Gross Profit"):
            qa, _, _ = aggregate(vdata, parents, q, sdata=sdata)
            qv = sum(qa[s][metric] for s in SEGMENTS)
            mv = 0.0
            for mk in qmonths:
                ma, _, _ = aggregate(vdata, parents, mk, sdata=sdata)
                mv += sum(ma[s][metric] for s in SEGMENTS)
            note = ("Independent quarterly fetch vs the summed months. Sales "
                    "must match exactly." if metric == "Sales" else
                    "A few per cent gap is EXPECTED: a quarterly weighted-average "
                    "cost rate is not the sum of three monthly ones. The report "
                    "columns use the summed months.")
            check(f"Q1 {metric}: quarterly fetch vs summed months", qv, mv,
                  note, warn=(metric != "Sales"),
                  tol=1.0 if metric == "Sales" else abs(mv) * 0.05)

    # voucher sales vs ledger sales, per company
    if sdata:
        for company, cdata in vdata["companies"].items():
            segs, _g, _nr, _b = company_segments(
                cdata, company, detail_period, parents.get(company, {}), sdata)
            v_sales = sum(segs[s]["Sales"] for s in SEGMENTS)
            led = sdata["companies"].get(company, {}).get("ledgers", {}).get(
                detail_period, [])
            l_sales = sum(x["amount"] for x in led if x["bucket"] == "sales")
            check(f"{COMPANY_ENTITY.get(company, ('?',))[0]} "
                  f"{'HO' if 'NOIDA' not in company.upper() else 'Noida'}: "
                  "voucher vs ledger sales",
                  v_sales, l_sales,
                  "Item-line sales vs sales ledgers (incl. branch and related "
                  "party). Gaps are non-inventory revenue - software, services, "
                  "freight recovered - which never appears on an inventory line.",
                  warn=True, tol=abs(l_sales) * 0.05)

    # how much of COGS rests on a fallback rate
    _a, no_rate, basis = aggregate(vdata, parents, detail_period, sdata=sdata)
    total_basis = sum(basis.values())
    for key, label in (("period", "period weighted average"),
                       ("ytd", "year-to-date average (fallback)"),
                       ("closing", "closing rate (fallback)"),
                       ("none", "NO cost rate available")):
        if not total_basis:
            break
        row = cursor[0]
        share = basis.get(key, 0.0) / total_basis
        ws.cell(row=row, column=1, value=f"Sales costed on {label}").font = NORMAL
        c = ws.cell(row=row, column=2, value=share)
        c.number_format = PCT
        c.font = NORMAL
        ws.cell(row=row, column=3, value="-").font = NORMAL
        bad = (key == "none" and share > 0.02)
        st = ws.cell(row=row, column=4, value="REVIEW" if bad else "OK")
        st.fill = WARN_FILL if bad else OK_FILL
        st.font = BOLD
        ws.cell(row=row, column=5,
                value=("Sales with no derivable cost show 100% margin and "
                       "overstate GP." if key == "none"
                       else "Share of sales value costed on this basis.")).font = NOTE_FONT
        cursor[0] += 1

    # segment credibility
    seg_agg, _, _ = aggregate(vdata, parents, detail_period, sdata=sdata)
    for s in SEGMENTS:
        if abs(seg_agg[s]["Sales"]) < 1:
            continue
        row = cursor[0]
        pct = seg_agg[s]["GP %"]
        bad = not (0.0 <= pct <= 0.75)
        ws.cell(row=row, column=1, value=f"{s}: GP% credible?").font = NORMAL
        c = ws.cell(row=row, column=2, value=pct)
        c.number_format = PCT
        c.font = NORMAL
        ws.cell(row=row, column=3, value="0% to 75%").font = NORMAL
        st = ws.cell(row=row, column=4, value="OK" if not bad else "REVIEW")
        st.fill = OK_FILL if not bad else WARN_FILL
        st.font = BOLD
        ws.cell(row=row, column=5,
                value="Within a plausible trading range." if not bad
                else "Outside the plausible range - investigate before quoting."
                ).font = NOTE_FONT
        cursor[0] += 1

    ws.column_dimensions["A"].width = 46
    for c, w in (("B", 18), ("C", 18), ("D", 10), ("E", 92)):
        ws.column_dimensions[c].width = w

    # --- Board reconciliation to the statutory P&L --------------------------
    q1_agg = aggregate_column(vdata, parents, detail_period, None, sdata)
    item_sales = sum(q1_agg[s]["Sales"] for s in SEGMENTS)
    item_cogs = sum(q1_agg[s]["COGS"] for s in SEGMENTS)
    item_gp = sum(q1_agg[s]["Gross Profit"] for s in SEGMENTS)

    branch_sales = 0.0
    if sdata:
        for _c, cd in sdata["companies"].items():
            for row in cd.get("ledgers", {}).get(detail_period, []):
                if row.get("is_branch") and row.get("bucket") == "sales":
                    branch_sales += row["amount"]

    pl_figs = read_pl_workbook(pl_path) if pl_path and Path(pl_path).exists() else None
    if pl_figs:
        write_reconciliation(wb, pl_figs, Path(pl_path).name, item_sales,
                             item_cogs, item_gp, branch_sales, detail_period)

    # --- Index --------------------------------------------------------------
    cons = "'MoM GP - Consolidated'"
    q1_col = period_keys.index(detail_period) * len(METRICS) + 2
    total_row = None
    ws_cons = wb["MoM GP - Consolidated"]
    for rr in range(1, ws_cons.max_row + 1):
        if str(ws_cons.cell(row=rr, column=1).value or "").upper().startswith("TOTAL"):
            total_row = rr
            break
    sales_ref = "{0}!{1}{2}".format(cons, get_column_letter(q1_col), total_row)
    gp_ref = "{0}!{1}{2}".format(cons, get_column_letter(q1_col + 2), total_row)

    headline = [
        ("{0} Sales".format(detail_period), "=" + sales_ref,
         "Live link to the Consolidated sheet."),
        ("{0} Gross Profit".format(detail_period), "=" + gp_ref,
         "Live link to the Consolidated sheet."),
        ("{0} GP %".format(detail_period),
         "=IF({0}=0,0,{1}/{0})".format(sales_ref, gp_ref),
         "Calculated, not typed."),
    ]

    meta_rows = [
        ("Report", "Gross Profit by segment - Machines, Heads, Ink, "
                   "Spare Parts, Other"),
        ("Source", "Tally Prime, live. Sales and purchase voucher inventory "
                   "lines, read item by item."),
        ("Generated", vdata.get("generated", "")),
        ("Period", "{0} to {1}   (monthly, plus a {2} column)".format(
            vdata.get("fy_start", ""), vdata.get("through", ""), detail_period)),
        ("Basis", "Segment comes from each stock item's own stock group, not "
                  "from the ledger name - which is why it covers all five "
                  "companies."),
        ("Reconciliation", "The Board Reconciliation sheet bridges this report "
                           "to the statutory P&L, line by line, to nil."),
    ]

    sheet_notes = [
        ("Basis & Method", "How gross profit is struck, what is deliberately "
                           "excluded, and the traps that corrupt this report."),
        ("MoM GP - Consolidated", "Group total by segment by month. Every "
                                  "figure is a live sum of the by-Entity "
                                  "sheet."),
        ("MoM GP - by Entity", "The three legal entities. This sheet feeds "
                               "Consolidated."),
        ("MoM GP - by Company", "All five Tally files separately."),
        ("Segment Detail", "Which stock groups drive each segment."),
        ("Board Reconciliation", "Bridges this report to the P&L - revenue, "
                                 "cost and gross profit, each to nil."),
        ("Cross Check", "Every validation run on the numbers. Read before "
                        "quoting anything."),
    ]
    sheet_notes = [(n, d) for n, d in sheet_notes if n in wb.sheetnames]

    write_index(wb, meta_rows, sheet_notes, headline)

    order = ["Index", "Basis & Method", "MoM GP - Consolidated",
             "MoM GP - by Entity", "MoM GP - by Company", "Segment Detail",
             "Board Reconciliation", "Cross Check"]
    for pos, name in enumerate([n for n in order if n in wb.sheetnames]):
        wb.move_sheet(name, offset=pos - wb.sheetnames.index(name))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--vouchers", default=str(PROJECT_ROOT / "Files" / "gp_voucher_cache.json"))
    ap.add_argument("--segments", default=str(PROJECT_ROOT / "Files" / "gp_segment_cache.json"))
    ap.add_argument("--out", default=None)
    ap.add_argument("--pl", default=None,
                    help="PL_BS_GL_Detail workbook to reconcile against. "
                         "Defaults to the newest one directly in Files\\.")
    args = ap.parse_args()
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out = Path(args.out) if args.out else (
        PROJECT_ROOT / "Files" / f"GP_Segment_ItemWise_{stamp}.xlsx")

    if args.pl:
        pl = Path(args.pl)
    else:
        cands = sorted((PROJECT_ROOT / "Files").glob("PL_BS_GL_Detail_*.xlsx"))
        pl = cands[-1] if cands else None
    if pl:
        print(f"Reconciling against {pl.name}")
    print(f"Wrote {build(Path(args.vouchers), Path(args.segments), out, pl)}")


if __name__ == "__main__":
    main()
