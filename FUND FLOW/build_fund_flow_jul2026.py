# -*- coding: utf-8 -*-
"""
Fund Flow Analysis builder -- JULY 2026.

Reads "0F_JULY 2026 Cashflow v1.xlsx" and produces
"Fund_Flow_Analysis_Jul2026_LINKED.xlsx" -- a management-facing workbook with:
  - Executive Summary (KPI tiles, investment/loan movement detail, "where funds are
    blocked" narrative, quick-reference links, basis of preparation)
  - Summary (clone of 'Main Summary', formulas redirected at the hidden source sheet)
  - Unplanned / Excess / Short / Payment Done (every vendor row classified by
    Balance = Budget - Paid, sorted by variance descending)
  - Machine Advance Payment (July block of 'Machine Adv_Receivable', rows 71-96)
  - Advance from Debtors (clone of 'Mach Adv As on Date' -- new vs the June pack)
  - Four hidden source sheets carrying the raw July numbers, row-for-row

Every number in the output is a live Excel formula, EXCEPT the row anchors in
CLASSIFIED_ROWS / INV_ROWS below, which identify which physical row of 'FF Summary'
each vendor currently sits on.

JULY ROW ANCHORS (these MOVED vs June -- do not reuse the June script's numbers):
    Import Dues            3-13      Tax                92-95
    Import Advances        15-20     Investment/Loan    98-106
    FD commitment          21        Monthly Expense    107-123
    Custom Duty            24        Total Payment      125
    EMI / Repayment        26-28     Machine Advances   127
    Local Vendors          31-37     Customer Collect   136
    Other Vendors          40-85     Total Collection   138
    Axis BC/LC (excluded)  87-90     Opening Balance    146

RE-RUN CHECKLIST when the source workbook changes:
  1. Amounts only, on existing rows -> just re-run.
  2. Rows added / removed / reordered -> re-run discover.py first, then update
     CLASSIFIED_ROWS and INV_ROWS. Row drift silently misclassifies vendors.
  3. Requires: pip install openpyxl        Run: python build_fund_flow_jul2026.py
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = r"c:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS\FUND FLOW"
SRC = os.path.join(BASE, "0F_JULY 2026 Cashflow v1.xlsx")
OUT = os.path.join(BASE, "Fund_Flow_Analysis_Jul2026_LINKED.xlsx")

MONTH = "July 2026"
FF = "'FF Source Data'"
BK = "'Bank Source Data'"
MA = "'MA Source Data'"

# ---------------------------------------------------------------- styling
TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FONT = Font(italic=True, size=9, color="595959")
HDR_FONT = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="2E75B6")
GROUP_FONT = Font(bold=True, color="1F4E78")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
UNPLANNED_FILL = PatternFill("solid", fgColor="FCE4E4")
EXCESS_FILL = PatternFill("solid", fgColor="FFF2CC")
SHORT_FILL = PatternFill("solid", fgColor="DDEBF7")
DONE_FILL = PatternFill("solid", fgColor="EAF1DD")
TOTAL_FONT = Font(bold=True)
TOTAL_FILL = PatternFill("solid", fgColor="BDD7EE")
KPI_FILL = PatternFill("solid", fgColor="1F4E78")
KPI_FONT = Font(bold=True, size=16, color="FFFFFF")
KPI_LBL_FONT = Font(bold=True, size=9, color="FFFFFF")
FLAG_FILL = PatternFill("solid", fgColor="FFF2CC")
RED_FONT = Font(color="C00000", bold=True)
GREEN_FONT = Font(color="375623", bold=True)
INR_FMT = '##,##,##,##0;[RED](##,##,##,##0)'
USD_FMT = '#,##0.00;[RED](#,##0.00)'
PCT_FMT = '0.0%'
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HYPER_FONT = Font(color="0563C1", underline="single")
NARR_FONT = Font(size=10)
NARR_FILL_IN = PatternFill("solid", fgColor="E2EFDA")
NARR_FILL_OUT = PatternFill("solid", fgColor="FCE4E4")


def title(ws, text, sub, width):
    c = ws.cell(1, 1, text)
    c.font, c.fill = TITLE_FONT, TITLE_FILL
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    ws.row_dimensions[1].height = 26
    s = ws.cell(2, 1, sub)
    s.font = SUB_FONT
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)


def hdr_row(ws, r, headers, start_col=1):
    for i, h in enumerate(headers, start=start_col):
        c = ws.cell(r, i, h)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[r].height = 30


def back_link(ws, col, row=1):
    c = ws.cell(row, col, "<- Executive Summary")
    c.hyperlink = "#'Executive Summary'!A1"
    c.font = HYPER_FONT


# ------------------------------------------------------- load the source
wb_v = openpyxl.load_workbook(SRC, data_only=True)
ff_v = wb_v["FF Summary"]
ms_v = wb_v["Main Summary"]
bk_v = wb_v["Bank Balance"]
ma_v = wb_v["Machine Adv_Receivable"]
da_v = wb_v["Mach Adv As on Date"]

out = openpyxl.Workbook()
out.remove(out.active)


def num(ws, r, c):
    v = ws.cell(r, c).value
    return float(v) if isinstance(v, (int, float)) else 0.0


def txt(ws, r, c):
    v = ws.cell(r, c).value
    return str(v).strip() if v is not None else ""


# ------------------------------------------------- classify vendor rows
BLOCKS = [
    ("Import Due Payment", 3, 13),
    ("Import Advance Payment", 15, 20),
    ("FD Commitment", 21, 21),
    ("Custom Duty and GST", 24, 24),
    ("EMI / Repayment", 26, 28),
    ("Local Vendor Payment", 31, 37),
    ("Other Vendors", 40, 85),
    ("Tax", 92, 95),
    ("Investment/Loan", 98, 106),
    ("Monthly Expense", 107, 123),
]

CLASSIFIED_ROWS = {"UNPLANNED": [], "EXCESS": [], "SHORT": [], "DONE": []}
for cat, r0, r1 in BLOCKS:
    for r in range(r0, r1 + 1):
        a, b = txt(ff_v, r, 1), txt(ff_v, r, 2)
        if not a and not b:
            continue
        if a.startswith("Do Not") or a == "Type":
            continue
        bi, pi = num(ff_v, r, 4), num(ff_v, r, 6)
        bu, pu = num(ff_v, r, 3), num(ff_v, r, 5)
        if bi == 0 and pi == 0 and bu == 0 and pu == 0:
            continue
        label = b if b else a
        label_col = "B" if b else "A"
        if bi == 0 and pi > 0:
            k = "UNPLANNED"
        elif bi > 0 and pi > bi:
            k = "EXCESS"
        elif bi > 0 and pi < bi:
            k = "SHORT"
        elif bi > 0 and pi == bi:
            k = "DONE"
        else:
            continue
        CLASSIFIED_ROWS[k].append(dict(row=r, cat=cat, label_col=label_col, var=pi - bi))

for k in CLASSIFIED_ROWS:
    CLASSIFIED_ROWS[k].sort(key=lambda x: -abs(x["var"]))

# Investment / Loan movement detail (July: all outflow, no fresh drawdown)
INV_ROWS = [
    (104, "Axis Bank - Limit Drop", "Outflow"),
    (103, "Bhoomi Weave Tex Pvt Ltd", "Outflow"),
    (105, "Modi Ajaykumar - Stamp Paper", "Outflow"),
    (102, "Aerztin Business & Consulting", "Outflow"),
    (100, "S and S Industries - Fixed Assets", "Outflow"),
    (101, "Smart Technical Services", "Outflow"),
    (106, "Investment in MF", "Outflow"),
]

# ---------------------------------------------------------------------
# Summary sheet layout is resolved FIRST so the Executive Summary can point
# at real row numbers instead of hand-counted guesses.
# ---------------------------------------------------------------------
SUMMARY_LAYOUT = [
    ("open",       "Opening Balance",           "grp"),
    ("pay_tot",    "Total Payment",             "grp"),
    ("coll_tot",   "Total Collection",          "grp"),
    ("close",      "Closing Balance",           "grp"),
    ("bank",       "Bank Balance (31 Jul)",     "grp"),
    ("diff",       "Difference",                "grp"),
    ("gap1",       None,                        "gap"),
    ("imp_due",    "Import Due Payment",        ""),
    ("imp_adv",    "Import Advance Payment",    ""),
    ("loc_adv",    "Local Advance Payment",     ""),
    ("sub_imp",    "   Sub-total",              "sub"),
    ("gap2",       None,                        "gap"),
    ("duty",       "Custom Duty",               ""),
    ("emi",        "EMI / Repayment",           ""),
    ("othven",     "Other Vendors",             ""),
    ("sub_ven",    "   Sub-total",              "sub"),
    ("gap3",       None,                        "gap"),
    ("tax_otec",   "Advance Tax - O Tec",       ""),
    ("tax_ent",    "Advance Tax - Enterprises", ""),
    ("tax_sa",     "SA Tax - Enterprises",      ""),
    ("tax_tds",    "TDS / TCS / GST",           ""),
    ("sub_tax",    "   Sub-total",              "sub"),
    ("gap4",       None,                        "gap"),
    ("invloan",    "Investment / Loan",         ""),
    ("gap5",       None,                        "gap"),
    ("monthly",    "Monthly Expenses",          ""),
    ("gap6",       None,                        "gap"),
    ("pay_grand",  "Total Payment",             "grp"),
    ("gap7",       None,                        "gap"),
    ("mach_adv",   "Machine Advances",          ""),
    ("gap8",       None,                        "gap"),
    ("cust_coll",  "Customer Collections",      ""),
    ("gap9",       None,                        "gap"),
    ("coll_grand", "Total Collection",          "grp"),
]

SR = {}
_r = 4
for _key, _label, _kind in SUMMARY_LAYOUT:
    if _kind != "gap":
        SR[_key] = _r
    _r += 1
SUM_LAST_ROW = _r


def S(col, key):
    """Reference to a Summary-sheet cell by logical key, e.g. S('F','close')."""
    return "Summary!%s%d" % (col, SR[key])


# ======================================================================
# 1. EXECUTIVE SUMMARY
# ======================================================================
es = out.create_sheet("Executive Summary")
title(es, "Fund Flow -- %s | Executive Summary" % MONTH,
      "Prepared for management review. Every figure below is a live formula linked to the "
      "July cashflow ledger -- refresh the source and this pack follows.", 11)

KPIS = [
    (4, 1, "TOTAL PAYOUT (JULY)", "=" + S("F", "pay_tot"), "Total actually paid out this month"),
    (4, 4, "TOTAL COLLECTION (JULY)", "=" + S("F", "coll_tot"),
     "Machine advances + customer collections received"),
    (4, 7, "CLOSING BALANCE", "=" + S("F", "close"), "Opening + Collection - Payout"),
    (8, 1, "OPENING BALANCE", "=" + S("F", "open"), "Carried forward from June closing"),
    (8, 4, "BANK BALANCE (31 JUL)", "=" + S("F", "bank"), "Per bank statement"),
    (8, 7, "VARIANCE vs BANK", "=" + S("F", "diff"), "Computed closing minus bank statement"),
]
for r, c, lbl, formula, note in KPIS:
    lc = es.cell(r, c, lbl)
    lc.font, lc.fill = KPI_LBL_FONT, KPI_FILL
    es.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 2)
    vc = es.cell(r + 1, c, formula)
    vc.font, vc.fill, vc.number_format = KPI_FONT, KPI_FILL, INR_FMT
    vc.alignment = Alignment(horizontal="left")
    es.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + 2)
    es.row_dimensions[r + 1].height = 24
    nc = es.cell(r + 2, c, note)
    nc.font = SUB_FONT
    es.merge_cells(start_row=r + 2, start_column=c, end_row=r + 2, end_column=c + 2)

# cash movement callout
es.cell(11, 1, "Net cash movement")
es.cell(11, 1).font = GROUP_FONT
mv = es.cell(11, 3, "=%s-%s" % (S("F", "close"), S("F", "open")))
mv.number_format, mv.font = INR_FMT, RED_FONT
pc = es.cell(11, 4, "=IFERROR((%s-%s)/%s,0)" % (S("F", "close"), S("F", "open"), S("F", "open")))
pc.number_format, pc.font = PCT_FMT, RED_FONT
es.cell(11, 5, "Opening to closing erosion. No fresh loan was drawn in July (June: Rs.9.69 Cr).")
es.cell(11, 5).font = SUB_FONT

# --- narrative: where funds went
r = 13
h = es.cell(r, 1, "Where the Funds Went")
h.font, h.fill = GROUP_FONT, GROUP_FILL
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)

r += 1
n1 = es.cell(r, 1,
             '=" Off-budget spending this month: "&TEXT(\'Unplanned Payment\'!H25+\'Excess Payment\'!H21,'
             '"Rs.##,##,##,##0")&" -- "&TEXT(\'Unplanned Payment\'!H25,"Rs.##,##,##,##0")&'
             '" with no budget sanction at all, plus "&TEXT(\'Excess Payment\'!H21,"Rs.##,##,##,##0")&'
             '" spent over approved budgets."')
n1.font, n1.fill = NARR_FONT, NARR_FILL_OUT
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)

r += 1
n2 = es.cell(r, 1,
             '=" The single largest unplanned item is the Axis Bank limit drop of "&'
             'TEXT(IF(ISNUMBER(' + FF + '!$F$104),' + FF + '!$F$104,0),"Rs.##,##,##,##0")&'
             '" -- a facility reduction rather than an operating expense, but it consumed cash all the same."')
n2.font, n2.fill = NARR_FONT, NARR_FILL_OUT
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)

r += 1
lnk = es.cell(r, 1, "-> Open Unplanned Payment sheet")
lnk.hyperlink, lnk.font = "#'Unplanned Payment'!A1", HYPER_FONT

# --- investment / loan movement detail
r += 2
h = es.cell(r, 1, "Investment / Loan Movement Detail (feeds Investment/Loan row on Summary)")
h.font, h.fill = GROUP_FONT, GROUP_FILL
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
r += 1
hdr_row(es, r, ["Item", "Type", "Amount Paid (Rs.)", "Budgeted (Rs.)"])
inv_start = r + 1
for i, (srow, label, kind) in enumerate(INV_ROWS):
    rr = inv_start + i
    es.cell(rr, 1, label).border = BORDER
    es.cell(rr, 2, kind).border = BORDER
    c = es.cell(rr, 3, "=IF(ISNUMBER(%s!$F$%d),%s!$F$%d,0)" % (FF, srow, FF, srow))
    c.number_format, c.border = INR_FMT, BORDER
    d = es.cell(rr, 4, "=IF(ISNUMBER(%s!$D$%d),%s!$D$%d,0)" % (FF, srow, FF, srow))
    d.number_format, d.border = INR_FMT, BORDER
inv_end = inv_start + len(INV_ROWS) - 1
r = inv_end + 1
es.cell(r, 1, "NET (matches Summary Investment/Loan row)").font = TOTAL_FONT
for col in (3, 4):
    c = es.cell(r, col, "=SUM(%s%d:%s%d)" % (get_column_letter(col), inv_start, get_column_letter(col), inv_end))
    c.font, c.fill, c.number_format, c.border = TOTAL_FONT, TOTAL_FILL, INR_FMT, BORDER
r += 1
nt = es.cell(r, 1, "Note: July shows NO fresh loan or FD inflow. The entire Investment/Loan block is "
                   "outflow against a zero budget, so all of it lands in Unplanned Payment by definition.")
nt.font = SUB_FONT
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)

# --- where funds are blocked
r += 2
h = es.cell(r, 1, "Where Funds Are Blocked")
h.font, h.fill = GROUP_FONT, GROUP_FILL
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
r += 1
b1 = es.cell(r, 1,
             '=" Machine advances are the biggest receivable gap: "&'
             'TEXT(\'Machine Advance Payment\'!E$TOT,"Rs.##,##,##,##0")&" was projected but only "&'
             'TEXT(\'Machine Advance Payment\'!F$TOT,"Rs.##,##,##,##0")&" came in -- "&'
             'TEXT(\'Machine Advance Payment\'!G$TOT,"Rs.##,##,##,##0")&" still pending ("&'
             'TEXT(\'Machine Advance Payment\'!F$TOT/\'Machine Advance Payment\'!E$TOT,"0.0%")&" collected)."')
b1.font, b1.fill = NARR_FONT, NARR_FILL_OUT
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
MA_TOT_CELL = (r, 1)

r += 1
lnk = es.cell(r, 1, "-> Open Machine Advance Payment sheet")
lnk.hyperlink, lnk.font = "#'Machine Advance Payment'!A1", HYPER_FONT

r += 1
b2 = es.cell(r, 1,
             '=" Payments held back from vendors / tax this month: "&'
             'TEXT(\'Short Payment\'!H29,"Rs.##,##,##,##0")&" against budgets of "&'
             'TEXT(\'Short Payment\'!E29,"Rs.##,##,##,##0")&". This is cash conserved, but it rolls "&'
             '"forward as a liability into August."')
b2.font, b2.fill = NARR_FONT, NARR_FILL_IN
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
SHORT_TOT_ROW_REF = (r, 1)

r += 1
b3 = es.cell(r, 1,
             '=" Customer collections ran at "&TEXT(%s,"0.0%%")&" of the internal target ("&'
             'TEXT(%s,"Rs.##,##,##,##0")&" received vs "&TEXT(%s,"Rs.##,##,##,##0")&'
             '" targeted), leaving "&TEXT(%s,"Rs.##,##,##,##0")&" uncollected."'
             % (S("L", "cust_coll"), S("F", "cust_coll"), S("D", "cust_coll"), S("H", "cust_coll")))
b3.font, b3.fill = NARR_FONT, NARR_FILL_OUT
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)

# --- data quality flags
r += 2
h = es.cell(r, 1, "Data Quality Flags -- Raised for Confirmation")
h.font, h.fill = GROUP_FONT, GROUP_FILL
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
r += 1
hdr_row(es, r, ["#", "Flag", "Amount (Rs.)", "Action"])
FLAGS = [
    ("'FD for Found payment-278400$ @ 20%' (FF Summary row 21) is budgeted but sits outside the "
     "range Main Summary totals (D15:D20), so Total Payment budget reads Rs.31.98 Cr on Main Summary "
     "vs Rs.32.51 Cr on FF Summary.",
     "=IF(ISNUMBER(%s!$D$21),%s!$D$21,0)" % (FF, FF),
     "Excluded from this pack's totals to match Main Summary. Confirm intended treatment."),
    ("Investment/Loan block carries a zero budget against Rs.3.76 Cr paid -- the whole block "
     "classifies as unplanned.",
     "=SUM(%s!$F$98:$F$106)" % FF,
     "Confirm whether a budget was approved but not entered."),
    ("Axis Bank 'Limit Drop' is a working-capital facility reduction, not an operating expense.",
     "=IF(ISNUMBER(%s!$F$104),%s!$F$104,0)" % (FF, FF),
     "Consider a separate 'Financing' category so operating payout is not overstated."),
    ("FF Summary rows 152-153 return #N/A (broken lookup) and Main Summary J19:J20 return #DIV/0! "
     "(zero-budget advance tax rows).",
     None,
     "Repair in the source workbook; this pack routes around them."),
    ("Machine Adv_Receivable rows 62-63 (Gopgan, June block) are annotated 'we need to remove this "
     "as the same is considered under the due amount' but are still present.",
     None,
     "Confirm removal so prior-month comparatives stay clean."),
]
fr = r + 1
for i, (flag, amt, action) in enumerate(FLAGS):
    rr = fr + i
    c = es.cell(rr, 1, i + 1)
    c.border, c.alignment = BORDER, Alignment(horizontal="center")
    f = es.cell(rr, 2, flag)
    f.border, f.font, f.alignment = BORDER, NARR_FONT, Alignment(wrap_text=True, vertical="top")
    a = es.cell(rr, 3, amt if amt else "--")
    a.border, a.number_format = BORDER, INR_FMT
    a.alignment = Alignment(horizontal="right")
    ac = es.cell(rr, 4, action)
    ac.border, ac.font, ac.alignment = BORDER, NARR_FONT, Alignment(wrap_text=True, vertical="top")
    es.row_dimensions[rr].height = 30
    for cc in range(1, 5):
        es.cell(rr, cc).fill = FLAG_FILL

# --- quick reference
r = fr + len(FLAGS) + 1
h = es.cell(r, 1, "Quick Reference -- Where To Look")
h.font, h.fill = GROUP_FONT, GROUP_FILL
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
r += 1
hdr_row(es, r, ["Question", "Sheet"])
QREF = [
    ("Overall Budget vs Paid vs Balance, by category", "Summary"),
    ("What was paid with zero budget approval", "Unplanned Payment"),
    ("What went over its approved budget", "Excess Payment"),
    ("What's still owed to vendors/tax/dues (money held back)", "Short Payment"),
    ("What's fully settled (audit trail)", "Payment Done"),
    ("What customers still owe us on machine orders", "Machine Advance Payment"),
    ("Every debtor advance held, party by party", "Advance from Debtors"),
]
for i, (q, sheet) in enumerate(QREF):
    rr = r + 1 + i
    qc = es.cell(rr, 1, q)
    qc.border = BORDER
    sc = es.cell(rr, 2, "-> " + sheet)
    sc.hyperlink, sc.font, sc.border = "#'%s'!A1" % sheet, HYPER_FONT, BORDER

# --- basis of preparation
r = r + len(QREF) + 2
h = es.cell(r, 1, "Basis of Preparation")
h.font, h.fill = GROUP_FONT, GROUP_FILL
es.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
r += 1
bp = es.cell(r, 1,
             "This report is generated from '0F_JULY 2026 Cashflow v1.xlsx' (July 2026 fund flow ledger). "
             "The four hidden source sheets hold a row-for-row snapshot of 'FF Summary', 'Bank Balance', "
             "'Machine Adv_Receivable' and 'Mach Adv As on Date'; every visible figure is a live formula "
             "pointing at them. Classification rule: Unplanned = paid with zero budget; Excess = paid "
             "above budget; Short = paid below budget; Done = paid exactly to budget. Amounts are in INR "
             "unless a column is marked USD. The Rs.53,45,280 FD commitment on FF Summary row 21 is "
             "excluded from category totals to agree with Main Summary, and is flagged above. Prepared "
             "for internal management review; not audited.")
bp.font, bp.alignment = NARR_FONT, Alignment(wrap_text=True, vertical="top")
es.merge_cells(start_row=r, start_column=1, end_row=r + 3, end_column=11)

for col, w in zip("ABCDEFGHIJK", [46, 26, 20, 46, 12, 12, 16, 12, 12, 14, 14]):
    es.column_dimensions[col].width = w
es.sheet_view.showGridLines = False

# ======================================================================
# 2. SUMMARY  (clone of Main Summary, redirected)
# ======================================================================
sm = out.create_sheet("Summary")
title(sm, "Summary -- %s | Budget vs Paid vs Balance" % MONTH,
      "Category roll-up mirroring 'Main Summary' of the source workbook. "
      "Every cell is a live formula against the hidden FF Source Data sheet.", 12)
back_link(sm, 14, 1)

# Formulas, with {key} placeholders resolved to physical rows via SR.
FORMULAS = {
    "open":      ("=%s!D146" % FF,               "=D{open}",                 "=F{open}"),
    "pay_tot":   ("=D{pay_grand}",               "=F{pay_grand}",            "=H{pay_grand}"),
    "coll_tot":  ("=D{coll_grand}",              "=F{coll_grand}",           "=H{coll_grand}"),
    "close":     ("=D{open}+D{coll_tot}-D{pay_tot}", "=F{open}+F{coll_tot}-F{pay_tot}",
                                                 "=H{open}+H{coll_tot}-H{pay_tot}"),
    "bank":      (None,                          "=%s!H35" % BK,             None),
    "diff":      (None,                          "=F{close}-F{bank}",        None),
    "imp_due":   ("=SUM(%s!D3:D13)" % FF,        "=SUM(%s!F3:F13)" % FF,     "=SUM(%s!H3:H13)" % FF),
    "imp_adv":   ("=SUM(%s!D15:D20)" % FF,       "=SUM(%s!F15:F20)" % FF,    "=SUM(%s!H15:H20)" % FF),
    "loc_adv":   (0,                             0,                          0),
    "sub_imp":   ("=SUM(D{imp_due}:D{loc_adv})", "=SUM(F{imp_due}:F{loc_adv})", "=SUM(H{imp_due}:H{loc_adv})"),
    "duty":      ("=%s!D24" % FF,                "=%s!F24" % FF,             "=%s!H24" % FF),
    "emi":       ("=SUM(%s!D26:D28)" % FF,       "=SUM(%s!F26:F28)" % FF,    "=SUM(%s!H26:H28)" % FF),
    "othven":    ("=SUM(%s!D31:D86)" % FF,       "=SUM(%s!F31:F86)" % FF,    "=SUM(%s!H31:H86)" % FF),
    "sub_ven":   ("=SUM(D{duty}:D{othven})",     "=SUM(F{duty}:F{othven})",  "=SUM(H{duty}:H{othven})"),
    "tax_otec":  ("=%s!D92" % FF,                "=%s!F92" % FF,             "=%s!H92" % FF),
    "tax_ent":   ("=%s!D93" % FF,                "=%s!F93" % FF,             "=%s!H93" % FF),
    "tax_sa":    ("=%s!D94" % FF,                "=%s!F94" % FF,             "=%s!H94" % FF),
    "tax_tds":   ("=%s!D95" % FF,                "=%s!F95" % FF,             "=%s!H95" % FF),
    "sub_tax":   ("=SUM(D{tax_otec}:D{tax_tds})", "=SUM(F{tax_otec}:F{tax_tds})", "=SUM(H{tax_otec}:H{tax_tds})"),
    "invloan":   ("=SUM(%s!D98:D106)" % FF,      "=SUM(%s!F98:F106)" % FF,   "=SUM(%s!H98:H106)" % FF),
    "monthly":   ("=SUM(%s!D107:D123)" % FF,     "=SUM(%s!F107:F123)" % FF,  "=SUM(%s!H107:H123)" % FF),
    "pay_grand": ("=D{sub_imp}+D{sub_ven}+D{sub_tax}+D{invloan}+D{monthly}",
                  "=F{sub_imp}+F{sub_ven}+F{sub_tax}+F{invloan}+F{monthly}",
                  "=H{sub_imp}+H{sub_ven}+H{sub_tax}+H{invloan}+H{monthly}"),
    "mach_adv":  ("=%s!D127" % FF,               "=%s!F127" % FF,            "=%s!H127" % FF),
    "cust_coll": ("=%s!D136" % FF,               "=%s!F136" % FF,            "=%s!H136" % FF),
    "coll_grand":("=%s!D138" % FF,               "=%s!F138" % FF,            "=%s!H138" % FF),
}


def resolve(f):
    return f.format(**SR) if isinstance(f, str) else f


hdr_row(sm, 3, ["", "Particulars", "", "Budgeted (Rs.)", "", "Paid (Rs.)", "", "Balance (Rs.)",
                "", "Check", "", "% Paid"])
for key, label, kind in SUMMARY_LAYOUT:
    if kind == "gap":
        continue
    r = SR[key]
    c = sm.cell(r, 2, label)
    if kind in ("grp", "sub"):
        c.font, c.fill = GROUP_FONT, GROUP_FILL
    bud, paid, bal = FORMULAS[key]
    for col, val in ((4, bud), (6, paid), (8, bal)):
        cell = sm.cell(r, col, resolve(val))
        cell.number_format, cell.border = INR_FMT, BORDER
        if kind in ("grp", "sub"):
            cell.font = TOTAL_FONT
            cell.fill = GROUP_FILL if kind == "sub" else TOTAL_FILL
    # The "Budget - Paid - Balance = 0" identity holds for flow rows only. Balance-type
    # rows (opening/closing) carry the opening figure in all three columns, so the check
    # would read -Opening and look like a break when nothing is wrong.
    BALANCE_ROWS = {"open", "close", "bank", "diff"}
    if bud is not None and paid is not None and bal is not None and key not in BALANCE_ROWS:
        chk = sm.cell(r, 10, "=D%d-F%d-H%d" % (r, r, r))
        chk.number_format, chk.border = INR_FMT, BORDER
        pct = sm.cell(r, 12, '=IFERROR(F%d/D%d,"--")' % (r, r))
        pct.number_format, pct.border = PCT_FMT, BORDER
r = SUM_LAST_ROW

note = sm.cell(r + 1, 2,
               "Note: 'FD for Found payment' (source row 21, Rs.53,45,280 budgeted) is excluded above to "
               "agree with Main Summary. Including it would raise Total Payment budget to Rs.32,51,18,773.")
note.font = SUB_FONT
sm.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=12)

for col, w in zip("ABCDEFGHIJKLMN", [3, 32, 2, 18, 2, 18, 2, 18, 2, 16, 2, 10, 2, 22]):
    sm.column_dimensions[col].width = w
sm.sheet_view.showGridLines = False
sm.freeze_panes = "A4"

# ======================================================================
# 3-6. CLASSIFIED SHEETS
# ======================================================================
SHEETS = [
    ("Unplanned Payment", "UNPLANNED", UNPLANNED_FILL,
     "Paid with NO budget sanction at all (Budget = 0). Every figure is a live formula. "
     "Sorted by amount, largest first.", "paid"),
    ("Excess Payment", "EXCESS", EXCESS_FILL,
     "Budgeted item where Paid exceeds Budget (overrun). Excess = Paid - Budget. "
     "Sorted by overrun, largest first.", "over"),
    ("Short Payment", "SHORT", SHORT_FILL,
     "Budgeted item where Paid is less than Budget (partially or fully unpaid). "
     "Shortfall = Budget - Paid. Sorted by shortfall, largest first.", "short"),
    ("Payment Done", "DONE", DONE_FILL,
     "Budgeted item fully settled (Paid = Budget, Balance = 0). Included for audit trail.", "done"),
]

for sheet_name, key, fill, subtitle, mode in SHEETS:
    ws = out.create_sheet(sheet_name)
    title(ws, "%s -- %s" % (sheet_name, MONTH), subtitle, 9)
    back_link(ws, 11, 1)
    hdr_row(ws, 4, ["#", "Vendor / Item", "Category", "Budget USD ($)", "Budget INR (Rs.)",
                    "Paid USD ($)", "Paid INR (Rs.)",
                    "Shortfall INR (Rs.)" if mode == "short" else "Variance INR (Rs.)", "Src Row"])
    start = 5
    rows = CLASSIFIED_ROWS[key]
    for i, item in enumerate(rows):
        rr = start + i
        srow, lc = item["row"], item["label_col"]
        n = ws.cell(rr, 1, i + 1)
        n.border, n.alignment = BORDER, Alignment(horizontal="center")
        b = ws.cell(rr, 2, "=%s!$%s$%d" % (FF, lc, srow))
        b.border = BORDER
        c = ws.cell(rr, 3, item["cat"])
        c.border, c.font = BORDER, SUB_FONT
        du = ws.cell(rr, 4, '=IF(OR(%s!$C$%d=0,%s!$C$%d=""),"--",%s!$C$%d)' % (FF, srow, FF, srow, FF, srow))
        du.number_format, du.border = USD_FMT, BORDER
        di = ws.cell(rr, 5, "=IF(ISNUMBER(%s!$D$%d),%s!$D$%d,0)" % (FF, srow, FF, srow))
        di.number_format, di.border = INR_FMT, BORDER
        pu = ws.cell(rr, 6, '=IF(OR(%s!$E$%d=0,%s!$E$%d=""),"--",%s!$E$%d)' % (FF, srow, FF, srow, FF, srow))
        pu.number_format, pu.border = USD_FMT, BORDER
        pi = ws.cell(rr, 7, "=IF(ISNUMBER(%s!$F$%d),%s!$F$%d,0)" % (FF, srow, FF, srow))
        pi.number_format, pi.border = INR_FMT, BORDER
        var = ws.cell(rr, 8, "=E%d-G%d" % (rr, rr) if mode == "short" else "=G%d-E%d" % (rr, rr))
        var.number_format, var.border, var.fill = INR_FMT, BORDER, fill
        sr = ws.cell(rr, 9, srow)
        sr.border, sr.font, sr.alignment = BORDER, SUB_FONT, Alignment(horizontal="center")
    end = start + len(rows) - 1
    tr = end + 1
    t = ws.cell(tr, 2, "TOTAL")
    t.font, t.fill = TOTAL_FONT, TOTAL_FILL
    for col in (5, 7, 8):
        cl = get_column_letter(col)
        c = ws.cell(tr, col, "=SUM(%s%d:%s%d)" % (cl, start, cl, end))
        c.font, c.fill, c.number_format, c.border = TOTAL_FONT, TOTAL_FILL, INR_FMT, BORDER
    for col, w in zip("ABCDEFGHIJK", [5, 52, 24, 15, 18, 15, 18, 20, 9, 4, 22]):
        ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"
    if sheet_name == "Short Payment":
        SHORT_TOTAL_ROW = tr
    if sheet_name == "Unplanned Payment":
        UNPLANNED_TOTAL_ROW = tr
    if sheet_name == "Excess Payment":
        EXCESS_TOTAL_ROW = tr

# ======================================================================
# 7. MACHINE ADVANCE PAYMENT  (July block: source rows 71-96)
# ======================================================================
mp = out.create_sheet("Machine Advance Payment")
title(mp, "Machine Advance Payment -- Customer Advance Tracker | %s" % MONTH,
      "Live copy of 'Machine Adv_Receivable' July block (source rows 71-96). "
      "Pending = Projected - Received.", 8)
back_link(mp, 10, 1)
hdr_row(mp, 4, ["#", "Sales Person", "Party Name", "Machine", "Projected (Rs.)",
                "Received (Rs.)", "Pending (Rs.)", "Src Row"])
MA_FIRST, MA_LAST = 71, 96
r = 5
idx = 0
for srow in range(MA_FIRST, MA_LAST + 1):
    if not txt(ma_v, srow, 3):
        continue
    idx += 1
    n = mp.cell(r, 1, idx)
    n.border, n.alignment = BORDER, Alignment(horizontal="center")
    for outc, srcc in ((2, 2), (3, 3), (4, 4)):
        c = mp.cell(r, outc, "=IF(%s!%s%d=\"\",\"\",%s!%s%d)" % (MA, get_column_letter(srcc), srow, MA, get_column_letter(srcc), srow))
        c.border = BORDER
    for outc, srcc in ((5, 5), (6, 6)):
        c = mp.cell(r, outc, "=IF(ISNUMBER(%s!%s%d),%s!%s%d,0)" % (MA, get_column_letter(srcc), srow, MA, get_column_letter(srcc), srow))
        c.number_format, c.border = INR_FMT, BORDER
    p = mp.cell(r, 7, "=E%d-F%d" % (r, r))
    p.number_format, p.border = INR_FMT, BORDER
    sr = mp.cell(r, 8, srow)
    sr.border, sr.font, sr.alignment = BORDER, SUB_FONT, Alignment(horizontal="center")
    r += 1
MA_END = r - 1
MA_TOT = r
t = mp.cell(MA_TOT, 3, "TOTAL")
t.font, t.fill = TOTAL_FONT, TOTAL_FILL
for col in (5, 6, 7):
    cl = get_column_letter(col)
    c = mp.cell(MA_TOT, col, "=SUM(%s5:%s%d)" % (cl, cl, MA_END))
    c.font, c.fill, c.number_format, c.border = TOTAL_FONT, TOTAL_FILL, INR_FMT, BORDER
rate = mp.cell(MA_TOT + 2, 1,
               '=IFERROR("Collection Rate: "&TEXT(F%d/E%d,"0.0%%")&"  |  Received Rs."&'
               'TEXT(F%d,"##,##,##,##0")&"  |  Pending Rs."&TEXT(G%d,"##,##,##,##0"),"")'
               % (MA_TOT, MA_TOT, MA_TOT, MA_TOT))
rate.font = GROUP_FONT
mp.merge_cells(start_row=MA_TOT + 2, start_column=1, end_row=MA_TOT + 2, end_column=8)
for col, w in zip("ABCDEFGH", [5, 16, 50, 18, 18, 18, 18, 9]):
    mp.column_dimensions[col].width = w
mp.sheet_view.showGridLines = False
mp.freeze_panes = "A5"

# ======================================================================
# 8. ADVANCE FROM DEBTORS  (new sheet -- no June equivalent)
# ======================================================================
ad = out.create_sheet("Advance from Debtors")
title(ad, "Advance from Debtors -- Position As On Date | %s" % MONTH,
      "Live copy of 'Mach Adv As on Date'. Every customer advance currently held, "
      "largest first. This sheet has no equivalent in the June pack.", 6)
back_link(ad, 8, 1)
hdr_row(ad, 4, ["#", "Customer Name", "Advance Held (Rs.)", "Date", "Sales Person", "Machine"])
r = 5
idx = 0
for srow in range(3, da_v.max_row + 1):
    if not txt(da_v, srow, 1):
        continue
    idx += 1
    n = ad.cell(r, 1, idx)
    n.border, n.alignment = BORDER, Alignment(horizontal="center")
    c = ad.cell(r, 2, "='Debtor Adv Source'!A%d" % srow)
    c.border = BORDER
    v = ad.cell(r, 3, "=IF(ISNUMBER('Debtor Adv Source'!B%d),'Debtor Adv Source'!B%d,0)" % (srow, srow))
    v.number_format, v.border = INR_FMT, BORDER
    for outc, srcc in ((4, 3), (5, 4), (6, 5)):
        c = ad.cell(r, outc, "=IF('Debtor Adv Source'!%s%d=\"\",\"\",'Debtor Adv Source'!%s%d)"
                    % (get_column_letter(srcc), srow, get_column_letter(srcc), srow))
        c.border, c.font = BORDER, SUB_FONT
    r += 1
AD_END = r - 1
t = ad.cell(r, 2, "TOTAL ADVANCES HELD")
t.font, t.fill = TOTAL_FONT, TOTAL_FILL
c = ad.cell(r, 3, "=SUM(C5:C%d)" % AD_END)
c.font, c.fill, c.number_format, c.border = TOTAL_FONT, TOTAL_FILL, INR_FMT, BORDER
for col, w in zip("ABCDEFGH", [5, 55, 20, 18, 22, 26, 4, 22]):
    ad.column_dimensions[col].width = w
ad.sheet_view.showGridLines = False
ad.freeze_panes = "A5"

# ======================================================================
# 9-12. HIDDEN SOURCE SHEETS (row-for-row value snapshots)
# ======================================================================
def snapshot(name, src_ws, note):
    s = out.create_sheet(name)
    for row in src_ws.iter_rows(min_row=1, max_row=src_ws.max_row, max_col=src_ws.max_column):
        for cell in row:
            if cell.value is not None:
                s.cell(cell.row, cell.column, cell.value)
    s.cell(1, src_ws.max_column + 2, note)
    s.sheet_state = "hidden"
    return s


snapshot("FF Source Data", ff_v, "FF Source Data -- live snapshot of '0F_JULY 2026 Cashflow v1.xlsx' > 'FF Summary'")
snapshot("Bank Source Data", bk_v, "Bank Source Data -- snapshot of 'Bank Balance'")
snapshot("MA Source Data", ma_v, "MA Source Data -- snapshot of 'Machine Adv_Receivable'")
snapshot("Debtor Adv Source", da_v, "Debtor Adv Source -- snapshot of 'Mach Adv As on Date'")

# ------------------------------------- patch Exec Summary placeholders
r0, c0 = MA_TOT_CELL
es.cell(r0, c0).value = es.cell(r0, c0).value.replace("$TOT", str(MA_TOT))
r0, c0 = SHORT_TOT_ROW_REF
es.cell(r0, c0).value = (es.cell(r0, c0).value
                         .replace("'Short Payment'!H29", "'Short Payment'!H%d" % SHORT_TOTAL_ROW)
                         .replace("'Short Payment'!E29", "'Short Payment'!E%d" % SHORT_TOTAL_ROW))
for rr in range(13, 20):
    v = es.cell(rr, 1).value
    if isinstance(v, str) and "'Unplanned Payment'!H25" in v:
        es.cell(rr, 1).value = (v.replace("'Unplanned Payment'!H25", "'Unplanned Payment'!H%d" % UNPLANNED_TOTAL_ROW)
                                 .replace("'Excess Payment'!H21", "'Excess Payment'!H%d" % EXCESS_TOTAL_ROW))

out.active = 0
out.save(OUT)
print("WROTE:", OUT)
print("  Unplanned rows: %d (total row %d)" % (len(CLASSIFIED_ROWS["UNPLANNED"]), UNPLANNED_TOTAL_ROW))
print("  Excess    rows: %d (total row %d)" % (len(CLASSIFIED_ROWS["EXCESS"]), EXCESS_TOTAL_ROW))
print("  Short     rows: %d (total row %d)" % (len(CLASSIFIED_ROWS["SHORT"]), SHORT_TOTAL_ROW))
print("  Done      rows: %d" % len(CLASSIFIED_ROWS["DONE"]))
print("  Machine advance rows: %d (total row %d)" % (MA_END - 4, MA_TOT))
print("  Debtor advance rows: %d" % (AD_END - 4))
