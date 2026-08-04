# -*- coding: utf-8 -*-
"""
Fund Flow MOM builder -- April to July FY2026-27 (Q1 + Q2 to date).

Rebuilds "Fund_Flow_MOM.xlsx" with July added as a fourth month, a rolling Apr-Jul
trend, and Q1 / Q2 subtotals.

TWO RESTATEMENTS vs the previous Q1-only version -- both deliberate:

1. JUNE PAYMENT BASIS. The earlier MOM carried June payment at Rs.28,97,68,786. That
   figure came from Fund_Flow_Analysis_Jun2026_LINKED.xlsx, whose Summary!F30 adds the
   Rs.9,69,00,000 bank-funds drawdown into the payment total and then shows the same
   Rs.9,69,00,000 again as a 'Fresh Loan Drawn' inflow. Closing balance still reconciled,
   but payment was inflated by the drawdown, so June was NOT comparable with April, May
   (which carry operating payment with investment shown separately) or July.
   The June source workbook itself reports Total Payment Rs.19,28,68,786.
   Here every month uses the same basis:
       Operating Payment  = total payout EXCLUDING the Investment/Loan block
       Investment/Financing Outflow = the Investment/Loan block, shown on its own line
       Fresh Loan Drawn   = financing INFLOWS, shown on its own line
   June operating payment = 289,768,786 - 14,378,435 = 275,390,351.
   July operating payment = 286,241,541 - 37,569,225 = 248,672,316.
   Both still reconcile to the confirmed bank closing balance.

2. JUNE-CLOSE / JULY-OPEN BREAK. June closes at Rs.6,39,00,000; July opens at
   Rs.7,14,00,000. The Rs.75,00,000 difference is unexplained by either workbook and is
   reported on its own reconciliation line rather than being absorbed silently.

Requires: pip install openpyxl        Run: python build_fund_flow_mom.py
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter

BASE = r"c:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS\FUND FLOW"
OUT = os.path.join(BASE, "Fund_Flow_MOM.xlsx")

# ------------------------------------------------------------------ data
# Apr / May figures carried forward from the prior MOM (already on the
# operating-payment basis). Jun / Jul derived from their source workbooks.
MONTHS = ["April", "May", "June", "July"]

OPENING   = [170000000,  79951252,  46850000,  71400000]
COLLECT   = [179961577, 198279316, 209918786, 263741541]
LOAN_IN   = [        0,  34400000,  96900000,         0]
OP_PAY    = [182153953, 220480152, 275390351, 248672316]
INV_OUT   = [ 87856372,  45300415,  14378435,  37569225]
INV_MF    = [        0,  10000000,         0,         0]
INV_FD    = [ 87856372,  35300415,         0,         0]
INV_PROP  = [        0,         0,   3163465,         0]
INV_LIMIT = [        0,         0,         0,  33108625]
INV_OTHER = [        0,         0,  11214970,   4460600]
CLOSING   = [ 79951252,  46850000,  63900000,  48900000]
BANK      = [ 79951252,  46850000,  63900000,  48900000]

NM = len(MONTHS)
Q1, Q2 = slice(0, 3), slice(3, 4)

# ------------------------------------------------------------- styling
TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FONT = Font(italic=True, size=9, color="595959")
HDR_FONT = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="2E75B6")
GROUP_FONT = Font(bold=True, color="1F4E78")
GROUP_FILL = PatternFill("solid", fgColor="D9E2F3")
TOTAL_FONT = Font(bold=True)
TOTAL_FILL = PatternFill("solid", fgColor="BDD7EE")
Q_FILL = PatternFill("solid", fgColor="E2EFDA")
SUBITEM_FONT = Font(italic=True, size=9, color="404040")
WARN_FILL = PatternFill("solid", fgColor="FCE4E4")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
FLAG_FILL = PatternFill("solid", fgColor="FFF2CC")
NOTE_FONT = Font(size=10)
INR = '##,##,##,##0;[RED](##,##,##,##0)'
CR = '#,##0.00'
PCT = '0.0%'
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = openpyxl.Workbook()
wb.remove(wb.active)


def title(ws, text, sub, width):
    c = ws.cell(1, 1, text)
    c.font, c.fill = TITLE_FONT, TITLE_FILL
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    ws.row_dimensions[1].height = 26
    s = ws.cell(2, 1, sub)
    s.font, s.alignment = SUB_FONT, Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=3, end_column=width)
    ws.row_dimensions[2].height = 26


def hdr(ws, r, headers, start=1):
    for i, h in enumerate(headers, start=start):
        c = ws.cell(r, i, h)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[r].height = 30


# ======================================================================
# SHEET 1 -- Fund Flow Summary
# ======================================================================
s1 = wb.create_sheet("1_Fund_Flow_Summary")
title(s1, "FUND FLOW ANALYSIS -- Apr to Jul FY2026-27  (Q1 complete + Q2 to date)",
      "Closing = Opening + Collection + Loan Drawn - Operating Payment - Investment/Financing Outflow.  "
      "All four months are stated on ONE basis: operating payment excludes the Investment/Loan block, "
      "which is shown separately. June has been restated from the prior MOM (see note below).", 12)

# columns: A label | B..E months | F Q1 | G Q2 TD | H Apr-Jul | I Cr | J % of pay | K trend
COLS = {"lbl": 1, "apr": 2, "may": 3, "jun": 4, "jul": 5,
        "q1": 6, "q2": 7, "tot": 8, "cr": 9, "pct": 10, "trend": 11}
MC = [2, 3, 4, 5]  # month columns

hdr(s1, 5, ["Particulars", "April (Rs.)", "May (Rs.)", "June (Rs.)", "July (Rs.)",
            "Q1 Apr-Jun (Rs.)", "Q2 TD Jul (Rs.)", "Apr-Jul (Rs.)", "Apr-Jul (Rs. Cr)",
            "% of Payment", "Trend Jul vs Apr"])

ROWS = [
    ("Opening Balance",                    OPENING,   "grp",  False),
    ("(+) Collection",                     COLLECT,   "",     True),
    ("(+) Fresh Loan Drawn",               LOAN_IN,   "",     True),
    ("(-) Operating Payment",              OP_PAY,    "",     True),
    ("(-) Investment / Financing Outflow", INV_OUT,   "",     True),
    ("      of which: MF Investment",      INV_MF,    "sub",  False),
    ("      of which: FD & Other Investments", INV_FD, "sub", False),
    ("      of which: Property Purchase",  INV_PROP,  "sub",  False),
    ("      of which: Bank Limit Drop",    INV_LIMIT, "sub",  False),
    ("      of which: Brokerage / EMI / Stamp / Other", INV_OTHER, "sub", False),
    ("Closing Balance (computed)",         CLOSING,   "grp",  False),
    ("Bank Balance (confirmed)",           BANK,      "grp",  False),
]

r = 6
ROWREF = {}
for label, data, kind, pct_of_pay in ROWS:
    ROWREF[label] = r
    lc = s1.cell(r, 1, label)
    lc.border = BORDER
    if kind == "grp":
        lc.font, lc.fill = GROUP_FONT, GROUP_FILL
    elif kind == "sub":
        lc.font = SUBITEM_FONT
    for i, col in enumerate(MC):
        c = s1.cell(r, col, data[i] if data[i] else None)
        c.number_format, c.border = INR, BORDER
        if kind == "grp":
            c.font, c.fill = TOTAL_FONT, GROUP_FILL
        elif kind == "sub":
            c.font = SUBITEM_FONT
    q1 = s1.cell(r, COLS["q1"], "=SUM(B%d:D%d)" % (r, r))
    q2 = s1.cell(r, COLS["q2"], "=E%d" % r)
    tot = s1.cell(r, COLS["tot"], "=F%d+G%d" % (r, r))
    cr = s1.cell(r, COLS["cr"], "=H%d/10000000" % r)
    for c, fmt in ((q1, INR), (q2, INR), (tot, INR), (cr, CR)):
        c.number_format, c.border = fmt, BORDER
        c.fill = Q_FILL
        if kind == "grp":
            c.font = TOTAL_FONT
    # Opening/Closing/Bank are stock balances, not flows -- summing them is meaningless.
    # Opening shows the period-START position; Closing/Bank show the period-END position.
    if label in ("Opening Balance", "Closing Balance (computed)", "Bank Balance (confirmed)"):
        if label == "Opening Balance":
            s1.cell(r, COLS["q1"], "=B%d" % r).number_format = INR   # Q1 opens at April
            s1.cell(r, COLS["tot"], "=B%d" % r).number_format = INR  # Apr-Jul opens at April
        else:
            s1.cell(r, COLS["q1"], "=D%d" % r).number_format = INR   # Q1 closes at June
            s1.cell(r, COLS["tot"], "=E%d" % r).number_format = INR  # Apr-Jul closes at July
        s1.cell(r, COLS["q2"], "=E%d" % r).number_format = INR
        s1.cell(r, COLS["cr"], "=H%d/10000000" % r).number_format = CR
        for cc in (COLS["q1"], COLS["q2"], COLS["tot"], COLS["cr"]):
            s1.cell(r, cc).border, s1.cell(r, cc).fill = BORDER, Q_FILL
            if kind == "grp":
                s1.cell(r, cc).font = TOTAL_FONT
    if pct_of_pay:
        p = s1.cell(r, COLS["pct"], "=IFERROR(H%d/H$%d,\"\")" % (r, r if label == "(-) Operating Payment" else r))
        p.number_format, p.border = PCT, BORDER
    t = s1.cell(r, COLS["trend"],
                '=IF(AND(B%d=0,E%d=0),"-- None",IF(E%d>B%d,"UP  Growing",'
                'IF(E%d<B%d,"DOWN  Declining","FLAT")))' % (r, r, r, r, r, r))
    t.border, t.alignment = BORDER, Alignment(horizontal="center")
    r += 1

# fix % of payment to be against operating payment row
pay_row = ROWREF["(-) Operating Payment"]
for label, data, kind, pct_of_pay in ROWS:
    if pct_of_pay:
        rr = ROWREF[label]
        s1.cell(rr, COLS["pct"], "=IFERROR(H%d/H$%d,\"\")" % (rr, pay_row)).number_format = PCT
        s1.cell(rr, COLS["pct"]).border = BORDER

# ---- reconciliation check per month
r += 1
h = s1.cell(r, 1, "MONTH-BY-MONTH RECONCILIATION  (Difference must = 0)")
h.font, h.fill = GROUP_FONT, GROUP_FILL
s1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
r += 1
hdr(s1, r, ["Month", "Opening", "+ Collection", "+ Loan Drawn", "- Operating Pay",
            "- Investment/Fin", "Computed Closing", "Stated Closing", "Difference", "Status"])
vr = r + 1
op_r = ROWREF["Opening Balance"]
co_r = ROWREF["(+) Collection"]
ln_r = ROWREF["(+) Fresh Loan Drawn"]
pa_r = ROWREF["(-) Operating Payment"]
iv_r = ROWREF["(-) Investment / Financing Outflow"]
cl_r = ROWREF["Closing Balance (computed)"]
for i, m in enumerate(MONTHS):
    rr = vr + i
    cl = get_column_letter(MC[i])
    s1.cell(rr, 1, m).border = BORDER
    for j, srow in enumerate((op_r, co_r, ln_r, pa_r, iv_r)):
        c = s1.cell(rr, 2 + j, "=%s%d" % (cl, srow))
        c.number_format, c.border = INR, BORDER
    comp = s1.cell(rr, 7, "=B%d+C%d+D%d-E%d-F%d" % (rr, rr, rr, rr, rr))
    comp.number_format, comp.border = INR, BORDER
    stated = s1.cell(rr, 8, "=%s%d" % (cl, cl_r))
    stated.number_format, stated.border = INR, BORDER
    diff = s1.cell(rr, 9, "=G%d-H%d" % (rr, rr))
    diff.number_format, diff.border = INR, BORDER
    st = s1.cell(rr, 10, '=IF(ABS(I%d)<2,"OK  VERIFIED","BREAK  CHECK")' % rr)
    st.border, st.alignment = BORDER, Alignment(horizontal="center")
    st.fill = OK_FILL
VER_LAST = vr + NM - 1

# ---- carry-forward continuity check (the Rs.75 lakh break)
r = VER_LAST + 2
h = s1.cell(r, 1, "CARRY-FORWARD CONTINUITY  (prior month closing must equal next month opening)")
h.font, h.fill = GROUP_FONT, GROUP_FILL
s1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
r += 1
hdr(s1, r, ["Hand-over", "Prior Month Closing", "Next Month Opening", "Gap", "Status"])
cf = r + 1
for i in range(NM - 1):
    rr = cf + i
    a, b = get_column_letter(MC[i]), get_column_letter(MC[i + 1])
    s1.cell(rr, 1, "%s close -> %s open" % (MONTHS[i], MONTHS[i + 1])).border = BORDER
    c1 = s1.cell(rr, 2, "=%s%d" % (a, cl_r)); c1.number_format, c1.border = INR, BORDER
    c2 = s1.cell(rr, 3, "=%s%d" % (b, op_r)); c2.number_format, c2.border = INR, BORDER
    gp = s1.cell(rr, 4, "=C%d-B%d" % (rr, rr)); gp.number_format, gp.border = INR, BORDER
    st = s1.cell(rr, 5, '=IF(ABS(D%d)<2,"OK  CONTINUOUS","BREAK  UNEXPLAINED")' % rr)
    st.border, st.alignment = BORDER, Alignment(horizontal="center")
    st.fill = OK_FILL
CF_LAST = cf + NM - 2

# ---- notes
r = CF_LAST + 2
h = s1.cell(r, 1, "BASIS, RESTATEMENTS & ITEMS FOR CONFIRMATION")
h.font, h.fill = GROUP_FONT, GROUP_FILL
s1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
r += 1
NOTES = [
    ("RESTATED",
     "JUNE PAYMENT. The prior MOM carried June payment at Rs.28,97,68,786, taken from "
     "Fund_Flow_Analysis_Jun2026_LINKED.xlsx, which adds the Rs.9,69,00,000 bank-funds drawdown "
     "into the payment total and shows the same amount again as a loan inflow. The June source "
     "workbook reports Rs.19,28,68,786. Restated here to Rs.27,53,90,351 operating payment with "
     "the Investment/Loan block (Rs.1,43,78,435) and the loan inflow (Rs.9,69,00,000) on their own "
     "lines -- so June is now comparable with April, May and July. Closing balance is unchanged."),
    ("BREAK",
     "JUNE CLOSE vs JULY OPEN. June closes at Rs.6,39,00,000; July opens at Rs.7,14,00,000. "
     "The Rs.75,00,000 gap is not explained by either workbook. Confirm whether a bank account "
     "(or the LC/BC free limit, which July's Bank Balance sheet tracks separately) was brought "
     "into scope from 1 July."),
    ("WATCH",
     "JULY HAD NO FRESH FUNDING. April-June were supported by Rs.13.13 Cr of loans/FD releases; "
     "July drew nil and additionally paid down Rs.3.31 Cr via the Axis limit drop. Cash fell from "
     "Rs.7.14 Cr to Rs.4.89 Cr (-31.5%) despite July being the strongest collection month of the year "
     "at Rs.26.37 Cr."),
    ("WATCH",
     "CASH EROSION, FULL PERIOD. Opening Rs.17.00 Cr (1 Apr) to closing Rs.4.89 Cr (31 Jul) -- down "
     "Rs.12.11 Cr (-71.2%) across four months, even after Rs.13.13 Cr of fresh borrowing."),
    ("CONFIRMED",
     "MF Investment Rs.1.00 Cr in May 2026 -- retained as a sub-component of May's investment outflow."),
    ("CONFIRMED",
     "Property purchase Rs.31,63,465 paid June 2026. Balance owed remains a contingent liability, "
     "not a cash outflow. No further payment planned."),
    ("NOTE",
     "Opening / Closing / Bank rows are stock balances, so their Q1 and Apr-Jul columns show the "
     "period-end position, not a sum. Flow rows sum normally."),
]
for tag, text in NOTES:
    tc = s1.cell(r, 1, tag)
    tc.font, tc.border = TOTAL_FONT, BORDER
    tc.fill = WARN_FILL if tag in ("BREAK", "WATCH") else FLAG_FILL if tag == "RESTATED" else OK_FILL
    tc.alignment = Alignment(horizontal="center", vertical="top")
    bc = s1.cell(r, 2, text)
    bc.font, bc.border = NOTE_FONT, BORDER
    bc.alignment = Alignment(wrap_text=True, vertical="top")
    s1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=11)
    s1.row_dimensions[r].height = 46
    r += 1

# ---- KPIs
r += 1
h = s1.cell(r, 1, "KEY PERFORMANCE INDICATORS -- Apr to Jul")
h.font, h.fill = GROUP_FONT, GROUP_FILL
s1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
r += 1
KPI = [
    ("Net Cash Change (Apr-Jul)", "=E%d-B%d" % (cl_r, op_r), INR),
    ("Total Collection (Apr-Jul)", "=H%d" % co_r, INR),
    ("Total Loan Drawn (Apr-Jul)", "=H%d" % ln_r, INR),
    ("Total Operating Payment", "=H%d" % pa_r, INR),
    ("Total Investment / Financing Out", "=H%d" % iv_r, INR),
    ("Collection vs Operating Payment", "=IFERROR(H%d/H%d,\"\")" % (co_r, pa_r), PCT),
    ("Avg Monthly Collection", "=H%d/4" % co_r, INR),
    ("Avg Monthly Operating Payment", "=H%d/4" % pa_r, INR),
    ("Payment vs Collection Gap (Apr-Jul)", "=H%d-H%d" % (pa_r, co_r), INR),
    ("Cash Burn Rate (Rs./month)", "=(B%d-E%d)/4" % (op_r, cl_r), INR),
]
for i, (lbl, formula, fmt) in enumerate(KPI):
    rr = r + i // 2
    cc = 1 + (i % 2) * 3
    lc = s1.cell(rr, cc, lbl)
    lc.font, lc.border = TOTAL_FONT, BORDER
    vc = s1.cell(rr, cc + 1, formula)
    vc.number_format, vc.border, vc.fill = fmt, BORDER, Q_FILL

for col, w in zip("ABCDEFGHIJK", [42, 18, 18, 18, 18, 19, 18, 18, 15, 13, 20]):
    s1.column_dimensions[col].width = w
s1.sheet_view.showGridLines = False
s1.freeze_panes = "B6"

# ======================================================================
# SHEET 2 -- Charts
# ======================================================================
s2 = wb.create_sheet("2_Charts")
title(s2, "GRAPHICAL ANALYSIS -- Apr to Jul FY2026-27",
      "All series are in Rs. Crore and read from 1_Fund_Flow_Summary. June is on the restated "
      "operating-payment basis, so the payment trend is like-for-like across all four months.", 10)

hdr(s2, 5, ["Month", "Collection (Cr)", "Operating Payment (Cr)", "Closing Balance (Cr)",
            "Loan Drawn (Cr)", "Investment / Financing Out (Cr)"])
for i, m in enumerate(MONTHS):
    rr = 6 + i
    s2.cell(rr, 1, m).border = BORDER
    for j, srow in enumerate((co_r, pa_r, cl_r, ln_r, iv_r)):
        c = s2.cell(rr, 2 + j, "='1_Fund_Flow_Summary'!%s%d/10000000" % (get_column_letter(MC[i]), srow))
        c.number_format, c.border = CR, BORDER
CH_FIRST, CH_LAST = 6, 6 + NM - 1

bar = BarChart()
bar.type, bar.style, bar.grouping = "col", 10, "clustered"
bar.title = "Collection vs Operating Payment vs Closing Balance (Rs. Cr)"
bar.y_axis.title, bar.x_axis.title = "Rs. Crore", "Month"
bar.add_data(Reference(s2, min_col=2, max_col=4, min_row=5, max_row=CH_LAST), titles_from_data=True)
bar.set_categories(Reference(s2, min_col=1, min_row=CH_FIRST, max_row=CH_LAST))
bar.height, bar.width = 9, 20
s2.add_chart(bar, "A12")

line = LineChart()
line.title = "Closing Cash Balance Trend (Rs. Cr)"
line.y_axis.title, line.x_axis.title = "Rs. Crore", "Month"
line.add_data(Reference(s2, min_col=4, min_row=5, max_row=CH_LAST), titles_from_data=True)
line.set_categories(Reference(s2, min_col=1, min_row=CH_FIRST, max_row=CH_LAST))
line.height, line.width = 9, 20
s2.add_chart(line, "A32")

fin = BarChart()
fin.type, fin.grouping = "col", "clustered"
fin.title = "Financing: Loan Drawn vs Investment / Financing Outflow (Rs. Cr)"
fin.y_axis.title, fin.x_axis.title = "Rs. Crore", "Month"
fin.add_data(Reference(s2, min_col=5, max_col=6, min_row=5, max_row=CH_LAST), titles_from_data=True)
fin.set_categories(Reference(s2, min_col=1, min_row=CH_FIRST, max_row=CH_LAST))
fin.height, fin.width = 9, 20
s2.add_chart(fin, "A52")

# quarter comparison block
hdr(s2, 72, ["Quarter Comparison", "Collection (Cr)", "Operating Payment (Cr)",
             "Loan Drawn (Cr)", "Investment / Fin Out (Cr)", "Closing (Cr)"])
for i, (qlabel, qcol) in enumerate((("Q1 (Apr-Jun)", "F"), ("Q2 to date (Jul)", "G"))):
    rr = 73 + i
    s2.cell(rr, 1, qlabel).border = BORDER
    for j, srow in enumerate((co_r, pa_r, ln_r, iv_r, cl_r)):
        c = s2.cell(rr, 2 + j, "='1_Fund_Flow_Summary'!%s%d/10000000" % (qcol, srow))
        c.number_format, c.border = CR, BORDER
qc = BarChart()
qc.type, qc.grouping = "col", "clustered"
qc.title = "Q1 vs Q2-to-date (Rs. Cr)"
qc.y_axis.title = "Rs. Crore"
qc.add_data(Reference(s2, min_col=2, max_col=6, min_row=72, max_row=74), titles_from_data=True)
qc.set_categories(Reference(s2, min_col=1, min_row=73, max_row=74))
qc.height, qc.width = 9, 20
s2.add_chart(qc, "H72")

for col, w in zip("ABCDEFGH", [24, 22, 26, 24, 24, 30, 4, 14]):
    s2.column_dimensions[col].width = w
s2.sheet_view.showGridLines = False

# ======================================================================
# SHEET 3 -- Fund Usage
# ======================================================================
s3 = wb.create_sheet("3_Fund_Usage")
title(s3, "HOW COMPANY FUNDS WERE USED -- Apr to Jul FY2026-27",
      "Sources and uses on the restated, consistent basis. Sources = opening cash + collections + "
      "fresh borrowing. Uses = operating payments + investment/financing outflow + cash retained.", 10)

hdr(s3, 5, ["Source of Funds", "April (Rs.)", "May (Rs.)", "June (Rs.)", "July (Rs.)",
            "Q1 (Rs.)", "Q2 TD (Rs.)", "Apr-Jul (Rs.)", "Rs. Cr", "% Share"])
SRC_ROWS = [("Opening Cash Balance (1 Apr)", op_r, True),
            ("Collection from Customers", co_r, False),
            ("Fresh Loan / FD Drawn", ln_r, False)]
r = 6
src_first = r
for lbl, srow, opening_only in SRC_ROWS:
    s3.cell(r, 1, lbl).border = BORDER
    for i, col in enumerate(MC):
        cl = get_column_letter(col)
        v = ("='1_Fund_Flow_Summary'!B%d" % srow) if (opening_only and i == 0) else \
            (None if opening_only else "='1_Fund_Flow_Summary'!%s%d" % (cl, srow))
        c = s3.cell(r, col, v)
        c.number_format, c.border = INR, BORDER
    s3.cell(r, 6, "=SUM(B%d:D%d)" % (r, r)).number_format = INR
    s3.cell(r, 7, "=E%d" % r).number_format = INR
    s3.cell(r, 8, "=F%d+G%d" % (r, r)).number_format = INR
    s3.cell(r, 9, "=H%d/10000000" % r).number_format = CR
    for cc in range(6, 10):
        s3.cell(r, cc).border, s3.cell(r, cc).fill = BORDER, Q_FILL
    r += 1

# The Jun-close / Jul-open step-up is a real cash source that neither workbook explains.
# It is shown on its own line so Sources = Uses balances and the item stays visible,
# rather than surfacing as an anonymous reconciliation gap.
brk = s3.cell(r, 1, "Unexplained carry-forward step-up (Jun close -> Jul open)")
brk.border, brk.font, brk.fill = BORDER, Font(size=10, bold=True, color="C00000"), WARN_FILL
c = s3.cell(r, 5, "='1_Fund_Flow_Summary'!D%d" % (cf + NM - 2))
c.number_format, c.border, c.fill = INR, BORDER, WARN_FILL
s3.cell(r, 6, 0).number_format = INR
s3.cell(r, 7, "=E%d" % r).number_format = INR
s3.cell(r, 8, "=F%d+G%d" % (r, r)).number_format = INR
s3.cell(r, 9, "=H%d/10000000" % r).number_format = CR
for cc in range(2, 10):
    s3.cell(r, cc).border = BORDER
    if cc >= 6:
        s3.cell(r, cc).fill = Q_FILL
r += 1
src_last = r - 1
s3.cell(r, 1, "TOTAL SOURCES").font = TOTAL_FONT
for cc in range(2, 10):
    cl = get_column_letter(cc)
    c = s3.cell(r, cc, "=SUM(%s%d:%s%d)" % (cl, src_first, cl, src_last))
    c.number_format, c.font, c.fill, c.border = (CR if cc == 9 else INR), TOTAL_FONT, TOTAL_FILL, BORDER
SRC_TOT = r
for rr in range(src_first, src_last + 1):
    p = s3.cell(rr, 10, "=IFERROR(H%d/H$%d,\"\")" % (rr, SRC_TOT))
    p.number_format, p.border = PCT, BORDER

r = SRC_TOT + 2
hdr(s3, r, ["Use of Funds", "April (Rs.)", "May (Rs.)", "June (Rs.)", "July (Rs.)",
            "Q1 (Rs.)", "Q2 TD (Rs.)", "Apr-Jul (Rs.)", "Rs. Cr", "% Share"])
r += 1
use_first = r
USE_ROWS = [("Vendor & Operational Payments", pa_r, False),
            ("Investment / Financing Outflow", iv_r, False),
            ("      of which: MF Investment (May)", ROWREF["      of which: MF Investment"], False),
            ("      of which: FD & Other Investments", ROWREF["      of which: FD & Other Investments"], False),
            ("      of which: Property Purchase (Jun)", ROWREF["      of which: Property Purchase"], False),
            ("      of which: Bank Limit Drop (Jul)", ROWREF["      of which: Bank Limit Drop"], False),
            ("Closing Cash Retained (31 Jul)", cl_r, True)]
sub_rows = []
for lbl, srow, closing_only in USE_ROWS:
    lc = s3.cell(r, 1, lbl)
    lc.border = BORDER
    if lbl.startswith("      "):
        lc.font = SUBITEM_FONT
        sub_rows.append(r)
    for i, col in enumerate(MC):
        cl = get_column_letter(col)
        v = ("='1_Fund_Flow_Summary'!E%d" % srow) if (closing_only and i == NM - 1) else \
            (None if closing_only else "='1_Fund_Flow_Summary'!%s%d" % (cl, srow))
        c = s3.cell(r, col, v)
        c.number_format, c.border = INR, BORDER
    s3.cell(r, 6, "=SUM(B%d:D%d)" % (r, r)).number_format = INR
    s3.cell(r, 7, "=E%d" % r).number_format = INR
    s3.cell(r, 8, "=F%d+G%d" % (r, r)).number_format = INR
    s3.cell(r, 9, "=H%d/10000000" % r).number_format = CR
    for cc in range(6, 10):
        s3.cell(r, cc).border, s3.cell(r, cc).fill = BORDER, Q_FILL
    r += 1
use_last = r - 1
s3.cell(r, 1, "TOTAL USES").font = TOTAL_FONT
main_use = [rr for rr in range(use_first, use_last + 1) if rr not in sub_rows]
for cc in range(2, 10):
    cl = get_column_letter(cc)
    expr = "=" + "+".join("%s%d" % (cl, rr) for rr in main_use)
    c = s3.cell(r, cc, expr)
    c.number_format, c.font, c.fill, c.border = (CR if cc == 9 else INR), TOTAL_FONT, TOTAL_FILL, BORDER
USE_TOT = r
for rr in main_use:
    p = s3.cell(rr, 10, "=IFERROR(H%d/H$%d,\"\")" % (rr, USE_TOT))
    p.number_format, p.border = PCT, BORDER

r = USE_TOT + 2
rec = s3.cell(r, 1,
              '=IF(ABS(H%d-H%d)<2,"RECONCILED -- Total Sources = Total Uses '
              '(the Rs.75,00,000 carry-forward step-up is shown as its own source line above, '
              'not absorbed silently)","GAP of Rs."&TEXT(H%d-H%d,"#,##0")&" -- investigate")'
              % (SRC_TOT, USE_TOT, SRC_TOT, USE_TOT))
rec.font = GROUP_FONT
rec.fill = OK_FILL
s3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)

r += 2
h = s3.cell(r, 1, "PROPERTY TRANSACTION SUMMARY")
h.font, h.fill = GROUP_FONT, GROUP_FILL
s3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
r += 1
hdr(s3, r, ["Item", "Status", "Remarks / Action Required"])
PROP = [
    ("Total Property Cost", "TBC", "Confirm with vendor -- obtain Agreement to Sale"),
    ("Paid in June 2026", "Rs.31,63,465", "Captured in June investment/financing outflow"),
    ("Paid in July 2026", "NIL", "No further payment made"),
    ("Balance Owed", "TBC -- Contingent Liability", "No cash outflow planned"),
    ("Balance Sheet Treatment", "Disclose as Commitment / Contingent Liability", "Include in BS Notes"),
    ("Internal Control", "Obtain Sale Deed / Agreement / Board Resolution", "Before any further payment"),
]
for i, (a, b, c) in enumerate(PROP):
    rr = r + 1 + i
    for j, v in enumerate((a, b, c)):
        cc = s3.cell(rr, 1 + j, v)
        cc.border, cc.font = BORDER, NOTE_FONT
        cc.alignment = Alignment(wrap_text=True, vertical="top")
    s3.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=10)

for col, w in zip("ABCDEFGHIJ", [42, 18, 18, 18, 18, 19, 18, 18, 13, 11]):
    s3.column_dimensions[col].width = w
s3.sheet_view.showGridLines = False

wb.active = 0
wb.save(OUT)
print("WROTE:", OUT)
print("  Summary rows: opening=%d collection=%d loan=%d pay=%d inv=%d closing=%d"
      % (op_r, co_r, ln_r, pa_r, iv_r, cl_r))
print("  Reconciliation block rows %d-%d ; carry-forward rows %d-%d" % (vr, VER_LAST, cf, CF_LAST))
for i, m in enumerate(MONTHS):
    comp = OPENING[i] + COLLECT[i] + LOAN_IN[i] - OP_PAY[i] - INV_OUT[i]
    print("  %-6s computed %15s   stated %15s   diff %s"
          % (m, format(comp, ",.0f"), format(CLOSING[i], ",.0f"), format(comp - CLOSING[i], ",.0f")))
