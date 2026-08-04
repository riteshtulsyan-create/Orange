# -*- coding: utf-8 -*-
"""
Builds a consolidated Q1 FY26-27 company review + action plan workbook,
synthesizing PL_BS_GL_Detail.xlsx, GP REPORT.xlsx, the two expense workbooks,
and Working_Capital_Review_Jul2026.xlsx.

The figures below were hand-extracted from those source workbooks on 2026-07-16
(read-only, values only). This is a snapshot synthesis, not a live multi-file
formula link — see the "Read Me First" sheet for what that means for refreshes.
"""
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.page import PageMargins

# Derived from this file's location so it survives a change of user profile.
OUT = str(Path(__file__).resolve().parents[4] / "Files" / "Company_Action_Plan_Q1_FY26-27.xlsx")

NAVY = "1F3864"
GREY_GRID = "E1E0D9"
GREY_MUTED = "898781"
GREEN_FILL = "C6EFCE"
GREEN_FONT = "006100"
RED_FILL = "FFC7CE"
RED_FONT = "9C0006"

wb = openpyxl.Workbook()
wb.remove(wb.active)

thin = Side(style="thin", color=GREY_GRID)
box = Border(top=thin, bottom=thin, left=thin, right=thin)


def style_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5, header=0.2, footer=0.2)


def table_header(ws, row, col_start, headers):
    for i, htext in enumerate(headers):
        c = ws.cell(row=row, column=col_start + i, value=htext)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = box


def data_row(ws, row, col_start, values, num_fmts=None, bold=False):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=col_start + i, value=v)
        c.border = box
        if num_fmts and num_fmts[i]:
            c.number_format = num_fmts[i]
            c.alignment = Alignment(horizontal="right")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if bold:
            c.font = Font(bold=True)


def green_red_cf(ws, cell_range, favorable_when_positive=True):
    green = PatternFill("solid", fgColor=GREEN_FILL)
    red = PatternFill("solid", fgColor=RED_FILL)
    green_font = Font(color=GREEN_FONT)
    red_font = Font(color=RED_FONT)
    pos_fill, pos_font = (green, green_font) if favorable_when_positive else (red, red_font)
    neg_fill, neg_font = (red, red_font) if favorable_when_positive else (green, green_font)
    ws.conditional_formatting.add(
        cell_range, CellIsRule(operator="greaterThan", formula=["0"], fill=pos_fill, font=pos_font))
    ws.conditional_formatting.add(
        cell_range, CellIsRule(operator="lessThan", formula=["0"], fill=neg_fill, font=neg_font))


def merge_text(ws, cell_range, text, size=11, bold=False, color=NAVY, align="left", italic=False):
    ws.merge_cells(cell_range)
    top_left = cell_range.split(":")[0]
    c = ws[top_left]
    c.value = text
    c.font = Font(size=size, bold=bold, color=color, italic=italic, name="Calibri")
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    return c


def add_link(ws, row, col, target_sheet, label, target_cell="A1"):
    c = ws.cell(row=row, column=col, value=f"→ {label}")
    c.hyperlink = f"#'{target_sheet}'!{target_cell}"
    c.font = Font(color="0563C1", underline="single", italic=True, size=10.5)
    c.alignment = Alignment(horizontal="left", vertical="center")
    return c


def table_header_spans(ws, row, spans, size=10, height=26):
    for htext, cs, ce in spans:
        if ce > cs:
            ws.merge_cells(start_row=row, start_column=cs, end_row=row, end_column=ce)
        for cc in range(cs, ce + 1):
            cell = ws.cell(row=row, column=cc)
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.border = box
        c = ws.cell(row=row, column=cs, value=htext)
        c.font = Font(bold=True, color="FFFFFF", size=size)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = height


CR = 10_000_000  # 1 crore, for converting raw rupee figures

# =====================================================================
# SOURCE DATA SHEETS  (hand-extracted figures from the 5 source workbooks,
# 2026-07-16 — see "Read Me First" sheet for refresh guidance)
# =====================================================================

# ---------------- Source Data - PL ----------------
ws = wb.create_sheet("Source Data - PL")
style_sheet(ws)
for col, w in {"A": 2, "B": 30, "C": 16}.items():
    ws.column_dimensions[col].width = w
merge_text(ws, "B2:C2", "P&L Summary — Q1 FY26-27 (Apr-Jun 2026), Grand Total (all companies)", size=13, bold=True)
r = 4
table_header(ws, r, 2, ["Line Item", "Rs. Crore"]); r += 1
pl_rows = [
    ("Total Turnover", 62327.5493720000e-3),
    ("Total Cost of Sales", -49.033007246),
    ("Gross Profit", 13.294542126),
    ("Other Income", 53.1250500000e-3),
    ("Total Operating Expenses", -6937.8499500000e-3),
    ("EBITDA", 6.409817226),
    ("Total Finance Cost", -678.1095670000e-3),
    ("PBT (Profit Before Tax)", 5.731707659),
    ("PAT (Net Profit)", 5.731707659),
]
pl_first = r
for label, val in pl_rows:
    data_row(ws, r, 2, [label, val], num_fmts=[None, '0.00" Cr"'])
    r += 1
pl_last = r - 1
r += 1
merge_text(ws, f"B{r}:C{r}", "PBT by Company — Q1 FY26-27 (why Grand Total masks a loss-making entity)", size=11, bold=True)
r += 1
table_header(ws, r, 2, ["Company", "PBT (Rs. Crore)"]); r += 1
pl_company_first = r
company_pbt = [
    ("Orange O Tec Pvt Ltd - Surat", -6372.5742829999e-3),
    ("Orange O Tec Pvt Ltd - Noida", 4586.5817260000e-3),
    ("Orange O Tec Enterprises - Surat", 4.495972436),
    ("Orange O Tec Enterprises - Noida", 1504.2669490000e-3),
    ("Colorix Digital Printing Solutions LLP", 1517.4608310000e-3),
]
for label, val in company_pbt:
    data_row(ws, r, 2, [label, val], num_fmts=[None, '0.00" Cr"'])
    r += 1
pl_company_last = r - 1
green_red_cf(ws, f"C{pl_company_first}:C{pl_company_last}", favorable_when_positive=True)
add_link(ws, pl_company_last + 2, 2, "Executive Summary", "Back to Executive Summary")
ws.print_area = f"A1:C{pl_company_last + 2}"
ws.print_options.horizontalCentered = True

# ---------------- Source Data - BS ----------------
ws = wb.create_sheet("Source Data - BS")
style_sheet(ws)
for col, w in {"A": 2, "B": 30, "C": 16}.items():
    ws.column_dimensions[col].width = w
merge_text(ws, "B2:C2", "Balance Sheet Summary — as of 30-Jun-2026, Grand Total", size=13, bold=True)
r = 4
table_header(ws, r, 2, ["Line Item", "Rs. Crore"]); r += 1
bs_rows = [
    ("Total Liabilities", 172.5983452),
    ("Total Assets", 172.5983452),
    ("Sundry Debtors", 65605.3576860000e-3),
    ("Sundry Creditors", 48.5566170),
    ("Bank OD (Overdraft)", 37366.2236100000e-3),
    ("Secured Loans", 11948.9196280000e-3),
]
bs_first = r
for label, val in bs_rows:
    data_row(ws, r, 2, [label, val], num_fmts=[None, '0.00" Cr"'])
    r += 1
bs_last = r - 1
r += 1
merge_text(ws, f"B{r}:C{r}", "Reconciliation Gap trend (Liabilities vs Assets; ~0 expected) — fully closed as of 17-Jul-2026", size=11, bold=True, color="1E7B34")
r += 1
table_header(ws, r, 2, ["Month-end", "Gap (Rs. Crore)"]); r += 1
gap_first = r
gap_rows = [("Apr-26", 1717.0441080000e-3), ("May-26", 1690.9766060000e-3), ("Jun-26", 0.0)]
for label, val in gap_rows:
    data_row(ws, r, 2, [label, val], num_fmts=[None, '0.00" Cr"'])
    r += 1
gap_last = r - 1
add_link(ws, gap_last + 2, 2, "Executive Summary", "Back to Executive Summary")
ws.print_area = f"A1:C{gap_last + 2}"
ws.print_options.horizontalCentered = True

# ---------------- Source Data - Ratios ----------------
ws = wb.create_sheet("Source Data - Ratios")
style_sheet(ws)
for col, w in {"A": 2, "B": 30, "C": 16}.items():
    ws.column_dimensions[col].width = w
merge_text(ws, "B2:C2", "Key Ratios — Q1 FY26-27, Grand Total", size=13, bold=True)
r = 4
table_header(ws, r, 2, ["Ratio", "Value"]); r += 1
ratio_rows = [
    ("Gross Profit Margin %", 0.213301, "0.0%"),
    ("EBITDA Margin %", 0.102841, "0.0%"),
    ("PBT Margin %", 0.091961, "0.0%"),
    ("Return on Assets %", 0.032826, "0.0%"),
    ("Return on Equity %", 0.112038, "0.0%"),
    ("Current Ratio", 1.5458, "0.00"),
    ("Quick Ratio", 0.9657, "0.00"),
    ("Debt-Equity Ratio", 1.1528, "0.00"),
    ("Interest Coverage Ratio", 9.3524, "0.00"),
    ("Debtor Days", 95.8, "0.0"),
    ("Creditor Days", 90.1, "0.0"),
]
ratio_first = r
for label, val, fmt in ratio_rows:
    data_row(ws, r, 2, [label, val], num_fmts=[None, fmt])
    r += 1
ratio_last = r - 1
r += 1
merge_text(ws, f"B{r}:C{r}", "Orange O Tec Pvt Ltd - Surat (standalone) — the loss-making entity flagged above", size=11, bold=True)
r += 1
table_header(ws, r, 2, ["Ratio", "Value"]); r += 1
surat_first = r
surat_rows = [
    ("Gross Profit Margin %", -0.039011, "0.0%"),
    ("EBITDA Margin %", -0.166439, "0.0%"),
    ("PBT Margin %", -0.185273, "0.0%"),
]
for label, val, fmt in surat_rows:
    data_row(ws, r, 2, [label, val], num_fmts=[None, fmt])
    r += 1
surat_last = r - 1
green_red_cf(ws, f"C{surat_first}:C{surat_last}", favorable_when_positive=True)
add_link(ws, surat_last + 2, 2, "Executive Summary", "Back to Executive Summary")
ws.print_area = f"A1:C{surat_last + 2}"
ws.print_options.horizontalCentered = True

# ---------------- Source Data - GP ----------------
ws = wb.create_sheet("Source Data - GP")
style_sheet(ws)
for col, w in {"A": 2, "B": 20, "C": 16, "D": 16, "E": 14}.items():
    ws.column_dimensions[col].width = w
merge_text(ws, "B2:E2", "Gross Profit Performance — Q1 FY25-26 vs Q1 FY26-27 (Apr-Jun, year-on-year)", size=13, bold=True)
r = 4
table_header(ws, r, 2, ["Metric", "Q1 FY25-26", "Q1 FY26-27", "Change %"]); r += 1
overall_first = r
gp_overall = [
    ("Net Sales (Rs. Cr)", 36887.5107e-3, 35363.9685e-3),
    ("Gross Profit (Rs. Cr)", 10875.0748e-3, 11331.5981e-3),
]
for label, before, now in gp_overall:
    data_row(ws, r, 2, [label, before, now, f"=(D{r}-C{r})/C{r}"], num_fmts=[None, '0.00" Cr"', '0.00" Cr"', "0.0%"])
    r += 1
data_row(ws, r, 2, ["Gross Profit %", 0.2948, 0.3204, "=D%d-C%d" % (r, r)], num_fmts=[None, "0.0%", "0.0%", "+0.0%;-0.0%"])
r += 1
overall_last = r - 1
r += 1
table_header(ws, r, 2, ["Segment", "Sales Q1 25-26", "Sales Q1 26-27", "Sales Chg %"]); r += 1
seg_sales_first = r
gp_segment = [
    ("Ink", 27926.3140e-3, 28429.2775e-3, 0.3253, 0.3736),
    ("Spare Parts", 1921.1708e-3, 1727.8684e-3, 0.3837, 0.2778),
    ("Heads", 7040.0260e-3, 5206.8226e-3, 0.1496, 0.0445),
]
for label, s1, s2, gp1, gp2 in gp_segment:
    data_row(ws, r, 2, [label, s1, s2, f"=(D{r}-C{r})/C{r}"], num_fmts=[None, '0.00" Cr"', '0.00" Cr"', "0.0%"])
    r += 1
seg_sales_last = r - 1
r += 1
table_header(ws, r, 2, ["Segment", "GP% Q1 25-26", "GP% Q1 26-27", "GP% Change"]); r += 1
seg_gp_first = r
for label, s1, s2, gp1, gp2 in gp_segment:
    data_row(ws, r, 2, [label, gp1, gp2, f"=D{r}-C{r}"], num_fmts=[None, "0.0%", "0.0%", "+0.0%;-0.0%"])
    r += 1
seg_gp_last = r - 1
green_red_cf(ws, f"D{overall_first}:D{overall_last}", favorable_when_positive=True)
green_red_cf(ws, f"D{seg_sales_first}:D{seg_sales_last}", favorable_when_positive=True)
green_red_cf(ws, f"D{seg_gp_first}:D{seg_gp_last}", favorable_when_positive=True)
GP_OVERALL_FIRST, GP_OVERALL_LAST = overall_first, overall_last
GP_SEG_SALES_FIRST = seg_sales_first
GP_SEG_GP_FIRST = seg_gp_first
add_link(ws, seg_gp_last + 2, 2, "Executive Summary", "Back to Executive Summary")
ws.print_area = f"A1:E{seg_gp_last + 2}"
ws.print_options.horizontalCentered = True

# ---------------- Source Data - Expenses ----------------
ws = wb.create_sheet("Source Data - Expenses")
style_sheet(ws)
for col, w in {"A": 2, "B": 26, "C": 16, "D": 16, "E": 14}.items():
    ws.column_dimensions[col].width = w
merge_text(ws, "B2:E2", "Expense Comparison — Q1 FY25-26 vs Q1 FY26-27 (Apr-Jun, year-on-year)", size=13, bold=True)
merge_text(ws, "B3:E3", "Source: EXPENSES SHEET FOR MARCH 26 (full FY25-26 data) vs Expense Sheet 2026-07-14 (Q1 FY26-27). Category names differ slightly between the two source files' schemas — matched by closest name.", size=9, italic=True, color=GREY_MUTED)
r = 5
table_header(ws, r, 2, ["Category", "Q1 FY25-26 (Rs. Cr)", "Q1 FY26-27 (Rs. Cr)", "Change %"]); r += 1
exp_total_first = r
exp_totals = [
    ("Direct Expenses (Total)", 2249.1903980000e-3, 4116.8874970000e-3),
    ("Indirect Expenses (Total)", 6873.0678240000e-3, 7501.2916750000e-3),
]
for label, before, now in exp_totals:
    data_row(ws, r, 2, [label, before, now, f"=(D{r}-C{r})/C{r}"], num_fmts=[None, '0.00" Cr"', '0.00" Cr"', "0.0%"])
    r += 1
data_row(ws, r, 2, ["Total Expenses (Direct + Indirect)", "=C%d+C%d" % (exp_total_first, exp_total_first + 1),
                    "=D%d+D%d" % (exp_total_first, exp_total_first + 1), f"=(D{r}-C{r})/C{r}"],
         num_fmts=[None, '0.00" Cr"', '0.00" Cr"', "0.0%"], bold=True)
r += 1
exp_total_last = r - 1
r += 1
table_header(ws, r, 2, ["Top Cost Categories", "Q1 FY25-26 (Rs. Cr)", "Q1 FY26-27 (Rs. Cr)", "Change %"]); r += 1
exp_cat_first = r
exp_categories = [
    ("Employee Cost (Indirect)", 3005.2860450000e-3, 3680.7472620000e-3),
    ("Custom Duty Expenses (Direct)", 1121.9920000000e-3, 1736.1570700000e-3),
    ("Freight Expenses (Direct)", 1036.3473620000e-3, 800.4078540000e-3),
    ("Finance Cost (Indirect)", 1032.7053360000e-3, 1055.7210850000e-3),
    ("Rent Expenses (Indirect)", 282.7230000000e-3, 415.3942920000e-3),
    ("Commission & Brokerage (Indirect)", 471.1520000000e-3, 598.9213000000e-3),
]
for label, before, now in exp_categories:
    data_row(ws, r, 2, [label, before, now, f"=(D{r}-C{r})/C{r}"], num_fmts=[None, '0.00" Cr"', '0.00" Cr"', "0.0%"])
    r += 1
exp_cat_last = r - 1
# For expenses, a DECREASE is favorable (green) and an INCREASE is unfavorable (red)
green_red_cf(ws, f"D{exp_total_first}:D{exp_total_last}", favorable_when_positive=False)
green_red_cf(ws, f"D{exp_cat_first}:D{exp_cat_last}", favorable_when_positive=False)
EXP_TOTAL_FIRST, EXP_TOTAL_LAST = exp_total_first, exp_total_last
EXP_CAT_FIRST, EXP_CAT_LAST = exp_cat_first, exp_cat_last
add_link(ws, exp_cat_last + 2, 2, "Expense Details", "Full category-by-category breakdown with reasons")
add_link(ws, exp_cat_last + 3, 2, "Executive Summary", "Back to Executive Summary")
ws.print_area = f"A1:E{exp_cat_last + 3}"
ws.print_options.horizontalCentered = True

# ---------------- Source Data - Working Capital ----------------
ws = wb.create_sheet("Source Data - Working Capital")
style_sheet(ws)
for col, w in {"A": 2, "B": 20, "C": 16, "D": 16, "E": 14, "F": 30}.items():
    ws.column_dimensions[col].width = w
merge_text(ws, "B2:F2", "Working Capital — 12-Jun-2026 vs 8-Jul-2026 (most recent snapshot)", size=13, bold=True)
r = 4
table_header(ws, r, 2, ["Item", "12-Jun-26 (Rs. Cr)", "8-Jul-26 (Rs. Cr)", "Change"]); r += 1
wc_glance_first = r
wc_glance = [
    ("Inventory", 45.441161369999996, 44.96605338599999),
    ("Debtors", 63.301696026, 68.64373913600001),
    ("Creditors", 66.37738612399998, 62.03340949799999),
    ("Net Working Capital", 42.36547127199999, 51.576383024000016),
]
for label, before, now in wc_glance:
    data_row(ws, r, 2, [label, before, now, f"=D{r}-C{r}"], num_fmts=[None, '0.00" Cr"', '0.00" Cr"', '+0.00" Cr";-0.00" Cr"'], bold=(label == "Net Working Capital"))
    r += 1
wc_glance_last = r - 1
r += 1
table_header(ws, r, 2, ["Segment", "Net WC (Rs. Cr)", "Aged Receivables >180d (Rs. Cr)", "Total Receivables (Rs. Cr)", "% Overdue"]); r += 1
wc_seg_first = r
wc_segment = [
    ("Machines", 24.268304259, 32.253591483, 34.211491482999996),
    ("Ink", 20.46707625200002, 26.501732944, 36.814162884),
    ("Head", 13.045370753, 5.865707522999999, 6.988300522999999),
    ("Spare Parts", 9.77343089499998, 1.690020007, 2.402842845),
    ("Others", -15.977799135, 0.7271831, 3.119704434),
]
for label, netwc, aged, total in wc_segment:
    data_row(ws, r, 2, [label, netwc, aged, total, f"=D{r}/E{r}"], num_fmts=[None, '0.00" Cr"', '0.00" Cr"', '0.00" Cr"', "0%"])
    r += 1
wc_seg_last = r - 1
WC_GLANCE_FIRST, WC_GLANCE_LAST = wc_glance_first, wc_glance_last
WC_SEG_FIRST, WC_SEG_LAST = wc_seg_first, wc_seg_last
add_link(ws, wc_seg_last + 2, 2, "Executive Summary", "Back to Executive Summary")
ws.print_area = f"A1:F{wc_seg_last + 2}"
ws.print_options.horizontalCentered = True

# =====================================================================
# READ ME FIRST
# =====================================================================
ws = wb.create_sheet("Read Me First", 0)
style_sheet(ws)
for col, w in {"A": 2, "B": 110}.items():
    ws.column_dimensions[col].width = w
merge_text(ws, "B2:B2", "Company Review & Action Plan — Q1 FY26-27 (Apr-Jun 2026)", size=18, bold=True)
r = 4
notes = [
    "This workbook pulls together five source files into one place: PL_BS_GL_Detail.xlsx (P&L, Balance Sheet, ratios), "
    "GP REPORT.xlsx (gross profit by segment), the two expense workbooks (Q1 FY26-27 and the FY25-26 file used for "
    "last year's comparison), and Working_Capital_Review_Jul2026.xlsx.",
    "",
    "This workbook has five sheets besides this one: 'Executive Summary' (the short, high-level summary — start "
    "here), 'Action Plan' (the full consolidated action plan), 'Expense Details' (the full Direct and Indirect "
    "expense category breakdown, with reasons), and the 'Source Data - ...' sheets (the underlying figures). "
    "Every table in 'Executive Summary', 'Action Plan', and 'Expense Details' either holds its own figures or is a "
    "formula pointing at a 'Source Data - ...' sheet — nothing is typed twice within this workbook. Blue links at "
    "the top and bottom of each sheet jump straight to the related sheet.",
    "",
    "The Source Data sheets themselves, however, are a snapshot: the figures were read out of the five source files "
    "on 16-Jul-2026 and typed in here as plain numbers, not live links back to those files. If any of those five "
    "files is updated later, this workbook will NOT pick up the change by itself — someone needs to re-extract the "
    "relevant figures and update the Source Data sheets. In particular, once tomorrow's Tally session closes the "
    "Balance Sheet gap noted below, the P&L/BS/Ratio figures in this workbook should be refreshed to match.",
    "",
    "Resolved item: the Balance Sheet previously had a small mismatch between total Liabilities and total Assets "
    "(should be ~0; it ran Rs. 1.72 Cr in April, Rs. 1.69 Cr in May, and Rs. 0.05 Cr in June once Tally was "
    "reconnected on 17-Jul-2026, after a Purchase Bills to Come entry (Enterprises Surat) and an Unadjusted Forex "
    "Gain/Loss entry (Colorix) were booked). A follow-up pass corrected a P&L-side sign error on the same "
    "Enterprises adjustment, raising Grand Total PBT by Rs. 0.068 Cr (now Rs. 5.73 Cr). The remaining Rs. 0.054 Cr "
    "Enterprises residual traced to one creditor balance (Usha Electro Trade Agencies) that had since been cleared "
    "in Tally; once removed, the gap closed fully. Independently re-verified via a full formula recalculation of "
    "the source workbook on 17-Jul-2026: Total Liabilities now equals Total Assets to the rupee for all five "
    "entities. This item is now closed — see 'Source Data - BS'.",
    "",
    "Colour key used throughout: green = favourable movement, red = unfavourable movement. For profit and cash "
    "figures, an increase is green. For expense figures, a DECREASE is green (spending less is good) and an "
    "increase is red.",
]
for n in notes:
    c = ws.cell(row=r, column=2, value=n)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if n:
        est_lines = max(-(-len(n) // 105), 1)
        ws.row_dimensions[r].height = max(15, est_lines * 15)
    r += 1
r += 1
add_link(ws, r, 2, "Executive Summary", "Go to Executive Summary")
ws.print_area = f"A1:B{r}"
ws.print_options.horizontalCentered = True

# =====================================================================
# EXECUTIVE SUMMARY
# =====================================================================
ws = wb.create_sheet("Executive Summary", 1)
style_sheet(ws)
for col, w in {"A": 2, "B": 26, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15}.items():
    ws.column_dimensions[col].width = w
EXEC_LAST_COL = "G"

merge_text(ws, f"B2:{EXEC_LAST_COL}2", "Company Review — Q1 FY26-27 (Apr-Jun 2026)", size=20, bold=True, align="center")
merge_text(ws, f"B3:{EXEC_LAST_COL}3", "Quarter snapshot and consolidated action plan", size=12, color=GREY_MUTED, align="center", italic=True)

add_link(ws, 4, 2, "Read Me First", "Read Me First")
add_link(ws, 4, 4, "Action Plan", "Action Plan")
add_link(ws, 4, 6, "Expense Details", "Expense Details")

r = 5
headline = ws.cell(row=r, column=2,
    value=f"=\"Q1 turnover: Rs. \"&TEXT('Source Data - PL'!C{pl_first},\"0.0\")&\" Cr, Gross Profit Rs. \""
          f"&TEXT('Source Data - PL'!C{pl_first+2},\"0.0\")&\" Cr, PBT Rs. \"&TEXT('Source Data - PL'!C{pl_first+7},\"0.0\")&\" Cr\"")
ws.merge_cells(f"B{r}:{EXEC_LAST_COL}{r}")
headline.font = Font(size=14, bold=True, color=NAVY)
headline.alignment = Alignment(horizontal="center")
r += 2

# ---- Section 1: Company Snapshot ----
merge_text(ws, f"B{r}:E{r}", "1. Company Snapshot (Rs. Crore)", size=13, bold=True)
r += 1
table_header(ws, r, 2, ["Metric", "Q1 FY26-27"])
snap_header = r
r += 1
snap_first = r
snap_rows = [("Turnover", pl_first), ("Gross Profit", pl_first + 2), ("EBITDA", pl_first + 5),
             ("PBT / PAT", pl_first + 7)]
for label, src_row in snap_rows:
    data_row(ws, r, 2, [label, f"='Source Data - PL'!C{src_row}"], num_fmts=[None, '0.00" Cr"'])
    r += 1
snap_last = r - 1
green_red_cf(ws, f"C{snap_first}:C{snap_last}", favorable_when_positive=True)
r += 1
table_header(ws, r, 2, ["Key Ratios", "Q1 FY26-27"])
r += 1
kr_first = r
key_ratios = [("Gross Profit %", ratio_first), ("EBITDA %", ratio_first + 1), ("PBT %", ratio_first + 2),
              ("Current Ratio", ratio_first + 5), ("Debtor Days", ratio_first + 9), ("Creditor Days", ratio_first + 10)]
for label, src_row in key_ratios:
    fmt = "0.0%" if "%" in label else "0.0"
    data_row(ws, r, 2, [label, f"='Source Data - Ratios'!C{src_row}"], num_fmts=[None, fmt])
    r += 1
kr_last = r - 1
r += 1
merge_text(ws, f"B{r}:E{r}", "PBT by Company — the Grand Total hides one loss-making entity", size=11, bold=True)
r += 1
table_header(ws, r, 2, ["Company", "PBT (Rs. Cr)"])
r += 1
pbtc_first = r
for i in range(pl_company_first, pl_company_last + 1):
    data_row(ws, r, 2, [f"='Source Data - PL'!B{i}", f"='Source Data - PL'!C{i}"], num_fmts=[None, '0.00" Cr"'])
    r += 1
pbtc_last = r - 1
green_red_cf(ws, f"C{pbtc_first}:C{pbtc_last}", favorable_when_positive=True)
r += 1
add_link(ws, r, 2, "Source Data - PL", "Full P&L and PBT-by-company detail")
add_link(ws, r, 6, "Source Data - Ratios", "Full ratio detail (incl. Surat standalone)")

section1_end = r
r = section1_end + 2

# ---- Section 2: Gross Profit Performance ----
merge_text(ws, f"B{r}:F{r}", "2. Gross Profit Performance (Q1 FY25-26 vs Q1 FY26-27)", size=13, bold=True)
r += 1
table_header(ws, r, 2, ["Metric", "Q1 FY25-26", "Q1 FY26-27", "Change %"])
r += 1
gp_ov_first = r
for i in range(GP_OVERALL_FIRST, GP_OVERALL_LAST + 1):
    is_pct = i == GP_OVERALL_LAST
    fmt_c = "0.0%" if is_pct else '0.00" Cr"'
    fmt_d = "+0.0%;-0.0%" if is_pct else "0.0%"
    data_row(ws, r, 2, [f"='Source Data - GP'!B{i}", f"='Source Data - GP'!C{i}", f"='Source Data - GP'!D{i}",
                        f"='Source Data - GP'!E{i}"], num_fmts=[None, fmt_c, fmt_c, fmt_d])
    r += 1
gp_ov_last = r - 1
green_red_cf(ws, f"E{gp_ov_first}:E{gp_ov_last}", favorable_when_positive=True)
r += 1
table_header(ws, r, 2, ["Segment", "Sales Q1 25-26", "Sales Q1 26-27", "Sales Chg %", "GP% Q1 25-26", "GP% Q1 26-27"])
r += 1
gp_seg_first = r
seg_labels = ["Ink", "Spare Parts", "Heads"]
for i, label in enumerate(seg_labels):
    sales_row = GP_SEG_SALES_FIRST + i
    gp_row = GP_SEG_GP_FIRST + i
    data_row(ws, r, 2, [label, f"='Source Data - GP'!C{sales_row}", f"='Source Data - GP'!D{sales_row}",
                        f"='Source Data - GP'!E{sales_row}", f"='Source Data - GP'!C{gp_row}",
                        f"='Source Data - GP'!D{gp_row}"],
              num_fmts=[None, '0.00" Cr"', '0.00" Cr"', "0.0%", "0.0%", "0.0%"])
    r += 1
gp_seg_last = r - 1
green_red_cf(ws, f"E{gp_seg_first}:E{gp_seg_last}", favorable_when_positive=True)
r += 1
gp_commentary = [
    "Ink is ~80% of sales and the sole growth engine this quarter, with margin up nearly 5 points.",
    "Spare Parts and Heads both lost sales and margin — Heads especially severely (GP% fell from 15.0% to 4.5%).",
    "Pigment Ink margin has collapsed (30.4% to 9.2% last year-on-year) — a cost or pricing issue worth checking directly.",
]
for c in gp_commentary:
    ws.cell(row=r, column=2, value=f"-  {c}").font = Font(size=10.5)
    ws.merge_cells(f"B{r}:G{r}")
    ws.row_dimensions[r].height = 18
    r += 1
add_link(ws, r, 2, "Source Data - GP", "Full Gross Profit detail (overall, segment, and source figures)")
r += 1

section2_end = r
r = section2_end + 1

# ---- Section 3: Expense Analysis ----
merge_text(ws, f"B{r}:E{r}", "3. Expense Analysis (Q1 FY25-26 vs Q1 FY26-27)", size=13, bold=True)
r += 1
table_header(ws, r, 2, ["Category", "Q1 FY25-26", "Q1 FY26-27", "Change %"])
r += 1
exp_tot_first = r
for i in range(EXP_TOTAL_FIRST, EXP_TOTAL_LAST + 1):
    bold = i == EXP_TOTAL_LAST
    data_row(ws, r, 2, [f"='Source Data - Expenses'!B{i}", f"='Source Data - Expenses'!C{i}",
                        f"='Source Data - Expenses'!D{i}", f"='Source Data - Expenses'!E{i}"],
              num_fmts=[None, '0.00" Cr"', '0.00" Cr"', "0.0%"], bold=bold)
    r += 1
exp_tot_last = r - 1
green_red_cf(ws, f"E{exp_tot_first}:E{exp_tot_last}", favorable_when_positive=False)
r += 1
exp_commentary = [
    "Direct expenses are up 83% year-on-year, well ahead of the ~4% dip in sales seen in the same comparison — "
    "Custom Duty alone is up 55%.",
    "Indirect expenses are up a more modest 9% — Employee Cost, Rent, and Commission & Brokerage are the biggest movers.",
]
for c in exp_commentary:
    ws.cell(row=r, column=2, value=f"-  {c}").font = Font(size=10.5)
    ws.merge_cells(f"B{r}:G{r}")
    ws.row_dimensions[r].height = 18
    r += 1
add_link(ws, r, 2, "Expense Details", "Full category-by-category breakdown, Direct and Indirect, with reasons")
r += 1

section3_end = r
r = section3_end + 1

# ---- Section 4: Where Funds Are Blocked ----
merge_text(ws, f"B{r}:E{r}", "4. Where Funds Are Blocked (Working Capital)", size=13, bold=True)
r += 1
table_header(ws, r, 2, ["Item", "12-Jun-26", "8-Jul-26", "Change"])
r += 1
wcg_first = r
for i in range(WC_GLANCE_FIRST, WC_GLANCE_LAST + 1):
    bold = i == WC_GLANCE_LAST
    data_row(ws, r, 2, [f"='Source Data - Working Capital'!B{i}", f"='Source Data - Working Capital'!C{i}",
                        f"='Source Data - Working Capital'!D{i}", f"='Source Data - Working Capital'!E{i}"],
              num_fmts=[None, '0.00" Cr"', '0.00" Cr"', '+0.00" Cr";-0.00" Cr"'], bold=bold)
    r += 1
wcg_last = r - 1
# Net Working Capital increasing = more cash blocked = unfavourable
green_red_cf(ws, f"E{wcg_last}", favorable_when_positive=False)
r += 1
table_header(ws, r, 2, ["Segment", "Net WC (Rs. Cr)", "Aged Receivables >180d", "Total Receivables", "% Overdue"])
r += 1
wcs_first = r
for i in range(WC_SEG_FIRST, WC_SEG_LAST + 1):
    data_row(ws, r, 2, [f"='Source Data - Working Capital'!B{i}", f"='Source Data - Working Capital'!C{i}",
                        f"='Source Data - Working Capital'!D{i}", f"='Source Data - Working Capital'!E{i}",
                        f"='Source Data - Working Capital'!F{i}"],
              num_fmts=[None, '0.00" Cr"', '0.00" Cr"', '0.00" Cr"', "0%"])
    r += 1
wcs_last = r - 1
# % Overdue: high overdue share is unfavourable (red), low share is favourable (green)
ws.conditional_formatting.add(
    f"F{wcs_first}:F{wcs_last}",
    CellIsRule(operator="greaterThan", formula=["0.7"], fill=PatternFill("solid", fgColor=RED_FILL), font=Font(color=RED_FONT)))
ws.conditional_formatting.add(
    f"F{wcs_first}:F{wcs_last}",
    CellIsRule(operator="lessThan", formula=["0.3"], fill=PatternFill("solid", fgColor=GREEN_FILL), font=Font(color=GREEN_FONT)))
r += 1
wc_commentary = [
    "Net Working Capital has grown Rs. 9.2 Cr (+22%) in just 26 days — cash conversion is getting worse, not better.",
    "Machines is where the most cash is stuck: Rs. 24.3 Cr net working capital, and 94% of its receivables are overdue beyond 180 days — the highest risk of any segment.",
    "Ink carries the largest single working-capital swing this period (+Rs. 5.7 Cr) — driven by both inventory build-up and slow collections (72% of Ink receivables also overdue beyond 180 days).",
    "Debtor Days (95.8) run higher than Creditor Days (90.1) — the company is effectively funding customers for longer than suppliers are funding it, a gap that drains cash every quarter it persists.",
]
for c in wc_commentary:
    ws.cell(row=r, column=2, value=f"-  {c}").font = Font(size=10.5)
    ws.merge_cells(f"B{r}:G{r}")
    ws.row_dimensions[r].height = 18
    r += 1
add_link(ws, r, 2, "Source Data - Working Capital", "Full working capital detail, all segments")
r += 1

section4_end = r
r = section4_end + 1
add_link(ws, r, 2, "Action Plan", "Full consolidated action plan (priority, findings, recommended action, basis)")
ws.row_dimensions[r].height = 18
r += 1

last_row = r + 1
ws.print_area = f"A1:G{last_row}"
ws.print_options.horizontalCentered = True

# ---------------------------------------------------------------
# Consolidated Action Plan items (used to build the "Action Plan" sheet below)
# ---------------------------------------------------------------
action_plan = [
    ("High", "Orange O Tec Pvt Ltd - Surat",
     "This one entity made a Rs. 6.4 Cr loss this quarter (GP margin -3.9%, PBT margin -18.5%) while every other "
     "group entity was profitable.",
     "Get a standalone review of this entity's pricing, cost allocation, and branch/related-party transactions "
     "within 2 weeks — the loss is large enough to be masking a structural problem, not just a bad quarter.",
     "See Source Data - PL (PBT by Company) and Source Data - Ratios (Surat standalone margins)."),
    ("High", "Machines — Receivables",
     "Rs. 32.3 Cr of the Rs. 34.2 Cr owed by Machines customers (94%) is overdue beyond 180 days — the worst "
     "collection quality of any segment, and the single biggest blocked-cash item in the business.",
     "Freeze further dispatch to any Machines customer overdue beyond 180 days; assign named owners to recover at "
     "least 25% of the aged balance (~Rs. 8 Cr) by end of next quarter.",
     "See Source Data - Working Capital (segment aged receivables)."),
    ("High", "Heads segment",
     "Heads gross profit % collapsed from 15.0% to 4.5% year-on-year, on a sales base that also fell 26%.",
     "Review Heads pricing and discount approval limits this month — discounting on this segment has nearly "
     "doubled year-on-year per the GP report; also confirm whether the vanished 'OTHER' head category "
     "(worth Rs. 5.2 Cr of gross profit last year) was lost, discontinued, or simply reclassified.",
     "See Source Data - GP and the GP REPORT.xlsx Executive Summary, Sections 2 and 4."),
    ("Medium", "Direct Expenses",
     "Direct expenses are up 83% year-on-year (Rs. 2.25 Cr to Rs. 4.12 Cr), well ahead of the ~4% dip in sales seen "
     "in the GP report over the same comparison — Custom Duty alone is up 55%.",
     "Get a line-by-line review of Direct Expenses and Custom Duty from the finance team before next quarter's "
     "numbers are closed, to confirm the increase is volume/rate-driven and not a coding or classification issue.",
     "See Source Data - Expenses."),
    ("Medium", "Pigment Ink",
     "Pigment Ink gross margin fell from 30.4% to 9.2% year-on-year — the steepest margin drop of any ink line.",
     "Review Pigment Ink's input cost and selling price directly with the product/pricing team — a drop this steep "
     "usually means a cost increase hasn't been passed through in price.",
     "See GP REPORT.xlsx Executive Summary, Section 4."),
    ("Medium", "Ink — Working Capital",
     "Ink's net working capital rose Rs. 5.7 Cr this period, the single largest segment increase, and 72% of Ink "
     "receivables are also overdue beyond 180 days.",
     "Set a quarter-end target to recover 30% of Ink's aged receivables (~Rs. 8 Cr) and review whether Ink "
     "inventory levels are being ordered ahead of actual demand.",
     "See Source Data - Working Capital."),
    ("Medium", "Debtor Days vs Creditor Days",
     "The business collects from customers in ~96 days on average but pays suppliers in ~90 days — a gap that "
     "funds customers at the company's own expense every quarter.",
     "Set a target to bring Debtor Days down to 80 or below over the next two quarters through tighter credit "
     "terms on new orders and faster follow-up on overdue accounts.",
     "See Source Data - Ratios (Debtor Days / Creditor Days)."),
    ("Low", "Spare Parts — quantity spike",
     "Spare Parts sale quantity jumped from ~483 to ~10,667 units between May and June — a jump large enough to "
     "look like a data or unit-of-measure issue rather than genuine demand.",
     "Verify this figure against source records before using it in any external report or forecast.",
     "See GP REPORT.xlsx Executive Summary, Section 3."),
]

# =====================================================================
# ACTION PLAN (own sheet)
# =====================================================================
ws = wb.create_sheet("Action Plan", 2)
style_sheet(ws)
for col, w in {"A": 2, "B": 10, "C": 24, "D": 16, "E": 16, "F": 16,
               "G": 15, "H": 15, "I": 15, "J": 15, "K": 13, "L": 13, "M": 13, "N": 13}.items():
    ws.column_dimensions[col].width = w

merge_text(ws, "B2:N2", "Consolidated Action Plan — Q1 FY26-27 (Apr-Jun 2026)", size=18, bold=True, align="center")
merge_text(ws, "B3:N3",
           "Priority actions from the Company Review, with the finding, recommended action, and the source "
           "basis behind each one.", size=11, color=GREY_MUTED, align="center", italic=True)
add_link(ws, 4, 2, "Executive Summary", "Back to Executive Summary")

r = 6
table_header_spans(ws, r, [
    ("Priority", 2, 2), ("Area", 3, 3), ("What We Found", 4, 6),
    ("Recommended Action", 7, 10), ("Basis / Justification", 11, 14),
], size=11, height=28)
r += 1

ap_first = r
for priority, area, found, action, basis in action_plan:
    ws.cell(row=r, column=2, value=priority).border = box
    ws.cell(row=r, column=2).font = Font(bold=True, color=("9C0006" if priority == "High" else "1F3864"), size=10.5)
    ws.cell(row=r, column=2).alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    ws.cell(row=r, column=3, value=area).border = box
    ws.cell(row=r, column=3).font = Font(bold=True, size=10.5)
    ws.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    c_found = ws.cell(row=r, column=4, value=found)
    c_found.alignment = Alignment(wrap_text=True, vertical="top")
    c_found.font = Font(size=10)
    ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=10)
    c_action = ws.cell(row=r, column=7, value=action)
    c_action.alignment = Alignment(wrap_text=True, vertical="top")
    c_action.font = Font(size=10)
    ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=14)
    c_basis = ws.cell(row=r, column=11, value=basis)
    c_basis.alignment = Alignment(wrap_text=True, vertical="top")
    c_basis.font = Font(size=9, italic=True, color=GREY_MUTED)
    for cc in range(2, 15):
        ws.cell(row=r, column=cc).border = box
    lines_found = -(-len(found) // 46)
    lines_action = -(-len(action) // 58)
    lines_basis = -(-len(basis) // 50)
    est_lines = max(lines_found, lines_action, lines_basis, 2)
    ws.row_dimensions[r].height = max(42, est_lines * 15)
    r += 1
ap_last = r - 1

ws.freeze_panes = f"A{ap_first}"
ws.print_area = f"A1:N{ap_last + 1}"
ws.print_options.horizontalCentered = True

# =====================================================================
# EXPENSE DETAILS (own sheet — full Direct + Indirect breakdown, Q1 FY25-26 vs Q1 FY26-27)
# =====================================================================
ws = wb.create_sheet("Expense Details", 3)
style_sheet(ws)
for col, w in {"A": 2, "B": 30, "C": 14, "D": 14, "E": 11, "F": 16, "G": 16, "H": 16, "I": 16, "J": 16}.items():
    ws.column_dimensions[col].width = w

merge_text(ws, "B2:J2", "Expense Details — Direct & Indirect Breakdown (Q1 FY25-26 vs Q1 FY26-27)", size=18, bold=True, align="center")
merge_text(ws, "B3:J3",
           "Source: EXPENSES SHEET FOR MARCH 26 (Apr-Jun 2025) vs Expense Sheet 2026-07-14 (Apr-Jun 2026). "
           "Category names differ slightly between the two source files' schemas — matched by closest name. "
           "'Reason / Note' is an indicative driver based on the nature and scale of each category's movement, "
           "not a confirmed root cause — large or unusual moves should still be confirmed with accounts.",
           size=9, color=GREY_MUTED, italic=True)
ws.row_dimensions[3].height = 32
add_link(ws, 4, 2, "Executive Summary", "Back to Executive Summary")

DIRECT_ITEMS = [
    ("Custom Duty Expenses", 1.121992, 1.736157,
     "Import duty scales with landed cost of raw material imports — check if driven by volume, HSN rate change, or currency."),
    ("Direct Expenses (General/Unclassified)", None, 1.168232,
     "New/unclassified direct-cost bucket this year — get it reclassified into a named sub-category."),
    ("Freight Expenses", 1.036347, 0.800408,
     "Lower than last year — check if this reflects fewer despatches, better freight rates, or a change in freight terms with customers."),
    ("Clearing & Forwarding Expenses", 0.195583, 0.149793,
     "Tracks with lower freight/import activity this quarter."),
    ("Carriage Inward Expenses", 0.135724, 0.149045,
     "Broadly in line with last year."),
    ("Other Direct Expenses", 0.037609, 0.081180,
     "Small base — a one-off item can swing the % sharply; confirm with the ledger."),
    ("Service Expenses", 0.039254, 0.041253,
     "In line with last year."),
    ("Tools & Spare Expenses", 0.000097, 0.027554,
     "Very small base last year — increase likely a handful of purchase entries, not a trend."),
    ("Packing Expenses", 0.011159, 0.026309,
     "Small base; worth a one-line check with the packing/dispatch team."),
    ("Sample Fabric Expenses", 0.032436, 0.017811,
     "Lower sampling activity than last year."),
    ("Labour Cost", None, 0.013370,
     "New direct-labour head this year — confirm what changed in job-work/contract arrangements."),
    ("Exchange Rate Gain/(Loss)", -0.369761, -0.094225,
     "Both years show a net FOREIGN EXCHANGE GAIN (shown here as a negative expense); the gain is smaller this "
     "year — linked to rupee movement on import/export settlements, not a cost problem."),
    ("Security Expenses (Direct)", 0.008750, None,
     "Not separately coded this year — likely merged into Indirect Security Expenses below; confirm classification is consistent."),
]

INDIRECT_ITEMS = [
    ("Employee Cost", 3.005286, 3.680747,
     "Annual increments/hiring — reconcile against the HR headcount and increment schedule."),
    ("Finance Cost (Interest & Bank Charges)", 1.032705, 1.055721,
     "Broadly stable — interest cost has not risen in line with the working capital increase (see Working Capital detail)."),
    ("Employee Travel Cost", 0.659682, 0.618630,
     "Broadly stable."),
    ("Commission & Brokerage", 0.471152, 0.598921,
     "Tracks with sales-linked payouts; confirm against the sales/commission schedule."),
    ("Rent Expenses", 0.282723, 0.415394,
     "Reflects a new/renewed lease or rent escalation — confirm with the admin/facilities team."),
    ("Insurance Expenses", 0.188199, 0.170381,
     "Broadly stable."),
    ("Transportation Expenses", 0.159467, 0.120176,
     "Lower than last year — check if linked to fewer local trips or a vendor change."),
    ("Subscription & Membership", 0.054942, 0.118018,
     "More than doubled on a small base — confirm which new subscriptions were added."),
    ("Electricity Expenses", 0.060691, 0.110312,
     "Sharp rise on a small base — worth checking against unit consumption, not just tariff."),
    ("Printing & Courier", 0.110890, 0.095467,
     "Broadly stable."),
    ("Learning & Development", None, 0.089471,
     "New training spend this year — confirm this reflects a planned initiative, not a coding change."),
    ("Repair & Maintenance", 0.180364, 0.088343,
     "Lower spend than last year — confirm this isn't deferred maintenance building up a future liability."),
    ("Advertising & Promotion", 0.335427, 0.077409,
     "Sharp cut in marketing spend — check whether this is a deliberate decision or a timing shift (spend not yet booked)."),
    ("Legal & Professional Fees", 0.121175, 0.074638,
     "Lower than last year; last year's base may have included one-off advisory work."),
    ("Office Expenses", 0.114655, 0.055655,
     "Lower than last year."),
    ("Environmental Protection", None, 0.031460,
     "New head this year — confirm nature of spend (compliance-driven vs discretionary)."),
    ("Security Expenses (Indirect)", 0.021351, 0.021371,
     "Stable year-on-year."),
    ("Computer & Printer Expenses", 0.041679, 0.020800,
     "Lower than last year."),
    ("Laboratory Expenses", 0.000145, 0.018274,
     "Effectively a new spend line — confirm what lab work started this year."),
    ("Carriage Outward Expenses", None, 0.016321,
     "New/separately coded this year — check if this was previously bundled into Freight or Transportation."),
    ("Telephone Expenses", 0.018375, 0.012800,
     "Lower than last year."),
    ("Donation & Subscription", 0.001575, 0.006000,
     "Tiny base — not material."),
    ("Taxes (Non-Income)", 0.009864, 0.004410,
     "Lower than last year; not material to total expenses."),
    ("Indirect Expenses (General/Unclassified)", None, 0.000571,
     "Negligible amount; not material."),
    ("Other Expenses", 0.002720, None,
     "Not separately coded this year; negligible amount, not material."),
]


def expense_table(ws, r, title, items):
    merge_text(ws, f"B{r}:J{r}", title, size=13, bold=True)
    r += 1
    table_header_spans(ws, r, [
        ("Category", 2, 2), ("Q1 FY25-26 (Rs. Cr)", 3, 3), ("Q1 FY26-27 (Rs. Cr)", 4, 4),
        ("Change %", 5, 5), ("Reason / Note", 6, 10),
    ], size=10, height=30)
    r += 1
    first = r
    for label, before, now, note in items:
        c_lbl = ws.cell(row=r, column=2, value=label)
        c_lbl.border = box
        c_lbl.font = Font(size=10)
        c_lbl.alignment = Alignment(wrap_text=True, vertical="top")

        c_b = ws.cell(row=r, column=3, value=before if before is not None else None)
        c_b.border = box
        c_b.number_format = '0.00" Cr"'
        c_b.alignment = Alignment(horizontal="right", vertical="top")
        if before is None:
            c_b.value = "—"
            c_b.alignment = Alignment(horizontal="center", vertical="top")

        c_n = ws.cell(row=r, column=4, value=now if now is not None else None)
        c_n.border = box
        c_n.number_format = '0.00" Cr"'
        c_n.alignment = Alignment(horizontal="right", vertical="top")
        if now is None:
            c_n.value = "—"
            c_n.alignment = Alignment(horizontal="center", vertical="top")

        c_chg = ws.cell(row=r, column=5)
        c_chg.border = box
        c_chg.alignment = Alignment(horizontal="center", vertical="top")
        if before is not None and now is not None and before != 0:
            c_chg.value = f"=(D{r}-C{r})/C{r}"
            c_chg.number_format = "0.0%"
        elif before is None:
            c_chg.value = "New"
            c_chg.font = Font(italic=True, color=GREY_MUTED, size=9)
        elif now is None:
            c_chg.value = "Discontinued"
            c_chg.font = Font(italic=True, color=GREY_MUTED, size=9)

        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=10)
        c_note = ws.cell(row=r, column=6, value=note)
        c_note.alignment = Alignment(wrap_text=True, vertical="top")
        c_note.font = Font(size=9, italic=True, color=GREY_MUTED)
        for cc in range(6, 11):
            ws.cell(row=r, column=cc).border = box

        lines = max(-(-len(note) // 78), 1)
        ws.row_dimensions[r].height = max(28, lines * 14)
        r += 1
    last = r - 1
    green_red_cf(ws, f"E{first}:E{last}", favorable_when_positive=False)

    # subtotal row
    b_range = f"C{first}:C{last}"
    n_range = f"D{first}:D{last}"
    ws.cell(row=r, column=2, value=f"{title} — Total").font = Font(bold=True, size=10.5)
    ws.cell(row=r, column=2).border = box
    c_bt = ws.cell(row=r, column=3, value=f"=SUM({b_range})")
    c_bt.number_format = '0.00" Cr"'
    c_bt.font = Font(bold=True)
    c_bt.border = box
    c_bt.alignment = Alignment(horizontal="right")
    c_nt = ws.cell(row=r, column=4, value=f"=SUM({n_range})")
    c_nt.number_format = '0.00" Cr"'
    c_nt.font = Font(bold=True)
    c_nt.border = box
    c_nt.alignment = Alignment(horizontal="right")
    c_pt = ws.cell(row=r, column=5, value=f"=(D{r}-C{r})/C{r}")
    c_pt.number_format = "0.0%"
    c_pt.font = Font(bold=True)
    c_pt.border = box
    c_pt.alignment = Alignment(horizontal="center")
    for cc in range(6, 11):
        ws.cell(row=r, column=cc).border = box
    green_red_cf(ws, f"E{r}", favorable_when_positive=False)
    total_row = r
    r += 2
    return r, total_row


r = 6
r, direct_total_row = expense_table(ws, r, "Direct Expenses", DIRECT_ITEMS)
r, indirect_total_row = expense_table(ws, r, "Indirect Expenses", INDIRECT_ITEMS)

ws.cell(row=r, column=2, value="Grand Total (Direct + Indirect)").font = Font(bold=True, size=12, color=NAVY)
ws.cell(row=r, column=2).border = box
c_bt = ws.cell(row=r, column=3, value=f"=C{direct_total_row}+C{indirect_total_row}")
c_bt.number_format = '0.00" Cr"'
c_bt.font = Font(bold=True, size=12)
c_bt.border = box
c_bt.alignment = Alignment(horizontal="right")
c_nt = ws.cell(row=r, column=4, value=f"=D{direct_total_row}+D{indirect_total_row}")
c_nt.number_format = '0.00" Cr"'
c_nt.font = Font(bold=True, size=12)
c_nt.border = box
c_nt.alignment = Alignment(horizontal="right")
c_pt = ws.cell(row=r, column=5, value=f"=(D{r}-C{r})/C{r}")
c_pt.number_format = "0.0%"
c_pt.font = Font(bold=True, size=12)
c_pt.border = box
c_pt.alignment = Alignment(horizontal="center")
for cc in range(6, 11):
    ws.cell(row=r, column=cc).border = box
green_red_cf(ws, f"E{r}", favorable_when_positive=False)
expense_details_last = r

ws.freeze_panes = "A6"
ws.print_area = f"A1:J{expense_details_last + 1}"
ws.print_options.horizontalCentered = True

wb.save(OUT)
print("Source Data sheets built. Row markers:")
print("PL:", pl_first, pl_last, pl_company_first, pl_company_last)
print("BS:", bs_first, bs_last, gap_first, gap_last)
print("Ratios:", ratio_first, ratio_last, surat_first, surat_last)
print("GP:", GP_OVERALL_FIRST, GP_OVERALL_LAST, GP_SEG_SALES_FIRST, seg_sales_last, GP_SEG_GP_FIRST, seg_gp_last)
print("Expenses:", EXP_TOTAL_FIRST, EXP_TOTAL_LAST, EXP_CAT_FIRST, EXP_CAT_LAST)
print("WC:", WC_GLANCE_FIRST, WC_GLANCE_LAST, WC_SEG_FIRST, WC_SEG_LAST)
print("Action Plan rows:", ap_first, ap_last)
print("Expense Details rows: direct_total", direct_total_row, "indirect_total", indirect_total_row, "grand_total", expense_details_last)
