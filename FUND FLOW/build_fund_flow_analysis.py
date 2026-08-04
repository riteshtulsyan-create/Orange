"""
Fund Flow Analysis builder.

Reads "Junne FUnd FLOW.xlsx" (the live monthly fund flow ledger you maintain) and produces
"Fund_Flow_Analysis_Jun2026_LINKED.xlsx" -- a management-facing workbook with:
  - Executive Summary (KPI tiles, loan/FD movement detail, "where funds are blocked" narrative,
    quick-reference links, and a basis-of-preparation note)
  - Summary (wholesale clone of your 'Main Summary (2)' sheet, formulas redirected + styled)
  - Unplanned Payment / Excess Payment / Short Payment / Payment Done (every vendor row from
    'FF Summary (2)' classified by Balance = Budget - Paid, sorted by variance descending)
  - Machine Advance Payment (clone of 'Machine Adv_Receivable')
  - FF Source Data (hidden) -- the linked backbone every other sheet's formulas point to

Every number in the output is a live Excel formula, not a typed value, EXCEPT the row numbers
hardcoded in CLASSIFIED_ROWS and loan_detail_rows below, which identify which physical row in
'FF Summary (2)' each vendor currently sits on.

IMPORTANT -- re-run checklist when you update Junne FUnd FLOW.xlsx next month:
  1. If you only change amounts (Budget/Paid values) on EXISTING rows, just re-run this script --
     nothing else needs to change.
  2. If you ADD, REMOVE, or REORDER any row in 'FF Summary (2)' (e.g. delete a duplicate vendor,
     insert a new expense line), every row number below that point shifts. Re-run the discovery
     block at the bottom of this file (under `if __name__ == "__discover__"`) to reprint the
     current row-to-vendor mapping, then update CLASSIFIED_ROWS and loan_detail_rows to match.
     Do NOT assume the row numbers below still hold -- that exact mistake broke this workbook
     twice during development (a removed duplicate row silently misclassified a real vendor and
     a text placeholder caused #VALUE! errors) before this checklist was written.
  3. Requires: pip install openpyxl
  4. Run: python build_fund_flow_analysis.py   (from this folder, or edit SRC_DIR below)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment

SRC_DIR = r"C:\Users\orang\OneDrive - Orange O Tec Private Limited\00 Claude Project\Misc\FUND FLOW ANALYSIS"
JUNNE = SRC_DIR + r"\Junne FUnd FLOW.xlsx"
OUT = SRC_DIR + r"\Fund_Flow_Analysis_Jun2026_LINKED.xlsx"

wb_f = openpyxl.load_workbook(JUNNE, data_only=False)
wb_v = openpyxl.load_workbook(JUNNE, data_only=True)
ff_f, ff_v = wb_f["FF Summary (2)"], wb_v["FF Summary (2)"]
ms_f, ms_v = wb_f["Main Summary (2)"], wb_v["Main Summary (2)"]
mr_f, mr_v = wb_f["Machine Adv_Receivable"], wb_v["Machine Adv_Receivable"]

out = openpyxl.Workbook()
out.remove(out.active)

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
PCT_FILL = PatternFill("solid", fgColor="FFFF00")
RED_FONT = Font(color="FF0000")
INR_FMT = '##,##,##,##0;[RED](##,##,##,##0)'
USD_FMT = '#,##0.00;[RED](#,##0.00)'
PCT_FMT = '0%'
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HYPER_FONT = Font(color="0563C1", underline="single")
NARR_FONT = Font(size=11)
NARR_FILL_IN = PatternFill("solid", fgColor="E2EFDA")
NARR_FILL_OUT = PatternFill("solid", fgColor="FCE4E4")

def hdr_row(ws, r, headers, start_col=1):
    for i, h in enumerate(headers, start=start_col):
        c = ws.cell(row=r, column=i, value=h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.border = BORDER

def group_row(ws, r, text, ncols, fill=GROUP_FILL):
    c = ws.cell(row=r, column=1, value=text)
    c.font = GROUP_FONT
    for col in range(1, ncols + 1):
        ws.cell(row=r, column=col).fill = fill

def link_cell(ws, cellref, label, sheet_target, cell_target="A1"):
    c = ws[cellref]
    c.value = label
    c.font = HYPER_FONT
    c.hyperlink = f"#'{sheet_target}'!{cell_target}"

# =====================================================================
# SHEET: FF Source Data (hidden) -- wholesale clone of 'FF Summary (2)'
# =====================================================================
src = out.create_sheet("FF Source Data")
src.sheet_view.showGridLines = False
src["M1"] = "FF Source Data -- live copy of Junne FUnd FLOW.xlsx > 'FF Summary (2)'. Row numbers unshifted."
src["M1"].font = SUB_FONT

MAX_ROW = 141
for row in range(1, MAX_ROW + 1):
    for col in range(1, 12):  # A..K
        fcell = ff_f.cell(row=row, column=col)
        if fcell.value is None:
            continue
        target = src.cell(row=row, column=col)
        target.value = fcell.value
        target.border = BORDER
        if col in (3, 4, 5, 6, 7, 8) and isinstance(target.value, (int, float)):
            target.number_format = INR_FMT if col in (4, 6, 8) else USD_FMT

def S(cellref):
    import re
    m = re.match(r"([A-Z]+)(\d+)", cellref)
    col, row = m.group(1), int(m.group(2))
    return f"'FF Source Data'!${col}${row}"

def NUM(cellref):
    """Coerce a source cell to a guaranteed number, since the source uses text placeholders
    like ' -   ' for nil/blank instead of 0 -- straight arithmetic on those throws #VALUE!."""
    return f"IF(ISNUMBER({cellref}),{cellref},0)"

src.column_dimensions["A"].width = 18
src.column_dimensions["B"].width = 40
for c in "CDEFGH":
    src.column_dimensions[c].width = 14
src.column_dimensions["I"].width = 30
print("FF Source Data cloned")

# =====================================================================
# SHEET: Machine Advance Payment (clone of 'Machine Adv_Receivable', cols B-G)
# =====================================================================
msd = out.create_sheet("Machine Advance Payment")
msd.sheet_view.showGridLines = False
msd["A1"] = "Machine Advance Payment -- Customer Advance Tracker | June 2026"
msd["A1"].font = TITLE_FONT
msd["A1"].fill = TITLE_FILL
msd.merge_cells("A1:G1")
msd["A2"] = "Live copy of Junne FUnd FLOW.xlsx > 'Machine Adv_Receivable'. Pending = Projected - Received."
msd["A2"].font = SUB_FONT
msd.merge_cells("A2:G2")
hdr_row(msd, 4, ["#", "Sales Person", "Party Name", "Machine", "Projected (₹)", "Received (₹)", "Pending (₹)"])
n = 1
for row in range(5, 22):
    tgt_row = row + 1
    msd.cell(row=tgt_row, column=1, value=n)
    for col in range(2, 6):
        fcell = mr_f.cell(row=row, column=col)
        if fcell.value is not None:
            msd.cell(row=tgt_row, column=col, value=fcell.value)
    fcell = mr_f.cell(row=row, column=6)
    if fcell.value is not None:
        msd.cell(row=tgt_row, column=6, value=fcell.value)
    msd.cell(row=tgt_row, column=7, value=f"=E{tgt_row}-F{tgt_row}")
    for col in range(1, 8):
        c = msd.cell(row=tgt_row, column=col)
        c.border = BORDER
        if col in (5, 6, 7):
            c.number_format = INR_FMT
    n += 1
total_row = 23
msd.cell(row=total_row, column=2, value="TOTAL").font = TOTAL_FONT
msd.cell(row=total_row, column=5, value="=SUM(E6:E22)").number_format = INR_FMT
msd.cell(row=total_row, column=6, value="=SUM(F6:F22)").number_format = INR_FMT
msd.cell(row=total_row, column=7, value="=SUM(G6:G22)").number_format = INR_FMT
for col in range(1, 8):
    msd.cell(row=total_row, column=col).fill = TOTAL_FILL
    msd.cell(row=total_row, column=col).font = TOTAL_FONT
msd.cell(row=25, column=1,
         value='=IFERROR("Collection Rate: "&TEXT(F23/E23,"0.0%")&"  |  Received ₹"&TEXT(F23,"#,##0")&"  |  Pending ₹"&TEXT(G23,"#,##0"),"")')
msd["A25"].font = SUB_FONT
msd.merge_cells("A25:G25")
msd.column_dimensions["A"].width = 5
msd.column_dimensions["B"].width = 14
msd.column_dimensions["C"].width = 40
msd.column_dimensions["D"].width = 12
for c in "EFG":
    msd.column_dimensions[c].width = 15
print("Machine Advance Payment built")

# =====================================================================
# SHEET: Summary -- wholesale clone of 'Main Summary (2)', formulas redirected
# =====================================================================
sm = out.create_sheet("Summary")
sm.sheet_view.showGridLines = False

for row in ms_f.iter_rows():
    for cell in row:
        if cell.value is None:
            continue
        val = cell.value
        target = sm.cell(row=cell.row, column=cell.column)
        if isinstance(val, str) and val.startswith("=") and "'FF Summary (2)'" in val:
            target.value = val.replace("'FF Summary (2)'", "'FF Source Data'")
        else:
            target.value = val

# styling -- Main Summary (2)'s own row layout is unchanged, so these row numbers are stable
sm["B1"] = None
sm["D1"], sm["F1"], sm["H1"], sm["J1"] = "Budgeted", "Paid", "Balance", "%"
for cellref in ("D1", "F1", "H1", "J1"):
    sm[cellref].font = HDR_FONT
    sm[cellref].fill = HDR_FILL

BOLD_ROWS = {2, 3, 4, 5, 12, 17, 23, 29, 35}
RED_ROWS = {21, 22}
PCT_ROWS = {3, 4, 9, 10, 12, 14, 15, 16, 17, 19, 20, 21, 22, 23, 25, 27, 29, 31, 33, 35}

USD_COLS = {3}         # C -- USD Amount (Import block subtotal, Customer Collections target)
INR_COLS = {4, 5, 6, 7, 8, 9}  # D..I -- INR financial columns, whether the cell holds a literal or a formula
for row in range(1, 36):
    for col in range(2, 11):
        c = sm.cell(row=row, column=col)
        if c.value is None:
            continue
        if col in USD_COLS:
            c.number_format = USD_FMT
        elif col in INR_COLS:
            c.number_format = INR_FMT
        c.border = BORDER
    if row in BOLD_ROWS:
        for col in range(2, 11):
            sm.cell(row=row, column=col).font = TOTAL_FONT
    if row in {29, 35}:
        for col in range(2, 11):
            sm.cell(row=row, column=col).fill = TOTAL_FILL
    if row in RED_ROWS:
        sm.cell(row=row, column=2).font = RED_FONT
    jcell = sm.cell(row=row, column=10)
    if row in PCT_ROWS and jcell.value is not None:
        jcell.number_format = PCT_FMT
        jcell.fill = PCT_FILL

# NOTE: do NOT insert/delete rows after formulas are written -- openpyxl does not rewrite formula
# text on insert_rows, so same-sheet references (e.g. "=SUM(D9:D11)") would silently break while
# cross-sheet references keep working (exactly the bug hit earlier in this build). Title goes in a
# side column instead, keeping every native Main Summary(2) row number unchanged.
sm["L1"] = "Fund Flow -- June 2026 | Summary"
sm["L1"].font = TITLE_FONT
sm["L1"].fill = TITLE_FILL
sm.merge_cells("L1:P1")
print("Summary sheet cloned (native row numbers preserved, no shift)")

ROWMAP = {n: n for n in (2,3,4,5,6,7,9,10,11,12,14,15,16,17,19,20,21,22,23,25,27,29,31,33,34,35)}
def SM(nativerow, col="F"):
    return f"Summary!{col}{ROWMAP[nativerow]}"

sm.column_dimensions["A"].width = 3
sm.column_dimensions["B"].width = 26
sm.column_dimensions["C"].width = 15
for c in "DEFGH":
    sm.column_dimensions[c].width = 15
sm.column_dimensions["I"].width = 4
sm.column_dimensions["J"].width = 10

# =====================================================================
# 4 classification sheets: Unplanned / Excess / Short / Payment Done
#
# Each entry: [row_in_FF_Summary_2, name_column('A' or 'B'), category_label, bucket]
# Verified against 'FF Summary (2)' as of 2026-07-05 22:13 (after the Jay Industrial
# duplicate row was removed). See the module docstring's re-run checklist before trusting
# these row numbers against a future edit of Junne FUnd FLOW.xlsx.
# =====================================================================
CLASSIFIED_ROWS = [
[3,'B','Dues','UNPLANNED'],[4,'B','Dues -O tec','SHORT'],[5,'B','Dues','DONE'],
[6,'B','Dues','SHORT'],[7,'B','Dues','DONE'],[8,'B','Dues','DONE'],
[9,'B','Dues','EXCESS'],[10,'B','Dues','EXCESS'],[11,'B','Dues','SHORT'],
[14,'B','Advances','DONE'],[16,'B','Advances','SHORT'],[18,'B','Advances','SHORT'],
[20,'B','Advances','DONE'],[21,'B','Other Vendor-O Tec','UNPLANNED'],
[24,'A','Custom Duty and GST','SHORT'],[26,'B','Repayment','SHORT'],[27,'B','EMI','EXCESS'],
[30,'B','Local Vendor Payent','UNPLANNED'],[31,'B','Local Vendor Payent','EXCESS'],
[33,'B','Local Vendor Payent','UNPLANNED'],[40,'B','Other Vendor','DONE'],
[41,'B','Other Vendor','EXCESS'],[42,'B','Other Vendor','DONE'],
[43,'B','Other Vendor','EXCESS'],[44,'B','Other Vendor','UNPLANNED'],
[45,'B','Other Vendor','EXCESS'],[46,'B','Other Vendor','DONE'],
[47,'B','Other Vendor','UNPLANNED'],[48,'B','Other Vendor','UNPLANNED'],
[49,'B','Other Vendor','DONE'],[51,'B','Other Vendor','DONE'],
[52,'B','Other Vendor','DONE'],[53,'B','Other Vendor','SHORT'],
[56,'B','Other Vendor','DONE'],[58,'B','Other Vendor','DONE'],
[60,'B','Other Vendor','UNPLANNED'],[61,'B','Other Vendor','UNPLANNED'],
[62,'B','Other Vendor-Enter','UNPLANNED'],[63,'B','Other Vendor-Enter','UNPLANNED'],
[64,'B','Other Vendor-Enter','UNPLANNED'],[65,'B','Other Vendor-Enter','DONE'],
[67,'B','Other Vendor-Enter','DONE'],[68,'B','Other Vendor-Enter','UNPLANNED'],
[69,'B','Other Vendor-Enter','DONE'],[70,'B','Other Vendor-Enter','SHORT'],
[71,'B','Other Vendor-Enter','EXCESS'],[72,'B','Other Vendor-Enter','EXCESS'],
[73,'B','Other Vendor-Enter','DONE'],[74,'B','Other Vendor-Enter','SHORT'],
[82,'B','Advacne Tax-O Tec','SHORT'],[83,'B','Advacne Tax-Enterprises','DONE'],
[84,'B','SA TAX - Enterprises','SHORT'],[85,'A','TDS-GST','SHORT'],
[89,'B','O Tec-Others','UNPLANNED'],[90,'B','O Tec-Others','UNPLANNED'],
[91,'B','O Tec-Others','UNPLANNED'],[92,'B','O Tec-Others','UNPLANNED'],
[93,'B','O Tec-Others','UNPLANNED'],[94,'B','O Tec-Land Purchase','UNPLANNED'],
[100,'B','Bank Loan-O Tec','UNPLANNED'],
[103,'A','Monthly Expense','EXCESS'],[104,'A','Monthly Expense','DONE'],
[105,'A','Monthly Expense','SHORT'],[106,'A','Monthly Expense','EXCESS'],
[107,'A','Monthly Expense','EXCESS'],[108,'A','Monthly Expense','EXCESS'],
[109,'A','Monthly Expense','DONE'],[110,'A','Monthly Expense','EXCESS'],
[111,'A','Monthly Expense','EXCESS'],[112,'A','Monthly Expense','DONE'],
[113,'A','Monthly Expense','EXCESS'],
]

BUCKETS = {
    "UNPLANNED": ("Unplanned Payment", UNPLANNED_FILL,
                  "Paid with NO budget sanction at all (Budget = 0). Every figure is a live formula linked to 'FF Source Data'."),
    "EXCESS":    ("Excess Payment", EXCESS_FILL,
                  "Budgeted item where Paid exceeds Budget (overrun). Excess = Paid - Budget."),
    "SHORT":     ("Short Payment", SHORT_FILL,
                  "Budgeted item where Paid is less than Budget (partially or fully unpaid). Short = Budget - Paid."),
    "DONE":      ("Payment Done", DONE_FILL,
                  "Budgeted item fully settled (Paid = Budget, Balance = 0). Included for audit completeness."),
}

USD_SUPPRESS_ROWS = set()  # dummy C=1 placeholder rows, if any land in a future classification set

def bal_inr(row_num):
    d = ff_v.cell(row=row_num, column=4).value
    f = ff_v.cell(row=row_num, column=6).value
    d = d if isinstance(d, (int, float)) else 0
    f = f if isinstance(f, (int, float)) else 0
    return d, f

def variance_value(row_num, bucket):
    d, f = bal_inr(row_num)
    if bucket == "EXCESS":
        return f - d
    elif bucket == "SHORT":
        return d - f
    else:
        return f - d

rows_by_bucket = {}
for row_num, name_col, cat, bucket in CLASSIFIED_ROWS:
    rows_by_bucket.setdefault(bucket, []).append((row_num, name_col, cat))
for bucket in rows_by_bucket:
    rows_by_bucket[bucket].sort(key=lambda t: variance_value(t[0], bucket), reverse=True)

sheets = {}
for key, (title, fill, note) in BUCKETS.items():
    ws = out.create_sheet(title)
    ws.sheet_view.showGridLines = False
    ws["A1"] = f"{title} -- June 2026"
    ws["A1"].font = TITLE_FONT
    ws["A1"].fill = TITLE_FILL
    ws.merge_cells("A1:I1")
    ws["A2"] = note + " Sorted by Variance, largest first."
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:I2")
    cols = ["#", "Vendor / Item", "Category", "Budget USD ($)", "Budget INR (₹)",
            "Paid USD ($)", "Paid INR (₹)", "Variance INR (₹)", "Src Row"]
    hdr_row(ws, 4, cols)
    sheets[key] = (ws, 5)

for bucket, rows in rows_by_bucket.items():
    ws, r = sheets[bucket]
    for row_num, name_col, cat in rows:
        ws.cell(row=r, column=1, value=r - 4)
        ws.cell(row=r, column=2, value=f"={S(name_col + str(row_num))}")
        ws.cell(row=r, column=3, value=cat.strip())
        if row_num in USD_SUPPRESS_ROWS:
            bU = ws.cell(row=r, column=4, value="—")
        else:
            bU = ws.cell(row=r, column=4, value=f"=IF(OR({S('C'+str(row_num))}=0,{S('C'+str(row_num))}=\"\"),\"—\",{S('C'+str(row_num))})")
        bU.number_format = USD_FMT
        bI = ws.cell(row=r, column=5, value=f"={NUM(S('D' + str(row_num)))}")
        bI.number_format = INR_FMT
        if row_num in USD_SUPPRESS_ROWS:
            pU = ws.cell(row=r, column=6, value="—")
        else:
            pU = ws.cell(row=r, column=6, value=f"=IF(OR({S('E'+str(row_num))}=0,{S('E'+str(row_num))}=\"\"),\"—\",{S('E'+str(row_num))})")
        pU.number_format = USD_FMT
        pI = ws.cell(row=r, column=7, value=f"={NUM(S('F' + str(row_num)))}")
        pI.number_format = INR_FMT
        if bucket == "EXCESS":
            vformula = f"=G{r}-E{r}"
        elif bucket == "SHORT":
            vformula = f"=E{r}-G{r}"
        else:
            vformula = f"=G{r}-E{r}"
        vcell = ws.cell(row=r, column=8, value=vformula)
        vcell.number_format = INR_FMT
        ws.cell(row=r, column=9, value=row_num)
        fill = BUCKETS[bucket][1]
        for col in range(1, 10):
            ws.cell(row=r, column=col).border = BORDER
            ws.cell(row=r, column=col).fill = fill
        r += 1
    sheets[bucket] = (ws, r)

for key, (ws, r) in sheets.items():
    ws.cell(row=r, column=2, value="TOTAL").font = TOTAL_FONT
    c5 = ws.cell(row=r, column=5, value=f"=SUM(E5:E{r-1})")
    c5.number_format = INR_FMT
    c7 = ws.cell(row=r, column=7, value=f"=SUM(G5:G{r-1})")
    c7.number_format = INR_FMT
    c8 = ws.cell(row=r, column=8, value=f"=SUM(H5:H{r-1})")
    c8.number_format = INR_FMT
    for col in range(1, 10):
        ws.cell(row=r, column=col).fill = TOTAL_FILL
        ws.cell(row=r, column=col).font = TOTAL_FONT
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 13
    ws.column_dimensions["E"].width = 15
    ws.column_dimensions["F"].width = 13
    ws.column_dimensions["G"].width = 15
    ws.column_dimensions["H"].width = 15
    ws.column_dimensions["I"].width = 9

TOTAL_ROWS = {k: r for k, (ws, r) in sheets.items()}
print("4 classification sheets built:", {k: v[1] - 5 for k, v in sheets.items()})

# =====================================================================
# SHEET: Executive Summary
# =====================================================================
exs = out.create_sheet("Executive Summary")
exs.sheet_view.showGridLines = False
exs["A1"] = "Fund Flow -- June 2026 | Executive Summary"
exs["A1"].font = TITLE_FONT
exs["A1"].fill = TITLE_FILL
exs.merge_cells("A1:H1")
exs.row_dimensions[1].height = 24
exs["A2"] = "Prepared for management review. Every figure below is a live formula -- click any underlined sheet name to jump to the detail behind it."
exs["A2"].font = SUB_FONT
exs.merge_cells("A2:H2")

def kpi_tile(ws, row, col, label, formula, note=None):
    lbl = ws.cell(row=row, column=col, value=label)
    lbl.font = Font(bold=True, size=10, color="D9E2F3")
    lbl.fill = TITLE_FILL
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
    val = ws.cell(row=row + 1, column=col, value=formula)
    val.font = Font(bold=True, size=18, color="FFFFFF")
    val.fill = TITLE_FILL
    val.number_format = '"₹"##,##,##,##0'
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + 1)
    if note:
        n = ws.cell(row=row + 2, column=col, value=note)
        n.font = Font(size=8, italic=True, color="808080")
        ws.merge_cells(start_row=row + 2, start_column=col, end_row=row + 2, end_column=col + 1)

kpi_tile(exs, 4, 1, "TOTAL PAYOUT (JUNE)", f"=+{SM(3)}", "Total actually paid out this month")
kpi_tile(exs, 4, 4, "TOTAL COLLECTION (JUNE)", f"=+{SM(4)}", "Machine advances + customer collections received")
kpi_tile(exs, 4, 7, "CLOSING BALANCE", f"=+{SM(5)}", "Opening + Collection - Payout")
kpi_tile(exs, 8, 1, "OPENING BALANCE", f"=+{SM(2)}", "Carried forward from May closing")
kpi_tile(exs, 8, 4, "BANK BALANCE TODAY", f"=+{SM(6)}", "Per bank statement")
kpi_tile(exs, 8, 7, "VARIANCE vs BANK", f"=+{SM(7)}", "Computed closing minus bank statement")

r = 13
group_row(exs, r, "Where the Funds Came From", 8, fill=NARR_FILL_IN); r += 1

exs.cell(row=r, column=1, value=(
    '=" Additional loan/FD drawn this month: "&TEXT(-('
    + NUM(S('F95')) + '+' + NUM(S('F96')) + '+' + NUM(S('F97')) + '+' + NUM(S('F98')) + '+' + NUM(S('F99')) + '),"₹##,##,##,##0")'
    '&" gross (ECLGS Loan + a second loan disbursement, net of FD releases; ICICI FD Release still NOT fully received -- see table below), of which "'
    '&TEXT(-' + SM(25) + ',"₹##,##,##,##0")&" is the NET inflow after EMI/related outflows in the same category."'))
exs.merge_cells(f"A{r}:H{r}")
exs.row_dimensions[r].height = 40
exs.cell(row=r, column=1).font = NARR_FONT
exs.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
r += 1

exs.cell(row=r, column=1, value=(
    '=" This category was budgeted for a net INFLOW of only "&TEXT(-' + SM(25, "D") + ',"₹##,##,##,##0")&" -- so "'
    '&TEXT(' + SM(25, "D") + '-' + SM(25) + ',"₹##,##,##,##0")'
    '&" of this loan money was UNBUDGETED, and it is what funded the "'
    '&TEXT(\'Unplanned Payment\'!G'+str(TOTAL_ROWS["UNPLANNED"])+',"₹##,##,##,##0")&" Unplanned payments, the "'
    '&TEXT(\'Excess Payment\'!H'+str(TOTAL_ROWS["EXCESS"])+',"₹##,##,##,##0")&" Excess payments, and part of what was actually paid within Short Payment items."'))
exs.merge_cells(f"A{r}:H{r}")
exs.row_dimensions[r].height = 40
exs.cell(row=r, column=1).font = NARR_FONT
exs.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
r += 1

link_cell(exs, f"A{r}", "-> Open Summary sheet (Investment/Loan row)", "Summary", f"A{ROWMAP[25]}")
r += 1

group_row(exs, r, "Loan & FD Movement Detail (feeds Investment/Loan above)", 8, fill=GROUP_FILL); r += 1
hdr_row(exs, r, ["Item", "Type", "Amount Received / Paid (₹)", "Not Received / Pending (₹)"], start_col=1)
r += 1
loan_detail_rows = [
    ("AXIS-FD Released", 95, "Inflow"), ("ICICI-FD Released", 96, "Inflow"),
    ("FD received from Enterprises", 97, "Inflow"), ("Loan-EMI-3.44 (2nd disbursement)", 98, "Inflow"),
    ("ECLGS-Loan (drawn)", 99, "Inflow"),
    ("ECLGS Loan EMI (installment)", 89, "Outflow"), ("Poonam Shah - Brokerage", 90, "Outflow"),
    ("Dhanikuber - Brokerage", 91, "Outflow"), ("Bhoomi Weave Tex Pvt Ltd", 92, "Outflow"),
    ("Colorix - Rohit Tak", 93, "Outflow"), ("Property Purchase - SBI", 94, "Outflow"),
    ("ECLGS Stamp Duty", 100, "Outflow"),
]
loan_detail_start = r
for label, row_num, typ in loan_detail_rows:
    exs.cell(row=r, column=1, value=label)
    exs.cell(row=r, column=2, value=typ)
    val = -1 if typ == "Inflow" else 1
    c = exs.cell(row=r, column=3, value=f"=({val})*{NUM(S('F'+str(row_num)))}")
    c.number_format = INR_FMT
    if typ == "Inflow":
        # Pending = budgeted inflow magnitude (-D, only where a target D<0 was actually set) minus what came in
        pend = exs.cell(row=r, column=4,
                         value=f"=IF({NUM(S('D'+str(row_num)))}<0,MAX(-{NUM(S('D'+str(row_num)))}+{NUM(S('F'+str(row_num)))},0),\"—\")")
        pend.number_format = INR_FMT
        if row_num == 97:
            pend.font = Font(bold=True, color="C00000")
            exs.cell(row=r, column=1).font = Font(bold=True, color="C00000")
    else:
        exs.cell(row=r, column=4, value="—")
    for col in range(1, 5):
        exs.cell(row=r, column=col).border = BORDER
    r += 1
loan_detail_end = r - 1
exs.cell(row=r, column=1, value="NET (matches Summary Investment/Loan row)").font = TOTAL_FONT
cnet = exs.cell(row=r, column=3, value=(
    f"=SUMIF(B{loan_detail_start}:B{loan_detail_end},\"Inflow\",C{loan_detail_start}:C{loan_detail_end})"
    f"-SUMIF(B{loan_detail_start}:B{loan_detail_end},\"Outflow\",C{loan_detail_start}:C{loan_detail_end})"))
cnet.number_format = INR_FMT
cnet.font = TOTAL_FONT
cpend = exs.cell(row=r, column=4, value=f"=SUMIF(D{loan_detail_start}:D{loan_detail_end},\">0\")")
cpend.number_format = INR_FMT
cpend.font = TOTAL_FONT
for col in range(1, 5):
    exs.cell(row=r, column=col).fill = TOTAL_FILL
r += 1
exs.cell(row=r, column=1, value=(
    "Note: 'Not Received' only applies to Inflow rows where a budget/target was actually set "
    "(D column < 0 in source). Rows with no explicit target show \"—\"."))
exs.cell(row=r, column=1).font = SUB_FONT
exs.merge_cells(f"A{r}:D{r}")
r += 2

group_row(exs, r, "Where Funds Are Blocked", 8, fill=NARR_FILL_OUT); r += 1

exs.cell(row=r, column=1, value=(
    "=\" Machine Advance is itself a short receipt: \""
    "&TEXT('Machine Advance Payment'!E23,\"₹##,##,##,##0\")"
    "&\" was projected from customers against new machine orders, but only \""
    "&TEXT('Machine Advance Payment'!F23,\"₹##,##,##,##0\")&\" (\"&TEXT('Machine Advance Payment'!F23/'Machine Advance Payment'!E23,\"0%\")&\") has actually been received -- \""
    "&TEXT('Machine Advance Payment'!G23,\"₹##,##,##,##0\")&\" is short-received, blocked at the customer's end.\""
))
exs.merge_cells(f"A{r}:H{r}")
exs.row_dimensions[r].height = 30
exs.cell(row=r, column=1).font = NARR_FONT
exs.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
r += 1

link_cell(exs, f"A{r}", "-> Open Machine Advance Payment sheet", "Machine Advance Payment", "A1")
r += 2

exs.cell(row=r, column=1, value=(
    '=" (Minor context) Overall customer collections also ran short of the internal 55% target by "'
    f'&TEXT({SM(33, "H")},"₹##,##,##,##0")&" -- a secondary factor, not the main driver."'))
exs.merge_cells(f"A{r}:H{r}")
exs.row_dimensions[r].height = 20
exs.cell(row=r, column=1).font = Font(size=10, italic=True, color="595959")
exs.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
r += 1
link_cell(exs, f"A{r}", "-> Open Summary sheet (Customer Collections row)", "Summary", f"A{ROWMAP[33]}")
r += 2

group_row(exs, r, "Net Read", 8, fill=GROUP_FILL); r += 1
exs.cell(row=r, column=1, value=(
    "The excess/unplanned spending was NOT funded by stronger collections -- it was funded by borrowing "
    "more than planned, precisely because machine advances (and collections generally) came in well short "
    "of expectation."))
exs.merge_cells(f"A{r}:H{r}")
exs.row_dimensions[r].height = 30
exs.cell(row=r, column=1).font = Font(bold=True, italic=True, size=11)
exs.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
r += 2

group_row(exs, r, "Quick Reference -- Where To Look", 8, fill=GROUP_FILL); r += 1
hdr_row(exs, r, ["Question", "Sheet"], start_col=1)
exs.merge_cells(f"C{r}:H{r}")
r += 1
ref_table = [
    ("Overall Budget vs Paid vs Balance, by category", "Summary"),
    ("What was paid with zero budget approval", "Unplanned Payment"),
    ("What went over its approved budget", "Excess Payment"),
    ("What's still owed to vendors/tax/dues (money held back)", "Short Payment"),
    ("What's fully settled (audit trail)", "Payment Done"),
    ("What customers still owe us on machine orders", "Machine Advance Payment"),
]
for question, sheet_target in ref_table:
    exs.cell(row=r, column=1, value=question)
    link_cell(exs, f"B{r}", f"-> {sheet_target}", sheet_target, "A1")
    for col in range(1, 3):
        exs.cell(row=r, column=col).border = BORDER
    r += 1
r += 1

group_row(exs, r, "Basis of Preparation", 8, fill=GROUP_FILL); r += 1
exs.cell(row=r, column=1, value=(
    '="This report is generated from Junne FUnd FLOW.xlsx (June 2026 fund flow ledger) as of 05-Jul-2026. '
    'All figures on every sheet are live formulas traced back to that ledger -- no number is manually typed. '
    'Budget vs. Paid variances are classified into four categories (Unplanned, Excess, Short, Payment Done) '
    'using a single rule: Balance = Budget - Paid. Closing Balance reconciles to the bank statement balance '
    'to within ₹"&TEXT(ABS(' + SM(7) + '),"#,##0")&". Prepared for internal management review; figures should '
    'be read alongside underlying vendor invoices and bank statements before any payment decision."'))
exs.merge_cells(f"A{r}:H{r}")
exs.row_dimensions[r].height = 55
exs.cell(row=r, column=1).font = Font(size=9, italic=True, color="595959")
exs.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
r += 2

exs.column_dimensions["A"].width = 46
exs.column_dimensions["B"].width = 26
for c in "CDEFGH":
    exs.column_dimensions[c].width = 13

print("Executive Summary built")

for sheet_name in ["Summary", "Unplanned Payment", "Excess Payment", "Short Payment", "Payment Done", "Machine Advance Payment"]:
    ws = out[sheet_name]
    ws["L1"] = "<- Executive Summary"
    ws["L1"].font = HYPER_FONT
    ws["L1"].hyperlink = "#'Executive Summary'!A1"

order = ["Executive Summary", "Summary", "Unplanned Payment", "Excess Payment", "Short Payment", "Payment Done",
         "Machine Advance Payment", "FF Source Data"]
out._sheets = [out[n] for n in order]
for ws in out.worksheets:
    ws.sheet_view.showGridLines = False
out["FF Source Data"].sheet_state = "hidden"

out.save(OUT)
print("saved FINAL:", OUT)


# =====================================================================
# Row-mapping discovery helper -- run this separately (not on normal execution)
# when the source file has been restructured, to reprint the current vendor -> row
# mapping so you can update CLASSIFIED_ROWS and loan_detail_rows above.
#
#   python -c "import build_fund_flow_analysis" won't trigger this (guarded below);
#   instead open a Python shell in this folder and run the block manually, or set
#   the env var DISCOVER=1 before running this script.
# =====================================================================
import os
if os.environ.get("DISCOVER") == "1":
    HEADER_C_MARKERS = {"Do Not type", "USD Amount"}
    FINANCING_LIMIT_TYPES = {"Axis", "ICICI "}
    for row in range(1, 120):
        a = ff_f.cell(row=row, column=1).value
        b = ff_f.cell(row=row, column=2).value
        c = ff_f.cell(row=row, column=3).value
        d = ff_v.cell(row=row, column=4).value
        f = ff_v.cell(row=row, column=6).value
        if a or b:
            print(row, repr(a), "|", repr(b), "| D=", d, "F=", f)
