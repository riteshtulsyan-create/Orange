"""Inventory Ageing — design sample builder.

Writes `Inventory\\SAMPLE LAYOUT - Stock_Ageing_Analysis.xlsx`, the proposed layout for the live
workbook: 10 cross-linked sheets, detail at LEGAL ENTITY level (Orange O Tec / Enterprises /
Colorix, each combining its HO + Noida Tally file), with a group Summary over all three.

Design decisions locked with the user 2026-08-01:
  - NO provision / write-down modelling. Ageing only.
  - NO "unaged" bucket. Stock with no purchase layer in the file is dated at the START OF THE
    PERIOD THE FILE MAKES AVAILABLE (e.g. 01-Apr-2025) and labelled "Opening" in the adjacent
    Receipt Source column. Every item therefore lands in a real bucket, and the Source column
    shows whether that date was traced to a purchase or assumed from the file's opening.

ALL FIGURES ARE ILLUSTRATIVE PLACEHOLDERS. Nothing here is fetched from Tally. This script exists
so the layout can be reviewed and changed before the live extract is built — edit it, re-run it,
and the sample refreshes in place.

Run:
    python "Stock Report\\.claude\\skills\\inventory-ageing\\scripts\\build_sample_layout.py"
"""
from datetime import date
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def project_root() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        if (parent / ".env").is_file():
            return parent
    return here.parents[5]


OUTDIR = project_root() / "Inventory"
OUTDIR.mkdir(parents=True, exist_ok=True)
OUT = OUTDIR / "SAMPLE LAYOUT - Stock_Ageing_Analysis.xlsx"

AS_AT = date(2026, 8, 1)

# ------------------------------------------------------------------ style kit
BANNER_FILL = PatternFill("solid", fgColor="C00000")
BANNER_FONT = Font(bold=True, color="FFFFFF", size=11)
NAV_FILL = PatternFill("solid", fgColor="EDEDED")
LINK_FONT = Font(color="0563C1", underline="single", size=9, bold=True)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
GRP_FILL = PatternFill("solid", fgColor="2E5C8A")
TITLE_FONT = Font(bold=True, size=13, color="1F3864")
SUB_FONT = Font(bold=True, size=10, color="1F3864")
NOTE_FONT = Font(italic=True, size=9, color="595959")
BLOCK_FILL = PatternFill("solid", fgColor="D9E1F2")
OPEN_FILL = PatternFill("solid", fgColor="FFF2CC")
OPEN_FONT = Font(bold=True, color="7F6000", size=10)
BAD_FILL = PatternFill("solid", fgColor="FFC7CE"); BAD_FONT = Font(color="9C0006", bold=True)
OK_FILL = PatternFill("solid", fgColor="C6EFCE"); OK_FONT = Font(color="006100", bold=True)
WARN_FILL = PatternFill("solid", fgColor="FFEB9C"); WARN_FONT = Font(color="9C5700", bold=True)
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '#,##0;[Red]-#,##0'
MONEY2 = '#,##0.00;[Red]-#,##0.00'
QTY = '#,##0.000'
RATE = '#,##0.00'
PCT = '0.0%'
DATE = 'dd-mmm-yy'
DAYS = '#,##0'

BANNER = ("SAMPLE LAYOUT  —  ILLUSTRATIVE FIGURES ONLY  —  NOT FETCHED FROM TALLY  "
          "—  Inventory Ageing design draft v2 for approval")

SHEETS = [
    ("Summary", "Summary"),
    ("Ageing by Segment", "Segment"),
    ("Orange O Tec", "Orange O Tec"),
    ("Enterprises", "Enterprises"),
    ("Colorix", "Colorix"),
    ("Slow & Non-Moving", "Slow-Moving"),
    ("Unvalued Ageing", "Unvalued"),
    ("Exception Register", "Exceptions"),
    ("Cross Check", "Cross Check"),
    ("Scope & Method", "Scope"),
]

BAD = {"FAIL", "DEAD", "NON-MOVING", "NEGATIVE", "NO VALUE", "OVERSTATED"}
GOOD = {"OK", "PASS", "FRESH", "MOVING", "NONE", "PURCHASE"}
WARN = {"REVIEW", "SLOW", "CHECK", "PARTIAL", "OPENING"}

# five real buckets — no "unaged" catch-all
BUCKETS = ["0 - 90 days", "91 - 180 days", "181 - 365 days", "1 - 2 years", "Over 2 years"]

# each Tally file only makes history available from its own start date
PERIOD_FROM = {
    "Orange O Tec": date(2025, 4, 1),
    "Enterprises": date(2026, 4, 1),
    "Colorix": date(2020, 4, 1),
}

# ---------------------------------------------------------------- the hard cap
# No item can be aged beyond the history its own file holds. As at 01-Aug-2026 that is
# 487 days for Orange O Tec, 122 for Enterprises, 2,313 for Colorix — so Enterprises
# CANNOT show anything past the 91-180 bucket, however old the stock really is.
BUCKET_MAX_DAYS = [90, 180, 365, 730, 10_000]


def reachable(entity):
    """Index of the oldest bucket this entity's file can reach."""
    days = (AS_AT - PERIOD_FROM[entity]).days
    return next(i for i, hi in enumerate(BUCKET_MAX_DAYS) if days <= hi)


# entity bucket totals, each respecting its own cap (zeros beyond the reachable bucket)
ENT_BUCKETS = {
    "Orange O Tec": [12_40_00_000, 8_10_00_000, 9_60_00_000, 16_08_20_000, 0],
    "Enterprises":  [12_90_00_000, 5_84_60_000, 0, 0, 0],
    "Colorix":      [1_80_00_000, 1_02_00_000, 1_28_00_000, 1_10_00_000, 1_22_10_000],
}
ENT_OPENING = {"Orange O Tec": 9_82_00_000, "Enterprises": 3_86_00_000, "Colorix": 94_00_000}
ENT_TOTAL = {e: sum(b) for e, b in ENT_BUCKETS.items()}

for _e, _b in ENT_BUCKETS.items():
    _cap = reachable(_e)
    assert all(v == 0 for v in _b[_cap + 1:]), \
        f"{_e}: value sits beyond bucket {_cap}, which its file cannot reach"

SEGMENTS = ["MACHINERY & MACHINERY PARTS", "MACHINERY HEAD", "PRINTING INK", "CHEMICALS",
            "SPARE PARTS", "PAPER ROLL", "SOFTWARE", "(no group assigned)"]

# closing value by segment x entity; the last segment is the balancing residual
SEG_ENTITY = {
    "Orange O Tec": [21_40_00_000, 8_90_00_000, 9_80_00_000, 2_10_00_000,
                     3_20_00_000, 60_00_000, 18_20_000, None],
    "Enterprises":  [6_10_00_000, 2_40_00_000, 6_20_00_000, 1_40_00_000,
                     1_90_00_000, 48_00_000, 6_60_000, None],
    "Colorix":      [3_80_00_000, 1_10_00_000, 90_00_000, 22_00_000,
                     30_00_000, 8_00_000, 2_10_000, None],
}
for _e, _v in SEG_ENTITY.items():
    _v[-1] = ENT_TOTAL[_e] - sum(x for x in _v[:-1])
    assert _v[-1] >= 0, f"{_e}: segment splits exceed the entity total"

# how each segment tends to age, before fitting to the real margins
SEG_SKEW = {
    "MACHINERY & MACHINERY PARTS": [0.6, 0.8, 1.0, 1.6, 1.8],
    "MACHINERY HEAD":              [0.7, 0.9, 1.1, 1.5, 1.6],
    "PRINTING INK":                [1.6, 1.3, 0.9, 0.5, 0.3],
    "CHEMICALS":                   [1.5, 1.2, 0.9, 0.6, 0.4],
    "SPARE PARTS":                 [0.9, 1.0, 1.1, 1.2, 1.3],
    "PAPER ROLL":                  [1.2, 1.1, 1.0, 0.8, 0.7],
    "SOFTWARE":                    [0.3, 0.4, 0.6, 1.5, 2.5],
    "(no group assigned)":         [1.0, 1.0, 1.0, 1.0, 1.0],
}


def fit_segment_buckets():
    """Iterative proportional fitting: give each segment a believable ageing shape while
    keeping BOTH margins exact — row = segment total, column = group bucket total."""
    row_tot = [sum(SEG_ENTITY[e][s] for e in SEG_ENTITY) for s in range(len(SEGMENTS))]
    col_tot = [sum(ENT_BUCKETS[e][b] for e in ENT_BUCKETS) for b in range(5)]
    m = [[row_tot[s] * SEG_SKEW[seg][b] / sum(SEG_SKEW[seg]) for b in range(5)]
         for s, seg in enumerate(SEGMENTS)]
    for _ in range(300):
        for b in range(5):                                  # scale columns
            cur = sum(m[s][b] for s in range(len(SEGMENTS)))
            if cur:
                f = col_tot[b] / cur
                for s in range(len(SEGMENTS)):
                    m[s][b] *= f
        for s in range(len(SEGMENTS)):                      # scale rows
            cur = sum(m[s])
            if cur:
                f = row_tot[s] / cur
                for b in range(5):
                    m[s][b] *= f
    return m, row_tot, col_tot


SEG_BUCKETS, SEG_ROW_TOTAL, GRP_BUCKETS = fit_segment_buckets()


def link(ws, row, col, text, target_sheet, target_cell="A1"):
    c = ws.cell(row=row, column=col, value=text)
    c.hyperlink = f"#'{target_sheet}'!{target_cell}"
    c.font = LINK_FONT
    return c


def frame(ws, ncols, current):
    """Red banner + nav bar on every sheet. Returns the next free row."""
    ws["A1"] = BANNER
    ws["A1"].fill, ws["A1"].font = BANNER_FILL, BANNER_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(ncols, 12))
    ws.row_dimensions[1].height = 20

    ws.cell(row=2, column=1, value="Go to:").font = Font(bold=True, size=9)
    col = 2
    for name, label in SHEETS:
        if name == current:
            c = ws.cell(row=2, column=col, value=f"[{label}]")
            c.font = Font(bold=True, size=9, color="1F3864")
        else:
            link(ws, 2, col, label, name)
        ws.cell(row=2, column=col).alignment = Alignment(horizontal="center")
        col += 1
    for c in range(1, col):
        ws.cell(row=2, column=c).fill = NAV_FILL
    ws.row_dimensions[2].height = 16
    return 4


def head(ws, row, headers, widths=None, group_row=None):
    """Styled header row, optionally with a merged grouping band above it."""
    if group_row:
        for label, c1, c2 in group_row:
            ws.cell(row=row - 1, column=c1, value=label)
            if c2 > c1:
                ws.merge_cells(start_row=row - 1, start_column=c1,
                               end_row=row - 1, end_column=c2)
            for c in range(c1, c2 + 1):
                cell = ws.cell(row=row - 1, column=c)
                cell.fill, cell.font = GRP_FILL, HDR_FONT
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = BORDER
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=i, value=h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = 36
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def paint(ws, r1, r2, ncols, formats, flag_cols=(), source_col=None):
    """Borders, number formats, flag colouring, total-row shading."""
    for ri in range(r1, r2 + 1):
        lbl = str(ws.cell(row=ri, column=1).value or "").upper()
        is_tot = lbl.startswith(("TOTAL", "GRAND", "GROUP", "ENTITY", "ALL ", "COMBINED"))
        for ci in range(1, ncols + 1):
            cell = ws.cell(row=ri, column=ci)
            cell.border = BORDER
            if ci in formats:
                cell.number_format = formats[ci]
            if is_tot:
                cell.fill = TOT_FILL
                cell.font = Font(bold=True, size=10)
            if ci in flag_cols and not is_tot:
                v = str(cell.value or "").upper()
                if v in BAD:
                    cell.fill, cell.font = BAD_FILL, BAD_FONT
                elif v in GOOD:
                    cell.fill, cell.font = OK_FILL, OK_FONT
                elif v in WARN:
                    cell.fill, cell.font = WARN_FILL, WARN_FONT
        # an Opening-dated row gets its date cell highlighted too
        if source_col and not is_tot:
            if str(ws.cell(row=ri, column=source_col).value or "").upper() == "OPENING":
                ws.cell(row=ri, column=source_col - 1).fill = OPEN_FILL
                ws.cell(row=ri, column=source_col - 1).font = OPEN_FONT


wb = Workbook()
wb.remove(wb.active)
for name, _ in SHEETS:
    wb.create_sheet(name)

# =================================================================== 1. Summary
ws = wb["Summary"]
r = frame(ws, 13, "Summary")
ws.cell(row=r, column=1, value="INVENTORY AGEING ANALYSIS — GROUP SUMMARY").font = TITLE_FONT
r += 1
ws.cell(row=r, column=1,
        value="As at 01-Aug-2026  ·  aged by date of receipt, FIFO (newest receipts are what "
              "remains)  ·  3 legal entities from 5 Tally files").font = NOTE_FONT
r += 1
ws.cell(row=r, column=1,
        value="Stock with no purchase layer in the file is dated at that file's OPENING and "
              "flagged as such — see the last two columns. No provision is modelled.").font = NOTE_FONT
r += 2

hdr = ["Legal Entity", "Items with stock", "Closing Stock Value",
       "0 - 90 days", "91 - 180 days", "181 - 365 days", "1 - 2 years", "Over 2 years",
       "% older than 1 year", "Value dated from Opening", "% resting on Opening date",
       "History available from", "open →"]
head(ws, r, hdr, widths=[24, 13, 17, 15, 15, 15, 15, 15, 12, 16, 13, 15, 10],
     group_row=[("Ageing buckets — value", 4, 8), ("How much rests on an assumed date", 10, 12)])
hr = r
rows = [("Orange O Tec", 3182), ("Enterprises", 1421), ("Colorix", 613)]
r += 1
first = r
for name, items in rows:
    buckets, opening = ENT_BUCKETS[name], ENT_OPENING[name]
    ws.cell(row=r, column=1, value=name)
    ws.cell(row=r, column=2, value=items)
    for i, v in enumerate(buckets, start=4):
        ws.cell(row=r, column=i, value=v)
    ws.cell(row=r, column=3, value=f"=SUM(D{r}:H{r})")
    ws.cell(row=r, column=9, value=f"=SUM(G{r}:H{r})/C{r}")
    ws.cell(row=r, column=10, value=opening)
    ws.cell(row=r, column=11, value=f"=J{r}/C{r}")
    ws.cell(row=r, column=12, value=PERIOD_FROM[name])
    link(ws, r, 13, "open →", name)
    r += 1
ws.cell(row=r, column=1, value="GROUP TOTAL")
for c in [2, 3, 4, 5, 6, 7, 8, 10]:
    L = get_column_letter(c)
    ws.cell(row=r, column=c, value=f"=SUM({L}{first}:{L}{r-1})")
ws.cell(row=r, column=9, value=f"=SUM(G{r}:H{r})/C{r}")
ws.cell(row=r, column=11, value=f"=J{r}/C{r}")
paint(ws, hr + 1, r, 13,
      {2: DAYS, 3: MONEY, 4: MONEY, 5: MONEY, 6: MONEY, 7: MONEY, 8: MONEY,
       9: PCT, 10: MONEY, 11: PCT, 12: DATE})
for ri in range(hr + 1, r):
    ws.cell(row=ri, column=12).fill = OPEN_FILL
    ws.cell(row=ri, column=12).font = OPEN_FONT
r += 2

ws.cell(row=r, column=1, value="READ THIS BEFORE READING THE BUCKETS").font = SUB_FONT
r += 1
head(ws, r, ["Point", "Illustrative answer", "Where it is worked", "→"])
hr2 = r
notes = [
    ("The three files start on three different dates.",
     "Colorix has 6 years of history; Orange O Tec has from 01-Apr-25; Enterprises only from "
     "01-Apr-26. The same physical item ages differently in each simply because of the file.",
     "Scope & Method", "Scope & Method"),
    ("Opening-dated stock cannot be older than its file.",
     "Enterprises stock dated 01-Apr-26 lands in 91-180 days and LOOKS FRESH. It may be years "
     "old. Treat the Enterprises bucket profile as a floor on age, not a measurement.",
     "Scope & Method", "Scope & Method"),
    ("Some stock has quantity but no value at all.",
     "₹0 carried against real quantity — the ageing cannot price it. Quantified separately.",
     "Unvalued Ageing", "Unvalued Ageing"),
    ("Purchase age is not the same question as movement.",
     "Stock bought recently but never issued is the real risk. Aged by last outward instead.",
     "Slow & Non-Moving", "Slow & Non-Moving"),
    ("Does the ageing tie to the balance sheet?",
     "Buckets sum to Tally closing stock, to the rupee.",
     "Cross Check", "Cross Check"),
]
r += 1
for q, a, where, target in notes:
    ws.cell(row=r, column=1, value=q)
    ws.cell(row=r, column=2, value=a)
    ws.cell(row=r, column=3, value=where)
    link(ws, r, 4, "→", target)
    for ci in (1, 2):
        ws.cell(row=r, column=ci).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30
    r += 1
paint(ws, hr2 + 1, r - 1, 4, {})
ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 76
ws.column_dimensions["C"].width = 20
ws.freeze_panes = "A6"

# =========================================================== 2. Ageing by Segment
ws = wb["Ageing by Segment"]
r = frame(ws, 13, "Ageing by Segment")
ws.cell(row=r, column=1, value="AGEING BY STOCK SEGMENT — ALL THREE ENTITIES").font = TITLE_FONT
r += 1
ws.cell(row=r, column=1,
        value="Main stock group from the Tally group tree. An item's own Parent is only its "
              "immediate group, so the full parent chain is walked to the root.").font = NOTE_FONT
r += 2
hdr = ["Main Category", "Orange O Tec", "Enterprises", "Colorix", "Total Value",
       "0 - 90", "91 - 180", "181 - 365", "1 - 2 yrs", "> 2 yrs",
       "% > 1 year", "Of which Opening-dated", "Flag"]
head(ws, r, hdr, widths=[34, 15, 15, 14, 16, 14, 14, 14, 14, 14, 11, 16, 13],
     group_row=[("Closing stock value by entity", 2, 4), ("Ageing buckets (value)", 6, 10)])
hr = r
SEG_FLAG = ["REVIEW", "REVIEW", "OK", "OK", "SLOW", "OK", "DEAD", "CHECK"]
r += 1
first = r
for s, name in enumerate(SEGMENTS):
    oot = SEG_ENTITY["Orange O Tec"][s]
    ent = SEG_ENTITY["Enterprises"][s]
    col = SEG_ENTITY["Colorix"][s]
    b = SEG_BUCKETS[s]
    # opening-dated value follows each entity's own segment share
    opening = sum(ENT_OPENING[e] * SEG_ENTITY[e][s] / ENT_TOTAL[e] for e in ENT_TOTAL)
    flag = SEG_FLAG[s]
    ws.cell(row=r, column=1, value=name)
    ws.cell(row=r, column=2, value=oot)
    ws.cell(row=r, column=3, value=ent)
    ws.cell(row=r, column=4, value=col)
    ws.cell(row=r, column=5, value=f"=SUM(B{r}:D{r})")
    for i, v in enumerate(b, start=6):
        ws.cell(row=r, column=i, value=v)
    ws.cell(row=r, column=11, value=f"=IF(E{r}=0,0,SUM(I{r}:J{r})/E{r})")
    ws.cell(row=r, column=12, value=opening)
    ws.cell(row=r, column=13, value=flag)
    r += 1
ws.cell(row=r, column=1, value="TOTAL")
for c in list(range(2, 11)) + [12]:
    L = get_column_letter(c)
    ws.cell(row=r, column=c, value=f"=SUM({L}{first}:{L}{r-1})")
ws.cell(row=r, column=11, value=f"=SUM(I{r}:J{r})/E{r}")
paint(ws, hr + 1, r, 13,
      {2: MONEY, 3: MONEY, 4: MONEY, 5: MONEY, 6: MONEY, 7: MONEY, 8: MONEY, 9: MONEY,
       10: MONEY, 11: PCT, 12: MONEY}, flag_cols={13})
ws.freeze_panes = "B7"
ws.auto_filter.ref = f"A{hr}:M{r-1}"

# ================================================= 3-5. the three entity sheets
DET_HDR = ["Rank", "Unit", "Main Category", "Sub Category", "Item Name", "UOM",
           "Closing Qty", "Closing Value", "Avg Rate",
           "0-90 Qty", "0-90 Value", "91-180 Qty", "91-180 Value",
           "181-365 Qty", "181-365 Value", "1-2 yr Qty", "1-2 yr Value",
           ">2 yr Qty", ">2 yr Value",
           "Oldest Receipt", "Receipt Source", "Age of Oldest (days)",
           "Last Inward", "Last Outward", "Days Since Movement", "Months of Cover",
           "Flag", "See Also"]
DET_W = [6, 8, 24, 22, 34, 7, 12, 14, 10,
         11, 13, 11, 13, 11, 13, 11, 13, 11, 13,
         12, 13, 11, 11, 11, 12, 11, 13, 16]
DET_GRP = [("Closing stock", 7, 9), ("0 - 90 days", 10, 11), ("91 - 180 days", 12, 13),
           ("181 - 365 days", 14, 15), ("1 - 2 years", 16, 17), ("Over 2 years", 18, 19),
           ("Ageing basis", 20, 22), ("Movement evidence", 23, 26)]
DET_FMT = {7: QTY, 8: MONEY2, 9: RATE, 10: QTY, 11: MONEY2, 12: QTY, 13: MONEY2,
           14: QTY, 15: MONEY2, 16: QTY, 17: MONEY2, 18: QTY, 19: MONEY2,
           20: DATE, 22: DAYS, 23: DATE, 24: DATE, 25: DAYS, 26: RATE}
SRC_COL = 21

ENTITIES = {
    "Orange O Tec": {
        "units": ("HO", "Noida Branch"),
        "roll": [
            ("Items with stock", 2190, 992, DAYS),
            ("Closing stock value", 33_10_00_000, 13_08_20_000, MONEY),
            ("Value over 1 year old", 11_40_00_000, 4_68_20_000, MONEY),
            ("Value dated from Opening (01-Apr-25)", 6_90_00_000, 2_92_00_000, MONEY),
            ("Value not moved in 12 months", 10_80_00_000, 4_20_00_000, MONEY),
            ("Items flagged", 214, 96, DAYS),
        ],
        "detail": [
            (1, "HO", "MACHINERY & MACHINERY PARTS", "MACHINE PROCUREMENT",
             "ATEXCO VEGA 3000 8-HEAD", "PCS", 2.000, None, 34_50_000,
             0, 0, 0, 0, 0, 0, 2.000, 69_00_000, 0, 0,
             date(2025, 4, 1), "Opening", 487, None, date(2024, 6, 2), 790, 0.0,
             "NON-MOVING", "Slow & Non-Moving"),
            (2, "HO", "MACHINERY HEAD", "KYOCERA",
             "KYOCERA KJ4B-QA06 HEAD", "PCS", 14.000, None, 1_86_000,
             2.000, 3_72_000, 3.000, 5_58_000, 4.000, 7_44_000, 5.000, 9_30_000, 0, 0,
             date(2025, 6, 12), "Purchase", 415, date(2026, 5, 18), date(2026, 6, 27), 35, 4.2,
             "SLOW", "Slow & Non-Moving"),
            (3, "Noida", "PRINTING INK", "FLOTEC INK",
             "REACTIVE INK H-SERIES BLACK", "KGS", 4820.000, None, 393.58,
             1800.000, 7_08_444, 1200.000, 4_72_296, 900.000, 3_54_222,
             920.000, 3_62_094, 0, 0,
             date(2025, 5, 20), "Purchase", 438, date(2026, 7, 2), date(2026, 7, 24), 8, 2.8,
             "MOVING", None),
            (4, "HO", "PRINTING INK", "ATEXCO REACTIVE",
             "ATEXCO REACTIVE INK GREY-D", "KGS", 20.000, None, 0,
             0, 0, 0, 0, 0, 0, 20.000, 0, 0, 0,
             date(2025, 4, 1), "Opening", 487, None, None, None, 0.0,
             "DEAD", "Exception Register"),
            (5, "HO", "PRINTING INK", "ACID INK",
             "ACID INK BLACK-D", "KGS", 10.000, None, 0,
             0, 0, 0, 0, 0, 0, 10.000, 0, 0, 0,
             date(2025, 4, 1), "Opening", 487, None, None, None, 0.0,
             "NO VALUE", "Unvalued Ageing"),
            (6, "Noida", "SPARE PARTS", "OTHER SPARE PARTS",
             "DAMPER ASSY DX5", "PCS", 340.000, None, 1_240.00,
             40.000, 49_600, 60.000, 74_400, 80.000, 99_200, 160.000, 1_98_400, 0, 0,
             date(2025, 4, 18), "Purchase", 470, date(2026, 6, 11), date(2026, 7, 19), 13, 6.1,
             "SLOW", None),
        ],
    },
    "Enterprises": {
        "units": ("HO", "Noida Branch"),
        "roll": [
            ("Items with stock", 890, 531, DAYS),
            ("Closing stock value", 12_40_00_000, 6_34_60_000, MONEY),
            ("Value over 1 year old", 0, 0, MONEY),
            ("Value dated from Opening (01-Apr-26)", 2_46_00_000, 1_40_00_000, MONEY),
            ("Value not moved since the file opened", 4_60_00_000, 2_40_00_000, MONEY),
            ("Items flagged", 88, 52, DAYS),
        ],
        "detail": [
            (1, "HO", "PRINTING INK", "SUBLIMATION INK",
             "SUBLI INK MAGENTA 5L", "LTR", 1240.000, None, 480.00,
             740.000, 3_55_200, 500.000, 2_40_000, 0, 0, 0, 0, 0, 0,
             date(2026, 4, 19), "Purchase", 104, date(2026, 6, 30), date(2026, 7, 21), 11, 3.4,
             "MOVING", None),
            (2, "Noida", "CHEMICALS", "PRE-TREATMENT",
             "50X50 PADDING CHEMICAL-D", "KGS", 100.000, None, 0,
             0, 0, 100.000, 0, 0, 0, 0, 0, 0, 0,
             date(2026, 4, 1), "Opening", 122, None, None, None, 0.0,
             "DEAD", "Exception Register"),
            (3, "HO", "SPARE PARTS", "BELTS & ROLLERS",
             "CONVEYOR BELT 1.8M", "NOS", 22.000, None, 18_400.00,
             14.000, 2_57_600, 8.000, 1_47_200, 0, 0, 0, 0, 0, 0,
             date(2026, 4, 25), "Purchase", 98, date(2026, 5, 27), date(2026, 7, 8), 24, 5.5,
             "SLOW", None),
        ],
    },
    "Colorix": {
        "units": ("Colorix",),
        "roll": [
            ("Items with stock", 613, None, DAYS),
            ("Closing stock value", 6_42_10_000, None, MONEY),
            ("Value over 1 year old", 2_32_10_000, None, MONEY),
            ("Value dated from Opening (01-Apr-20)", 94_00_000, None, MONEY),
            ("Value not moved in 12 months", 2_10_00_000, None, MONEY),
            ("Items flagged", 41, None, DAYS),
        ],
        "detail": [
            (1, "Colorix", "MACHINERY & MACHINERY PARTS", "MACHINE PROCUREMENT",
             "COLORIX DTF 60 PRINTER", "PCS", 3.000, None, 8_20_000.00,
             1.000, 8_20_000, 0, 0, 1.000, 8_20_000, 1.000, 8_20_000, 0, 0,
             date(2024, 12, 3), "Purchase", 606, date(2026, 3, 11), date(2026, 6, 19), 43, 3.0,
             "MOVING", None),
            (2, "Colorix", "PRINTING INK", "DTF INK",
             "DTF WHITE INK 1L", "LTR", 640.000, None, 720.00,
             200.000, 1_44_000, 160.000, 1_15_200, 120.000, 86_400,
             0, 0, 160.000, 1_15_200,
             date(2024, 6, 14), "Purchase", 778, date(2026, 7, 5), date(2026, 7, 28), 4, 2.2,
             "MOVING", None),
            (3, "Colorix", "SOFTWARE", "DONGLE",
             "444-027 SOFT DONGLE", "PCS", 2.000, None, 10_000.00,
             0, 0, 0, 0, 0, 0, 0, 0, 2.000, 20_000,
             date(2020, 4, 1), "Opening", 2313, None, None, None, 0.0,
             "NON-MOVING", "Slow & Non-Moving"),
        ],
    },
}

for ent, cfg in ENTITIES.items():
    ws = wb[ent]
    r = frame(ws, 28, ent)
    ws.cell(row=r, column=1, value=f"{ent.upper()} — INVENTORY AGEING").font = TITLE_FONT
    r += 1
    ws.cell(row=r, column=1,
            value=f"History available from {PERIOD_FROM[ent]:%d-%b-%Y}. Stock with no purchase "
                  f"layer in the file is dated {PERIOD_FROM[ent]:%d-%b-%Y} and marked 'Opening' "
                  f"in the Receipt Source column — it cannot be shown as older than the file "
                  f"itself.").font = NOTE_FONT
    r += 1
    if len(cfg["units"]) > 1:
        ws.cell(row=r, column=1,
                value="HO and Noida are the SAME legal company. Both roll into the Combined "
                      "column; the Unit column keeps them separable.").font = NOTE_FONT
        r += 1
    r += 1

    # --- roll-up block
    ws.cell(row=r, column=1, value="ENTITY ROLL-UP").font = SUB_FONT
    r += 1
    units = cfg["units"]
    cols = ["Measure", units[0]] + ([units[1]] if len(units) > 1 else []) + ["Entity Combined"]
    head(ws, r, cols)
    hr = r
    r += 1
    for label, a, b, fmt in cfg["roll"]:
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=a)
        if len(units) > 1:
            ws.cell(row=r, column=3, value=b)
            ws.cell(row=r, column=4, value=f"=B{r}+C{r}")
        else:
            ws.cell(row=r, column=3, value=f"=B{r}")
        for c in range(2, len(cols) + 1):
            ws.cell(row=r, column=c).number_format = fmt
        r += 1
    paint(ws, hr + 1, r - 1, len(cols), {})
    for i, w in enumerate([38, 18, 18, 18][:len(cols)], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    r += 2

    # --- item detail
    ws.cell(row=r, column=1,
            value="ITEM DETAIL — ranked by value over one year old").font = SUB_FONT
    r += 1
    ws.cell(row=r, column=1,
            value="Closing Value is a live formula summing the five bucket values. A yellow "
                  "Oldest Receipt date means the date was assumed from the file's opening, not "
                  "traced to a purchase.").font = NOTE_FONT
    r += 2
    head(ws, r, DET_HDR, widths=DET_W, group_row=DET_GRP)
    hr = r
    r += 1
    first = r
    for row in cfg["detail"]:
        for i, v in enumerate(row, start=1):
            if v is None or i == 8:
                continue
            if i == 28:
                link(ws, r, 28, v, v)
                continue
            ws.cell(row=r, column=i, value=v)
        # Closing Value = sum of the five bucket values
        ws.cell(row=r, column=8, value=f"=K{r}+M{r}+O{r}+Q{r}+S{r}")
        r += 1
    ws.cell(row=r, column=1, value="TOTAL")
    for c in [7, 8] + list(range(10, 20)):
        L = get_column_letter(c)
        ws.cell(row=r, column=c, value=f"=SUM({L}{first}:{L}{r-1})")
    paint(ws, hr + 1, r, 28, DET_FMT, flag_cols={SRC_COL, 27}, source_col=SRC_COL)
    ws.auto_filter.ref = f"A{hr}:AB{r-1}"
    ws.freeze_panes = ws.cell(row=hr + 1, column=6).coordinate

# ==================================================== 6. Slow & Non-Moving
ws = wb["Slow & Non-Moving"]
r = frame(ws, 14, "Slow & Non-Moving")
ws.cell(row=r, column=1, value="SLOW-MOVING AND NON-MOVING STOCK").font = TITLE_FONT
r += 1
ws.cell(row=r, column=1,
        value="Purchase age and movement are different questions. An item bought 400 days ago "
              "that sells every month is fine; an item bought last month that has never issued "
              "is not. This sheet ages by LAST OUTWARD MOVEMENT, not by receipt date — and it is "
              "unaffected by how far back the Tally file goes.").font = NOTE_FONT
r += 2
hdr = ["Entity", "Unit", "Main Category", "Item Name", "UOM", "Closing Qty", "Closing Value",
       "Last Outward", "Days Since Movement", "Issued in last 12m (qty)",
       "Avg Monthly Consumption", "Months of Cover", "Status", "See Also"]
head(ws, r, hdr, widths=[16, 9, 26, 34, 7, 13, 15, 12, 13, 15, 15, 12, 15, 16])
hr = r
data = [
    ("Orange O Tec", "HO", "MACHINERY & MACHINERY PARTS", "ATEXCO VEGA 3000 8-HEAD", "PCS",
     2.000, 69_00_000, date(2024, 6, 2), 790, 0.000, 0.000, None, "NON-MOVING", "Orange O Tec"),
    ("Colorix", "Colorix", "SOFTWARE", "444-027 SOFT DONGLE", "PCS",
     2.000, 20_000, None, 2313, 0.000, 0.000, None, "NON-MOVING", "Colorix"),
    ("Orange O Tec", "HO", "PRINTING INK", "ATEXCO REACTIVE INK GREY-D", "KGS",
     20.000, 0, None, None, 0.000, 0.000, None, "DEAD", "Exception Register"),
    ("Enterprises", "Noida", "CHEMICALS", "50X50 PADDING CHEMICAL-D", "KGS",
     100.000, 0, None, None, 0.000, 0.000, None, "DEAD", "Exception Register"),
    ("Orange O Tec", "Noida", "SPARE PARTS", "DAMPER ASSY DX5", "PCS",
     340.000, 4_21_600, date(2026, 7, 19), 13, 668.000, 55.667, 6.1, "SLOW", None),
    ("Orange O Tec", "HO", "MACHINERY HEAD", "KYOCERA KJ4B-QA06 HEAD", "PCS",
     14.000, 26_04_000, date(2026, 6, 27), 35, 40.000, 3.333, 4.2, "SLOW", None),
    ("Enterprises", "HO", "SPARE PARTS", "CONVEYOR BELT 1.8M", "NOS",
     22.000, 4_04_800, date(2026, 7, 8), 24, 48.000, 4.000, 5.5, "SLOW", None),
    ("Colorix", "Colorix", "PRINTING INK", "DTF WHITE INK 1L", "LTR",
     640.000, 4_60_800, date(2026, 7, 28), 4, 3480.000, 290.000, 2.2, "MOVING", None),
]
r += 1
first = r
for row in data:
    for i, v in enumerate(row, start=1):
        if v is None:
            continue
        if i == 14:
            link(ws, r, 14, v, v)
            continue
        ws.cell(row=r, column=i, value=v)
    r += 1
ws.cell(row=r, column=1, value="TOTAL")
for c in (6, 7, 10):
    L = get_column_letter(c)
    ws.cell(row=r, column=c, value=f"=SUM({L}{first}:{L}{r-1})")
paint(ws, hr + 1, r, 14,
      {6: QTY, 7: MONEY2, 8: DATE, 9: DAYS, 10: QTY, 11: QTY, 12: RATE}, flag_cols={13})
ws.auto_filter.ref = f"A{hr}:N{r-1}"
ws.freeze_panes = ws.cell(row=hr + 1, column=5).coordinate
r += 2
ws.cell(row=r, column=1,
        value="A blank Last Outward means the item has never issued in the period the file "
              "covers — the strongest dead-stock signal in the pack.").font = NOTE_FONT

# ==================================================== 7. Unvalued Ageing
ws = wb["Unvalued Ageing"]
r = frame(ws, 11, "Unvalued Ageing")
ws.cell(row=r, column=1,
        value="STOCK CARRYING QUANTITY BUT NO VALUE").font = TITLE_FONT
r += 1
ws.cell(row=r, column=1,
        value="Straight off your Tally screen: ACID INK BLACK-D shows 10.000 KGS in the >180 day "
              "bucket with NO value — as do ACID INK BLUE / CYAN / MAGENTA / ORANGE / YELLOW, "
              "50X50 PADDING CHEMICAL-D and BLACK FOILJET INK-D. 'Valued based on actual "
              "purchase' means Tally found no purchase to price them from.").font = NOTE_FONT
r += 1
ws.cell(row=r, column=1,
        value="This is a DIFFERENT problem from an unknown date. The date is now solved by dating "
              "such stock at the file's opening. The missing VALUE is not — the ageing simply "
              "cannot price it, so these quantities sit in a bucket contributing ₹0.").font = NOTE_FONT
r += 2
hdr = ["Entity", "Unit", "Item Name", "UOM", "Closing Qty", "Tally Carrying Value",
       "Bucket it falls in", "Receipt Source", "Why there is no value", "Impact", "Action"]
head(ws, r, hdr, widths=[15, 9, 34, 7, 13, 16, 15, 13, 42, 13, 36])
hr = r
data = [
    ("Orange O Tec", "HO", "ACID INK BLACK-D", "KGS", 10.000, 0, "1 - 2 years", "Opening",
     "Opening stock brought in without a value", "NO VALUE",
     "Confirm it physically exists; write off if not"),
    ("Orange O Tec", "HO", "ATEXCO REACTIVE INK GREY-D", "KGS", 20.000, 0, "1 - 2 years",
     "Opening", "Opening stock brought in without a value", "NO VALUE",
     "Value at last known purchase rate, or de-activate"),
    ("Enterprises", "Noida", "50X50 PADDING CHEMICAL-D", "KGS", 100.000, 0, "91 - 180 days",
     "Opening", "Received by stock journal at nil rate", "NO VALUE",
     "Trace the journal; restate at the transferring unit's cost"),
    ("Orange O Tec", "Noida", "ANTELOS ORANGE R-RCD XD-DEAD", "KGS", 5.000, 0, "1 - 2 years",
     "Opening", "Free receipt / sample — no purchase invoice", "NO VALUE",
     "Confirm nil cost is correct, then de-activate the item"),
    ("Orange O Tec", "HO", "BLACK FOILJET INK-D", "KGS", 80.000, 0, "1 - 2 years", "Opening",
     "Opening stock brought in without a value", "NO VALUE",
     "Value at last known purchase rate, or write off"),
    ("Colorix", "Colorix", "CLEANER", "KGS", 10.000, 0, "Over 2 years", "Opening",
     "Opening stock brought in without a value", "NO VALUE",
     "Value at last known purchase rate, or write off"),
]
r += 1
first = r
for row in data:
    for i, v in enumerate(row, start=1):
        ws.cell(row=r, column=i, value=v)
    r += 1
ws.cell(row=r, column=1, value="TOTAL — quantity carrying no value")
for c in (5, 6):
    L = get_column_letter(c)
    ws.cell(row=r, column=c, value=f"=SUM({L}{first}:{L}{r-1})")
paint(ws, hr + 1, r, 11, {5: QTY, 6: MONEY2}, flag_cols={8, 10})
ws.auto_filter.ref = f"A{hr}:K{r-1}"
r += 2
ws.cell(row=r, column=1,
        value="Coverage test — the figure that decides how much of the ageing is real:").font = SUB_FONT
r += 1
for label, v, fmt in [
    ("Tally closing stock value (all 3 entities)", 71_34_90_000, MONEY),
    ("Value that could be priced into a bucket", 64_61_00_000, MONEY),
    ("Quantity present but carrying NO value", 6_73_90_000, MONEY),
    ("Ageing value coverage", 0.9055, PCT),
]:
    ws.cell(row=r, column=1, value=label)
    c = ws.cell(row=r, column=3, value=v)
    c.number_format = fmt
    c.font = Font(bold=True)
    ws.cell(row=r, column=1).fill = BLOCK_FILL
    ws.cell(row=r, column=3).fill = BLOCK_FILL
    r += 1

# ==================================================== 8. Exception Register
ws = wb["Exception Register"]
r = frame(ws, 11, "Exception Register")
ws.cell(row=r, column=1, value="EXCEPTION REGISTER").font = TITLE_FONT
r += 1
ws.cell(row=r, column=1,
        value="Data defects the ageing exposes. These are accounting issues, not report bugs — "
              "they are reported, never suppressed.").font = NOTE_FONT
r += 2
hdr = ["#", "Entity", "Unit", "Item Name", "Exception Type", "Qty", "Value",
       "Why Flagged", "Est. ₹ Impact", "Suggested Action", "See Also"]
head(ws, r, hdr, widths=[5, 15, 9, 32, 16, 12, 14, 48, 14, 38, 16])
hr = r
data = [
    (1, "Orange O Tec", "HO", "ATEXCO REACTIVE INK GREY-D", "DEAD", 20.000, 0,
     "Name marked '-D'. No outward movement at all in the period the file covers.",
     0, "De-activate the master so it stops appearing in every stock report",
     "Slow & Non-Moving"),
    (2, "Enterprises", "Noida", "50X50 PADDING CHEMICAL-D", "DEAD", 100.000, 0,
     "Marked '-D', 100 KGS held, never issued. Shows as 91-180 days ONLY because the "
     "Enterprises file starts 01-Apr-26 — it is almost certainly far older.",
     0, "Physical verification, then write off", "Slow & Non-Moving"),
    (3, "Orange O Tec", "Noida", "ANTELOS ORANGE R-RCD XD-DEAD", "DEAD", 5.000, 0,
     "Name explicitly says DEAD but the item is still live in the master",
     0, "De-activate", "Unvalued Ageing"),
    (4, "Orange O Tec", "HO", "SUZHOU TRANSFER PAPER 1.6M", "NEGATIVE", -240.000, -1_31_000,
     "Negative closing quantity — issues booked without a matching receipt. Cannot be aged.",
     1_31_000, "Locate the missing purchase / receipt note and post it", None),
    (5, "Enterprises", "HO", "SPARE KIT — MISC", "REVIEW", 0.000, 42_000,
     "Value with zero quantity — stranded value, nothing on hand",
     42_000, "Write the residual value off", None),
    (6, "Colorix", "Colorix", "444-027 SOFT DONGLE", "NON-MOVING", 2.000, 20_000,
     "Dated 01-Apr-20 opening, never issued in 6 years of file history",
     20_000, "Write off — the strongest evidence in the pack", "Slow & Non-Moving"),
    (7, "Orange O Tec", "HO", "(various — 38 items)", "CHECK", 0.000, 0,
     "Stock items with no stock group assigned — cannot be segmented",
     0, "Assign a stock group in the item master", "Ageing by Segment"),
]
r += 1
first = r
for row in data:
    for i, v in enumerate(row, start=1):
        if v is None:
            continue
        if i == 11:
            link(ws, r, 11, v, v)
            continue
        ws.cell(row=r, column=i, value=v)
    ws.cell(row=r, column=8).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30
    r += 1
ws.cell(row=r, column=1, value="TOTAL")
for c in (6, 7, 9):
    L = get_column_letter(c)
    ws.cell(row=r, column=c, value=f"=SUM({L}{first}:{L}{r-1})")
paint(ws, hr + 1, r, 11, {6: QTY, 7: MONEY2, 9: MONEY2}, flag_cols={5})
ws.auto_filter.ref = f"A{hr}:K{r-1}"

# ==================================================== 9. Cross Check
ws = wb["Cross Check"]
r = frame(ws, 9, "Cross Check")
ws.cell(row=r, column=1, value="CROSS CHECK — RUN BEFORE ANY FIGURE IS QUOTED").font = TITLE_FONT
r += 1
ws.cell(row=r, column=1,
        value="Tests 1-5 must PASS. Tests 6-8 are EXPECTED to fail where the underlying data is "
              "genuinely defective — a failure there is a finding, not something to hide.").font = NOTE_FONT
r += 2
hdr = ["#", "Check", "Scope", "Expected", "Computed", "Difference", "Status", "Note", "Sheet"]
head(ws, r, hdr, widths=[5, 46, 16, 17, 17, 14, 10, 46, 16])
hr = r
data = [
    (1, "Five buckets sum to Tally closing stock value", "All 3 entities",
     71_34_90_000, 71_34_90_000, 0, "PASS", "Ties to the rupee", "Summary"),
    (2, "Ageing total = latest Stock Summary workbook", "All 3 entities",
     71_34_90_000, 71_34_90_000, 0, "PASS",
     "Reconciled to Stock_Summary_ALL_COMPANIES", "Summary"),
    (3, "Closing stock = Balance Sheet inventory line", "All 3 entities",
     71_34_90_000, 71_34_90_000, 0, "PASS", "Agrees to PL_BS_GL_Detail", "Summary"),
    (4, "Bucket quantities sum to closing quantity", "Every item",
     None, None, 0, "PASS", "5,216 of 5,216 items tie", None),
    (5, "Entity combined = sum of its Tally files", "3 entities / 5 files",
     71_34_90_000, 71_34_90_000, 0, "PASS", "HO + Noida roll-up correct", None),
    (6, "No negative closing quantity", "Every item",
     0, 11, 11, "FAIL", "11 items negative — real data defect, cannot be aged",
     "Exception Register"),
    (7, "No value held against zero quantity", "Every item",
     0, 6, 6, "FAIL", "6 items stranded value", "Exception Register"),
    (8, "Every item has a stock group", "Every item",
     0, 38, 38, "FAIL", "38 unassigned — segment view incomplete", "Ageing by Segment"),
    (9, "Top 20 ages re-proved from raw voucher XML", "Top 20 by value",
     20, 20, 0, "PASS", "Independent recomputation, not a re-read", None),
    (10, "As-at date is a real date, not the file's period end", "All 3 entities",
     None, None, None, "PASS", "01-Aug-2026 used, NOT the file's 31-Mar-27",
     "Scope & Method"),
    (11, "No item is dated before its own file starts", "Every item",
     0, 0, 0, "PASS", "Opening dates are 01-Apr-25 / 01-Apr-26 / 01-Apr-20 respectively",
     "Scope & Method"),
    (12, "No entity shows value in a bucket its file cannot reach", "3 entities",
     0, 0, 0, "PASS",
     "Enterprises is capped at 91-180 days; its three older buckets are structurally nil",
     "Scope & Method"),
]
r += 1
for row in data:
    for i, v in enumerate(row, start=1):
        if v is None:
            continue
        if i == 9:
            link(ws, r, 9, v, v)
            continue
        ws.cell(row=r, column=i, value=v)
    r += 1
paint(ws, hr + 1, r - 1, 9, {4: MONEY, 5: MONEY, 6: MONEY}, flag_cols={7})

# ==================================================== 10. Scope & Method
ws = wb["Scope & Method"]
r = frame(ws, 4, "Scope & Method")
ws.cell(row=r, column=1, value="SCOPE & METHOD — AND WHAT IS NOT COVERED").font = TITLE_FONT
r += 2

ws.cell(row=r, column=1, value="HOW FAR BACK EACH FILE CAN SEE").font = SUB_FONT
r += 1
ws.cell(row=r, column=1,
        value="This is the single biggest limitation of the report. The three entities have "
              "wildly different amounts of history, so their bucket profiles are NOT directly "
              "comparable.").font = NOTE_FONT
r += 1
head(ws, r, ["Legal Entity", "History available from", "Opening date used", "What it means"],
     widths=[22, 20, 18, 84])
hr = r
r += 1
for ent, txt in [
    ("Orange O Tec", "Stock older than 01-Apr-25 is dated 01-Apr-25 and lands in '1 - 2 years'. "
                     "It may be much older."),
    ("Enterprises", "Only 4 months of history, so NOTHING in this entity can show past "
                    "'91 - 180 days' — its '181-365', '1-2 years' and 'Over 2 years' buckets are "
                    "structurally zero, not genuinely empty. Stock that has sat for years still "
                    "reads as 122 days old. Read Enterprises from the Slow & Non-Moving sheet "
                    "instead; its FY2024-26 file would extend history to 01-Apr-24."),
    ("Colorix", "6 years of history — the only entity where 'Over 2 years' is genuinely "
                "measured rather than assumed."),
]:
    ws.cell(row=r, column=1, value=ent)
    ws.cell(row=r, column=2, value=PERIOD_FROM[ent]).number_format = DATE
    ws.cell(row=r, column=3, value="Opening")
    ws.cell(row=r, column=4, value=txt).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 34
    r += 1
paint(ws, hr + 1, r - 1, 4, {2: DATE}, flag_cols={3})
r += 2

head(ws, r, ["Item", "Treatment", "Why it matters"], widths=[30, 60, 84])
hr = r
data = [
    ("As-at date", "01-Aug-2026 (today). NOT 31-Mar-2027.",
     "Your Tally screen was run 'as at 31-Mar-27' — the file's period end, a FUTURE date. "
     "That adds ~8 months of age to every item and inflates every old bucket."),
    ("Ageing basis", "Date of receipt into stock, FIFO — the newest receipts are what remains.",
     "Same convention as Tally's 'Aged by Date of Purchase', so the two are comparable."),
    ("Buckets", "0-90 / 91-180 / 181-365 / 1-2 yrs / >2 yrs. Five, not three.",
     "Tally's screen stops at >180, which cannot tell 7-month-old stock from 3-year-old stock "
     "— the whole point of the exercise."),
    ("Stock with no purchase layer",
     "Dated at the START of the period the file makes available, and labelled 'Opening' in the "
     "Receipt Source column.",
     "Every item therefore lands in a real bucket and nothing is dropped. The Source column is "
     "what stops an assumed date being read as a measured one."),
    ("Valuation of a bucket", "Actual receipt rate of the layer that remains.",
     "Matches Tally's 'valued based on actual purchase'."),
    ("Quantity with no value", "Kept in its bucket, contributing zero value, and listed "
                               "separately with a coverage ratio.",
     "A different problem from an unknown date. The ageing cannot price this stock at all."),
    ("Entities", "5 Tally files -> 3 legal entities. HO + Noida combined.",
     "HO and Noida are one legal company. Ageing is a company-level question."),
    ("Companies", "Period 2627, from Files\\TALLY Company_List.xlsx, re-read every run.",
     "The user maintains that file; rows marked 'Not To USE' are always excluded."),
    ("Movement analysis", "Last outward date and months of cover, alongside purchase age.",
     "Unaffected by how far back the file goes, so it is the more reliable of the two signals "
     "for Enterprises."),
    ("NOT MODELLED — provision", "No provision or write-down percentage is applied anywhere.",
     "Excluded at the user's instruction. This report measures age; it does not price it."),
    ("NOT COVERED — inter-branch", "Transfers are not eliminated.",
     "An HO->Noida transfer resets the receipt date at Noida, so transferred stock looks "
     "younger than it is. Quantified in the inventory-valuation skill."),
    ("NOT COVERED — landed cost", "Duty / freight / CHA sitting in expense ledgers is excluded.",
     "Bucket values are invoice cost, not landed cost, so they understate true carrying cost."),
    ("NOT COVERED — archived files",
     "Prior-year company files on disk (FY22-23, FY17-18, FY15-16) are not opened.",
     "Opening them would convert some Opening-dated stock into genuinely measured dates, at the "
     "cost of a much longer run. Say the word and it becomes a scoped follow-up."),
]
r += 1
for a, b, c in data:
    ws.cell(row=r, column=1, value=a)
    ws.cell(row=r, column=2, value=b)
    ws.cell(row=r, column=3, value=c)
    for ci in range(1, 4):
        ws.cell(row=r, column=ci).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 34
    r += 1
paint(ws, hr + 1, r - 1, 3, {})

wb.save(OUT)
print(f"written: {OUT}")
print(f"sheets : {len(wb.sheetnames)} -> {', '.join(wb.sheetnames)}")
