"""Inventory Valuation — design sample builder.

Rebuilds `Inventory\\SAMPLE LAYOUT - Stock_Valuation_Check.xlsx`, the approved layout for the
live workbook: 10 cross-linked sheets, detail at LEGAL ENTITY level (Orange O Tec / Enterprises /
Colorix, each combining its HO + Noida Tally file), with a group Summary over all three.

ALL FIGURES ARE ILLUSTRATIVE PLACEHOLDERS. Nothing here is fetched from Tally. This script exists
so the layout can be reviewed and changed without touching live data — edit it, re-run it, and the
sample refreshes in place.

Run:
    python "Stock Report\\.claude\\skills\\inventory-valuation\\scripts\\build_sample_layout.py"
"""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def project_root() -> Path:
    """Walk up from this script until the folder holding .env is found."""
    here = Path(__file__).resolve()
    for parent in here.parents:
        if (parent / ".env").is_file():
            return parent
    # fallback: scripts -> inventory-valuation -> skills -> .claude -> Stock Report -> root
    return here.parents[5]


OUTDIR = project_root() / "Inventory"
OUTDIR.mkdir(parents=True, exist_ok=True)
OUT = OUTDIR / "SAMPLE LAYOUT - Stock_Valuation_Check.xlsx"

BANNER_FILL = PatternFill("solid", fgColor="C00000")
BANNER_FONT = Font(bold=True, color="FFFFFF", size=11)
NAV_FILL = PatternFill("solid", fgColor="EDEDED")
LINK_FONT = Font(color="0563C1", underline="single", size=9, bold=True)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(bold=True, size=13, color="1F3864")
SUB_FONT = Font(bold=True, size=10, color="1F3864")
BLOCK_FILL = PatternFill("solid", fgColor="D9E1F2")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE"); BAD_FONT = Font(color="9C0006", bold=True)
OK_FILL = PatternFill("solid", fgColor="C6EFCE"); OK_FONT = Font(color="006100", bold=True)
WARN_FILL = PatternFill("solid", fgColor="FFEB9C"); WARN_FONT = Font(color="9C5700", bold=True)
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '#,##0.00;[Red]-#,##0.00'
MONEY0 = '#,##0;[Red]-#,##0'
QTY = '#,##0.000'
RATE = '#,##0.00'
PCT = '0.0%'

BANNER = ("SAMPLE LAYOUT  —  ILLUSTRATIVE FIGURES ONLY  —  NOT FETCHED FROM TALLY  "
          "—  design draft v2 for approval")

SHEETS = [
    ("Summary", "Summary"),
    ("Orange O Tec", "Orange O Tec"),
    ("Enterprises", "Enterprises"),
    ("Colorix", "Colorix"),
    ("Costing Method Audit", "Costing"),
    ("Landed Cost", "Landed Cost"),
    ("NRV Exceptions", "NRV"),
    ("Inter-Branch", "Inter-Branch"),
    ("Exception Register", "Exceptions"),
    ("Cross Check", "Cross Check"),
]

BAD = {"FROZEN", "OVERSTATED", "FAIL", "WRITE-DOWN", "NOT CAPITALISED", "ELIMINATE"}
GOOD = {"OK", "TRACKS COST", "PASS", "CAPITALISED", "NONE"}
WARN = {"REVIEW", "PARTIAL", "CHECK"}


def link(ws, row, col, text, target_sheet, target_cell="A1"):
    c = ws.cell(row=row, column=col, value=text)
    c.hyperlink = f"#'{target_sheet}'!{target_cell}"
    c.font = LINK_FONT
    return c


def frame(ws, ncols, current):
    """Banner + nav bar on every sheet. Returns next free row."""
    ws["A1"] = BANNER
    ws["A1"].fill, ws["A1"].font = BANNER_FILL, BANNER_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(ncols, 10))
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


def style_table(ws, hdr_row, ncols, headers, formats):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=hdr_row, column=c)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER
    ws.row_dimensions[hdr_row].height = 34
    for ri in range(hdr_row + 1, ws.max_row + 1):
        lbl = str(ws.cell(row=ri, column=1).value or "").upper()
        is_tot = lbl.startswith(("TOTAL", "GRAND", "CONSOLIDATED", "ENTITY", "GROUP"))
        for ci in range(1, ncols + 1):
            cell = ws.cell(row=ri, column=ci)
            cell.border = BORDER
            h = headers[ci - 1]
            if h in formats:
                cell.number_format = formats[h]
            if is_tot:
                cell.fill = TOT_FILL
                if cell.font.color is None or cell.font.color.rgb != "FF0563C1":
                    cell.font = Font(bold=True)
            v = cell.value
            if isinstance(v, str):
                if v in BAD:
                    cell.fill, cell.font = BAD_FILL, BAD_FONT
                elif v in GOOD:
                    cell.fill, cell.font = OK_FILL, OK_FONT
                elif v in WARN:
                    cell.fill, cell.font = WARN_FILL, WARN_FONT
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
    ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(ncols)}{ws.max_row}"


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


wb = Workbook()
ws_by_name = {}
for i, (name, _) in enumerate(SHEETS):
    ws_by_name[name] = wb.active if i == 0 else wb.create_sheet()
    ws_by_name[name].title = name

# ================================================================= SUMMARY
ws = ws_by_name["Summary"]
r = frame(ws, 11, "Summary")
ws.cell(row=r, column=1, value="STOCK VALUATION CHECK — GROUP SUMMARY").font = TITLE_FONT
ws.cell(row=r + 1, column=1,
        value="Period 1-Apr-26 to 31-Jul-26  ·  Period code 2627  ·  5 Tally companies "
              "presented as 3 legal entities  ·  Source: Tally Prime @ localhost:9010")
r += 3

hdr = ["Legal Entity", "Tally Files Combined", "Items with Stock", "Tally Carrying Value",
       "FIFO Value", "Wtd Avg Value", "Gap vs FIFO", "Var %", "Items Flagged",
       "Verdict", "Detail"]
ws.append([]) if False else None
for i, h in enumerate(hdr, start=1):
    ws.cell(row=r, column=i, value=h)
hdr_row = r
rows = [
    ["ORANGE O TEC", "HO + Noida (2 files)", 680, 115730000, 109290000, 108980000,
     6440000, 0.0556, 92, "OVERSTATED", "Orange O Tec"],
    ["ORANGE O TEC ENTERPRISES", "HO + Noida (2 files)", 254, 19140000, 19085000, 19070000,
     55000, 0.0029, 6, "OK", "Enterprises"],
    ["COLORIX DIGITAL PRINTING LLP", "Standalone (1 file)", 74, 4180000, 4120000, 4115000,
     60000, 0.0144, 6, "REVIEW", "Colorix"],
    ["GROUP TOTAL", "5 files", 1008, 139050000, 132495000, 132165000, 6555000, 0.0471,
     104, "OVERSTATED", ""],
    ["Less: inter-branch unrealised profit in stock", "", None, None, None, None,
     -890000, None, None, "", "Inter-Branch"],
    ["CONSOLIDATED GROUP EFFECT", "", None, None, None, None, 5665000, None, None, "", ""],
]
for row in rows:
    r += 1
    for i, v in enumerate(row[:-1], start=1):
        ws.cell(row=r, column=i, value=v)
    if row[-1]:
        link(ws, r, 11, "open →", row[-1])
    if row[0] in ("ORANGE O TEC", "ORANGE O TEC ENTERPRISES", "COLORIX DIGITAL PRINTING LLP"):
        link(ws, r, 1, row[0], row[-1])

fmts = {"Tally Carrying Value": MONEY0, "FIFO Value": MONEY0, "Wtd Avg Value": MONEY0,
        "Gap vs FIFO": MONEY0, "Var %": PCT}
style_table(ws, hdr_row, 11, hdr, fmts)

# contents block
r += 3
ws.cell(row=r, column=1, value="WHERE TO LOOK").font = SUB_FONT
r += 1
for c, h in enumerate(["Sheet", "What it holds", "Why it matters"], start=1):
    cell = ws.cell(row=r, column=c, value=h)
    cell.fill, cell.font = HDR_FILL, HDR_FONT
    cell.border = BORDER
contents = [
    ("Orange O Tec", "Item-level gap detail, HO and Noida split then combined",
     "Largest exposure — 92 flagged items"),
    ("Enterprises", "Item-level gap detail, HO and Noida split then combined",
     "Clean so far"),
    ("Colorix", "Item-level gap detail, standalone", "Small but non-nil"),
    ("Costing Method Audit", "Costing method per item, all entities",
     "ROOT CAUSE — decides if the gap is real or display-only"),
    ("Landed Cost", "Duty / freight / CHA per import consignment",
     "Narrows the gap — raises true cost"),
    ("NRV Exceptions", "Items carried above realisable value", "Write-down candidates"),
    ("Inter-Branch", "HO↔Noida transfers and unrealised profit",
     "Eliminated at entity and group level"),
    ("Exception Register", "Re-1 issues, rate mismatches, stock journals, negative stock",
     "Voucher-level fixes"),
    ("Cross Check", "10 reconciliation tests + self-proof of top-20 gaps",
     "Read this before trusting any figure"),
]
for sheet_name, what, why in contents:
    r += 1
    link(ws, r, 1, sheet_name, sheet_name)
    ws.cell(row=r, column=2, value=what)
    ws.cell(row=r, column=3, value=why)
    for c in range(1, 4):
        ws.cell(row=r, column=c).border = BORDER
widths(ws, [34, 22, 13, 18, 16, 16, 15, 9, 11, 14, 10])
ws.column_dimensions["B"].width = 22

# ====================================================== ENTITY DETAIL SHEETS
ENTITY_DATA = {
    "Orange O Tec": {
        "title": "ORANGE O TEC — legal entity view (HO + Noida branch combined)",
        "files": ["ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)  [HO]",
                  "ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)  [Noida]"],
        "block": [
            ["Items with stock", 412, 268, 680],
            ["Tally carrying value", 84250000, 31480000, 115730000],
            ["FIFO value", 79180000, 30110000, 109290000],
            ["Wtd avg value", 78940000, 30040000, 108980000],
            ["Gap vs FIFO", 5070000, 1370000, 6440000],
            ["Var %", 0.0602, 0.0435, 0.0556],
            ["Items flagged", 63, 29, 92],
            ["Less: inter-branch unrealised profit", None, None, -890000],
            ["ENTITY NET EFFECT", None, None, 5550000],
        ],
        "items": [
            [1, "HO", "PRINTING INK", "REACTIVE INK", "REACTIVE INK H-SERIES BLACK", "KGS",
             4515.000, 393.58, 1777013.70, 315.80, 1425837.00, 312.91, 1412788.65, 316.09,
             351176.70, 0.2463, "Avg. Cost", "FROZEN", "Exception Register"],
            [2, "HO", "PRINTING INK", "REACTIVE INK", "REACTIVE INK H-SERIES CYAN", "KGS",
             2180.000, 401.20, 874616.00, 324.10, 706538.00, 322.45, 702941.00, 322.45,
             168078.00, 0.2379, "Avg. Cost", "FROZEN", ""],
            [3, "Noida", "PRINTING INK", "REACTIVE INK", "REACTIVE INK H-SERIES BLACK", "KGS",
             1250.000, 380.00, 475000.00, 316.09, 395112.50, 315.20, 394000.00, 316.09,
             79887.50, 0.2022, "Avg. Cost", "FROZEN", "Inter-Branch"],
            [4, "HO", "CHEMICALS", "", "SODA ASH LIGHT", "KGS",
             155.400, 628.62, 97687.55, 631.00, 98057.40, 630.50, 97979.70, 631.00,
             -369.85, -0.0038, "FIFO", "OK", ""],
            ["…", "…", "…", "…", "(one row per item — 680 rows, ranked by gap, autofiltered)",
             "…", None, None, None, None, None, None, None, None, None, None, "…", "…", ""],
        ],
    },
    "Enterprises": {
        "title": "ORANGE O TEC ENTERPRISES — legal entity view (HO + Noida branch combined)",
        "files": ["ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)  [HO]",
                  "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27  [Noida]"],
        "block": [
            ["Items with stock", 156, 98, 254],
            ["Tally carrying value", 12730000, 6410000, 19140000],
            ["FIFO value", 12690000, 6395000, 19085000],
            ["Wtd avg value", 12680000, 6390000, 19070000],
            ["Gap vs FIFO", 40000, 15000, 55000],
            ["Var %", 0.0031, 0.0023, 0.0029],
            ["Items flagged", 4, 2, 6],
            ["Less: inter-branch unrealised profit", None, None, 0],
            ["ENTITY NET EFFECT", None, None, 55000],
        ],
        "items": [
            [1, "HO", "MACHINERY & MACHINERY PARTS", "MACHINARY PARTS", "PRINT HEAD KYOCERA",
             "PCS", 12.000, 68500.00, 822000.00, 66200.00, 794400.00, 66050.00, 792600.00,
             66200.00, 27600.00, 0.0347, "FIFO", "REVIEW", ""],
            [2, "Noida", "SPARES", "OTHER SPARE PARTS", "DAMPER ASSEMBLY", "PCS",
             240.000, 1450.00, 348000.00, 1398.00, 335520.00, 1395.00, 334800.00, 1398.00,
             12480.00, 0.0372, "FIFO", "REVIEW", ""],
            ["…", "…", "…", "…", "(one row per item — 254 rows)", "…", None, None, None,
             None, None, None, None, None, None, None, "…", "…", ""],
        ],
    },
    "Colorix": {
        "title": "COLORIX DIGITAL PRINTING SOLUTIONS LLP — standalone",
        "files": ["COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)"],
        "block": [
            ["Items with stock", 74, None, 74],
            ["Tally carrying value", 4180000, None, 4180000],
            ["FIFO value", 4120000, None, 4120000],
            ["Wtd avg value", 4115000, None, 4115000],
            ["Gap vs FIFO", 60000, None, 60000],
            ["Var %", 0.0144, None, 0.0144],
            ["Items flagged", 6, None, 6],
            ["Less: inter-branch unrealised profit", None, None, 0],
            ["ENTITY NET EFFECT", None, None, 60000],
        ],
        "items": [
            [1, "—", "PAPER ROLL", "SUBLIMATION", "SUBLIMATION PAPER 1.6M", "MTR",
             9000.000, 5.46, 49140.00, 5.41, 48690.00, 5.40, 48600.00, 5.41, 450.00,
             0.0092, "FIFO", "OK", ""],
            [2, "—", "PRINTING INK", "", "DISPERSE INK MAGENTA", "KGS",
             820.000, 512.00, 419840.00, 448.00, 367360.00, 446.50, 366130.00, 448.00,
             52480.00, 0.1429, "Avg. Cost", "FROZEN", "NRV Exceptions"],
            ["…", "…", "…", "…", "(one row per item — 74 rows)", "…", None, None, None,
             None, None, None, None, None, None, None, "…", "…", ""],
        ],
    },
}

ITEM_HDR = ["Rank", "Unit", "Main Category", "Sub Category", "Item Name", "UOM",
            "Closing Qty", "Tally Rate", "Tally Value", "FIFO Rate", "FIFO Value",
            "Wtd Avg Rate", "Wtd Avg Value", "Last Pur Rate", "Gap vs FIFO", "Gap %",
            "Costing Method", "Flag", "See Also"]
ITEM_FMT = {"Closing Qty": QTY, "Tally Rate": RATE, "Tally Value": MONEY,
            "FIFO Rate": RATE, "FIFO Value": MONEY, "Wtd Avg Rate": RATE,
            "Wtd Avg Value": MONEY, "Last Pur Rate": RATE, "Gap vs FIFO": MONEY,
            "Gap %": PCT}
ITEM_W = [6, 8, 24, 18, 34, 7, 12, 11, 14, 11, 14, 12, 14, 12, 14, 9, 14, 12, 18]

for ent, d in ENTITY_DATA.items():
    ws = ws_by_name[ent]
    r = frame(ws, len(ITEM_HDR), ent)
    ws.cell(row=r, column=1, value=d["title"]).font = TITLE_FONT
    r += 1
    ws.cell(row=r, column=1, value="Tally files: " + "   |   ".join(d["files"])).font = \
        Font(italic=True, size=9)
    r += 2

    ws.cell(row=r, column=1, value="ENTITY ROLL-UP").font = SUB_FONT
    r += 1
    blk_hdr = ["Measure", "HO", "Noida Branch", "Entity Combined"]
    for i, h in enumerate(blk_hdr, start=1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BORDER
    blk_start = r
    for brow in d["block"]:
        r += 1
        for i, v in enumerate(brow, start=1):
            cell = ws.cell(row=r, column=i, value=v)
            cell.border = BORDER
            if i > 1:
                if brow[0] == "Var %":
                    cell.number_format = PCT
                elif brow[0] in ("Items with stock", "Items flagged"):
                    cell.number_format = '#,##0'
                else:
                    cell.number_format = MONEY0
        if brow[0].startswith("ENTITY"):
            for i in range(1, 5):
                ws.cell(row=r, column=i).fill = TOT_FILL
                ws.cell(row=r, column=i).font = Font(bold=True)
        if brow[0].startswith("Less:"):
            link(ws, r, 1, brow[0], "Inter-Branch")
    r += 2

    ws.cell(row=r, column=1, value="ITEM-LEVEL VALUATION GAP").font = SUB_FONT
    r += 1
    for i, h in enumerate(ITEM_HDR, start=1):
        ws.cell(row=r, column=i, value=h)
    item_hdr_row = r
    for irow in d["items"]:
        r += 1
        for i, v in enumerate(irow[:-1], start=1):
            ws.cell(row=r, column=i, value=v)
        if irow[-1]:
            link(ws, r, len(ITEM_HDR), irow[-1], irow[-1])
    style_table(ws, item_hdr_row, len(ITEM_HDR), ITEM_HDR, ITEM_FMT)
    widths(ws, ITEM_W)
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 10

# =========================================================== SUPPORT SHEETS
def support(name, headers, rows, ws_widths, formats, backlinks=None):
    ws = ws_by_name[name]
    r = frame(ws, len(headers), name)
    ws.cell(row=r, column=1, value=name.upper()).font = TITLE_FONT
    r += 2
    for i, h in enumerate(headers, start=1):
        ws.cell(row=r, column=i, value=h)
    hr = r
    for row in rows:
        r += 1
        for i, v in enumerate(row, start=1):
            ws.cell(row=r, column=i, value=v)
        if backlinks:
            tgt = backlinks(row)
            if tgt:
                link(ws, r, 1, str(row[0]), tgt)
    style_table(ws, hr, len(headers), headers, formats)
    widths(ws, ws_widths)


ENT_OF = {"ORANGE O TEC": "Orange O Tec", "ORANGE O TEC ENTERPRISES": "Enterprises",
          "COLORIX": "Colorix"}

support(
    "Costing Method Audit",
    ["Entity", "Unit", "Stock Group", "Item Name", "Costing Method",
     "Market Valuation Method", "Std Cost Defined?", "Std Rate", "Closing Qty",
     "Tally Closing Rate", "Last Purchase Rate", "Rate Moved in Period?", "Flag"],
    [
        ["ORANGE O TEC", "HO", "PRINTING INK", "REACTIVE INK H-SERIES BLACK", "Avg. Cost",
         "Last Sale Price", "Yes", 393.58, 4515.000, 393.58, 316.09, "No", "FROZEN"],
        ["ORANGE O TEC", "HO", "PRINTING INK", "REACTIVE INK H-SERIES CYAN", "Avg. Cost",
         "Last Sale Price", "Yes", 401.20, 2180.000, 401.20, 322.45, "No", "FROZEN"],
        ["ORANGE O TEC", "Noida", "PRINTING INK", "REACTIVE INK H-SERIES BLACK", "Avg. Cost",
         "Last Sale Price", "Yes", 380.00, 1250.000, 380.00, 316.09, "No", "FROZEN"],
        ["ORANGE O TEC", "HO", "CHEMICALS", "SODA ASH LIGHT", "FIFO", "Avg. Price", "No",
         None, 155.400, 628.62, 631.00, "Yes", "TRACKS COST"],
        ["COLORIX", "—", "PRINTING INK", "DISPERSE INK MAGENTA", "Avg. Cost",
         "Last Sale Price", "Yes", 512.00, 820.000, 512.00, 448.00, "No", "FROZEN"],
        ["…", "…", "…", "(one row per stock item across all 5 files)", "…", "…", "…",
         None, None, None, None, "…", "…"],
    ],
    [22, 8, 20, 34, 15, 17, 12, 11, 13, 14, 14, 13, 14],
    {"Std Rate": RATE, "Closing Qty": QTY, "Tally Closing Rate": RATE,
     "Last Purchase Rate": RATE},
    backlinks=lambda row: ENT_OF.get(row[0]),
)

support(
    "Landed Cost",
    ["Entity", "Date", "Supplier", "Vch No", "Item", "Qty", "UOM", "Invoice Value",
     "Invoice Rate", "Customs Duty", "Freight", "CHA / Clearing", "Forex Diff",
     "Landed Value", "Landed Rate", "Booked to Stock?", "Cost Not Capitalised"],
    [
        ["ORANGE O TEC", "2026-05-20", "HAN GLORY HONG KONG LIMITED-INK", "CI260414048I",
         "REACTIVE INK H-SERIES BLACK", 1000.000, "KGS", 306150.00, 306.15, 30615.00,
         12400.00, 4800.00, 1850.00, 355815.00, 355.82, "NOT CAPITALISED", 49665.00],
        ["ORANGE O TEC", "2026-06-04", "HAN GLORY HONG KONG LIMITED-INK", "CI260422051I",
         "REACTIVE INK H-SERIES BLACK", 2800.000, "KGS", 865865.01, 309.24, 86586.50,
         31200.00, 9600.00, -2100.00, 991151.51, 353.98, "NOT CAPITALISED", 125286.50],
        ["ORANGE O TEC", "2026-06-05", "HAN GLORY HONG KONG LIMITED-INK", "CI260430054I",
         "REACTIVE INK H-SERIES BLACK", 3000.000, "KGS", 946237.51, 315.41, 94623.75,
         33400.00, 9600.00, 3200.00, 1087061.26, 362.35, "NOT CAPITALISED", 140823.75],
        ["ORANGE O TEC", "2026-06-29", "HAN GLORY HONG KONG LIMITED-INK", "CI260515057I",
         "REACTIVE INK H-SERIES BLACK", 3000.000, "KGS", 948276.00, 316.09, 94827.60,
         34100.00, 9600.00, 0.00, 1086803.60, 362.27, "NOT CAPITALISED", 138527.60],
        ["TOTAL — import consignments in period", "", "", "", "", 9800.000, "KGS",
         3066528.52, None, 306652.85, 111100.00, 33600.00, 2950.00, 3520831.37, 359.27,
         "", 454302.85],
    ],
    [22, 11, 32, 15, 30, 11, 7, 14, 12, 13, 11, 14, 11, 14, 12, 16, 17],
    {"Qty": QTY, "Invoice Value": MONEY, "Invoice Rate": RATE, "Customs Duty": MONEY,
     "Freight": MONEY, "CHA / Clearing": MONEY, "Forex Diff": MONEY, "Landed Value": MONEY,
     "Landed Rate": RATE, "Cost Not Capitalised": MONEY},
    backlinks=lambda row: ENT_OF.get(row[0]),
)

support(
    "NRV Exceptions",
    ["Entity", "Unit", "Item", "Closing Qty", "UOM", "Carrying Rate", "Carrying Value",
     "Lowest Arm's-Length Sale Rate", "Sale Date", "Party", "Est. Selling Cost %",
     "NRV per Unit", "Shortfall per Unit", "Write-down Required", "Status"],
    [
        ["ORANGE O TEC", "HO", "REACTIVE INK H-SERIES BLACK", 4515.000, "KGS", 393.58,
         1777013.70, 400.00, "2026-06-25", "AJANTA DIGITAL INDUSTRIES-SURAT", 0.05, 380.00,
         13.58, 61313.70, "WRITE-DOWN"],
        ["COLORIX", "—", "DISPERSE INK MAGENTA", 820.000, "KGS", 512.00, 419840.00, 495.00,
         "2026-07-08", "(third party)", 0.05, 470.25, 41.75, 34235.00, "WRITE-DOWN"],
        ["ORANGE O TEC", "HO", "REACTIVE INK H-SERIES CYAN", 2180.000, "KGS", 401.20,
         874616.00, 430.00, "2026-06-18", "AJANTA DIGITAL INDUSTRIES-SURAT", 0.05, 408.50,
         0.00, 0.00, "PASS"],
        ["TOTAL — write-down required", "", "", None, "", None, None, None, "", "", None,
         None, None, 95548.70, ""],
    ],
    [22, 8, 32, 12, 7, 13, 15, 18, 11, 32, 13, 12, 13, 15, 13],
    {"Closing Qty": QTY, "Carrying Rate": RATE, "Carrying Value": MONEY,
     "Lowest Arm's-Length Sale Rate": RATE, "Est. Selling Cost %": PCT,
     "NRV per Unit": RATE, "Shortfall per Unit": RATE, "Write-down Required": MONEY},
    backlinks=lambda row: ENT_OF.get(row[0]),
)

support(
    "Inter-Branch",
    ["Entity", "From (Tally file)", "To (Tally file)", "Item", "Qty Transferred", "UOM",
     "Transfer Rate", "Transfer Value", "Qty Received",
     "Qty Still in Stock at 31-Jul-26", "Group Cost Rate", "Margin per Unit",
     "Unrealised Profit in Stock", "Status"],
    [
        ["ORANGE O TEC", "OOT HO", "OOT Noida", "REACTIVE INK H-SERIES BLACK", 5100.000,
         "KGS", 380.00, 1938000.00, 5100.000, 1250.000, 312.91, 67.09, 83862.50, "ELIMINATE"],
        ["ORANGE O TEC", "OOT Noida", "OOT HO", "REACTIVE INK H-SERIES BLACK", 200.000,
         "KGS", 380.00, 76000.00, 200.000, 0.000, 312.91, 67.09, 0.00, "CHECK"],
        ["ORANGE O TEC ENTERPRISES", "Ent HO", "Ent Noida", "(other items)", None, "",
         None, None, None, None, None, None, 0.00, "NONE"],
        ["ENTITY TOTAL — Orange O Tec", "", "", "", None, "", None, None, None, None, None,
         None, 890000.00, ""],
        ["GROUP TOTAL — unrealised profit to eliminate", "", "", "", None, "", None, None,
         None, None, None, None, 890000.00, ""],
    ],
    [26, 16, 16, 30, 14, 7, 12, 14, 12, 18, 12, 12, 17, 12],
    {"Qty Transferred": QTY, "Transfer Rate": RATE, "Transfer Value": MONEY,
     "Qty Received": QTY, "Qty Still in Stock at 31-Jul-26": QTY, "Group Cost Rate": RATE,
     "Margin per Unit": RATE, "Unrealised Profit in Stock": MONEY},
    backlinks=lambda row: ENT_OF.get(row[0]),
)

support(
    "Exception Register",
    ["Entity", "Unit", "Date", "Voucher Type", "Vch No", "Item", "Party", "Qty", "Rate",
     "Value", "Exception Type", "Why Flagged", "Est. ₹ Impact", "Suggested Action"],
    [
        ["ORANGE O TEC", "HO", "2026-06-26", "DELIVERY CHALLAN", "DC/INK/2627/46",
         "REACTIVE INK H-SERIES BLACK", "BHOOMI FASHION PVT LTD", 10.000, 1.00, 10.00,
         "Near-zero value issue",
         "Stock relieved at Re 1/kg vs cost ~313/kg; challan may never have been invoiced",
         3129.10, "Convert to invoice or book as sample expense"],
        ["ORANGE O TEC", "HO", "2026-06-11", "GST SALES RETURN", "G/SR/26-27/112",
         "REACTIVE INK H-SERIES BLACK", "O FAB", -25.000, 690.00, -17250.00,
         "Return rate mismatch",
         "Returned at 690 vs original sale rate 650 to the same party", 1000.00,
         "Reverse and re-book credit note at 650"],
        ["ORANGE O TEC", "HO", "2026-05-09", "STOCK JOURNAL", "(auto)",
         "REACTIVE INK H-SERIES BLACK", "(internal)", -20.000, 501.03, -10020.60,
         "Unexplained rate",
         "Consumption priced at neither cost (~313) nor any observed sale rate", 3762.60,
         "Trace destination and re-price at cost"],
        ["…", "…", "…", "…", "…", "(one row per exception, all entities)", "…", None, None,
         None, "…", "…", None, "…"],
    ],
    [22, 8, 11, 17, 16, 30, 26, 11, 10, 13, 20, 46, 13, 34],
    {"Qty": QTY, "Rate": RATE, "Value": MONEY, "Est. ₹ Impact": MONEY},
    backlinks=lambda row: ENT_OF.get(row[0]),
)

# ------------------------------------------------------------- Cross Check
ws = ws_by_name["Cross Check"]
r = frame(ws, 9, "Cross Check")
ws.cell(row=r, column=1, value="CROSS CHECK — read before trusting any figure").font = TITLE_FONT
r += 2
cc_hdr = ["#", "Check", "Scope", "Expected", "Computed", "Difference", "Status",
          "Note", "Sheet"]
for i, h in enumerate(cc_hdr, start=1):
    ws.cell(row=r, column=i, value=h)
hr = r
cc_rows = [
    [1, "Closing qty ties to Tally stock summary", "All items, 5 files", 1008, 1008, 0,
     "PASS", "Item count reconciled", "Summary"],
    [2, "Closing value ties to Tally stock summary", "Sum of item closing values",
     139050000, 139050000, 0, "PASS", "Carrying value agrees before adjustment", "Summary"],
    [3, "Opening + Inward - Outward = Closing (qty)", "Per item roll-forward", 0, 0, 0,
     "PASS", "Quantity flow internally consistent", "Orange O Tec"],
    [4, "FIFO rebuild consumes exactly the available layers", "Per item", 0, 0, 0, "PASS",
     "No negative layer draws", "Orange O Tec"],
    [5, "Entity roll-up = sum of its Tally files", "HO + Noida per entity", 0, 0, 0,
     "PASS", "No double count across branch files", "Summary"],
    [6, "Items with negative closing qty", "Exception scan", 0, 3, 3, "FAIL",
     "3 items show negative stock", "Exception Register"],
    [7, "Items with value but zero qty", "Exception scan", 0, 2, 2, "FAIL",
     "Stranded value", "Exception Register"],
    [8, "Items with no stock group", "Master-data scan", 0, 0, 0, "PASS", "", "Costing Method Audit"],
    [9, "Inter-branch out = inter-branch in (qty)", "Cross-file match", 5300.000, 5300.000,
     0, "PASS", "Both sides agree", "Inter-Branch"],
    [10, "Top-20 gaps re-proved from raw vouchers", "Independent recalc", 20, 20, 0, "PASS",
     "Mandatory self-check before any figure is quoted", "Orange O Tec"],
]
for row in cc_rows:
    r += 1
    for i, v in enumerate(row[:-1], start=1):
        ws.cell(row=r, column=i, value=v)
    link(ws, r, 9, row[-1], row[-1])
style_table(ws, hr, 9, cc_hdr,
            {"Expected": MONEY0, "Computed": MONEY0, "Difference": MONEY0})
widths(ws, [5, 44, 26, 15, 15, 13, 10, 46, 22])

wb.save(OUT)
print("Written:", OUT)
for w in wb.worksheets:
    print(f"  {w.title:<24} cols={w.max_column:>2}  rows={w.max_row}")
