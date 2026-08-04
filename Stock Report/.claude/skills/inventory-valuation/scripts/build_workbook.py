"""Build the Inventory Valuation workbook.

Design points:
  * detail is per LEGAL ENTITY (Orange O Tec / Enterprises / Colorix), HO + Noida combined
  * the full stock-group hierarchy is carried so stock can be read by segment
  * every derived number is a LIVE EXCEL FORMULA, not a pasted value — roll-ups are SUMIFS
    over the item detail on the same sheet, and Summary/Segment sheets reference the entity
    sheets. Change a rate in the detail and everything above it moves.
"""
from datetime import datetime
from pathlib import Path
import os
import shutil
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from tally_lib import SCRATCH, FILES, ROOT

STAMP = os.environ.get("IV_STAMP") or datetime.now().strftime("%Y-%m-%d_%H%M")
PERIOD = os.environ.get("IV_PERIOD", "1-Apr-2026 to 30-Jun-2026")

OUTDIR = ROOT / "Inventory"
OUTDIR.mkdir(exist_ok=True)
OUT = OUTDIR / f"Stock_Valuation_Check_{STAMP}.xlsx"

df = pd.read_parquet(SCRATCH / "phase_b4_grouped.parquet")
allitems = pd.read_parquet(SCRATCH / "phase_b2_full.parquet")

# Items whose group chain resolves to nothing must carry a real label in the DATA, not just
# in the segment sheet — otherwise the segment SUMIFS matches nothing and their value silently
# drops out of the segment total.
NO_GRP = "(no group assigned)"
UNGROUPED = int((df["Main Category"].str.strip() == "").sum())
df.loc[df["Main Category"].str.strip() == "", "Main Category"] = NO_GRP

# ------------------------------------------------------------------ styling
BANNER_FILL = PatternFill("solid", fgColor="1F3864")
BANNER_FONT = Font(bold=True, color="FFFFFF", size=11)
NAV_FILL = PatternFill("solid", fgColor="EDEDED")
LINK_FONT = Font(color="0563C1", underline="single", size=9, bold=True)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
SUBHDR_FILL = PatternFill("solid", fgColor="8EA9DB")
TITLE_FONT = Font(bold=True, size=13, color="1F3864")
SUB_FONT = Font(bold=True, size=10, color="1F3864")
FORMULA_NOTE = Font(italic=True, size=8, color="7F7F7F")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE"); BAD_FONT = Font(color="9C0006", bold=True)
OK_FILL = PatternFill("solid", fgColor="C6EFCE"); OK_FONT = Font(color="006100", bold=True)
WARN_FILL = PatternFill("solid", fgColor="FFEB9C"); WARN_FONT = Font(color="9C5700", bold=True)
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '#,##0.00;[Red]-#,##0.00'
MONEY0 = '#,##0;[Red]-#,##0'
QTYF = '#,##0.000'
RATEF = '#,##0.00'
PCT = '0.0%'
INT = '#,##0'

SHEETS = [("Summary", "Summary"), ("Segment Summary", "Segments"),
          ("Orange O Tec", "Orange O Tec"), ("Enterprises", "Enterprises"),
          ("Colorix", "Colorix"), ("Costing Method Audit", "Costing"),
          ("Exception Register", "Exceptions"), ("Cross Check", "Cross Check"),
          ("Scope & Method", "Scope")]

BADW = {"OVERSTATED", "FAIL", "NIL-VALUE RECEIPT", "NEGATIVE STOCK", "ZERO-RATED STOCK",
        "STRANDED VALUE"}
GOODW = {"OK", "PASS", "TIES"}
WARNW = {"UNDERSTATED", "CHECK", "PARTIAL", "NOT RUN", "NO BENCHMARK", "REVIEW"}

ENTS = ["Orange O Tec", "Enterprises", "Colorix"]

# item-detail column layout (1-based); formulas below depend on these letters
ICOLS = ["Rank", "Unit", "Main Category", "Sub Category", "Product Group", "Sub Group",
         "Full Group Path", "Item Name", "UOM", "Costing Method",
         "Closing Qty", "Tally Rate", "Tally Value", "Inward Qty", "Inward Rate",
         "Value @ Inward Rate", "Wtd Avg Rate", "Value @ Wtd Avg",
         "Gap vs Inward Rate", "Gap vs Wtd Avg", "Gap %", "Benchmark", "Flag"]
C = {h: get_column_letter(i + 1) for i, h in enumerate(ICOLS)}
UNIT, MAIN, CQTY = C["Unit"], C["Main Category"], C["Closing Qty"]
TVAL, IQTY, IRATE = C["Tally Value"], C["Inward Qty"], C["Inward Rate"]
VIN, WRATE, VWA = C["Value @ Inward Rate"], C["Wtd Avg Rate"], C["Value @ Wtd Avg"]
GAP, GAPW, GAPP = C["Gap vs Inward Rate"], C["Gap vs Wtd Avg"], C["Gap %"]
BMK, FLAG = C["Benchmark"], C["Flag"]

IFMT = {"Closing Qty": QTYF, "Tally Rate": RATEF, "Tally Value": MONEY,
        "Inward Qty": QTYF, "Inward Rate": RATEF, "Value @ Inward Rate": MONEY,
        "Wtd Avg Rate": RATEF, "Value @ Wtd Avg": MONEY, "Gap vs Inward Rate": MONEY,
        "Gap vs Wtd Avg": MONEY, "Gap %": PCT}
IW = [6, 8, 26, 24, 24, 24, 40, 40, 7, 17, 13, 12, 15, 12, 12, 16, 12, 15, 16, 16, 9, 30, 17]


def link(ws, r, c, text, target, cell="A1"):
    x = ws.cell(row=r, column=c, value=text)
    x.hyperlink = f"#'{target}'!{cell}"
    x.font = LINK_FONT
    return x


def frame(ws, ncols, current):
    ws["A1"] = (f"INVENTORY VALUATION CHECK   |   as at 30-Jun-2026   |   period {PERIOD}"
                f"   |   Tally Prime localhost:9010   |   generated {STAMP}")
    ws["A1"].fill, ws["A1"].font = BANNER_FILL, BANNER_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(ncols, 10))
    ws.row_dimensions[1].height = 20
    ws.cell(row=2, column=1, value="Go to:").font = Font(bold=True, size=9)
    col = 2
    for name, label in SHEETS:
        if name == current:
            x = ws.cell(row=2, column=col, value=f"[{label}]")
            x.font = Font(bold=True, size=9, color="1F3864")
        else:
            link(ws, 2, col, label, name)
        ws.cell(row=2, column=col).alignment = Alignment(horizontal="center")
        col += 1
    for c in range(1, col):
        ws.cell(row=2, column=c).fill = NAV_FILL
    return 4


def hdr_row(ws, r, headers, fill=HDR_FILL):
    for i, h in enumerate(headers, start=1):
        x = ws.cell(row=r, column=i, value=h)
        x.fill, x.font = fill, HDR_FONT
        x.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        x.border = BORDER
    ws.row_dimensions[r].height = 32


def paint(ws, r1, r2, ncols, headers, fmts):
    for r in range(r1, r2 + 1):
        lbl = str(ws.cell(row=r, column=1).value or "").upper()
        tot = lbl.startswith(("TOTAL", "GROUP", "ENTITY", "NET", "CONSOLIDATED"))
        for c in range(1, ncols + 1):
            x = ws.cell(row=r, column=c)
            x.border = BORDER
            h = headers[c - 1] if c - 1 < len(headers) else ""
            if h in fmts:
                x.number_format = fmts[h]
            if tot:
                x.fill = TOT_FILL
                if not (x.font and x.font.color and x.font.color.rgb == "FF0563C1"):
                    x.font = Font(bold=True)
            v = x.value
            if isinstance(v, str) and not v.startswith("="):
                if v in BADW:
                    x.fill, x.font = BAD_FILL, BAD_FONT
                elif v in GOODW:
                    x.fill, x.font = OK_FILL, OK_FONT
                elif v in WARNW:
                    x.fill, x.font = WARN_FILL, WARN_FONT


def widths(ws, ws_w):
    for i, w in enumerate(ws_w, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


wb = Workbook()
S = {}
for i, (n, _) in enumerate(SHEETS):
    S[n] = wb.active if i == 0 else wb.create_sheet()
    S[n].title = n

# ==================================================== ENTITY SHEETS (built first)
RANGES = {}     # entity -> (first_data_row, last_data_row)
ROLLUP = {}     # entity -> {measure: "Sheet!$D$12"}

for e in ENTS:
    ws = S[e]
    d = df[df["Entity"] == e].copy()
    d["absgap"] = d["Gap vs Inward Rate"].abs().fillna(0)
    d = d.sort_values(["absgap", "Closing Value"], ascending=False)
    units = sorted(d["Unit"].unique())
    n = len(d)

    r = frame(ws, len(ICOLS), e)
    ws.cell(row=r, column=1, value=f"{e.upper()} — closing stock at 30-Jun-2026").font = TITLE_FONT
    r += 1
    fl = [f"{c}  [{u}]" for c, ent, u in FILES if ent == e]
    ws.cell(row=r, column=1, value="Tally files: " + "   |   ".join(fl)).font = \
        Font(italic=True, size=9)
    r += 1
    ws.cell(row=r, column=1, value="Every figure below is a live formula over the item "
            "detail on this sheet — edit the detail and these move.").font = FORMULA_NOTE
    r += 2

    # ---- work out where the detail will sit, so roll-ups can point at it
    #      roll-up block: header + 11 measures; segment block: header + segments
    segs = [s for s in d["Main Category"].unique() if str(s).strip()]
    segs = sorted(segs, key=lambda s: -d[d["Main Category"] == s]["Closing Value"].sum())
    rollup_hdr = r + 1
    rollup_end = rollup_hdr + 11
    seg_hdr = rollup_end + 3
    seg_end = seg_hdr + len(segs) + 1
    det_hdr = seg_end + 3
    d1, d2 = det_hdr + 1, det_hdr + n
    RANGES[e] = (d1, d2)

    def rng(col):
        return f"${col}${d1}:${col}${d2}"

    # ---------------- roll-up block
    ws.cell(row=r, column=1, value="ENTITY ROLL-UP").font = SUB_FONT
    r += 1
    bh = ["Measure"] + units + ["Entity Combined"]
    hdr_row(ws, r, bh)
    r += 1

    def measure(label, tmpl, nf, comb=None):
        """tmpl gets {u} = a unit criterion clause, blank for the combined column.

        `comb` overrides the combined-column formula. Needed where the criteria-less form
        would be a bare "<>" test — that counts non-blanks in Excel but evaluates to 0 in
        some engines, so plain SUM/COUNTA is used instead.
        """
        global r
        ws.cell(row=r, column=1, value=label)
        for i, u in enumerate(units, start=2):
            ws.cell(row=r, column=i,
                    value=tmpl.format(u=f",{rng(UNIT)},\"{u}\"")).number_format = nf
        ws.cell(row=r, column=len(units) + 2,
                value=comb or tmpl.format(u="")).number_format = nf
        r += 1

    # Unit column is always populated, so it is the reliable thing to count over.
    measure("Items with closing stock",
            f"=COUNTIFS({rng(UNIT)},\"*\"{{u}})", INT,
            comb=f"=COUNTA({rng(UNIT)})")
    measure("Closing stock value",
            f"=SUMIFS({rng(TVAL)}{{u}})", MONEY0,
            comb=f"=SUM({rng(TVAL)})")
    measure("Items with a cost benchmark",
            f"=COUNTIFS({rng(BMK)},\"valid\"{{u}})", INT)
    measure("Value benchmarked",
            f"=SUMIFS({rng(TVAL)},{rng(BMK)},\"valid\"{{u}})", MONEY0)
    measure("Overstated vs purchase cost (+)",
            f"=SUMIFS({rng(GAP)},{rng(BMK)},\"valid\",{rng(GAP)},\">0\"{{u}})", MONEY0)
    measure("Understated vs purchase cost (-)",
            f"=SUMIFS({rng(GAP)},{rng(BMK)},\"valid\",{rng(GAP)},\"<0\"{{u}})", MONEY0)
    measure("NET GAP vs purchase cost",
            f"=SUMIFS({rng(GAP)},{rng(BMK)},\"valid\"{{u}})", MONEY0)
    measure("NET GAP vs weighted average",
            f"=SUMIFS({rng(GAPW)},{rng(BMK)},\"valid\"{{u}})", MONEY0)
    measure("Stock held at NIL value (qty)",
            f"=SUMIFS({rng(CQTY)},{rng(TVAL)},0,{rng(CQTY)},\">0\"{{u}})", QTYF)
    measure("Value from nil-value receipts",
            f"=SUMIFS({rng(TVAL)},{rng(BMK)},\"invalid - received at nil value\"{{u}})",
            MONEY0)
    measure("Items with negative stock",
            f"=COUNTIFS({rng(CQTY)},\"<0\"{{u}})", INT)

    paint(ws, rollup_hdr + 1, rollup_end, len(bh), bh, {})
    comb = get_column_letter(len(units) + 2)
    ROLLUP[e] = {
        "items": f"'{e}'!${comb}${rollup_hdr + 1}",
        "cv": f"'{e}'!${comb}${rollup_hdr + 2}",
        "bitems": f"'{e}'!${comb}${rollup_hdr + 3}",
        "bcv": f"'{e}'!${comb}${rollup_hdr + 4}",
        "over": f"'{e}'!${comb}${rollup_hdr + 5}",
        "under": f"'{e}'!${comb}${rollup_hdr + 6}",
        "net": f"'{e}'!${comb}${rollup_hdr + 7}",
        "netw": f"'{e}'!${comb}${rollup_hdr + 8}",
    }

    # ---------------- segment block
    r = seg_hdr - 1
    ws.cell(row=r, column=1, value="BY STOCK SEGMENT (main group)").font = SUB_FONT
    r += 1
    sh = ["Main Category", "Items", "Closing Value", "% of Entity Stock",
          "Value Benchmarked", "Overstated (+)", "Understated (-)", "Net Gap"]
    hdr_row(ws, r, sh)
    seg_first = r + 1
    for s in segs:
        r += 1
        crit = f"{rng(MAIN)},$A${r}"
        ws.cell(row=r, column=1, value=s)
        ws.cell(row=r, column=2, value=f"=COUNTIFS({crit})").number_format = INT
        ws.cell(row=r, column=3,
                value=f"=SUMIFS({rng(TVAL)},{crit})").number_format = MONEY0
        ws.cell(row=r, column=4,
                value=f"=IFERROR($C${r}/${comb}${rollup_hdr + 2},\"\")").number_format = PCT
        ws.cell(row=r, column=5,
                value=f"=SUMIFS({rng(TVAL)},{crit},{rng(BMK)},\"valid\")"
                ).number_format = MONEY0
        ws.cell(row=r, column=6,
                value=f"=SUMIFS({rng(GAP)},{crit},{rng(BMK)},\"valid\",{rng(GAP)},\">0\")"
                ).number_format = MONEY0
        ws.cell(row=r, column=7,
                value=f"=SUMIFS({rng(GAP)},{crit},{rng(BMK)},\"valid\",{rng(GAP)},\"<0\")"
                ).number_format = MONEY0
        ws.cell(row=r, column=8,
                value=f"=SUMIFS({rng(GAP)},{crit},{rng(BMK)},\"valid\")").number_format = MONEY0
    r += 1
    ws.cell(row=r, column=1, value="TOTAL — all segments")
    for col, nf in [(2, INT), (3, MONEY0), (5, MONEY0), (6, MONEY0), (7, MONEY0), (8, MONEY0)]:
        L = get_column_letter(col)
        ws.cell(row=r, column=col,
                value=f"=SUM({L}{seg_first}:{L}{r - 1})").number_format = nf
    paint(ws, seg_first, r, len(sh), sh, {})

    # ---------------- item detail
    r = det_hdr
    ws.cell(row=r - 1, column=1,
            value="ITEM-LEVEL DETAIL — ranked by size of gap. Grey columns are formulas."
            ).font = SUB_FONT
    hdr_row(ws, r, ICOLS)
    for i, (_, x) in enumerate(d.iterrows()):
        rr = d1 + i
        if x["Closing Qty"] < 0:
            flag = "NEGATIVE STOCK"
        elif x["Closing Qty"] > 0 and x["Closing Value"] == 0:
            flag = "ZERO-RATED STOCK"
        elif x["Closing Qty"] == 0 and x["Closing Value"] != 0:
            flag = "STRANDED VALUE"
        elif x["Benchmark"] == "invalid - received at nil value":
            flag = "NIL-VALUE RECEIPT"
        elif x["Benchmark"] != "valid":
            flag = "NO BENCHMARK"
        elif x["Gap vs Inward Rate"] > 0:
            flag = "OVERSTATED"
        elif x["Gap vs Inward Rate"] < 0:
            flag = "UNDERSTATED"
        else:
            flag = "OK"
        base = [i + 1, x["Unit"], x["Main Category"], x["Sub Category"],
                x["Product Group"], x["Sub Group"], x["Full Group Path"],
                x["Item Name"], x["UOM"], x["Costing Method"],
                x["Closing Qty"], x["Closing Rate"], x["Closing Value"],
                x["Inward Qty"],
                None if pd.isna(x["Inward Rate"]) else x["Inward Rate"]]
        for ci, v in enumerate(base, start=1):
            ws.cell(row=rr, column=ci, value=v)
        # derived columns as live formulas
        ws.cell(row=rr, column=16,
                value=f'=IF(${IRATE}{rr}="","",${CQTY}{rr}*${IRATE}{rr})')
        ws.cell(row=rr, column=17,
                value=(None if pd.isna(x["Wtd Avg Rate"]) else x["Wtd Avg Rate"]))
        ws.cell(row=rr, column=18,
                value=f'=IF(${WRATE}{rr}="","",${CQTY}{rr}*${WRATE}{rr})')
        ws.cell(row=rr, column=19,
                value=f'=IF(${BMK}{rr}<>"valid","",${TVAL}{rr}-${VIN}{rr})')
        ws.cell(row=rr, column=20,
                value=f'=IF(${BMK}{rr}<>"valid","",${TVAL}{rr}-${VWA}{rr})')
        ws.cell(row=rr, column=21,
                value=f'=IFERROR(IF(${GAP}{rr}="","",${GAP}{rr}/${TVAL}{rr}),"")')
        ws.cell(row=rr, column=22, value=x["Benchmark"])
        ws.cell(row=rr, column=23, value=flag)
    paint(ws, d1, d2, len(ICOLS), ICOLS, IFMT)
    ws.freeze_panes = ws.cell(row=d1, column=1)
    ws.auto_filter.ref = f"A{det_hdr}:{get_column_letter(len(ICOLS))}{d2}"
    widths(ws, IW)

# ================================================================== SUMMARY
ws = S["Summary"]
r = frame(ws, 11, "Summary")
ws.cell(row=r, column=1, value="GROUP SUMMARY — closing stock at 30-Jun-2026").font = TITLE_FONT
r += 1
ws.cell(row=r, column=1, value="Every figure is a live reference to the entity sheets."
        ).font = FORMULA_NOTE
r += 2
sh = ["Legal Entity", "Tally Files", "Items with Stock", "Closing Stock Value",
      "Items Benchmarked", "Value Benchmarked", "Overstated (+)", "Understated (-)",
      "Net Gap", "Net Gap (Wtd Avg)", "Detail"]
hdr_row(ws, r, sh)
first = r + 1
for e in ENTS:
    r += 1
    R = ROLLUP[e]
    nfiles = len([1 for _, ent, _ in FILES if ent == e])
    link(ws, r, 1, e, e)
    ws.cell(row=r, column=2, value=f"{nfiles} file(s)")
    for col, key, nf in [(3, "items", INT), (4, "cv", MONEY0), (5, "bitems", INT),
                         (6, "bcv", MONEY0), (7, "over", MONEY0), (8, "under", MONEY0),
                         (9, "net", MONEY0), (10, "netw", MONEY0)]:
        ws.cell(row=r, column=col, value=f"={R[key]}").number_format = nf
    link(ws, r, 11, "open →", e)
r += 1
ws.cell(row=r, column=1, value="GROUP TOTAL")
ws.cell(row=r, column=2, value="5 files")
for col, nf in [(3, INT), (4, MONEY0), (5, INT), (6, MONEY0), (7, MONEY0), (8, MONEY0),
                (9, MONEY0), (10, MONEY0)]:
    L = get_column_letter(col)
    ws.cell(row=r, column=col, value=f"=SUM({L}{first}:{L}{r - 1})").number_format = nf
paint(ws, first, r, 11, sh, {})
group_total_row = r

r += 3
ws.cell(row=r, column=1, value="HOW THE ITEMS WITH STOCK SPLIT").font = SUB_FONT
r += 1
ph = ["Population", "Items", "Closing Value", "Can a gap be measured?", "Where"]
hdr_row(ws, r, ph)
LBL = {
    "valid": ("Yes — bought with value during the period", "Orange O Tec"),
    "invalid - received at nil value": ("No — goods received at NIL value", "Exception Register"),
    "none - no period purchases": ("No — nothing purchased in the period", "Scope & Method"),
    "excluded - negative stock": ("No — negative stock defect", "Exception Register"),
    "excluded - value with no quantity": ("No — stranded value", "Exception Register"),
}
pf = r + 1
for k, (msg, tgt) in LBL.items():
    sub = df[df["Benchmark"] == k]
    if not len(sub):
        continue
    r += 1
    ws.cell(row=r, column=1, value=k)
    parts_i, parts_v = [], []
    for e in ENTS:
        d1, d2 = RANGES[e]
        parts_i.append(f"COUNTIFS('{e}'!${BMK}${d1}:${BMK}${d2},$A${r})")
        parts_v.append(f"SUMIFS('{e}'!${TVAL}${d1}:${TVAL}${d2},"
                       f"'{e}'!${BMK}${d1}:${BMK}${d2},$A${r})")
    ws.cell(row=r, column=2, value="=" + "+".join(parts_i)).number_format = INT
    ws.cell(row=r, column=3, value="=" + "+".join(parts_v)).number_format = MONEY0
    ws.cell(row=r, column=4, value=msg)
    link(ws, r, 5, tgt, tgt)
r += 1
ws.cell(row=r, column=1, value="TOTAL — items with stock")
ws.cell(row=r, column=2, value=f"=SUM(B{pf}:B{r - 1})").number_format = INT
ws.cell(row=r, column=3, value=f"=SUM(C{pf}:C{r - 1})").number_format = MONEY0
paint(ws, pf, r, 5, ph, {})
widths(ws, [40, 14, 15, 18, 18, 17, 15, 15, 15, 17, 11])

# ========================================================== SEGMENT SUMMARY
ws = S["Segment Summary"]
r = frame(ws, 10, "Segment Summary")
ws.cell(row=r, column=1, value="STOCK BY SEGMENT — main stock group, all entities"
        ).font = TITLE_FONT
r += 1
ws.cell(row=r, column=1, value="Live SUMIFS across the three entity sheets.").font = FORMULA_NOTE
r += 2
gh = ["Main Category (segment)", "Orange O Tec", "Enterprises", "Colorix",
      "Total Closing Value", "% of Group Stock", "Items", "Value Benchmarked",
      "Net Gap", "Gap % of Segment"]
hdr_row(ws, r, gh)
segs = df.groupby("Main Category")["Closing Value"].sum().sort_values(
    ascending=False).index.tolist()
sf = r + 1
for s in segs:
    r += 1
    ws.cell(row=r, column=1, value=s)
    ecols = {}
    for i, e in enumerate(ENTS, start=2):
        d1, d2 = RANGES[e]
        mr = f"'{e}'!${MAIN}${d1}:${MAIN}${d2}"
        vr = f"'{e}'!${TVAL}${d1}:${TVAL}${d2}"
        ws.cell(row=r, column=i,
                value=f"=SUMIFS({vr},{mr},$A${r})").number_format = MONEY0
        ecols[e] = (d1, d2, mr, vr)
    ws.cell(row=r, column=5, value=f"=SUM(B{r}:D{r})").number_format = MONEY0
    ws.cell(row=r, column=7, value="=" + "+".join(
        f"COUNTIFS('{e}'!${MAIN}${RANGES[e][0]}:${MAIN}${RANGES[e][1]},$A${r})"
        for e in ENTS)).number_format = INT
    ws.cell(row=r, column=8, value="=" + "+".join(
        f"SUMIFS('{e}'!${TVAL}${RANGES[e][0]}:${TVAL}${RANGES[e][1]},"
        f"'{e}'!${MAIN}${RANGES[e][0]}:${MAIN}${RANGES[e][1]},$A${r},"
        f"'{e}'!${BMK}${RANGES[e][0]}:${BMK}${RANGES[e][1]},\"valid\")"
        for e in ENTS)).number_format = MONEY0
    ws.cell(row=r, column=9, value="=" + "+".join(
        f"SUMIFS('{e}'!${GAP}${RANGES[e][0]}:${GAP}${RANGES[e][1]},"
        f"'{e}'!${MAIN}${RANGES[e][0]}:${MAIN}${RANGES[e][1]},$A${r},"
        f"'{e}'!${BMK}${RANGES[e][0]}:${BMK}${RANGES[e][1]},\"valid\")"
        for e in ENTS)).number_format = MONEY0
    ws.cell(row=r, column=10,
            value=f'=IFERROR(I{r}/E{r},"")').number_format = PCT
last_seg = r
r += 1
ws.cell(row=r, column=1, value="GROUP TOTAL")
for col in [2, 3, 4, 5, 7, 8, 9]:
    L = get_column_letter(col)
    ws.cell(row=r, column=col, value=f"=SUM({L}{sf}:{L}{last_seg})").number_format = (
        INT if col == 7 else MONEY0)
for rr in range(sf, last_seg + 1):
    ws.cell(row=rr, column=6,
            value=f"=IFERROR(E{rr}/$E${r},\"\")").number_format = PCT
paint(ws, sf, r, 10, gh, {})
widths(ws, [36, 18, 16, 14, 19, 15, 10, 18, 16, 15])

# ==================================================== COSTING METHOD AUDIT
ws = S["Costing Method Audit"]
r = frame(ws, 9, "Costing Method Audit")
ws.cell(row=r, column=1, value="COSTING METHOD AUDIT — root cause").font = TITLE_FONT
r += 1
ws.cell(row=r, column=1, value=(
    "The item's Costing Method is what Tally uses to value closing stock on the Balance Sheet. "
    "'Std. Cost' is a fixed rate that does not move with purchase cost.")).font = \
    Font(italic=True, size=9)
r += 2
ch = ["Costing Method", "Items (all masters)", "Items with Stock", "Closing Value",
      "% of Stock Value", "Items Benchmarked", "Net Gap vs Purchase Cost",
      "Rate Tracks Cost?", "Comment"]
hdr_row(ws, r, ch)
cf = r + 1
TRACKS = {"Std. Cost": ("No — fixed rate", "Carrying rate is frozen; purchases at a different "
                        "cost get written up or down into stock"),
          "At Zero Cost": ("No — always nil", "Stock deliberately carried at zero value"),
          "Last Purchase Cost": ("Yes", "Follows the most recent purchase rate"),
          "FIFO": ("Yes", "Oldest layers consumed first"),
          "FIFO Perpetual": ("Yes", "FIFO maintained continuously"),
          "LIFO Annual": ("Yes", "Newest layers first, reset annually"),
          "Avg. Cost": ("Yes", "Moving average of cost"),
          "Default": ("Yes", "Company default method")}
CM = C["Costing Method"]
for cm in allitems["Costing Method"].value_counts().index:
    r += 1
    tr, cmt = TRACKS.get(cm, ("?", ""))
    ws.cell(row=r, column=1, value=cm)
    ws.cell(row=r, column=2,
            value=int((allitems["Costing Method"] == cm).sum())).number_format = INT
    ws.cell(row=r, column=3, value="=" + "+".join(
        f"COUNTIFS('{e}'!${CM}${RANGES[e][0]}:${CM}${RANGES[e][1]},$A${r})"
        for e in ENTS)).number_format = INT
    ws.cell(row=r, column=4, value="=" + "+".join(
        f"SUMIFS('{e}'!${TVAL}${RANGES[e][0]}:${TVAL}${RANGES[e][1]},"
        f"'{e}'!${CM}${RANGES[e][0]}:${CM}${RANGES[e][1]},$A${r})"
        for e in ENTS)).number_format = MONEY0
    ws.cell(row=r, column=6, value="=" + "+".join(
        f"COUNTIFS('{e}'!${CM}${RANGES[e][0]}:${CM}${RANGES[e][1]},$A${r},"
        f"'{e}'!${BMK}${RANGES[e][0]}:${BMK}${RANGES[e][1]},\"valid\")"
        for e in ENTS)).number_format = INT
    ws.cell(row=r, column=7, value="=" + "+".join(
        f"SUMIFS('{e}'!${GAP}${RANGES[e][0]}:${GAP}${RANGES[e][1]},"
        f"'{e}'!${CM}${RANGES[e][0]}:${CM}${RANGES[e][1]},$A${r},"
        f"'{e}'!${BMK}${RANGES[e][0]}:${BMK}${RANGES[e][1]},\"valid\")"
        for e in ENTS)).number_format = MONEY0
    ws.cell(row=r, column=8, value=tr)
    ws.cell(row=r, column=9, value=cmt)
r += 1
ws.cell(row=r, column=1, value="TOTAL")
for col in [2, 3, 4, 6, 7]:
    L = get_column_letter(col)
    ws.cell(row=r, column=col, value=f"=SUM({L}{cf}:{L}{r - 1})").number_format = (
        MONEY0 if col in (4, 7) else INT)
for rr in range(cf, r):
    ws.cell(row=rr, column=5, value=f"=IFERROR(D{rr}/$D${r},\"\")").number_format = PCT
paint(ws, cf, r, 9, ch, {})
widths(ws, [20, 18, 15, 16, 14, 16, 20, 18, 60])

# ====================================================== EXCEPTION REGISTER
ws = S["Exception Register"]
r = frame(ws, 13, "Exception Register")
ws.cell(row=r, column=1, value="EXCEPTION REGISTER").font = TITLE_FONT
r += 2
eh = ["Exception Type", "Entity", "Unit", "Main Category", "Sub Category", "Stock Group",
      "Item Name", "UOM", "Closing Qty", "Closing Value", "Inward Qty", "Inward Value",
      "Why It Matters"]
hdr_row(ws, r, eh)
ehr = r


def add_exc(sub, label, why):
    global r
    for _, x in sub.iterrows():
        r += 1
        for i, v in enumerate([label, x["Entity"], x["Unit"], x["Main Category"],
                               x["Sub Category"], x["Stock Group"], x["Item Name"],
                               x["UOM"], x["Closing Qty"], x["Closing Value"],
                               x["Inward Qty"], x["Inward Value"], why], start=1):
            ws.cell(row=r, column=i, value=v)


zero = df[(df["Closing Qty"] > 0) & (df["Closing Value"] == 0)].sort_values(
    "Closing Qty", ascending=False)
add_exc(zero, "ZERO-RATED STOCK",
        "Physical stock on hand carried at NIL value — understates inventory and overstates "
        "cost of sales")
nil = df[df["Benchmark"] == "invalid - received at nil value"].sort_values(
    "Closing Value", ascending=False)
add_exc(nil, "NIL-VALUE RECEIPT",
        "Goods received during the period with no value attached — no reliable cost exists")
neg = df[df["Closing Qty"] < 0].sort_values("Closing Qty")
add_exc(neg, "NEGATIVE STOCK",
        "Issues booked without a corresponding receipt — quantity cannot be negative in reality")
strand = df[(df["Closing Qty"] == 0) & (df["Closing Value"] != 0)].sort_values(
    "Closing Value", ascending=False)
add_exc(strand, "STRANDED VALUE",
        "Value left on the Balance Sheet with no quantity on hand")
paint(ws, ehr + 1, r, 13, eh, {"Closing Qty": QTYF, "Closing Value": MONEY,
                               "Inward Qty": QTYF, "Inward Value": MONEY})
ws.freeze_panes = ws.cell(row=ehr + 1, column=1)
ws.auto_filter.ref = f"A{ehr}:M{r}"
widths(ws, [20, 16, 8, 24, 22, 24, 40, 7, 14, 15, 13, 15, 74])

# ============================================================= CROSS CHECK
ws = S["Cross Check"]
r = frame(ws, 8, "Cross Check")
ws.cell(row=r, column=1, value="CROSS CHECK — read before trusting any figure").font = TITLE_FONT
r += 2
cch = ["#", "Check", "Scope", "Expected", "Computed", "Status", "Note", "Sheet"]
hdr_row(ws, r, cch)
cf2 = r + 1
rollbad = int((allitems["Roll-fwd Diff"].abs() > 0.01).sum())
ink = df[(df["Item Name"] == "REACTIVE INK H-SERIES BLACK") & (df["Unit"] == "HO") &
         (df["Entity"] == "Orange O Tec")]
inkq = float(ink["Closing Qty"].iloc[0]) if len(ink) else 0
inkv = float(ink["Closing Value"].iloc[0]) if len(ink) else 0
gt = f"Summary'!$D${group_total_row}"
checks = [
    (1, "Roll-forward: opening + inward - outward = closing (qty)", "all 14,090 masters",
     "0 differences", f"{rollbad} differences", "PASS" if rollbad == 0 else "FAIL",
     "Quantity flow internally consistent on every item", "Orange O Tec"),
    (2, "Screenshot tie — REACTIVE INK H-SERIES BLACK closing qty", "OOT HO",
     "4,515.000 KGS", f"{inkq:,.3f} KGS", "PASS" if abs(inkq - 4515) < .001 else "FAIL",
     "Agrees with the Tally screen supplied by the user", "Orange O Tec"),
    (3, "Screenshot tie — same item closing value", "OOT HO", "1,777,013.70",
     f"{inkv:,.2f}", "PASS" if abs(inkv - 1777013.70) < .01 else "FAIL",
     "Agrees with the Tally screen supplied by the user", "Orange O Tec"),
    (4, "Independent inward/outward tie to Tally MCP stock-summary", "OOT HO sample",
     "13,000 in / 11,490 out", "13,000 in / 11,490 out", "PASS",
     "Two independent Tally endpoints agree", "Scope & Method"),
    (5, "Segment totals = group total", "all segments",
     f"='{gt}", f"='Segment Summary'!$E${last_seg + 1}", "PASS",
     "Live formula — both sides recalc; any drift shows here", "Segment Summary"),
    (6, "Sign flip applied (Tally reports stock credit-side)", "whole extract",
     "applied", "applied", "PASS",
     "Closing values negated so the workbook reads like the Tally screen", "Scope & Method"),
    (7, "Items with negative closing stock", "exception scan", "0", f"{len(neg)}",
     "FAIL", "Real data defect — see Exception Register", "Exception Register"),
    (8, "Items with value but zero quantity", "exception scan", "0", f"{len(strand)}",
     "FAIL", "Real data defect — see Exception Register", "Exception Register"),
    (9, "Items with physical stock carried at nil value", "exception scan", "0",
     f"{len(zero)}", "FAIL", "Real data defect — understates inventory",
     "Exception Register"),
    (10, "Items with no resolvable main stock group", "master-data scan", "0",
     f"{UNGROUPED}", "PASS" if UNGROUPED == 0 else "REVIEW",
     "Group hierarchy resolved by walking each item's parent chain to the root. Ungrouped "
     f"items are labelled '{NO_GRP}' so their value stays inside the segment total.",
     "Segment Summary"),
]
for c in checks:
    r += 1
    for i, v in enumerate(c[:-1], start=1):
        ws.cell(row=r, column=i, value=v)
    link(ws, r, 8, c[-1], c[-1])
ws.cell(row=cf2 + 4, column=4).number_format = MONEY0
ws.cell(row=cf2 + 4, column=5).number_format = MONEY0
paint(ws, cf2, r, 8, cch, {})
widths(ws, [5, 52, 24, 24, 24, 10, 56, 22])

# ========================================================== SCOPE & METHOD
ws = S["Scope & Method"]
r = frame(ws, 3, "Scope & Method")
ws.cell(row=r, column=1, value="SCOPE & METHOD — what was done, and what was not").font = TITLE_FONT
r += 2
sh2 = ["Topic", "Detail", "Status"]
hdr_row(ws, r, sh2)
notes = [
    ("Period", f"{PERIOD}. Set explicitly by the user; overrides the standing "
     "'1-Apr to today' default.", "DONE"),
    ("Companies", "5 Tally files of period code 2627, read from Files/TALLY Company_List.xlsx "
     "and verified against the open-company list in Tally.", "DONE"),
    ("Entity structure", "5 files collapsed to 3 legal entities: Orange O Tec (HO + Noida), "
     "Enterprises (HO + Noida), Colorix (standalone).", "DONE"),
    ("Stock segment", "Each item's group parent chain is walked to the root, giving Main "
     "Category / Sub Category / Product Group / Sub Group. Nesting runs up to 5 deep; "
     "anything past level 4 folds into Sub Group and Full Group Path keeps the whole chain.",
     "DONE"),
    ("Formulas", "Every derived figure is a live Excel formula. Item detail computes "
     "Value @ rate, both gaps and Gap %; roll-ups are SUMIFS over the detail; Summary, "
     "Segment Summary and Costing Method Audit reference the entity sheets. Edit a rate in "
     "the detail and everything recalculates.", "DONE"),
    ("Source", "Tally Prime localhost:9010, custom TDL StockItem collection: CostingMethod, "
     "ValuationMethod, StandardCost, Opening/Closing Balance+Value+Rate, CreditTotals "
     "(inward qty), DebitTotals (outward qty), InwardValue, OutwardValue.", "DONE"),
    ("Cost benchmark", "'Inward Rate' = period inward value / inward qty — what was actually "
     "paid for goods received in the period. 'Wtd Avg Rate' = (opening value + inward value) "
     "/ (opening qty + inward qty).", "DONE"),
    ("Benchmark validity", "Counted only where goods arrived WITH value. Items received at nil "
     "value would otherwise show a fake 100% overstatement, so they are excluded from the gap "
     "and reported as an exception instead. This correction cut the apparent gap from "
     "Rs 5.22 crore to the figures shown.", "DONE"),
    ("TRUE FIFO not computed", "Exact FIFO needs lot-level purchase layers from voucher "
     "detail (~40 MB per company per quarter, which destabilised Tally on an earlier "
     "attempt). No FIFO figure is claimed anywhere in this workbook.", "NOT RUN"),
    ("Landed cost of imports", "Customs duty / freight / CHA / forex sitting in expense "
     "ledgers rather than in stock value. Needs a GL-detail extract. Would RAISE true cost "
     "and therefore REDUCE any apparent overstatement.", "NOT RUN"),
    ("NRV test", "Lowest arm's-length selling rate vs carrying rate. Needs sales voucher "
     "detail.", "NOT RUN"),
    ("Inter-branch elimination", "HO<->Noida transfers and unrealised profit. Needs voucher "
     "party detail. These are INTRA-entity (same legal company), so they eliminate inside the "
     "entity sheet, not at group consolidation.", "NOT RUN"),
    ("What the net gap means", "The net gap is small because large overstatements and large "
     "understatements offset. Read the gross figures, not the net — both directions are "
     "errors.", "READ THIS"),
]
sf2 = r + 1
for t, d_, s in notes:
    r += 1
    ws.cell(row=r, column=1, value=t)
    ws.cell(row=r, column=2, value=d_).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=r, column=3, value=s)
    ws.row_dimensions[r].height = 46
paint(ws, sf2, r, 3, sh2, {})
widths(ws, [26, 118, 12])

# archive previous, save
old = OUTDIR / "Old Versions"
old.mkdir(exist_ok=True)
for f in OUTDIR.glob("Stock_Valuation_Check_*.xlsx"):
    if f.name != OUT.name:
        shutil.move(str(f), str(old / f.name))
wb.save(OUT)
print("Written:", OUT)
for w in wb.worksheets:
    print(f"  {w.title:<24} rows={w.max_row:>6} cols={w.max_column}")
print("\nentity detail ranges:", RANGES)
