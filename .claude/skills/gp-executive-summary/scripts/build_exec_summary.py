# -*- coding: utf-8 -*-
import argparse
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Derived from this file's location (.../PL and BS/.claude/skills/<skill>/scripts/<file>.py)
# so the script keeps working regardless of which user profile the project sits under.
PROJECT_ROOT = str(Path(__file__).resolve().parents[4])
DEFAULT_SRC = PROJECT_ROOT + r"\Files\GP REPORT.xlsx"

parser = argparse.ArgumentParser(description="Add/refresh the Executive Summary sheet inside GP REPORT.xlsx.")
parser.add_argument("--source", default=DEFAULT_SRC, help="Path to GP REPORT.xlsx (read AND written in place)")
args = parser.parse_args()

SRC = args.source

wb = openpyxl.load_workbook(SRC)

# Remove old Executive Summary sheet if re-running
if "Executive Summary" in wb.sheetnames:
    del wb["Executive Summary"]

ws = wb.create_sheet("Executive Summary", 0)

# ---- Styles ----
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
SUBTITLE_FONT = Font(italic=True, size=10, color="595959")
SECTION_FONT = Font(bold=True, size=12, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(bold=True, size=10, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
TOTAL_FONT = Font(bold=True, size=10)
TOTAL_FILL = PatternFill("solid", fgColor="D9E2F3")
NOTE_FONT = Font(italic=True, size=9, color="7F7F7F")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CUR = '#,##0'
PCT = '0.0%'
PCTPT = '+0.0pt;-0.0pt'

r = 1

def title(text, row):
    c = ws.cell(row=row, column=1, value=text)
    c.font = TITLE_FONT
    return row + 1

def subtitle(text, row):
    c = ws.cell(row=row, column=1, value=text)
    c.font = SUBTITLE_FONT
    return row + 1

def section(text, row, span=6):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = SECTION_FONT
    c.fill = SECTION_FILL
    c.alignment = Alignment(vertical="center")
    for col in range(1, span + 1):
        ws.cell(row=row, column=col).fill = SECTION_FILL
    return row + 1

def header_row(headers, row):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    return row + 1

def data_row(values, row, formats=None, bold=False, fill=None):
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.border = BORDER
        if formats and formats.get(i):
            c.number_format = formats[i]
        if bold:
            c.font = TOTAL_FONT
        if fill:
            c.fill = fill
        if i == 1:
            c.alignment = Alignment(horizontal="left")
        else:
            c.alignment = Alignment(horizontal="right")
    return row + 1

def note(text, row, span=6):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = NOTE_FONT
    c.alignment = Alignment(wrap_text=True, vertical="top")
    return row + 1

def blank(row, n=1):
    return row + n

# ================= TITLE =================
r = title("Executive Summary — Gross Profit Performance", r)
r = subtitle("Basis: GP REPORT.xlsx  |  Comparison period: Q1 FY26-27 (Apr-Jun 2026) vs Q1 FY25-26 (Apr-Jun 2025) — like-for-like quarters", r)
r = note("Note: The workbook's \"25-26 YTD\" column is a full 12-month figure while \"26-27 YTD\" covers only Q1 so far — the two are not directly comparable. "
         "This summary uses the workbook's Q1-vs-Q1 columns throughout for a fair year-on-year read. Figures below are net of discounts unless marked \"pre-discount\".", r)
r = blank(r)

# ================= SECTION 1: OVERALL =================
r = section("1. Overall Performance", r, span=4)
r = header_row(["Metric", "Q1 FY25-26", "Q1 FY26-27", "Change"], r)
r = data_row(["Net Sales (₹)", 368875107, 353639685, -0.041], r, {2: CUR, 3: CUR, 4: PCT})
r = data_row(["Gross Profit (₹)", 108750748, 113315981, 0.042], r, {2: CUR, 3: CUR, 4: PCT})
r = data_row(["Gross Profit %", 0.2948, 0.3204, 0.0256], r, {2: PCT, 3: PCT, 4: PCTPT}, bold=True, fill=TOTAL_FILL)
r = note("Sales dipped ~4% but gross profit rose ~4% — margin expansion (Ink-led) more than offset the sales decline. Net effect: GP% improved by roughly 2.6 percentage points.", r, span=4)
r = blank(r)

# ================= SECTION 2: SEGMENT VIEW =================
r = section("2. Segment-wise Performance", r, span=6)
r = header_row(["Segment", "Sales Q1 25-26 (₹)", "Sales Q1 26-27 (₹)", "Sales Chg %", "GP% Q1 25-26", "GP% Q1 26-27"], r)
r = data_row(["Ink", 279263140, 284292775, 0.018, 0.3253, 0.3736], r, {2: CUR, 3: CUR, 4: PCT, 5: PCT, 6: PCT})
r = data_row(["Spare Parts", 19211708, 17278684, -0.101, 0.3837, 0.2778], r, {2: CUR, 3: CUR, 4: PCT, 5: PCT, 6: PCT})
r = data_row(["Heads", 70400260, 52068226, -0.260, 0.1496, 0.0445], r, {2: CUR, 3: CUR, 4: PCT, 5: PCT, 6: PCT})
r = data_row(["Total", 368875107, 353639685, -0.041, 0.2948, 0.3204], r, {2: CUR, 3: CUR, 4: PCT, 5: PCT, 6: PCT}, bold=True, fill=TOTAL_FILL)
r = note("Ink is ~80% of sales and the sole growth engine, with margin up ~4.8 points. Spare Parts and Heads both lost sales AND margin — Heads especially severely (GP% fell from 15.0% to 4.5%).", r, span=6)
r = note("Machine Sale (units only: 6 / 22 / 11 in Apr/May/Jun 2026) has no revenue or GP captured in this report — flagged as a data gap in Section 5, not included in totals above.", r, span=6)
r = blank(r)

# ================= SECTION 3: VOLUME / QUANTITY TRENDS =================
r = section("3. Volume / Quantity Trends", r, span=7)
r = header_row(["Segment", "Qty Q1 25-26", "Qty Q1 26-27", "Qty Chg %", "Realization/Unit 25-26 (₹)", "Realization/Unit 26-27 (₹)", "Realization Chg %"], r)
r = data_row(["Ink (kg/ltr)", 418787.25, 483317.566, 0.1541, 666.86, 588.16, -0.1180], r, {2: '#,##0.0', 3: '#,##0.0', 4: PCT, 5: CUR, 6: CUR, 7: PCT})
r = data_row(["Heads (nos.)", 346, 279, -0.1936, 203470, 186624, -0.0828], r, {2: CUR, 3: CUR, 4: PCT, 5: CUR, 6: CUR, 7: PCT})
r = data_row(["Spare Parts (mixed units)*", 2297.1, 11867.8, 4.1662, 8363, 1456, None], r, {2: '#,##0.0', 3: '#,##0.0', 4: PCT, 5: CUR, 6: CUR})
r = note("* Spare Parts quantities span many different unit types (single components through full assemblies), so the blended quantity/realization figures are not meaningful on a standalone basis — see Section 5.", r, span=7)
r = blank(r)
r = header_row(["Segment", "April 2026", "May 2026", "June 2026", "Trend", "", ""], r)
r = data_row(["Machine (nos.)", 6, 22, 11, "No clear trend — small base, high month-to-month variance", "", ""], r, {2: CUR, 3: CUR, 4: CUR})
r = data_row(["Ink (kg/ltr)", 175585.7, 140398.1, 167333.3, "Dipped in May, recovered by June", "", ""], r, {2: '#,##0.0', 3: '#,##0.0', 4: '#,##0.0'})
r = data_row(["Heads (nos.)", 81, 90, 108, "Steady month-on-month growth", "", ""], r, {2: CUR, 3: CUR, 4: CUR})
r = data_row(["Spare Parts (mixed units)", 720, 483.3, 10666.5, "Sharp spike in June — verify genuine demand vs. a unit-of-measure/data-entry anomaly", "", ""], r, {2: CUR, 3: '#,##0.0', 4: '#,##0.0'})
r = note("Ink volume grew ~15% year-on-year while realization per unit fell ~12% — a volume-led quarter where cost efficiency gains preserved and expanded margin despite softer pricing (ties back to Section 2's Ink GP% gain). "
         "Heads lost on both volume (-19%) and price (-8%) simultaneously — a more concerning combination than the margin figure in Section 2 alone suggests.", r, span=7)
r = blank(r)

# ================= SECTION 4: KEY DRIVERS =================
r = section("4. Key Drivers & Anomalies", r, span=5)
r = header_row(["Area", "Q1 25-26", "Q1 26-27", "Change", "What's driving it"], r)
r = data_row(["Reactive Ink GP%", 0.3459, 0.3991, 0.0532, "Largest ink line (₹198.7M sales); margin expansion added ~₹8.4M gross profit"], r, {2: PCT, 3: PCT, 4: PCTPT})
r = data_row(["Sublimation Ink GP%", 0.3372, 0.4053, 0.0681, "Second-largest ink line; margin expansion added ~₹6.7M gross profit"], r, {2: PCT, 3: PCT, 4: PCTPT})
r = data_row(["Pigment Ink GP%", 0.3038, 0.0918, -0.2120, "Margin collapsed to near cost — small volume (~₹3.2M sales) but a sharp cost/price signal"], r, {2: PCT, 3: PCT, 4: PCTPT})
r = data_row(["Disperse Ink GP%", 0.3474, 0.2491, -0.0984, "Margin down ~10 points on ~₹10.6M sales"], r, {2: PCT, 3: PCT, 4: PCTPT})
r = data_row(["Head \"OTHER\" category", 5243600, 0, "Gone", "₹5.24M gross profit in the base quarter vanished entirely — this alone explains ~64% of the Heads segment's GP decline"], r, {2: CUR, 3: CUR})
r = data_row(["Head discounts (pre-GP)", -3645449, -6319653, 0.733, "Discounting nearly doubled — explains most of the remaining Heads GP decline"], r, {2: CUR, 3: CUR, 4: PCT})
r = data_row(["Spare Parts discounts", -573910, -1725672, 2.007, "Discounting roughly tripled — the single largest drag on Spare Parts margin"], r, {2: CUR, 3: CUR, 4: PCT})
r = data_row(["MS Spare Parts GP%", 0.3408, 0.2655, -0.0753, "Largest spare-parts line (₹9.2M sales); margin compression here also weighs on the segment"], r, {2: PCT, 3: PCT, 4: PCTPT})
r = blank(r)

# ================= SECTION 5: DATA GAPS =================
r = section("5. Data Gaps & Caveats", r, span=3)
r = header_row(["Item", "Issue", "Recommendation"], r)
r = data_row(["Machine Sale", "Quantities recorded (6/22/11 units, Apr-Jun 2026) but no sale value or GP captured", "Add machine sale value/GP capture to the report for a complete revenue picture"], r)
r = data_row(["YTD columns", "\"25-26 YTD\" = full 12 months; \"26-27 YTD\" = Q1 only — not comparable at face value", "Label these columns clearly, or replace with a rolling like-for-like period in future versions"], r)
r = data_row(["Discount lines", "Discounts sit as separate negative-GP rows rather than netted into each product line", "Fine for totals, but obscures which specific product discounts are moving — consider tagging by category"], r)
r = data_row(["Spare Parts quantity", "Aggregates many different unit types (individual parts through full assemblies), so blended qty/realization is not meaningful", "Break out Section 3's Spare Parts volume by sub-category if unit-level analysis is needed"], r)
r = blank(r)

# ================= SECTION 6: RECOMMENDED ACTIONS =================
r = section("6. Watch Items / Recommended Actions", r, span=4)
r = header_row(["Priority", "Area", "Recommended Action", "Rationale"], r)
r = data_row(["High", "Heads", "Confirm whether the \"OTHER\" head category (₹21.9M sales / ₹5.24M GP last year) was discontinued, reclassified, or is a lost account", "Single largest identifiable driver of the Heads segment's margin collapse"], r)
r = data_row(["High", "Pigment Ink", "Review pricing and input cost on pigment ink", "GP% fell from 30.4% to 9.2% — steep enough to suggest a cost pass-through gap"], r)
r = data_row(["Medium", "Spare Parts", "Review the ~3x rise in discounting and MS Spare Parts pricing/costing", "Discount growth outpaces sales decline; margin erosion is broad-based, not one-off"], r)
r = data_row(["Medium", "Heads", "Review discount approval thresholds", "Discounting nearly doubled year-on-year on a shrinking revenue base"], r)
r = data_row(["Low", "Machine Sale", "Add revenue/GP tracking for machine sales to this report", "Currently invisible in the GP view despite units being sold"], r)
r = data_row(["Medium", "Spare Parts", "Verify the June quantity spike (720 → 483 → 10,666.5 units across Apr/May/Jun)", "Confirm it's genuine demand and not a unit-of-measure or data-entry anomaly before acting on it"], r)

# ---- Column widths ----
widths = {1: 30, 2: 20, 3: 20, 4: 16, 5: 45, 6: 20, 7: 20}
for col, w in widths.items():
    ws.column_dimensions[get_column_letter(col)].width = w

ws.sheet_view.showGridLines = False
ws.freeze_panes = "A5"

wb.save(SRC)
print("Saved Executive Summary sheet into:", SRC)
