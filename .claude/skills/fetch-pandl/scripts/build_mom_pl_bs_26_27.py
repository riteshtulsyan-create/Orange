"""Builds Files/MOM PL and BS 26-27.xlsx - the standalone month-on-month P&L / Balance Sheet
deliverable for FY 2026-27.

WHY THIS EXISTS
---------------
This workbook used to have NO builder at all. It was hand-assembled once (01-Aug-2026) out of the
"SAMPLE_New_Sheets_Preview_v4" exploration, so "refresh it" was not a runnable operation - the
fetch-pandl pipeline rebuilds Files/PL_BS_GL_Detail_<ts>.xlsx, which has a completely different
sheet set. Asked to refresh this file on 04-Aug-2026, the only honest options were to hand-patch
it again or to write this. See build_mom_gl_detail.py's note near build_pl_mom_extended_sheet:
those Main/Sub/Full-GL sheets were deliberately removed from the production workbook (the user
found 39 tabs overwhelming) and continued life here instead, as a separate, focused deliverable.

RELATIONSHIP TO THE REST OF THE PIPELINE
----------------------------------------
This script is a PRESENTATION layer only. It performs no fetching and owns no figures: it reads
the same gl_data/ + gl_data_mom_detail/ CSVs and the same classification/adjustment constants that
build_gl_detail_sidebyside.py and build_mom_gl_detail.py use, via build_multi_period_tree(). Run
fetch_gl_data.py first (Step A of CFO mode); this is a downstream consumer, exactly like
build_mom_gl_detail.py is.

    python fetch_gl_data.py --all
    python build_mom_pl_bs_26_27.py

Because every figure traces to those shared constants, this workbook can never disagree with
PL_BS_GL_Detail_<ts>.xlsx about a number - if it does, that is a bug in this file, not a
difference of opinion, and the Cross Check sheet is what surfaces it.

SHEET DESIGN (three levels, each linked UP from the one below - never a re-keyed literal)
----------------------------------------------------------------------------------------
    <PL|BS> Full GL MOM    ledger-level. The ONLY sheet holding literal figures.
    <PL|BS> Sub Group MOM  one row per sub-group, = a link to that sub-group's TOTAL row on Full GL
    <PL|BS> Main Group MOM one row per main group, = a link to that group's TOTAL row on Sub Group

so a reader can start at any level and drill down without the levels ever disagreeing. Per-company
columns are links/literals; the Total Orange O Tec / Total Enterprises / Grand Total columns are
always same-row arithmetic on THIS sheet's own CO1..CO5 (matching write_period_cells' convention
in build_gl_detail_sidebyside.py), never copied from another sheet.

PERIOD BLOCKS come straight from build_mom_gl_detail.PL_PERIODS / BS_PERIODS, so this file
automatically follows the pipeline as the year extends - adding Aug-26 there adds it here with no
edit to this script. P&L carries both cumulative columns (Q1 Total, Apr-Jul Total) because they
are independent fetches, not derived sums; Balance Sheet carries none, because a balance sheet is
a position and the Jul-26 column already IS the as-at-31-Jul position.
"""

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))
import build_gl_detail_sidebyside as b          # noqa: E402
import build_mom_gl_detail as m                 # noqa: E402

FILES_DIR = SCRIPT_DIR.parent.parent.parent.parent / "Files"
OUTPUT_NAME = "MOM PL and BS 26-27.xlsx"
ARCHIVE_DIR = FILES_DIR / "Old Version of PL BS"

THIN = Side(style="thin", color="D9D9D9")
BORDER_THIN = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TITLE_FONT = Font(bold=True, size=13, name="Calibri", color="1F3864")
NOTE_FONT = Font(italic=True, size=9, name="Calibri", color="595959")
WARN_FONT = Font(italic=True, bold=True, size=9, name="Calibri", color="C00000")
TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
GAP_FILL = PatternFill("solid", fgColor="FCE4D6")

HEADER_ROW = 5
FIRST_DATA_ROW = 6

# P&L main-group layout: (display section, [bucket, ...]). Tree section keys on the left,
# the label the user reads on the right - 'OPERATING EXPENSES' shows as 'INDIRECT COST'.
PL_SECTIONS = [
    ("TURNOVER", "TURNOVER", ["Sales Accounts", "Direct Incomes"]),
    ("OTHER INCOME", "OTHER INCOME", ["Indirect Incomes"]),
    ("COST OF SALES", "COST OF SALES",
     ["Purchase Accounts", "Direct Expenses", "Stock Valuation (Inventory)"]),
    ("OPERATING EXPENSES", "INDIRECT COST", ["Indirect Expenses"]),
    ("BRANCH TRANSACTIONS", "BRANCH TRANSACTIONS (memo)", ["Branch Sales", "Branch Purchase"]),
    ("FINANCE COST", "FINANCE COST", ["Finance Cost"]),
]
# SIGN CONVENTION - the single easiest thing to get wrong in this file.
#
# build_multi_period_tree() already multiplies every leaf by b.display_sign(bucket), so what
# reaches this script is a DISPLAY MAGNITUDE: income and liabilities positive, and expenses and
# assets ALSO positive (they are shown as positive numbers under a "LESS:" heading, not as
# negatives). That is enough for every SUM-based subtotal, because a subtotal only ever adds
# leaves that share one bucket - but it means the headline waterfall rows MUST re-apply the sign
# themselves. Summing bucket totals blindly gives nonsense (caught here for real: PAT came out at
# Rs.1,940cr against Rs.62cr of turnover, because Cost of Sales was being added to income instead
# of subtracted). Mirrors write_multi_computed()'s explicit "+"/"-" operand list in
# build_mom_gl_detail.py - keep the two in step.
EBITDA_ADD = ["Sales Accounts", "Direct Incomes", "Indirect Incomes", "Branch Sales"]
EBITDA_LESS = ["Purchase Accounts", "Direct Expenses", "Stock Valuation (Inventory)",
               "Indirect Expenses", "Branch Purchase"]
RP_BUCKETS = ["Related Party Sales", "Related Party Purchase"]


def block_start(n_label_cols, idx):
    return n_label_cols + 1 + idx * 8


def col(n_label_cols, idx, offset):
    return get_column_letter(block_start(n_label_cols, idx) + offset)


def write_headers(ws, n_label_cols, label_names, periods, title, subtitle, warning=None):
    ws.cell(1, 1, value=title).font = TITLE_FONT
    ws.cell(2, 1, value=subtitle).font = NOTE_FONT
    if warning:
        ws.cell(3, 1, value=warning).font = WARN_FONT
    for i, name in enumerate(label_names, start=1):
        c = ws.cell(HEADER_ROW, i, value=name)
        c.font = m.NEW_HEADER_FONT
        c.fill = m.NEW_HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for bidx, (_key, plabel) in enumerate(periods):
        start = block_start(n_label_cols, bidx)
        cumulative = _key in ("qtr", "ytd")
        pc = ws.cell(HEADER_ROW - 1, start, value=plabel)
        pc.font = m.NEW_PERIOD_FONT
        pc.fill = m.NEW_QTR_FILL if cumulative else m.NEW_PERIOD_FILL
        for i, clabel in enumerate(m.NEW_COMPANY_LABELS):
            c = ws.cell(HEADER_ROW, start + i, value=clabel)
            c.font = m.NEW_HEADER_FONT
            c.fill = m.NEW_HEADER_FILL
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(start + i)].outline_level = m.NEW_OUTLINE_LEVELS[i]
    ws.row_dimensions[HEADER_ROW - 1].height = 16
    ws.row_dimensions[HEADER_ROW].height = 26
    ws.sheet_properties.outlinePr.summaryRight = True
    ws.freeze_panes = ws.cell(FIRST_DATA_ROW, n_label_cols + 1).coordinate
    for i in range(1, n_label_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = [16, 26, 30, 40][i - 1] if i <= 4 else 18
    for bidx in range(len(periods)):
        start = block_start(n_label_cols, bidx)
        for i in range(8):
            ws.column_dimensions[get_column_letter(start + i)].width = 17


def write_totals(ws, row, n_label_cols, periods, font):
    """Total Orange / Total Enterprises / Grand Total, same-row arithmetic on this sheet."""
    for bidx in range(len(periods)):
        s = block_start(n_label_cols, bidx)
        co1, co2 = get_column_letter(s), get_column_letter(s + 1)
        co3, co4 = get_column_letter(s + 3), get_column_letter(s + 4)
        co5 = get_column_letter(s + 6)
        to, te = get_column_letter(s + 2), get_column_letter(s + 5)
        ws.cell(row, s + 2, value=f"={co1}{row}+{co2}{row}")
        ws.cell(row, s + 5, value=f"={co3}{row}+{co4}{row}")
        ws.cell(row, s + 7, value=f"={to}{row}+{te}{row}+{co5}{row}")
        for off in (2, 5, 7):
            c = ws.cell(row, s + off)
            c.number_format = b.NUM_FORMAT
            c.font = font


def style_row(ws, row, n_label_cols, periods, font, fill=None):
    for c in range(1, block_start(n_label_cols, len(periods))):
        cell = ws.cell(row, c)
        cell.border = BORDER_THIN
        if fill:
            cell.fill = fill
        if c > n_label_cols:
            cell.font = font
            cell.number_format = b.NUM_FORMAT


def section_header(ws, row, n_label_cols, periods, label, memo=False):
    c = ws.cell(row, 1, value=label)
    c.font = m.NEW_MEMO_FONT if memo else m.NEW_SECTION_FONT
    fill = m.NEW_MEMO_FILL if memo else m.NEW_SECTION_FILL
    for cc in range(1, block_start(n_label_cols, len(periods))):
        ws.cell(row, cc).fill = fill
    return row + 1


# ---------------------------------------------------------------------------
# Level 1: Full GL (the only sheet with literals)
# ---------------------------------------------------------------------------

def write_full_gl(ws, tree, sections, periods, title, warning):
    write_headers(ws, 4, ["Section", "Grouping", "Sub Grouping", "GL Name"], periods, title,
                  "Raw ledger figures (literals). Sub Group and Main Group link UP from this sheet - "
                  "this is the only level holding numbers, so the three levels can never disagree.",
                  warning)
    row = FIRST_DATA_ROW
    sub_total_rows = {}   # (bucket, sub_grp) -> row on this sheet
    for tree_key, display, buckets in sections:
        row = section_header(ws, row, 4, periods, display,
                             memo="memo" in display.lower())
        for bucket in buckets:
            leaves = tree.get(tree_key, {}).get(bucket, {})
            by_sub = {}
            for (sub_grp, gl_name), pdata in leaves.items():
                by_sub.setdefault(sub_grp, []).append((gl_name, pdata))
            for sub_grp in sorted(by_sub):
                first = row
                for gl_name, pdata in sorted(by_sub[sub_grp]):
                    ws.cell(row, 2, value=bucket).font = m.NEW_LABEL_FONT
                    ws.cell(row, 3, value=sub_grp).font = m.NEW_LABEL_FONT
                    ws.cell(row, 4, value=gl_name).font = m.NEW_LABEL_FONT
                    for bidx, (pkey, _) in enumerate(periods):
                        vals = pdata.get(pkey, [0.0] * 5)
                        s = block_start(4, bidx)
                        for i, off in enumerate((0, 1, 3, 4, 6)):
                            cell = ws.cell(row, s + off, value=round(vals[i], 2))
                            cell.number_format = b.NUM_FORMAT
                            cell.font = m.NEW_LABEL_FONT
                    write_totals(ws, row, 4, periods, m.NEW_LABEL_FONT)
                    style_row(ws, row, 4, periods, m.NEW_LABEL_FONT)
                    row += 1
                ws.cell(row, 2, value=bucket).font = m.NEW_TOTAL_FONT
                ws.cell(row, 3, value=f"TOTAL: {sub_grp}").font = m.NEW_TOTAL_FONT
                for bidx in range(len(periods)):
                    s = block_start(4, bidx)
                    for off in (0, 1, 3, 4, 6):
                        L = get_column_letter(s + off)
                        cell = ws.cell(row, s + off,
                                       value=(f"=SUM({L}{first}:{L}{row - 1})" if row > first else 0))
                        cell.number_format = b.NUM_FORMAT
                        cell.font = m.NEW_TOTAL_FONT
                write_totals(ws, row, 4, periods, m.NEW_TOTAL_FONT)
                style_row(ws, row, 4, periods, m.NEW_TOTAL_FONT, TOTAL_FILL)
                sub_total_rows[(bucket, sub_grp)] = row
                row += 1
        row += 1
    return sub_total_rows


# ---------------------------------------------------------------------------
# Level 2: Sub Group (links to Full GL)
# ---------------------------------------------------------------------------

def write_sub_group(ws, sections, periods, sub_total_rows, source, title, warning):
    write_headers(ws, 3, ["Section", "Grouping", "Sub Grouping"], periods, title,
                  f"Every figure is a live link to '{source}' - one row per sub-group. "
                  "Nothing here is re-keyed.", warning)
    row = FIRST_DATA_ROW
    bucket_total_rows = {}
    for _tree_key, display, buckets in sections:
        row = section_header(ws, row, 3, periods, display, memo="memo" in display.lower())
        for bucket in buckets:
            subs = sorted(s for (bk, s) in sub_total_rows if bk == bucket)
            member_rows = []
            for sub_grp in subs:
                src_row = sub_total_rows[(bucket, sub_grp)]
                ws.cell(row, 2, value=bucket).font = m.NEW_LABEL_FONT
                ws.cell(row, 3, value=sub_grp).font = m.NEW_LABEL_FONT
                for bidx in range(len(periods)):
                    s_dest = block_start(3, bidx)
                    s_src = block_start(4, bidx)
                    for i, off in enumerate((0, 1, 3, 4, 6)):
                        src_letter = get_column_letter(s_src + off)
                        cell = ws.cell(row, s_dest + off,
                                       value=f"='{source}'!{src_letter}{src_row}")
                        cell.number_format = b.NUM_FORMAT
                        cell.font = m.NEW_LABEL_FONT
                write_totals(ws, row, 3, periods, m.NEW_LABEL_FONT)
                style_row(ws, row, 3, periods, m.NEW_LABEL_FONT)
                member_rows.append(row)
                row += 1
            ws.cell(row, 2, value=bucket).font = m.NEW_TOTAL_FONT
            ws.cell(row, 3, value=f"TOTAL: {bucket}").font = m.NEW_TOTAL_FONT
            for bidx in range(len(periods)):
                s = block_start(3, bidx)
                for off in (0, 1, 3, 4, 6):
                    L = get_column_letter(s + off)
                    ref = ",".join(f"{L}{r}" for r in member_rows) if member_rows else "0"
                    cell = ws.cell(row, s + off, value=f"=SUM({ref})" if member_rows else 0)
                    cell.number_format = b.NUM_FORMAT
                    cell.font = m.NEW_TOTAL_FONT
            write_totals(ws, row, 3, periods, m.NEW_TOTAL_FONT)
            style_row(ws, row, 3, periods, m.NEW_TOTAL_FONT, TOTAL_FILL)
            bucket_total_rows[bucket] = row
            row += 1
        row += 1
    return bucket_total_rows


# ---------------------------------------------------------------------------
# Level 3: Main Group (links to Sub Group)
# ---------------------------------------------------------------------------

def link_row(ws, row, label_cells, periods, source, src_row, font, fill=None):
    for cidx, val in label_cells:
        c = ws.cell(row, cidx, value=val)
        c.font = font
    for bidx in range(len(periods)):
        s_dest = block_start(2, bidx)
        s_src = block_start(3, bidx)
        for i, off in enumerate((0, 1, 3, 4, 6)):
            src_letter = get_column_letter(s_src + off)
            cell = ws.cell(row, s_dest + off, value=f"='{source}'!{src_letter}{src_row}")
            cell.number_format = b.NUM_FORMAT
            cell.font = font
    write_totals(ws, row, 2, periods, font)
    style_row(ws, row, 2, periods, font, fill)


def formula_row(ws, row, label, periods, parts, font, fill=None):
    """parts is a list of (source_row, "+"|"-") - see the SIGN CONVENTION note above for why the
    operator is explicit rather than a bare SUM()."""
    ws.cell(row, 1, value=label).font = font
    for bidx in range(len(periods)):
        s = block_start(2, bidx)
        for off in (0, 1, 3, 4, 6):
            L = get_column_letter(s + off)
            if parts:
                expr = "".join(f"{op}{L}{r}" for r, op in parts)
                cell = ws.cell(row, s + off, value=f"={expr.lstrip('+')}"
                               if parts[0][1] == "+" else f"={expr}")
            else:
                cell = ws.cell(row, s + off, value=0)
            cell.number_format = b.NUM_FORMAT
            cell.font = font
    write_totals(ws, row, 2, periods, font)
    style_row(ws, row, 2, periods, font, fill)


def write_pl_main(ws, periods, bucket_rows, rp_bucket_rows, title, warning):
    write_headers(ws, 2, ["Section", "Grouping"], periods, title,
                  "Every figure is a live link to 'PL Sub Group MOM' - a change there flows "
                  "straight up into this sheet.", warning)
    row = FIRST_DATA_ROW
    rows_of = {}
    for _tree_key, display, buckets in PL_SECTIONS:
        row = section_header(ws, row, 2, periods, display, memo="memo" in display.lower())
        for bucket in buckets:
            if bucket not in bucket_rows:
                continue
            link_row(ws, row, [(2, bucket)], periods, "PL Sub Group MOM", bucket_rows[bucket],
                     m.NEW_LABEL_FONT)
            rows_of[bucket] = row
            row += 1
        row += 1
    if rp_bucket_rows:
        row = section_header(ws, row, 2, periods, "RELATED PARTY TRANSACTIONS (memo)", memo=True)
        for bucket in RP_BUCKETS:
            if bucket not in rp_bucket_rows:
                continue
            link_row(ws, row, [(2, bucket)], periods, "PL Sub Group MOM", rp_bucket_rows[bucket],
                     m.NEW_LABEL_FONT)
            row += 1
        row += 1

    ebitda_parts = ([(rows_of[bk], "+") for bk in EBITDA_ADD if bk in rows_of]
                    + [(rows_of[bk], "-") for bk in EBITDA_LESS if bk in rows_of])
    formula_row(ws, row, "EBITDA", periods, ebitda_parts, m.NEW_MILESTONE_FONT, TOTAL_FILL)
    r_ebitda = row
    row += 1
    fin_parts = [(rows_of[bk], "+") for bk in ["Finance Cost"] if bk in rows_of]
    formula_row(ws, row, "TOTAL FINANCE COST", periods, fin_parts, m.NEW_TOTAL_FONT)
    r_fin = row
    row += 1
    formula_row(ws, row, "LESS: DEPRECIATION (not booked this period)", periods, [],
                m.NEW_TOTAL_FONT)
    r_dep = row
    row += 1
    formula_row(ws, row, "PBT (PROFIT BEFORE TAX)", periods,
                [(r_ebitda, "+"), (r_fin, "-"), (r_dep, "-")], m.NEW_MILESTONE_FONT, TOTAL_FILL)
    r_pbt = row
    row += 1
    formula_row(ws, row, "LESS: TAX (not booked this period)", periods, [], m.NEW_TOTAL_FONT)
    r_tax = row
    row += 1
    formula_row(ws, row, "PAT (NET PROFIT)", periods, [(r_pbt, "+"), (r_tax, "-")],
                m.NEW_MILESTONE_FONT, TOTAL_FILL)
    return {"PAT": row, "EBITDA": r_ebitda, "PBT": r_pbt}


def write_bs_main(ws, periods, bucket_rows, title, warning):
    write_headers(ws, 2, ["Section", "Grouping"], periods, title,
                  "Every figure is a live link to 'BS Sub Group MOM'. Balance Sheet is a position: "
                  "use each month's own column, never a sum across months.", warning)
    row = FIRST_DATA_ROW
    liab_rows, asset_rows = [], []
    row = section_header(ws, row, 2, periods, "LIABILITIES (Sources of Funds)")
    for bucket in b.BS_LIABILITY_ORDER:
        if bucket not in bucket_rows:
            continue
        label = b.BS_GROUP_LABEL.get(bucket, bucket)
        link_row(ws, row, [(2, label)], periods, "BS Sub Group MOM", bucket_rows[bucket],
                 m.NEW_LABEL_FONT)
        liab_rows.append(row)
        row += 1
    row += 1
    row = section_header(ws, row, 2, periods, "ASSETS (Application of Funds)")
    for bucket in b.BS_ASSET_ORDER:
        if bucket not in bucket_rows:
            continue
        label = b.BS_GROUP_LABEL.get(bucket, bucket)
        link_row(ws, row, [(2, label)], periods, "BS Sub Group MOM", bucket_rows[bucket],
                 m.NEW_LABEL_FONT)
        asset_rows.append(row)
        row += 1
    row += 1
    formula_row(ws, row, "TOTAL LIABILITIES", periods, [(r, "+") for r in liab_rows],
                m.NEW_MILESTONE_FONT, TOTAL_FILL)
    r_liab = row
    row += 1
    formula_row(ws, row, "TOTAL ASSETS", periods, [(r, "+") for r in asset_rows],
                m.NEW_MILESTONE_FONT, TOTAL_FILL)
    r_asset = row
    row += 1
    # Both totals are display magnitudes (see SIGN CONVENTION), so a balanced sheet is
    # Liabilities MINUS Assets = 0 - not a sum. Same wording as the pipeline's own gap row.
    formula_row(ws, row, "RECONCILIATION GAP (Liabilities - Assets; ~0 expected)", periods,
                [(r_liab, "+"), (r_asset, "-")], m.NEW_MILESTONE_FONT, GAP_FILL)
    return {"gap": row, "liab": r_liab, "asset": r_asset}


# ---------------------------------------------------------------------------
# Cross Check
# ---------------------------------------------------------------------------

def write_cross_check(ws, pl_periods, bs_periods, pl_main, bs_main, notes):
    ws.cell(1, 1, value="Cross Check").font = TITLE_FONT
    ws.cell(2, 1, value=(
        "Independent checks on this workbook. Every figure below is a live formula reading the "
        "other sheets - nothing is a pasted result, so these stay true on any recalculation."
    )).font = NOTE_FONT
    ws.column_dimensions["A"].width = 52
    for i in range(2, 12):
        ws.column_dimensions[get_column_letter(i)].width = 18
    row = 4

    ws.cell(row, 1, value="1. Balance Sheet must balance (Reconciliation Gap ~ 0)").font = Font(
        bold=True, size=10, name="Calibri", color="1F3864")
    row += 1
    for i, (_k, label) in enumerate(bs_periods):
        ws.cell(row, 1 + i + 1, value=label).font = m.NEW_TOTAL_FONT
    ws.cell(row, 1, value="Grand Total gap, per period").font = m.NEW_LABEL_FONT
    row += 1
    for i, _ in enumerate(bs_periods):
        L = col(2, i, 7)
        c = ws.cell(row, 2 + i, value=f"='BS Main Group MOM'!{L}{bs_main['gap']}")
        c.number_format = b.NUM_FORMAT
        c.font = m.NEW_TOTAL_FONT
    ws.cell(row, 1, value="Total Liabilities + Total Assets").font = m.NEW_LABEL_FONT
    row += 2

    ws.cell(row, 1, value="2. P&L: Apr + May + Jun must equal the Q1 Total column").font = Font(
        bold=True, size=10, name="Calibri", color="1F3864")
    row += 1
    keys = [k for k, _ in pl_periods]
    if all(k in keys for k in ("apr", "may", "jun", "qtr")):
        ia, im, ij, iq = (keys.index(k) for k in ("apr", "may", "jun", "qtr"))
        ws.cell(row, 1, value="PAT: Sum(Apr:Jun) vs Q1 Total  (variance)").font = m.NEW_LABEL_FONT
        pa, pm, pj, pq = (col(2, i, 7) for i in (ia, im, ij, iq))
        r = pl_main["PAT"]
        c = ws.cell(row, 2, value=f"=('PL Main Group MOM'!{pa}{r}+'PL Main Group MOM'!{pm}{r}"
                                  f"+'PL Main Group MOM'!{pj}{r})-'PL Main Group MOM'!{pq}{r}")
        c.number_format = b.NUM_FORMAT
        c.font = m.NEW_TOTAL_FONT
        row += 1
    if all(k in keys for k in ("apr", "may", "jun", "jul", "ytd")):
        idxs = [keys.index(k) for k in ("apr", "may", "jun", "jul")]
        iy = keys.index("ytd")
        ws.cell(row, 1, value="PAT: Sum(Apr:Jul) vs Apr-Jul Total  (variance)").font = m.NEW_LABEL_FONT
        r = pl_main["PAT"]
        terms = "+".join(f"'PL Main Group MOM'!{col(2, i, 7)}{r}" for i in idxs)
        c = ws.cell(row, 2, value=f"=({terms})-'PL Main Group MOM'!{col(2, iy, 7)}{r}")
        c.number_format = b.NUM_FORMAT
        c.font = m.NEW_TOTAL_FONT
        row += 1
    row += 1

    ws.cell(row, 1, value="3. Notes recorded at build time").font = Font(
        bold=True, size=10, name="Calibri", color="1F3864")
    row += 1
    for note in notes:
        c = ws.cell(row, 1, value=note)
        c.font = WARN_FONT if note.startswith("OPEN") else NOTE_FONT
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = 30
        row += 1
    ws.freeze_panes = "B4"


INDEX_DESC = {
    "PL Main Group MOM": "Profit & Loss - one row per main grouping, linked from Sub Group. Includes EBITDA / PBT / PAT / Branch Total / Related Party Total.",
    "PL Sub Group MOM": "Profit & Loss - one row per sub-group, each a formula linked to Full GL.",
    "PL Full GL MOM": "Profit & Loss - full ledger-level detail. This is the base level; Sub Group and Main Group link UP from here.",
    "BS Main Group MOM": "Balance Sheet - one row per main grouping, linked from Sub Group. Includes Total Liabilities / Total Assets / Reconciliation Gap.",
    "BS Sub Group MOM": "Balance Sheet - one row per sub-group, each a formula linked to Full GL.",
    "BS Full GL MOM": "Balance Sheet - full ledger-level detail. This is the base level; Sub Group and Main Group link UP from here.",
    "Cross Check": "Live self-checks: Balance Sheet gap per period, and Apr+May+Jun vs Q1 / Apr..Jul vs Apr-Jul Total.",
}


def build_index(wb, order, built_at, source_note):
    ws = wb.create_sheet(title="Index")
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 118
    ws.cell(1, 1, value="INDEX").font = TITLE_FONT
    ws.cell(2, 1, value=f"Built {built_at} by build_mom_pl_bs_26_27.py. {source_note}").font = NOTE_FONT
    row = 4
    for name in order:
        if name not in wb.sheetnames:
            continue
        c = ws.cell(row, 1, value=name)
        c.hyperlink = f"#'{name}'!A1"
        c.font = Font(size=10, name="Calibri", color="0563C1", underline="single")
        d = ws.cell(row, 2, value=INDEX_DESC.get(name, ""))
        d.font = Font(size=9, name="Calibri", color="595959")
        d.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1
    return ws


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--output", default=None, help=f"defaults to Files/{OUTPUT_NAME}")
    ap.add_argument("--no-archive", action="store_true",
                    help="do not move the previous version into 'Old Version of PL BS'")
    args = ap.parse_args()

    out = Path(args.output) if args.output else FILES_DIR / OUTPUT_NAME
    built_at = datetime.now().strftime("%d-%b-%Y %H:%M")

    pl_tree, rp_tree = m.build_multi_period_tree("pl", m.PL_PERIODS)
    bs_tree, _ = m.build_multi_period_tree("bs", m.BS_PERIODS)

    pl_periods = list(m.PL_PERIODS)
    bs_periods = list(m.BS_PERIODS)

    # Related Party is a memo view of rows that also live inside Turnover / Cost of Sales, so it
    # is built as its own mini-section rather than being spliced into PL_SECTIONS (which would
    # double-count it into EBITDA).
    rp_sections = [("RP", "RELATED PARTY TRANSACTIONS (memo)", RP_BUCKETS)]

    if out.exists() and not args.no_archive:
        ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.fromtimestamp(out.stat().st_mtime).strftime("%Y-%m-%d_%H%M")
        shutil.move(str(out), str(ARCHIVE_DIR / f"MOM PL and BS 26-27_{stamp}.xlsx"))
        print(f"Archived previous version -> Old Version of PL BS/MOM PL and BS 26-27_{stamp}.xlsx")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    warn_pl = None
    ws = wb.create_sheet(title="PL Full GL MOM")
    pl_sub_rows = write_full_gl(ws, pl_tree, PL_SECTIONS, pl_periods,
                                "PL FULL GL - MONTH ON MONTH (Apr-26 to Jul-26)", warn_pl)
    rp_sub_rows = write_full_gl_append(ws, rp_tree, rp_sections, pl_periods)

    ws = wb.create_sheet(title="PL Sub Group MOM")
    pl_bucket_rows = write_sub_group(ws, PL_SECTIONS, pl_periods, pl_sub_rows, "PL Full GL MOM",
                                     "PL SUB GROUP - MONTH ON MONTH (Apr-26 to Jul-26)", warn_pl)
    rp_bucket_rows = write_sub_group_append(ws, rp_sections, pl_periods, rp_sub_rows,
                                            "PL Full GL MOM")

    ws = wb.create_sheet(title="PL Main Group MOM")
    pl_main = write_pl_main(ws, pl_periods, pl_bucket_rows, rp_bucket_rows,
                            "PL MAIN - MONTH ON MONTH (Apr-26 to Jul-26)", warn_pl)

    bs_sections = [("LIABILITIES", "LIABILITIES (Sources of Funds)", b.BS_LIABILITY_ORDER),
                   ("ASSETS", "ASSETS (Application of Funds)", b.BS_ASSET_ORDER)]
    ws = wb.create_sheet(title="BS Full GL MOM")
    bs_sub_rows = write_full_gl(ws, bs_tree, bs_sections, bs_periods,
                                "BS FULL GL - MONTH ON MONTH (as at each month end)", None)
    ws = wb.create_sheet(title="BS Sub Group MOM")
    bs_bucket_rows = write_sub_group(ws, bs_sections, bs_periods, bs_sub_rows, "BS Full GL MOM",
                                     "BS SUB GROUP - MONTH ON MONTH (as at each month end)", None)
    ws = wb.create_sheet(title="BS Main Group MOM")
    bs_main = write_bs_main(ws, bs_periods, bs_bucket_rows,
                            "BS MAIN - MONTH ON MONTH (as at each month end)", None)

    notes = build_notes()
    ws = wb.create_sheet(title="Cross Check")
    write_cross_check(ws, pl_periods, bs_periods, pl_main, bs_main, notes)

    order = ["PL Main Group MOM", "PL Sub Group MOM", "PL Full GL MOM",
             "BS Main Group MOM", "BS Sub Group MOM", "BS Full GL MOM", "Cross Check"]
    build_index(wb, order, built_at,
                "Source: the gl_data/ + gl_data_mom_detail/ CSVs written by fetch_gl_data.py "
                "straight from Tally - the same inputs PL_BS_GL_Detail_<ts>.xlsx uses.")
    wb._sheets = [wb["Index"]] + [wb[n] for n in order if n in wb.sheetnames]

    try:
        wb.save(out)
    except PermissionError:
        raise RuntimeError(f"Cannot save to {out!r} - file may be open in Excel. Close it and retry.")
    print(f"File saved: {out}")


def write_full_gl_append(ws, tree, sections, periods):
    """Appends extra sections (Related Party memo) below whatever is already on the sheet."""
    row = ws.max_row + 2
    sub_total_rows = {}
    for tree_key, display, buckets in sections:
        row = section_header(ws, row, 4, periods, display, memo=True)
        for bucket in buckets:
            leaves = tree.get(bucket, {})
            by_sub = {}
            for (sub_grp, gl_name), pdata in leaves.items():
                by_sub.setdefault(sub_grp, []).append((gl_name, pdata))
            for sub_grp in sorted(by_sub):
                first = row
                for gl_name, pdata in sorted(by_sub[sub_grp]):
                    ws.cell(row, 2, value=bucket).font = m.NEW_LABEL_FONT
                    ws.cell(row, 3, value=sub_grp).font = m.NEW_LABEL_FONT
                    ws.cell(row, 4, value=gl_name).font = m.NEW_LABEL_FONT
                    for bidx, (pkey, _) in enumerate(periods):
                        vals = pdata.get(pkey, [0.0] * 5)
                        s = block_start(4, bidx)
                        for i, off in enumerate((0, 1, 3, 4, 6)):
                            cell = ws.cell(row, s + off, value=round(vals[i], 2))
                            cell.number_format = b.NUM_FORMAT
                            cell.font = m.NEW_LABEL_FONT
                    write_totals(ws, row, 4, periods, m.NEW_LABEL_FONT)
                    style_row(ws, row, 4, periods, m.NEW_LABEL_FONT)
                    row += 1
                ws.cell(row, 2, value=bucket).font = m.NEW_TOTAL_FONT
                ws.cell(row, 3, value=f"TOTAL: {sub_grp}").font = m.NEW_TOTAL_FONT
                for bidx in range(len(periods)):
                    s = block_start(4, bidx)
                    for off in (0, 1, 3, 4, 6):
                        L = get_column_letter(s + off)
                        cell = ws.cell(row, s + off,
                                       value=(f"=SUM({L}{first}:{L}{row - 1})" if row > first else 0))
                        cell.number_format = b.NUM_FORMAT
                        cell.font = m.NEW_TOTAL_FONT
                write_totals(ws, row, 4, periods, m.NEW_TOTAL_FONT)
                style_row(ws, row, 4, periods, m.NEW_TOTAL_FONT, TOTAL_FILL)
                sub_total_rows[(bucket, sub_grp)] = row
                row += 1
        row += 1
    return sub_total_rows


def write_sub_group_append(ws, sections, periods, sub_total_rows, source):
    row = ws.max_row + 2
    bucket_total_rows = {}
    for _tree_key, display, buckets in sections:
        row = section_header(ws, row, 3, periods, display, memo=True)
        for bucket in buckets:
            subs = sorted(s for (bk, s) in sub_total_rows if bk == bucket)
            member_rows = []
            for sub_grp in subs:
                src_row = sub_total_rows[(bucket, sub_grp)]
                ws.cell(row, 2, value=bucket).font = m.NEW_LABEL_FONT
                ws.cell(row, 3, value=sub_grp).font = m.NEW_LABEL_FONT
                for bidx in range(len(periods)):
                    s_dest = block_start(3, bidx)
                    s_src = block_start(4, bidx)
                    for i, off in enumerate((0, 1, 3, 4, 6)):
                        src_letter = get_column_letter(s_src + off)
                        cell = ws.cell(row, s_dest + off,
                                       value=f"='{source}'!{src_letter}{src_row}")
                        cell.number_format = b.NUM_FORMAT
                        cell.font = m.NEW_LABEL_FONT
                write_totals(ws, row, 3, periods, m.NEW_LABEL_FONT)
                style_row(ws, row, 3, periods, m.NEW_LABEL_FONT)
                member_rows.append(row)
                row += 1
            ws.cell(row, 2, value=bucket).font = m.NEW_TOTAL_FONT
            ws.cell(row, 3, value=f"TOTAL: {bucket}").font = m.NEW_TOTAL_FONT
            for bidx in range(len(periods)):
                s = block_start(3, bidx)
                for off in (0, 1, 3, 4, 6):
                    L = get_column_letter(s + off)
                    ref = ",".join(f"{L}{r}" for r in member_rows) if member_rows else "0"
                    cell = ws.cell(row, s + off, value=f"=SUM({ref})" if member_rows else 0)
                    cell.number_format = b.NUM_FORMAT
                    cell.font = m.NEW_TOTAL_FONT
            write_totals(ws, row, 3, periods, m.NEW_TOTAL_FONT)
            style_row(ws, row, 3, periods, m.NEW_TOTAL_FONT, TOTAL_FILL)
            bucket_total_rows[bucket] = row
            row += 1
        row += 1
    return bucket_total_rows


def build_notes():
    """Open items worth carrying on the face of the workbook rather than only in a chat message."""
    notes = []
    fx = b.UNADJUSTED_FOREX_GAIN_LOSS_ASSET.get("ORANGE_O_TEC_PRIVATE_LIMITED", {})
    if fx:
        notes.append(
            "OPEN ITEM - Orange O Tec Pvt Ltd carries an off-ledger 'Unadjusted Forex Gain/Loss "
            f"(Asset)' constant (Jun {fx.get('qtr', 0):,.0f} / Jul {fx.get('jul', 0):,.0f}). Since "
            "fetch_gl_data.py switched to Tally's Trial Balance (which balances by construction), "
            "adding it leaves a residual Reconciliation Gap for those two months. Confirm the item "
            "on Tally's Balance Sheet > Assets screen and either re-base or remove the constant.")
    notes.append(
        "Stock is not a ledger and comes from STOCK / STOCK_MOM in the build scripts, not from any "
        "CSV. It nets out of the Balance Sheet gap by construction, so a zero gap does NOT confirm "
        "it - see the chat/report accompanying this refresh for the live-vs-script comparison.")
    return notes


if __name__ == "__main__":
    main()
