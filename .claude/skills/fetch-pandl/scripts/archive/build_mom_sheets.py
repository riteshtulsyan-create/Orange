"""
Adds two MOM (month-on-month) sheets to Files/PL_BS_GL_Detail.xlsx:
  - "MOM PL and BS": consolidated (all 5 companies combined) P&L and BS, trended
    Apr-26 / May-26 / Jun-26 / Q1 Total, same statement structure as the main sheets
    (Turnover -> Cost of Sales -> GP -> EBITDA -> PBT -> PAT / Liabilities -> Assets).
  - "MOM PL and BS - By Company": same statement structure, but each month (and Q1
    Total) broken out by company + family totals + grand total (35 value columns).

Row detail is at Grouping + Sub Grouping level (no individual GL Name) - a monthly
trend at full ledger granularity would be enormous and not useful for MIS review.

Source data: CSVs pre-fetched via tally-prime MCP tools, saved to gl_data_mom/ next to
this script's scratch location, aggregated to (primary_group, group_name, closing_balance)
per company per month. Balance sheet month-end snapshots use fromDate=2026-04-01 (start
of the quarter/period) so the embedded 'Profit & Loss A/c (current)' line is cumulative
from period start, consistent with the main GL Detail sheets.
"""

import csv
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

from build_gl_detail_sidebyside import (
    ALL_COMPANIES, CO1, CO2, CO3, CO4, CO5,
    classify_pl_row, classify_bs_row,
    PL_TURNOVER_BUCKETS, PL_COST_OF_SALES_BUCKETS, PL_OTHER_INCOME_BUCKETS,
    PL_OPERATING_EXPENSE_BUCKETS, PL_FINANCE_COST_BUCKETS,
    BS_LIABILITY_ORDER, BS_ASSET_ORDER, BS_GROUP_LABEL,
    TITLE_FONT, HEADER_FONT, HEADER_FILL, TOTALCOL_FILL, GROUP_FONT, SUBGROUP_FONT,
    NORMAL_FONT, SUBTOTAL_SUBGROUP_FONT, SUBTOTAL_SUBGROUP_FILL, SUBTOTAL_FONT, SUBTOTAL_FILL,
    SECTION_FONT, SECTION_FILL, SECTION_TOTAL_FONT, SECTION_TOTAL_FILL,
    GP_FONT, GP_FILL, EBITDA_FONT, EBITDA_FILL, PBT_FONT, PBT_FILL, PAT_FONT, PAT_FILL,
    ZERO_FONT, ZERO_FILL, GAP_FONT, GAP_FILL, NUM_FORMAT,
)

SCRIPT_DIR = Path(__file__).parent
MOM_DATA_DIR = Path(
    r"C:\Users\orang\AppData\Local\Temp\claude\c--Users-orang-OneDrive---Orange-O-Tec-Private-Limited-00-Claude-Project-PL-and-BS"
    r"\1692684d-78ab-405d-8ee7-bf023df8abe6\scratchpad\gl_data_mom"
)
QUARTER_DATA_DIR = Path(
    r"C:\Users\orang\AppData\Local\Temp\claude\c--Users-orang-OneDrive---Orange-O-Tec-Private-Limited-00-Claude-Project-PL-and-BS"
    r"\1692684d-78ab-405d-8ee7-bf023df8abe6\scratchpad\gl_data"
)

MONTHS = [("Apr-26", "apr"), ("May-26", "may"), ("Jun-26", "jun")]

STOCK_MOM = {
    "ORANGE_O_TEC_PRIVATE_LIMITED": {"opening": 292508691.08, "apr": 279429066.97, "may": 265672654.09, "jun": 279830678.42},
    "PRIVATE_LIMITED_NOIDA": {"opening": 32685329.56, "apr": 33592810.11, "may": 28224526.73, "jun": 29258524.91},
    "ENTERPRISES_PVT_LTD": {"opening": 71324050.22, "apr": 69800418.24, "may": 86519323.49, "jun": 91179012.34},
    "ENTERPRISES_NOIDA": {"opening": 32959492.39, "apr": 34367665.06, "may": 35201591.7, "jun": 37902468.1},
    "COLORIX": {"opening": 105648, "apr": 294943.76, "may": 294943.76, "jun": 105648},
}
MONTH_OPEN_KEY = {"apr": "opening", "may": "apr", "jun": "may"}


def load_agg_csv(path):
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append((r["primary_group"], r["group_name"], float(r["closing_balance"])))
    return rows


def aggregate_gl_csv(path):
    sums = {}
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            k = (r["primary_group"], r["group_name"])
            sums[k] = sums.get(k, 0.0) + float(r["closing_balance"])
    return [(pg, gn, bal) for (pg, gn), bal in sums.items()]


def stock_rows_for_month(stat_type, company_key, month_key):
    s = STOCK_MOM[company_key]
    open_val = s[MONTH_OPEN_KEY[month_key]]
    close_val = s[month_key]
    if stat_type == "pl":
        return [("Stock Valuation (Inventory)", "Stock Valuation (Inventory)", -open_val),
                ("Stock Valuation (Inventory)", "Stock Valuation (Inventory)", close_val)]
    return [("Current Assets", "Closing Stock (Inventory)", -close_val)]


def load_month_rows(stat_type, company_key, month_key):
    if stat_type == "pl":
        rows = load_agg_csv(MOM_DATA_DIR / f"pl_{month_key}_{company_key}.csv")
    else:
        if month_key == "jun":
            rows = aggregate_gl_csv(QUARTER_DATA_DIR / f"bs_{company_key}.csv")
        else:
            rows = load_agg_csv(MOM_DATA_DIR / f"bs_{month_key}_{company_key}.csv")
    rows += stock_rows_for_month(stat_type, company_key, month_key)
    return rows


def classify(stat_type, pg, sub_grp):
    if stat_type == "pl":
        return classify_pl_row(pg, sub_grp)
    return classify_bs_row(pg)


def label_for(stat_type, bucket):
    if stat_type == "bs":
        return BS_GROUP_LABEL.get(bucket, bucket) if bucket != "" else BS_GROUP_LABEL.get("", bucket)
    return bucket


def bucket_orders(stat_type):
    if stat_type == "pl":
        return {
            "TURNOVER": PL_TURNOVER_BUCKETS,
            "COST OF SALES": PL_COST_OF_SALES_BUCKETS,
            "OTHER INCOME": PL_OTHER_INCOME_BUCKETS,
            "OPERATING EXPENSES": PL_OPERATING_EXPENSE_BUCKETS,
            "FINANCE COST": PL_FINANCE_COST_BUCKETS,
        }
    return {"LIABILITIES": BS_LIABILITY_ORDER, "ASSETS": BS_ASSET_ORDER}


# ---------------------------------------------------------------------------
# Data tree: tree[section][bucket][sub_grp][company_key] = [apr, may, jun]
# ---------------------------------------------------------------------------

def build_tree(stat_type):
    tree = {}
    for month_idx, (_, month_key) in enumerate(MONTHS):
        for _, key in ALL_COMPANIES:
            for pg, sub_grp, bal in load_month_rows(stat_type, key, month_key):
                section, bucket = classify(stat_type, pg, sub_grp)
                slot = (tree.setdefault(section, {}).setdefault(bucket, {})
                        .setdefault(sub_grp, {}).setdefault(key, [0.0, 0.0, 0.0]))
                slot[month_idx] += bal
    return tree


def qtr(vals3):
    return vals3[0] + vals3[1] + vals3[2]


def sum_company_lists(lists):
    out = [0.0, 0.0, 0.0]
    for lst in lists:
        for i in range(3):
            out[i] += lst[i]
    return out


# ---------------------------------------------------------------------------
# Sheet 1: consolidated (all companies combined), Apr/May/Jun/Qtr columns
# ---------------------------------------------------------------------------

C_SECTION, C_GROUPING, C_SUBGROUPING, C_APR, C_MAY, C_JUN, C_QTR = 1, 2, 3, 4, 5, 6, 7


def write_consolidated_values(ws, row, vals3, font, fill=None):
    apr, may, jun = vals3
    for col, val in ((C_APR, apr), (C_MAY, may), (C_JUN, jun), (C_QTR, apr + may + jun)):
        cell = ws.cell(row, col, value=val)
        cell.number_format = NUM_FORMAT
        cell.font = font
    if fill:
        for c in range(1, C_QTR + 1):
            ws.cell(row, c).fill = fill


def consolidated_company_totals(node):
    """node: {sub_grp: {company_key: [apr,may,jun]}} -> sub_grp: [apr,may,jun] summed across companies."""
    out = {}
    for sub_grp, by_company in node.items():
        out[sub_grp] = sum_company_lists(by_company.values())
    return out


def write_consolidated_bucket(ws, row, display_label, subgroup_vals, base_level):
    if not subgroup_vals:
        return row, [0.0, 0.0, 0.0]
    items = sorted(subgroup_vals.items(), key=lambda kv: abs(qtr(kv[1])), reverse=True)
    bucket_totals = [0.0, 0.0, 0.0]
    for sub_grp, vals in items:
        ws.cell(row, C_SUBGROUPING, value=sub_grp).font = SUBGROUP_FONT
        write_consolidated_values(ws, row, vals, NORMAL_FONT)
        ws.row_dimensions[row].outlineLevel = base_level + 1
        row += 1
        for i in range(3):
            bucket_totals[i] += vals[i]
    ws.cell(row, C_GROUPING, value=display_label).font = GROUP_FONT
    ws.cell(row, C_SUBGROUPING, value=f"TOTAL: {display_label}").font = SUBTOTAL_FONT
    write_consolidated_values(ws, row, bucket_totals, SUBTOTAL_FONT, fill=SUBTOTAL_FILL)
    ws.row_dimensions[row].outlineLevel = base_level
    row += 1
    return row, bucket_totals


def write_consolidated_section(ws, row, section_label, tree_section, order_map, total_label=None, base_level=1):
    total_label = total_label or f"TOTAL {section_label.split(':')[-1].strip()}"
    ws.cell(row, C_SECTION, value=section_label).font = SECTION_FONT
    for c in range(1, C_QTR + 1):
        ws.cell(row, c).fill = SECTION_FILL
    ws.row_dimensions[row].outlineLevel = 0
    row += 1

    section_totals = [0.0, 0.0, 0.0]
    for bucket in order_map:
        if bucket not in tree_section:
            continue
        display_label = label_for("bs" if bucket in BS_LIABILITY_ORDER + BS_ASSET_ORDER else "pl", bucket)
        subgroup_vals = consolidated_company_totals(tree_section[bucket])
        row, bucket_totals = write_consolidated_bucket(ws, row, display_label, subgroup_vals, base_level)
        for i in range(3):
            section_totals[i] += bucket_totals[i]

    ws.cell(row, C_SUBGROUPING, value=total_label).font = SECTION_TOTAL_FONT
    write_consolidated_values(ws, row, section_totals, SECTION_TOTAL_FONT, fill=SECTION_TOTAL_FILL)
    ws.row_dimensions[row].outlineLevel = 0
    row += 1
    return row, section_totals


def write_consolidated_computed(ws, row, label, vals3, font, fill):
    ws.cell(row, C_SECTION, value=label).font = font
    write_consolidated_values(ws, row, vals3, font, fill=fill)
    ws.row_dimensions[row].outlineLevel = 0
    return row + 1


def write_consolidated_header(ws, row, title):
    ws.cell(row, 1, value=title).font = TITLE_FONT
    row += 2
    ws.cell(row, C_SECTION, value="Statement Section")
    ws.cell(row, C_GROUPING, value="Grouping")
    ws.cell(row, C_SUBGROUPING, value="Sub Grouping")
    ws.cell(row, C_APR, value=MONTHS[0][0])
    ws.cell(row, C_MAY, value=MONTHS[1][0])
    ws.cell(row, C_JUN, value=MONTHS[2][0])
    ws.cell(row, C_QTR, value="Q1 Total")
    for c in range(1, C_QTR + 1):
        cell = ws.cell(row, c)
        cell.font = HEADER_FONT
        cell.fill = TOTALCOL_FILL if c == C_QTR else HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="bottom")
    ws.freeze_panes = f"D{row + 1}"
    return row + 1


def write_consolidated_sheet(ws):
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 30
    for col in "DEFG":
        ws.column_dimensions[col].width = 17
    ws.sheet_properties.outlinePr.summaryBelow = True

    pl_tree = build_tree("pl")
    bs_tree = build_tree("bs")

    row = write_consolidated_header(ws, 1, "MOM - Profit & Loss A/c (all companies consolidated)")

    row, turnover = write_consolidated_section(ws, row, "TURNOVER", pl_tree.get("TURNOVER", {}),
                                                PL_TURNOVER_BUCKETS, total_label="TOTAL TURNOVER")
    row, cos = write_consolidated_section(ws, row, "LESS: COST OF SALES", pl_tree.get("COST OF SALES", {}),
                                           PL_COST_OF_SALES_BUCKETS, total_label="TOTAL COST OF SALES")
    gp = [turnover[i] + cos[i] for i in range(3)]
    row = write_consolidated_computed(ws, row, "GROSS PROFIT", gp, GP_FONT, GP_FILL)

    row, oi = write_consolidated_section(ws, row, "ADD: OTHER INCOME", pl_tree.get("OTHER INCOME", {}),
                                          PL_OTHER_INCOME_BUCKETS, total_label="TOTAL OTHER INCOME")
    row, opex = write_consolidated_section(ws, row, "LESS: OPERATING EXPENSES", pl_tree.get("OPERATING EXPENSES", {}),
                                            PL_OPERATING_EXPENSE_BUCKETS, total_label="TOTAL OPERATING EXPENSES")
    ebitda = [gp[i] + oi[i] + opex[i] for i in range(3)]
    row = write_consolidated_computed(ws, row, "EBITDA", ebitda, EBITDA_FONT, EBITDA_FILL)

    row, fc = write_consolidated_section(ws, row, "LESS: FINANCE COST", pl_tree.get("FINANCE COST", {}),
                                          PL_FINANCE_COST_BUCKETS, total_label="TOTAL FINANCE COST")
    dep = [0.0, 0.0, 0.0]
    row = write_consolidated_computed(ws, row, "LESS: DEPRECIATION (not booked this period)", dep, ZERO_FONT, ZERO_FILL)
    pbt = [ebitda[i] + fc[i] + dep[i] for i in range(3)]
    row = write_consolidated_computed(ws, row, "PBT (PROFIT BEFORE TAX)", pbt, PBT_FONT, PBT_FILL)
    tax = [0.0, 0.0, 0.0]
    row = write_consolidated_computed(ws, row, "LESS: TAX (not booked this period)", tax, ZERO_FONT, ZERO_FILL)
    pat = [pbt[i] + tax[i] for i in range(3)]
    row = write_consolidated_computed(ws, row, "PAT (NET PROFIT)", pat, PAT_FONT, PAT_FILL)

    row += 2
    row = write_consolidated_header(
        ws, row,
        "MOM - Balance Sheet (all companies consolidated)  "
        "[Note: BS is point-in-time - the 'Q1 Total' column sums 3 snapshots and is NOT a "
        "meaningful balance. Use the Jun-26 column for the quarter-end position.]")
    row, liab = write_consolidated_section(ws, row, "LIABILITIES (Sources of Funds)", bs_tree.get("LIABILITIES", {}),
                                            BS_LIABILITY_ORDER, total_label="TOTAL LIABILITIES")
    row, assets = write_consolidated_section(ws, row, "ASSETS (Application of Funds)", bs_tree.get("ASSETS", {}),
                                              BS_ASSET_ORDER, total_label="TOTAL ASSETS")
    gap = [liab[i] + assets[i] for i in range(3)]
    write_consolidated_computed(ws, row, "RECONCILIATION GAP (Liabilities + Assets; ~0 expected)", gap, GAP_FONT, GAP_FILL)


# ---------------------------------------------------------------------------
# Sheet 2: by company, 8 column-blocks (CO1,CO2,Orange,CO3,CO4,Enterprises,CO5,Grand)
#          each block has Apr/May/Jun/Qtr sub-columns
# ---------------------------------------------------------------------------

W_LABEL_COLS = 3  # Section, Grouping, Sub Grouping
BLOCKS = [
    ("CO1", CO1[0], [CO1[1]]),
    ("CO2", CO2[0], [CO2[1]]),
    ("ORANGE", "Total of Orange O Tec", [CO1[1], CO2[1]]),
    ("CO3", CO3[0], [CO3[1]]),
    ("CO4", CO4[0], [CO4[1]]),
    ("ENTERPRISES", "Total of Enterprises", [CO3[1], CO4[1]]),
    ("CO5", CO5[0], [CO5[1]]),
    ("GRAND", "Grand Total", [CO1[1], CO2[1], CO3[1], CO4[1], CO5[1]]),
]
TOTAL_BLOCK_KEYS = {"ORANGE", "ENTERPRISES", "GRAND"}


def block_start_col(block_idx):
    return W_LABEL_COLS + 1 + block_idx * 4


W_TOTAL_COLS = W_LABEL_COLS + len(BLOCKS) * 4


def write_wide_values(ws, row, by_company, font):
    for idx, (bkey, _, members) in enumerate(BLOCKS):
        start = block_start_col(idx)
        vals3 = sum_company_lists(by_company.get(m, [0.0, 0.0, 0.0]) for m in members)
        apr, may, jun = vals3
        for offset, val in enumerate((apr, may, jun, apr + may + jun)):
            cell = ws.cell(row, start + offset, value=val)
            cell.number_format = NUM_FORMAT
            cell.font = font


def shade_row(ws, row, fill, up_to=W_TOTAL_COLS):
    for c in range(1, up_to + 1):
        ws.cell(row, c).fill = fill


def write_wide_bucket(ws, row, display_label, node, base_level):
    """node: {sub_grp: {company_key: [apr,may,jun]}}"""
    if not node:
        return row, {}
    items = sorted(node.items(), key=lambda kv: abs(qtr(sum_company_lists(kv[1].values()))), reverse=True)
    bucket_totals = {}  # company_key -> [apr,may,jun]
    for sub_grp, by_company in items:
        ws.cell(row, C_SUBGROUPING, value=sub_grp).font = SUBGROUP_FONT
        write_wide_values(ws, row, by_company, NORMAL_FONT)
        ws.row_dimensions[row].outlineLevel = base_level + 1
        row += 1
        for ck, vals in by_company.items():
            slot = bucket_totals.setdefault(ck, [0.0, 0.0, 0.0])
            for i in range(3):
                slot[i] += vals[i]
    ws.cell(row, C_GROUPING, value=display_label).font = GROUP_FONT
    ws.cell(row, C_SUBGROUPING, value=f"TOTAL: {display_label}").font = SUBTOTAL_FONT
    write_wide_values(ws, row, bucket_totals, SUBTOTAL_FONT)
    shade_row(ws, row, SUBTOTAL_FILL)
    ws.row_dimensions[row].outlineLevel = base_level
    row += 1
    return row, bucket_totals


def write_wide_section(ws, row, section_label, tree_section, order_map, total_label=None, base_level=1):
    total_label = total_label or f"TOTAL {section_label}"
    ws.cell(row, C_SECTION, value=section_label).font = SECTION_FONT
    shade_row(ws, row, SECTION_FILL)
    ws.row_dimensions[row].outlineLevel = 0
    row += 1

    section_totals = {}
    for bucket in order_map:
        if bucket not in tree_section:
            continue
        display_label = label_for("bs" if bucket in BS_LIABILITY_ORDER + BS_ASSET_ORDER else "pl", bucket)
        row, bucket_totals = write_wide_bucket(ws, row, display_label, tree_section[bucket], base_level)
        for ck, vals in bucket_totals.items():
            slot = section_totals.setdefault(ck, [0.0, 0.0, 0.0])
            for i in range(3):
                slot[i] += vals[i]

    ws.cell(row, C_SUBGROUPING, value=total_label).font = SECTION_TOTAL_FONT
    write_wide_values(ws, row, section_totals, SECTION_TOTAL_FONT)
    shade_row(ws, row, SECTION_TOTAL_FILL)
    ws.row_dimensions[row].outlineLevel = 0
    row += 1
    return row, section_totals


def write_wide_computed(ws, row, label, by_company, font, fill):
    ws.cell(row, C_SECTION, value=label).font = font
    write_wide_values(ws, row, by_company, font)
    shade_row(ws, row, fill)
    ws.row_dimensions[row].outlineLevel = 0
    return row + 1


def combine(a, b, sign_b=1):
    out = dict(a)
    for ck, vals in b.items():
        slot = out.setdefault(ck, [0.0, 0.0, 0.0])
        out[ck] = [slot[i] + sign_b * vals[i] for i in range(3)]
    return out


def write_wide_header(ws, row, title):
    ws.cell(row, 1, value=title).font = TITLE_FONT
    row += 2
    header_row1 = row
    header_row2 = row + 1
    ws.cell(header_row1, C_SECTION, value="Statement Section")
    ws.cell(header_row1, C_GROUPING, value="Grouping")
    ws.cell(header_row1, C_SUBGROUPING, value="Sub Grouping")
    ws.merge_cells(start_row=header_row1, start_column=C_SECTION, end_row=header_row2, end_column=C_SECTION)
    ws.merge_cells(start_row=header_row1, start_column=C_GROUPING, end_row=header_row2, end_column=C_GROUPING)
    ws.merge_cells(start_row=header_row1, start_column=C_SUBGROUPING, end_row=header_row2, end_column=C_SUBGROUPING)
    for idx, (bkey, label, _) in enumerate(BLOCKS):
        start = block_start_col(idx)
        ws.merge_cells(start_row=header_row1, start_column=start, end_row=header_row1, end_column=start + 3)
        cell = ws.cell(header_row1, start, value=label)
        cell.font = HEADER_FONT
        cell.fill = TOTALCOL_FILL if bkey in TOTAL_BLOCK_KEYS else HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        for offset, sub_label in enumerate(("Apr-26", "May-26", "Jun-26", "Q1 Total")):
            c2 = ws.cell(header_row2, start + offset, value=sub_label)
            c2.font = HEADER_FONT
            c2.fill = TOTALCOL_FILL if bkey in TOTAL_BLOCK_KEYS else HEADER_FILL
            c2.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    for c in range(1, C_SUBGROUPING + 1):
        cell = ws.cell(header_row1, c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    ws.row_dimensions[header_row1].height = 20
    ws.row_dimensions[header_row2].height = 20
    ws.freeze_panes = ws.cell(header_row2 + 1, W_LABEL_COLS + 1).coordinate
    return header_row2 + 1


def write_by_company_sheet(ws):
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 30
    for idx in range(len(BLOCKS)):
        start = block_start_col(idx)
        for offset in range(4):
            ws.column_dimensions[get_column_letter(start + offset)].width = 15
    ws.sheet_properties.outlinePr.summaryBelow = True

    pl_tree = build_tree("pl")
    bs_tree = build_tree("bs")

    row = write_wide_header(ws, 1, "MOM - Profit & Loss A/c (by company)")

    row, turnover = write_wide_section(ws, row, "TURNOVER", pl_tree.get("TURNOVER", {}),
                                        PL_TURNOVER_BUCKETS, total_label="TOTAL TURNOVER")
    row, cos = write_wide_section(ws, row, "LESS: COST OF SALES", pl_tree.get("COST OF SALES", {}),
                                   PL_COST_OF_SALES_BUCKETS, total_label="TOTAL COST OF SALES")
    gp = combine(turnover, cos)
    row = write_wide_computed(ws, row, "GROSS PROFIT", gp, GP_FONT, GP_FILL)

    row, oi = write_wide_section(ws, row, "ADD: OTHER INCOME", pl_tree.get("OTHER INCOME", {}),
                                  PL_OTHER_INCOME_BUCKETS, total_label="TOTAL OTHER INCOME")
    row, opex = write_wide_section(ws, row, "LESS: OPERATING EXPENSES", pl_tree.get("OPERATING EXPENSES", {}),
                                    PL_OPERATING_EXPENSE_BUCKETS, total_label="TOTAL OPERATING EXPENSES")
    ebitda = combine(combine(gp, oi), opex)
    row = write_wide_computed(ws, row, "EBITDA", ebitda, EBITDA_FONT, EBITDA_FILL)

    row, fc = write_wide_section(ws, row, "LESS: FINANCE COST", pl_tree.get("FINANCE COST", {}),
                                  PL_FINANCE_COST_BUCKETS, total_label="TOTAL FINANCE COST")
    dep = {}
    row = write_wide_computed(ws, row, "LESS: DEPRECIATION (not booked this period)", dep, ZERO_FONT, ZERO_FILL)
    pbt = combine(combine(ebitda, fc), dep)
    row = write_wide_computed(ws, row, "PBT (PROFIT BEFORE TAX)", pbt, PBT_FONT, PBT_FILL)
    tax = {}
    row = write_wide_computed(ws, row, "LESS: TAX (not booked this period)", tax, ZERO_FONT, ZERO_FILL)
    pat = combine(pbt, tax)
    row = write_wide_computed(ws, row, "PAT (NET PROFIT)", pat, PAT_FONT, PAT_FILL)

    row += 2
    row = write_wide_header(
        ws, row,
        "MOM - Balance Sheet (by company)  "
        "[Note: BS is point-in-time - each block's 'Q1 Total' sums 3 snapshots and is NOT a "
        "meaningful balance. Use the Jun-26 column for the quarter-end position.]")
    row, liab = write_wide_section(ws, row, "LIABILITIES (Sources of Funds)", bs_tree.get("LIABILITIES", {}),
                                    BS_LIABILITY_ORDER, total_label="TOTAL LIABILITIES")
    row, assets = write_wide_section(ws, row, "ASSETS (Application of Funds)", bs_tree.get("ASSETS", {}),
                                      BS_ASSET_ORDER, total_label="TOTAL ASSETS")
    gap = combine(liab, assets)
    write_wide_computed(ws, row, "RECONCILIATION GAP (Liabilities + Assets; ~0 expected)", gap, GAP_FONT, GAP_FILL)


def main():
    output_path = Path(__file__).parent.parent.parent.parent.parent / "Files" / "PL_BS_GL_Detail.xlsx"
    wb = openpyxl.load_workbook(output_path)

    for name in ("MOM PL and BS", "MOM PL and BS - By Company"):
        if name in wb.sheetnames:
            del wb[name]

    ws1 = wb.create_sheet(title="MOM PL and BS")
    write_consolidated_sheet(ws1)

    ws2 = wb.create_sheet(title="MOM PL and BS - By Company")
    write_by_company_sheet(ws2)

    try:
        wb.save(output_path)
    except PermissionError:
        raise RuntimeError(f"Cannot save to {output_path!r} - file may be open in Excel.\nClose the file and retry.")

    print(f"File saved: {output_path}")


if __name__ == "__main__":
    main()
