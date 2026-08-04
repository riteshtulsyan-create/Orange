"""
Builds the MOM (month-on-month) sheets with the same Schedule + Summary architecture as the
main workbook:
  - "MOM PL GL Detail" / "MOM BS GL Detail": full ledger-level schedules, one row per unique
    (Sub Grouping, GL Name), with a separate column-block per period (Apr-26/May-26/Jun-26/
    Q1 Total for P&L; Apr-26/May-26/Jun-26 for BS - no Q1 Total for BS since it's point-in-time).
    Every subtotal/total/computed row is a live Excel formula (SUM/+/-), never a pre-computed
    number - exactly like PL GL Detail / BS GL Detail.
  - "MOM PL and BS" / "MOM PL and BS - By Company": formula-linked summaries referencing the
    schedules above, with the same Sub Group column, grouping headers, Branch/Related Party
    memo sections, Indirect Expense sub-group breakdown, and green/red PAT coloring as
    PL Summary / BS Summary.

Source data: April and May use freshly-fetched ledger-level CSVs in gl_data_mom_detail/
(fromDate/toDate scoped to that month for P&L; fromDate=2026-04-01 with toDate=month-end for
BS, so the embedded 'Profit & Loss A/c (current)' line is cumulative from period start). June
P&L uses gl_data_mom_detail/pl_jun_<key>.csv (June's own flow only). June BS and the Q1 Total
P&L column both reuse the main gl_data/ CSVs via build_gl_detail_sidebyside.load_company_rows,
so they automatically pick up the Stock and Sales Bills to Make adjustments already applied
there. Sales Bills to Make has no confirmed monthly split, so it is applied only in the Jun-26
column (both schedules), per the agreed approximation - not in Apr/May.
"""

import argparse
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))
import build_gl_detail_sidebyside as b

MOM_DETAIL_DIR = Path(__file__).parent.parent / "gl_data_mom_detail"

MONTH_OPEN_KEY = {"apr": "opening", "may": "apr", "jun": "may", "jul": "jun"}
STOCK_MOM = {
    "ORANGE_O_TEC_PRIVATE_LIMITED": {"opening": 292508691.08, "apr": 279429066.97, "may": 265672654.09, "jun": 277198530.32, "jul": 260716650.16},
    "PRIVATE_LIMITED_NOIDA": {"opening": 32685329.56, "apr": 33592810.11, "may": 28224526.73, "jun": 29258524.91, "jul": 26995737.57},
    "ENTERPRISES_PVT_LTD": {"opening": 71324050.22, "apr": 69800418.24, "may": 86519323.49, "jun": 91177937.29, "jul": 111934955.18},
    "ENTERPRISES_NOIDA": {"opening": 32959492.39, "apr": 34367665.06, "may": 35201591.7, "jun": 37910968.1, "jul": 37008890.83},
    "COLORIX": {"opening": 105648, "apr": 294943.76, "may": 294943.76, "jun": 105648, "jul": 105648},
}
# 2026-07-31: July closing stock values above are PROVISIONAL - fetched while July's Tally
# voucher entry was still visibly incomplete for 4 of 5 companies (thin P&L/BS row counts vs
# Apr/May/Jun - see MOM PL and BS 26-27.xlsx's July column note). Re-fetch and update once July
# is confirmed fully entered.

# Period columns. Q1 Total (Apr-Jun) is deliberately KEPT as its own column so every prior
# report stays comparable, with Jul-26 and a cumulative Apr-Jul Total appended after it rather
# than folded into it (user decision, 2026-08-03 - "keep Q1, add Jul as extra column and show
# the April to Jul total also").
#
# P&L columns are FLOWS, so both cumulative columns are real fetches, not derived sums:
# 'qtr' reads gl_data/pl_<KEY>.csv (01-Apr..30-Jun) and 'ytd' reads gl_data/pl_ytd_<KEY>.csv
# (01-Apr..31-Jul). Keeping 'ytd' as its own fetch rather than a Q1+Jul formula is what lets the
# Cross Check sheet verify Sum(Apr:Jul) against it as a genuine reconciliation instead of a
# tautology.
#
# BS columns are POSITIONS, so there is no cumulative column to add: the Jul-26 column IS the
# as-at-31-Jul position, which is exactly what an "Apr-Jul" balance sheet would be.
PL_PERIODS = [("apr", "Apr-26"), ("may", "May-26"), ("jun", "Jun-26"), ("qtr", "Q1 Total"),
              ("jul", "Jul-26"), ("ytd", "Apr-Jul Total")]
BS_PERIODS = [("apr", "Apr-26"), ("may", "May-26"), ("jun", "Jun-26"), ("jul", "Jul-26")]


def load_period_rows(stat_type, key, period):
    """Returns [(primary_group, sub_group, gl_name, closing_balance), ...] for one company/
    period, including the same Stock Valuation synthetic rows the main schedule uses (Q1 Total
    P&L and Jun BS reuse load_company_rows directly so they can never drift from the main GL
    Detail sheets). BS periods (apr/may/jun) also get the period-specific Sales Bills to Make /
    Purchase Bills to Come / Unadjusted Forex Gain/Loss values from build_gl_detail_sidebyside's
    per-period dicts, so each month's own Balance Sheet Reconciliation Gap reads ~0 - these three
    off-ledger items genuinely grow/shrink month to month, so a single quarter-end constant cannot
    be reused across April/May the way it can for June (which reuses the quarter file directly).
    P&L flow for Apr/May deliberately does NOT get these three items (no confirmed monthly P&L
    split exists for them - see the "Sales Bills to Make has no confirmed monthly split" note in
    this file's module docstring); only June's own P&L flow gets the quarter-end approximation, as
    before - this is a documented, deliberate limitation, unrelated to the BS-side fix here."""
    if stat_type == "pl" and period == "qtr":
        return b.load_company_rows("pl", key, "pl")
    if stat_type == "bs" and period == "jun":
        return b.load_company_rows("bs", key, "bs")
    if stat_type == "pl" and period == "ytd":
        # Cumulative 01-Apr..31-Jul P&L flow, fetched in its own right (see PL_PERIODS note).
        # Stock brackets the WHOLE period: opening is the 01-Apr opening, closing is 31-Jul -
        # not July's own opening, which would understate cost of sales for the four months.
        rows = b.load_csv(b.GL_DATA_DIR / f"pl_ytd_{key}.csv")
        rows.append(("Stock Valuation (Inventory)", b.STOCK_VALUATION_SUBGROUP,
                     "Opening Stock", -b.STOCK[key]["opening"]))
        rows.append(("Stock Valuation (Inventory)", b.STOCK_VALUATION_SUBGROUP,
                     "Closing Stock", STOCK_MOM[key]["jul"]))
        sbtm = b.SALES_BILLS_TO_MAKE.get(key, {}).get("ytd", 0.0)
        if sbtm:
            rows.append(("Sales Accounts", "Sales Bills to Make", "Sales Bills to Make", sbtm))
        pbtc = b.PURCHASE_BILLS_TO_COME.get(key, {}).get("ytd", 0.0)
        if pbtc:
            rows.append(("Purchase Accounts", "Purchase Bills to Come", "Purchase Bills to Come", pbtc))
        return rows
    if stat_type == "pl" and period == "jul":
        # July's own flow. The off-ledger adjustments are as-at BALANCES, not flows, so July's
        # share is the INCREMENT over the quarter-end balance (ytd - qtr). That is what keeps
        # both cumulative columns internally consistent at once: Apr+May+Jun already sums to the
        # quarter-end balance (Apr/May get nothing, Jun carries the full 'qtr' figure), so adding
        # the increment here makes Apr..Jul sum to the 31-Jul balance, matching the 'ytd' column.
        # Handing July the full 'ytd' balance instead would double-count the quarter's portion.
        rows = b.load_csv(MOM_DETAIL_DIR / f"pl_jul_{key}.csv")
        rows.append(("Stock Valuation (Inventory)", b.STOCK_VALUATION_SUBGROUP,
                     "Opening Stock", -STOCK_MOM[key][MONTH_OPEN_KEY["jul"]]))
        rows.append(("Stock Valuation (Inventory)", b.STOCK_VALUATION_SUBGROUP,
                     "Closing Stock", STOCK_MOM[key]["jul"]))
        sbtm_inc = (b.SALES_BILLS_TO_MAKE.get(key, {}).get("ytd", 0.0)
                    - b.SALES_BILLS_TO_MAKE.get(key, {}).get("qtr", 0.0))
        if sbtm_inc:
            rows.append(("Sales Accounts", "Sales Bills to Make", "Sales Bills to Make", sbtm_inc))
        pbtc_inc = (b.PURCHASE_BILLS_TO_COME.get(key, {}).get("ytd", 0.0)
                    - b.PURCHASE_BILLS_TO_COME.get(key, {}).get("qtr", 0.0))
        if pbtc_inc:
            rows.append(("Purchase Accounts", "Purchase Bills to Come", "Purchase Bills to Come", pbtc_inc))
        return rows
    if stat_type == "pl" and period == "jun":
        rows = b.load_csv(MOM_DETAIL_DIR / f"pl_jun_{key}.csv")
        open_val = STOCK_MOM[key][MONTH_OPEN_KEY["jun"]]
        close_val = STOCK_MOM[key]["jun"]
        rows.append(("Stock Valuation (Inventory)", b.STOCK_VALUATION_SUBGROUP, "Opening Stock", -open_val))
        rows.append(("Stock Valuation (Inventory)", b.STOCK_VALUATION_SUBGROUP, "Closing Stock", close_val))
        sbtm = b.SALES_BILLS_TO_MAKE.get(key, {}).get("qtr", 0.0)
        if sbtm:
            rows.append(("Sales Accounts", "Sales Bills to Make", "Sales Bills to Make", sbtm))
        pbtc = b.PURCHASE_BILLS_TO_COME.get(key, {}).get("qtr", 0.0)
        if pbtc:
            rows.append(("Purchase Accounts", "Purchase Bills to Come", "Purchase Bills to Come", pbtc))
        return rows
    rows = b.load_csv(MOM_DETAIL_DIR / f"{stat_type}_{period}_{key}.csv")
    if stat_type == "pl":
        open_val = STOCK_MOM[key][MONTH_OPEN_KEY[period]]
        close_val = STOCK_MOM[key][period]
        rows.append(("Stock Valuation (Inventory)", b.STOCK_VALUATION_SUBGROUP, "Opening Stock", -open_val))
        rows.append(("Stock Valuation (Inventory)", b.STOCK_VALUATION_SUBGROUP, "Closing Stock", close_val))
    else:
        close_val = STOCK_MOM[key][period]
        rows.append(("Current Assets", "Closing Stock (Inventory)", "Closing Stock (Inventory)", -close_val))
        # Balance Sheet is cumulative, so the stock movement inside retained earnings runs from
        # the FY-START opening to this month's close - not from the prior month. See
        # apply_stock_to_pl_ac's docstring.
        b.apply_stock_to_pl_ac(rows, b.STOCK[key]["opening"], close_val)
        sbtm = b.SALES_BILLS_TO_MAKE.get(key, {}).get(period, 0.0)
        if sbtm:
            rows.append(("Sundry Debtors", "Sales Bills to Make", "Sales Bills to Make", -sbtm))
        pbtc = b.PURCHASE_BILLS_TO_COME.get(key, {}).get(period, 0.0)
        if pbtc:
            rows.append(("Sundry Creditors", "Purchase Bills to Come", "Purchase Bills to Come", -pbtc))
        forex = b.UNADJUSTED_FOREX_GAIN_LOSS.get(key, {}).get(period, 0.0)
        if forex:
            rows.append(("Unadjusted Forex Gain/Loss", "Unadjusted Forex Gain/Loss", "Unadjusted Forex Gain/Loss", forex))
        forex_asset = b.UNADJUSTED_FOREX_GAIN_LOSS_ASSET.get(key, {}).get(period, 0.0)
        if forex_asset:
            rows.append(("Unadjusted Forex Gain/Loss (Asset)", "Unadjusted Forex Gain/Loss (Asset)",
                          "Unadjusted Forex Gain/Loss (Asset)", -forex_asset))
    return rows


def build_multi_period_tree(stat_type, periods):
    """tree[section][bucket][(sub_grp, gl_name)][period_key] = [co1..co5]. rp_tree (P&L only):
    [bucket][(sub_grp, gl_name)][period_key] = [co1..co5] for the Related Party memo section -
    these rows stay INSIDE tree's TURNOVER/COST OF SALES too (not removed), matching the main
    schedule's design."""
    tree = {}
    rp_tree = {}
    for period_key, _ in periods:
        for idx, (_, key) in enumerate(b.ALL_COMPANIES):
            for pg, sub_grp, gl_name, bal in load_period_rows(stat_type, key, period_key):
                if stat_type == "pl":
                    section, bucket = b.classify_pl_row(pg, sub_grp, gl_name)
                else:
                    section, bucket = b.classify_bs_row(pg)
                # Display sign convention: expenses and assets are shown as POSITIVE magnitudes
                # (see build_gl_detail_sidebyside.DISPLAY_NEGATED_BUCKETS). Flipping the leaf
                # values here is enough for every SUM-based subtotal above them; only the headline
                # waterfall rows and the two mixed memo sections need explicit sign handling.
                node = tree.setdefault(section, {}).setdefault(bucket, {}).setdefault((sub_grp, gl_name), {})
                slot = node.setdefault(period_key, [0.0] * 5)
                slot[idx] += bal * b.display_sign(bucket)
                if stat_type == "pl" and b.is_related_party_row(pg, sub_grp):
                    rp_bucket = "Related Party Sales" if pg == "Sales Accounts" else "Related Party Purchase"
                    rnode = rp_tree.setdefault(rp_bucket, {}).setdefault((sub_grp, gl_name), {})
                    rslot = rnode.setdefault(period_key, [0.0] * 5)
                    rslot[idx] += bal * b.display_sign(rp_bucket)
    return tree, rp_tree


# ---------------------------------------------------------------------------
# Schedule column layout: A Section, B Grouping, C Sub Grouping, D GL Name,
# then one 8-column company block per period (CO1,CO2,TotalOrange,CO3,CO4,TotalEnterprises,
# CO5,GrandTotal) - identical relative layout to PL GL Detail / BS GL Detail, just repeated.
# ---------------------------------------------------------------------------

LABEL_COLS = 4
C_SECTION, C_GROUPING, C_SUBGROUPING, C_GLNAME = 1, 2, 3, 4


def period_start_col(idx):
    return LABEL_COLS + 1 + idx * 8


def period_cols(idx):
    s = period_start_col(idx)
    return {"CO1": s, "CO2": s + 1, "TOTAL_ORANGE": s + 2, "CO3": s + 3, "CO4": s + 4,
            "TOTAL_ENTERPRISES": s + 5, "CO5": s + 6, "GRANDTOTAL": s + 7}


def schedule_total_cols(periods):
    return LABEL_COLS + len(periods) * 8


def shade_row(ws, row, fill, total_cols):
    for c in range(1, total_cols + 1):
        ws.cell(row, c).fill = fill


def write_period_cells(ws, row, period_idx, plain_values, font, skip_zero=False):
    """plain_values: {'CO1':val,...,'CO5':val}, val a number or formula string. Total columns
    (Total of Orange O Tec/Enterprises/Grand Total) are ALWAYS written as formulas referencing
    this period's own CO1..CO5 cells in the same row - never an independently computed number."""
    cols = period_cols(period_idx)
    for key in ("CO1", "CO2", "CO3", "CO4", "CO5"):
        val = plain_values.get(key)
        if val is None:
            continue
        if skip_zero and isinstance(val, (int, float)) and abs(val) < 0.005:
            continue
        cell = ws.cell(row, cols[key], value=val)
        cell.number_format = b.NUM_FORMAT
        cell.font = font
    l = {k: get_column_letter(v) for k, v in cols.items()}
    ws.cell(row, cols["TOTAL_ORANGE"], value=f"={l['CO1']}{row}+{l['CO2']}{row}")
    ws.cell(row, cols["TOTAL_ENTERPRISES"], value=f"={l['CO3']}{row}+{l['CO4']}{row}")
    ws.cell(row, cols["GRANDTOTAL"], value=f"={l['TOTAL_ORANGE']}{row}+{l['TOTAL_ENTERPRISES']}{row}+{l['CO5']}{row}")
    for key in ("TOTAL_ORANGE", "TOTAL_ENTERPRISES", "GRANDTOTAL"):
        c = ws.cell(row, cols[key])
        c.number_format = b.NUM_FORMAT
        c.font = font


def plain_dict5(vals5):
    return {"CO1": vals5[0], "CO2": vals5[1], "CO3": vals5[2], "CO4": vals5[3], "CO5": vals5[4]}


def sum_range_dict5(period_idx, first_row, last_row):
    cols = period_cols(period_idx)
    return {key: f"=SUM({get_column_letter(cols[key])}{first_row}:{get_column_letter(cols[key])}{last_row})"
            for key in ("CO1", "CO2", "CO3", "CO4", "CO5")}


def sum_scattered_dict5(period_idx, rows):
    cols = period_cols(period_idx)
    return {key: "=SUM(" + ",".join(f"{get_column_letter(cols[key])}{r}" for r in rows) + ")"
            for key in ("CO1", "CO2", "CO3", "CO4", "CO5")}


def signed_scattered_dict5(period_idx, rows_with_signs):
    """Signed equivalent of sum_scattered_dict5, for the MIXED memo sections (Branch, Related
    Party) whose income and expense legs must NET once expenses are shown positive."""
    cols = period_cols(period_idx)
    out = {}
    for key in ("CO1", "CO2", "CO3", "CO4", "CO5"):
        l = get_column_letter(cols[key])
        parts = []
        for i, (r, sign) in enumerate(rows_with_signs):
            op = "-" if sign < 0 else ("+" if i else "")
            parts.append(f"{op}{l}{r}")
        out[key] = "=" + "".join(parts) if parts else "=0"
    return out


def write_multi_bucket(ws, row, display_label, node, base_level, periods):
    """node: {(sub_grp, gl_name): {period_key: [co1..co5]}}. Returns (row, bucket_total_row or
    None, subgroup_rows {sub_grp: row}, leaf_rows {(sub_grp, gl_name): row}). leaf_rows is
    additive/read-only bookkeeping - it doesn't change any cell this function already writes -
    added so other sheets (e.g. the PL Main/Sub/Full-GL and per-BS-Group MOM sheets) can reference
    these exact ledger rows by formula instead of recomputing values independently."""
    if not node:
        return row, None, {}, {}

    def combined_mag(period_map):
        return sum(abs(b.col_sum(v)) for v in period_map.values())

    subgroup_lines = {}
    for (sub_grp, gl_name), period_map in node.items():
        subgroup_lines.setdefault(sub_grp, []).append((gl_name, period_map))
    subgroup_order = sorted(subgroup_lines.keys(),
                             key=lambda sg: sum(combined_mag(pm) for _, pm in subgroup_lines[sg]),
                             reverse=True)

    group_start_row = row
    subtotal_rows = []
    subgroup_rows = {}
    leaf_rows = {}
    for sub_grp in subgroup_order:
        lines = subgroup_lines[sub_grp]
        lines.sort(key=lambda item: combined_mag(item[1]), reverse=True)
        subgroup_start_row = row
        first = True
        for gl_name, period_map in lines:
            ws.cell(row, C_SUBGROUPING, value=sub_grp if first else None).font = b.SUBGROUP_FONT
            ws.cell(row, C_GLNAME, value=gl_name).font = b.NORMAL_FONT
            for pidx, (pkey, _) in enumerate(periods):
                vals = period_map.get(pkey, [0.0] * 5)
                write_period_cells(ws, row, pidx, plain_dict5(vals), b.NORMAL_FONT, skip_zero=True)
            ws.row_dimensions[row].outlineLevel = base_level + 2
            leaf_rows[(sub_grp, gl_name)] = row
            first = False
            row += 1
        subgroup_end_row = row - 1
        ws.cell(row, C_GLNAME, value=f"Subtotal: {sub_grp}").font = b.SUBTOTAL_SUBGROUP_FONT
        for pidx in range(len(periods)):
            write_period_cells(ws, row, pidx, sum_range_dict5(pidx, subgroup_start_row, subgroup_end_row),
                                b.SUBTOTAL_SUBGROUP_FONT)
        shade_row(ws, row, b.SUBTOTAL_SUBGROUP_FILL, schedule_total_cols(periods))
        ws.row_dimensions[row].outlineLevel = base_level + 1
        subtotal_rows.append(row)
        subgroup_rows[sub_grp] = row
        row += 1

    ws.cell(group_start_row, C_GROUPING, value=display_label).font = b.GROUP_FONT
    ws.cell(row, C_GROUPING, value=display_label).font = b.GROUP_FONT
    ws.cell(row, C_GLNAME, value=f"TOTAL: {display_label}").font = b.SUBTOTAL_FONT
    for pidx in range(len(periods)):
        write_period_cells(ws, row, pidx, sum_scattered_dict5(pidx, subtotal_rows), b.SUBTOTAL_FONT)
    shade_row(ws, row, b.SUBTOTAL_FILL, schedule_total_cols(periods))
    ws.row_dimensions[row].outlineLevel = base_level
    bucket_total_row = row
    row += 1
    return row, bucket_total_row, subgroup_rows, leaf_rows


def write_multi_section(ws, row, section_label, tree_section, order_map, periods,
                         total_label=None, base_level=1, section_style="normal", label_map=None,
                         bucket_signs=None):
    total_label = total_label or f"TOTAL {section_label}"
    header_font = b.SECTION_FONT if section_style == "normal" else b.ADJUSTMENT_FONT
    header_fill = b.SECTION_FILL if section_style == "normal" else b.ADJUSTMENT_FILL
    total_font = b.SECTION_TOTAL_FONT if section_style == "normal" else b.ADJUSTMENT_TOTAL_FONT
    total_fill = b.SECTION_TOTAL_FILL if section_style == "normal" else b.ADJUSTMENT_TOTAL_FILL
    total_cols = schedule_total_cols(periods)

    ws.cell(row, C_SECTION, value=section_label).font = header_font
    shade_row(ws, row, header_fill, total_cols)
    ws.row_dimensions[row].outlineLevel = 0
    row += 1

    bucket_rows = {}
    detail_rows = {}
    leaf_detail_rows = {}
    bucket_total_rows = []

    merged_order = []
    merged_nodes = {}
    for pg in order_map:
        if pg not in tree_section:
            continue
        label = label_map.get(pg, pg) if label_map else pg
        if label not in merged_nodes:
            merged_nodes[label] = {}
            merged_order.append(label)
        for gl_key, period_map in tree_section[pg].items():
            dest = merged_nodes[label].setdefault(gl_key, {})
            for pkey, vals in period_map.items():
                slot = dest.setdefault(pkey, [0.0] * 5)
                for i in range(5):
                    slot[i] += vals[i]

    for label in merged_order:
        row, bucket_total_row, subgroup_rows, leaf_rows = write_multi_bucket(
            ws, row, label, merged_nodes[label], base_level, periods)
        if bucket_total_row is not None:
            bucket_rows[label] = bucket_total_row
            bucket_total_rows.append(bucket_total_row)
        detail_rows[label] = subgroup_rows
        leaf_detail_rows[label] = leaf_rows

    ws.cell(row, C_GLNAME, value=total_label).font = total_font
    signed_rows = ([(bucket_rows[lbl], bucket_signs.get(lbl, 1))
                    for lbl in bucket_rows if bucket_rows[lbl] in bucket_total_rows]
                   if bucket_signs else None)
    for pidx in range(len(periods)):
        if bucket_total_rows and signed_rows:
            write_period_cells(ws, row, pidx, signed_scattered_dict5(pidx, signed_rows), total_font)
        elif bucket_total_rows:
            write_period_cells(ws, row, pidx, sum_scattered_dict5(pidx, bucket_total_rows), total_font)
        else:
            write_period_cells(ws, row, pidx, plain_dict5([0.0] * 5), total_font)
    shade_row(ws, row, total_fill, total_cols)
    ws.row_dimensions[row].outlineLevel = 0
    row += 1

    return row, bucket_rows, detail_rows, leaf_detail_rows


def write_multi_computed(ws, row, label, component_rows, font, fill, periods):
    ws.cell(row, C_SECTION, value=label).font = font
    for pidx in range(len(periods)):
        cols = period_cols(pidx)
        formulas = {}
        for key in ("CO1", "CO2", "CO3", "CO4", "CO5"):
            l = get_column_letter(cols[key])
            parts = []
            for i, (ref_row, sign) in enumerate(component_rows):
                term = f"{l}{ref_row}"
                parts.append((f"-{term}" if sign == "-" else term) if i == 0
                              else f"{'-' if sign == '-' else '+'}{term}")
            formulas[key] = "=" + "".join(parts)
        write_period_cells(ws, row, pidx, formulas, font)
    shade_row(ws, row, fill, schedule_total_cols(periods))
    ws.row_dimensions[row].outlineLevel = 0
    return row + 1


def write_multi_literal(ws, row, label, font, fill, periods):
    ws.cell(row, C_SECTION, value=label).font = font
    for pidx in range(len(periods)):
        write_period_cells(ws, row, pidx, plain_dict5([0.0] * 5), font)
    shade_row(ws, row, fill, schedule_total_cols(periods))
    ws.row_dimensions[row].outlineLevel = 0
    return row + 1


def write_schedule_header(ws, title, periods):
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 40
    for pidx in range(len(periods)):
        cols = period_cols(pidx)
        for key in ("CO1", "CO2", "CO3", "CO4", "CO5"):
            ws.column_dimensions[get_column_letter(cols[key])].width = 15
        for key in ("TOTAL_ORANGE", "TOTAL_ENTERPRISES", "GRANDTOTAL"):
            ws.column_dimensions[get_column_letter(cols[key])].width = 16
        b.apply_family_column_outline(ws, cols["CO1"], cols["CO2"], cols["TOTAL_ORANGE"], cols["CO3"],
                                       cols["CO4"], cols["TOTAL_ENTERPRISES"], cols["CO5"], cols["GRANDTOTAL"])
    ws.sheet_properties.outlinePr.summaryBelow = True

    row = 1
    ws.cell(row, 1, value=title).font = b.TITLE_FONT
    row += 2

    header_row1 = row
    header_row2 = row + 1
    ws.cell(header_row1, C_SECTION, value="Statement Section")
    ws.cell(header_row1, C_GROUPING, value="Grouping")
    ws.cell(header_row1, C_SUBGROUPING, value="Sub Grouping")
    ws.cell(header_row1, C_GLNAME, value="GL Name")
    for c in (C_SECTION, C_GROUPING, C_SUBGROUPING, C_GLNAME):
        ws.merge_cells(start_row=header_row1, start_column=c, end_row=header_row2, end_column=c)
        cell = ws.cell(header_row1, c)
        cell.font = b.HEADER_FONT
        cell.fill = b.HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")

    for pidx, (pkey, plabel) in enumerate(periods):
        cols = period_cols(pidx)
        start, end = cols["CO1"], cols["GRANDTOTAL"]
        ws.merge_cells(start_row=header_row1, start_column=start, end_row=header_row1, end_column=end)
        cell = ws.cell(header_row1, start, value=plabel)
        cell.font = b.HEADER_FONT
        cell.fill = b.TOTALCOL_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        sub_labels = {cols["CO1"]: b.CO1[0], cols["CO2"]: b.CO2[0], cols["TOTAL_ORANGE"]: "Total of Orange O Tec",
                      cols["CO3"]: b.CO3[0], cols["CO4"]: b.CO4[0], cols["TOTAL_ENTERPRISES"]: "Total of Enterprises",
                      cols["CO5"]: b.CO5[0], cols["GRANDTOTAL"]: "Grand Total"}
        for c, lbl in sub_labels.items():
            cell = ws.cell(header_row2, c, value=lbl)
            cell.font = b.HEADER_FONT
            cell.fill = (b.TOTALCOL_FILL if c in (cols["TOTAL_ORANGE"], cols["TOTAL_ENTERPRISES"], cols["GRANDTOTAL"])
                         else b.HEADER_FILL)
            cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="bottom")
    ws.row_dimensions[header_row1].height = 20
    ws.row_dimensions[header_row2].height = 50
    ws.freeze_panes = ws.cell(header_row2 + 1, period_start_col(0)).coordinate
    return header_row2 + 1


# ---------------------------------------------------------------------------
# "% of Sales" columns (2026-08-04 onward, user request)
# ---------------------------------------------------------------------------
# Three extra columns appended AFTER the last period block, expressing every line as a
# percentage of that period's own Total Turnover. Only the three reporting periods get one
# (Q1 Total / Jul-26 / Apr-Jul Total) - the individual months already sit next to each other for
# comparison, and a % column per month would double the sheet width for little gain.
#
# The user originally added these by hand in Excel. A hand-added column does NOT survive a
# rebuild (both schedules are regenerated from scratch every run), so they are generated here
# instead - that was the whole point of the request.
#
# Every cell is a live formula dividing the row's own Grand Total by the Grand Total of the
# TOTAL TURNOVER row, IFERROR-wrapped so blank/zero-turnover rows render empty rather than
# #DIV/0!. The turnover row reference is absolute ($) so the column can be filled safely.
PCT_FORMAT = "0.0%"
PCT_PERIOD_KEYS = ("qtr", "jul", "ytd")


def add_pct_of_sales_columns(ws, periods, turnover_row, first_row, last_row,
                             header_row1, header_row2):
    """Appends one '% of Sales' column per reporting period. Returns the number added."""
    base = schedule_total_cols(periods)
    added = 0
    for i, pkey in enumerate(PCT_PERIOD_KEYS):
        pidx = next((n for n, (k, _) in enumerate(periods) if k == pkey), None)
        if pidx is None:
            continue
        src = get_column_letter(period_cols(pidx)["GRANDTOTAL"])
        label = dict(periods)[pkey]
        col = base + 1 + added
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 13

        top = ws.cell(header_row1, col, value=label)
        top.font = b.HEADER_FONT
        top.fill = b.TOTALCOL_FILL
        top.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        sub = ws.cell(header_row2, col, value="% of Sales")
        sub.font = b.HEADER_FONT
        sub.fill = b.TOTALCOL_FILL
        sub.alignment = Alignment(horizontal="center", wrap_text=True, vertical="bottom")

        for r in range(first_row, last_row + 1):
            if ws.cell(r, period_cols(pidx)["GRANDTOTAL"]).value in (None, ""):
                continue
            c = ws.cell(r, col, value=f'=IFERROR({src}{r}/{src}${turnover_row},"")')
            c.number_format = PCT_FORMAT
            c.font = b.NORMAL_FONT
        added += 1
    return added


def write_mom_pl_schedule(ws):
    periods = PL_PERIODS
    row = write_schedule_header(ws, "MOM - Profit & Loss A/c GL Detail Schedule (full ledger detail, by month)", periods)
    tree, rp_tree = build_multi_period_tree("pl", periods)
    row_map = {}
    section_detail = {}
    section_buckets = {}
    leaf_detail = {}

    row, section_buckets["TURNOVER"], section_detail["TURNOVER"], leaf_detail["TURNOVER"] = write_multi_section(
        ws, row, "TURNOVER", tree.get("TURNOVER", {}), b.PL_TURNOVER_BUCKETS, periods, total_label="TOTAL TURNOVER")
    row_map["TOTAL TURNOVER"] = row - 1
    row, section_buckets["COST OF SALES"], section_detail["COST OF SALES"], leaf_detail["COST OF SALES"] = write_multi_section(
        ws, row, "LESS: COST OF SALES", tree.get("COST OF SALES", {}), b.PL_COST_OF_SALES_BUCKETS, periods,
        total_label="TOTAL COST OF SALES")
    row_map["TOTAL COST OF SALES"] = row - 1

    row, section_buckets["BRANCH TRANSACTIONS"], section_detail["BRANCH TRANSACTIONS"], leaf_detail["BRANCH TRANSACTIONS"] = write_multi_section(
        ws, row, "BRANCH TRANSACTIONS (Inter-company - shown separately for visibility, included in Gross Profit below)",
        tree.get("BRANCH TRANSACTIONS", {}), b.PL_BRANCH_BUCKETS, periods,
        total_label="NET BRANCH TRANSACTIONS (included in Gross Profit below, matches each company's own Tally P&L)",
        section_style="adjustment", bucket_signs={"Branch Sales": 1, "Branch Purchase": -1})
    row_map["NET BRANCH TRANSACTIONS"] = row - 1

    row = write_multi_computed(ws, row, "GROSS PROFIT",
                                [(row_map["TOTAL TURNOVER"], "+"), (row_map["TOTAL COST OF SALES"], "-"),
                                 (row_map["NET BRANCH TRANSACTIONS"], "+")],
                                b.GP_FONT, b.GP_FILL, periods)
    row_map["GROSS PROFIT"] = row - 1

    row, section_buckets["OTHER INCOME"], section_detail["OTHER INCOME"], leaf_detail["OTHER INCOME"] = write_multi_section(
        ws, row, "ADD: OTHER INCOME", tree.get("OTHER INCOME", {}), b.PL_OTHER_INCOME_BUCKETS, periods,
        total_label="TOTAL OTHER INCOME")
    row_map["TOTAL OTHER INCOME"] = row - 1
    row, section_buckets["OPERATING EXPENSES"], section_detail["OPERATING EXPENSES"], leaf_detail["OPERATING EXPENSES"] = write_multi_section(
        ws, row, "LESS: OPERATING EXPENSES", tree.get("OPERATING EXPENSES", {}), b.PL_OPERATING_EXPENSE_BUCKETS,
        periods, total_label="TOTAL OPERATING EXPENSES")
    row_map["TOTAL OPERATING EXPENSES"] = row - 1

    row = write_multi_computed(ws, row, "EBITDA",
                                [(row_map["GROSS PROFIT"], "+"), (row_map["TOTAL OTHER INCOME"], "+"),
                                 (row_map["TOTAL OPERATING EXPENSES"], "-")], b.EBITDA_FONT, b.EBITDA_FILL, periods)
    row_map["EBITDA"] = row - 1

    row, section_buckets["FINANCE COST"], section_detail["FINANCE COST"], leaf_detail["FINANCE COST"] = write_multi_section(
        ws, row, "LESS: FINANCE COST", tree.get("FINANCE COST", {}), b.PL_FINANCE_COST_BUCKETS, periods,
        total_label="TOTAL FINANCE COST")
    row_map["TOTAL FINANCE COST"] = row - 1

    row = write_multi_literal(ws, row, "LESS: DEPRECIATION (not booked this period)", b.ZERO_FONT, b.ZERO_FILL, periods)
    row_map["LESS: DEPRECIATION"] = row - 1

    row = write_multi_computed(ws, row, "PBT (PROFIT BEFORE TAX)",
                                [(row_map["EBITDA"], "+"), (row_map["TOTAL FINANCE COST"], "-"),
                                 (row_map["LESS: DEPRECIATION"], "-")], b.PBT_FONT, b.PBT_FILL, periods)
    row_map["PBT"] = row - 1

    row = write_multi_literal(ws, row, "LESS: TAX (not booked this period)", b.ZERO_FONT, b.ZERO_FILL, periods)
    row_map["LESS: TAX"] = row - 1

    row = write_multi_computed(ws, row, "PAT (NET PROFIT)",
                                [(row_map["PBT"], "+"), (row_map["LESS: TAX"], "-")], b.PAT_FONT, b.PAT_FILL, periods)
    row_map["PAT"] = row - 1

    row, section_buckets["RELATED PARTY TRANSACTIONS"], section_detail["RELATED PARTY TRANSACTIONS"], leaf_detail["RELATED PARTY TRANSACTIONS"] = write_multi_section(
        ws, row, "RELATED PARTY TRANSACTIONS (included in PAT above - shown for reference)",
        rp_tree, b.PL_RELATED_PARTY_BUCKETS, periods,
        total_label="NET RELATED PARTY TRANSACTIONS", section_style="adjustment",
        bucket_signs={"Related Party Sales": 1, "Related Party Purchase": -1})
    row_map["NET RELATED PARTY TRANSACTIONS"] = row - 1

    row = write_multi_computed(ws, row, "PAT (NET OF RELATED PARTY TRANSACTIONS)",
                                [(row_map["PAT"], "+"), (row_map["NET RELATED PARTY TRANSACTIONS"], "-")],
                                b.GRANDTOTAL_FONT, b.GRANDTOTAL_FILL, periods)
    row_map["PAT (NET OF RELATED PARTY)"] = row - 1

    row = write_mom_valuation_bridge(ws, row, tree, periods, row_map)

    add_pct_of_sales_columns(ws, periods, row_map["TOTAL TURNOVER"],
                             first_row=5, last_row=row - 1,
                             header_row1=3, header_row2=4)

    return row_map, section_buckets, section_detail, leaf_detail


def write_mom_valuation_bridge(ws, row, tree, periods, row_map):
    """Per-period version of the Valuation-related Provisions bridge that PL GL Detail / PL
    Summary already carry (see build_gl_detail_sidebyside.write_pl_sheet and SKILL.md
    "Valuation-related provisions bridge").

    Same selection rule as the quarter sheets: whatever Tally files DIRECTLY under Direct
    Expenses with no finer Sub Grouping of its own - in practice PROVISION FOR IMPORT-FOC(PL),
    INVENTORY OVERHEAD COST and EXCHANGE RATE DIFF EXP ACCOUNTS. These are the most volatile,
    judgment-heavy items in the statement and are repeatedly the top PAT drivers, so they get
    called out as a standalone bridge rather than staying buried inside Cost of Sales.

    Added 2026-08-04: the bridge previously existed ONLY on the two quarter sheets, so the
    month-on-month view - the one actually used to compare periods - could not show how much of
    each month's profit depended on them. Profit Before Valuation is PAT net of related party
    LESS this block; Profit After Valuation is that same PAT unchanged (a lens on PAT, not a new
    PAT), exactly as on the quarter sheets."""
    node = tree.get("COST OF SALES", {}).get("Direct Expenses", {})
    valuation_node = {k: v for k, v in node.items() if k[0] == "Direct Expenses"}
    if not valuation_node:
        return row

    total_cols = schedule_total_cols(periods)
    row += 1
    ws.cell(row, C_SECTION,
            value="VALUATION-RELATED PROVISIONS (most volatile Direct Expenses items - bridge to reported PAT)"
            ).font = b.ADJUSTMENT_FONT
    shade_row(ws, row, b.ADJUSTMENT_FILL, total_cols)
    ws.row_dimensions[row].outlineLevel = 0
    row += 1

    row, valuation_total_row, _, _ = write_multi_bucket(
        ws, row, "Valuation-related Provisions", valuation_node, 1, periods)

    row = write_multi_computed(ws, row, "Profit Before Valuation (excl. items above)",
                               [(row_map["PAT (NET OF RELATED PARTY)"], "+"),
                                (valuation_total_row, "-")],
                               b.GRANDTOTAL_FONT, b.ADJUSTMENT_FILL, periods)
    row_map["PROFIT BEFORE VALUATION"] = row - 1

    row = write_multi_computed(ws, row, "Profit After Valuation (= Reported PAT)",
                               [(row_map["PAT (NET OF RELATED PARTY)"], "+")],
                               b.GRANDTOTAL_FONT, b.GRANDTOTAL_FILL, periods)
    row_map["PROFIT AFTER VALUATION"] = row - 1
    return row


def write_mom_bs_schedule(ws):
    periods = BS_PERIODS
    row = write_schedule_header(
        ws, "MOM - Balance Sheet GL Detail Schedule (full ledger detail, by month)  "
            "[Point-in-time: each month is its own snapshot, there is no Q1 Total column]", periods)
    tree, _ = build_multi_period_tree("bs", periods)
    row_map = {}
    section_detail = {}
    section_buckets = {}
    leaf_detail = {}

    row, liability_bucket_rows, section_detail["LIABILITIES"], leaf_detail["LIABILITIES"] = write_multi_section(
        ws, row, "LIABILITIES (Sources of Funds)", tree.get("LIABILITIES", {}), b.BS_LIABILITY_ORDER, periods,
        total_label="TOTAL LIABILITIES", label_map=b.BS_GROUP_LABEL)
    row_map["TOTAL LIABILITIES"] = row - 1
    row_map.update(liability_bucket_rows)
    section_buckets["LIABILITIES"] = liability_bucket_rows

    row, asset_bucket_rows, section_detail["ASSETS"], leaf_detail["ASSETS"] = write_multi_section(
        ws, row, "ASSETS (Application of Funds)", tree.get("ASSETS", {}), b.BS_ASSET_ORDER, periods,
        total_label="TOTAL ASSETS", label_map=b.BS_GROUP_LABEL)
    row_map["TOTAL ASSETS"] = row - 1
    row_map.update(asset_bucket_rows)
    section_buckets["ASSETS"] = asset_bucket_rows

    row = write_multi_computed(ws, row, "RECONCILIATION GAP (Liabilities - Assets; ~0 expected)",
                                [(row_map["TOTAL LIABILITIES"], "+"), (row_map["TOTAL ASSETS"], "-")],
                                b.GAP_FONT, b.GAP_FILL, periods)
    row_map["RECONCILIATION GAP"] = row - 1

    return row_map, section_buckets, section_detail, leaf_detail


# ---------------------------------------------------------------------------
# "MOM PL and BS": consolidated (Grand Total only), trended across periods.
# Columns: Label | Sub Group | one column per period (linked to that period's Grand Total).
# ---------------------------------------------------------------------------

def write_consolidated_summary_header(ws, title, periods, start_row):
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 24
    for i in range(len(periods)):
        ws.column_dimensions[get_column_letter(3 + i)].width = 16

    row = start_row
    ws.cell(row, 1, value=title).font = b.TITLE_FONT
    row += 2
    header_row = row
    ws.cell(row, 1, value="Line Item")
    ws.cell(row, 2, value="Sub Group")
    for i, (_, plabel) in enumerate(periods):
        ws.cell(row, 3 + i, value=plabel)
    total_cols = 2 + len(periods)
    for c in range(1, total_cols + 1):
        cell = ws.cell(row, c)
        cell.font = b.HEADER_FONT
        cell.fill = b.TOTALCOL_FILL if c == total_cols else b.HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="bottom")
    ws.row_dimensions[row].height = 40
    ws.freeze_panes = f"C{header_row + 1}"
    return header_row + 1


def write_consolidated_linked_row(ws, row, label, schedule_name, schedule_row, font, fill, periods, subgroup=None):
    ws.cell(row, 1, value=label).font = font
    if subgroup:
        ws.cell(row, 2, value=subgroup).font = font
    for i in range(len(periods)):
        cols = period_cols(i)
        gt = get_column_letter(cols["GRANDTOTAL"])
        cell = ws.cell(row, 3 + i, value=f"='{schedule_name}'!{gt}{schedule_row}")
        cell.number_format = b.NUM_FORMAT
        cell.font = font
    if fill:
        for c in range(1, 2 + len(periods) + 1):
            ws.cell(row, c).fill = fill


def write_consolidated_grouping_header(ws, row, label, periods):
    ws.cell(row, 1, value=label).font = b.SECTION_FONT
    for c in range(1, 2 + len(periods) + 1):
        ws.cell(row, c).fill = b.SECTION_FILL


def write_consolidated_subgroup_rows(ws, row, schedule_name, bucket_rows, periods):
    for label, srow in bucket_rows.items():
        write_consolidated_linked_row(ws, row, "", schedule_name, srow, b.SUBGROUP_ROW_FONT, None, periods, subgroup=label)
        row += 1
    return row


def write_consolidated_profit_loss_row(ws, row, label, schedule_name, schedule_row, periods):
    ws.cell(row, 1, value=label).font = b.PROFIT_LOSS_LABEL_FONT
    for i in range(len(periods)):
        cols = period_cols(i)
        gt = get_column_letter(cols["GRANDTOTAL"])
        cell = ws.cell(row, 3 + i, value=f"='{schedule_name}'!{gt}{schedule_row}")
        cell.number_format = b.NUM_FORMAT
        cell.font = b.PROFIT_LOSS_VALUE_FONT
    value_range = f"C{row}:{get_column_letter(2 + len(periods))}{row}"
    ws.conditional_formatting.add(value_range, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=b.PROFIT_FILL))
    ws.conditional_formatting.add(value_range, CellIsRule(operator="lessThan", formula=["0"], fill=b.LOSS_FILL))


def write_mom_pl_consolidated_summary(ws, schedule_name, row_map, section_buckets, section_detail, start_row):
    periods = PL_PERIODS
    row = write_consolidated_summary_header(ws, "MOM - Profit & Loss A/c Summary (all companies consolidated)",
                                             periods, start_row)

    write_consolidated_grouping_header(ws, row, "TURNOVER", periods); row += 1
    row = write_consolidated_subgroup_rows(ws, row, schedule_name, section_buckets.get("TURNOVER", {}), periods)
    write_consolidated_linked_row(ws, row, "Total Turnover", schedule_name, row_map["TOTAL TURNOVER"],
                                   b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_consolidated_grouping_header(ws, row, "LESS: COST OF SALES", periods); row += 1
    row = write_consolidated_subgroup_rows(ws, row, schedule_name, section_buckets.get("COST OF SALES", {}), periods)
    write_consolidated_linked_row(ws, row, "Total Cost of Sales", schedule_name, row_map["TOTAL COST OF SALES"],
                                   b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_consolidated_grouping_header(ws, row, "BRANCH TRANSACTIONS (shown separately, included in Gross Profit)", periods); row += 1
    row = write_consolidated_subgroup_rows(ws, row, schedule_name, section_buckets.get("BRANCH TRANSACTIONS", {}), periods)
    write_consolidated_linked_row(ws, row, "Net Branch Transactions", schedule_name, row_map["NET BRANCH TRANSACTIONS"],
                                   b.ADJUSTMENT_TOTAL_FONT, b.ADJUSTMENT_TOTAL_FILL, periods); row += 1

    write_consolidated_profit_loss_row(ws, row, "GROSS PROFIT", schedule_name, row_map["GROSS PROFIT"], periods); row += 1

    write_consolidated_grouping_header(ws, row, "ADD: OTHER INCOME", periods); row += 1
    row = write_consolidated_subgroup_rows(ws, row, schedule_name, section_buckets.get("OTHER INCOME", {}), periods)
    write_consolidated_linked_row(ws, row, "Total Other Income", schedule_name, row_map["TOTAL OTHER INCOME"],
                                   b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_consolidated_grouping_header(ws, row, "LESS: OPERATING EXPENSES", periods); row += 1
    indirect = section_detail.get("OPERATING EXPENSES", {}).get("Indirect Expenses", {})
    row = write_consolidated_subgroup_rows(ws, row, schedule_name, indirect, periods)
    write_consolidated_linked_row(ws, row, "Total Operating Expenses", schedule_name, row_map["TOTAL OPERATING EXPENSES"],
                                   b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_consolidated_profit_loss_row(ws, row, "EBITDA", schedule_name, row_map["EBITDA"], periods); row += 1

    write_consolidated_grouping_header(ws, row, "LESS: FINANCE COST", periods); row += 1
    row = write_consolidated_subgroup_rows(ws, row, schedule_name, section_buckets.get("FINANCE COST", {}), periods)
    write_consolidated_linked_row(ws, row, "Total Finance Cost", schedule_name, row_map["TOTAL FINANCE COST"],
                                   b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1
    write_consolidated_linked_row(ws, row, "Less: Depreciation", schedule_name, row_map["LESS: DEPRECIATION"],
                                   b.ZERO_FONT, b.ZERO_FILL, periods); row += 1

    write_consolidated_profit_loss_row(ws, row, "PBT (Profit Before Tax)", schedule_name, row_map["PBT"], periods); row += 1

    write_consolidated_linked_row(ws, row, "Less: Tax", schedule_name, row_map["LESS: TAX"], b.ZERO_FONT, b.ZERO_FILL,
                                   periods); row += 1

    write_consolidated_profit_loss_row(ws, row, "PAT (Net Profit)", schedule_name, row_map["PAT"], periods); row += 1

    write_consolidated_grouping_header(ws, row, "RELATED PARTY TRANSACTIONS (reference only)", periods); row += 1
    row = write_consolidated_subgroup_rows(ws, row, schedule_name, section_buckets.get("RELATED PARTY TRANSACTIONS", {}), periods)
    write_consolidated_linked_row(ws, row, "Net Related Party Transactions", schedule_name,
                                   row_map["NET RELATED PARTY TRANSACTIONS"], b.ADJUSTMENT_TOTAL_FONT,
                                   b.ADJUSTMENT_TOTAL_FILL, periods); row += 1

    write_consolidated_profit_loss_row(ws, row, "PAT (Net of Related Party Transactions)", schedule_name,
                                        row_map["PAT (NET OF RELATED PARTY)"], periods)
    row += 1

    # Valuation bridge, mirroring the schedule and the quarter summary.
    if "PROFIT BEFORE VALUATION" in row_map:
        row += 1
        ws.cell(row, 1, value="VALUATION-RELATED PROVISIONS (bridge to reported PAT)").font = b.ADJUSTMENT_FONT
        for c in range(1, 2 + len(periods) + 1):
            ws.cell(row, c).fill = b.ADJUSTMENT_FILL
        row += 1
        write_consolidated_linked_row(ws, row, "Profit Before Valuation (excl. items above)", schedule_name,
                                       row_map["PROFIT BEFORE VALUATION"], b.GRANDTOTAL_FONT,
                                       b.ADJUSTMENT_FILL, periods); row += 1
        write_consolidated_linked_row(ws, row, "Profit After Valuation (= Reported PAT)", schedule_name,
                                       row_map["PROFIT AFTER VALUATION"], b.GRANDTOTAL_FONT,
                                       b.GRANDTOTAL_FILL, periods); row += 1

    add_summary_pct_columns(ws, periods, turnover_row=find_row(ws, "Total Turnover"),
                            first_row=start_row + 3, last_row=row - 1,
                            header_row=start_row + 2, label_cols=2)
    return row + 2


def add_summary_pct_columns(ws, periods, turnover_row, first_row, last_row, header_row,
                            label_cols=2):
    """'% of Sales' for the SUMMARY sheets, whose periods sit in consecutive single columns
    (label_cols + 1 ...) rather than in 8-wide blocks like the schedules. Same three reporting
    periods, same IFERROR-over-own-column rule - see add_pct_of_sales_columns."""
    if not turnover_row:
        return 0
    base = label_cols + len(periods)
    added = 0
    for pkey in PCT_PERIOD_KEYS:
        pidx = next((n for n, (k, _) in enumerate(periods) if k == pkey), None)
        if pidx is None:
            continue
        src = get_column_letter(label_cols + 1 + pidx)
        col = base + 1 + added
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 13
        head = ws.cell(header_row, col, value=f"{dict(periods)[pkey]} % of Sales")
        head.font = b.HEADER_FONT
        head.fill = b.TOTALCOL_FILL
        head.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        for r in range(first_row, last_row + 1):
            if ws.cell(r, label_cols + 1 + pidx).value in (None, ""):
                continue
            c = ws.cell(r, col, value=f'=IFERROR({src}{r}/{src}${turnover_row},"")')
            c.number_format = PCT_FORMAT
            c.font = b.NORMAL_FONT
        added += 1
    return added


def write_mom_bs_consolidated_summary(ws, schedule_name, row_map, section_buckets, start_row):
    periods = BS_PERIODS
    row = write_consolidated_summary_header(
        ws, "MOM - Balance Sheet Summary (all companies consolidated)  "
            "[Point-in-time: use each month's own column, do not sum across months]",
        periods, start_row)

    write_consolidated_grouping_header(ws, row, "LIABILITIES (Sources of Funds)", periods); row += 1
    row = write_consolidated_subgroup_rows(ws, row, schedule_name, section_buckets.get("LIABILITIES", {}), periods)
    write_consolidated_linked_row(ws, row, "Total Liabilities", schedule_name, row_map["TOTAL LIABILITIES"],
                                   b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_consolidated_grouping_header(ws, row, "ASSETS (Application of Funds)", periods); row += 1
    row = write_consolidated_subgroup_rows(ws, row, schedule_name, section_buckets.get("ASSETS", {}), periods)
    write_consolidated_linked_row(ws, row, "Total Assets", schedule_name, row_map["TOTAL ASSETS"],
                                   b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_consolidated_linked_row(ws, row, "Reconciliation Gap (~0 expected)", schedule_name,
                                   row_map["RECONCILIATION GAP"], b.GAP_FONT, b.GAP_FILL, periods)


# ---------------------------------------------------------------------------
# "MOM PL and BS - By Company": 8 company/family blocks x period sub-columns (kept the same
# visual layout as before), values linked to the schedules above.
# ---------------------------------------------------------------------------

BLOCKS = [
    ("CO1", b.CO1[0], "CO1"), ("CO2", b.CO2[0], "CO2"), ("ORANGE", "Total of Orange O Tec", "TOTAL_ORANGE"),
    ("CO3", b.CO3[0], "CO3"), ("CO4", b.CO4[0], "CO4"), ("ENTERPRISES", "Total of Enterprises", "TOTAL_ENTERPRISES"),
    ("CO5", b.CO5[0], "CO5"), ("GRAND", "Grand Total", "GRANDTOTAL"),
]
TOTAL_BLOCK_KEYS = {"ORANGE", "ENTERPRISES", "GRAND"}
W_LABEL_COLS = 2


def by_company_block_start(block_idx, periods):
    return W_LABEL_COLS + 1 + block_idx * len(periods)


def by_company_total_cols(periods):
    return W_LABEL_COLS + len(BLOCKS) * len(periods)


def write_by_company_summary_header(ws, title, periods, start_row):
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 24
    for bidx in range(len(BLOCKS)):
        start = by_company_block_start(bidx, periods)
        for i in range(len(periods)):
            ws.column_dimensions[get_column_letter(start + i)].width = 15
    for bidx, (bkey, _, _) in enumerate(BLOCKS):
        start = by_company_block_start(bidx, periods)
        level = 0 if bkey == "GRAND" else (1 if bkey in TOTAL_BLOCK_KEYS else 2)
        for i in range(len(periods)):
            ws.column_dimensions[get_column_letter(start + i)].outlineLevel = level
    ws.sheet_properties.outlinePr.summaryRight = True

    row = start_row
    ws.cell(row, 1, value=title).font = b.TITLE_FONT
    row += 2
    header_row1, header_row2 = row, row + 1
    ws.cell(header_row1, 1, value="Line Item")
    ws.cell(header_row1, 2, value="Sub Group")
    ws.merge_cells(start_row=header_row1, start_column=1, end_row=header_row2, end_column=1)
    ws.merge_cells(start_row=header_row1, start_column=2, end_row=header_row2, end_column=2)
    for c in (1, 2):
        cell = ws.cell(header_row1, c)
        cell.font = b.HEADER_FONT
        cell.fill = b.HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    for bidx, (bkey, blabel, _) in enumerate(BLOCKS):
        start = by_company_block_start(bidx, periods)
        end = start + len(periods) - 1
        ws.merge_cells(start_row=header_row1, start_column=start, end_row=header_row1, end_column=end)
        cell = ws.cell(header_row1, start, value=blabel)
        cell.font = b.HEADER_FONT
        cell.fill = b.TOTALCOL_FILL if bkey in TOTAL_BLOCK_KEYS else b.HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        for i, (_, plabel) in enumerate(periods):
            c2 = ws.cell(header_row2, start + i, value=plabel)
            c2.font = b.HEADER_FONT
            c2.fill = b.TOTALCOL_FILL if bkey in TOTAL_BLOCK_KEYS else b.HEADER_FILL
            c2.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    ws.row_dimensions[header_row1].height = 20
    ws.row_dimensions[header_row2].height = 20
    ws.freeze_panes = ws.cell(header_row2 + 1, W_LABEL_COLS + 1).coordinate
    return header_row2 + 1


def write_by_company_linked_row(ws, row, label, schedule_name, schedule_row, font, fill, periods, subgroup=None):
    ws.cell(row, 1, value=label).font = font
    if subgroup:
        ws.cell(row, 2, value=subgroup).font = font
    for bidx, (_, _, field) in enumerate(BLOCKS):
        start = by_company_block_start(bidx, periods)
        for i in range(len(periods)):
            cols = period_cols(i)
            col_letter = get_column_letter(cols[field])
            cell = ws.cell(row, start + i, value=f"='{schedule_name}'!{col_letter}{schedule_row}")
            cell.number_format = b.NUM_FORMAT
            cell.font = font
    if fill:
        for c in range(1, by_company_total_cols(periods) + 1):
            ws.cell(row, c).fill = fill


def write_by_company_subgroup_rows(ws, row, schedule_name, bucket_rows, periods):
    for label, srow in bucket_rows.items():
        write_by_company_linked_row(ws, row, "", schedule_name, srow, b.SUBGROUP_ROW_FONT, None, periods, subgroup=label)
        row += 1
    return row


def write_by_company_grouping_header(ws, row, label, periods):
    ws.cell(row, 1, value=label).font = b.SECTION_FONT
    for c in range(1, by_company_total_cols(periods) + 1):
        ws.cell(row, c).fill = b.SECTION_FILL


def write_by_company_profit_loss_row(ws, row, label, schedule_name, schedule_row, periods):
    ws.cell(row, 1, value=label).font = b.PROFIT_LOSS_LABEL_FONT
    for bidx, (_, _, field) in enumerate(BLOCKS):
        start = by_company_block_start(bidx, periods)
        for i in range(len(periods)):
            cols = period_cols(i)
            col_letter = get_column_letter(cols[field])
            cell = ws.cell(row, start + i, value=f"='{schedule_name}'!{col_letter}{schedule_row}")
            cell.number_format = b.NUM_FORMAT
            cell.font = b.PROFIT_LOSS_VALUE_FONT
    value_range = f"{get_column_letter(W_LABEL_COLS + 1)}{row}:{get_column_letter(by_company_total_cols(periods))}{row}"
    ws.conditional_formatting.add(value_range, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=b.PROFIT_FILL))
    ws.conditional_formatting.add(value_range, CellIsRule(operator="lessThan", formula=["0"], fill=b.LOSS_FILL))


def write_mom_pl_by_company_summary(ws, schedule_name, row_map, section_buckets, section_detail, start_row):
    periods = PL_PERIODS
    row = write_by_company_summary_header(ws, "MOM - Profit & Loss A/c (by company)", periods, start_row)

    write_by_company_grouping_header(ws, row, "TURNOVER", periods); row += 1
    row = write_by_company_subgroup_rows(ws, row, schedule_name, section_buckets.get("TURNOVER", {}), periods)
    write_by_company_linked_row(ws, row, "Total Turnover", schedule_name, row_map["TOTAL TURNOVER"],
                                 b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_by_company_grouping_header(ws, row, "LESS: COST OF SALES", periods); row += 1
    row = write_by_company_subgroup_rows(ws, row, schedule_name, section_buckets.get("COST OF SALES", {}), periods)
    write_by_company_linked_row(ws, row, "Total Cost of Sales", schedule_name, row_map["TOTAL COST OF SALES"],
                                 b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_by_company_grouping_header(ws, row, "BRANCH TRANSACTIONS (shown separately, included in Gross Profit)", periods); row += 1
    row = write_by_company_subgroup_rows(ws, row, schedule_name, section_buckets.get("BRANCH TRANSACTIONS", {}), periods)
    write_by_company_linked_row(ws, row, "Net Branch Transactions", schedule_name, row_map["NET BRANCH TRANSACTIONS"],
                                 b.ADJUSTMENT_TOTAL_FONT, b.ADJUSTMENT_TOTAL_FILL, periods); row += 1

    write_by_company_profit_loss_row(ws, row, "GROSS PROFIT", schedule_name, row_map["GROSS PROFIT"], periods); row += 1

    write_by_company_grouping_header(ws, row, "ADD: OTHER INCOME", periods); row += 1
    row = write_by_company_subgroup_rows(ws, row, schedule_name, section_buckets.get("OTHER INCOME", {}), periods)
    write_by_company_linked_row(ws, row, "Total Other Income", schedule_name, row_map["TOTAL OTHER INCOME"],
                                 b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_by_company_grouping_header(ws, row, "LESS: OPERATING EXPENSES", periods); row += 1
    indirect = section_detail.get("OPERATING EXPENSES", {}).get("Indirect Expenses", {})
    row = write_by_company_subgroup_rows(ws, row, schedule_name, indirect, periods)
    write_by_company_linked_row(ws, row, "Total Operating Expenses", schedule_name, row_map["TOTAL OPERATING EXPENSES"],
                                 b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_by_company_profit_loss_row(ws, row, "EBITDA", schedule_name, row_map["EBITDA"], periods); row += 1

    write_by_company_grouping_header(ws, row, "LESS: FINANCE COST", periods); row += 1
    row = write_by_company_subgroup_rows(ws, row, schedule_name, section_buckets.get("FINANCE COST", {}), periods)
    write_by_company_linked_row(ws, row, "Total Finance Cost", schedule_name, row_map["TOTAL FINANCE COST"],
                                 b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1
    write_by_company_linked_row(ws, row, "Less: Depreciation", schedule_name, row_map["LESS: DEPRECIATION"],
                                 b.ZERO_FONT, b.ZERO_FILL, periods); row += 1

    write_by_company_profit_loss_row(ws, row, "PBT (Profit Before Tax)", schedule_name, row_map["PBT"], periods); row += 1

    write_by_company_linked_row(ws, row, "Less: Tax", schedule_name, row_map["LESS: TAX"], b.ZERO_FONT, b.ZERO_FILL,
                                 periods); row += 1

    write_by_company_profit_loss_row(ws, row, "PAT (Net Profit)", schedule_name, row_map["PAT"], periods); row += 1

    write_by_company_grouping_header(ws, row, "RELATED PARTY TRANSACTIONS (reference only)", periods); row += 1
    row = write_by_company_subgroup_rows(ws, row, schedule_name, section_buckets.get("RELATED PARTY TRANSACTIONS", {}), periods)
    write_by_company_linked_row(ws, row, "Net Related Party Transactions", schedule_name,
                                 row_map["NET RELATED PARTY TRANSACTIONS"], b.ADJUSTMENT_TOTAL_FONT,
                                 b.ADJUSTMENT_TOTAL_FILL, periods); row += 1

    write_by_company_profit_loss_row(ws, row, "PAT (Net of Related Party Transactions)", schedule_name,
                                      row_map["PAT (NET OF RELATED PARTY)"], periods)
    row += 1

    if "PROFIT BEFORE VALUATION" in row_map:
        row += 1
        ws.cell(row, 1, value="VALUATION-RELATED PROVISIONS (bridge to reported PAT)").font = b.ADJUSTMENT_FONT
        for c in range(1, by_company_total_cols(periods) + 1):
            ws.cell(row, c).fill = b.ADJUSTMENT_FILL
        row += 1
        write_by_company_linked_row(ws, row, "Profit Before Valuation (excl. items above)", schedule_name,
                                     row_map["PROFIT BEFORE VALUATION"], b.GRANDTOTAL_FONT,
                                     b.ADJUSTMENT_FILL, periods); row += 1
        write_by_company_linked_row(ws, row, "Profit After Valuation (= Reported PAT)", schedule_name,
                                     row_map["PROFIT AFTER VALUATION"], b.GRANDTOTAL_FONT,
                                     b.GRANDTOTAL_FILL, periods); row += 1

    # Grand Total block only (user decision 2026-08-04) - a % column per company block would add
    # 24 columns to an already very wide sheet.
    gt_block = by_company_block_start(len(BLOCKS) - 1, periods)
    add_by_company_pct_columns(ws, periods, gt_block, turnover_row=find_row(ws, "Total Turnover"),
                               first_row=start_row + 3, last_row=row - 1,
                               header_row=start_row + 2)
    return row + 2


def add_by_company_pct_columns(ws, periods, gt_block_start, turnover_row, first_row, last_row,
                               header_row):
    """'% of Sales' for the By-Company summary, computed off the GRAND TOTAL block only."""
    if not turnover_row:
        return 0
    base = by_company_total_cols(periods)
    added = 0
    for pkey in PCT_PERIOD_KEYS:
        pidx = next((n for n, (k, _) in enumerate(periods) if k == pkey), None)
        if pidx is None:
            continue
        src = get_column_letter(gt_block_start + pidx)
        col = base + 1 + added
        ws.column_dimensions[get_column_letter(col)].width = 13
        head = ws.cell(header_row, col, value=f"{dict(periods)[pkey]} % of Sales")
        head.font = b.HEADER_FONT
        head.fill = b.TOTALCOL_FILL
        head.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        for r in range(first_row, last_row + 1):
            if ws.cell(r, gt_block_start + pidx).value in (None, ""):
                continue
            c = ws.cell(r, col, value=f'=IFERROR({src}{r}/{src}${turnover_row},"")')
            c.number_format = PCT_FORMAT
            c.font = b.NORMAL_FONT
        added += 1
    return added


def write_mom_bs_by_company_summary(ws, schedule_name, row_map, section_buckets, start_row):
    periods = BS_PERIODS
    row = write_by_company_summary_header(
        ws, "MOM - Balance Sheet (by company)  "
            "[Point-in-time: use each month's own column, do not sum across months]",
        periods, start_row)

    write_by_company_grouping_header(ws, row, "LIABILITIES (Sources of Funds)", periods); row += 1
    row = write_by_company_subgroup_rows(ws, row, schedule_name, section_buckets.get("LIABILITIES", {}), periods)
    write_by_company_linked_row(ws, row, "Total Liabilities", schedule_name, row_map["TOTAL LIABILITIES"],
                                 b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_by_company_grouping_header(ws, row, "ASSETS (Application of Funds)", periods); row += 1
    row = write_by_company_subgroup_rows(ws, row, schedule_name, section_buckets.get("ASSETS", {}), periods)
    write_by_company_linked_row(ws, row, "Total Assets", schedule_name, row_map["TOTAL ASSETS"],
                                 b.SECTION_TOTAL_FONT, b.SECTION_TOTAL_FILL, periods); row += 1

    write_by_company_linked_row(ws, row, "Reconciliation Gap (~0 expected)", schedule_name,
                                 row_map["RECONCILIATION GAP"], b.GAP_FONT, b.GAP_FILL, periods)


# ---------------------------------------------------------------------------
# "MOM Ratio Analysis": the same ratio definitions and Summary/Component Figures/Calculations
# three-section layout as the quarterly "Ratio Analysis" sheet (build_gl_detail_sidebyside.
# ratio_definitions / ratio_sheet_layout / write_ratio_summary_section / write_ratio_workings_
# section), computed independently for each month from that month's own MOM PL/BS GL Detail
# schedule columns - not a month-over-month average or a slice of the quarter total. Quarterly
# Ratio Analysis remains the Q1 view; this sheet is the monthly view; both are always kept in sync
# going forward (see fetch-pandl SKILL.md).
# ---------------------------------------------------------------------------

MONTH_DAYS = {"apr": 30, "may": 31, "jun": 30, "jul": 31}
RATIO_LABEL_COLS = 1


def ratio_period_start_col(idx):
    return RATIO_LABEL_COLS + 1 + idx * 8


def ratio_period_cols(idx):
    s = ratio_period_start_col(idx)
    return {"CO1": s, "CO2": s + 1, "TOTAL_ORANGE": s + 2, "CO3": s + 3, "CO4": s + 4,
            "TOTAL_ENTERPRISES": s + 5, "CO5": s + 6, "GRANDTOTAL": s + 7}


def write_mom_ratio_sheet(ws, pl_row_map, bs_row_map):
    periods = BS_PERIODS  # Apr/May/Jun only - Q1 Total is already the quarterly Ratio Analysis sheet

    ws.column_dimensions["A"].width = 30
    for pidx in range(len(periods)):
        cols = ratio_period_cols(pidx)
        for key in ("CO1", "CO2", "CO3", "CO4", "CO5"):
            ws.column_dimensions[get_column_letter(cols[key])].width = 15
        for key in ("TOTAL_ORANGE", "TOTAL_ENTERPRISES", "GRANDTOTAL"):
            ws.column_dimensions[get_column_letter(cols[key])].width = 16
        b.apply_family_column_outline(ws, cols["CO1"], cols["CO2"], cols["TOTAL_ORANGE"], cols["CO3"],
                                       cols["CO4"], cols["TOTAL_ENTERPRISES"], cols["CO5"], cols["GRANDTOTAL"])
    ws.sheet_properties.outlinePr.summaryRight = True

    row = 1
    ws.cell(row, 1, value="MOM Ratio Analysis").font = b.TITLE_FONT
    row += 1
    ws.cell(row, 1, value=(
        "Each month's ratios are computed independently from that month's own MOM PL/BS GL Detail "
        "columns (not a slice of the quarter) - actual day-count per month used for Debtor/Creditor "
        "Days, not annualised. RATIO SUMMARY below links to RATIO WORKINGS further down this sheet - "
        "every input figure behind each ratio is shown there for cross-check. Asset Turnover / "
        "Debtor Days / Creditor Days use the AVERAGE of opening and closing balance (May/Jun use the "
        "prior month's own closing as opening; Apr-26 falls back to closing-only - no 31-Mar-2026 "
        "opening BS fetched yet). See 'Ratio Analysis' for the Q1 quarterly view of the same ratios."
    )).font = b.NORMAL_FONT
    row += 2

    header_row1, header_row2 = row, row + 1
    ws.cell(header_row1, 1, value="Line Item").font = b.HEADER_FONT
    ws.cell(header_row1, 1).fill = b.HEADER_FILL
    ws.merge_cells(start_row=header_row1, start_column=1, end_row=header_row2, end_column=1)
    for pidx, (pkey, plabel) in enumerate(periods):
        cols = ratio_period_cols(pidx)
        start, end = cols["CO1"], cols["GRANDTOTAL"]
        ws.merge_cells(start_row=header_row1, start_column=start, end_row=header_row1, end_column=end)
        cell = ws.cell(header_row1, start, value=plabel)
        cell.font = b.HEADER_FONT
        cell.fill = b.TOTALCOL_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
        sub_labels = {cols["CO1"]: b.CO1[0], cols["CO2"]: b.CO2[0], cols["TOTAL_ORANGE"]: "Total of Orange O Tec",
                      cols["CO3"]: b.CO3[0], cols["CO4"]: b.CO4[0], cols["TOTAL_ENTERPRISES"]: "Total of Enterprises",
                      cols["CO5"]: b.CO5[0], cols["GRANDTOTAL"]: "Grand Total"}
        for c, lbl in sub_labels.items():
            cell = ws.cell(header_row2, c, value=lbl)
            cell.font = b.HEADER_FONT
            cell.fill = (b.TOTALCOL_FILL if c in (cols["TOTAL_ORANGE"], cols["TOTAL_ENTERPRISES"], cols["GRANDTOTAL"])
                         else b.HEADER_FILL)
            cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="bottom")
    ws.row_dimensions[header_row1].height = 20
    ws.row_dimensions[header_row2].height = 40
    ws.freeze_panes = ws.cell(header_row2 + 1, ratio_period_start_col(0)).coordinate
    row = header_row2 + 1

    def pl(key, col_letter):
        return f"'MOM PL GL Detail'!{col_letter}{pl_row_map[key]}"

    def bs(key, col_letter):
        return f"'MOM BS GL Detail'!{col_letter}{bs_row_map[key]}"

    categories = b.ratio_definitions()
    layout = b.ratio_sheet_layout(categories, row)

    # For each period block: (key, dest_col-on-this-sheet, closing_col-on-MOM-schedule,
    # opening_col-on-MOM-schedule, day_count). dest_col comes from THIS sheet's own layout
    # (ratio_period_cols); closing_col/opening_col from the MOM PL/BS GL Detail schedule's layout
    # (period_cols) - the two are different column-numbering schemes (see RATIO_LABEL_COLS vs
    # LABEL_COLS above), the exact mix-up that caused the wrong-column bug this sheet had earlier -
    # keep them looked up separately, never reuse one dict for both purposes. opening_col is the
    # PRIOR period's closing column (May's opening = Apr's closing, June's opening = May's closing)
    # - for April itself there is no prior MOM period, so it falls back to its own closing column
    # (see AVG_COMPONENTS docstring in build_gl_detail_sidebyside.py for what this fallback means).
    summary_columns = []
    workings_columns = []
    for pidx, (pkey, _) in enumerate(periods):
        dest_cols = ratio_period_cols(pidx)
        sched_cols = period_cols(pidx)
        opening_sched_cols = period_cols(pidx - 1) if pidx > 0 else sched_cols
        for key in ("CO1", "CO2", "TOTAL_ORANGE", "CO3", "CO4", "TOTAL_ENTERPRISES", "CO5", "GRANDTOTAL"):
            summary_columns.append((key, dest_cols[key]))
            workings_columns.append((key, dest_cols[key], get_column_letter(sched_cols[key]),
                                      get_column_letter(opening_sched_cols[key]), MONTH_DAYS[pkey]))

    b.write_ratio_summary_section(ws, layout, categories, summary_columns, label_col=1)
    b.write_ratio_workings_section(ws, layout, categories, workings_columns, pl, bs, label_col=1)


# ---------------------------------------------------------------------------
# "Cross Check": live formulas (no hardcoded numbers) that verify the MOM sheets tie back to the
# quarter sheets, at both consolidated and by-company level. Two distinct checks per line item:
#   1. Var: Sum(Apr:Jun) vs MOM Q1 Total - internal consistency of "MOM PL and BS" itself. This is
#      EXPECTED to be non-zero until gl_data_mom_detail/ (the monthly ledger CSVs) is refreshed
#      with live Tally data - it is a data-vintage gap, not a formula bug (Q1 Total reuses the
#      live gl_data/ CSVs directly, the monthly columns do not).
#   2. Var: MOM Q1 Total vs PL/BS Summary Grand Total (or company column) - a true formula/data
#      integrity check. Both pull from the same live gl_data/ CSVs, so this should be ~0; any
#      non-zero value here signals a real bug, not a data-vintage difference.
CHECK_FILL_ALERT = PatternFill("solid", fgColor="FFFF00")
CHECK_FONT_ALERT = Font(bold=True, color="9C0006")
CHECK_TOLERANCE = 10  # rupees - absorbs floating-point noise, not real gaps

PL_CHECK_LINES = [
    "Total Turnover", "Total Cost of Sales", "Net Branch Transactions", "GROSS PROFIT",
    "Total Other Income", "Total Operating Expenses", "EBITDA", "Total Finance Cost",
    "PBT (Profit Before Tax)", "PAT (Net Profit)", "Net Related Party Transactions",
    "PAT (Net of Related Party Transactions)",
]
BS_CHECK_LINES = ["Total Liabilities", "Total Assets", "Reconciliation Gap (~0 expected)"]

SUM2_COL_BY_FIELD = {
    "CO1": b.SUM2_COL_CO1, "CO2": b.SUM2_COL_CO2, "TOTAL_ORANGE": b.SUM2_COL_TOTAL_ORANGE,
    "CO3": b.SUM2_COL_CO3, "CO4": b.SUM2_COL_CO4, "TOTAL_ENTERPRISES": b.SUM2_COL_TOTAL_ENTERPRISES,
    "CO5": b.SUM2_COL_CO5, "GRANDTOTAL": b.SUM2_COL_GRANDTOTAL,
}


def find_row(ws, label, col=1):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, col).value == label:
            return r
    raise ValueError(f"Row not found for label {label!r} in sheet {ws.title!r}")


def write_cross_check_header(ws, row, title, col_labels):
    ws.cell(row, 1, value=title).font = b.TITLE_FONT
    row += 2
    ws.cell(row, 1, value="Line Item").font = b.HEADER_FONT
    for i, lbl in enumerate(col_labels):
        cell = ws.cell(row, 2 + i, value=lbl)
        cell.font = b.HEADER_FONT
        cell.fill = b.HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="bottom")
    ws.cell(row, 1).fill = b.HEADER_FILL
    ws.row_dimensions[row].height = 32
    return row + 1


def write_pl_consolidated_checks(ws, row, mom_ws_name, mom_row_map, pls_ws_name, pls_row_map):
    row = write_cross_check_header(
        ws, row, "P&L - Consolidated Cross-Check (Grand Total)",
        ["Apr-26", "May-26", "Jun-26", "Sum (Apr:Jun)", "MOM Q1 Total", "PL Summary (Grand Total)",
         "Var: Sum vs MOM Q1 Total", "Var: MOM Q1 Total vs PL Summary",
         "Jul-26", "Sum (Apr:Jul)", "MOM Apr-Jul Total", "Var: Sum(Apr:Jul) vs Apr-Jul Total"])
    var_cells = []
    for label in PL_CHECK_LINES:
        mrow = mom_row_map[label]
        prow = pls_row_map[label]
        ws.cell(row, 1, value=label).font = b.NORMAL_FONT
        ws.cell(row, 2, value=f"='{mom_ws_name}'!C{mrow}").number_format = b.NUM_FORMAT
        ws.cell(row, 3, value=f"='{mom_ws_name}'!D{mrow}").number_format = b.NUM_FORMAT
        ws.cell(row, 4, value=f"='{mom_ws_name}'!E{mrow}").number_format = b.NUM_FORMAT
        ws.cell(row, 5, value=f"=B{row}+C{row}+D{row}").number_format = b.NUM_FORMAT
        ws.cell(row, 6, value=f"='{mom_ws_name}'!F{mrow}").number_format = b.NUM_FORMAT
        gt_letter = get_column_letter(b.SUM2_COL_GRANDTOTAL)
        ws.cell(row, 7, value=f"='{pls_ws_name}'!{gt_letter}{prow}").number_format = b.NUM_FORMAT
        ws.cell(row, 8, value=f"=E{row}-F{row}").number_format = b.NUM_FORMAT
        ws.cell(row, 9, value=f"=F{row}-G{row}").number_format = b.NUM_FORMAT
        # Apr-Jul leg (2026-08-04): the same additivity test one period longer. 'ytd' is fetched
        # in its own right rather than derived from Q1+Jul, so this is a genuine reconciliation.
        ws.cell(row, 10, value=f"='{mom_ws_name}'!G{mrow}").number_format = b.NUM_FORMAT
        ws.cell(row, 11, value=f"=B{row}+C{row}+D{row}+J{row}").number_format = b.NUM_FORMAT
        ws.cell(row, 12, value=f"='{mom_ws_name}'!H{mrow}").number_format = b.NUM_FORMAT
        ws.cell(row, 13, value=f"=K{row}-L{row}").number_format = b.NUM_FORMAT
        var_cells.append(row)
        row += 1
    for r in var_cells:
        for c in (8, 9, 13):
            ws.conditional_formatting.add(
                ws.cell(r, c).coordinate,
                FormulaRule(formula=[f"ABS({get_column_letter(c)}{r})>{CHECK_TOLERANCE}"],
                            fill=CHECK_FILL_ALERT, font=CHECK_FONT_ALERT))
    return row + 1


def write_pl_by_company_checks(ws, row, mom_ws_name, mom_by_co_row_map, periods, pls_ws_name, pls_row_map):
    for bidx, (bkey, blabel, field) in enumerate(BLOCKS):
        row = write_cross_check_header(
            ws, row, f"P&L - By Company Cross-Check: {blabel}",
            ["Apr-26", "May-26", "Jun-26", "Sum (Apr:Jun)", "MOM Q1 Total", f"PL Summary ({blabel})",
             "Var: Sum vs MOM Q1 Total", "Var: MOM Q1 Total vs PL Summary"])
        var_cells = []
        block_start = by_company_block_start(bidx, periods)
        for label in PL_CHECK_LINES:
            mrow = mom_by_co_row_map[label]
            prow = pls_row_map[label]
            ws.cell(row, 1, value=label).font = b.NORMAL_FONT
            for i in range(3):
                col_letter = get_column_letter(block_start + i)
                ws.cell(row, 2 + i, value=f"='{mom_ws_name}'!{col_letter}{mrow}").number_format = b.NUM_FORMAT
            ws.cell(row, 5, value=f"=B{row}+C{row}+D{row}").number_format = b.NUM_FORMAT
            q1_col_letter = get_column_letter(block_start + 3)
            ws.cell(row, 6, value=f"='{mom_ws_name}'!{q1_col_letter}{mrow}").number_format = b.NUM_FORMAT
            pls_col_letter = get_column_letter(SUM2_COL_BY_FIELD[field])
            ws.cell(row, 7, value=f"='{pls_ws_name}'!{pls_col_letter}{prow}").number_format = b.NUM_FORMAT
            ws.cell(row, 8, value=f"=E{row}-F{row}").number_format = b.NUM_FORMAT
            ws.cell(row, 9, value=f"=F{row}-G{row}").number_format = b.NUM_FORMAT
            var_cells.append(row)
            row += 1
        for r in var_cells:
            for c in (8, 9):
                ws.conditional_formatting.add(
                    ws.cell(r, c).coordinate,
                    FormulaRule(formula=[f"ABS({get_column_letter(c)}{r})>{CHECK_TOLERANCE}"],
                                fill=CHECK_FILL_ALERT, font=CHECK_FONT_ALERT))
        row += 1
    return row + 1


def write_bs_consolidated_checks(ws, row, mom_ws_name, mom_row_map, bss_ws_name, bss_row_map):
    row = write_cross_check_header(
        ws, row, "Balance Sheet - Consolidated Cross-Check (point-in-time: Jun-26 close vs BS Summary only)",
        ["Apr-26", "May-26", "Jun-26", "BS Summary (Grand Total)", "Var: Jun-26 (MOM) vs BS Summary"])
    var_cells = []
    for label in BS_CHECK_LINES:
        mrow = mom_row_map[label]
        brow = bss_row_map[label]
        ws.cell(row, 1, value=label).font = b.NORMAL_FONT
        ws.cell(row, 2, value=f"='{mom_ws_name}'!C{mrow}").number_format = b.NUM_FORMAT
        ws.cell(row, 3, value=f"='{mom_ws_name}'!D{mrow}").number_format = b.NUM_FORMAT
        ws.cell(row, 4, value=f"='{mom_ws_name}'!E{mrow}").number_format = b.NUM_FORMAT
        gt_letter = get_column_letter(b.SUM2_COL_GRANDTOTAL)
        ws.cell(row, 5, value=f"='{bss_ws_name}'!{gt_letter}{brow}").number_format = b.NUM_FORMAT
        ws.cell(row, 6, value=f"=D{row}-E{row}").number_format = b.NUM_FORMAT
        var_cells.append(row)
        row += 1
    for r in var_cells:
        ws.conditional_formatting.add(
            ws.cell(r, 6).coordinate,
            FormulaRule(formula=[f"ABS(F{r})>{CHECK_TOLERANCE}"], fill=CHECK_FILL_ALERT, font=CHECK_FONT_ALERT))
    return row + 1


def write_bs_by_company_checks(ws, row, mom_ws_name, mom_by_co_row_map, periods, bss_ws_name, bss_row_map):
    for bidx, (bkey, blabel, field) in enumerate(BLOCKS):
        row = write_cross_check_header(
            ws, row, f"Balance Sheet - By Company Cross-Check: {blabel} (point-in-time)",
            ["Apr-26", "May-26", "Jun-26", f"BS Summary ({blabel})", "Var: Jun-26 (MOM) vs BS Summary"])
        var_cells = []
        block_start = by_company_block_start(bidx, periods)
        for label in BS_CHECK_LINES:
            mrow = mom_by_co_row_map[label]
            brow = bss_row_map[label]
            ws.cell(row, 1, value=label).font = b.NORMAL_FONT
            for i in range(3):
                col_letter = get_column_letter(block_start + i)
                ws.cell(row, 2 + i, value=f"='{mom_ws_name}'!{col_letter}{mrow}").number_format = b.NUM_FORMAT
            bss_col_letter = get_column_letter(SUM2_COL_BY_FIELD[field])
            ws.cell(row, 5, value=f"='{bss_ws_name}'!{bss_col_letter}{brow}").number_format = b.NUM_FORMAT
            ws.cell(row, 6, value=f"=D{row}-E{row}").number_format = b.NUM_FORMAT
            var_cells.append(row)
            row += 1
        for r in var_cells:
            ws.conditional_formatting.add(
                ws.cell(r, 6).coordinate,
                FormulaRule(formula=[f"ABS(F{r})>{CHECK_TOLERANCE}"], fill=CHECK_FILL_ALERT, font=CHECK_FONT_ALERT))
        row += 1
    return row + 1



# ---------------------------------------------------------------------------
# "Related Party" sheet (2026-08-04 onward, user request)
# ---------------------------------------------------------------------------
# A standalone reconciliation of every inter-company item, built so the NET of each block can be
# read at a glance - the whole point being that a genuine group elimination nets to ZERO.
#
# What it showed on the day it was built: the P&L side eliminates cleanly (Rs 9.97cr of related
# party sales against the same purchases, netting to Rs 0.80 across Apr-Jul), but the BALANCE
# SHEET side does not - roughly Rs 42 lakh of related-party balances and Rs 8.4 lakh of branch
# balances were unreconciled at both 30-Jun and 31-Jul. The counterparty matrix exists to locate
# exactly which company pair disagrees, which a flat ledger listing cannot do.
#
# Detail cells carry values; every NET/total row is a live formula over the cells above it, so the
# reconciliation can be audited in Excel rather than taken on trust.
RP_TOL = 10.0


def _rp_period_rows(stat_type, key, period):
    """(primary_group, sub_group, ledger, balance, counterparty) for one company/period, limited
    to rows that are inter-company by either test in classify_bs_memo / is_related_party_row."""
    out = []
    for pg, sg, gl, bal in load_period_rows(stat_type, key, period):
        if stat_type == "pl":
            if not b.is_related_party_row(pg, sg):
                continue
            kind = "RELATED"
        else:
            kind = b.classify_bs_memo(pg, sg, gl)
            if kind is None:
                continue
        out.append((pg, sg, gl, bal, b.counterparty_of(gl), kind))
    return out


def _rp_header(ws, row, title, period_labels, first_col_width=46):
    ws.column_dimensions["A"].width = first_col_width
    ws.cell(row, 1, value=title).font = b.SECTION_FONT
    for c in range(1, 2 + len(period_labels)):
        ws.cell(row, c).fill = b.SECTION_FILL
    row += 1
    ws.cell(row, 1, value="Line Item").font = b.HEADER_FONT
    ws.cell(row, 1).fill = b.HEADER_FILL
    for i, lbl in enumerate(period_labels):
        c = ws.cell(row, 2 + i, value=lbl)
        c.font = b.HEADER_FONT
        c.fill = b.TOTALCOL_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(2 + i)].width = 18
    return row + 1


def _rp_net_row(ws, row, label, value_rows, ncols):
    """NET row as a live SUM over the detail rows above - highlighted when it breaches RP_TOL."""
    ws.cell(row, 1, value=label).font = b.GRANDTOTAL_FONT
    for i in range(ncols):
        col = get_column_letter(2 + i)
        cell = ws.cell(row, 2 + i,
                       value="=" + "+".join(f"{col}{r}" for r in value_rows) if value_rows else "=0")
        cell.number_format = b.NUM_FORMAT
        cell.font = b.GRANDTOTAL_FONT
        cell.fill = b.GRANDTOTAL_FILL
        ws.conditional_formatting.add(
            cell.coordinate,
            FormulaRule(formula=[f"ABS({col}{row})>{RP_TOL}"], fill=CHECK_FILL_ALERT,
                        font=CHECK_FONT_ALERT))
    return row + 1


def write_related_party_sheet(wb):
    ws = wb.create_sheet(title="Related Party")
    pl_periods = PL_PERIODS
    bs_periods = BS_PERIODS
    names = {k: n for n, k in b.ALL_COMPANIES}

    row = 1
    ws.cell(row, 1, value="Related Party & Inter-Company Reconciliation").font = b.TITLE_FONT
    row += 1
    ws.cell(row, 1, value=(
        "Every inter-company item in one place. A genuine group elimination NETS TO ZERO, so each "
        "NET row below is the check - amber means it breaches Rs.10 and something is unmatched. "
        "P&L blocks are FLOWS (per period); Balance Sheet blocks are POSITIONS (as at each "
        "month-end, never summed across months). Related-party ledgers are identified both by "
        "sub-group name containing 'RELATED' and by the ledger name carrying a group company's "
        "name - about a third sit in ordinary groups such as Unsecured Loans and are only caught "
        "by the second test. Branch is kept separate because the intra-group Surat/Noida pair is "
        "~Rs 58cr each way and would otherwise swamp the related-party totals."
    )).font = b.NORMAL_FONT
    ws.row_dimensions[row].height = 60
    row += 2

    # ---- 1. P&L flows ------------------------------------------------------------------------
    pl_labels = [lbl for _, lbl in pl_periods]
    row = _rp_header(ws, row, "1. PROFIT & LOSS - RELATED PARTY TRANSACTIONS (flows)", pl_labels)
    detail_rows = []
    for bucket, want_sales in (("Related Party Sales", True), ("Related Party Purchase", False)):
        ws.cell(row, 1, value=bucket).font = b.NORMAL_FONT
        for i, (pkey, _) in enumerate(pl_periods):
            tot = 0.0
            for _, key in b.ALL_COMPANIES:
                for pg, sg, gl, bal, cp, kind in _rp_period_rows("pl", key, pkey):
                    is_sales = pg == "Sales Accounts"
                    if is_sales == want_sales:
                        tot += bal
            c = ws.cell(row, 2 + i, value=tot)
            c.number_format = b.NUM_FORMAT
            c.font = b.NORMAL_FONT
        detail_rows.append(row)
        row += 1
    row = _rp_net_row(ws, row, "NET RELATED PARTY P&L (should be 0)", detail_rows, len(pl_periods))
    row += 1

    # ---- 2 & 3. Balance Sheet positions, related party then branch ----------------------------
    bs_labels = [lbl for _, lbl in bs_periods]
    for idx, (kind_wanted, title) in enumerate((
        ("RELATED", "2. BALANCE SHEET - RELATED PARTY BALANCES (as at month-end)"),
        ("BRANCH", "3. BALANCE SHEET - BRANCH BALANCES (as at month-end)"),
    )):
        row = _rp_header(ws, row, title, bs_labels)
        comp_rows = []
        for _, key in b.ALL_COMPANIES:
            ws.cell(row, 1, value=names[key]).font = b.NORMAL_FONT
            for i, (pkey, _) in enumerate(bs_periods):
                tot = sum(bal for _, _, _, bal, _, kind in _rp_period_rows("bs", key, pkey)
                          if kind == kind_wanted)
                c = ws.cell(row, 2 + i, value=tot)
                c.number_format = b.NUM_FORMAT
                c.font = b.NORMAL_FONT
            comp_rows.append(row)
            row += 1
        row = _rp_net_row(ws, row, f"NET {kind_wanted.title()} BALANCES (should be 0)",
                          comp_rows, len(bs_periods))
        row += 1

    # ---- 4. Counterparty matrix, per Balance Sheet period -------------------------------------
    ws.cell(row, 1, value="4. COUNTERPARTY MATRIX - related party only, who books against whom").font = b.SECTION_FONT
    for c in range(1, 2 + len(b.ALL_COMPANIES)):
        ws.cell(row, c).fill = b.SECTION_FILL
    row += 1
    ws.cell(row, 1, value=(
        "Rows = the company whose books the balance sits in; columns = the counterparty it is "
        "booked against. For any pair, the two cells (A against B, and B against A) are mirror "
        "images and must sum to zero - the PAIR CHECK block beneath each matrix does exactly that, "
        "so a non-zero cell names the two companies whose records disagree."
    )).font = b.NORMAL_FONT
    ws.row_dimensions[row].height = 32
    row += 2

    for pkey, plabel in bs_periods:
        mat = {a: {c: 0.0 for c, _ in [(k, 0) for _, k in b.ALL_COMPANIES]} for _, a in b.ALL_COMPANIES}
        for _, key in b.ALL_COMPANIES:
            for pg, sg, gl, bal, cp, kind in _rp_period_rows("bs", key, pkey):
                # RELATED only. Branch is deliberately excluded: Surat and Noida are the same
                # legal entity in two Tally files, and a branch ledger named after its OWN
                # company (e.g. "ORANGE O TEC PVT. LTD.(SURAT BRANCH)" sitting in the Surat
                # books) maps to itself, so one leg of the pair would be dropped and the pair
                # check would report a spurious multi-crore break. Branch still gets its own NET
                # block in section 3, which is the meaningful test for it.
                if kind != "RELATED" or cp is None or cp == key:
                    continue
                mat[key][cp] += bal

        ws.cell(row, 1, value=f"{plabel} - counterparty matrix").font = b.SUBTOTAL_FONT
        row += 1
        ws.cell(row, 1, value="Books of \\ Counterparty").font = b.HEADER_FONT
        ws.cell(row, 1).fill = b.HEADER_FILL
        for j, (_, ck) in enumerate(b.ALL_COMPANIES):
            c = ws.cell(row, 2 + j, value=names[ck])
            c.font = b.HEADER_FONT
            c.fill = b.TOTALCOL_FILL
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        row += 1
        first_matrix_row = row
        row_of = {}
        for _, ak in b.ALL_COMPANIES:
            ws.cell(row, 1, value=names[ak]).font = b.NORMAL_FONT
            for j, (_, ck) in enumerate(b.ALL_COMPANIES):
                v = mat[ak][ck]
                cell = ws.cell(row, 2 + j, value=v if ak != ck else None)
                cell.number_format = b.NUM_FORMAT
                cell.font = b.NORMAL_FONT
            row_of[ak] = row
            row += 1
        row += 1

        ws.cell(row, 1, value=f"{plabel} - PAIR CHECK (A vs B + B vs A; should be 0)").font = b.SUBTOTAL_FONT
        row += 1
        keys = [k for _, k in b.ALL_COMPANIES]
        for i, ak in enumerate(keys):
            for ck in keys[i + 1:]:
                a_col = get_column_letter(2 + keys.index(ck))
                c_col = get_column_letter(2 + keys.index(ak))
                ws.cell(row, 1, value=f"{names[ak]}  vs  {names[ck]}").font = b.NORMAL_FONT
                cell = ws.cell(row, 2,
                               value=f"={a_col}{row_of[ak]}+{c_col}{row_of[ck]}")
                cell.number_format = b.NUM_FORMAT
                cell.font = b.NORMAL_FONT
                ws.conditional_formatting.add(
                    cell.coordinate,
                    FormulaRule(formula=[f"ABS(B{row})>{RP_TOL}"], fill=CHECK_FILL_ALERT,
                                font=CHECK_FONT_ALERT))
                row += 1
        row += 1

    # ---- 5. Ledger detail ---------------------------------------------------------------------
    ws.cell(row, 1, value="5. LEDGER DETAIL - every inter-company ledger, by month-end").font = b.SECTION_FONT
    for c in range(1, 6 + len(bs_periods)):
        ws.cell(row, c).fill = b.SECTION_FILL
    row += 1
    heads = ["Company (books)", "Type", "Grouping", "Sub Group", "Ledger", "Counterparty"] + bs_labels
    for j, h in enumerate(heads):
        c = ws.cell(row, 1 + j, value=h)
        c.font = b.HEADER_FONT
        c.fill = b.HEADER_FILL if j < 6 else b.TOTALCOL_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for j, w in enumerate((30, 10, 24, 28, 46, 26)):
        ws.column_dimensions[get_column_letter(1 + j)].width = w
    for j in range(len(bs_periods)):
        ws.column_dimensions[get_column_letter(7 + j)].width = 18
    row += 1

    index = {}
    for _, key in b.ALL_COMPANIES:
        for i, (pkey, _) in enumerate(bs_periods):
            for pg, sg, gl, bal, cp, kind in _rp_period_rows("bs", key, pkey):
                index.setdefault((key, kind, pg, sg, gl, cp), {})[pkey] = bal
    for (key, kind, pg, sg, gl, cp), vals in sorted(
            index.items(), key=lambda kv: (kv[0][1], kv[0][0], kv[0][2], kv[0][4])):
        ws.cell(row, 1, value=names[key]).font = b.NORMAL_FONT
        ws.cell(row, 2, value=kind.title()).font = b.NORMAL_FONT
        ws.cell(row, 3, value=pg).font = b.NORMAL_FONT
        ws.cell(row, 4, value=sg).font = b.NORMAL_FONT
        ws.cell(row, 5, value=gl).font = b.NORMAL_FONT
        ws.cell(row, 6, value=names.get(cp, "-") if cp else "(unmapped)").font = b.NORMAL_FONT
        for i, (pkey, _) in enumerate(bs_periods):
            c = ws.cell(row, 7 + i, value=vals.get(pkey, 0.0))
            c.number_format = b.NUM_FORMAT
            c.font = b.NORMAL_FONT
        row += 1

    ws.freeze_panes = "B5"
    return ws

def write_cross_check_sheet(wb):
    ws = wb.create_sheet(title="Cross Check")
    ws.column_dimensions["A"].width = 42
    for col_letter in ("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"):
        ws.column_dimensions[col_letter].width = 19

    pl_summary_ws = wb["PL Summary"]
    bs_summary_ws = wb["BS Summary"]
    mom_ws = wb["MOM PL and BS"]
    mom_by_co_ws = wb["MOM PL and BS - By Company"]

    pls_row_map = {label: find_row(pl_summary_ws, label) for label in PL_CHECK_LINES}
    bss_row_map = {label: find_row(bs_summary_ws, label) for label in BS_CHECK_LINES}
    mom_pl_row_map = {label: find_row(mom_ws, label) for label in PL_CHECK_LINES}
    mom_bs_row_map = {label: find_row(mom_ws, label) for label in BS_CHECK_LINES}
    mom_by_co_pl_row_map = {label: find_row(mom_by_co_ws, label) for label in PL_CHECK_LINES}
    mom_by_co_bs_row_map = {label: find_row(mom_by_co_ws, label) for label in BS_CHECK_LINES}

    row = 1
    ws.cell(row, 1, value="Cross-Check: MOM sheets vs Quarter (Q1) sheets").font = b.TITLE_FONT
    row += 1
    ws.cell(row, 1, value=(
        "Two distinct checks per line. (1) Var: Sum(Apr:Jun) vs MOM Q1 Total - internal MOM "
        "consistency; EXPECTED to be non-zero until gl_data_mom_detail/ (the monthly ledger CSVs) "
        "is refreshed with live Tally data - a data-vintage gap, not a formula bug, since Q1 Total "
        "already reuses the live quarter CSVs directly. (2) Var: MOM Q1 Total vs PL/BS Summary - "
        "both pull from the identical live quarter data, so this should be ~0; any non-zero value "
        "here is a real formula/data integrity issue. Cells beyond Rs.10 are highlighted yellow."
    )).font = b.NORMAL_FONT
    ws.row_dimensions[row].height = 45
    row += 2

    row = write_pl_consolidated_checks(ws, row, "MOM PL and BS", mom_pl_row_map, "PL Summary", pls_row_map)
    row = write_pl_by_company_checks(ws, row, "MOM PL and BS - By Company", mom_by_co_pl_row_map,
                                      PL_PERIODS, "PL Summary", pls_row_map)
    row = write_bs_consolidated_checks(ws, row, "MOM PL and BS", mom_bs_row_map, "BS Summary", bss_row_map)
    write_bs_by_company_checks(ws, row, "MOM PL and BS - By Company", mom_by_co_bs_row_map,
                                BS_PERIODS, "BS Summary", bss_row_map)

    ws.freeze_panes = "B5"


# ---------------------------------------------------------------------------
# NEW (2026-07-25): PL Main Group / Sub Group / Full GL MOM sheets, and one sheet per real
# BS Group - all fully formula-linked back to MOM PL GL Detail / MOM BS GL Detail (never a
# literal), per explicit user request ("do not merge cells", "keep month as each column",
# "the figures in the cell are not linked - please keep all the formulae", "show Branch and
# Related Party total also"). Column layout per row: Apr-26, May-26, Jun-26, then a QUARTER
# TOTAL block - Q1 TOTAL for P&L (a same-sheet SUM of the three month blocks, since MOM PL GL
# Detail itself has no Q1 Total column), "AS AT 30-JUN-26" for BS (a same-sheet COPY of the Jun
# block, since Balance Sheet is point-in-time, not additive - summing three month-end snapshots
# would be meaningless). Jul-26 (and eventually a Q2 TOTAL block) append the same way in future
# refreshes - never a rebuild-from-scratch of these sheets.
# ---------------------------------------------------------------------------

NEW_LABEL_FONT = Font(size=9, name="Calibri")
NEW_HEADER_FONT = Font(bold=True, size=9, name="Calibri", color="FFFFFF")
NEW_HEADER_FILL = PatternFill("solid", fgColor="2E5395")
NEW_PERIOD_FONT = Font(bold=True, size=9, name="Calibri", color="FFFFFF")
NEW_PERIOD_FILL = PatternFill("solid", fgColor="4472C4")
NEW_QTR_FILL = PatternFill("solid", fgColor="833C0C")
NEW_SECTION_FONT = Font(bold=True, size=10, name="Calibri", color="FFFFFF")
NEW_SECTION_FILL = PatternFill("solid", fgColor="1F3864")
NEW_MEMO_FONT = Font(italic=True, bold=True, size=10, name="Calibri", color="FFFFFF")
NEW_MEMO_FILL = PatternFill("solid", fgColor="7F7F7F")
NEW_MILESTONE_FONT = Font(bold=True, size=9, name="Calibri", color="FFFFFF")
NEW_MILESTONE_FILL = PatternFill("solid", fgColor="1F3864")
NEW_TOTAL_FONT = Font(bold=True, size=9, name="Calibri")
NEW_COMPANY_LABELS = ["Orange O Tec Pvt Ltd", "Orange O Tec NOIDA", "Total Orange O Tec",
                      "Enterprises Pvt Ltd", "Enterprises NOIDA", "Total Enterprises",
                      "Colorix", "Grand Total"]
NEW_OUTLINE_LEVELS = [2, 2, 1, 2, 2, 1, 1, 0]  # 2=innermost detail, 1=family total, 0=Grand Total


def new_block_start_col(n_label_cols, block_idx):
    return n_label_cols + 1 + block_idx * 8


def write_new_sheet_headers(ws, n_label_cols, label_names, header_row, blocks):
    period_row = header_row - 1
    for i, name in enumerate(label_names, start=1):
        c = ws.cell(header_row, i, value=name)
        c.font = NEW_HEADER_FONT
        c.fill = NEW_HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for bidx, (_key, plabel, kind) in enumerate(blocks):
        start = new_block_start_col(n_label_cols, bidx)
        pc = ws.cell(period_row, start, value=plabel)
        pc.font = NEW_PERIOD_FONT
        pc.fill = NEW_QTR_FILL if kind == "quarter" else NEW_PERIOD_FILL
        for i, clabel in enumerate(NEW_COMPANY_LABELS):
            col = start + i
            c = ws.cell(header_row, col, value=clabel)
            c.font = NEW_HEADER_FONT
            c.fill = NEW_HEADER_FILL
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(col)].outline_level = NEW_OUTLINE_LEVELS[i]
    ws.row_dimensions[period_row].height = 16
    ws.row_dimensions[header_row].height = 26
    ws.sheet_properties.outlinePr.summaryRight = True


def write_new_linked_row(ws, target_row, n_label_cols, labels, schedule_name, source_row, blocks,
                          font=None, fill=None, quarter_mode="direct"):
    """Writes one row of column-blocks. schedule_name/source_row is the ONE row in MOM PL/BS GL
    Detail this whole row is linked to (that sheet is one row per bucket/sub-group/ledger, with a
    separate column-block per period in that same row - not a separate row per month).

    quarter_mode="direct" (P&L): every block, including the last (Q1 TOTAL), maps 1:1 to a
    column-block that genuinely exists in schedule_name (MOM PL GL Detail already carries its own
    Q1 Total block, sourced from the quarter CSV directly - not summed from Apr/May/Jun - so
    linking to it directly is both simpler AND more correct than re-summing here, which would
    reproduce the small, already-documented Apr+May+Jun-vs-quarter timing gap the Cross Check
    sheet tracks). quarter_mode="asof" (BS): MOM BS GL Detail has no such block (Balance Sheet is
    point-in-time - summing month-end snapshots is meaningless), so the last block is a same-sheet
    formula copying the previous (Jun) block instead of a cross-sheet reference.

    Total Orange/Enterprises/Grand Total are always same-row, same-sheet formulas from THIS
    sheet's own CO1..CO5 - never copied from the source - matching this project's established
    convention (see write_period_cells)."""
    for i, lbl in enumerate(labels, start=1):
        c = ws.cell(target_row, i, value=lbl)
        c.font = font or NEW_LABEL_FONT
        if fill:
            c.fill = fill
    n_direct = len(blocks) - 1 if quarter_mode == "asof" else len(blocks)
    for bidx in range(len(blocks)):
        dest_start = new_block_start_col(n_label_cols, bidx)
        if bidx < n_direct:
            src_cols = period_cols(bidx)
            for i, key in enumerate(("CO1", "CO2", "CO3", "CO4", "CO5")):
                src_letter = get_column_letter(src_cols[key])
                formula = f"='{schedule_name}'!{src_letter}{source_row}"
                cell = ws.cell(target_row, dest_start + i, value=formula)
                cell.number_format = b.NUM_FORMAT
                cell.font = font or Font(size=9, name="Calibri")
            co1 = get_column_letter(dest_start)
            co2 = get_column_letter(dest_start + 1)
            co3 = get_column_letter(dest_start + 3)
            co4 = get_column_letter(dest_start + 4)
            co5 = get_column_letter(dest_start + 6)
            to = get_column_letter(dest_start + 2)
            te = get_column_letter(dest_start + 5)
            ws.cell(target_row, dest_start + 2, value=f"={co1}{target_row}+{co2}{target_row}")
            ws.cell(target_row, dest_start + 5, value=f"={co3}{target_row}+{co4}{target_row}")
            ws.cell(target_row, dest_start + 7, value=f"={to}{target_row}+{te}{target_row}+{co5}{target_row}")
            for off in (2, 5, 7):
                c = ws.cell(target_row, dest_start + off)
                c.number_format = b.NUM_FORMAT
                c.font = font or Font(size=9, name="Calibri")
        else:
            # "asof" trailing block: same-sheet copy of the immediately preceding (Jun) block
            prev_start = new_block_start_col(n_label_cols, bidx - 1)
            for i in range(8):
                src_letter = get_column_letter(prev_start + i)
                cell = ws.cell(target_row, dest_start + i, value=f"={src_letter}{target_row}")
                cell.number_format = b.NUM_FORMAT
                cell.font = font or Font(size=9, name="Calibri")
    for c in range(1, new_block_start_col(n_label_cols, len(blocks))):
        ws.cell(target_row, c).border = BORDER_THIN


thin_side = Side(style="thin", color="D9D9D9")
BORDER_THIN = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)


def write_new_section_header(ws, row, n_label_cols, n_blocks, text, font=NEW_SECTION_FONT, fill=NEW_SECTION_FILL):
    ncols = n_label_cols + 8 * n_blocks
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.fill = fill
        if c == 1:
            cell.value = text
            cell.font = font
    return row + 1


PL_MOM_SECTION_ORDER = ["TURNOVER", "OTHER INCOME", "COST OF SALES", "OPERATING EXPENSES"]
PL_MOM_SECTION_DISPLAY = {"TURNOVER": "TURNOVER", "OTHER INCOME": "OTHER INCOME",
                          "COST OF SALES": "COST OF SALES", "OPERATING EXPENSES": "INDIRECT COST"}


def build_pl_mom_extended_sheet(ws, title, granularity, schedule_name, pl_row_map, pl_section_buckets,
                                 pl_section_detail, pl_leaf_detail, periods):
    # `periods` is PL_PERIODS as-is (apr, may, jun, qtr) - MOM PL GL Detail already carries its
    # own Q1 Total block (4th), sourced from the quarter CSV directly, so every block here maps
    # 1:1 onto a real column-block in that sheet (quarter_mode="direct", the default).
    blocks = [(k, l, "quarter" if k == "qtr" else "month") for k, l in periods]
    ws.sheet_view.showGridLines = False
    n_label_cols = {"main": 2, "sub": 3, "gl": 4}[granularity]
    label_names = {"main": ["Section", "Grouping"],
                   "sub": ["Section", "Grouping", "Sub Grouping"],
                   "gl": ["Section", "Grouping", "Sub Grouping", "GL Name"]}[granularity]
    ws.cell(1, 1, value=title).font = Font(bold=True, size=13, name="Calibri", color="1F3864")
    note_row = 2
    ws.cell(note_row, 1, value=(
        "Every figure below is a live formula linked to '" + schedule_name + "' - never a fixed "
        "number. Collapse the column groups ([-] above the column headers) to see just Grand "
        "Total per period. Jul-26 (and Q2 TOTAL once Jul/Aug/Sep exist) append as new columns on "
        "future refreshes.")).font = Font(italic=True, size=9, name="Calibri", color="595959")
    ws.row_dimensions[note_row].height = 28
    header_row = 5
    write_new_sheet_headers(ws, n_label_cols, label_names, header_row, blocks)
    row = header_row + 1

    def write_bucket_rows(section, source_dict, extra_label_after_grouping=None):
        nonlocal row
        if granularity == "main":
            for bucket, src_row in source_dict.items():
                labels = [None] * (n_label_cols - 1) + [bucket]
                write_new_linked_row(ws, row, n_label_cols, labels, schedule_name, src_row, blocks)
                row += 1
        elif granularity == "sub":
            for bucket, subrows in source_dict.items():
                for sg, src_row in subrows.items():
                    labels = [None, bucket, sg]
                    write_new_linked_row(ws, row, n_label_cols, labels, schedule_name, src_row, blocks)
                    row += 1
        else:
            for bucket, leafrows in source_dict.items():
                for (sg, gl), src_row in leafrows.items():
                    labels = [None, bucket, sg, gl]
                    write_new_linked_row(ws, row, n_label_cols, labels, schedule_name, src_row, blocks)
                    row += 1

    source_for = {"main": pl_section_buckets, "sub": pl_section_detail, "gl": pl_leaf_detail}[granularity]

    for section in PL_MOM_SECTION_ORDER:
        row = write_new_section_header(ws, row, n_label_cols, len(blocks), PL_MOM_SECTION_DISPLAY[section])
        write_bucket_rows(section, source_for.get(section, {}))
        if section == "COST OF SALES":
            branch_src = source_for.get("BRANCH TRANSACTIONS", {})
            if branch_src:
                row = write_new_section_header(ws, row, n_label_cols, len(blocks),
                                                "BRANCH TRANSACTIONS (memo)", font=NEW_MEMO_FONT, fill=NEW_MEMO_FILL)
                write_bucket_rows("BRANCH TRANSACTIONS", branch_src)
                write_new_linked_row(ws, row, n_label_cols, ["NET BRANCH TRANSACTIONS (memo total)"] + [None] * (n_label_cols - 1),
                                      schedule_name, pl_row_map["NET BRANCH TRANSACTIONS"], blocks,
                                      font=NEW_TOTAL_FONT)
                row += 1

    write_new_linked_row(ws, row, n_label_cols, ["EBITDA"] + [None] * (n_label_cols - 1),
                          schedule_name, pl_row_map["EBITDA"], blocks, font=NEW_MILESTONE_FONT, fill=NEW_MILESTONE_FILL)
    row += 1

    row = write_new_section_header(ws, row, n_label_cols, len(blocks), "FINANCE COST")
    write_bucket_rows("FINANCE COST", source_for.get("FINANCE COST", {}))

    write_new_linked_row(ws, row, n_label_cols, ["LESS: DEPRECIATION (not booked this period)"] + [None] * (n_label_cols - 1),
                          schedule_name, pl_row_map["LESS: DEPRECIATION"], blocks)
    row += 1
    write_new_linked_row(ws, row, n_label_cols, ["PBT (PROFIT BEFORE TAX)"] + [None] * (n_label_cols - 1),
                          schedule_name, pl_row_map["PBT"], blocks, font=NEW_MILESTONE_FONT, fill=NEW_MILESTONE_FILL)
    row += 1
    write_new_linked_row(ws, row, n_label_cols, ["LESS: TAX (not booked this period)"] + [None] * (n_label_cols - 1),
                          schedule_name, pl_row_map["LESS: TAX"], blocks)
    row += 1
    write_new_linked_row(ws, row, n_label_cols, ["PAT (NET PROFIT)"] + [None] * (n_label_cols - 1),
                          schedule_name, pl_row_map["PAT"], blocks, font=NEW_MILESTONE_FONT, fill=NEW_MILESTONE_FILL)
    row += 2

    rp_src = source_for.get("RELATED PARTY TRANSACTIONS", {})
    if rp_src:
        row = write_new_section_header(ws, row, n_label_cols, len(blocks),
                                        "RELATED PARTY TRANSACTIONS (memo)", font=NEW_MEMO_FONT, fill=NEW_MEMO_FILL)
        write_bucket_rows("RELATED PARTY TRANSACTIONS", rp_src)
        write_new_linked_row(ws, row, n_label_cols, ["NET RELATED PARTY TRANSACTIONS (memo total)"] + [None] * (n_label_cols - 1),
                              schedule_name, pl_row_map["NET RELATED PARTY TRANSACTIONS"], blocks, font=NEW_TOTAL_FONT)
        row += 1

    label_widths = {"main": [16, 30], "sub": [16, 22, 26], "gl": [14, 20, 22, 32]}[granularity]
    widths = label_widths + [16] * (8 * len(blocks))
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(header_row + 1, n_label_cols + 1).coordinate


def build_bs_group_mom_sheet(ws, section, bucket_label, schedule_name, bs_section_buckets,
                              bs_section_detail, bs_leaf_detail, periods):
    blocks = [(k, l, "month") for k, l in periods] + [("asof", "AS AT 30-JUN-26", "quarter")]
    ws.sheet_view.showGridLines = False
    n_label_cols = 2
    ws.cell(1, 1, value=f"BS GROUP: {bucket_label.upper()} - MONTH ON MONTH").font = Font(
        bold=True, size=13, name="Calibri", color="1F3864")
    ws.cell(2, 1, value=(
        "Every figure below is a live formula linked to '" + schedule_name + "'. Balance Sheet is "
        "point-in-time, not additive, so 'AS AT 30-JUN-26' is the same balance as the Jun-26 "
        "column (a same-sheet copy), not a sum of the 3 months.")).font = Font(
        italic=True, size=9, name="Calibri", color="595959")
    ws.row_dimensions[2].height = 28
    header_row = 5
    write_new_sheet_headers(ws, n_label_cols, ["Sub Grouping", "GL Name"], header_row, blocks)
    row = header_row + 1
    leaf_rows = bs_leaf_detail.get(section, {}).get(bucket_label, {})
    for (sg, gl), src_row in leaf_rows.items():
        write_new_linked_row(ws, row, n_label_cols, [sg, gl], schedule_name, src_row, blocks, quarter_mode="asof")
        row += 1
    total_row = bs_section_buckets.get(section, {}).get(bucket_label)
    if total_row is not None:
        write_new_linked_row(ws, row, n_label_cols, ["TOTAL", bucket_label], schedule_name, total_row, blocks,
                              font=NEW_TOTAL_FONT, quarter_mode="asof")
    widths = [26, 34] + [16] * (8 * len(blocks))
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(header_row + 1, n_label_cols + 1).coordinate


def bs_sheet_title(section, bucket_label):
    short = bucket_label.replace("/", "-")[:22]
    prefix = "BS-Liab" if section == "LIABILITIES" else "BS-Asset"
    return f"{prefix} {short}"


INDEX_DESCRIPTIONS = {
    "MOM PL GL Detail": "Month-on-month Profit & Loss, full GL-level detail (Apr / May / Jun + Q1 Total)",
    "MOM BS GL Detail": "Month-on-month Balance Sheet, full GL-level detail (as at Apr / May / Jun end)",
    "MOM PL and BS": "Consolidated month-on-month summary, P&L and BS combined (Grand Total view)",
    "MOM Ratio Analysis": "Monthly ratio analysis (Apr / May / Jun), each month computed independently",
    "Ratio Analysis": "Quarterly ratio analysis (Q1 Total), with full Ratio Workings shown below each figure",
    "Q1 Performance Report": "Board-style commentary: margin analysis, tax planning, cost reduction action items",
    "PL GL Detail": "Profit & Loss, full GL-level detail for the quarter, side-by-side by company",
    "BS GL Detail": "Balance Sheet, full GL-level detail as at quarter-end, side-by-side by company",
    "PL Summary": "Profit & Loss, summary down to Sub Group level - linked live to PL GL Detail",
    "BS Summary": "Balance Sheet, summary down to Sub Group level - linked live to BS GL Detail",
    "MOM PL and BS - By Company": "Month-on-month summary, one company per column block",
    "Notes": "Methodology notes: sign convention, known adjustments, how figures are derived",
    "Related Party": "All inter-company items in one place - P&L flows, BS balances, a counterparty matrix and full ledger detail. Every NET row should read 0.",
    "Cross Check": "Live formula-based verification: MOM vs Quarter totals, Balance Sheet Reconciliation Gap",
}


def build_index_sheet(wb, sheet_order):
    """Creates an 'Index' sheet as the first tab, with a clickable hyperlink + one-line
    description per sheet in sheet_order. Regenerated fresh every rebuild (deleted and recreated
    if already present) so it never goes stale - to change what it lists, edit
    INDEX_DESCRIPTIONS / desired_order in main(), not this sheet directly."""
    if "Index" in wb.sheetnames:
        del wb["Index"]
    ws = wb.create_sheet(title="Index", index=0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 90

    title_font = Font(bold=True, size=14, name="Calibri", color="0F243E")
    note_font = Font(italic=True, size=10, name="Calibri", color="595959")
    link_font = Font(underline="single", color="0563C1", size=11, name="Calibri", bold=True)
    desc_font = Font(size=10, name="Calibri")

    ws.cell(1, 1, value="INDEX").font = title_font
    ws.cell(2, 1, value="Click a sheet name below to jump straight to it.").font = note_font

    row = 4
    for name in sheet_order:
        if name not in wb.sheetnames:
            continue
        c = ws.cell(row, 1, value=name)
        c.hyperlink = f"#'{name}'!A1"
        c.font = link_font
        ws.cell(row, 2, value=INDEX_DESCRIPTIONS.get(name, "")).font = desc_font
        row += 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--timestamp", default=None,
                     help="Must match the --timestamp passed to build_gl_detail_sidebyside.py in "
                          "this same refresh. If omitted, uses whichever dated PL_BS_GL_Detail_*.xlsx "
                          "is newest in Files/ (i.e. the one that script just created).")
    args = ap.parse_args()

    files_dir = Path(__file__).parent.parent.parent.parent.parent / "Files"
    if args.timestamp:
        output_path = files_dir / f"{b.BASE_NAME}_{args.timestamp}.xlsx"
    else:
        output_path = b.find_latest_pl_bs_file(files_dir)
    if output_path is None or not output_path.exists():
        raise RuntimeError(f"No PL_BS_GL_Detail_*.xlsx found to open (looked for {output_path!r}). "
                            f"Run build_gl_detail_sidebyside.py first.")
    wb = openpyxl.load_workbook(output_path)

    # Load the grouping snapshot build_gl_detail_sidebyside.py captured (from the file as it stood
    # BEFORE this whole rebuild chain started) and persisted to a JSON sidecar - the MOM sheets
    # about to be deleted and recreated below were already wiped once by that script's own
    # from-scratch rebuild, so snapshotting this in-memory `wb` here would be too late to recover
    # their original hidden state. Restored just before saving.
    snap = b.load_grouping_snapshot(output_path)

    for name in ("MOM PL and BS", "MOM PL and BS - By Company", "MOM PL GL Detail", "MOM BS GL Detail",
                 "MOM Ratio Analysis", "Cross Check",
                 "PL Main Group MOM", "PL Sub Group MOM", "PL Full GL MOM"):
        if name in wb.sheetnames:
            del wb[name]
    for name in list(wb.sheetnames):
        if name.startswith("BS-Liab ") or name.startswith("BS-Asset "):
            del wb[name]

    ws_pl_sched = wb.create_sheet(title="MOM PL GL Detail")
    pl_row_map, pl_section_buckets, pl_section_detail, pl_leaf_detail = write_mom_pl_schedule(ws_pl_sched)

    ws_bs_sched = wb.create_sheet(title="MOM BS GL Detail")
    bs_row_map, bs_section_buckets, bs_section_detail, bs_leaf_detail = write_mom_bs_schedule(ws_bs_sched)

    ws1 = wb.create_sheet(title="MOM PL and BS")
    next_row = write_mom_pl_consolidated_summary(ws1, "MOM PL GL Detail", pl_row_map, pl_section_buckets,
                                                  pl_section_detail, 1)
    write_mom_bs_consolidated_summary(ws1, "MOM BS GL Detail", bs_row_map, bs_section_buckets, next_row)

    ws2 = wb.create_sheet(title="MOM PL and BS - By Company")
    next_row2 = write_mom_pl_by_company_summary(ws2, "MOM PL GL Detail", pl_row_map, pl_section_buckets,
                                                 pl_section_detail, 1)
    write_mom_bs_by_company_summary(ws2, "MOM BS GL Detail", bs_row_map, bs_section_buckets, next_row2)

    ws_mom_ratios = wb.create_sheet(title="MOM Ratio Analysis")
    write_mom_ratio_sheet(ws_mom_ratios, pl_row_map, bs_row_map)

    # NOTE: an earlier attempt (2026-07-25) added PL Main/Sub/Full-GL MOM sheets + one sheet per
    # real BS Group directly into this production workbook (39 sheets total). The user found that
    # overwhelming ("There are Mutiple SHeet now created I am Completely lost. Please stop this
    # activity as of now.") and this was reverted. That exploration continued instead as a
    # standalone sample workbook (Files/SAMPLE_New_Sheets_Preview_v4.xlsx) - do not re-wire
    # build_pl_mom_extended_sheet / build_bs_group_mom_sheet back into this production script
    # unless the user explicitly asks for the production workbook itself to include them.

    write_related_party_sheet(wb)
    write_cross_check_sheet(wb)

    b.restore_column_row_hidden(wb, snap)

    # Final tab order for the whole workbook (this script runs last): MOM sheets first, then the
    # quarterly Ratio Analysis, then the base PL/BS Summary + GL Detail + Notes sheets, then
    # Cross Check last as a diagnostic. Fixed here per standing instruction - do not reorder by
    # hand in Excel, since every rebuild re-applies this order.
    # Order reflects the user's own manual tab rearrangement in Excel (2026-07-22) - baked in
    # here so a rebuild reproduces it rather than reverting to the original script-default order.
    # Any sheet NOT listed here is silently dropped by the wb._sheets reassignment below, so new
    # sheet names must always be added to this list.
    desired_order = [
        "MOM PL GL Detail", "MOM BS GL Detail", "MOM PL and BS",
        "MOM Ratio Analysis",
        "Ratio Analysis",
        "Q1 Performance Report",
        "PL GL Detail", "BS GL Detail", "PL Summary", "BS Summary",
        "MOM PL and BS - By Company", "Notes",
        "Related Party",
        "Cross Check",
    ]
    build_index_sheet(wb, desired_order)
    desired_order = ["Index"] + desired_order
    wb._sheets = [wb[name] for name in desired_order if name in wb.sheetnames]

    try:
        wb.save(output_path)
    except PermissionError:
        raise RuntimeError(f"Cannot save to {output_path!r} - file may be open in Excel.\nClose the file and retry.")

    print(f"File saved: {output_path}")


if __name__ == "__main__":
    main()
