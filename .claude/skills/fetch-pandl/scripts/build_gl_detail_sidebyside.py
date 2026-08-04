"""
Builds Files/PL_BS_GL_Detail_<timestamp>.xlsx: full GL-level detail, side-by-side by company
(Grouping | Sub Grouping | GL Name | <company columns> | Total of Orange O Tec |
<company columns> | Total of Enterprises | Colorix | Grand Total), with a subtotal row
per Sub Grouping and per Grouping (both collapsible via Excel row outlining, 2 levels)
and a final GRAND TOTAL row. Period: 01-Apr-2026 to 30-Jun-2026.

Source data: CSVs pre-fetched via the tally-prime MCP tools, saved to gl_data/ next to this
script (inside the skill folder itself, NOT a session scratchpad - so this rebuilds correctly
from any future session regardless of session ID). Each CSV: primary_group, group_name,
ledger_name, closing_balance.
"""

import argparse
import csv
import json
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

GL_DATA_DIR = Path(__file__).parent.parent / "gl_data"

PERIOD_LABEL = "01-Apr-2026 to 30-Jun-2026"

# Column order matches the user-provided mockup exactly.
CO1 = ("ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)", "ORANGE_O_TEC_PRIVATE_LIMITED")
CO2 = ("ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)", "PRIVATE_LIMITED_NOIDA")
CO3 = ("ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)", "ENTERPRISES_PVT_LTD")
CO4 = ("ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27", "ENTERPRISES_NOIDA")
CO5 = ("COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)", "COLORIX")
ALL_COMPANIES = [CO1, CO2, CO3, CO4, CO5]

STOCK = {
    "ORANGE_O_TEC_PRIVATE_LIMITED": {"opening": 292508691.08, "closing": 277198530.32},
    "PRIVATE_LIMITED_NOIDA": {"opening": 32685329.56, "closing": 29258524.91},
    "ENTERPRISES_PVT_LTD": {"opening": 71324050.22, "closing": 91177937.29},
    "ENTERPRISES_NOIDA": {"opening": 32959492.39, "closing": 37910968.1},
    "COLORIX": {"opening": 105648, "closing": 105648},
}

# "Sales Bills to Make": goods dispatched (revenue already earned per Tally's own books) where
# the formal sales invoice has not yet been raised. Tally tracks this as a real receivable under
# Sundry Debtors and includes it in the Balance Sheet / P&L, but it is computed from pending
# Delivery Note vouchers - it is NOT a ledger or group master, so it never appears via the
# chart-of-accounts / ledger / group API. Confirmed manually against Tally's own Balance Sheet ->
# Current Assets -> Sundry Debtors -> Group Summary screen for the quarter (01-Apr-2026 to
# 30-Jun-2026); reconciles both P&L and BS to the rupee once added. Only confirmed for companies
# listed here - other companies' residual BS gaps remain unexplained until checked the same way.
# Each entry is per-period, not a single constant - this balance genuinely grows/shrinks month to
# month (pending dispatches get invoiced, new ones accrue), so the 30-Jun figure cannot be reused
# for 30-Apr/31-May without breaking the Balance Sheet Reconciliation Gap for those months. "apr"/
# "may" values were back-solved from the observed Reconciliation Gap for that company/period (the
# gap IS the missing adjustment, since this is the only known unexplained item for that company)
# and cross-verified 2026-07-18 against Tally's own Balance Sheet -> Sundry Debtors -> Group
# Summary screen directly, for both Apr-26 (exact match, 1,70,17,545.41) and May-26 (screen showed
# 1,58,42,402.19 for Sales Bills to Make specifically - the remaining 12,032.00 of the originally
# back-solved 1,58,54,434.19 turned out to be a SEPARATE small "Unadjusted Forex Gain/Loss" line
# for this company, on the Assets side - see UNADJUSTED_FOREX_GAIN_LOSS_ASSET below). "qtr" remains
# the original manually-confirmed 30-Jun value and is what June (which reuses load_company_rows via
# MOM's "jun" BS path) and the quarter itself use.
#
# ENTERPRISES_PVT_LTD "may" was initially misattributed to PURCHASE_BILLS_TO_COME below (a pure
# back-solve from the Reconciliation Gap total, before either was screen-verified - the same
# rupee amount zeros the gap regardless of which side/bucket it's placed in, since Gap =
# Liabilities + Assets sums both buckets together, so a wrong bucket can still "work" arithmetically
# while being factually wrong). Corrected 2026-07-18 after isolating the actual discrepancy to
# Sundry Debtors specifically (Tally's own Current Assets -> Group Summary showed Sundry Debtors
# short by exactly this amount) and then confirming directly against Tally's own Sundry Debtors ->
# Group Summary screen: "Sales Bills to Make" = 3,51,050.00 exactly, 31-May-26. Apr needed no
# adjustment for this company (that gap was already 0) - this item apparently arose during May and
# was fully invoiced/cleared by 30-Jun (the quarter-end reconciliation was already confirmed
# complete using only PURCHASE_BILLS_TO_COME below, no Sales Bills to Make needed at quarter-end).
#
# 2026-08-03: these are no longer hand-transcribed from a Tally screenshot. A LEDGER-WISE Trial
# Balance (fetch_gl_data.py --computed-lines) returns Tally's own computed Balance Sheet lines
# directly, and every previously screen-verified figure below came back from it to the paisa
# (Orange O Tec apr/may/qtr, Enterprises may, Enterprises PURCHASE_BILLS_TO_COME qtr, Colorix
# apr). That is the preferred source for any new period now - run --computed-lines, read the
# value, paste it here. The values are stored as POSITIVE magnitudes; the raw Trial Balance row
# is negative because it is an asset-side/contra figure.
SALES_BILLS_TO_MAKE = {
    "ORANGE_O_TEC_PRIVATE_LIMITED": {"qtr": 35479998.84, "apr": 17017545.41, "may": 15842402.19,
                                      "jul": 51030000.00, "ytd": 51030000.00},
    "ENTERPRISES_PVT_LTD": {"may": 351050.00, "jul": 56100.00, "ytd": 56100.00},
    "ENTERPRISES_NOIDA": {"jul": 173022.00, "ytd": 173022.00},
    "PRIVATE_LIMITED_NOIDA": {"jul": 30.00, "ytd": 30.00},
}

# "Purchase Bills to Come": the mirror-image of Sales Bills to Make - goods/services received
# from a related party where the formal purchase invoice has not yet been booked. Confirmed
# against Tally's own Balance Sheet -> Current Liabilities -> Sundry Creditors -> Group Summary
# screen (shown there as a contra reduction) AND independently against Tally's own Purchase Group
# Summary screen (shown there as its own line reducing Purchase Accounts) - both screens show the
# identical amount, confirming it belongs on both statements, like Sales Bills to Make. Added as a
# positive value to Purchase Accounts in P&L (reduces Cost of Sales, raises PAT) and as a negative
# value to Sundry Creditors in BS (reduces the liability) - see load_company_rows. Only confirmed
# at quarter-end (30-Jun) for this company - Apr/May need no Purchase Bills to Come adjustment at
# all (their small residual gaps turned out to be a SALES_BILLS_TO_MAKE item instead - see above).
PURCHASE_BILLS_TO_COME = {
    "ENTERPRISES_PVT_LTD": {"qtr": 678615.00, "jul": 678615.00, "ytd": 678615.00},
}

# "Unadjusted Forex Gain/Loss": a standalone line in Tally's own Balance Sheet. Not tied to any
# ledger/group master, so it never appears via the chart-of-accounts / ledger API. BS-side only.
# For Colorix this sits on the LIABILITIES side (Sources of Funds); confirmed via Tally's own
# screen for both Apr-26 and May-26 - exact match.
# 2026-08-03: the "qtr" and "may" entries for Colorix (both 45,600.00) were REMOVED. They were
# never a computed Balance Sheet line at all - a ledger-wise Trial Balance shows Colorix has no
# such line for May or June. The 45,600.00 is a REAL LEDGER, 'EXCHANGE RATE DIFF EXP ACCOUNTS'
# (Direct Expenses), which the old MCP profit-loss/balance-sheet queries silently omitted; the
# constant was back-solved to plug the hole it left. fetch_gl_data.py picks that ledger up, and
# the arithmetic ties exactly: the old CSV's Profit & Loss A/c of 38,09,604.33 plus 45,600.00 is
# the 38,55,204.33 now computed. Keeping the constant on top of the live ledger would
# double-count it. "apr" stays - Tally does report a genuine computed line for April.
UNADJUSTED_FOREX_GAIN_LOSS = {
    # SIGNED, unlike the other three dicts: this one is written straight onto the liabilities
    # side, so Tally's own direction has to survive. April is a forex GAIN (+44,266.20), July a
    # forex LOSS (-12,987.45) - storing July as a positive magnitude put it on the wrong side and
    # left Colorix's July Reconciliation Gap at 25,974.90, exactly twice the figure.
    "COLORIX": {"apr": 44266.20, "jul": -12987.45, "ytd": -12987.45},
}

# Same concept as UNADJUSTED_FOREX_GAIN_LOSS above, but for a company where Tally shows it on the
# ASSETS side of the Balance Sheet instead of Liabilities (confirmed via screen for Orange O Tec
# Pvt Ltd, 31-May-26 - a separate, smaller item from that same company's Sales Bills to Make).
# Kept as its own dict/primary_group ("Unadjusted Forex Gain/Loss (Asset)", see BS_ASSET_ORDER)
# rather than reusing UNADJUSTED_FOREX_GAIN_LOSS, since the same primary_group string cannot be
# classified as both a liability and an asset depending on which company it belongs to.
# 2026-08-03: refreshed from the live computed-line report. May moved 12,032.00 -> 56,318.00 and
# a quarter-end / 31-Jul figure now exists where previously none was recorded (postings against
# this item continued after the original back-solve). Verify each against that period's own
# Balance Sheet Reconciliation Gap after a rebuild before trusting it - a figure that zeroes the
# Gap can still sit on the wrong side, which is the trap SKILL.md documents at length.
UNADJUSTED_FOREX_GAIN_LOSS_ASSET = {
    # Re-read from Tally 2026-08-04 (fetch_gl_data.py --computed-lines). These moved since the
    # 08-03 read: the quarter figure fell 77,750.00 -> 55,000.00 and the 31-Jul line was cleared
    # entirely. Leaving the stale values in place produced Reconciliation Gaps of exactly the
    # difference - 22,750.00 on the quarter sheets and 2,48,118.00 on the MOM Balance Sheet -
    # which is the signature of a stale constant: the gap equals the drift, to the rupee.
    "ORANGE_O_TEC_PRIVATE_LIMITED": {"may": 56318.00, "qtr": 55000.00},
}

BS_GROUP_LABEL = {"": "Profit & Loss A/c (current)", "Bank Account": "Bank Accounts"}

STOCK_VALUATION_SUBGROUP = "Stock Valuation (Inventory)"

FINANCE_SUBGROUPS = {"Finance Cost"}
RELATED_PARTY_SUBGROUPS = {"RELATED PARTY SALES", "RELATED PARTY PURCHASE", "RELATED SALES", "RELATED PURCHASE"}

# P&L: bucket order within each statement section (primary_group values)
PL_TURNOVER_BUCKETS = ["Sales Accounts", "Direct Incomes"]
PL_COST_OF_SALES_BUCKETS = ["Purchase Accounts", "Direct Expenses", "Stock Valuation (Inventory)"]
PL_OTHER_INCOME_BUCKETS = ["Indirect Incomes"]
PL_OPERATING_EXPENSE_BUCKETS = ["Indirect Expenses"]
PL_FINANCE_COST_BUCKETS = ["Finance Cost"]
PL_BRANCH_BUCKETS = ["Branch Sales", "Branch Purchase"]
PL_RELATED_PARTY_BUCKETS = ["Related Party Sales", "Related Party Purchase"]


# BS: which primary_group values sit on which side of the balance sheet
BS_LIABILITY_ORDER = ["Capital Account", "Reserves & Surplus", "Secured Loans", "Unsecured Loans",
                       "Bank OD A/c", "Duties & Taxes", "Provisions", "Sundry Creditors",
                       "Current Liabilities", "Loans (Liability)", "Branch / Divisions",
                       "Suspense A/c", "Unadjusted Forex Gain/Loss", ""]
BS_ASSET_ORDER = ["Fixed Assets", "Investments", "Current Assets", "Bank Accounts", "Bank Account",
                   "Cash-in-hand", "Deposits (Asset)", "Loans & Advances (Asset)", "Sundry Debtors",
                   "Misc. Expenses (ASSET)", "Unadjusted Forex Gain/Loss (Asset)"]

# Case-insensitive lookup so Tally's inconsistent casing for the same conceptual group across
# companies/fetches (e.g. 'Cash-in-Hand' vs 'Cash-in-hand') never silently drops a ledger from
# classify_bs_row() / buckets_for(), which both do an exact string match against the two lists
# above. See load_csv() guard 3.
CANONICAL_PRIMARY_GROUP = {pg.lower(): pg for pg in BS_LIABILITY_ORDER + BS_ASSET_ORDER if pg}

# ---------------------------------------------------------------------------
# Display sign convention (2026-08-03 onward, user request: "represent all expenses in positive")
# ---------------------------------------------------------------------------
# The CSVs and every load_* function keep Tally's native convention untouched - assets and
# expenses NEGATIVE - because the Balance Sheet reconciliation, the adjustment constants and the
# fetcher's own validation all depend on it. The flip below is applied ONCE, at the point where
# classified rows are handed to the sheet writers, so what a reader SEES is a conventional
# statement (expenses and assets as positive magnitudes) while everything upstream is unchanged.
#
# Because every subtotal, bucket total and section total in this workbook is a live SUM() over the
# leaf cells, negating the leaves makes all of them follow automatically. Only two things need
# hand-adjustment, and both are done explicitly:
#   1. the headline waterfall rows (Gross Profit, EBITDA, PBT, PAT) - their component signs flip
#      from '+' to '-' for expense inputs, see write_pl_sheet;
#   2. the two MIXED memo sections (Branch, Related Party), which hold an income leg AND an
#      expense leg - a plain SUM of the two would now ADD them instead of netting, so those
#      section totals are written as signed formulas via write_section(bucket_signs=...).
# The acceptance test for the whole change is that Gross Profit, EBITDA, PBT and PAT come out
# bit-identical to the pre-flip workbook.
DISPLAY_NEGATED_BUCKETS = frozenset(
    PL_COST_OF_SALES_BUCKETS + PL_OPERATING_EXPENSE_BUCKETS + PL_FINANCE_COST_BUCKETS
    + ["Branch Purchase", "Related Party Purchase"]
    + BS_ASSET_ORDER
)


def display_sign(bucket: str) -> int:
    return -1 if bucket in DISPLAY_NEGATED_BUCKETS else 1


def apply_display_signs(data):
    """Negates every value in a {section: {bucket: {company: [(sub_grp, gl_name, bal), ...]}}}
    tree whose bucket is an expense/asset bucket. Returns the same shape."""
    for section in data.values():
        for bucket, by_company in section.items():
            s = display_sign(bucket)
            if s == 1:
                continue
            for key, rows in by_company.items():
                by_company[key] = [(sg, gl, -bal) for sg, gl, bal in rows]
    return data

# Canonical Indirect Expense sub-group list, agreed with the user 2026-07-17 (see
# Indirect_Expense_Grouping_Proposal.xlsx for the company-wise before/after this was built from).
# Tally's own chart of accounts is not consistent about group_name spelling/grouping across the 5
# companies, so the same conceptual cost was showing up under several different sub-group labels
# (e.g. four separate "Employee ..." rows, three separate bank-charge rows) - read as duplicated or
# scattered lines in the PL schedule/summary/MOM sheets. Applied via CANONICAL_SUB_GROUP
# (case-insensitive) in load_csv(), so every sheet that reads from these CSVs gets it automatically:
#   - Employee Cost: Salary + Benefits + Managerial Remuneration + Staff Welfare + Employee Cost
#     (confirmed against Tally's own P&L Group Summary, which rolls these four into one figure)
#   - Employee Travel Cost: Travelling + Vehicle + Hotel (confirmed against Tally's own P&L screen
#     for Orange O Tec Pvt Ltd, whose "Employee Travel Cost" line tied to the rupee against exactly
#     these three)
#   - Finance Cost: Interest Cost + all bank/LC/import-payment charges (user confirmed 2026-07-17)
#   - Insurance Exp: also absorbs Colorix's un-subgrouped "Indirect Expenses" catch-all row (its
#     only ledger there, MARINE INSURANCE EXP, is an insurance cost with no proper sub-group from
#     Tally's chart of accounts)
#   - Printing and Courier: Printing Stationery + Courier & Postage
#   - Repair & Maintenance Cost: also absorbs Repair & Maintenance Exp (Office)
#   - Subscription & Membership Exp: also absorbs AI Subscription
#   - Electricity Expenses: also absorbs Electrical Expenses
#   - Transportation Exp: also absorbs Carriage Outward (both spellings) + Loading & Unloading
#     (user confirmed 2026-07-17 - these had no clean home of their own in the canonical list)
CANONICAL_SUB_GROUP = {
    "employee salary expenses": "Employee Cost",
    "employee benefits expenses": "Employee Cost",
    "managerial remuneration": "Employee Cost",
    "staff welfare expenses": "Employee Cost",
    "employee cost": "Employee Cost",
    "employee travel cost": "Employee Travel Cost",
    "travelling exp": "Employee Travel Cost",
    "vehical exp": "Employee Travel Cost",
    "hotel expenses": "Employee Travel Cost",
    "finance cost": "Finance Cost",
    "interest cost": "Finance Cost",
    "other bank charges": "Finance Cost",
    "bank charges": "Finance Cost",
    "bank charges-import direct payment": "Finance Cost",
    "bank charges-lc-bc": "Finance Cost",
    "insurance exp": "Insurance Exp",
    "indirect expenses": "Insurance Exp",
    "printing and courier": "Printing and Courier",
    "printing stationery exp": "Printing and Courier",
    "courier & postage exp": "Printing and Courier",
    "repair & maintenance cost": "Repair & Maintenance Cost",
    "repair & maintenance exp (office)": "Repair & Maintenance Cost",
    "subscription & membership exp": "Subscription & Membership Exp",
    "ai subscription": "Subscription & Membership Exp",
    "electricity expenses": "Electricity Expenses",
    "electrical expenses": "Electricity Expenses",
    "transportation exp": "Transportation Exp",
    "carriage outward accounts": "Transportation Exp",
    "carriage outward exp": "Transportation Exp",
    "loading & unloading exp": "Transportation Exp",
    "advertising promotion exp": "Advertising Promotion Exp",
    "commission & brokerage exp": "Commission & Brokerage Exp",
    "computer & printer exp": "Computer & Printer Exp",
    "donation subscription": "Donation Subscription",
    "environmental protection": "Environmental Protection",
    "laboratory expences": "Laboratory Expences",
    "learning and development": "Learning and Development",
    "legal and professional exp": "Legal and Professional Exp",
    "office expenses": "Office Expenses",
    "rent expenses": "Rent Expenses",
    "security exp": "Security Exp",
    "taxes exp": "Taxes Exp",
    "telephone exp": "Telephone Exp",
}


def classify_pl_row(primary_group, sub_group, ledger_name=""):
    if primary_group in ("Sales Accounts", "Purchase Accounts") and "BRANCH" in ledger_name.upper():
        return "BRANCH TRANSACTIONS", "Branch Sales" if primary_group == "Sales Accounts" else "Branch Purchase"
    if primary_group in ("Sales Accounts", "Direct Incomes"):
        return "TURNOVER", primary_group
    if primary_group in ("Purchase Accounts", "Direct Expenses", "Stock Valuation (Inventory)"):
        return "COST OF SALES", primary_group
    if primary_group == "Indirect Incomes":
        return "OTHER INCOME", primary_group
    if primary_group == "Indirect Expenses":
        if sub_group in FINANCE_SUBGROUPS:
            return "FINANCE COST", "Finance Cost"
        return "OPERATING EXPENSES", primary_group
    return "OPERATING EXPENSES", primary_group  # safety net for any unclassified group


def is_related_party_row(primary_group, sub_group):
    return sub_group in RELATED_PARTY_SUBGROUPS and primary_group in ("Sales Accounts", "Purchase Accounts")


def classify_bs_row(primary_group):
    if primary_group in BS_LIABILITY_ORDER:
        return "LIABILITIES", primary_group
    return "ASSETS", primary_group


def load_csv(path: Path):
    """Loads (primary_group, group_name, ledger_name, closing_balance) rows, with two guards
    against Tally API quirks observed empirically (inconsistently, across different companies
    and date ranges - not tied to any one company or period):
    1. Skips any ledger_name seen more than once (keeps the first occurrence). Ledger names are
       unique per Tally company, but the balance-sheet API has been observed to return the same
       ledger twice for certain date ranges (e.g. 'Profit & Loss A/c' under both '' and
       'Reserves & Surplus' group_name, same balance) - without this guard that ledger would be
       double-counted.
    2. Skips rows under group_name 'Stock-in-Hand' (Tally's own native 'Opening Stock'/'Closing
       Stock' entries). For some companies/date ranges the P&L and Balance Sheet API responses
       already include these - confirmed to exactly match the figures we compute independently
       via the stock-summary tool (STOCK / STOCK_MOM). Relying solely on our own stock-summary
       based addition avoids double-counting and avoids these rows landing in the wrong section
       (their primary_group is blank, which would otherwise misclassify them).
    3. Normalizes primary_group casing against BS_LIABILITY_ORDER/BS_ASSET_ORDER (case-insensitive
       match). Tally returns inconsistent casing for the same conceptual group across different
       companies/fetches (e.g. 'Cash-in-Hand' vs 'Cash-in-hand') - classify_bs_row() and
       buckets_for() both do an exact string match against those two lists, so an unnormalized
       casing mismatch causes the ledger to be SILENTLY DROPPED from the workbook entirely (not
       misclassified - just missing, with no error), which is exactly what caused a real ~Rs
       5,43,776 Balance Sheet Reconciliation Gap for Enterprises Pvt Ltd (Cash + Petty Cash) before
       this guard was added. Any primary_group not found in either list (case-insensitively) is
       left as-is.
    4. Normalizes group_name against CANONICAL_SUB_GROUP (case-insensitive) so employee-cost
       ledgers filed under different Tally group names (EMPLOYEE SALARY EXPENSES, EMPLOYEE
       BENEFITS EXPENSES, MANAGERIAL REMUNERATION, STAFF WELFARE EXPENSES, EMPLOYEE COST) collapse
       into one "Employee Cost" sub-group row instead of appearing as several separate lines."""
    rows = []
    seen_ledgers = set()
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            if r["group_name"] == "Stock-in-Hand":
                continue
            ledger_name = r["ledger_name"]
            if ledger_name in seen_ledgers:
                continue
            seen_ledgers.add(ledger_name)
            primary_group = CANONICAL_PRIMARY_GROUP.get(r["primary_group"].strip().lower(), r["primary_group"])
            group_name = CANONICAL_SUB_GROUP.get(r["group_name"].strip().lower(), r["group_name"])
            rows.append((primary_group, group_name, ledger_name, float(r["closing_balance"])))
    return rows


PL_AC_LEDGER = "Profit & Loss A/c"


def apply_stock_to_pl_ac(rows, opening: float, closing: float):
    """Adds the period's stock movement (closing - opening) to the Balance Sheet's
    'Profit & Loss A/c' line, in place.

    Why this is needed: this project deliberately replaces Tally's own inventory figures with
    its STOCK / STOCK_MOM values, injecting 'Closing Stock (Inventory)' as an ASSET. Tally's own
    Balance Sheet carries that same asset but ALSO carries the matching stock movement inside
    retained earnings (its P&L A/c line includes +closing - opening). Injecting only the asset
    leaves the Balance Sheet short by exactly that movement on the liabilities side.

    Confirmed to the rupee against the Reconciliation Gap for every company and month on
    2026-08-03 - e.g. Enterprises Noida Apr-26 gap of -14,08,172.67 is precisely its
    3,43,67,665.06 closing stock less its 3,29,59,492.39 opening stock. Note 'opening' is always
    the FY-START opening, not the prior month's close, because the Balance Sheet is cumulative.
    """
    delta = closing - opening
    if not delta:
        return rows
    for i, (pg, sg, gl, bal) in enumerate(rows):
        if gl.strip() == PL_AC_LEDGER:
            rows[i] = (pg, sg, gl, bal + delta)
            return rows
    rows.append(("", "", PL_AC_LEDGER, delta))
    return rows


def load_company_rows(file_prefix: str, file_key: str, stat_type: str):
    csv_path = GL_DATA_DIR / f"{file_prefix}_{file_key}.csv"
    rows = load_csv(csv_path)
    stock = STOCK[file_key]
    sales_bills_to_make = SALES_BILLS_TO_MAKE.get(file_key, {}).get("qtr", 0.0)
    if stat_type == "pl":
        rows.append(("Stock Valuation (Inventory)", STOCK_VALUATION_SUBGROUP, "Opening Stock", -stock["opening"]))
        rows.append(("Stock Valuation (Inventory)", STOCK_VALUATION_SUBGROUP, "Closing Stock", stock["closing"]))
        if sales_bills_to_make:
            rows.append(("Sales Accounts", "Sales Bills to Make", "Sales Bills to Make", sales_bills_to_make))
        purchase_bills_to_come = PURCHASE_BILLS_TO_COME.get(file_key, {}).get("qtr", 0.0)
        if purchase_bills_to_come:
            rows.append(("Purchase Accounts", "Purchase Bills to Come", "Purchase Bills to Come", purchase_bills_to_come))
    elif stat_type == "bs":
        rows.append(("Current Assets", "Closing Stock (Inventory)", "Closing Stock (Inventory)", -stock["closing"]))
        apply_stock_to_pl_ac(rows, stock["opening"], stock["closing"])
        if sales_bills_to_make:
            rows.append(("Sundry Debtors", "Sales Bills to Make", "Sales Bills to Make", -sales_bills_to_make))
        purchase_bills_to_come = PURCHASE_BILLS_TO_COME.get(file_key, {}).get("qtr", 0.0)
        if purchase_bills_to_come:
            rows.append(("Sundry Creditors", "Purchase Bills to Come", "Purchase Bills to Come", -purchase_bills_to_come))
        unadjusted_forex = UNADJUSTED_FOREX_GAIN_LOSS.get(file_key, {}).get("qtr", 0.0)
        if unadjusted_forex:
            rows.append(("Unadjusted Forex Gain/Loss", "Unadjusted Forex Gain/Loss", "Unadjusted Forex Gain/Loss", unadjusted_forex))
        # The ASSET-side variant was previously handled ONLY in build_mom_gl_detail's
        # load_period_rows (i.e. only for Apr/May), so the quarter sheets - and June's BS, which
        # reuses this function - could never pick one up. That was fine while the dict held a
        # May-only value, but the live computed-line report now shows a 30-Jun figure too, and
        # without this branch it would silently vanish from the quarter's Reconciliation Gap.
        unadjusted_forex_asset = UNADJUSTED_FOREX_GAIN_LOSS_ASSET.get(file_key, {}).get("qtr", 0.0)
        if unadjusted_forex_asset:
            rows.append(("Unadjusted Forex Gain/Loss (Asset)", "Unadjusted Forex Gain/Loss (Asset)",
                         "Unadjusted Forex Gain/Loss (Asset)", -unadjusted_forex_asset))
    return rows


# Professional navy / slate palette. The P&L waterfall (GP -> EBITDA -> PBT -> PAT) deepens
# through one consistent hue rather than cycling through unrelated colors, so the reader's eye
# reads "closer to the bottom line" as "darker", not as a jump between unrelated categories.
# The one deliberate exception is the amber/orange used for genuine reconciliation warnings.
TITLE_FONT = Font(bold=True, size=13, name="Calibri", color="1F3864")
HEADER_FONT = Font(bold=True, size=10, name="Calibri", color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
TOTALCOL_FILL = PatternFill("solid", fgColor="2E5395")
GROUP_FONT = Font(bold=True, size=10, name="Calibri", color="1F3864")
SUBGROUP_FONT = Font(bold=True, italic=True, size=9, name="Calibri", color="44546A")
NORMAL_FONT = Font(size=10, name="Calibri")
SUBTOTAL_SUBGROUP_FONT = Font(bold=True, italic=True, size=9, name="Calibri", color="1F3864")
SUBTOTAL_SUBGROUP_FILL = PatternFill("solid", fgColor="EDF1F8")
SUBTOTAL_FONT = Font(bold=True, italic=True, size=10, name="Calibri", color="1F3864")
SUBTOTAL_FILL = PatternFill("solid", fgColor="DCE4F0")
GRANDTOTAL_FONT = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
GRANDTOTAL_FILL = PatternFill("solid", fgColor="0F243E")
NUM_FORMAT = "#,##0.00"

SECTION_FONT = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="44546A")
SECTION_TOTAL_FONT = Font(bold=True, size=10, name="Calibri", color="1F3864")
SECTION_TOTAL_FILL = PatternFill("solid", fgColor="D6DCE5")
GP_FONT = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
GP_FILL = PatternFill("solid", fgColor="2E5395")
EBITDA_FONT = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
EBITDA_FILL = PatternFill("solid", fgColor="1F3864")
PBT_FONT = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
PBT_FILL = PatternFill("solid", fgColor="17375E")
PAT_FONT = Font(bold=True, size=12, name="Calibri", color="FFFFFF")
PAT_FILL = PatternFill("solid", fgColor="0F243E")
ZERO_FONT = Font(italic=True, size=9, name="Calibri", color="808080")
ZERO_FILL = PatternFill("solid", fgColor="F2F2F2")
GAP_FONT = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
GAP_FILL = PatternFill("solid", fgColor="C55A11")

# Adjustment / memo sections (Branch Transactions, Related Party Transactions) - neutral grey
# so they visually read as supplementary disclosure, not part of the main P&L waterfall.
ADJUSTMENT_FONT = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
ADJUSTMENT_FILL = PatternFill("solid", fgColor="757575")
ADJUSTMENT_TOTAL_FONT = Font(bold=True, size=10, name="Calibri", color="404040")
ADJUSTMENT_TOTAL_FILL = PatternFill("solid", fgColor="E7E6E6")

# Column layout: A Section, B Grouping, C Sub Grouping, D GL Name, E..L values
COL_SECTION = 1
COL_GROUPING = 2
COL_SUBGROUPING = 3
COL_GLNAME = 4
COL_CO1 = 5
COL_CO2 = 6
COL_TOTAL_ORANGE = 7
COL_CO3 = 8
COL_CO4 = 9
COL_TOTAL_ENTERPRISES = 10
COL_CO5 = 11
COL_GRANDTOTAL = 12


def col_sum(vals):
    return vals[0] + vals[1] + vals[2] + vals[3] + vals[4]


def family_totals(vals5):
    co1, co2, co3, co4, co5 = vals5
    total_orange = co1 + co2
    total_enterprises = co3 + co4
    grand_total = total_orange + total_enterprises + co5
    return total_orange, total_enterprises, grand_total


PLAIN_COLS = (COL_CO1, COL_CO2, COL_CO3, COL_CO4, COL_CO5)
PLAIN_COL_LETTERS = {col: get_column_letter(col) for col in PLAIN_COLS}
TOTAL_ORANGE_L = get_column_letter(COL_TOTAL_ORANGE)
TOTAL_ENTERPRISES_L = get_column_letter(COL_TOTAL_ENTERPRISES)


def plain_dict(vals5):
    """Maps a raw [co1..co5] list to {col: value} for the five source-data columns."""
    return dict(zip(PLAIN_COLS, vals5))


def write_value_cells(ws, row, plain_cell_values, font, fill=None, skip_zero=False):
    """plain_cell_values: {col: value} for the five company columns, where each value is either
    a raw number (GL detail leaf rows - the source figure, nothing to add) or an Excel formula
    string (any subtotal/total/computed row - a live formula over its child rows). The three
    Total columns (Total of Orange O Tec, Total of Enterprises, Grand Total) are ALWAYS written
    as Excel formulas summing the plain columns of this same row - never an independently
    computed number - so pressing F2 on any total in the workbook shows exactly how it was
    built, on every row without exception."""
    for col, val in plain_cell_values.items():
        if skip_zero and isinstance(val, (int, float)) and abs(val) < 0.005:
            continue
        cell = ws.cell(row, col, value=val)
        cell.number_format = NUM_FORMAT
        cell.font = font
    l = PLAIN_COL_LETTERS
    ws.cell(row, COL_TOTAL_ORANGE, value=f"={l[COL_CO1]}{row}+{l[COL_CO2]}{row}")
    ws.cell(row, COL_TOTAL_ENTERPRISES, value=f"={l[COL_CO3]}{row}+{l[COL_CO4]}{row}")
    ws.cell(row, COL_GRANDTOTAL, value=f"={TOTAL_ORANGE_L}{row}+{TOTAL_ENTERPRISES_L}{row}+{l[COL_CO5]}{row}")
    for col in (COL_TOTAL_ORANGE, COL_TOTAL_ENTERPRISES, COL_GRANDTOTAL):
        cell = ws.cell(row, col)
        cell.number_format = NUM_FORMAT
        cell.font = font
    if fill:
        for c in range(1, COL_GRANDTOTAL + 1):
            ws.cell(row, c).fill = fill


def apply_family_column_outline(ws, co1, co2, total_orange, co3, co4, total_enterprises, co5, grandtotal):
    """Column outline mirroring the row outline used for expense detail: collapsing level 2 hides
    the four individual-company columns (down to family totals + Colorix + Grand Total);
    collapsing level 1 as well hides the family totals and Colorix too, leaving only Grand Total."""
    for c in (co1, co2, co3, co4):
        ws.column_dimensions[get_column_letter(c)].outlineLevel = 2
    for c in (total_orange, total_enterprises, co5):
        ws.column_dimensions[get_column_letter(c)].outlineLevel = 1
    ws.column_dimensions[get_column_letter(grandtotal)].outlineLevel = 0
    ws.sheet_properties.outlinePr.summaryRight = True


def snapshot_column_row_hidden(wb):
    """Captures which columns/rows the user has manually collapsed in Excel (Data > Group, which
    openpyxl records as dimension.hidden = True) on any sheet of an existing workbook. A full
    rebuild regenerates every sheet from scratch and has no way to know this display preference on
    its own, so it must be captured before the rebuild and reapplied after - see
    restore_column_row_hidden. Returns {sheet_name: {"cols": [letters], "rows": [numbers] OR
    {"outline_level_gte": K}}}.

    Rows on the GL-level schedule sheets (PL/BS GL Detail, MOM PL/BS GL Detail) are sorted by
    magnitude and re-laid-out on every rebuild, so a literal row-number list captured today can
    point at completely different ledgers after the next Tally refresh - collapsing "individual
    GL rows" would silently degrade into collapsing an arbitrary 344-ish rows while ~60 detail
    rows that should be hidden stay visible (hit once in practice). Whenever a sheet's hidden-row
    set exactly equals "every row at outline level >= K" for some K, that's a native Excel
    collapse-button click, not an arbitrary hand-pick - store it as that RULE instead of a row
    list, so restore_column_row_hidden can re-derive the correct rows from the fresh rebuild's own
    (correctly recomputed) outline levels rather than trusting stale row numbers. Columns don't
    have this problem (company/total column positions are structurally fixed, never
    magnitude-sorted), so column hides are always stored as a literal letter list."""
    snap = {}
    for name in wb.sheetnames:
        ws = wb[name]
        cols = [c for c, dim in ws.column_dimensions.items() if dim.hidden]
        hidden_rows = set(r for r, dim in ws.row_dimensions.items() if dim.hidden)
        rows_state = list(hidden_rows)
        if hidden_rows:
            for level in range(1, 5):
                at_or_above = set(r for r, dim in ws.row_dimensions.items() if (dim.outlineLevel or 0) >= level)
                if at_or_above and at_or_above == hidden_rows:
                    rows_state = {"outline_level_gte": level}
                    break
        if cols or hidden_rows:
            snap[name] = {"cols": cols, "rows": rows_state}
    return snap


def restore_column_row_hidden(wb, snap):
    """Reapplies a snapshot from snapshot_column_row_hidden onto a freshly rebuilt workbook. Skips
    any sheet/column/row that no longer exists rather than erroring - a rebuild can legitimately
    change row counts (new ledgers, a refreshed period, etc.). A rows state of
    {"outline_level_gte": K} is re-derived from THIS workbook's own current outline levels (set
    moments ago by this same rebuild) rather than replayed as stale row numbers - see
    snapshot_column_row_hidden docstring for why that matters."""
    for name, state in snap.items():
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        for c in state["cols"]:
            ws.column_dimensions[c].hidden = True
        rows_state = state["rows"]
        if isinstance(rows_state, dict):
            level = rows_state["outline_level_gte"]
            for r, dim in ws.row_dimensions.items():
                if (dim.outlineLevel or 0) >= level:
                    dim.hidden = True
        else:
            for r in rows_state:
                if r <= ws.max_row:
                    ws.row_dimensions[r].hidden = True


GROUPING_SNAPSHOT_SUFFIX = ".grouping_snapshot.json"

BASE_NAME = "PL_BS_GL_Detail"
TIMESTAMP_FORMAT = "%Y-%m-%d_%H%M"


def find_latest_pl_bs_file(files_dir, exclude=None):
    """Every refresh produces a new dated PL_BS_GL_Detail_<timestamp>.xlsx (no fixed filename -
    see SKILL.md 'File naming and version history'). This finds the most recent one currently
    sitting in Files/ (i.e. not yet moved to the Old Version of PL BS archive)."""
    candidates = sorted(files_dir.glob(f"{BASE_NAME}_*.xlsx"))
    if exclude is not None:
        exclude = Path(exclude).resolve()
        candidates = [c for c in candidates if c.resolve() != exclude]
    return candidates[-1] if candidates else None


def grouping_snapshot_path(output_path):
    # Fixed name regardless of the dated output filename - this is transient handoff state between
    # this script and build_mom_gl_detail.py within a single run, not a user-facing deliverable, so
    # it deliberately does NOT get a timestamp (that would leave one orphaned JSON file per refresh).
    return output_path.parent / (BASE_NAME + GROUPING_SNAPSHOT_SUFFIX)


def save_grouping_snapshot(snap, output_path):
    """Persists a hidden-column/row snapshot to a JSON sidecar next to output_path. Needed because
    build_gl_detail_sidebyside.py's own rebuild wipes the whole workbook (a brand new
    openpyxl.Workbook()), including the MOM/Cross Check sheets it never touches - by the time
    build_mom_gl_detail.py runs afterwards and deletes+recreates those sheets, their original
    hidden state would already be gone unless it was captured here, in THIS script's run, before
    the overwrite, and handed off via this file rather than in-memory."""
    grouping_snapshot_path(output_path).write_text(json.dumps(snap), encoding="utf-8")


def load_grouping_snapshot(output_path):
    path = grouping_snapshot_path(output_path)
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def write_sheet_header(ws, title: str):
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 26
    ws.column_dimensions["D"].width = 40
    for col_letter in "EFHIK":
        ws.column_dimensions[col_letter].width = 15
    for col_letter in "GJL":
        ws.column_dimensions[col_letter].width = 16
    apply_family_column_outline(ws, COL_CO1, COL_CO2, COL_TOTAL_ORANGE, COL_CO3, COL_CO4,
                                 COL_TOTAL_ENTERPRISES, COL_CO5, COL_GRANDTOTAL)

    ws.sheet_properties.outlinePr.summaryBelow = True

    row = 1
    ws.cell(row, 1, value=title).font = TITLE_FONT
    row += 1
    ws.cell(row, 1, value=f"Period: {PERIOD_LABEL}").font = NORMAL_FONT
    row += 2

    header_row = row
    ws.cell(row, COL_SECTION, value="Statement Section")
    ws.cell(row, COL_GROUPING, value="Grouping")
    ws.cell(row, COL_SUBGROUPING, value="Sub Grouping")
    ws.cell(row, COL_GLNAME, value="GL Name")
    ws.cell(row, COL_CO1, value=CO1[0])
    ws.cell(row, COL_CO2, value=CO2[0])
    ws.cell(row, COL_TOTAL_ORANGE, value="Total of Orange O Tec")
    ws.cell(row, COL_CO3, value=CO3[0])
    ws.cell(row, COL_CO4, value=CO4[0])
    ws.cell(row, COL_TOTAL_ENTERPRISES, value="Total of Enterprises")
    ws.cell(row, COL_CO5, value=CO5[0])
    ws.cell(row, COL_GRANDTOTAL, value="Grand Total")
    for c in range(1, COL_GRANDTOTAL + 1):
        cell = ws.cell(row, c)
        cell.font = HEADER_FONT
        cell.fill = TOTALCOL_FILL if c in (COL_TOTAL_ORANGE, COL_TOTAL_ENTERPRISES, COL_GRANDTOTAL) else HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="bottom")
    ws.row_dimensions[row].height = 60
    ws.freeze_panes = f"E{header_row + 1}"
    return header_row + 1


def sum_formula_dict(row_lists):
    """row_lists: list of row numbers (contiguous range or scattered rows) to SUM per company
    column. Accepts either a (first_row, last_row) contiguous range or a list of individual row
    numbers - both render as a single Excel SUM() so every subtotal is a live formula over its
    child rows, never a pre-computed number."""
    if isinstance(row_lists, tuple):
        first_row, last_row = row_lists
        return {col: f"=SUM({l}{first_row}:{l}{last_row})" for col, l in PLAIN_COL_LETTERS.items()}
    return {col: "=SUM(" + ",".join(f"{l}{r}" for r in row_lists) + ")" for col, l in PLAIN_COL_LETTERS.items()}


def signed_formula_dict(rows_with_signs):
    """rows_with_signs: [(row_number, +1 or -1), ...]. Renders '=A10-A20' style per company
    column. Needed for the MIXED memo sections (Branch, Related Party) whose income leg and
    expense leg must NET rather than add once expenses are shown as positive magnitudes - a plain
    SUM() there would report roughly double the gross flow instead of ~0."""
    out = {}
    for col, l in PLAIN_COL_LETTERS.items():
        parts = []
        for i, (r, sign) in enumerate(rows_with_signs):
            op = "-" if sign < 0 else ("+" if i else "")
            parts.append(f"{op}{l}{r}")
        out[col] = "=" + "".join(parts) if parts else "=0"
    return out


def write_bucket_block(ws, row, display_label, bucket_rows_by_company, base_level):
    """Writes one Grouping bucket (GL detail -> Sub Grouping subtotal -> bucket TOTAL). Every
    subtotal above the individual GL rows is written as a live Excel SUM() formula over its own
    child rows, not a pre-computed number, so pressing F2 on any subtotal shows its derivation.
    bucket_rows_by_company: {company_key: [(sub_grp, gl_name, bal), ...]}. Returns
    (row, group_col_sums[5], bucket_total_row, subgroup_rows) - group_col_sums is still computed
    in Python purely to decide sort order and let callers combine bucket/section totals into
    higher-level +/- formulas (e.g. Gross Profit); it is never written into a cell directly.
    bucket_total_row is the row number of this bucket's own "TOTAL: <label>" row (None if the
    bucket had no data). subgroup_rows maps each Tally Sub Grouping name within this bucket to
    the row number of its own "Subtotal: <name>" row, so callers needing finer-than-bucket detail
    (e.g. an Indirect Expenses breakdown in PL Summary) can link straight to it."""
    line_values = {}
    line_order = []
    for idx, (_, key) in enumerate(ALL_COMPANIES):
        for sub_grp, gl_name, bal in bucket_rows_by_company.get(key, []):
            k = (sub_grp, gl_name)
            if k not in line_values:
                line_values[k] = [0.0] * 5
                line_order.append(k)
            line_values[k][idx] += bal

    if not line_order:
        return row, [0.0] * 5, None, {}

    subgroup_lines = {}
    subgroup_totals = {}
    subgroup_order = []
    for (sub_grp, gl_name) in line_order:
        if sub_grp not in subgroup_lines:
            subgroup_lines[sub_grp] = []
            subgroup_totals[sub_grp] = [0.0] * 5
            subgroup_order.append(sub_grp)
        vals = line_values[(sub_grp, gl_name)]
        subgroup_lines[sub_grp].append((gl_name, vals))
        for i in range(5):
            subgroup_totals[sub_grp][i] += vals[i]

    subgroup_order.sort(key=lambda sg: abs(col_sum(subgroup_totals[sg])), reverse=True)

    group_col_sums = [0.0] * 5
    group_start_row = row
    subtotal_rows = []
    subgroup_rows = {}
    for sub_grp in subgroup_order:
        lines = subgroup_lines[sub_grp]
        lines.sort(key=lambda item: abs(col_sum(item[1])), reverse=True)

        subgroup_start_row = row
        first_in_subgroup = True
        for gl_name, vals in lines:
            ws.cell(row, COL_SUBGROUPING, value=sub_grp if first_in_subgroup else None).font = SUBGROUP_FONT
            ws.cell(row, COL_GLNAME, value=gl_name).font = NORMAL_FONT
            write_value_cells(ws, row, plain_dict(vals), NORMAL_FONT, skip_zero=True)
            ws.row_dimensions[row].outlineLevel = base_level + 2
            first_in_subgroup = False
            row += 1
        subgroup_end_row = row - 1

        ws.cell(row, COL_GLNAME, value=f"Subtotal: {sub_grp}").font = SUBTOTAL_SUBGROUP_FONT
        write_value_cells(ws, row, sum_formula_dict((subgroup_start_row, subgroup_end_row)),
                          SUBTOTAL_SUBGROUP_FONT, fill=SUBTOTAL_SUBGROUP_FILL)
        ws.row_dimensions[row].outlineLevel = base_level + 1
        subtotal_rows.append(row)
        subgroup_rows[sub_grp] = row
        row += 1

        for i in range(5):
            group_col_sums[i] += subgroup_totals[sub_grp][i]

    ws.cell(group_start_row, COL_GROUPING, value=display_label).font = GROUP_FONT
    ws.cell(row, COL_GROUPING, value=display_label).font = GROUP_FONT
    ws.cell(row, COL_GLNAME, value=f"TOTAL: {display_label}").font = SUBTOTAL_FONT
    write_value_cells(ws, row, sum_formula_dict(subtotal_rows), SUBTOTAL_FONT, fill=SUBTOTAL_FILL)
    ws.row_dimensions[row].outlineLevel = base_level
    bucket_total_row = row
    row += 1

    return row, group_col_sums, bucket_total_row, subgroup_rows


def write_section(ws, row, section_label, buckets, total_label=None, base_level=1, section_style="normal",
                   bucket_signs=None):
    """buckets: list of (display_label, rows_by_company). Writes a level-0 section header,
    each bucket nested at base_level, and a level-0 section total row (a live SUM() of the
    bucket TOTAL rows, not a pre-computed number). Returns (row, totals[5], bucket_rows,
    detail_rows) where bucket_rows maps each bucket's display_label to the row number of its own
    "TOTAL: <label>" row, so callers (e.g. the Ratio Analysis sheet) can link directly to a
    specific bucket (Sundry Debtors, Secured Loans, ...) without needing the whole section
    total. detail_rows maps each bucket's display_label to its own {sub_grp: row} dict (one
    level finer than bucket_rows - Tally's Sub Grouping), for callers needing a full sub-group
    breakdown of one specific bucket (e.g. Indirect Expenses in PL Summary). section_style:
    "normal" for the main P&L/BS flow, "adjustment" for memo/disclosure sections (Branch
    Transactions, Related Party Transactions) that get a distinct neutral grey treatment so they
    read as supplementary, not part of the main waterfall."""
    total_label = total_label or f"TOTAL {section_label}"
    header_font = SECTION_FONT if section_style == "normal" else ADJUSTMENT_FONT
    header_fill = SECTION_FILL if section_style == "normal" else ADJUSTMENT_FILL
    total_font = SECTION_TOTAL_FONT if section_style == "normal" else ADJUSTMENT_TOTAL_FONT
    total_fill = SECTION_TOTAL_FILL if section_style == "normal" else ADJUSTMENT_TOTAL_FILL

    ws.cell(row, COL_SECTION, value=section_label).font = header_font
    for c in range(1, COL_GRANDTOTAL + 1):
        ws.cell(row, c).fill = header_fill
    ws.row_dimensions[row].outlineLevel = 0
    row += 1

    section_totals = [0.0] * 5
    bucket_rows = {}
    detail_rows = {}
    bucket_total_rows = []
    for display_label, rows_by_company in buckets:
        row, bucket_totals, bucket_total_row, subgroup_rows = write_bucket_block(
            ws, row, display_label, rows_by_company, base_level)
        if bucket_total_row is not None:
            bucket_rows[display_label] = bucket_total_row
            bucket_total_rows.append(bucket_total_row)
        detail_rows[display_label] = subgroup_rows
        for i in range(5):
            section_totals[i] += bucket_totals[i]

    ws.cell(row, COL_GLNAME, value=total_label).font = total_font
    if bucket_total_rows and bucket_signs:
        signed = [(bucket_rows[lbl], bucket_signs.get(lbl, 1))
                  for lbl in bucket_rows if bucket_rows[lbl] in bucket_total_rows]
        write_value_cells(ws, row, signed_formula_dict(signed), total_font, fill=total_fill)
    elif bucket_total_rows:
        write_value_cells(ws, row, sum_formula_dict(bucket_total_rows), total_font, fill=total_fill)
    else:
        write_value_cells(ws, row, plain_dict([0.0] * 5), total_font, fill=total_fill)
    ws.row_dimensions[row].outlineLevel = 0
    row += 1

    return row, section_totals, bucket_rows, detail_rows


def write_computed_row(ws, row, label, component_rows, font, fill):
    """Writes a computed headline row (Gross Profit, EBITDA, PBT, PAT, Reconciliation Gap, ...)
    as a live Excel formula combining other rows' company columns - never an independently
    computed number. component_rows: list of (row_number, sign) with sign '+' or '-'."""
    ws.cell(row, COL_SECTION, value=label).font = font
    formulas = {}
    for col, l in PLAIN_COL_LETTERS.items():
        parts = []
        for i, (ref_row, sign) in enumerate(component_rows):
            term = f"{l}{ref_row}"
            if i == 0:
                parts.append(f"-{term}" if sign == "-" else term)
            else:
                parts.append(f"{'-' if sign == '-' else '+'}{term}")
        formulas[col] = "=" + "".join(parts)
    write_value_cells(ws, row, formulas, font, fill=fill)
    ws.row_dimensions[row].outlineLevel = 0
    return row + 1


def write_literal_row(ws, row, label, vals5, font, fill):
    """Writes a literal (non-formula) placeholder row - used only where the figure is genuinely
    zero this quarter with nothing to sum or combine (Depreciation/Tax: no provision booked)."""
    ws.cell(row, COL_SECTION, value=label).font = font
    write_value_cells(ws, row, plain_dict(vals5), font, fill=fill)
    ws.row_dimensions[row].outlineLevel = 0
    return row + 1


def build_pl_sections(all_rows_by_company):
    """Classifies every ledger row into (section, bucket) and returns
    {section: {bucket: {company_key: [(sub_grp, gl_name, bal), ...]}}}."""
    data = {}
    for _, key in ALL_COMPANIES:
        for pg, sub_grp, gl_name, bal in all_rows_by_company[key]:
            section, bucket = classify_pl_row(pg, sub_grp, gl_name)
            data.setdefault(section, {}).setdefault(bucket, {}).setdefault(key, []).append((sub_grp, gl_name, bal))
    return apply_display_signs(data)


def build_related_party_sections(all_rows_by_company):
    """Separately extracts Related Party Sales/Purchase rows for the memo disclosure below PAT.
    These rows are NOT removed from the main TURNOVER/COST OF SALES sections above - Gross
    Profit, EBITDA, PBT and PAT are computed on the full (related-party-inclusive) basis, and
    this section exists purely to compute the final 'PAT net of Related Party' memo row."""
    data = {}
    for _, key in ALL_COMPANIES:
        for pg, sub_grp, gl_name, bal in all_rows_by_company[key]:
            if not is_related_party_row(pg, sub_grp):
                continue
            bucket = "Related Party Sales" if pg == "Sales Accounts" else "Related Party Purchase"
            data.setdefault(bucket, {}).setdefault(key, []).append((sub_grp, gl_name, bal))
    return apply_display_signs({"RELATED PARTY": data})


def build_bs_sections(all_rows_by_company):
    data = {}
    for _, key in ALL_COMPANIES:
        for pg, sub_grp, gl_name, bal in all_rows_by_company[key]:
            section, bucket = classify_bs_row(pg)
            data.setdefault(section, {}).setdefault(bucket, {}).setdefault(key, []).append((sub_grp, gl_name, bal))
    return apply_display_signs(data)


def buckets_for(section_data, section_key, bucket_order, label_map=None):
    """Returns [(display_label, rows_by_company), ...] in bucket_order. Where two different
    primary_group values map to the same display_label via label_map (e.g. 'Bank Accounts' and
    Colorix's 'Bank Account' both display as 'Bank Accounts'), their company data is merged into
    a single bucket/row rather than written as two separate rows with the same label - otherwise
    a later lookup by display_label (e.g. from the Ratio Analysis sheet) would only ever see one
    of the two and silently miss the other companies' figures."""
    section_buckets = section_data.get(section_key, {})
    merged = {}  # display_label -> rows_by_company (merged across raw pg keys)
    order = []
    for pg in bucket_order:
        if pg not in section_buckets:
            continue
        label = label_map.get(pg, pg) if label_map else pg
        if label not in merged:
            merged[label] = {}
            order.append(label)
        for company_key, rows in section_buckets[pg].items():
            merged[label].setdefault(company_key, []).extend(rows)
    return [(label, merged[label]) for label in order]


def write_pl_sheet(ws):
    """Writes the PL schedule (full GL detail) and returns a {line_item_label: row_number}
    map so the Summary sheet can link its cells directly to these rows."""
    row = write_sheet_header(ws, "Profit & Loss A/c - GL Detail Schedule (side-by-side)")
    all_rows_by_company = {key: load_company_rows("pl", key, "pl") for _, key in ALL_COMPANIES}
    sections = build_pl_sections(all_rows_by_company)
    row_map = {}
    section_buckets = {}
    section_detail = {}

    row, turnover, section_buckets["TURNOVER"], section_detail["TURNOVER"] = write_section(
        ws, row, "TURNOVER", buckets_for(sections, "TURNOVER", PL_TURNOVER_BUCKETS),
        total_label="TOTAL TURNOVER")
    row_map["TOTAL TURNOVER"] = row - 1
    row, cos, section_buckets["COST OF SALES"], section_detail["COST OF SALES"] = write_section(
        ws, row, "LESS: COST OF SALES", buckets_for(sections, "COST OF SALES", PL_COST_OF_SALES_BUCKETS),
        total_label="TOTAL COST OF SALES")
    row_map["TOTAL COST OF SALES"] = row - 1

    row, branch_totals, section_buckets["BRANCH TRANSACTIONS"], section_detail["BRANCH TRANSACTIONS"] = write_section(
        ws, row, "BRANCH TRANSACTIONS (Inter-company - shown separately for visibility, included in Gross Profit below)",
        buckets_for(sections, "BRANCH TRANSACTIONS", PL_BRANCH_BUCKETS),
        total_label="NET BRANCH TRANSACTIONS (included in Gross Profit below, matches each company's own Tally P&L)",
        section_style="adjustment",
        bucket_signs={"Branch Sales": 1, "Branch Purchase": -1})
    row_map["NET BRANCH TRANSACTIONS"] = row - 1

    row = write_computed_row(ws, row, "GROSS PROFIT",
                              [(row_map["TOTAL TURNOVER"], "+"), (row_map["TOTAL COST OF SALES"], "-"),
                               (row_map["NET BRANCH TRANSACTIONS"], "+")],
                              GP_FONT, GP_FILL)
    row_map["GROSS PROFIT"] = row - 1

    row, other_income, section_buckets["OTHER INCOME"], section_detail["OTHER INCOME"] = write_section(
        ws, row, "ADD: OTHER INCOME", buckets_for(sections, "OTHER INCOME", PL_OTHER_INCOME_BUCKETS),
        total_label="TOTAL OTHER INCOME")
    row_map["TOTAL OTHER INCOME"] = row - 1
    row, opex, section_buckets["OPERATING EXPENSES"], section_detail["OPERATING EXPENSES"] = write_section(
        ws, row, "LESS: OPERATING EXPENSES",
        buckets_for(sections, "OPERATING EXPENSES", PL_OPERATING_EXPENSE_BUCKETS),
        total_label="TOTAL OPERATING EXPENSES")
    row_map["TOTAL OPERATING EXPENSES"] = row - 1
    row = write_computed_row(ws, row, "EBITDA",
                              [(row_map["GROSS PROFIT"], "+"), (row_map["TOTAL OTHER INCOME"], "+"),
                               (row_map["TOTAL OPERATING EXPENSES"], "-")],
                              EBITDA_FONT, EBITDA_FILL)
    row_map["EBITDA"] = row - 1

    row, finance_cost, section_buckets["FINANCE COST"], section_detail["FINANCE COST"] = write_section(
        ws, row, "LESS: FINANCE COST", buckets_for(sections, "FINANCE COST", PL_FINANCE_COST_BUCKETS),
        total_label="TOTAL FINANCE COST")
    row_map["TOTAL FINANCE COST"] = row - 1
    row = write_literal_row(ws, row, "LESS: DEPRECIATION (not booked this period)", [0.0] * 5,
                             ZERO_FONT, ZERO_FILL)
    row_map["LESS: DEPRECIATION"] = row - 1
    row = write_computed_row(ws, row, "PBT (PROFIT BEFORE TAX)",
                              [(row_map["EBITDA"], "+"), (row_map["TOTAL FINANCE COST"], "-"),
                               (row_map["LESS: DEPRECIATION"], "-")],
                              PBT_FONT, PBT_FILL)
    row_map["PBT"] = row - 1

    row = write_literal_row(ws, row, "LESS: TAX (not booked this period)", [0.0] * 5, ZERO_FONT, ZERO_FILL)
    row_map["LESS: TAX"] = row - 1
    row = write_computed_row(ws, row, "PAT (NET PROFIT)",
                              [(row_map["PBT"], "+"), (row_map["LESS: TAX"], "-")],
                              PAT_FONT, PAT_FILL)
    row_map["PAT"] = row - 1

    rp_sections = build_related_party_sections(all_rows_by_company)
    row, rp_totals, section_buckets["RELATED PARTY TRANSACTIONS"], section_detail["RELATED PARTY TRANSACTIONS"] = write_section(
        ws, row, "RELATED PARTY TRANSACTIONS (included in PAT above - shown for reference)",
        buckets_for(rp_sections, "RELATED PARTY", PL_RELATED_PARTY_BUCKETS),
        total_label="NET RELATED PARTY TRANSACTIONS",
        section_style="adjustment",
        bucket_signs={"Related Party Sales": 1, "Related Party Purchase": -1})
    row_map["NET RELATED PARTY TRANSACTIONS"] = row - 1
    row = write_computed_row(ws, row, "PAT (NET OF RELATED PARTY TRANSACTIONS)",
                              [(row_map["PAT"], "+"), (row_map["NET RELATED PARTY TRANSACTIONS"], "-")],
                              GRANDTOTAL_FONT, GRANDTOTAL_FILL)
    row_map["PAT (NET OF RELATED PARTY)"] = row - 1

    # Valuation-related provisions bridge (permanent, 2026-07-31 onward): the ledgers Tally files
    # directly under "Direct Expenses" with no finer Sub Grouping of their own (e.g. "PROVISION
    # FOR IMPORT-FOC(PL)", "INVENTORY OVERHEAD COST") are consistently the most volatile,
    # judgment-heavy line items quarter to quarter - see Profit_Change_Reconciliation reports,
    # where they are repeatedly the top PAT drivers. Called out here as a standalone bridge so a
    # reader sees, at a glance, how much of reported PAT depends on them, without hunting through
    # the ordinary Cost of Sales detail above. This replaces an earlier one-off patch (found and
    # removed 2026-07-31) that added these items back and then immediately subtracted them again,
    # so "Profit After Valuation" always collapsed back to PBT regardless of these items' values -
    # "Profit Before Valuation" here is genuinely PAT EXCLUDING this bucket, so the bridge is real.
    valuation_rows_by_company = {}
    for _, key in ALL_COMPANIES:
        entries = [(sub_grp, gl_name, bal) for sub_grp, gl_name, bal in
                   sections.get("COST OF SALES", {}).get("Direct Expenses", {}).get(key, [])
                   if sub_grp == "Direct Expenses"]
        if entries:
            valuation_rows_by_company[key] = entries

    if valuation_rows_by_company:
        row += 1
        ws.cell(row, COL_SECTION,
                value="VALUATION-RELATED PROVISIONS (most volatile Direct Expenses items - bridge to reported PAT)").font = ADJUSTMENT_FONT
        for c in range(1, COL_GRANDTOTAL + 1):
            ws.cell(row, c).fill = ADJUSTMENT_FILL
        ws.row_dimensions[row].outlineLevel = 0
        row += 1

        row, _, valuation_total_row, _ = write_bucket_block(
            ws, row, "Valuation-related Provisions", valuation_rows_by_company, base_level=1)

        row = write_computed_row(ws, row, "Profit Before Valuation (excl. items above)",
                                  [(row_map["PAT (NET OF RELATED PARTY)"], "+"), (valuation_total_row, "-")],
                                  GRANDTOTAL_FONT, ADJUSTMENT_FILL)
        row_map["PROFIT BEFORE VALUATION"] = row - 1

        row = write_computed_row(ws, row, "Profit After Valuation (= Reported PAT)",
                                  [(row_map["PAT (NET OF RELATED PARTY)"], "+")],
                                  GRANDTOTAL_FONT, GRANDTOTAL_FILL)
        row_map["PROFIT AFTER VALUATION"] = row - 1

    add_grand_total_pct_column(ws, COL_GRANDTOTAL, row_map["TOTAL TURNOVER"],
                               first_row=6, last_row=row - 1, header_row=5)

    return row_map, section_buckets, section_detail


# "% of Sales" on the quarter sheets. These are laid out by COMPANY (one period only), so unlike
# the MOM sheets there is a single percentage column, against the Grand Total column - agreed
# with the user 2026-08-04 ("Grand Total % only") to keep the sheets readable rather than adding
# eight percentage columns per sheet.
PCT_FORMAT = "0.0%"


def add_grand_total_pct_column(ws, gt_col, turnover_row, first_row, last_row,
                               header_row, label="% of Sales"):
    """Appends one '% of Sales' column immediately right of the Grand Total column. Every cell is
    a live IFERROR-wrapped formula over that sheet's own Grand Total column, so it can never
    disagree with the figures beside it."""
    col = gt_col + 1
    letter = get_column_letter(col)
    src = get_column_letter(gt_col)
    ws.column_dimensions[letter].width = 13
    head = ws.cell(header_row, col, value=label)
    head.font = HEADER_FONT
    head.fill = TOTALCOL_FILL
    head.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    for r in range(first_row, last_row + 1):
        if ws.cell(r, gt_col).value in (None, ""):
            continue
        c = ws.cell(r, col, value=f'=IFERROR({src}{r}/{src}${turnover_row},"")')
        c.number_format = PCT_FORMAT
        c.font = NORMAL_FONT
    return col


# ---------------------------------------------------------------------------
# Balance Sheet related-party / branch memo sections (2026-08-04 onward, user request:
# "show all related party ledgers in BS below as done in PL")
# ---------------------------------------------------------------------------
# The P&L has carried Branch Transactions and Related Party Transactions as two separate memo
# blocks for a while; the Balance Sheet had neither, so the closing BALANCES with group entities
# were invisible even though the flows were disclosed.
#
# DETECTION NEEDS BOTH TESTS - a sub-group-name rule alone silently misses about a third of them.
# Verified against live data 2026-08-04: 16 ledgers sit in sub-groups that say "RELATED"
# (BALANCE WITH RELATED PARTY, CREDITORS FOR RELATED, BALANCE WITH RELATED PARTY(Debtors)), but
# another 11 sit in perfectly ordinary groups - Unsecured Loans, Loans & Advances (Asset),
# Sundry Debtors, Current Assets, Branch / Divisions - and are identifiable only by the LEDGER
# NAME carrying a group company's name (e.g. "ORANGE O TEC PRIVATE LIMITED-USL-SURAT" under
# plain "Unsecured Loans"). Matching on name as well is what catches those.
#
# Branch is kept as its own block rather than folded in, mirroring the P&L: the intra-group
# Surat<->Noida pair alone is ~Rs 58cr each way and would otherwise swamp the genuine
# related-party balances and make the memo total meaningless.
BS_RELATED_PARTY_SUBGROUP_HINT = "RELATED"
BS_BRANCH_PRIMARY_GROUPS = {"Branch / Divisions"}
GROUP_ENTITY_NAME_HINTS = ("ORANGE O TEC", "COLORIX")


def is_group_entity_ledger(ledger_name: str) -> bool:
    u = (ledger_name or "").upper()
    return any(h in u for h in GROUP_ENTITY_NAME_HINTS)


# Maps a related-party ledger to the GROUP COMPANY it represents, so a balance can be matched
# against the mirror entry in that company's own books. Order matters: "ORANGE O TEC ENTERPRISES
# PRIVATE LIMITED-NOIDA" contains all three of ORANGE O TEC / ENTERPRISES / NOIDA, so Colorix is
# tested first, then Enterprises (with its own NOIDA qualifier), and only then plain Orange O Tec.
# SURAT denotes the Orange O Tec / Enterprises head office, i.e. the non-NOIDA entity, so it needs
# no rule of its own - it falls through to the correct default.
def counterparty_of(ledger_name: str):
    u = (ledger_name or "").upper()
    if "COLORIX" in u:
        return "COLORIX"
    if "ENTERPRISES" in u:
        return "ENTERPRISES_NOIDA" if "NOIDA" in u else "ENTERPRISES_PVT_LTD"
    if "ORANGE O TEC" in u:
        return "PRIVATE_LIMITED_NOIDA" if "NOIDA" in u else "ORANGE_O_TEC_PRIVATE_LIMITED"
    return None

def classify_bs_memo(primary_group: str, sub_group: str, ledger_name: str):
    """Returns 'BRANCH', 'RELATED' or None for the Balance Sheet memo sections."""
    if primary_group in BS_BRANCH_PRIMARY_GROUPS:
        return "BRANCH"
    if BS_RELATED_PARTY_SUBGROUP_HINT in (sub_group or "").upper():
        return "RELATED"
    if is_group_entity_ledger(ledger_name):
        return "RELATED"
    return None


def build_bs_memo_sections(all_rows_by_company):
    """{'BS BRANCH': {bucket: {company: [...]}}, 'BS RELATED PARTY': {...}} - rows stay in their
    normal Balance Sheet position too; these blocks are disclosure, not a reclassification, so
    they are NOT subtracted from Total Assets / Total Liabilities."""
    branch, related = {}, {}
    for _, key in ALL_COMPANIES:
        for pg, sub_grp, gl_name, bal in all_rows_by_company[key]:
            kind = classify_bs_memo(pg, sub_grp, gl_name)
            if kind is None:
                continue
            target = branch if kind == "BRANCH" else related
            target.setdefault(pg, {}).setdefault(key, []).append((sub_grp, gl_name, bal))
    return {"BS BRANCH": branch, "BS RELATED PARTY": related}


def write_bs_sheet(ws):
    """Writes the BS schedule (full GL detail) and returns a {line_item_label: row_number}
    map so the Summary and Ratio Analysis sheets can link their cells directly to these rows.
    Bucket-level rows (e.g. 'Sundry Debtors', 'Secured Loans') are included alongside the
    section totals ('TOTAL LIABILITIES' etc.) for ratios that need a specific line item."""
    row = write_sheet_header(ws, "Balance Sheet - GL Detail Schedule (side-by-side)")
    all_rows_by_company = {key: load_company_rows("bs", key, "bs") for _, key in ALL_COMPANIES}
    sections = build_bs_sections(all_rows_by_company)
    row_map = {}

    row, liabilities, liability_bucket_rows, liability_detail = write_section(
        ws, row, "LIABILITIES (Sources of Funds)",
        buckets_for(sections, "LIABILITIES", BS_LIABILITY_ORDER, BS_GROUP_LABEL),
        total_label="TOTAL LIABILITIES")
    row_map["TOTAL LIABILITIES"] = row - 1
    row_map.update(liability_bucket_rows)
    row, assets, asset_bucket_rows, asset_detail = write_section(
        ws, row, "ASSETS (Application of Funds)",
        buckets_for(sections, "ASSETS", BS_ASSET_ORDER, BS_GROUP_LABEL),
        total_label="TOTAL ASSETS")
    row_map["TOTAL ASSETS"] = row - 1
    row_map.update(asset_bucket_rows)
    row = write_computed_row(ws, row, "RECONCILIATION GAP (Liabilities - Assets; ~0 expected)",
                              [(row_map["TOTAL LIABILITIES"], "+"), (row_map["TOTAL ASSETS"], "-")],
                              GAP_FONT, GAP_FILL)
    row_map["RECONCILIATION GAP"] = row - 1

    # Memo disclosure blocks - see build_bs_memo_sections. These mirror the P&L's Branch /
    # Related Party sections and are DISCLOSURE ONLY: the same ledgers remain counted in Total
    # Assets / Total Liabilities above, so the Reconciliation Gap is unaffected by them.
    memo = build_bs_memo_sections(all_rows_by_company)
    for section_key, title, total_label in (
        ("BS BRANCH",
         "BRANCH BALANCES (inter-branch - included in Assets/Liabilities above, shown for reference)",
         "NET BRANCH BALANCES"),
        ("BS RELATED PARTY",
         "RELATED PARTY BALANCES (included in Assets/Liabilities above, shown for reference)",
         "NET RELATED PARTY BALANCES"),
    ):
        data = memo.get(section_key, {})
        if not data:
            continue
        order = [pg for pg in BS_LIABILITY_ORDER + BS_ASSET_ORDER if pg in data]
        order += [pg for pg in data if pg not in order]
        row += 1
        row, _, bucket_rows_memo, _ = write_section(
            ws, row, title, [(pg, data[pg]) for pg in order],
            total_label=total_label, section_style="adjustment")
        row_map[total_label] = row - 1

    section_buckets = {"LIABILITIES": liability_bucket_rows, "ASSETS": asset_bucket_rows}
    section_detail = {"LIABILITIES": liability_detail, "ASSETS": asset_detail}
    return row_map, section_buckets, section_detail


# --- Summary sheets: short, clean sheets whose cells are formulas linking straight into the
# schedule sheets above (PL GL Detail / BS GL Detail) - edit a number on the schedule and the
# summary recalculates. No values are independently recomputed here.

SUM_COL_LABEL = 1
SUM_COL_CO1 = 2
SUM_COL_CO2 = 3
SUM_COL_TOTAL_ORANGE = 4
SUM_COL_CO3 = 5
SUM_COL_CO4 = 6
SUM_COL_TOTAL_ENTERPRISES = 7
SUM_COL_CO5 = 8
SUM_COL_GRANDTOTAL = 9

# Maps each Summary value column to the Schedule sheet's column holding the same series.
SUMMARY_TO_SCHEDULE_COL = {
    SUM_COL_CO1: COL_CO1, SUM_COL_CO2: COL_CO2, SUM_COL_TOTAL_ORANGE: COL_TOTAL_ORANGE,
    SUM_COL_CO3: COL_CO3, SUM_COL_CO4: COL_CO4, SUM_COL_TOTAL_ENTERPRISES: COL_TOTAL_ENTERPRISES,
    SUM_COL_CO5: COL_CO5, SUM_COL_GRANDTOTAL: COL_GRANDTOTAL,
}

# The 8 company/total keys shared by every "8-column family block" in this workbook (Summary
# sheets, Ratio Analysis, MOM Ratio Analysis period blocks, ...), and how they map onto the
# Summary-style column numbers - reused by the Ratio Workings helpers below so the quarterly and
# monthly ratio sheets can build their column lists the same way.
RATIO_KEYS = ("CO1", "CO2", "TOTAL_ORANGE", "CO3", "CO4", "TOTAL_ENTERPRISES", "CO5", "GRANDTOTAL")
RATIO_KEY_TO_SUM_COL = {
    "CO1": SUM_COL_CO1, "CO2": SUM_COL_CO2, "TOTAL_ORANGE": SUM_COL_TOTAL_ORANGE,
    "CO3": SUM_COL_CO3, "CO4": SUM_COL_CO4, "TOTAL_ENTERPRISES": SUM_COL_TOTAL_ENTERPRISES,
    "CO5": SUM_COL_CO5, "GRANDTOTAL": SUM_COL_GRANDTOTAL,
}


def write_summary_header(ws, title: str):
    ws.column_dimensions["A"].width = 42
    for col_letter in "BCEFH":
        ws.column_dimensions[col_letter].width = 15
    for col_letter in "DGI":
        ws.column_dimensions[col_letter].width = 16
    apply_family_column_outline(ws, SUM_COL_CO1, SUM_COL_CO2, SUM_COL_TOTAL_ORANGE, SUM_COL_CO3,
                                 SUM_COL_CO4, SUM_COL_TOTAL_ENTERPRISES, SUM_COL_CO5, SUM_COL_GRANDTOTAL)

    row = 1
    ws.cell(row, 1, value=title).font = TITLE_FONT
    row += 1
    ws.cell(row, 1, value=f"Period: {PERIOD_LABEL}  (all figures linked live from the GL Detail schedule sheets)").font = NORMAL_FONT
    row += 2

    header_row = row
    ws.cell(row, SUM_COL_LABEL, value="Line Item")
    ws.cell(row, SUM_COL_CO1, value=CO1[0])
    ws.cell(row, SUM_COL_CO2, value=CO2[0])
    ws.cell(row, SUM_COL_TOTAL_ORANGE, value="Total of Orange O Tec")
    ws.cell(row, SUM_COL_CO3, value=CO3[0])
    ws.cell(row, SUM_COL_CO4, value=CO4[0])
    ws.cell(row, SUM_COL_TOTAL_ENTERPRISES, value="Total of Enterprises")
    ws.cell(row, SUM_COL_CO5, value=CO5[0])
    ws.cell(row, SUM_COL_GRANDTOTAL, value="Grand Total")
    for c in range(1, SUM_COL_GRANDTOTAL + 1):
        cell = ws.cell(row, c)
        cell.font = HEADER_FONT
        cell.fill = TOTALCOL_FILL if c in (SUM_COL_TOTAL_ORANGE, SUM_COL_TOTAL_ENTERPRISES, SUM_COL_GRANDTOTAL) else HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="bottom")
    ws.row_dimensions[row].height = 60
    ws.freeze_panes = f"B{header_row + 1}"
    return header_row + 1


PROFIT_LOSS_LABEL_FONT = Font(bold=True, size=11, name="Calibri", color="1F3864")
# NOTE: Excel's conditional-formatting engine renders the *background* color slot for a
# solid dxf fill, not the foreground slot used by normal cell fills - so both must be set here,
# otherwise the CF fill renders black (the unset bgColor default) instead of green/red.
PROFIT_FILL = PatternFill("solid", fgColor="2E7D32", bgColor="2E7D32")
LOSS_FILL = PatternFill("solid", fgColor="B71C1C", bgColor="B71C1C")
PROFIT_LOSS_VALUE_FONT = Font(bold=True, size=11, name="Calibri", color="FFFFFF")


def write_linked_row(ws, row, label, schedule_sheet_name, schedule_row, font, fill=None):
    """Writes one Summary row whose value cells are formulas referencing schedule_row on the
    schedule sheet (same column series), not independently computed values."""
    ws.cell(row, SUM_COL_LABEL, value=label).font = font
    for summary_col, schedule_col in SUMMARY_TO_SCHEDULE_COL.items():
        schedule_col_letter = get_column_letter(schedule_col)
        cell = ws.cell(row, summary_col, value=f"='{schedule_sheet_name}'!{schedule_col_letter}{schedule_row}")
        cell.number_format = NUM_FORMAT
        cell.font = font
    if fill:
        for c in range(1, SUM_COL_GRANDTOTAL + 1):
            ws.cell(row, c).fill = fill


def write_profit_loss_row(ws, row, label, schedule_sheet_name, schedule_row):
    """Writes a profit/loss milestone row (Gross Profit, EBITDA, PBT, PAT...) whose value cells
    are conditionally colored green when >= 0 (profit) and red when < 0 (loss) - the fill isn't
    a fixed color, it reacts live to whatever the linked formula evaluates to."""
    ws.cell(row, SUM_COL_LABEL, value=label).font = PROFIT_LOSS_LABEL_FONT
    for summary_col, schedule_col in SUMMARY_TO_SCHEDULE_COL.items():
        schedule_col_letter = get_column_letter(schedule_col)
        cell = ws.cell(row, summary_col, value=f"='{schedule_sheet_name}'!{schedule_col_letter}{schedule_row}")
        cell.number_format = NUM_FORMAT
        cell.font = PROFIT_LOSS_VALUE_FONT
    value_range = f"{get_column_letter(SUM_COL_CO1)}{row}:{get_column_letter(SUM_COL_GRANDTOTAL)}{row}"
    ws.conditional_formatting.add(
        value_range, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PROFIT_FILL))
    ws.conditional_formatting.add(
        value_range, CellIsRule(operator="lessThan", formula=["0"], fill=LOSS_FILL))


def write_grouping_header(ws, row, label):
    ws.cell(row, SUM_COL_LABEL, value=label).font = SECTION_FONT
    for c in range(1, SUM_COL_GRANDTOTAL + 1):
        ws.cell(row, c).fill = SECTION_FILL


# --- PL Summary / BS Summary get an extra "Sub Group" column (between Line Item and the value
# columns) so each grouping heading can list its bucket-level sub-groups (e.g. under TURNOVER:
# Sales Accounts, Direct Incomes), not just a single total row. Ratio Analysis has no sub-group
# concept and keeps the plain SUM_COL_* layout above.
SUM2_COL_LABEL = 1
SUM2_COL_SUBGROUP = 2
SUM2_COL_CO1 = 3
SUM2_COL_CO2 = 4
SUM2_COL_TOTAL_ORANGE = 5
SUM2_COL_CO3 = 6
SUM2_COL_CO4 = 7
SUM2_COL_TOTAL_ENTERPRISES = 8
SUM2_COL_CO5 = 9
SUM2_COL_GRANDTOTAL = 10

SUMMARY2_TO_SCHEDULE_COL = {
    SUM2_COL_CO1: COL_CO1, SUM2_COL_CO2: COL_CO2, SUM2_COL_TOTAL_ORANGE: COL_TOTAL_ORANGE,
    SUM2_COL_CO3: COL_CO3, SUM2_COL_CO4: COL_CO4, SUM2_COL_TOTAL_ENTERPRISES: COL_TOTAL_ENTERPRISES,
    SUM2_COL_CO5: COL_CO5, SUM2_COL_GRANDTOTAL: COL_GRANDTOTAL,
}


def write_summary_header2(ws, title: str):
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 24
    for col_letter in "CDFGI":
        ws.column_dimensions[col_letter].width = 15
    for col_letter in "EHJ":
        ws.column_dimensions[col_letter].width = 16
    apply_family_column_outline(ws, SUM2_COL_CO1, SUM2_COL_CO2, SUM2_COL_TOTAL_ORANGE, SUM2_COL_CO3,
                                 SUM2_COL_CO4, SUM2_COL_TOTAL_ENTERPRISES, SUM2_COL_CO5, SUM2_COL_GRANDTOTAL)

    row = 1
    ws.cell(row, 1, value=title).font = TITLE_FONT
    row += 1
    ws.cell(row, 1, value=f"Period: {PERIOD_LABEL}  (all figures linked live from the GL Detail schedule sheets)").font = NORMAL_FONT
    row += 2

    header_row = row
    ws.cell(row, SUM2_COL_LABEL, value="Line Item")
    ws.cell(row, SUM2_COL_SUBGROUP, value="Sub Group")
    ws.cell(row, SUM2_COL_CO1, value=CO1[0])
    ws.cell(row, SUM2_COL_CO2, value=CO2[0])
    ws.cell(row, SUM2_COL_TOTAL_ORANGE, value="Total of Orange O Tec")
    ws.cell(row, SUM2_COL_CO3, value=CO3[0])
    ws.cell(row, SUM2_COL_CO4, value=CO4[0])
    ws.cell(row, SUM2_COL_TOTAL_ENTERPRISES, value="Total of Enterprises")
    ws.cell(row, SUM2_COL_CO5, value=CO5[0])
    ws.cell(row, SUM2_COL_GRANDTOTAL, value="Grand Total")
    for c in range(1, SUM2_COL_GRANDTOTAL + 1):
        cell = ws.cell(row, c)
        cell.font = HEADER_FONT
        cell.fill = TOTALCOL_FILL if c in (SUM2_COL_TOTAL_ORANGE, SUM2_COL_TOTAL_ENTERPRISES, SUM2_COL_GRANDTOTAL) else HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="bottom")
    ws.row_dimensions[row].height = 60
    ws.freeze_panes = f"C{header_row + 1}"
    return header_row + 1


SUBGROUP_ROW_FONT = Font(italic=True, size=9, name="Calibri", color="44546A")


def write_linked_row2(ws, row, label, schedule_sheet_name, schedule_row, font, fill=None, subgroup=None):
    """Writes one PL/BS Summary row whose value cells are formulas referencing schedule_row on
    the schedule sheet. If subgroup is given, it is written into the Sub Group column instead of
    (or alongside) the Line Item label."""
    ws.cell(row, SUM2_COL_LABEL, value=label).font = font
    if subgroup:
        ws.cell(row, SUM2_COL_SUBGROUP, value=subgroup).font = font
    for summary_col, schedule_col in SUMMARY2_TO_SCHEDULE_COL.items():
        schedule_col_letter = get_column_letter(schedule_col)
        cell = ws.cell(row, summary_col, value=f"='{schedule_sheet_name}'!{schedule_col_letter}{schedule_row}")
        cell.number_format = NUM_FORMAT
        cell.font = font
    if fill:
        for c in range(1, SUM2_COL_GRANDTOTAL + 1):
            ws.cell(row, c).fill = fill


def write_subgroup_rows2(ws, row, schedule_sheet_name, bucket_rows):
    """Writes one row per sub-group in bucket_rows ({display_label: schedule_row}), each linked
    straight to that sub-group's own 'TOTAL: <label>' row on the schedule sheet. Returns the next
    free row."""
    for label, schedule_row in bucket_rows.items():
        write_linked_row2(ws, row, "", schedule_sheet_name, schedule_row, SUBGROUP_ROW_FONT, subgroup=label)
        row += 1
    return row


def write_profit_loss_row2(ws, row, label, schedule_sheet_name, schedule_row):
    """Same as write_profit_loss_row but on the Sub-Group-column layout used by PL/BS Summary."""
    ws.cell(row, SUM2_COL_LABEL, value=label).font = PROFIT_LOSS_LABEL_FONT
    for summary_col, schedule_col in SUMMARY2_TO_SCHEDULE_COL.items():
        schedule_col_letter = get_column_letter(schedule_col)
        cell = ws.cell(row, summary_col, value=f"='{schedule_sheet_name}'!{schedule_col_letter}{schedule_row}")
        cell.number_format = NUM_FORMAT
        cell.font = PROFIT_LOSS_VALUE_FONT
    value_range = f"{get_column_letter(SUM2_COL_CO1)}{row}:{get_column_letter(SUM2_COL_GRANDTOTAL)}{row}"
    ws.conditional_formatting.add(
        value_range, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PROFIT_FILL))
    ws.conditional_formatting.add(
        value_range, CellIsRule(operator="lessThan", formula=["0"], fill=LOSS_FILL))


def write_grouping_header2(ws, row, label):
    ws.cell(row, SUM2_COL_LABEL, value=label).font = SECTION_FONT
    for c in range(1, SUM2_COL_GRANDTOTAL + 1):
        ws.cell(row, c).fill = SECTION_FILL


def find_summary_row(ws, label):
    """Row number whose column-A label matches `label` (case-insensitive). Rows are located by
    TEXT, never by a remembered number - see SKILL.md verification checklist item 0."""
    target = label.strip().upper()
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip().upper() == target:
            return r
    return None


def write_pl_summary_sheet(ws, schedule_sheet_name, row_map, section_buckets, section_detail):
    row = write_summary_header2(ws, "Profit & Loss A/c - Summary")

    write_grouping_header2(ws, row, "TURNOVER"); row += 1
    row = write_subgroup_rows2(ws, row, schedule_sheet_name, section_buckets.get("TURNOVER", {}))
    write_linked_row2(ws, row, "Total Turnover", schedule_sheet_name, row_map["TOTAL TURNOVER"],
                       SECTION_TOTAL_FONT, SECTION_TOTAL_FILL); row += 1

    write_grouping_header2(ws, row, "LESS: COST OF SALES"); row += 1
    row = write_subgroup_rows2(ws, row, schedule_sheet_name, section_buckets.get("COST OF SALES", {}))
    write_linked_row2(ws, row, "Total Cost of Sales", schedule_sheet_name, row_map["TOTAL COST OF SALES"],
                       SECTION_TOTAL_FONT, SECTION_TOTAL_FILL); row += 1

    write_grouping_header2(ws, row, "BRANCH TRANSACTIONS (shown separately, included in Gross Profit)"); row += 1
    row = write_subgroup_rows2(ws, row, schedule_sheet_name, section_buckets.get("BRANCH TRANSACTIONS", {}))
    write_linked_row2(ws, row, "Net Branch Transactions", schedule_sheet_name, row_map["NET BRANCH TRANSACTIONS"],
                       ADJUSTMENT_TOTAL_FONT, ADJUSTMENT_TOTAL_FILL); row += 1

    write_profit_loss_row2(ws, row, "GROSS PROFIT", schedule_sheet_name, row_map["GROSS PROFIT"]); row += 1

    write_grouping_header2(ws, row, "ADD: OTHER INCOME"); row += 1
    row = write_subgroup_rows2(ws, row, schedule_sheet_name, section_buckets.get("OTHER INCOME", {}))
    write_linked_row2(ws, row, "Total Other Income", schedule_sheet_name, row_map["TOTAL OTHER INCOME"],
                       SECTION_TOTAL_FONT, SECTION_TOTAL_FILL); row += 1

    write_grouping_header2(ws, row, "LESS: OPERATING EXPENSES"); row += 1
    indirect_expense_subgroups = section_detail.get("OPERATING EXPENSES", {}).get("Indirect Expenses", {})
    row = write_subgroup_rows2(ws, row, schedule_sheet_name, indirect_expense_subgroups)
    write_linked_row2(ws, row, "Total Operating Expenses", schedule_sheet_name, row_map["TOTAL OPERATING EXPENSES"],
                       SECTION_TOTAL_FONT, SECTION_TOTAL_FILL); row += 1

    write_profit_loss_row2(ws, row, "EBITDA", schedule_sheet_name, row_map["EBITDA"]); row += 1

    write_grouping_header2(ws, row, "LESS: FINANCE COST"); row += 1
    row = write_subgroup_rows2(ws, row, schedule_sheet_name, section_buckets.get("FINANCE COST", {}))
    write_linked_row2(ws, row, "Total Finance Cost", schedule_sheet_name, row_map["TOTAL FINANCE COST"],
                       SECTION_TOTAL_FONT, SECTION_TOTAL_FILL); row += 1
    write_linked_row2(ws, row, "Less: Depreciation", schedule_sheet_name, row_map["LESS: DEPRECIATION"],
                       ZERO_FONT, ZERO_FILL); row += 1

    write_profit_loss_row2(ws, row, "PBT (Profit Before Tax)", schedule_sheet_name, row_map["PBT"]); row += 1

    write_linked_row2(ws, row, "Less: Tax", schedule_sheet_name, row_map["LESS: TAX"], ZERO_FONT, ZERO_FILL); row += 1

    write_profit_loss_row2(ws, row, "PAT (Net Profit)", schedule_sheet_name, row_map["PAT"]); row += 1

    write_grouping_header2(ws, row, "RELATED PARTY TRANSACTIONS (reference only)"); row += 1
    row = write_subgroup_rows2(ws, row, schedule_sheet_name, section_buckets.get("RELATED PARTY TRANSACTIONS", {}))
    write_linked_row2(ws, row, "Net Related Party Transactions", schedule_sheet_name,
                       row_map["NET RELATED PARTY TRANSACTIONS"], ADJUSTMENT_TOTAL_FONT, ADJUSTMENT_TOTAL_FILL); row += 1

    write_profit_loss_row2(ws, row, "PAT (Net of Related Party Transactions)", schedule_sheet_name,
                            row_map["PAT (NET OF RELATED PARTY)"]); row += 1

    # Valuation-related provisions bridge, linked to the same section on the schedule sheet -
    # see write_pl_sheet() for why this is a genuine bridge (Before excludes the items, After is
    # the true reported PAT), not the double-counted loop an earlier one-off patch had.
    direct_expense_subgroups = section_detail.get("COST OF SALES", {}).get("Direct Expenses", {})
    valuation_schedule_row = direct_expense_subgroups.get("Direct Expenses")
    if valuation_schedule_row is not None and "PROFIT BEFORE VALUATION" in row_map:
        row += 1
        write_grouping_header2(ws, row, "VALUATION-RELATED PROVISIONS (most volatile Direct Expenses items)"); row += 1
        write_linked_row2(ws, row, "Valuation-related Provisions", schedule_sheet_name,
                           valuation_schedule_row, ADJUSTMENT_TOTAL_FONT, ADJUSTMENT_TOTAL_FILL); row += 1
        write_profit_loss_row2(ws, row, "Profit Before Valuation (excl. items above)", schedule_sheet_name,
                                row_map["PROFIT BEFORE VALUATION"]); row += 1
        write_profit_loss_row2(ws, row, "Profit After Valuation (= Reported PAT)", schedule_sheet_name,
                                row_map["PROFIT AFTER VALUATION"])

    turnover_row = find_summary_row(ws, "Total Turnover")
    if turnover_row:
        add_grand_total_pct_column(ws, SUM2_COL_GRANDTOTAL, turnover_row,
                                   first_row=turnover_row - 3, last_row=row,
                                   header_row=turnover_row - 4)


def write_bs_summary_sheet(ws, schedule_sheet_name, row_map, section_buckets):
    row = write_summary_header2(ws, "Balance Sheet - Summary")

    write_grouping_header2(ws, row, "LIABILITIES (Sources of Funds)"); row += 1
    row = write_subgroup_rows2(ws, row, schedule_sheet_name, section_buckets.get("LIABILITIES", {}))
    write_linked_row2(ws, row, "Total Liabilities", schedule_sheet_name, row_map["TOTAL LIABILITIES"],
                       SECTION_TOTAL_FONT, SECTION_TOTAL_FILL); row += 1

    write_grouping_header2(ws, row, "ASSETS (Application of Funds)"); row += 1
    row = write_subgroup_rows2(ws, row, schedule_sheet_name, section_buckets.get("ASSETS", {}))
    write_linked_row2(ws, row, "Total Assets", schedule_sheet_name, row_map["TOTAL ASSETS"],
                       SECTION_TOTAL_FONT, SECTION_TOTAL_FILL); row += 1

    write_linked_row2(ws, row, "Reconciliation Gap (~0 expected)", schedule_sheet_name,
                       row_map["RECONCILIATION GAP"], GAP_FONT, GAP_FILL)


# --- Ratio Analysis: every figure is a formula combining cells on the PL/BS GL Detail
# schedule sheets (never an independently computed value). Quarter length = 91 days
# (01-Apr-2026 to 30-Jun-2026) for the day-count ratios.

QUARTER_DAYS = 91


def _cellref(sheet_name, row_map, key, col_letter):
    return f"'{sheet_name}'!{col_letter}{row_map[key]}"


# Ratio Workings: every ratio sheet (quarterly Ratio Analysis, monthly MOM Ratio Analysis) is laid
# out in three stacked sections top to bottom so every number can be traced by eye, not just
# trusted:
#   1. RATIO SUMMARY (top) - the headline numbers. Each cell is a same-sheet link down to its row
#      in section 3 - never computed directly, so the top figure and the workings below it can
#      never show different numbers.
#   2. COMPONENT FIGURES (middle) - every raw P&L/BS figure any ratio needs, pulled ONCE per
#      company column directly from PL/BS GL Detail (or the MOM schedules for the monthly sheet),
#      as a plain same-sheet-visible cell link.
#   3. RATIO CALCULATIONS (bottom) - the actual ratio formulas, built purely from same-sheet
#      references into section 2 above (e.g. =C40/C13*100) - no cross-sheet references at all, so
#      anyone auditing a ratio can see every input value a few rows up on the SAME sheet.
# PL_COMPONENTS / BS_COMPONENTS / ratio_definitions() are the single source of truth for both
# sheets - add a component or ratio here once and it appears in both, in both sections, on the
# next rebuild.

PL_COMPONENTS = [
    ("GP", "Gross Profit", "GROSS PROFIT"),
    ("TURNOVER", "Total Turnover", "TOTAL TURNOVER"),
    ("EBITDA", "EBITDA", "EBITDA"),
    ("PBT", "PBT", "PBT"),
    ("PAT", "PAT", "PAT"),
    ("FINANCE_COST", "Total Finance Cost", "TOTAL FINANCE COST"),
    ("COGS", "Total Cost of Sales", "TOTAL COST OF SALES"),
]

BS_COMPONENTS = [
    ("TOTAL_ASSETS", "Total Assets", "TOTAL ASSETS"),
    ("CAPITAL", "Capital Account", "Capital Account"),
    ("RESERVES", "Reserves & Surplus", "Reserves & Surplus"),
    ("CURR_ASSETS", "Current Assets", "Current Assets"),
    ("BANK", "Bank Accounts", "Bank Accounts"),
    ("CASH", "Cash-in-hand", "Cash-in-hand"),
    ("DEBTORS", "Sundry Debtors", "Sundry Debtors"),
    ("LOANS_ADV_ASSET", "Loans & Advances (Asset)", "Loans & Advances (Asset)"),
    ("DEPOSITS_ASSET", "Deposits (Asset)", "Deposits (Asset)"),
    ("CREDITORS", "Sundry Creditors", "Sundry Creditors"),
    ("CURR_LIAB", "Current Liabilities", "Current Liabilities"),
    ("DUTIES_TAXES", "Duties & Taxes", "Duties & Taxes"),
    ("PROVISIONS", "Provisions", "Provisions"),
    ("BANK_OD", "Bank OD A/c", "Bank OD A/c"),
    ("SECURED_LOANS", "Secured Loans", "Secured Loans"),
    ("UNSECURED_LOANS", "Unsecured Loans", "Unsecured Loans"),
    ("LOANS_LIAB", "Loans (Liability)", "Loans (Liability)"),
]

# Average-balance components: standard ratio methodology averages a balance-sheet figure against
# a period FLOW figure (PAT, Turnover, Cost of Sales) - e.g. Debtor Days should divide by average
# Debtors, not the closing balance, since Turnover accrued across the whole period while Debtors is
# a snapshot. Only used by ratios that mix a flow with a balance (Asset Turnover, Debtor Days,
# Creditor Days) - Liquidity (Current/Quick Ratio) and Debt-Equity are balance-to-balance snapshots
# and correctly stay on the closing-only components above. Each AVG component reuses its closing
# counterpart's tally_key; the (opening_cell + closing_cell)/2 formula is built directly in
# write_ratio_workings_section from two source columns the caller supplies per company per period -
# when no true opening balance exists for a period (April, and the quarter itself, both of which
# would need the 31-Mar-2026 opening Balance Sheet, not yet fetched from Tally), the caller passes
# the closing column as BOTH the opening and closing source, so the average formula degenerates to
# the closing value itself - a deliberate, documented fallback, not a bug.
AVG_COMPONENTS = [
    ("TOTAL_ASSETS_AVG", "Total Assets (Average of Opening & Closing)", "TOTAL ASSETS"),
    ("DEBTORS_AVG", "Sundry Debtors (Average of Opening & Closing)", "Sundry Debtors"),
    ("CREDITORS_AVG", "Sundry Creditors (Average of Opening & Closing)", "Sundry Creditors"),
]

ALL_COMPONENTS = PL_COMPONENTS + BS_COMPONENTS + AVG_COMPONENTS


def ratio_definitions():
    """(category_label, [(ratio_label, calc_fn), ...]) tuples shared by the quarterly Ratio
    Analysis sheet and the monthly MOM Ratio Analysis sheet. calc_fn(r, day_count) takes r - a
    dict mapping component key (see ALL_COMPONENTS) -> the same-sheet cell reference (e.g. "C40")
    holding that component's Component Figures row, for the column currently being written - and
    returns the Excel formula string for the matching Ratio Calculations row. day_count is only
    used by the day-count ratios (Debtor Days, Creditor Days) - never annualised."""
    return [
        ("PROFITABILITY", [
            ("Gross Profit Margin %", lambda r, d: f"=IFERROR({r['GP']}/{r['TURNOVER']}*100,0)"),
            ("EBITDA Margin %", lambda r, d: f"=IFERROR({r['EBITDA']}/{r['TURNOVER']}*100,0)"),
            ("PBT Margin %", lambda r, d: f"=IFERROR({r['PBT']}/{r['TURNOVER']}*100,0)"),
            ("PAT Margin %", lambda r, d: f"=IFERROR({r['PAT']}/{r['TURNOVER']}*100,0)"),
            ("Return on Assets (ROA) %", lambda r, d: f"=IFERROR({r['PAT']}/ABS({r['TOTAL_ASSETS']})*100,0)"),
            ("Return on Equity (ROE) %",
             lambda r, d: f"=IFERROR({r['PAT']}/({r['CAPITAL']}+{r['RESERVES']})*100,0)"),
        ]),
        ("LIQUIDITY", [
            ("Current Ratio", lambda r, d: (
                f"=IFERROR(ABS({r['CURR_ASSETS']}+{r['BANK']}+{r['CASH']}+{r['DEBTORS']}"
                f"+{r['LOANS_ADV_ASSET']}+{r['DEPOSITS_ASSET']})"
                f"/({r['CREDITORS']}+{r['CURR_LIAB']}+{r['DUTIES_TAXES']}+{r['PROVISIONS']}+{r['BANK_OD']}),0)")),
            ("Quick Ratio (excl. inventory)", lambda r, d: (
                f"=IFERROR(ABS({r['BANK']}+{r['CASH']}+{r['DEBTORS']}"
                f"+{r['LOANS_ADV_ASSET']}+{r['DEPOSITS_ASSET']})"
                f"/({r['CREDITORS']}+{r['CURR_LIAB']}+{r['DUTIES_TAXES']}+{r['PROVISIONS']}+{r['BANK_OD']}),0)")),
        ]),
        ("LEVERAGE", [
            ("Debt-Equity Ratio", lambda r, d: (
                f"=IFERROR(({r['SECURED_LOANS']}+{r['UNSECURED_LOANS']}+{r['BANK_OD']}+{r['LOANS_LIAB']})"
                f"/({r['CAPITAL']}+{r['RESERVES']}),0)")),
            ("Interest Coverage Ratio", lambda r, d: f"=IFERROR(ABS({r['EBITDA']})/ABS({r['FINANCE_COST']}),0)"),
        ]),
        ("EFFICIENCY", [
            ("Asset Turnover Ratio", lambda r, d: f"=IFERROR({r['TURNOVER']}/ABS({r['TOTAL_ASSETS_AVG']}),0)"),
            ("Debtor Days", lambda r, d: f"=IFERROR(ABS({r['DEBTORS_AVG']})/{r['TURNOVER']}*{d},0)"),
            ("Creditor Days", lambda r, d: f"=IFERROR({r['CREDITORS_AVG']}/ABS({r['COGS']})*{d},0)"),
        ]),
    ]


def ratio_sheet_layout(categories, start_row):
    """Precomputes row numbers for the three stacked Ratio Workings sections before any cell is
    written, since the Summary section (written first, at the top) needs to know which
    Calculations row each ratio will land on (written last, at the bottom) in order to build its
    link formula. Row count is fully determined by len(categories)/len(ratios)/len(ALL_COMPONENTS)
    - pure arithmetic, no sheet access."""
    row = start_row
    summary_row_of = {}
    for ci, (_, ratios) in enumerate(categories):
        row += 1  # category header
        for ri in range(len(ratios)):
            summary_row_of[(ci, ri)] = row
            row += 1
        row += 1  # spacer
    summary_end = row

    components_header_row = summary_end + 1
    component_row_of = {}
    row = components_header_row + 2  # header row + "From Profit & Loss" subheader row
    for key, _, _ in PL_COMPONENTS:
        component_row_of[key] = row
        row += 1
    row += 1  # spacer
    row += 1  # "From Balance Sheet" subheader row
    for key, _, _ in BS_COMPONENTS:
        component_row_of[key] = row
        row += 1
    row += 1  # spacer
    row += 1  # "Average Balances" subheader row
    for key, _, _ in AVG_COMPONENTS:
        component_row_of[key] = row
        row += 1
    components_end = row

    calc_header_row = components_end + 1
    calc_row_of = {}
    row = calc_header_row + 1
    for ci, (_, ratios) in enumerate(categories):
        row += 1  # category header
        for ri in range(len(ratios)):
            calc_row_of[(ci, ri)] = row
            row += 1
        row += 1  # spacer
    calc_end = row

    return {
        "summary_row_of": summary_row_of, "summary_end": summary_end,
        "components_header_row": components_header_row, "component_row_of": component_row_of,
        "components_end": components_end,
        "calc_header_row": calc_header_row, "calc_row_of": calc_row_of, "calc_end": calc_end,
    }


def write_ratio_summary_section(ws, layout, categories, columns, label_col=SUM_COL_LABEL):
    """columns: [(key, write_col), ...] - key is one of RATIO_KEYS (used only to decide
    subtotal-column shading); write_col is the column each ratio value lands in on this sheet.
    Writes simple same-sheet links down to the matching Ratio Calculations row - never a computed
    formula directly, so this section can never drift from the workings below it."""
    max_col = max(c for _, c in columns)
    for ci, (category_label, ratios) in enumerate(categories):
        hdr_row = layout["summary_row_of"][(ci, 0)] - 1
        ws.cell(hdr_row, label_col, value=category_label).font = SECTION_FONT
        for c in range(1, max_col + 1):
            ws.cell(hdr_row, c).fill = SECTION_FILL
        for ri, (ratio_label, _) in enumerate(ratios):
            r = layout["summary_row_of"][(ci, ri)]
            ws.cell(r, label_col, value=ratio_label).font = NORMAL_FONT
            calc_r = layout["calc_row_of"][(ci, ri)]
            for key, write_col in columns:
                col_letter = get_column_letter(write_col)
                cell = ws.cell(r, write_col, value=f"={col_letter}{calc_r}")
                cell.number_format = "0.00"
                cell.font = NORMAL_FONT
                if key in ("TOTAL_ORANGE", "TOTAL_ENTERPRISES", "GRANDTOTAL"):
                    cell.fill = SUBTOTAL_FILL


def write_ratio_workings_section(ws, layout, categories, columns, pl, bs, label_col=SUM_COL_LABEL):
    """columns: [(key, write_col, closing_col_letter, opening_col_letter, day_count), ...] -
    write_col is where the Component Figures / Ratio Calculations value lands on THIS sheet;
    closing_col_letter is the column to read the period's own closing balance from on the PL/BS
    GL Detail (or MOM equivalent) schedule sheet for that same company; opening_col_letter is the
    column holding that same balance-sheet figure's opening-of-period value (the PRIOR period's
    closing column) - used only by the Average Balances components below, and equal to
    closing_col_letter when no true opening figure is available (see AVG_COMPONENTS docstring for
    the fallback rule this produces). day_count lets the monthly sheet pass a different day-count
    per period block while sharing this one function with the single-period quarterly sheet.
    Writes the Component Figures and Ratio Calculations sections (Summary is written separately by
    write_ratio_summary_section, using the calc_row_of positions this function's layout already
    reserved)."""
    max_col = max(c for _, c, _, _, _ in columns)

    ws.cell(layout["components_header_row"], label_col,
            value="RATIO WORKINGS — COMPONENT FIGURES (linked directly from the GL Detail "
                  "schedule; every Ratio Calculation below is built only from these rows)").font = SECTION_FONT
    for c in range(1, max_col + 1):
        ws.cell(layout["components_header_row"], c).fill = SECTION_FILL

    row = layout["components_header_row"] + 1
    ws.cell(row, label_col, value="From Profit & Loss").font = SUBGROUP_FONT
    row += 1
    for key, label, tally_key in PL_COMPONENTS:
        ws.cell(row, label_col, value=label).font = NORMAL_FONT
        for _, write_col, src_col, _, _ in columns:
            cell = ws.cell(row, write_col, value=f"={pl(tally_key, src_col)}")
            cell.number_format = NUM_FORMAT
            cell.font = NORMAL_FONT
        row += 1
    row += 1
    ws.cell(row, label_col, value="From Balance Sheet").font = SUBGROUP_FONT
    row += 1
    for key, label, tally_key in BS_COMPONENTS:
        ws.cell(row, label_col, value=label).font = NORMAL_FONT
        for _, write_col, src_col, _, _ in columns:
            cell = ws.cell(row, write_col, value=f"={bs(tally_key, src_col)}")
            cell.number_format = NUM_FORMAT
            cell.font = NORMAL_FONT
        row += 1
    row += 1
    ws.cell(row, label_col, value=(
        "Average Balances (Opening + Closing) ÷ 2 — used only by ratios that mix a "
        "period flow figure with a balance-sheet snapshot (Asset Turnover, Debtor Days, Creditor "
        "Days); Apr-26 and the quarterly view fall back to the closing balance alone until the "
        "31-Mar-2026 opening Balance Sheet is fetched from Tally")).font = SUBGROUP_FONT
    row += 1
    for key, label, tally_key in AVG_COMPONENTS:
        ws.cell(row, label_col, value=label).font = NORMAL_FONT
        for _, write_col, src_col, open_col, _ in columns:
            cell = ws.cell(row, write_col, value=f"=({bs(tally_key, open_col)}+{bs(tally_key, src_col)})/2")
            cell.number_format = NUM_FORMAT
            cell.font = NORMAL_FONT
        row += 1

    ws.cell(layout["calc_header_row"], label_col,
            value="RATIO WORKINGS — CALCULATIONS (computed from the component figures above; "
                  "the Ratio Summary at the top of this sheet links to these rows)").font = SECTION_FONT
    for c in range(1, max_col + 1):
        ws.cell(layout["calc_header_row"], c).fill = SECTION_FILL

    for ci, (category_label, ratios) in enumerate(categories):
        hdr_row = layout["calc_row_of"][(ci, 0)] - 1
        ws.cell(hdr_row, label_col, value=category_label).font = SECTION_FONT
        for c in range(1, max_col + 1):
            ws.cell(hdr_row, c).fill = SECTION_FILL
        for ri, (ratio_label, calc_fn) in enumerate(ratios):
            r = layout["calc_row_of"][(ci, ri)]
            ws.cell(r, label_col, value=ratio_label).font = NORMAL_FONT
            for key, write_col, _, _, day_count in columns:
                col_letter = get_column_letter(write_col)
                refs = {ck: f"{col_letter}{layout['component_row_of'][ck]}" for ck, _, _ in ALL_COMPONENTS}
                formula = calc_fn(refs, day_count)
                cell = ws.cell(r, write_col, value=formula)
                cell.number_format = "0.00"
                cell.font = NORMAL_FONT
                if key in ("TOTAL_ORANGE", "TOTAL_ENTERPRISES", "GRANDTOTAL"):
                    cell.fill = SUBTOTAL_FILL


def write_ratio_sheet(ws, pl_row_map, bs_row_map):
    row = write_summary_header(ws, "Ratio Analysis")
    ws.cell(2, 1, value=(
        f"Period: {PERIOD_LABEL}  ({QUARTER_DAYS}-day quarter used for day-count ratios, not "
        "annualised). RATIO SUMMARY below links to RATIO WORKINGS further down this sheet - every "
        "input figure behind each ratio is shown there for cross-check. Average Balances fall back "
        "to the closing balance alone (no distinct opening figure - the quarter's opening would be "
        "the 31-Mar-2026 Balance Sheet, not yet fetched from Tally)."))

    def pl(key, col_letter):
        return _cellref("PL GL Detail", pl_row_map, key, col_letter)

    def bs(key, col_letter):
        return _cellref("BS GL Detail", bs_row_map, key, col_letter)

    categories = ratio_definitions()
    layout = ratio_sheet_layout(categories, row)

    summary_columns = [(key, RATIO_KEY_TO_SUM_COL[key]) for key in RATIO_KEYS]
    # No 31-Mar-2026 opening Balance Sheet is fetched yet, so the quarter has no distinct opening
    # figure - opening_col == closing_col, making the Average Balances components degenerate to
    # the closing balance alone (see AVG_COMPONENTS docstring).
    workings_columns = [(key, RATIO_KEY_TO_SUM_COL[key],
                          get_column_letter(SUMMARY_TO_SCHEDULE_COL[RATIO_KEY_TO_SUM_COL[key]]),
                          get_column_letter(SUMMARY_TO_SCHEDULE_COL[RATIO_KEY_TO_SUM_COL[key]]), QUARTER_DAYS)
                         for key in RATIO_KEYS]

    write_ratio_summary_section(ws, layout, categories, summary_columns)
    write_ratio_workings_section(ws, layout, categories, workings_columns, pl, bs)


# --- Q1 Performance Report: a small, single-screen management sheet - headline P&L/BS figures
# plus a compact action-plan table. Every value cell is a formula linking into 'Ratio Analysis'
# (never a duplicated static number) - and since 'Ratio Analysis' itself links from PL GL Detail /
# BS GL Detail, every figure here is fully traceable back to source with no independent
# computation anywhere in the chain. Deliberately NOT a rebuild of the separate company-review
# skill's multi-tab board pack - this is one compact sheet only, per explicit user request ("keep
# it simple and small").

def find_row_by_label(ws, label, col=1):
    """Scans column `col` of `ws` top to bottom for an exact match on `label` and returns its row
    number. Used so a source row is always located by its label text, never a hardcoded row
    number - row numbers on 'Ratio Analysis' shift whenever a ratio/component is added or removed,
    so hardcoding one here would silently point at the wrong row after a future rebuild (the same
    staleness risk flagged throughout this codebase, e.g. see Cross Check's own find_row())."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, col).value == label:
            return r
    raise ValueError(f"Row not found for label {label!r} in sheet {ws.title!r}")


# Company-wise reporting reflects the REAL legal-entity structure, corrected per management:
# Orange O Tec Pvt Ltd (Surat) and Orange O Tec NOIDA are ONE company operating through two
# branches; likewise Enterprises Pvt Ltd (Surat) and Enterprises NOIDA are ONE company. So there
# are three real companies in this group - Orange O Tec, Enterprises, and Colorix (an LLP) - not
# five. The GL Detail / Ratio Analysis sheets elsewhere in this workbook keep the branch-level
# CO1..CO5 breakdown (useful operational MIS, left untouched), but THIS management-facing sheet
# reports at the real-company level: it pulls only the TotalOrange / TotalEnterprises / CO5 /
# GrandTotal columns from 'Ratio Analysis' (already branch-combined for the two two-branch
# companies), and uses its own compact 5-column header rather than the shared 9-column
# write_summary_header()/write_grouping_header() used by every branch-level sheet in this file.
Q1_DEST_COLS = (2, 3, 4, 5)  # B, C, D, E on this sheet: Orange O Tec, Enterprises, Colorix, Grand Total
Q1_SRC_COLS = (SUM_COL_TOTAL_ORANGE, SUM_COL_TOTAL_ENTERPRISES, SUM_COL_CO5, SUM_COL_GRANDTOTAL)
Q1_COMPANY_LABELS = ("Orange O Tec\n(Surat + NOIDA branches)", "Enterprises\n(Surat + NOIDA branches)",
                     "Colorix (LLP)", "Grand Total")


def _q1_section_header(ws, row, label, ncols=5):
    ws.cell(row, 1, value=label).font = SECTION_FONT
    for c in range(1, ncols + 1):
        ws.cell(row, c).fill = SECTION_FILL


def write_q1_header(ws, title):
    """Compact 5-column header (Label + 3 real companies + Grand Total) for this sheet only."""
    ws.column_dimensions["A"].width = 30
    for col_letter in "BCD":
        ws.column_dimensions[col_letter].width = 24
    ws.column_dimensions["E"].width = 18

    row = 1
    ws.cell(row, 1, value=title).font = TITLE_FONT
    row += 1
    ws.cell(row, 1, value=f"Period: {PERIOD_LABEL}  (figures linked live from the GL Detail schedule sheets, combined at the real-company level)").font = NORMAL_FONT
    row += 2

    header_row = row
    ws.cell(row, 1, value="Line Item")
    for c, label in zip(Q1_DEST_COLS, Q1_COMPANY_LABELS):
        ws.cell(row, c, value=label)
    for c in range(1, 6):
        cell = ws.cell(row, c)
        cell.font = HEADER_FONT
        cell.fill = TOTALCOL_FILL if c == 5 else HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="bottom")
    ws.row_dimensions[row].height = 32
    return header_row + 2


def write_q1_linked_row(ws, row, label, ratio_ws, source_row, font, number_format=NUM_FORMAT):
    """Writes one row whose 4 value cells (Orange O Tec / Enterprises / Colorix / Grand Total)
    link to the corresponding TotalOrange/TotalEnterprises/CO5/GrandTotal column on 'Ratio
    Analysis' for the given source row."""
    ws.cell(row, 1, value=label).font = font
    for dest_col, src_col in zip(Q1_DEST_COLS, Q1_SRC_COLS):
        col_letter = get_column_letter(src_col)
        cell = ws.cell(row, dest_col, value=f"='{ratio_ws.title}'!{col_letter}{source_row}")
        cell.number_format = number_format
        cell.font = font
    return row + 1


def write_q1_sum_row(ws, row, label, ratio_ws, source_rows, font):
    """Writes a row that is the SUM of two or more 'Ratio Analysis' rows per company (used for
    Net Worth = Capital Account + Reserves & Surplus) - still a pure formula link."""
    ws.cell(row, 1, value=label).font = font
    for dest_col, src_col in zip(Q1_DEST_COLS, Q1_SRC_COLS):
        col_letter = get_column_letter(src_col)
        formula = "=" + "+".join(f"'{ratio_ws.title}'!{col_letter}{r}" for r in source_rows)
        cell = ws.cell(row, dest_col, value=formula)
        cell.number_format = NUM_FORMAT
        cell.font = font
    return row + 1


def write_q1_profit_row(ws, row, label, ratio_ws, source_row):
    """Same as write_q1_linked_row but with the green/red profit-loss conditional coloring used
    for PAT elsewhere in this workbook."""
    ws.cell(row, 1, value=label).font = PROFIT_LOSS_LABEL_FONT
    for dest_col, src_col in zip(Q1_DEST_COLS, Q1_SRC_COLS):
        col_letter = get_column_letter(src_col)
        cell = ws.cell(row, dest_col, value=f"='{ratio_ws.title}'!{col_letter}{source_row}")
        cell.number_format = NUM_FORMAT
        cell.font = PROFIT_LOSS_VALUE_FONT
    value_range = f"B{row}:E{row}"
    ws.conditional_formatting.add(
        value_range, CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=PROFIT_FILL))
    ws.conditional_formatting.add(
        value_range, CellIsRule(operator="lessThan", formula=["0"], fill=LOSS_FILL))
    return row + 1


def write_q1_table(ws, row, title, headers, rows, ncols):
    """Generic labeled-column table writer shared by the commentary, cost/tax, tax-planning, and
    action-plan sections below - same visual pattern (section header -> column headers -> wrapped
    body rows), parameterised so several structurally-similar tables share one implementation."""
    ws.cell(row, 1, value=title).font = SECTION_FONT
    for c in range(1, ncols + 1):
        ws.cell(row, c).fill = SECTION_FILL
    row += 1
    for c, label in enumerate(headers, start=1):
        cell = ws.cell(row, c, value=label)
        cell.font = HEADER_FONT
        cell.fill = TOTALCOL_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 20
    row += 1
    for data_row in rows:
        for c, val in enumerate(data_row, start=1):
            cell = ws.cell(row, c, value=val)
            cell.font = (Font(bold=True, size=10, name="Calibri", color="1F3864") if c == 1
                         else NORMAL_FONT)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1
    return row


def _q1_est_row_height(texts, chars_per_line=55, pt_per_line=14, minimum=40):
    """Rough wrapped-row height estimate (chars-per-line tuned for the wide merged columns used
    by write_q1_merged_table) - not exact, but keeps long rows from rendering too short."""
    longest = max((len(t) for t in texts if t), default=0)
    lines = max(1, -(-longest // chars_per_line))
    return max(minimum, lines * pt_per_line)


def write_q1_merged_table(ws, row, title, headers, rows):
    """4-column labeled table (Area | Observation | Recommendation | Note), like write_q1_table,
    but with each row's Observation/Recommendation/Note cells merged across several physical grid
    columns (B:D, E:I, J:K) for extra width - matches the layout the user manually applied to the
    Cost Reduction Focus table in Excel, now baked in here so every rebuild reproduces it rather
    than requiring another one-off patch. Area (column A) stays a single, unmerged column."""
    area_font = Font(bold=True, size=10, name="Calibri", color="1F3864")
    wrap_top = Alignment(wrap_text=True, vertical="top")
    wrap_top_center = Alignment(horizontal="center", wrap_text=True, vertical="top")
    merge_spans = ((2, 4), (5, 9), (10, 11))  # Observation, Recommendation, Note

    ws.cell(row, 1, value=title).font = SECTION_FONT
    for c in range(1, 12):
        ws.cell(row, c).fill = SECTION_FILL
    row += 1

    for (start_col, end_col) in merge_spans:
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    for col, label in zip((1, 2, 5, 10), headers):
        cell = ws.cell(row, col, value=label)
        cell.font = HEADER_FONT
        cell.fill = TOTALCOL_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 20
    row += 1

    for area, observation, recommendation, note in rows:
        for (start_col, end_col) in merge_spans:
            ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
        ws.cell(row, 1, value=area).font = area_font
        ws.cell(row, 1).alignment = wrap_top
        ws.cell(row, 2, value=observation).font = NORMAL_FONT
        ws.cell(row, 2).alignment = wrap_top_center
        ws.cell(row, 5, value=recommendation).font = NORMAL_FONT
        ws.cell(row, 5).alignment = wrap_top_center
        ws.cell(row, 10, value=note).font = NORMAL_FONT
        ws.cell(row, 10).alignment = wrap_top_center
        ws.row_dimensions[row].height = _q1_est_row_height([observation, recommendation])
        row += 1
    return row


def write_q1_company_commentary(ws, row, ratio_ws, commentary_text):
    """One row per REAL company (Orange O Tec, Enterprises, Colorix, Grand Total) - Orange O Tec
    and Enterprises are each combined across their Surat + NOIDA branches, since those branches
    are the same legal entity, not separate companies. Turnover/PAT % are formula links to
    'Ratio Analysis' (already branch-combined); only the commentary text is authored."""
    turnover_row = find_row_by_label(ratio_ws, "Total Turnover")
    patpct_row = find_row_by_label(ratio_ws, "PAT Margin %")
    company_labels = ["Orange O Tec", "Enterprises", "Colorix", "Grand Total"]
    ws.cell(row, 1, value="MANAGEMENT COMMENTARY - COMPANY WISE").font = SECTION_FONT
    for c in range(1, 5):
        ws.cell(row, c).fill = SECTION_FILL
    row += 1
    for c, label in enumerate(Q1_COMMENTARY_HEADER, start=1):
        cell = ws.cell(row, c, value=label)
        cell.font = HEADER_FONT
        cell.fill = TOTALCOL_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 20
    row += 1
    for label, src_col, text in zip(company_labels, Q1_SRC_COLS, commentary_text):
        col_letter = get_column_letter(src_col)
        ws.cell(row, 1, value=label).font = Font(bold=True, size=10, name="Calibri", color="1F3864")
        turn_cell = ws.cell(row, 2, value=f"='{ratio_ws.title}'!{col_letter}{turnover_row}")
        turn_cell.number_format = NUM_FORMAT
        turn_cell.font = NORMAL_FONT
        pat_cell = ws.cell(row, 3, value=f"='{ratio_ws.title}'!{col_letter}{patpct_row}")
        pat_cell.number_format = "0.00"
        pat_cell.font = NORMAL_FONT
        ws.cell(row, 4, value=text).font = NORMAL_FONT
        for c in range(1, 5):
            ws.cell(row, c).alignment = Alignment(wrap_text=True, vertical="top")
        row += 1
    return row


def write_q1_action_plan(ws, row, rows):
    return write_q1_table(ws, row, "ACTION PLAN", ACTION_PLAN_HEADER, rows, ncols=4)


# Action plan: labeled-column table (Area | Observation | Action | Owner/Timeline), not narrative
# bullet prose - filled in from the live, verified Q1 figures once the refresh completes (never
# generic filler - see build note in main() for how this list is populated each rebuild).
ACTION_PLAN_HEADER = ["Area", "Observation", "Action", "Owner / Timeline"]

# Populated from the live, verified Q1 FY26-27 figures (Ratio Analysis / PL Summary / BS Summary)
# after the 2026-07-21 refresh, corrected to the real 3-company structure (Orange O Tec and
# Enterprises are each ONE company across two branches - branch-level items below are flagged as
# such). Update this list whenever the underlying figures materially change.
Q1_ACTION_PLAN = [
    ("Orange O Tec - Group Margin", "PAT margin just 1.38% company-wide (Rs 55.44L on Rs 40.1cr Q1 turnover) - thinnest margin of the three real companies despite the largest volume", "Review pricing and cost pass-through, especially at the Surat branch which carries the bulk of the shortfall; identify which cost lines compress margin most", "Finance / before Q2 close"),
    ("Enterprises - NOIDA branch losses", "Branch-level PAT margin -11.9% this quarter at the NOIDA branch, though the company overall (with the Surat branch) is highly profitable at 45.2%", "Review cost structure and pricing for the NOIDA branch specifically; assess whether it needs a standalone turnaround plan", "Unit head / immediate"),
    ("Group-wide - Debtor Days", "Debtor days range 90-144 days across branches (Enterprises' NOIDA branch highest at 144 days)", "Tighten collections on oldest-ageing accounts; review credit terms for repeat-offender customers", "Collections team / ongoing, review monthly"),
    ("Orange O Tec / Enterprises NOIDA branches - Current Ratio", "Current ratio shows as -101.8 and -10.0 respectively at branch level - mathematically implies a near-zero or negative Current Liabilities base in that branch-level MIS split", "Validate Current Liabilities classification at branch level before relying on this ratio for decisions", "Finance / before next quarter's review"),
    ("Enterprises - Margin & Tax Profile", "Company-wide PAT margin 45.2% (Rs 6.33cr on Rs 14.0cr turnover), on a concessional ~20% tax rate - far above group average", "See FUTURE TAX PLANNING section below - priorities #1 and #2 apply directly here", "Finance & tax advisor / before board review"),
    ("Future tax planning", "See FUTURE TAX PLANNING - ADVICE TO MANAGEMENT section above for the full 5-point plan", "Work through the 5 priorities with your tax advisor ahead of FY26-27 return filing", "Tax advisor / per section timelines"),
]

# Company-wise management commentary: one row per REAL company (not per branch), grounded in the
# same live Q1 figures as the headline table above. Order matches Q1_SRC_COLS.
Q1_COMMENTARY_HEADER = ["Company", "Turnover", "PAT %", "Commentary"]
Q1_COMPANY_COMMENTARY = [
    "Turnover Rs 40.1cr (Surat branch Rs 34.4cr + NOIDA branch Rs 5.7cr), PAT margin just 1.38% company-wide - essentially breakeven. Cost reduction here has the largest absolute rupee impact of the three companies; the NOIDA branch alone runs a healthier 8.3% margin, worth understanding what it does differently.",
    "Turnover Rs 14.0cr (Surat branch Rs 12.0cr + NOIDA branch Rs 2.0cr), PAT margin 45.2% company-wide - the group's most profitable entity by far, on a concessional ~20% manufacturing tax rate. The Surat branch alone shows 54.7% margin; the NOIDA branch is loss-making (-11.9%) but too small to offset it. Confirm the high margin reflects genuine trading/manufacturing profit, not a related-party pricing effect.",
    "Turnover Rs 8.2cr, PAT margin 18.4% - a consistent, healthy performer, no immediate concerns this quarter. As an LLP, it does not have access to the concessional company tax rates available to Orange O Tec or Enterprises.",
    "Group turnover Rs 62.3cr, PAT margin 13.5% blended - masks a very uneven spread across the three real companies: from 1.4% (Orange O Tec) to 45.2% (Enterprises).",
]

# Cost reduction focus + inter-company cost allocation considered through a tax-efficiency lens,
# per management's explicit request - now correctly scoped to the THREE real companies (Orange O
# Tec, Enterprises, Colorix), since Surat/NOIDA branches of the same company cannot meaningfully
# "shift cost" between each other - there is no separate taxpayer to shift to. This is directional
# financial analysis, not tax/legal advice - every row says so, and the compliance guardrail row
# is deliberately non-negotiable.
Q1_COST_TAX_HEADER = ["Area", "Observation", "Recommendation", "Note"]
Q1_COST_TAX_ROWS = [
    ("Where to focus operational cost reduction",
     "Orange O Tec (combined, both branches) has the largest revenue (Rs 40.1cr) and the thinnest margin (1.38% PAT) of the three real companies in the group",
     "Prioritise a direct cost-line review here (purchase cost, freight/carriage, employee cost) - this is where operational savings have the largest rupee impact",
     "-"),
    ("Enterprises - tax profile",
     "Manufacturing company on a concessional ~20% effective tax rate (per management), PAT margin 45.2% combined - by far the most profitable of the three companies",
     "At a ~20% rate, each Rs 1 of expense booked here shields only ~20 paise of tax - the lowest shield value of the three companies",
     "Confirm the exact section (115BAA/115BAB) and rate with your tax advisor"),
    ("Direction of any genuine shared-cost allocation",
     "Standard group tax logic: a deduction is worth more at the entity with the HIGHER marginal tax rate; profit is retained more efficiently at the entity with the LOWER rate. Orange O Tec's normal-regime rate is presumed higher than Enterprises' ~20% concession, and Colorix (an LLP) is taxed at the highest flat rate of the three, with no access to the concessional company rates at all",
     "If any real, commercially-substantiated shared cost exists across the group, it is more tax-efficient to allocate it to Orange O Tec or Colorix than to Enterprises - the opposite direction from shifting cost toward Enterprises",
     "Validate actual rates and Orange O Tec's tax/loss position with your advisor before acting"),
    ("Compliance guardrail",
     "Orange O Tec and Enterprises are each a single legal entity (their Surat and NOIDA operations are branches, not separate companies) - no cost shifting is needed or possible between branches of the same company",
     "Any real reallocation should only be considered BETWEEN the three actual companies (Orange O Tec, Enterprises, Colorix), each backed by a documented cost driver and arm's-length pricing, reviewed by your tax advisor for Section 40A(2)/92BA compliance",
     "Non-negotiable for audit/assessment defensibility"),
    ("Enterprises NOIDA branch - standalone loss",
     "The NOIDA branch is loss-making this quarter (branch-level PAT margin -11.9%), offset by the Surat branch's strong performance in the combined company result",
     "Fix directly through branch-level pricing/cost control - internal branch performance still matters for management even though it does not affect the company's overall tax position",
     "-"),
    ("Import & customs cost concentration - Enterprises",
     "Enterprises carries a large import-driven cost base (GST on imported goods, customs duty, ocean freight) alongside its manufacturing operations - even at a 45%+ margin, a percentage-point improvement here has real absolute rupee value",
     "Review the local-vs-import sourcing mix; pair with the RoDTEP/Advance Authorization/EPCG review under Future Tax Planning below",
     "-"),
    ("Forex hedging - Enterprises",
     "The P&L shows meaningful exchange-rate gain/loss swings tied to import purchases - unhedged foreign currency exposure",
     "A forward-cover policy on large import orders would reduce this volatility-driven cost and make margins more predictable quarter to quarter",
     "-"),
    ("Financing cost structure - Enterprises",
     "Loan processing charges appear disproportionately large relative to the interest cost itself",
     "Renegotiate with the lender or consolidate facilities rather than processing multiple loans/limits separately",
     "-"),
    ("Related-party rent/services benchmarking",
     "Rent is paid partly to a related party alongside regular market rent",
     "Confirm the related-party rent is at market rate - this also strengthens the arm's-length defense already flagged above",
     "-"),
    ("Debtor days - real financing cost, not just a ratio",
     "90-144 days of debtor cycle across the group means working capital is being funded through CC/OD limits",
     "Tightening collections by even 15-20 days cuts real interest expense, not just improves a ratio",
     "-"),
    ("Commission & brokerage vs. direct relationships",
     "A meaningful commission/brokerage cost line exists in the verified data",
     "Review whether higher-value, repeat accounts could move to direct relationships to reduce ongoing commission leakage",
     "-"),
]

# Future tax planning - added at management's explicit request, framed as tax-partner-style
# directional advice. Every row is a recommendation to validate/execute WITH the actual tax
# advisor, not a final position - the concessional-regime conditions (115BAA/115BAB), MAT
# exemption thresholds, and LLP taxation rules cited here are the general framework a tax partner
# would raise, not a substitute for reviewing this group's actual filings and elections.
Q1_TAX_PLANNING_HEADER = ["Priority", "Recommendation", "Why It Matters", "Timeline"]
Q1_TAX_PLANNING_ROWS = [
    ("1 - Protect the concessional regime",
     "Get Enterprises' concessional manufacturing tax rate (Section 115BAB/115BAA, whichever applies) health-checked now: confirm manufacturing start date, no reconstruction of an existing business, plant & machinery usage limits, and MAT non-applicability are all still satisfied",
     "Enterprises carries ~45% of group PAT at a ~20% rate; losing this status would be reassessed at the normal rate and would be the single largest tax risk in the group",
     "Before FY26-27 return filing"),
    ("2 - Benchmark related-party pricing",
     "Commission a formal arm's-length pricing review of transactions between Orange O Tec and Enterprises, given the stark margin gap (1.4% vs 45.2% PAT) between two related group companies",
     "Protects Enterprises' concessional profit from being reallocated to Orange O Tec on assessment under Section 40A(2), and protects Orange O Tec from a notional-income addition",
     "Before Q2 close, ahead of any tax audit"),
    ("3 - Advance tax by entity, not blended",
     "Confirm advance tax instalments for Enterprises reflect its actual quarter-by-quarter profit concentration, not a blended group estimate",
     "Enterprises holds the bulk of group profit this quarter; underestimating its instalment risks interest under Sections 234B/234C",
     "Each quarter, ongoing"),
    ("4 - Review Orange O Tec's regime election",
     "Given its razor-thin 1.38% margin this quarter, run a comparison of Orange O Tec's current tax regime (old-regime deductions vs Section 115BAA's flat ~25.17%) against this year's actual numbers",
     "A regime that made sense at a higher margin may not be optimal now, and vice versa - worth revisiting annually rather than assuming it is set once",
     "Before advance tax planning for FY27-28"),
    ("5 - Colorix (LLP) has the highest marginal rate",
     "Since an LLP cannot access the 115BAA/115BAB concessional company rates, any genuinely-incurred shared group cost should be allocated here (not to Enterprises) wherever it can be commercially substantiated",
     "Colorix's flat LLP tax rate is the highest of the three entities - the opposite direction from shifting cost toward Enterprises",
     "Ongoing, as part of cost allocation review"),
    ("6 - MSME payment timeline (Section 43B(h))",
     "Orange O Tec shows large MSME creditor balances (one supplier alone at roughly Rs 1.56cr under local-MSME creditors) - if these remain unpaid beyond the agreed period (max 45 days under the MSME Development Act), the ENTIRE unpaid amount is disallowed as a deduction and taxed as income this year",
     "This is the single most avoidable tax increase in the group right now - a payment-timing issue, not a business cost, that fully reverses if paid before year-end",
     "Immediate - review the MSME creditor ageing this month"),
    ("7 - Additional depreciation on new plant & machinery",
     "If Orange O Tec is on the old tax regime (not Section 115BAA), confirm all new plant & machinery purchased this year has claimed the extra 20% additional depreciation under Section 32(1)(iia). Enterprises, on a 115BAA/115BAB concessional regime, is NOT eligible for this and should not attempt to claim it",
     "An easy-to-miss deduction that directly reduces Orange O Tec's already-thin taxable profit",
     "Before FY26-27 return filing"),
    ("8 - Brought-forward losses / unabsorbed depreciation",
     "Given Orange O Tec's historically thin margins, confirm any carried-forward business losses or unabsorbed depreciation (Section 72) are being tracked and set off against this year's profit before the 8-year carry-forward window lapses",
     "A rupee of old loss set off against current profit is a full tax saving at Orange O Tec's marginal rate - easy to lose track of if nobody is actively monitoring it year to year",
     "Before FY26-27 return filing"),
    ("9 - Colorix (LLP) partner remuneration & interest",
     "Confirm partner remuneration and interest on partners' capital are being claimed up to the maximum permitted under Section 40(b) - LLPs commonly under-claim this",
     "A direct, fully legitimate reduction in Colorix's own taxable profit, not dependent on any group-level restructuring",
     "Review at each year-end closing"),
    ("10 - Statutory dues paid on time (Section 43B)",
     "Confirm PF, ESIC, bonus, and leave encashment amounts are actually paid (not just provided for) before the return filing due date",
     "Any unpaid statutory due at filing date is disallowed as a deduction regardless of when it was accrued - a fully avoidable addback",
     "Before each return filing"),
    ("11 - Customs/export incentive review - Orange O Tec",
     "Given visible import volumes and customs duty cost in the group, get a customs consultant to check RoDTEP, Advance Authorization, or EPCG scheme eligibility",
     "Reduces the underlying cost base (and indirect tax outgo) rather than income tax directly, but has a real, sustained margin impact given how thin Orange O Tec's PAT already is",
     "Separate workstream, alongside your customs/GST consultant"),
    ("12 - GST input credit health check - Enterprises",
     "As a manufacturer, confirm Enterprises is claiming full input tax credit and check eligibility for an inverted-duty-structure GST refund if input GST rates exceed output rates",
     "Unclaimed ITC or an unfiled refund is cash left on the table even though it does not affect income-tax computation directly",
     "Ongoing, each GST filing cycle"),
]


def write_q1_performance_report_sheet(ws, ratio_ws, action_plan_rows):
    row = write_q1_header(ws, "Q1 Performance Report")

    def r(label):
        return find_row_by_label(ratio_ws, label)

    _q1_section_header(ws, row, "Q1 FY26-27 HEADLINE PERFORMANCE (every cell links to 'Ratio Analysis'; Orange O Tec and Enterprises are each combined across their two branches)")
    row += 1
    row = write_q1_linked_row(ws, row, "Turnover", ratio_ws, r("Total Turnover"), NORMAL_FONT)
    row = write_q1_linked_row(ws, row, "Gross Profit", ratio_ws, r("Gross Profit"), NORMAL_FONT)
    row = write_q1_linked_row(ws, row, "Gross Profit %", ratio_ws, r("Gross Profit Margin %"), NORMAL_FONT, number_format="0.00")
    row = write_q1_profit_row(ws, row, "PAT", ratio_ws, r("PAT"))
    row = write_q1_linked_row(ws, row, "PAT %", ratio_ws, r("PAT Margin %"), NORMAL_FONT, number_format="0.00")
    row = write_q1_linked_row(ws, row, "Total Assets", ratio_ws, r("Total Assets"), NORMAL_FONT)
    row = write_q1_sum_row(ws, row, "Net Worth", ratio_ws, [r("Capital Account"), r("Reserves & Surplus")], NORMAL_FONT)
    row = write_q1_linked_row(ws, row, "Current Ratio", ratio_ws, r("Current Ratio"), NORMAL_FONT, number_format="0.00")
    row = write_q1_linked_row(ws, row, "Debtor Days", ratio_ws, r("Debtor Days"), NORMAL_FONT, number_format="0.00")
    row = write_q1_linked_row(ws, row, "Creditor Days", ratio_ws, r("Creditor Days"), NORMAL_FONT, number_format="0.00")

    row += 2
    row = write_q1_company_commentary(ws, row, ratio_ws, Q1_COMPANY_COMMENTARY)

    row += 2
    row = write_q1_merged_table(ws, row, "COST REDUCTION FOCUS & INTER-COMPANY TAX EFFICIENCY (directional analysis - validate with your tax advisor before acting)",
                                 Q1_COST_TAX_HEADER, Q1_COST_TAX_ROWS)

    row += 2
    row = write_q1_table(ws, row, "FUTURE TAX PLANNING - ADVICE TO MANAGEMENT (acting as tax partner; directional, needs CA/tax advisor confirmation)",
                          Q1_TAX_PLANNING_HEADER, Q1_TAX_PLANNING_ROWS, ncols=4)

    row += 2
    write_q1_action_plan(ws, row, action_plan_rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--timestamp", default=None,
                     help="Shared timestamp (format %s) so this script and build_mom_gl_detail.py "
                          "write to the exact same dated file. Defaults to now()." % TIMESTAMP_FORMAT)
    args = ap.parse_args()

    files_dir = Path(__file__).parent.parent.parent.parent.parent / "Files"
    timestamp = args.timestamp or datetime.now().strftime(TIMESTAMP_FORMAT)
    output_path = files_dir / f"{BASE_NAME}_{timestamp}.xlsx"

    # Snapshot any manual column/row grouping the user collapsed in Excel on the MOST RECENT
    # existing dated file (this run's own output_path is brand new and won't exist yet) - restored
    # just before saving. See "File naming and version history" in SKILL.md.
    old_snap = {}
    prev_path = find_latest_pl_bs_file(files_dir, exclude=output_path)
    if prev_path is not None:
        old_wb = openpyxl.load_workbook(prev_path)
        old_snap = snapshot_column_row_hidden(old_wb)
        old_wb.close()
        save_grouping_snapshot(old_snap, output_path)

    wb = openpyxl.Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    # Schedule sheets first (source of truth), so the Summary sheets below can reference them.
    ws_pl = wb.create_sheet(title="PL GL Detail")
    pl_row_map, pl_section_buckets, pl_section_detail = write_pl_sheet(ws_pl)

    ws_bs = wb.create_sheet(title="BS GL Detail")
    bs_row_map, bs_section_buckets, bs_section_detail = write_bs_sheet(ws_bs)

    ws_pl_summary = wb.create_sheet(title="PL Summary")
    write_pl_summary_sheet(ws_pl_summary, "PL GL Detail", pl_row_map, pl_section_buckets, pl_section_detail)

    ws_bs_summary = wb.create_sheet(title="BS Summary")
    write_bs_summary_sheet(ws_bs_summary, "BS GL Detail", bs_row_map, bs_section_buckets)

    ws_ratios = wb.create_sheet(title="Ratio Analysis")
    write_ratio_sheet(ws_ratios, pl_row_map, bs_row_map)

    ws_q1_report = wb.create_sheet(title="Q1 Performance Report")
    write_q1_performance_report_sheet(ws_q1_report, ws_ratios, Q1_ACTION_PLAN)

    ws_notes = wb.create_sheet(title="Notes")
    ws_notes.column_dimensions["A"].width = 105
    notes = [
        "NOTES ON THIS WORKBOOK",
        "",
        "Source: Tally Prime (localhost:9010), all 5 companies open as of 2026-07-15.",
        "Period: 01-Apr-2026 to 30-Jun-2026 (Q1 FY26-27). Only non-zero ledgers shown.",
        "",
        "P&L layout: proper Trading + P&L account -",
        "  Turnover (Sales Accounts + Direct Incomes)",
        "  less Cost of Sales (Purchase Accounts + Direct Expenses + Stock Valuation)",
        "  = Gross Profit",
        "  add Other Income (Indirect Incomes)",
        "  less Operating Expenses (Indirect Expenses, excluding Interest/Finance Cost)",
        "  = EBITDA",
        "  less Finance Cost (Interest Cost / Finance Cost sub-groups, pulled out of",
        "    Indirect Expenses) and Depreciation",
        "  = PBT (Profit Before Tax)",
        "  less Tax",
        "  = PAT (Net Profit)",
        "IMPORTANT: no Depreciation or Tax-provision ledgers exist in this quarter's",
        "closing balances (these are normally booked only at year-end via journal in",
        "Tally), so those two rows show Rs 0 and PBT = PAT this quarter. Shown",
        "explicitly rather than hidden - update manually if you book a provisional",
        "entry for either.",
        "",
        "BS layout: two sections - LIABILITIES (Sources of Funds: Capital, Reserves,",
        "Secured/Unsecured Loans, Bank OD, Duties & Taxes, Provisions, Creditors,",
        "Current Liabilities...) and ASSETS (Application of Funds: Fixed Assets,",
        "Investments, Current Assets incl. Closing Stock, Bank, Cash, Loans & Advances,",
        "Debtors...), each with its own TOTAL row, plus a Reconciliation Gap row at the",
        "bottom (Liabilities + Assets, where Assets carry a negative sign by Tally's",
        "convention, so a true balance shows ~0).",
        "",
        "Column A (Statement Section) marks the always-visible top-level rows (Turnover,",
        "Cost of Sales, Gross Profit, EBITDA, PBT, PAT / Liabilities, Assets). Column B",
        "(Grouping) is Tally's primary group. Column C (Sub Grouping) is Tally's",
        "immediate parent group per ledger. Three levels of row outline (+/- controls",
        "in the left margin, or Data > Group):",
        "  - Level 2 (innermost): individual GL rows",
        "  - Level 1: Sub Grouping subtotal rows",
        "  - Level 0: Grouping TOTAL rows (Statement Section / computed rows are also",
        "    level 0 and always stay visible)",
        "",
        "Because each company has its own chart of accounts, most rows carry a value in",
        "only one company's column - that's expected, not a gap. Where a GL name",
        "genuinely repeats across companies (e.g. 'ROUND OFF', 'Opening Stock'), it",
        "lands on one shared row.",
        "",
        "Total of Orange O Tec = ORANGE O TEC PRIVATE LIMITED + ...-NOIDA.",
        "Total of Enterprises = ORANGE O TEC ENTERPRISES PVT LTD + ...-NOIDA.",
        "Grand Total = Total of Orange O Tec + Total of Enterprises + COLORIX.",
        "",
        "Stock Valuation: Closing/Opening Stock is not a ledger account in Tally (it's",
        "valued via the inventory/stock-summary report), so it never appears in the raw",
        "ledger pull. Added back explicitly: PL sheet includes it under Cost of Sales",
        "(Opening Stock as a cost, Closing Stock as a credit) so PAT equals true Net",
        "Profit, not just revenue less purchases. BS sheet adds Closing Stock under",
        "Current Assets so the Reconciliation Gap row is meaningful.",
        "",
        "BS residuals after the Closing Stock adjustment (small, disclosed, not hidden):",
        "  - ORANGE O TEC PRIVATE LIMITED-NOIDA: fully reconciled (0.00)",
        "  - ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA: fully reconciled (0.00)",
        "  - ORANGE O TEC PRIVATE LIMITED: RESOLVED (2026-07-16) - the residual was 'Sales Bills",
        "    to Make' (goods dispatched, formal sales invoice not yet raised - a real Tally",
        "    receivable computed from pending Delivery Notes, not a ledger/group, so it never",
        "    appears via the API). Confirmed against Tally's own Balance Sheet -> Current Assets",
        "    -> Sundry Debtors -> Group Summary screen: Rs 3,54,79,998.84. Added explicitly as",
        "    its own line (Sales Accounts in P&L, Sundry Debtors in BS) - both statements now",
        "    reconcile to the rupee against Tally's native figures.",
        "  - ORANGE O TEC ENTERPRISES PVT LTD: RESOLVED (2026-07-17) - the residual was",
        "    'Purchase Bills to Come' (goods/services received from a related party, formal",
        "    purchase invoice not yet booked) - confirmed against Tally's own Balance Sheet ->",
        "    Current Liabilities -> Sundry Creditors -> Group Summary screen (Rs 6,78,615, shown",
        "    as a contra reduction within Sundry Creditors) AND independently confirmed against",
        "    Tally's own Purchase Group Summary screen, where the identical Rs 6,78,615 appears",
        "    as its own line reducing Purchase Accounts. Added to BOTH statements accordingly:",
        "    Purchase Accounts in P&L (reducing Cost of Sales, raising PAT by Rs 6,78,615) and",
        "    Sundry Creditors in BS (reducing the liability by the same amount) - mirroring how",
        "    Sales Bills to Make was added to both statements for ORANGE O TEC PRIVATE LIMITED.",
        "    Both statements now reconcile to the rupee against Tally's native figures.",
        "  - COLORIX DIGITAL PRINTING SOLUTIONS LLP: RESOLVED (2026-07-17) - the residual was",
        "    'Unadjusted Forex Gain/Loss', a standalone line in Tally's own Balance Sheet ->",
        "    Sources of Funds (not tied to any ledger/group master) - confirmed Rs 45,600 against",
        "    Tally's screen. Added as its own BS-only line; Reconciliation Gap row now nets to",
        "    0.00. Same open question as above applies in principle, but the PL-vs-BS cross-check",
        "    for Colorix shows no unexplained gap once the ledger's Opening Balance component is",
        "    accounted for, so no P&L-side action is indicated here.",
        "",
        "SCHEDULE / SUMMARY ARCHITECTURE:",
        "'PL GL Detail' and 'BS GL Detail' are the schedules - the full GL-level source of",
        "truth, including every subtotal. 'PL Summary' and 'BS Summary' are short sheets whose",
        "cells are Excel formulas linking directly to the matching total row on the schedule",
        "sheets (e.g. ='PL GL Detail'!E67) - they are NOT independently recomputed. Edit a",
        "number on a schedule sheet and the corresponding Summary cell recalculates. Each",
        "grouping heading (TURNOVER, COST OF SALES, ...) is followed by one row per Sub Group",
        "(e.g. Sales Accounts, Direct Incomes) in the Sub Group column, each linked to that",
        "sub-group's own TOTAL row on the schedule sheet, then a bold section total row.",
        "Exception: LESS: OPERATING EXPENSES goes one level deeper - since it is a single large",
        "bucket (Indirect Expenses), it lists all of Tally's Sub Grouping-level expense heads",
        "(Employee Salary, Rent, Travelling, etc.), each linked to its own Subtotal row on the",
        "schedule sheet, largest first.",
        "",
        "BRANCH TRANSACTIONS: ledgers with 'BRANCH' in the name (inter-company trade between",
        "the paired entities, e.g. ORANGE O TEC PRIVATE LIMITED <-> ...-NOIDA) are pulled out of",
        "the normal Sales Accounts / Purchase Accounts buckets and shown in their own memo",
        "section for visibility, so Total Turnover and Total Cost of Sales above are net of them -",
        "but the net branch amount IS added back into Gross Profit (and everything below it:",
        "EBITDA, PBT, PAT), so each individual company's PAT matches that company's own Tally P&L",
        "screen exactly. Since the trade is reciprocal between the two paired entities, this has",
        "~Nil net effect at the family (Orange O Tec / Enterprises) and Grand Total level either",
        "way - branch inclusion only changes how PAT is split between the two individual companies,",
        "never the family or Grand Total figure.",
        "",
        "RELATED PARTY TRANSACTIONS: 'RELATED PARTY SALES' / 'RELATED PARTY PURCHASE' ledgers",
        "stay INSIDE Turnover/Cost of Sales as normal (so Gross Profit, EBITDA, PBT and PAT",
        "above are the true, full-basis figures, unaffected by this adjustment). A memo section",
        "after PAT shows these amounts, and a final 'PAT (Net of Related Party Transactions)'",
        "row shows PAT with their net contribution removed - useful for judging profit quality",
        "excluding related-party dealings, without altering the headline PAT above it.",
        "",
        "RATIO ANALYSIS: three sections stacked on each sheet, top to bottom - (1) RATIO SUMMARY,",
        "the headline numbers, each cell a same-sheet link down to (3); (2) RATIO WORKINGS -",
        "COMPONENT FIGURES, every raw P&L/BS figure any ratio needs, pulled once per company",
        "column directly from PL GL Detail / BS GL Detail; (3) RATIO WORKINGS - CALCULATIONS, the",
        "actual ratio formulas, built only from same-sheet references into (2) - never a",
        "cross-sheet reference. Scroll down on either ratio sheet to see every input figure behind",
        "any number at the top, without needing to open another sheet. Profitability, Liquidity,",
        "Leverage, Efficiency. Day-count ratios (Debtor Days, Creditor Days) use the 91-day quarter",
        "itself, not an annualised estimate. Quick Ratio excludes the 'Current Assets' bucket",
        "(which contains Closing Stock) and counts only Bank, Cash, Debtors, Loans & Advances and",
        "Deposits as liquid assets.",
        "This is the QUARTERLY view (Q1 Total). See 'MOM Ratio Analysis' for the same ratio set",
        "computed independently for each of Apr-26/May-26/Jun-26 (each month's own actual",
        "day-count used for Debtor/Creditor Days, not the 91-day quarter) - both sheets share one",
        "ratio-definition function so they can never drift apart; a new ratio added to one always",
        "appears in the other on the next rebuild.",
        "",
        "Cash Flow and Fund Flow Statements (which need the 31-Mar-2026 opening Balance Sheet) are",
        "planned as a follow-up once that data is fetched from Tally.",
    ]
    for i, line in enumerate(notes, start=1):
        cell = ws_notes.cell(i, 1, value=line)
        cell.font = Font(bold=True, size=12) if i == 1 else Font(size=10, name="Calibri")
        cell.alignment = Alignment(wrap_text=True)

    wb._sheets = [wb[name] for name in
                  ["PL Summary", "BS Summary", "Ratio Analysis", "Q1 Performance Report",
                   "PL GL Detail", "BS GL Detail", "Notes"]]

    restore_column_row_hidden(wb, old_snap)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(output_path)
    except PermissionError:
        raise RuntimeError(f"Cannot save to {output_path!r} - file may be open in Excel.\nClose the file and retry.")

    print(f"File saved: {output_path}")


if __name__ == "__main__":
    main()
