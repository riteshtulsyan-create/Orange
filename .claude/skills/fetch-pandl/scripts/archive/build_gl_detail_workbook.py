"""
Builds Files/PL_BS_GL_Detail.xlsx: full GL-level detail (Grouping | GL Name | Balance)
with a subtotal row per Grouping, for all 5 companies open in Tally, stacked vertically
in two sheets (PL, BS). Period: 01-Apr-2026 to 30-Jun-2026.

Source data: CSVs pre-fetched via the tally-prime MCP tools (profit-loss, balance-sheet,
chart-of-accounts, query-database with non-zero filter), saved to the gl_data folder next
to this script's scratch location. Each CSV has columns: primary_group, ledger_name,
closing_balance, pre-sorted by primary_group then |closing_balance| descending.
"""

import csv
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

GL_DATA_DIR = Path(
    r"C:\Users\orang\AppData\Local\Temp\claude\c--Users-orang-OneDrive---Orange-O-Tec-Private-Limited-00-Claude-Project-PL-and-BS"
    r"\1692684d-78ab-405d-8ee7-bf023df8abe6\scratchpad\gl_data"
)

PERIOD_LABEL = "01-Apr-2026 to 30-Jun-2026"

COMPANIES = [
    ("ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)", "ORANGE_O_TEC_PRIVATE_LIMITED"),
    ("ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)", "PRIVATE_LIMITED_NOIDA"),
    ("ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)", "ENTERPRISES_PVT_LTD"),
    ("ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27", "ENTERPRISES_NOIDA"),
    ("COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)", "COLORIX"),
]

# Stock is not a ledger account (Tally values it via inventory, not a GL balance), so it never
# appears in the ledger-level pull. Without it, PL "GRAND TOTAL" understates true Net Profit
# (Purchases show as full expense, no stock adjustment) and BS "GRAND TOTAL" doesn't reconcile
# to zero. Add it back explicitly, sourced from the stock-summary tool (fetched earlier).
STOCK = {
    "ORANGE_O_TEC_PRIVATE_LIMITED": {"opening": 292508691.08, "closing": 284571181.51},
    "PRIVATE_LIMITED_NOIDA": {"opening": 32685329.56, "closing": 29258524.91},
    "ENTERPRISES_PVT_LTD": {"opening": 71324050.22, "closing": 91179012.34},
    "ENTERPRISES_NOIDA": {"opening": 32959492.39, "closing": 37902468.1},
    "COLORIX": {"opening": 105648, "closing": 105648},
}

# Sensible display order for primary groups (income first, then costs / assets then liabilities-ish)
PL_GROUP_ORDER = ["Sales Accounts", "Direct Incomes", "Purchase Accounts", "Stock Valuation (Inventory)",
                   "Direct Expenses", "Indirect Incomes", "Indirect Expenses"]
BS_GROUP_ORDER = ["Capital Account", "Reserves & Surplus", "Secured Loans", "Unsecured Loans",
                   "Bank OD A/c", "Duties & Taxes", "Provisions", "Sundry Creditors",
                   "Current Liabilities", "Loans (Liability)", "Branch / Divisions",
                   "Suspense A/c", "", "Fixed Assets", "Investments", "Current Assets",
                   "Bank Accounts", "Bank Account", "Cash-in-hand", "Deposits (Asset)",
                   "Loans & Advances (Asset)", "Sundry Debtors", "Misc. Expenses (ASSET)"]
BS_GROUP_LABEL = {"": "Profit & Loss A/c (current)", "Bank Account": "Bank Accounts"}


def load_csv(path: Path):
    rows = []
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append((r["primary_group"], r["ledger_name"], float(r["closing_balance"])))
    return rows


def ordered_groups(rows, preferred_order):
    present = []
    seen = set()
    for pg, _, _ in rows:
        if pg not in seen:
            seen.add(pg)
            present.append(pg)
    ordered = [g for g in preferred_order if g in seen]
    remaining = sorted(g for g in present if g not in preferred_order)
    return ordered + remaining


TITLE_FONT = Font(bold=True, size=13, name="Calibri")
COMPANY_FONT = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
COMPANY_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(bold=True, size=10, name="Calibri")
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
GROUP_FONT = Font(bold=True, size=10, name="Calibri")
GROUP_FILL = PatternFill("solid", fgColor="F2F2F2")
NORMAL_FONT = Font(size=10, name="Calibri")
SUBTOTAL_FONT = Font(bold=True, italic=True, size=10, name="Calibri")
GRANDTOTAL_FONT = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
GRANDTOTAL_FILL = PatternFill("solid", fgColor="203864")
NUM_FORMAT = "#,##0.00"


def write_statement_sheet(ws, title: str, group_order: list, group_label: dict, file_prefix: str,
                            stat_type: str):
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 20

    row = 1
    ws.cell(row, 1, value=title).font = TITLE_FONT
    row += 1
    ws.cell(row, 1, value=f"Period: {PERIOD_LABEL}").font = NORMAL_FONT
    if stat_type == "pl":
        ws.cell(row, 3, value="Includes Stock Valuation (not a ledger; see Notes)").font = Font(italic=True, size=9, color="806000")
    row += 2

    for company_name, file_key in COMPANIES:
        csv_path = GL_DATA_DIR / f"{file_prefix}_{file_key}.csv"
        rows = load_csv(csv_path)

        stock = STOCK[file_key]
        if stat_type == "pl":
            rows.append(("Stock Valuation (Inventory)", "Opening Stock", -stock["opening"]))
            rows.append(("Stock Valuation (Inventory)", "Closing Stock", stock["closing"]))
        elif stat_type == "bs":
            rows.append(("Current Assets", "Closing Stock (Inventory)", -stock["closing"]))

        # Company header band
        ws.cell(row, 1, value=company_name).font = COMPANY_FONT
        for c in (1, 2, 3):
            ws.cell(row, c).fill = COMPANY_FILL
        row += 1

        # Column header
        ws.cell(row, 1, value="Grouping").font = HEADER_FONT
        ws.cell(row, 2, value="GL Name").font = HEADER_FONT
        ws.cell(row, 3, value="Balance").font = HEADER_FONT
        for c in (1, 2, 3):
            ws.cell(row, c).fill = HEADER_FILL
        row += 1

        groups_present = ordered_groups(rows, group_order)
        grand_total = 0.0
        for grp in groups_present:
            grp_rows = [(nm, bal) for pg, nm, bal in rows if pg == grp]
            # already sorted by |balance| desc from the source query
            display_grp = group_label.get(grp, grp) if grp != "" else group_label.get("", "(no group)")
            group_sum = sum(bal for _, bal in grp_rows)
            grand_total += group_sum

            first = True
            for gl_name, bal in grp_rows:
                ws.cell(row, 1, value=display_grp if first else None).font = NORMAL_FONT
                ws.cell(row, 2, value=gl_name).font = NORMAL_FONT
                cell = ws.cell(row, 3, value=bal)
                cell.number_format = NUM_FORMAT
                cell.font = NORMAL_FONT
                first = False
                row += 1

            # subtotal row
            ws.cell(row, 1, value=None)
            ws.cell(row, 2, value=f"Subtotal: {display_grp}").font = SUBTOTAL_FONT
            cell = ws.cell(row, 3, value=group_sum)
            cell.number_format = NUM_FORMAT
            cell.font = SUBTOTAL_FONT
            for c in (1, 2, 3):
                ws.cell(row, c).fill = GROUP_FILL
            row += 1

        # grand total row
        ws.cell(row, 2, value=f"GRAND TOTAL - {company_name}").font = GRANDTOTAL_FONT
        cell = ws.cell(row, 3, value=grand_total)
        cell.number_format = NUM_FORMAT
        cell.font = GRANDTOTAL_FONT
        for c in (1, 2, 3):
            ws.cell(row, c).fill = GRANDTOTAL_FILL
        row += 3  # blank separator rows before next company

    ws.freeze_panes = "A4"
    return


def main():
    wb = openpyxl.Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    ws_pl = wb.create_sheet(title="PL GL Detail")
    write_statement_sheet(ws_pl, "Profit & Loss A/c - GL Detail", PL_GROUP_ORDER, {}, "pl", "pl")

    ws_bs = wb.create_sheet(title="BS GL Detail")
    write_statement_sheet(ws_bs, "Balance Sheet - GL Detail", BS_GROUP_ORDER, BS_GROUP_LABEL, "bs", "bs")

    ws_notes = wb.create_sheet(title="Notes")
    ws_notes.column_dimensions["A"].width = 105
    notes = [
        "NOTES ON THIS WORKBOOK",
        "",
        "Source: Tally Prime (localhost:9010), all 5 companies open as of 2026-07-15.",
        "Period: 01-Apr-2026 to 30-Jun-2026 (Q1 FY26-27).",
        "Only ledgers with a non-zero closing balance for the period are listed.",
        "",
        "Layout: each sheet lists every company's ledgers grouped by Tally's top-level",
        "primary group (Grouping), with the individual GL name and balance, sorted by",
        "balance magnitude (largest first) within each group, followed by a subtotal row",
        "per group and a GRAND TOTAL row per company.",
        "",
        "Stock Valuation: Closing/Opening Stock is not a ledger account in Tally (it is",
        "valued via the inventory/stock-summary report, not a GL balance), so it never",
        "appears in the raw ledger pull. We added it back explicitly:",
        "  - PL sheet: a 'Stock Valuation (Inventory)' group with Opening Stock (shown as",
        "    a cost) and Closing Stock (shown as a credit) - this is what makes the GRAND",
        "    TOTAL equal true Net Profit instead of just revenue less purchases.",
        "  - BS sheet: Closing Stock is added as a line under 'Current Assets'.",
        "",
        "BS reconciliation: with Closing Stock included, the BS GRAND TOTAL for each",
        "company should be ~0 (Sources of funds = Application of funds, in Tally's sign",
        "convention where assets are negative and liabilities/capital are positive). Two",
        "companies reconcile exactly; three have a small residual shown in the GRAND TOTAL",
        "row itself (not hidden), consistent with the earlier PL_BS_Combined_Output.xlsx:",
        "  - ORANGE O TEC PRIVATE LIMITED-NOIDA: 0.00 (fully reconciled)",
        "  - ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA: 0.00 (fully reconciled)",
        "  - ORANGE O TEC ENTERPRISES PVT LTD: residual Rs 6,78,615 (~0.2% of total)",
        "  - ORANGE O TEC PRIVATE LIMITED: residual Rs 3,54,79,998.84 (~3% of total) -",
        "    this exact figure matches 'Sales Bills to Make' from a prior report, likely",
        "    a goods-in-transit/billing adjustment not captured as a standard ledger",
        "  - COLORIX DIGITAL PRINTING SOLUTIONS LLP: residual ~Rs -45,600 (small)",
        "",
        "Recommend cross-checking against Tally's native reports before relying on this",
        "for filings or board reporting, particularly for ORANGE O TEC PRIVATE LIMITED's",
        "Balance Sheet given the larger residual.",
    ]
    for i, line in enumerate(notes, start=1):
        cell = ws_notes.cell(i, 1, value=line)
        cell.font = Font(bold=True, size=12) if i == 1 else Font(size=10, name="Calibri")
        cell.alignment = Alignment(wrap_text=True)

    output_path = Path(__file__).parent.parent.parent.parent.parent / "Files" / "PL_BS_GL_Detail.xlsx"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(output_path)
    except PermissionError:
        raise RuntimeError(f"Cannot save to {output_path!r} - file may be open in Excel.\nClose the file and retry.")

    print(f"File saved: {output_path}")


if __name__ == "__main__":
    main()
