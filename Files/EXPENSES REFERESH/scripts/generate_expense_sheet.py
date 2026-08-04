"""
Tool: generate_expense_sheet.py
Fetches Direct and Indirect Expense data from Tally for all 4 companies
and writes the raw-data sheet of the Expense Sheet workbook (23 columns).

This produces ONLY the "RAW DATA" sheet. Run build_analysis_sheets.py afterwards
to build the Summary/Direct/Indirect Exp/Advertisement/Direct by Component/
per-main-group indirect sheets and Index on top of it.

Usage (run from the EXPENSES REFERESH folder):
    python scripts/generate_expense_sheet.py --from 20260401 --to 20260630
    python scripts/generate_expense_sheet.py --from 20260401 --to 20260630 --out "custom.xlsx"

Output columns (RAW DATA):
    Company | EXP TYPE | MAIN GROUP | SUB-GROUP | SUB-SUB GROUP | TYPES |
    DATE | LEDGER NAME | VCH. TYPE | VCH NO. | NET AMOUNT | MONTH |
    LOCATION | NARRATION | SECTIONS | REMARK-1 | REMARK-2 |
    REMARK-3 (Specified period) | EMPLOYEE NAME | DEPARTMENT | Financial Quarter |
    Company Name | COMPONENT CATEGORY
"""

import sys
import argparse
import os
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from fetch_tally_expenses import fetch_expense_rows, get_open_companies

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip install openpyxl")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Company resolution - which Tally company FILE to use for each of the 5
# operating entities depends on the period, because Tally periodically
# "splits" a company into a new data file per financial year. The canonical
# mapping of period -> correct Tally company file lives in
# Files/TALLY Company_List.xlsx (columns "26-27"/"25-26"/"24-25", a "1"
# means "use this file for this period"; "Not To USE" rows are group/
# consolidation shells, always excluded). Do not hardcode a single fixed
# company list here - resolve it per period via companies_for_period().
#
# _ENTITY_META maps every known Tally company-file name (across all periods
# found in that workbook) to the stable display metadata for the entity it
# belongs to (company_col/location/company_name stay the same across periods
# for a given entity; only the underlying tally_name changes).
# ---------------------------------------------------------------------------

_ENTITY_META = {
    "ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)":
        dict(key="OOTPL", company_col="Otec-Surat", location="SURAT", company_name="Otec"),
    "ORANGE O TEC PRIVATE LIMITED (01-04-24TO31-03-25)":
        dict(key="OOTPL", company_col="Otec-Surat", location="SURAT", company_name="Otec"),
    "ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)":
        dict(key="OOTPL_NOIDA", company_col="Otec-Noida", location="NOIDA", company_name="Otec"),
    "ORANGE O TEC PRIVATE LIMITED-NOIDA - (from 1-Apr-24)":
        dict(key="OOTPL_NOIDA", company_col="Otec-Noida", location="NOIDA", company_name="Otec"),
    "ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)":
        dict(key="OOTEPL", company_col="Ent-Surat", location="SURAT", company_name="Ent"),
    "ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2024-26)":
        dict(key="OOTEPL", company_col="Ent-Surat", location="SURAT", company_name="Ent"),
    "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27":
        dict(key="OOTEPL_NOIDA", company_col="Ent-Noida", location="NOIDA", company_name="Ent"),
    "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA - FY25-26":
        dict(key="OOTEPL_NOIDA", company_col="Ent-Noida", location="NOIDA", company_name="Ent"),
    "COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)":
        dict(key="COLORIX", company_col="Colorix", location="SURAT", company_name="Colorix"),
}

_PERIOD_COL_MAP = {"2627": "26-27", "2526": "25-26", "2425": "24-25"}


def _company_list_path() -> Path:
    candidates = [
        Path.cwd() / "Files" / "TALLY Company_List.xlsx",
        Path(__file__).resolve().parents[3] / "Files" / "TALLY Company_List.xlsx",
    ]
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(
        "Cannot find Files/TALLY Company_List.xlsx. Run from the project root."
    )


def _read_company_list_rows() -> list[tuple]:
    """Returns the raw rows of Files/TALLY Company_List.xlsx's "Company List"
    sheet. Falls back to reading via a running Excel COM instance if the file
    is open (and therefore locked for plain file access) - this reference
    file is small and likely to be left open, so don't force the user to
    close it before every refresh."""
    path = _company_list_path()
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
        rows = [tuple(r) for r in wb.active.iter_rows(values_only=True)]
        wb.close()
        return rows
    except PermissionError:
        import win32com.client as win32
        try:
            excel = win32.GetActiveObject("Excel.Application")
        except Exception:
            raise PermissionError(
                f"{path} is open (locked) but no running Excel instance was found to "
                "read it from. Close the file or reopen it and retry."
            )
        wb_com = None
        for w in excel.Workbooks:
            if w.Name == path.name:
                wb_com = w
                break
        if wb_com is None:
            raise PermissionError(
                f"{path} is locked but not open in the running Excel instance. "
                "Close whatever has it locked and retry."
            )
        ws_com = wb_com.Worksheets(1)
        used = ws_com.UsedRange
        n_rows, n_cols = used.Rows.Count, used.Columns.Count
        return [
            tuple(ws_com.Cells(r, c).Value for c in range(1, n_cols + 1))
            for r in range(1, n_rows + 1)
        ]


def companies_for_period(period_code: str) -> list[dict]:
    """Resolve the 5-company list for a period code (2627/2526/2425) from
    Files/TALLY Company_List.xlsx, in the same {tally_name, key, company_col,
    location, company_name} shape the rest of this module expects."""
    col_name = _PERIOD_COL_MAP.get(str(period_code))
    if not col_name:
        raise ValueError(f"Unknown period code '{period_code}'. Use 2627, 2526, or 2425.")

    headers = None
    col_idx = None
    tally_names = []
    for row in _read_company_list_rows():
        if headers is None:
            if row[0] == "Company Name":
                headers = list(row)
                col_idx = headers.index(col_name)
            continue
        company_name = row[0]
        comment = row[3] if len(row) > 3 else None
        if not company_name or comment == "Not To USE":
            continue
        val = row[col_idx] if col_idx is not None and len(row) > col_idx else None
        if val == 1:
            tally_names.append(str(company_name).strip())

    result = []
    for name in tally_names:
        meta = _ENTITY_META.get(name)
        if meta is None:
            raise ValueError(
                f"'{name}' is marked for period {period_code} in TALLY Company_List.xlsx "
                "but is not recognised in generate_expense_sheet._ENTITY_META. "
                "Add it there (company_col/location/company_name for the entity it belongs "
                "to) before refreshing."
            )
        result.append({"tally_name": name, **meta})
    return result


# ---------------------------------------------------------------------------
# Excel column definitions
# ---------------------------------------------------------------------------

COLUMNS = [
    "Company",
    "EXP TYPE",
    "MAIN GROUP",
    "SUB-GROUP",
    "SUB-SUB GROUP",
    "TYPES",
    "DATE",
    "LEDGER NAME",
    "VCH. TYPE",
    "VCH NO.",
    "NET AMOUNT",
    "MONTH",
    "LOCATION",
    "NARRATION",
    "SECTIONS",
    "REMARK-1",
    "REMARK-2",
    "REMARK-3 (Specified period)",
    "EMPLOYEE NAME",
    "DEPARTMENT",
    "Financial Quarter",
    "Company Name",
    "COMPONENT CATEGORY",
]

# Column letters for formula references (computed once after import)
_DATE_COL_LETTER  = get_column_letter(COLUMNS.index("DATE")  + 1)   # G
_MONTH_COL_LETTER = get_column_letter(COLUMNS.index("MONTH") + 1)   # L

_HEADER_FILL  = PatternFill("solid", fgColor="FFC000")  # Excel Gold (matches reference)
_HEADER_FONT  = Font(bold=True, color="000000", size=10)
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
_THIN         = Side(style="thin", color="CCCCCC")
_BORDER       = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

COL_WIDTHS = {
    "Company":                    16,
    "EXP TYPE":                   10,
    "MAIN GROUP":                 30,
    "SUB-GROUP":                  30,
    "SUB-SUB GROUP":              30,
    "TYPES":                      34,
    "DATE":                       13,
    "LEDGER NAME":                40,
    "VCH. TYPE":                  24,
    "VCH NO.":                    16,
    "NET AMOUNT":                 14,
    "MONTH":                      12,
    "LOCATION":                   10,
    "NARRATION":                  48,
    "SECTIONS":                   20,
    "REMARK-1":                   20,
    "REMARK-2":                   20,
    "REMARK-3 (Specified period)": 24,
    "EMPLOYEE NAME":              22,
    "DEPARTMENT":                 18,
    "Financial Quarter":          14,
    "Company Name":               14,
    "COMPONENT CATEGORY":         20,
}


def write_excel(rows: list[dict], output_path: str):
    """Write all expense rows to a formatted Excel workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RAW DATA"

    # Header row
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font      = _HEADER_FONT
        cell.fill      = _HEADER_FILL
        cell.alignment = _HEADER_ALIGN
        cell.border    = _BORDER

    ws.row_dimensions[1].height = 30

    # Data rows
    for row_idx, row in enumerate(rows, start=2):
        for col_idx, col_name in enumerate(COLUMNS, start=1):

            if col_name == "MONTH":
                # Excel formula: format DATE column as "MAY-25" etc.
                cell = ws.cell(
                    row=row_idx, column=col_idx,
                    value=f'=UPPER(TEXT({_DATE_COL_LETTER}{row_idx},"MMM-YY"))',
                )

            elif col_name == "Financial Quarter":
                # Excel formula: compute Q1-Q4 from MONTH column
                cell = ws.cell(
                    row=row_idx, column=col_idx,
                    value=(
                        f'="Q"&INT(MOD(MONTH(DATEVALUE("1-"&'
                        f'{_MONTH_COL_LETTER}{row_idx}))-4,12)/3)+1'
                    ),
                )

            else:
                val  = row.get(col_name, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=val)

                if col_name == "DATE" and isinstance(val, datetime):
                    cell.number_format = "DD-MM-YYYY"

                if col_name == "NET AMOUNT" and isinstance(val, (int, float)):
                    cell.number_format = "#,##0.00"

            cell.border    = _BORDER
            cell.alignment = Alignment(vertical="center")

    # Column widths
    for col_idx, col_name in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = COL_WIDTHS.get(col_name, 14)

    # Freeze top row and enable auto-filter
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    wb.save(output_path)
    print(f"\nSaved: {output_path}  ({len(rows)} rows)")


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def _check_tally():
    """Quick TCP connectivity check before starting the full run."""
    import socket
    from fetch_tally_expenses import TALLY_HOST, TALLY_PORT
    try:
        s = socket.create_connection((TALLY_HOST, int(TALLY_PORT)), timeout=5)
        s.close()
    except OSError:
        print(f"ERROR: Tally is not running or Gateway is not enabled on {TALLY_HOST}:{TALLY_PORT}.")
        print(f"       Please open Tally, enable the Gateway ({TALLY_HOST}:{TALLY_PORT}), and try again.")
        sys.exit(1)


def fetch_all_rows(from_date: str, to_date: str, companies: list[dict] = None,
                    period: str = "2627", open_names_lower: set = None) -> list[dict]:
    """Fetch expense rows for every open company in `companies` (default:
    resolved from Files/TALLY Company_List.xlsx for `period`) over one shared
    date range. Returns the combined row list without writing anything -
    callers decide how to combine/write it (a single write_excel() call, or
    merged with rows from other date ranges for a one-off historical
    backfill)."""
    companies = companies if companies is not None else companies_for_period(period)
    if open_names_lower is None:
        open_companies = get_open_companies()
        open_names_lower = {n.lower() for n in open_companies}

    all_rows = []
    for cfg in companies:
        if cfg["tally_name"].lower() not in open_names_lower:
            print(f"SKIPPED (not open): {cfg['key']}  ({cfg['tally_name']})")
            continue

        print(f"\nFetching: {cfg['key']}  ({cfg['tally_name']})  [{from_date} to {to_date}]")
        company_cfg = {
            "company_col":  cfg["company_col"],
            "location":     cfg["location"],
            "company_name": cfg["company_name"],
        }
        try:
            rows = fetch_expense_rows(
                company=cfg["tally_name"],
                company_cfg=company_cfg,
                from_date=from_date,
                to_date=to_date,
            )
        except RuntimeError as e:
            print(f"  ERROR: {e}")
            continue

        all_rows.extend(rows)
        print(f"  Rows added: {len(rows)}")

    return all_rows


def run(from_date: str, to_date: str, output_path: str, period: str = "2627"):
    print("=" * 60)
    print("Expense Sheet Generator")
    print(f"Date range : {from_date} to {to_date}")
    print(f"Period     : {period} (Files/TALLY Company_List.xlsx)")
    print(f"Output     : {output_path}")
    print("=" * 60)

    _check_tally()
    print("Tally connection: OK\n")

    try:
        open_companies = get_open_companies()
    except RuntimeError as e:
        print(f"ERROR: Could not list open companies: {e}")
        sys.exit(1)

    open_names_lower = {n.lower() for n in open_companies}
    print(f"Open companies in Tally ({len(open_companies)}):")
    for n in open_companies:
        print(f"  - {n}")

    companies = companies_for_period(period)
    print(f"\nCompanies required for period {period}:")
    for cfg in companies:
        print(f"  - {cfg['company_col']}: {cfg['tally_name']}")

    required_names = {c["tally_name"] for c in companies}
    missing = [c["tally_name"] for c in companies if c["tally_name"].lower() not in open_names_lower]
    extra = [n for n in open_companies if n not in required_names]
    if missing:
        print(f"\nERROR: Period {period} requires companies that are NOT currently open in Tally:")
        for n in missing:
            print(f"  X missing (please OPEN in Tally): {n}")
        if extra:
            print("\n  Currently open but not needed for this period (you can CLOSE these):")
            for n in extra:
                print(f"  - {n}")
        print("\nOpen the missing companies in Tally, then re-run.")
        sys.exit(1)
    if extra:
        print(f"\nNote: {len(extra)} open company(ies) are not part of period {period} and will be skipped:")
        for n in extra:
            print(f"  - {n}")
    print()

    all_rows = fetch_all_rows(from_date, to_date, companies=companies, open_names_lower=open_names_lower)

    print(f"\nTotal rows across all companies: {len(all_rows)}")

    if not all_rows:
        print(
            "\nNo expense data found for the given date range. "
            "Check that companies are open and there are expense vouchers in this period."
        )
        sys.exit(0)

    write_excel(all_rows, output_path)
    print("Done.")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate Expense Sheet from Tally")
    parser.add_argument("--from", dest="from_date", required=True,
                        help="Start date YYYYMMDD (e.g. 20250401)")
    parser.add_argument("--to",   dest="to_date",   required=True,
                        help="End date YYYYMMDD (e.g. 20260331)")
    parser.add_argument("--out",  dest="output_path", default=None,
                        help="Output Excel path (default: ./Expense Sheet <timestamp>.xlsx)")
    parser.add_argument("--period", dest="period", default="2627",
                        help="Period code 2627/2526/2425 - selects companies via "
                             "Files/TALLY Company_List.xlsx (default: 2627)")
    args = parser.parse_args()

    if args.output_path is None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        args.output_path = f"Expense Sheet {timestamp}.xlsx"

    run(args.from_date, args.to_date, args.output_path, period=args.period)
