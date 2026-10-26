"""
build_cfo_pack.py - the CFO Pack: the six MOM P&L / BS / Ratio sheets of one PL_BS_GL_Detail
workbook, plus month-on-month Cash Flow and Fund Flow (CFS and SFS), an Opening Balance Sheet at
31-Mar-26, a CFO Analysis sheet, and print sheets that Excel exports to a board-ready PDF.

    python build_cfo_pack.py                                   # newest PL_BS_GL_Detail_*.xlsx
    python build_cfo_pack.py --source "Files\\Orange FS\\PL_BS_GL_Detail_2026-10-16_1523.xlsx"
    python build_cfo_pack.py --no-pdf                          # skip the Excel/PDF step

Output (never overwrites): Files\\Results\\CFO_Pack_PL_BS_CF_FF_<YYYY-MM-DD_HHMM>.xlsx / .pdf

WHERE EACH NUMBER COMES FROM
  * The six MOM sheets are COPIED from the source workbook, formulas and formatting intact. They
    reference only each other (checked 2026-10-24), so dropping every other sheet breaks nothing.
  * Cash Flow / Fund Flow are live formulas over those sheets: flow = balance sheet at period end
    minus balance sheet at period start, with profit from the matching P&L column.
  * The Balance Sheet in the source starts at 30-Apr, so April (and Q1 and every Apr-* column)
    needs the 31-Mar-26 position. That comes from the Opening BS sheet built here:
      - ConnectWave (read-only Supabase mirror of Tally), v_ledger_detail.opening, for every
        company it can serve;
      - Orange O Tec Surat from the saved Tally pull gl_data\\bs_q4_2526_<KEY>.csv, because the
        mirror times out on that book (user decision 2026-10-24);
      - opening stock from the source's own P&L (April 'Opening Stock' row);
      - Profit & Loss A/c (b/f) as the balancing figure, shown beside both source values.
    Mapped by primary group, ConnectWave's opening equals the Tally 31-Mar-26 pull to the rupee
    for every group of Noida / Enterprises / Ent Noida / Colorix (measured 2026-10-24); the few
    carry-forward differences are listed on the Opening sheet, never plugged.

COMPANY SPLIT GUARD
  Tally company files get split by year (Enterprises already is: 'F.Y.2024-26' and 'F.Y.2026-27';
  Orange O Tec may follow). A wrong file returns figures without any error. So the companies are
  never hard-coded here: FY 26-27 and FY 25-26 names are read from Files\\TALLY Company_List.xlsx
  (columns 26-27 / 25-26), matched EXACTLY against ConnectWave's v_company, and an entity whose
  26-27 name differs from its 25-26 name is treated as SPLIT: its opening is the 26-27 file's
  carried-forward opening, compared ledger by ledger with the 25-26 file's closing. Any mismatch
  in the company list stops the build.
"""
from __future__ import annotations

import argparse
import collections
import csv
import datetime as dt
import json
import re
import shutil
import sys
import tempfile
import time
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col_letter
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
PROJECT = SKILL_DIR.parent.parent.parent
FILES = PROJECT / "Files"
RESULTS = FILES / "Results"
GL_DATA = SKILL_DIR / "gl_data"
COMPANY_LIST = FILES / "TALLY Company_List.xlsx"
CONNECTWAVE_TOOLS = PROJECT.parent / "Orange One Supabase Connect" / "tools"
EXPENSE_SHEET = FILES / "EXPENSES REFERESH" / "Expense Sheet 2026-10-10 12-40-56.xlsx"

KEEP_SHEETS = ["MOM PL GL CFS", "MOM BS GL CFS", "MOM Ratio Analysis CFS",
               "MOM PL GL SFS", "MOM BS GL SFS", "MOM Ratio Analysis SFS"]

# Entity keys in the column order every MOM sheet uses. Totals are formulas over these.
ENTITY_KEYS = ["ORANGE_O_TEC_PRIVATE_LIMITED", "PRIVATE_LIMITED_NOIDA",
               "ENTERPRISES_PVT_LTD", "ENTERPRISES_NOIDA", "COLORIX"]
SHORT = {"ORANGE_O_TEC_PRIVATE_LIMITED": "O Tec Surat", "PRIVATE_LIMITED_NOIDA": "O Tec Noida",
         "ENTERPRISES_PVT_LTD": "Enterprises Surat", "ENTERPRISES_NOIDA": "Enterprises Noida",
         "COLORIX": "Colorix"}
# Position of each of the 8 columns inside a block: 5 companies, 2 entity totals, Grand Total.
BLOCK8 = ["ORANGE_O_TEC_PRIVATE_LIMITED", "PRIVATE_LIMITED_NOIDA", "TOT_OTEC",
          "ENTERPRISES_PVT_LTD", "ENTERPRISES_NOIDA", "TOT_ENT", "COLORIX", "GT"]
TOTAL_LABEL = {"TOT_OTEC": "Total of Orange O Tec", "TOT_ENT": "Total of Enterprises", "GT": "Grand Total"}

# The mirror cannot serve this book (v_ledger_detail times out, rpt_ledger_master is a stale
# snapshot) - user decision 2026-10-24: use the saved Tally pull instead.
TALLY_CSV_KEYS = {"ORANGE_O_TEC_PRIVATE_LIMITED"}

MONTHS = ["Apr-26", "May-26", "Jun-26", "Jul-26", "Aug-26", "Sep-26"]
# period label -> (BS start month or None for the 31-Mar opening, BS end month)
PERIODS = [("Apr-26", None, "Apr-26"), ("May-26", "Apr-26", "May-26"), ("Jun-26", "May-26", "Jun-26"),
           ("Q1 Total", None, "Jun-26"), ("Jul-26", "Jun-26", "Jul-26"), ("Apr-Jul Total", None, "Jul-26"),
           ("Aug-26", "Jul-26", "Aug-26"), ("Apr-Aug Total", None, "Aug-26"), ("Sep-26", "Aug-26", "Sep-26"),
           ("Q2 Total", "Jun-26", "Sep-26"), ("Apr-Sep Total", None, "Sep-26")]
PERIOD_SPAN = {"Apr-26": "31-Mar-26 to 30-Apr-26", "May-26": "30-Apr-26 to 31-May-26",
               "Jun-26": "31-May-26 to 30-Jun-26", "Q1 Total": "31-Mar-26 to 30-Jun-26",
               "Jul-26": "30-Jun-26 to 31-Jul-26", "Apr-Jul Total": "31-Mar-26 to 31-Jul-26",
               "Aug-26": "31-Jul-26 to 31-Aug-26", "Apr-Aug Total": "31-Mar-26 to 31-Aug-26",
               "Sep-26": "31-Aug-26 to 30-Sep-26", "Q2 Total": "30-Jun-26 to 30-Sep-26",
               "Apr-Sep Total": "31-Mar-26 to 30-Sep-26"}

# Balance-sheet groups in the order the MOM BS sheets use. Side and the cash/fund-flow line each
# feeds. A group in the sheet but NOT here falls to 'other current' on its side, so the cash check
# still closes - and it is printed so it gets named.
LIAB_GROUPS = ["Capital Account", "Reserves & Surplus", "Secured Loans", "Unsecured Loans", "Bank OD A/c",
               "Duties & Taxes", "Provisions", "Sundry Creditors", "Current Liabilities", "Loans (Liability)",
               "Branch / Divisions", "Suspense A/c", "Unadjusted Forex Gain/Loss", "Profit & Loss A/c (current)"]
ASSET_GROUPS = ["Fixed Assets", "Investments", "Current Assets", "Bank Accounts", "Cash-in-hand",
                "Deposits (Asset)", "Loans & Advances (Asset)", "Sundry Debtors", "Non-Current Investments",
                "Misc. Expenses (ASSET)", "Unadjusted Forex Gain/Loss (Asset)"]
STOCK_LINE = "of which: Closing Stock (Inventory)"
PL_LINE = "Profit & Loss A/c (current)"

# Tally primary-group spellings -> the MOM BS label (mirrors build_gl_detail_sidebyside's
# BS_GROUP_LABEL / CANONICAL_PRIMARY_GROUP).
GROUP_ALIAS = {"bank account": "Bank Accounts", "cash-in-hand": "Cash-in-hand",
               "non current assest": "Non-Current Investments", "non current assets": "Non-Current Investments",
               "": PL_LINE, "primary": PL_LINE}
for _g in LIAB_GROUPS + ASSET_GROUPS:
    GROUP_ALIAS.setdefault(_g.lower(), _g)
PL_PRIMARY = {"sales accounts", "purchase accounts", "direct expenses", "direct incomes",
              "indirect expenses", "indirect incomes"}

CR = 1e7  # rupees per reporting unit: crore by default; --unit lakh / abs resets it (and the labels)
UNIT_DIV = {"cr": 1e7, "lakh": 1e5, "abs": 1.0}
UNIT_SUFFIX = {"cr": "", "lakh": "_lakh", "abs": "_abs"}
# The short download: the reading pages only, as values (no GL sheets, no formulas).
SHORT_SHEETS = ["Directors Summary", "CFO Briefing", "CFO Analysis", "PDF - Statements", "PDF - MOM"]

# ---------------------------------------------------------------------------------------------
# Styles - the navy / slate palette of the MOM sheets.
# ---------------------------------------------------------------------------------------------
NAVY, BLUE, SLATE = "1F3864", "2E5395", "44546A"
F_TITLE = Font(bold=True, size=13, name="Calibri", color=NAVY)
F_HDR = Font(bold=True, size=10, name="Calibri", color="FFFFFF")
F_NOTE = Font(italic=True, size=9, name="Calibri", color="595959")
F_BOLD = Font(bold=True, size=10, name="Calibri", color=NAVY)
F_NORM = Font(size=10, name="Calibri")
F_SECTION = Font(bold=True, size=11, name="Calibri", color="FFFFFF")
F_RED = Font(bold=True, size=10, name="Calibri", color="C00000")
F_GREEN = Font(bold=True, size=10, name="Calibri", color="00703C")
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_TOT = PatternFill("solid", fgColor=BLUE)
FILL_SECTION = PatternFill("solid", fgColor=SLATE)
FILL_SUB = PatternFill("solid", fgColor="DCE4F0")
FILL_KEY = PatternFill("solid", fgColor="EDF1F8")
FILL_CHECK = PatternFill("solid", fgColor="FFF2CC")
FILL_FLAG = PatternFill("solid", fgColor="FBE5D6")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
NUM = "#,##0;[Red](#,##0);-"
NUM_CR = '#,##0.00;[Red](#,##0.00);-'
PCT = '0.0%;[Red]-0.0%;-'
WRAP = Alignment(wrap_text=True, vertical="top")


def say(msg):
    print(msg, flush=True)


# =============================================================================================
# 1. Source workbook and its layout
# =============================================================================================
def newest_source() -> Path:
    cands = [p for d in (FILES, FILES / "Orange FS") if d.exists()
             for p in d.glob("PL_BS_GL_Detail_*.xlsx") if not p.name.startswith("~$")]
    if not cands:
        sys.exit("No PL_BS_GL_Detail_*.xlsx in Files\\ or Files\\Orange FS\\")
    return max(cands, key=lambda p: p.name)


def readable_copy(path: Path, tmp: Path) -> Path:
    """The published workbook is often open in Excel / held by OneDrive; read a private copy."""
    dst = tmp / path.name
    shutil.copy2(path, dst)
    return dst


class Layout:
    """Where things are on one MOM sheet: block start columns by (period, entity) and row numbers
    by label. Found by reading the headers, never assumed, because the 8- and 12-column blocks
    (the latter carry % of Sales columns) put the same entity at different offsets."""

    def __init__(self, ws_formula, ws_values):
        self.title = ws_formula.title
        rows = list(ws_values.iter_rows(min_row=1, max_row=4, values_only=True))
        hdr_period, hdr_entity = rows[2], rows[3]
        self.col = {}
        period = None
        for i, (p, e) in enumerate(zip(hdr_period, hdr_entity), start=1):
            if p:
                period = str(p).strip()
            if not e or period is None:
                continue
            e = str(e).strip()
            if "% of Sales" in e:
                continue
            key = self._entity_key(e)
            if key:
                self.col[(period, key)] = i
        self.label_row = {}
        self.cache = {}
        for r, row in enumerate(ws_values.iter_rows(min_row=1, values_only=True), start=1):
            text = next((str(v).strip() for v in row[:4] if isinstance(v, str) and v.strip()), None)
            if text is not None:
                self.label_row.setdefault(text, r)
            self.cache[r] = row
        self.max_row = ws_values.max_row

    @staticmethod
    def _entity_key(name: str):
        n = name.upper()
        if n.startswith("TOTAL OF ORANGE"):
            return "TOT_OTEC"
        if n.startswith("TOTAL OF ENTERPRISES"):
            return "TOT_ENT"
        if n.startswith("GRAND TOTAL"):
            return "GT"
        if "COLORIX" in n:
            return "COLORIX"
        ent = "ENTERPRISES" in n
        noida = "NOIDA" in n
        if ent:
            return "ENTERPRISES_NOIDA" if noida else "ENTERPRISES_PVT_LTD"
        if "ORANGE O TEC" in n:
            return "PRIVATE_LIMITED_NOIDA" if noida else "ORANGE_O_TEC_PRIVATE_LIMITED"
        return None

    def row(self, label, startswith=True):
        if label in self.label_row:
            return self.label_row[label]
        if startswith:
            for k, r in self.label_row.items():
                if k.startswith(label):
                    return r
        return None

    def cell(self, r, c):
        row = self.cache.get(r, ())
        return row[c - 1] if 0 < c <= len(row) else None

    def value(self, label, period, entity):
        r, c = self.row(label), self.col.get((period, entity))
        if r is None or c is None:
            return None
        v = self.cell(r, c)
        return v if isinstance(v, (int, float)) else 0.0

    def ref(self, label, period, entity):
        r, c = self.row(label), self.col.get((period, entity))
        if r is None or c is None:
            return None
        return f"'{self.title}'!${col_letter(c)}${r}"


def bs_group_rows(lay: Layout):
    """{group label: row of its 'TOTAL: <group>' line} and the side each sits on."""
    groups, side = {}, None
    for r in range(1, lay.max_row + 1):
        row = lay.cache.get(r) or (None,) * 4
        a = row[0]
        if isinstance(a, str):
            if a.startswith("LIABILITIES"):
                side = "L"
            elif a.startswith("ASSETS"):
                side = "A"
            elif a.startswith("RECONCILIATION") or a.startswith("ELIMINATED"):
                side = None
        d = row[3]
        if side and isinstance(d, str) and d.startswith("TOTAL: "):
            groups[d[len("TOTAL: "):].strip()] = (r, side)
    return groups


# =============================================================================================
# 2. Companies (split guard) and the 31-Mar-26 opening position
# =============================================================================================
def entity_of(name: str):
    return Layout._entity_key(name)


def read_company_list(tmp: Path):
    wb = openpyxl.load_workbook(readable_copy(COMPANY_LIST, tmp), data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = next(i for i, r in enumerate(rows) if r and r[0] == "Company Name")
    hdr = [str(h).strip() if h else "" for h in rows[hdr_i]]
    c26, c25 = hdr.index("26-27"), hdr.index("25-26")
    by_year = {"2627": {}, "2526": {}}
    for r in rows[hdr_i + 1:]:
        if not r or not r[0] or (r[3] and "NOT TO USE" in str(r[3]).upper()):
            continue
        name = str(r[0]).strip()
        key = entity_of(name)
        for col, fy in ((c26, "2627"), (c25, "2526")):
            if r[col] in (1, "1") and key:
                if key in by_year[fy]:
                    sys.exit(f"Company list marks two {fy} files for {key}: {by_year[fy][key]!r} and {name!r}")
                by_year[fy][key] = name
    for fy in by_year:
        missing = set(ENTITY_KEYS) - set(by_year[fy])
        if missing:
            sys.exit(f"Company list has no {fy} file for: {sorted(missing)}")
    return by_year


def connectwave():
    sys.path.insert(0, str(CONNECTWAVE_TOOLS))
    import supabase_client as sb  # read-only by construction (see its docstring)
    return sb


def resolve_companies(tmp: Path, sb):
    by_year = read_company_list(tmp)
    cw = {c["company_name"]: c["tenant_id"] for c in sb.list_companies()} if hasattr(sb, "list_companies") \
        else {c["company_name"]: c["tenant_id"] for c in sb.select("v_company", order="tenant_id")}
    plan = {}
    for key in ENTITY_KEYS:
        n27, n25 = by_year["2627"][key], by_year["2526"][key]
        split = n27 != n25
        for n in ({n27, n25} if split else {n27}):
            if n not in cw:
                sys.exit(f"STOP - company {n!r} (from the company list) is not in ConnectWave v_company. "
                         f"ConnectWave holds: {sorted(cw)}")
        plan[key] = {"fy2627": n27, "fy2526": n25, "split": split,
                     "tenant": cw[n27], "tenant_prior": cw[n25] if split else None}
    return plan


def canon_group(chain_or_primary, ledger=""):
    """Tally primary group (or a ConnectWave group_chain, nearest first) -> MOM BS label.
    Returns None for P&L heads, 'STOCK' for Tally's Stock-in-Hand pseudo-rows."""
    items = chain_or_primary if isinstance(chain_or_primary, list) else [chain_or_primary]
    for g in items:
        g = (g or "").strip()
        lg = g.lower()
        if lg == "stock-in-hand":
            return "STOCK"
        if lg in PL_PRIMARY:
            return None
        if lg in GROUP_ALIAS and (lg or ledger.strip() == "Profit & Loss A/c"):
            return GROUP_ALIAS[lg]
    return "?" + (items[-1] if items else "")


def opening_from_connectwave(sb, tenant):
    """{ledger: (group label, sub_group, balance in Tally sign: Cr +, Dr -)} from the mirror's
    FY 26-27 opening. ConnectWave reports Dr as +, hence the negation."""
    out = {}
    for r in sb.select("v_ledger_detail", filters={"tenant_id": tenant}, order="guid"):
        if not r.get("opening"):
            continue
        g = canon_group(r.get("group_chain") or [r.get("sub_group"), r.get("grouping")], r["ledger"])
        if g is None:
            continue
        out[r["ledger"]] = (g, r.get("sub_group") or "", -float(r["opening"]))
    return out


def closing_from_connectwave(sb, tenant):
    out = {}
    for r in sb.select("v_ledger_detail", filters={"tenant_id": tenant}, order="guid"):
        if not r.get("closing"):
            continue
        g = canon_group(r.get("group_chain") or [r.get("sub_group"), r.get("grouping")], r["ledger"])
        if g is None:
            continue
        sign = 1 if (r.get("closing_dr_cr") or "Dr") == "Dr" else -1
        out[r["ledger"]] = (g, r.get("sub_group") or "", -sign * abs(float(r["closing"])))
    return out


def opening_from_tally_csv(key):
    path = GL_DATA / f"bs_q4_2526_{key}.csv"
    if not path.exists():
        sys.exit(f"Missing {path} - the saved Tally 31-Mar-26 pull for {key}")
    out, seen = {}, set()
    with open(path, newline="", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            name = r["ledger_name"]
            if name in seen:
                continue
            seen.add(name)
            if r["group_name"] == "Stock-in-Hand":
                continue  # the TB's 'Opening Stock' is the FY 25-26 OPENING, not 31-Mar-26
            g = canon_group(r["primary_group"], name)
            if g is None:
                continue
            out[name] = (g, r["group_name"], float(r["closing_balance"]))
    return out, dt.datetime.fromtimestamp(path.stat().st_mtime)


def group_totals(ledgers):
    t = collections.defaultdict(float)
    for g, _sg, bal in ledgers.values():
        t[g] += bal
    return t


def eliminated_ledgers(cfs_lay: Layout):
    """Ledger names the CFS sheet eliminates (its memo block), by entity column that carries them."""
    start = cfs_lay.row("ELIMINATED ON CONSOLIDATION")
    names = collections.defaultdict(set)
    if not start:
        return names
    for r in range(start + 1, cfs_lay.max_row + 1):
        d = cfs_lay.cell(r, 4)
        if not isinstance(d, str) or d.startswith(("Subtotal", "TOTAL")):
            continue
        for key in ENTITY_KEYS:
            for m in MONTHS:
                c = cfs_lay.col.get((m, key))
                x = cfs_lay.cell(r, c) if c else None
                if isinstance(x, (int, float)) and abs(x) > 0.005:
                    names[key].add(d.strip())
                    break
    return names


def is_intercompany(group, sub_group, ledger, elim_names):
    return (group == "Branch / Divisions" or ledger in elim_names
            or "RELATED PARTY" in (sub_group or "").upper())


# =============================================================================================
# 3. Writing helpers
# =============================================================================================
def hdr_cell(ws, r, c, v, fill=FILL_HDR, width=None):
    cell = ws.cell(r, c, v)
    cell.font, cell.fill, cell.border = F_HDR, fill, BOX
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return cell


def num_cell(ws, r, c, v, fmt=NUM, font=F_NORM, fill=None):
    cell = ws.cell(r, c, v)
    cell.number_format, cell.font = fmt, font
    if fill:
        cell.fill = fill
    return cell


def label(ws, r, text, font=F_NORM, fill=None, col=1, indent=0):
    cell = ws.cell(r, col, text)
    cell.font = font
    if indent:
        cell.alignment = Alignment(indent=indent)
    if fill:
        cell.fill = fill
    return cell


def band(ws, r, text, last_col, font=F_SECTION, fill=FILL_SECTION):
    for c in range(1, last_col + 1):
        ws.cell(r, c).fill = fill
    label(ws, r, text, font)


def block_headers(ws, first_col, entity_names, periods, r_period=3, r_entity=4, r_span=5):
    """Period / entity header rows in the MOM layout: 8 columns per period block."""
    pos = {}
    c = first_col
    for p in periods:
        hdr_cell(ws, r_period, c, p)
        ws.merge_cells(start_row=r_period, start_column=c, end_row=r_period, end_column=c + 7)
        span = ws.cell(r_span, c, PERIOD_SPAN.get(p, p))
        span.font = F_NOTE
        ws.merge_cells(start_row=r_span, start_column=c, end_row=r_span, end_column=c + 7)
        span.alignment = Alignment(horizontal="center")
        for i, key in enumerate(BLOCK8):
            hdr_cell(ws, r_entity, c + i, entity_names[key], FILL_TOT if key in TOTAL_LABEL else FILL_HDR)
            pos[(p, key)] = c + i
            ws.column_dimensions[col_letter(c + i)].width = 15
        c += 8
    return pos


def total_formula(pos, period, key, r):
    L = lambda k: f"{col_letter(pos[(period, k)])}{r}"
    if key == "TOT_OTEC":
        return f"={L('ORANGE_O_TEC_PRIVATE_LIMITED')}+{L('PRIVATE_LIMITED_NOIDA')}"
    if key == "TOT_ENT":
        return f"={L('ENTERPRISES_PVT_LTD')}+{L('ENTERPRISES_NOIDA')}"
    return f"={L('TOT_OTEC')}+{L('TOT_ENT')}+{L('COLORIX')}"


# =============================================================================================
# 4. Opening Balance Sheet 31-Mar-26
# =============================================================================================
def build_opening(wb, sfs_lay, cfs_lay, entity_names, opening, sources, compare, stock_refs):
    """Two blocks (SFS, CFS) in the MOM BS column layout: rows = groups, E..L = 8 columns.
    Returns {('SFS'|'CFS', group|'STOCK'|'GAP'|'TL'|'TA'): row}."""
    ws = wb.create_sheet("Opening BS 31-Mar-26")
    ws.sheet_properties.tabColor = SLATE
    ws["A1"] = "OPENING BALANCE SHEET as at 31-Mar-2026 (= 1-Apr-2026 opening of FY 2026-27) - group level"
    ws["A1"].font = F_TITLE
    ws["A2"] = ("Rs. Liabilities and assets both shown positive, as on the MOM BS sheets. Source per company in row 5. "
                "Closing stock is the source workbook's own April 'Opening Stock'. Profit & Loss A/c is the balancing "
                "figure (FY 25-26 profit is not a ledger at 31-Mar); the two source values are shown beside it.")
    ws["A2"].font = F_NOTE
    ws.column_dimensions["A"].width = 46
    for c in "BCD":
        ws.column_dimensions[c].width = 2
    rows = {}
    E = 5  # first value column, as on the MOM BS sheets
    for i, key in enumerate(BLOCK8):
        hdr_cell(ws, 4, E + i, entity_names[key], FILL_TOT if key in TOTAL_LABEL else FILL_HDR)
        ws.column_dimensions[col_letter(E + i)].width = 17
        if key in sources:
            cell = ws.cell(5, E + i, sources[key])
            cell.font, cell.alignment = F_NOTE, Alignment(wrap_text=True, horizontal="center")
    ws.row_dimensions[5].height = 42
    r = 7
    for basis, lay in (("SFS", sfs_lay), ("CFS", cfs_lay)):
        sheet_groups = bs_group_rows(lay)
        band(ws, r, f"{basis} - " + ("STANDALONE (every company as in Tally)" if basis == "SFS" else
                                     "CONSOLIDATED (inter-company and branch balances eliminated)"), E + 7)
        r += 1
        for side, groups in (("L", LIAB_GROUPS), ("A", ASSET_GROUPS)):
            label(ws, r, "LIABILITIES (Sources of Funds)" if side == "L" else "ASSETS (Application of Funds)", F_BOLD)
            r += 1
            first = r
            present = [g for g in groups if g in sheet_groups or any(opening[basis][k].get(g) for k in ENTITY_KEYS)]
            for g in present:
                rows[(basis, g)] = r
                label(ws, r, g, indent=1)
                for i, key in enumerate(BLOCK8):
                    c = E + i
                    if key in TOTAL_LABEL:
                        num_cell(ws, r, c, total_formula({("x", k): E + j for j, k in enumerate(BLOCK8)}, "x", key, r),
                                 font=F_BOLD)
                    elif g == PL_LINE:
                        continue  # balancing figure, written below
                    elif g == "Current Assets":
                        # non-stock current assets from the source + stock linked from the P&L sheet
                        v = opening[basis][key].get(g, 0.0)
                        num_cell(ws, r, c, f"={round(v, 2)}+{stock_refs[key]}")
                    else:
                        num_cell(ws, r, c, round(opening[basis][key].get(g, 0.0), 2))
                if g == "Current Assets":
                    r += 1
                    rows[(basis, "STOCK")] = r
                    label(ws, r, STOCK_LINE, F_NOTE, indent=2)
                    for i, key in enumerate(BLOCK8):
                        c = E + i
                        if key in TOTAL_LABEL:
                            num_cell(ws, r, c, total_formula({("x", k): E + j for j, k in enumerate(BLOCK8)}, "x", key, r),
                                     font=F_NOTE)
                        else:
                            num_cell(ws, r, c, f"={stock_refs[key]}", font=F_NOTE)
                r += 1
            tot_r = r
            label(ws, r, "TOTAL LIABILITIES" if side == "L" else "TOTAL ASSETS", F_BOLD, FILL_SUB)
            for i in range(8):
                c = E + i
                body = [f"{col_letter(c)}{rows[(basis, g)]}" for g in present]
                num_cell(ws, r, c, "=" + "+".join(body) if body else 0, font=F_BOLD, fill=FILL_SUB)
            rows[(basis, "TL" if side == "L" else "TA")] = tot_r
            r += 2
        # P&L balancing figure: TOTAL ASSETS - (TOTAL LIABILITIES excluding the P&L line itself),
        # per company. For CFS it is the SFS figure (b/f profit is not eliminated); the CFS gap row
        # then shows the unmatched inter-company balance.
        pl_r = rows[(basis, PL_LINE)]
        for i, key in enumerate(BLOCK8):
            if key in TOTAL_LABEL:
                continue
            c = col_letter(E + i)
            if basis == "SFS":
                others = "+".join(f"{c}{rows[(basis, g)]}" for g in LIAB_GROUPS
                                  if (basis, g) in rows and g != PL_LINE)
                num_cell(ws, pl_r, E + i, f"={c}{rows[(basis, 'TA')]}-({others})", fill=FILL_KEY)
            else:
                num_cell(ws, pl_r, E + i, f"={c}{rows[('SFS', PL_LINE)]}", fill=FILL_KEY)
        rows[(basis, "GAP")] = r
        label(ws, r, "RECONCILIATION GAP (Liabilities - Assets)" + (" - must be 0" if basis == "SFS" else
                                                                    " = unmatched inter-company at 31-Mar"), F_BOLD,
              FILL_CHECK)
        for i in range(8):
            c = col_letter(E + i)
            num_cell(ws, r, E + i, f"={c}{rows[(basis, 'TL')]}-{c}{rows[(basis, 'TA')]}", font=F_BOLD, fill=FILL_CHECK)
        r += 3

    # --- Source cross-check: ConnectWave vs the saved Tally pull, by group and by ledger --------------
    band(ws, r, "SOURCE CROSS-CHECK - ConnectWave opening 1-Apr-26 vs saved Tally pull 31-Mar-26 (Rs, Tally sign: Cr +)", E + 7)
    r += 1
    for c, h in enumerate(["Company", "", "", "Group", "ConnectWave", "Tally pull", "Difference", "Comment"], start=1):
        if h:
            hdr_cell(ws, r, c, h)
    r += 1
    for key, rowsets in compare.items():
        for g, cwv, tv, comment in rowsets:
            label(ws, r, SHORT[key])
            ws.cell(r, 4, g).font = F_NORM
            num_cell(ws, r, 5, round(cwv, 2))
            num_cell(ws, r, 6, round(tv, 2))
            num_cell(ws, r, 7, f"=E{r}-F{r}", font=F_RED if abs(cwv - tv) >= 1 else F_NORM)
            ws.cell(r, 8, comment).font = F_NOTE
            r += 1
    return ws, rows


# =============================================================================================
# 5. Cash Flow and Fund Flow
# =============================================================================================
def flow_sources(basis, lay_bs, lay_pl, open_rows):
    """Closure that returns, for (period, entity, group): an Excel expression for the BS balance at
    the period END and START, and P&L references."""
    groups = bs_group_rows(lay_bs)
    stock_row = lay_bs.row("Subtotal: Closing Stock (Inventory)")
    gap_row = lay_bs.row("RECONCILIATION GAP")
    op_sheet = "'Opening BS 31-Mar-26'"
    op_col = {k: 5 + i for i, k in enumerate(BLOCK8)}

    def bs(group, month, key):
        if month is None:  # 31-Mar opening
            if group == "STOCK":
                r = open_rows.get((basis, "STOCK"))
            elif group == "GAP":
                r = open_rows.get((basis, "GAP"))
            else:
                r = open_rows.get((basis, group))
            return f"{op_sheet}!${col_letter(op_col[key])}${r}" if r else "0"
        c = lay_bs.col.get((month, key))
        if group == "STOCK":
            r = stock_row
        elif group == "GAP":
            r = gap_row
        else:
            r = groups.get(group, (None,))[0]
        return f"'{lay_bs.title}'!${col_letter(c)}${r}" if (r and c) else "0"

    def delta(group, period, key):
        _, start, end = next(p for p in PERIODS if p[0] == period)
        return f"({bs(group, end, key)}-{bs(group, start, key)})"

    def pl(label_, period, key):
        ref = lay_pl.ref(label_, period, key)
        return ref or "0"

    return groups, bs, delta, pl


def classify_extra_groups(groups):
    known = set(LIAB_GROUPS) | set(ASSET_GROUPS)
    extra_l = [g for g, (_r, s) in groups.items() if s == "L" and g not in known]
    extra_a = [g for g, (_r, s) in groups.items() if s == "A" and g not in known]
    for g in extra_l + extra_a:
        say(f"   ! BS group {g!r} is not mapped to a named line - carried in 'other current' so the cash check still closes")
    return extra_l, extra_a


def build_cash_flow(wb, basis, lay_bs, lay_pl, open_rows, entity_names):
    ws = wb.create_sheet(f"MOM Cash Flow {basis}")
    ws.sheet_properties.tabColor = "00703C"
    groups, bs, delta, pl = flow_sources(basis, lay_bs, lay_pl, open_rows)
    extra_l, extra_a = classify_extra_groups(groups)
    title = ("MOM - CONSOLIDATED Cash Flow Statement (indirect method; inter-company and branch eliminated)" if basis == "CFS"
             else "MOM - STANDALONE Cash Flow Statement (indirect method; every company as in Tally)")
    ws["A1"], ws["A1"].font = title, F_TITLE
    ws["A2"] = ("Rs. Every figure is a live formula: movement between the opening and closing Balance Sheet of the period "
                f"('{lay_bs.title}' and 'Opening BS 31-Mar-26'), profit from '{lay_pl.title}'. Depreciation and tax are NOT "
                "booked for Apr-Sep, so cash flow is before them.")
    ws["A2"].font = F_NOTE
    ws.column_dimensions["A"].width = 58
    pos = block_headers(ws, 2, entity_names, [p[0] for p in PERIODS])
    ws.freeze_panes = "B6"

    def neg(expr):
        return f"-{expr}"

    lines = []  # (key, label, kind, builder) ; builder(period, entity) -> formula body (no '=')
    A = lambda *k: None
    lines += [
        ("hdrA", "A. CASH FLOW FROM OPERATING ACTIVITIES", "band", None),
        ("pbt", "Profit before tax", "line", lambda p, k: pl("PBT (PROFIT BEFORE TAX)", p, k)),
        ("tax", "Less: tax charge (PAT - PBT; not booked Apr-Sep - tax paid sits in advance tax below)", "line",
         lambda p, k: pl("PAT (NET PROFIT)", p, k) + "-" + pl("PBT (PROFIT BEFORE TAX)", p, k)),
        ("dep", "Add: Depreciation & amortisation (not booked Apr-Sep)", "line", lambda p, k: pl("LESS: DEPRECIATION", p, k)),
        ("fin", "Add: Finance cost (shown under financing)", "line", lambda p, k: pl("TOTAL FINANCE COST", p, k)),
        ("opwc", "Operating profit before working capital changes", "sub", ["pbt", "tax", "dep", "fin"]),
        ("inv", "(Increase)/decrease in inventory", "line", lambda p, k: neg(delta("STOCK", p, k))),
        ("deb", "(Increase)/decrease in sundry debtors", "line", lambda p, k: neg(delta("Sundry Debtors", p, k))),
        ("lna", "(Increase)/decrease in loans & advances", "line", lambda p, k: neg(delta("Loans & Advances (Asset)", p, k))),
        ("dep_a", "(Increase)/decrease in deposits", "line", lambda p, k: neg(delta("Deposits (Asset)", p, k))),
        ("oca", "(Increase)/decrease in other current assets (incl. advance tax / TDS)", "line",
         lambda p, k: "-(" + "+".join([f"({delta('Current Assets', p, k)}-{delta('STOCK', p, k)})"]
                                      + [delta(g, p, k) for g in ["Misc. Expenses (ASSET)", "Unadjusted Forex Gain/Loss (Asset)"] + extra_a]) + ")"),
        ("cred", "Increase/(decrease) in sundry creditors", "line", lambda p, k: delta("Sundry Creditors", p, k)),
        ("ocl", "Increase/(decrease) in other current liabilities", "line",
         lambda p, k: "+".join(delta(g, p, k) for g in ["Current Liabilities", "Suspense A/c", "Unadjusted Forex Gain/Loss"] + extra_l)),
        ("dt", "Increase/(decrease) in duties & taxes", "line", lambda p, k: delta("Duties & Taxes", p, k)),
        ("prov", "Increase/(decrease) in provisions", "line", lambda p, k: delta("Provisions", p, k)),
        ("netA", "NET CASH FROM OPERATING ACTIVITIES (A)", "total", ["opwc", "inv", "deb", "lna", "dep_a", "oca", "cred", "ocl", "dt", "prov"]),
        ("hdrB", "B. CASH FLOW FROM INVESTING ACTIVITIES", "band", None),
        ("fa", "Purchase of fixed assets (net of disposals; no depreciation booked)", "line",
         lambda p, k: neg(delta("Fixed Assets", p, k)) + "-" + pl("LESS: DEPRECIATION", p, k)),
        ("invst", "(Purchase)/sale of investments", "line",
         lambda p, k: "-(" + delta("Investments", p, k) + "+" + delta("Non-Current Investments", p, k) + ")"),
        ("netB", "NET CASH FROM INVESTING ACTIVITIES (B)", "total", ["fa", "invst"]),
        ("hdrC", "C. CASH FLOW FROM FINANCING ACTIVITIES", "band", None),
        ("cap", "Capital introduced / (withdrawn)", "line", lambda p, k: delta("Capital Account", p, k)),
        ("res", "Reserves & surplus - movement", "line", lambda p, k: delta("Reserves & Surplus", p, k)),
        ("sec", "Term / secured loans raised / (repaid)", "line", lambda p, k: delta("Secured Loans", p, k)),
        ("uns", "Unsecured loans raised / (repaid)", "line", lambda p, k: delta("Unsecured Loans", p, k)),
        ("oth", "Other loans raised / (repaid)", "line", lambda p, k: delta("Loans (Liability)", p, k)),
        ("od", "Bank OD / cash credit drawn / (repaid)", "line", lambda p, k: delta("Bank OD A/c", p, k)),
        ("finp", "Finance cost paid", "line", lambda p, k: neg(pl("TOTAL FINANCE COST", p, k))),
        ("netC", "NET CASH FROM FINANCING ACTIVITIES (C)", "total", ["cap", "res", "sec", "uns", "oth", "od", "finp"]),
        ("D", "D. P&L A/c movement not matching the period's PAT (see note 3)", "line",
         lambda p, k: delta(PL_LINE, p, k) + "-" + pl("PAT (NET PROFIT)", p, k)),
        ("E", "E. Movement in branch / unmatched inter-company balances", "line",
         lambda p, k: delta("Branch / Divisions", p, k) + "-" + delta("GAP", p, k)),
        ("net", "NET INCREASE / (DECREASE) IN CASH & BANK (A+B+C+D+E)", "total", ["netA", "netB", "netC", "D", "E"]),
        ("ocash", "Opening cash & bank", "line",
         lambda p, k: (lambda s: f"{bs('Bank Accounts', s, k)}+{bs('Cash-in-hand', s, k)}")(next(x for x in PERIODS if x[0] == p)[1])),
        ("ccash", "Closing cash & bank", "total", ["ocash", "net"]),
        ("bcash", "Check: closing cash & bank per Balance Sheet", "line",
         lambda p, k: (lambda e: f"{bs('Bank Accounts', e, k)}+{bs('Cash-in-hand', e, k)}")(next(x for x in PERIODS if x[0] == p)[2])),
        ("chk", "Check: difference (must be nil)", "check", ("ccash", "bcash")),
    ]
    rows = write_lines(ws, lines, pos, start_row=7)
    notes = rows["_next"] + 1
    for i, t in enumerate([
        "Notes",
        "1. Bank OD / cash credit is under financing; cash & bank is the Bank Accounts and Cash-in-hand groups only.",
        "2. Depreciation and tax are not booked for Apr-Sep 2026, so operating cash is before tax and fixed-asset additions are net of nothing.",
        "3. Line D: the Balance Sheet's P&L A/c moved by a different amount than the period's PAT. On the SFS it is nil for Q1, Q2 and "
        "Apr-Sep for every company - which also proves the 31-Mar opening ties to the books. Inside Q1 the source workbook's Orange O Tec "
        "Surat balance sheet carries about Rs 1.70 cr more profit at 30-Apr than April's P&L, reversing in May-June: a month-allocation "
        "difference in the source, not a cash item. On the CFS entity columns it also carries the inter-company profit that "
        "consolidation moves between entities (nil on the Grand Total).",
        "4. Line E: branch balances (SFS) and inter-company balances that do not mirror each other (CFS entity columns). "
        "On the CFS Grand Total it is the movement in the unreconciled inter-company gap.",
        "5. CFS entity columns show external trade only (inter-company removed from seller and buyer) - use SFS for entity performance.",
    ]):
        label(ws, notes + i, t, F_BOLD if i == 0 else F_NOTE)
    return ws, rows, pos


def write_lines(ws, lines, pos, start_row):
    rows = {}
    r = start_row
    periods = [p[0] for p in PERIODS]
    for key, text, kind, spec in lines:
        if kind == "band":
            band(ws, r, text, max(pos.values()))
            r += 1
            continue
        rows[key] = r
        font = F_BOLD if kind in ("sub", "total", "check") else F_NORM
        fill = FILL_SUB if kind in ("sub", "total") else (FILL_CHECK if kind == "check" else None)
        label(ws, r, text, font, fill, indent=0 if kind != "line" else 1)
        for p in periods:
            for ek in BLOCK8:
                c = pos[(p, ek)]
                if ek in TOTAL_LABEL:
                    f = total_formula(pos, p, ek, r)
                elif kind == "line":
                    f = "=" + spec(p, ek)
                elif kind in ("sub", "total"):
                    f = "=" + "+".join(f"{col_letter(c)}{rows[x]}" for x in spec)
                elif kind == "check":
                    a, b = spec
                    f = f"=ROUND({col_letter(c)}{rows[a]}-{col_letter(c)}{rows[b]},2)"
                num_cell(ws, r, c, f, font=font, fill=fill)
        r += 1
    rows["_next"] = r
    return rows


def build_fund_flow(wb, basis, lay_bs, lay_pl, open_rows, entity_names):
    ws = wb.create_sheet(f"MOM Fund Flow {basis}")
    ws.sheet_properties.tabColor = "7030A0"
    groups, bs, delta, pl = flow_sources(basis, lay_bs, lay_pl, open_rows)
    extra_l, extra_a = classify_extra_groups(groups)
    ws["A1"] = ("MOM - CONSOLIDATED Fund Flow Statement (CMA style)" if basis == "CFS"
                else "MOM - STANDALONE Fund Flow Statement (CMA style)")
    ws["A1"].font = F_TITLE
    ws["A2"] = ("Rs. Long-term funds (equity, term / unsecured / other loans, fixed assets, investments) are sources and "
                "applications; Bank OD / cash credit is a current liability, as in the bank CMA format. Live formulas.")
    ws["A2"].font = F_NOTE
    ws.column_dimensions["A"].width = 58
    pos = block_headers(ws, 2, entity_names, [p[0] for p in PERIODS])
    ws.freeze_panes = "B6"
    pos_ = lambda x: f"MAX({x},0)"
    negpos = lambda x: f"MAX(-{x},0)"
    lines = [
        ("hdrS", "SOURCES OF FUNDS", "band", None),
        ("fop", "Funds from operations (PAT + depreciation)", "line",
         lambda p, k: pl("PAT (NET PROFIT)", p, k) + "+" + pl("LESS: DEPRECIATION", p, k)),
        ("pya", "P&L A/c movement not matching PAT (see Cash Flow note 3)", "line",
         lambda p, k: delta(PL_LINE, p, k) + "-" + pl("PAT (NET PROFIT)", p, k)),
        ("capi", "Capital introduced / reserves added", "line",
         lambda p, k: pos_(delta("Capital Account", p, k) + "+" + delta("Reserves & Surplus", p, k))),
        ("seci", "Term / secured loans raised", "line", lambda p, k: pos_(delta("Secured Loans", p, k))),
        ("unsi", "Unsecured loans raised", "line", lambda p, k: pos_(delta("Unsecured Loans", p, k))),
        ("othi", "Other loans raised", "line", lambda p, k: pos_(delta("Loans (Liability)", p, k))),
        ("invs", "Sale of investments", "line",
         lambda p, k: negpos(delta("Investments", p, k) + "+" + delta("Non-Current Investments", p, k))),
        ("fas", "Sale of fixed assets", "line",
         lambda p, k: negpos(delta("Fixed Assets", p, k) + "+" + pl("LESS: DEPRECIATION", p, k))),
        ("ic", "Movement in branch / unmatched inter-company balances", "line",
         lambda p, k: delta("Branch / Divisions", p, k) + "-" + delta("GAP", p, k)),
        ("TS", "TOTAL SOURCES", "total", ["fop", "pya", "capi", "seci", "unsi", "othi", "invs", "fas", "ic"]),
        ("hdrA", "APPLICATION OF FUNDS", "band", None),
        ("faa", "Purchase of fixed assets", "line",
         lambda p, k: pos_(delta("Fixed Assets", p, k) + "+" + pl("LESS: DEPRECIATION", p, k))),
        ("inva", "Purchase of investments", "line",
         lambda p, k: pos_(delta("Investments", p, k) + "+" + delta("Non-Current Investments", p, k))),
        ("capw", "Capital withdrawn / reserves reduced", "line",
         lambda p, k: negpos(delta("Capital Account", p, k) + "+" + delta("Reserves & Surplus", p, k))),
        ("secr", "Term / secured loans repaid", "line", lambda p, k: negpos(delta("Secured Loans", p, k))),
        ("unsr", "Unsecured loans repaid", "line", lambda p, k: negpos(delta("Unsecured Loans", p, k))),
        ("othr", "Other loans repaid", "line", lambda p, k: negpos(delta("Loans (Liability)", p, k))),
        ("TA", "TOTAL APPLICATIONS", "total", ["faa", "inva", "capw", "secr", "unsr", "othr"]),
        ("NWC", "INCREASE / (DECREASE) IN NET WORKING CAPITAL (Sources less Applications)", "diff", ("TS", "TA")),
        ("hdrW", "SCHEDULE OF CHANGES IN WORKING CAPITAL", "band", None),
        ("w_inv", "Current assets - Inventory", "line", lambda p, k: delta("STOCK", p, k)),
        ("w_deb", "Current assets - Sundry debtors", "line", lambda p, k: delta("Sundry Debtors", p, k)),
        ("w_lna", "Current assets - Loans & advances", "line", lambda p, k: delta("Loans & Advances (Asset)", p, k)),
        ("w_dep", "Current assets - Deposits", "line", lambda p, k: delta("Deposits (Asset)", p, k)),
        ("w_oca", "Current assets - Other (incl. advance tax / TDS)", "line",
         lambda p, k: "+".join([f"({delta('Current Assets', p, k)}-{delta('STOCK', p, k)})"]
                               + [delta(g, p, k) for g in ["Misc. Expenses (ASSET)", "Unadjusted Forex Gain/Loss (Asset)"] + extra_a])),
        ("w_cash", "Current assets - Cash & bank", "line",
         lambda p, k: delta("Bank Accounts", p, k) + "+" + delta("Cash-in-hand", p, k)),
        ("CA", "Increase / (decrease) in current assets", "total", ["w_inv", "w_deb", "w_lna", "w_dep", "w_oca", "w_cash"]),
        ("w_cr", "Current liabilities - Sundry creditors", "line", lambda p, k: delta("Sundry Creditors", p, k)),
        ("w_ocl", "Current liabilities - Other", "line",
         lambda p, k: "+".join(delta(g, p, k) for g in ["Current Liabilities", "Suspense A/c", "Unadjusted Forex Gain/Loss"] + extra_l)),
        ("w_dt", "Current liabilities - Duties & taxes", "line", lambda p, k: delta("Duties & Taxes", p, k)),
        ("w_pr", "Current liabilities - Provisions", "line", lambda p, k: delta("Provisions", p, k)),
        ("w_od", "Current liabilities - Bank OD / cash credit", "line", lambda p, k: delta("Bank OD A/c", p, k)),
        ("CL", "Increase / (decrease) in current liabilities", "total", ["w_cr", "w_ocl", "w_dt", "w_pr", "w_od"]),
        ("WC", "NET INCREASE / (DECREASE) IN WORKING CAPITAL", "diff", ("CA", "CL")),
        ("chk", "Check: difference to Sources less Applications (must be nil)", "check", ("WC", "NWC")),
    ]
    # 'diff' rows: a - b
    expanded = []
    for key, text, kind, spec in lines:
        if kind == "diff":
            a, b = spec
            expanded.append((key, text, "difftotal", (a, b)))
        else:
            expanded.append((key, text, kind, spec))
    rows = write_lines_ff(ws, expanded, pos)
    label(ws, rows["_next"] + 1, "Sources and applications are split per company column (a loan raised in one company is not netted against "
                                 "a repayment in another); the total columns add the companies. Months therefore need not add to the "
                                 "quarter on the gross lines - the NET lines do.", F_NOTE)
    return ws, rows, pos


def write_lines_ff(ws, lines, pos):
    rows, r = {}, 7
    for key, text, kind, spec in lines:
        if kind == "band":
            band(ws, r, text, max(pos.values()))
            r += 1
            continue
        rows[key] = r
        font = F_BOLD if kind != "line" else F_NORM
        fill = FILL_SUB if kind in ("total", "difftotal") else (FILL_CHECK if kind == "check" else None)
        label(ws, r, text, font, fill, indent=1 if kind == "line" else 0)
        for p, _s, _e in PERIODS:
            for ek in BLOCK8:
                c = pos[(p, ek)]
                L = col_letter(c)
                if ek in TOTAL_LABEL:
                    f = total_formula(pos, p, ek, r)
                elif kind == "line":
                    f = "=" + spec(p, ek)
                elif kind == "total":
                    f = "=" + "+".join(f"{L}{rows[x]}" for x in spec)
                elif kind == "difftotal":
                    f = f"={L}{rows[spec[0]]}-{L}{rows[spec[1]]}"
                else:
                    f = f"=ROUND({L}{rows[spec[0]]}-{L}{rows[spec[1]]},2)"
                num_cell(ws, r, c, f, font=font, fill=fill)
        r += 1
    rows["_next"] = r
    return rows


# =============================================================================================
# 6. CFO Analysis and the print sheets
# =============================================================================================
def ref_or_zero(x):
    return x if x else "0"


def build_cfo_analysis(wb, L, cf_rows, cf_pos, ff_rows, ff_pos, facts, tieout, movers):
    """One sheet: the one-page summary, then KPIs, entity view, Q2 vs Q1 bridge, cash story,
    expense tie-out and the Checks block. Figures are formulas (Rs crore) on the sheets above."""
    ws = wb.create_sheet("CFO Analysis", 1)
    ws.sheet_properties.tabColor = "C00000"
    pl_c, bs_c, pl_s = L["PL_CFS"], L["BS_CFS"], L["PL_SFS"]
    cf_c, cf_s = "'MOM Cash Flow CFS'", "'MOM Cash Flow SFS'"
    widths = {"A": 46, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 13, "I": 13, "J": 13, "K": 13, "L": 13}
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    r = 1
    ws.cell(r, 1, "CFO ANALYSIS - Orange O Tec Group, FY 2026-27 April to September").font = F_TITLE
    r += 1
    ws.cell(r, 1, f"Rs crore unless stated. Source: {facts['source_name']} (live Tally pull {facts['source_stamp']}). "
                  "Every table is a formula on the sheets in this file. Depreciation and tax are NOT booked for Apr-Sep.").font = F_NOTE
    r += 2
    marks = {}

    def briefing_block(sections, title):
        nonlocal r
        band(ws, r, title, 12)
        r += 1
        for head, items in sections:
            label(ws, r, head, F_BOLD)
            r += 1
            for it in items:
                cell = ws.cell(r, 1, "•  " + it)
                cell.font, cell.alignment = F_NORM, WRAP
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
                ws.row_dimensions[r].height = 15 * (1 + len(it) // 190)
                r += 1
            r += 1

    def page_break():
        ws.row_breaks.append(Break(id=r - 1))

    # The summary, red flags, board Q&A and caveats live on their own 'CFO Briefing' tab, which is
    # kept out of the PDF (user, 2026-10-24) - this sheet is the figures only.

    def cr(expr):
        return f"=({expr})/{int(CR)}"

    def gt(lay, label_, period):
        return ref_or_zero(lay.ref(label_, period, "GT"))

    # ---- 1. Headline numbers ----------------------------------------------------------------------
    band(ws, r, "1. HEADLINE NUMBERS - GROUP (CFS)", 12)
    r += 1
    hdr = ["Rs crore", "Q1 (Apr-Jun)", "Q2 (Jul-Sep)", "Q2 vs Q1", "Apr-Sep", "Sep-26"]
    for i, h in enumerate(hdr, start=1):
        hdr_cell(ws, r, i, h)
    r += 1
    marks["head_first"] = r
    heads = [("Turnover", "TOTAL TURNOVER"), ("Gross profit", "GROSS PROFIT"), ("EBITDA", "EBITDA"),
             ("Finance cost", "TOTAL FINANCE COST"), ("Valuation provisions (inventory overhead, FOC)", "Total Valuation Provisions"),
             ("PAT (before depreciation & tax - not booked)", "PAT (NET PROFIT)")]
    for name, lab in heads:
        label(ws, r, name, F_BOLD if name.startswith(("Turnover", "EBITDA", "PAT")) else F_NORM)
        num_cell(ws, r, 2, cr(gt(pl_c, lab, "Q1 Total")), NUM_CR)
        num_cell(ws, r, 3, cr(gt(pl_c, lab, "Q2 Total")), NUM_CR)
        num_cell(ws, r, 4, f"=IFERROR(C{r}/B{r}-1,\"\")", PCT)
        num_cell(ws, r, 5, cr(gt(pl_c, lab, "Apr-Sep Total")), NUM_CR)
        num_cell(ws, r, 6, cr(gt(pl_c, lab, "Sep-26")), NUM_CR)
        marks[lab] = r
        r += 1
    for name, num, den in (("Gross margin %", "GROSS PROFIT", "TOTAL TURNOVER"), ("EBITDA margin %", "EBITDA", "TOTAL TURNOVER"),
                           ("PAT margin %", "PAT (NET PROFIT)", "TOTAL TURNOVER")):
        label(ws, r, name)
        for c in (2, 3, 5, 6):
            L_ = col_letter(c)
            num_cell(ws, r, c, f"=IFERROR({L_}{marks[num]}/{L_}{marks[den]},\"\")", PCT)
        num_cell(ws, r, 4, f"=IFERROR(C{r}-B{r},\"\")", '+0.0%;[Red]-0.0%;-')
        marks[name] = r
        r += 1
    # Adjusted PAT estimate
    label(ws, r, "Estimated depreciation for the period (half of FY 25-26 Rs crore, input)", F_NOTE)
    num_cell(ws, r, 5, round(facts["dep_half_year_cr"], 2), NUM_CR, fill=FILL_KEY)
    marks["dep_est"] = r
    r += 1
    label(ws, r, "Estimated tax @ 25.17% on (PAT - depreciation) (input rate in B)", F_NOTE)
    num_cell(ws, r, 2, 0.2517, '0.00%', fill=FILL_KEY)
    num_cell(ws, r, 5, f"=MAX(E{marks['PAT (NET PROFIT)']}-E{marks['dep_est']},0)*B{r}", NUM_CR)
    marks["tax_est"] = r
    r += 1
    label(ws, r, "ADJUSTED PAT - ESTIMATE (after depreciation and tax)", F_RED)
    num_cell(ws, r, 5, f"=E{marks['PAT (NET PROFIT)']}-E{marks['dep_est']}-E{marks['tax_est']}", NUM_CR, F_RED)
    marks["adj_pat"] = r
    r += 2

    # ---- 2. Balance sheet & cash ------------------------------------------------------------------
    band(ws, r, "2. BALANCE SHEET, BORROWINGS AND CASH - GROUP (CFS)", 12)
    r += 1
    bs_hdr = ["Rs crore", "31-Mar-26", "30-Jun-26", "30-Sep-26", "Change Apr-Sep"]
    for i, h in enumerate(bs_hdr, start=1):
        hdr_cell(ws, r, i, h)
    r += 1
    op = "'Opening BS 31-Mar-26'"
    open_rows = facts["open_rows"]
    gt_col = col_letter(5 + BLOCK8.index("GT"))
    bs_groups = bs_group_rows(L["BS_CFS_LAY"])

    def bs_ref(group, month):
        if month is None:
            rr = open_rows.get(("CFS", group))
            return f"{op}!${gt_col}${rr}" if rr else "0"
        rr = (L["BS_CFS_LAY"].row("Subtotal: Closing Stock (Inventory)") if group == "STOCK"
              else bs_groups.get(group, (None,))[0])
        c = L["BS_CFS_LAY"].col.get((month, "GT"))
        return f"'{L['BS_CFS_LAY'].title}'!${col_letter(c)}${rr}" if rr and c else "0"

    bsl = [("Sundry debtors", ["Sundry Debtors"]), ("Inventory", ["STOCK"]),
           ("Sundry creditors", ["Sundry Creditors"]), ("Bank OD / cash credit", ["Bank OD A/c"]),
           ("Term / secured loans", ["Secured Loans"]), ("Other loans", ["Loans (Liability)"]),
           ("Cash & bank", ["Bank Accounts", "Cash-in-hand"]),
           ("Equity (capital + reserves + P&L)", ["Capital Account", "Reserves & Surplus", PL_LINE])]
    for name, grs in bsl:
        label(ws, r, name, F_BOLD if name.startswith(("Sundry debtors", "Bank OD", "Cash")) else F_NORM)
        for c, m in ((2, None), (3, "Jun-26"), (4, "Sep-26")):
            num_cell(ws, r, c, cr("+".join(bs_ref(g, m) for g in grs)), NUM_CR)
        num_cell(ws, r, 5, f"=D{r}-B{r}", NUM_CR)
        marks[name] = r
        r += 1
    label(ws, r, "Gross debt (OD + secured + other loans)", F_BOLD)
    for c in (2, 3, 4, 5):
        Lc = col_letter(c)
        num_cell(ws, r, c, f"={Lc}{marks['Bank OD / cash credit']}+{Lc}{marks['Term / secured loans']}+{Lc}{marks['Other loans']}", NUM_CR, F_BOLD)
    marks["debt"] = r
    r += 1
    label(ws, r, "Net debt (gross debt - cash & bank)", F_BOLD)
    for c in (2, 3, 4, 5):
        Lc = col_letter(c)
        num_cell(ws, r, c, f"={Lc}{marks['debt']}-{Lc}{marks['Cash & bank']}", NUM_CR, F_BOLD)
    marks["netdebt"] = r
    r += 1
    label(ws, r, "Gross debt / equity (x)")
    for c in (2, 3, 4):
        Lc = col_letter(c)
        num_cell(ws, r, c, f"=IFERROR({Lc}{marks['debt']}/{Lc}{marks['Equity (capital + reserves + P&L)']},\"\")", '0.00"x"')
    r += 1
    label(ws, r, "Interest cover (EBITDA / finance cost, x)")
    for c, src in ((2, 2), (3, 3), (4, 5)):
        num_cell(ws, r, c, f"=IFERROR({col_letter(src)}{marks['EBITDA']}/{col_letter(src)}{marks['TOTAL FINANCE COST']},\"\")", '0.0"x"')
    ws.cell(r, 5, "(Q1 | Q2 | Apr-Sep)").font = F_NOTE
    r += 1
    label(ws, r, "Debtor days (Sep debtors / Apr-Sep turnover x 183)")
    num_cell(ws, r, 4, f"=IFERROR(D{marks['Sundry debtors']}/E{marks['TOTAL TURNOVER']}*183,\"\")", '0')
    marks["debtor_days"] = r
    r += 1
    label(ws, r, "Inventory days (Sep stock / Apr-Sep cost of sales x 183)")
    num_cell(ws, r, 4, f"=IFERROR(D{marks['Inventory']}/(({gt(pl_c, 'TOTAL COST OF SALES', 'Apr-Sep Total')})/{int(CR)})*183,\"\")", '0')
    r += 1
    label(ws, r, "Creditor days (Sep creditors / Apr-Sep cost of sales x 183)")
    num_cell(ws, r, 4, f"=IFERROR(D{marks['Sundry creditors']}/(({gt(pl_c, 'TOTAL COST OF SALES', 'Apr-Sep Total')})/{int(CR)})*183,\"\")", '0')
    r += 2
    page_break()

    # ---- 3. Cash story ----------------------------------------------------------------------------
    band(ws, r, "3. WHERE THE CASH WENT - GROUP CASH FLOW (CFS, Rs crore)", 12)
    r += 1
    for i, h in enumerate(["Rs crore", "Q1", "Q2", "Apr-Sep"], start=1):
        hdr_cell(ws, r, i, h)
    r += 1

    def cfref(key, period):
        return f"{cf_c}!${col_letter(cf_pos[(period, 'GT')])}${cf_rows[key]}"

    # The working: every operating line of the cash flow sheet (user, 2026-10-24: "show the working").
    story = [("Profit before tax", "pbt"), ("Add: finance cost (shown under financing)", "fin"),
             ("Tax and depreciation (not booked Apr-Sep)", "taxdep"),
             ("Operating profit before working capital", "opwc"), ("Inventory (increase) / decrease", "inv"),
             ("Debtors (increase) / decrease", "deb"), ("Loans & advances (increase) / decrease", "lna"),
             ("Deposits (increase) / decrease", "dep_a"), ("Other current assets incl. advance tax / TDS", "oca"),
             ("Creditors increase / (decrease)", "cred"), ("Other current liabilities increase / (decrease)", "ocl"),
             ("Duties & taxes increase / (decrease)", "dt"), ("Provisions increase / (decrease)", "prov"),
             ("NET CASH FROM OPERATIONS (A)", "netA"),
             ("Capex and investments (B)", "netB"), ("Loans and OD raised, net of finance cost (C)", "netC"),
             ("P&L A/c vs PAT timing (D)", "D"), ("Inter-company unmatched (E)", "E"),
             ("NET CHANGE IN CASH & BANK", "net"), ("Check: cash per Balance Sheet difference (must be nil)", "chk")]
    for name, key in story:
        bold = name.isupper() or name.startswith("NET")
        label(ws, r, name, F_BOLD if bold else F_NORM, FILL_CHECK if key == "chk" else None)
        for c, p in ((2, "Q1 Total"), (3, "Q2 Total"), (4, "Apr-Sep Total")):
            if key == "taxdep":
                num_cell(ws, r, c, cr(f"{cfref('tax', p)}+{cfref('dep', p)}"), NUM_CR)
            else:
                num_cell(ws, r, c, cr(cfref(key, p)), NUM_CR, F_BOLD if bold else F_NORM)
        marks["cf_" + key] = r
        r += 1
    r += 1

    # ---- 4. Entity view (SFS) ---------------------------------------------------------------------
    band(ws, r, "4. ENTITY PERFORMANCE - STANDALONE (SFS). Use these, not the CFS entity columns, for entity results.", 12)
    r += 1
    ents = [("TOT_OTEC", "Orange O Tec"), ("TOT_ENT", "Enterprises"), ("COLORIX", "Colorix"), ("GT", "Group (SFS)")]
    hdr_cell(ws, r, 1, "Rs crore")
    c = 2
    for _k, n in ents:
        for sub in ("Q1", "Q2", "Apr-Sep"):
            hdr_cell(ws, r, c, f"{n}\n{sub}")
            c += 1
    ws.row_dimensions[r].height = 42
    r += 1
    for name, lab in (("Turnover", "TOTAL TURNOVER"), ("Gross profit", "GROSS PROFIT"), ("EBITDA", "EBITDA"),
                      ("Finance cost", "TOTAL FINANCE COST"), ("PAT (pre dep & tax)", "PAT (NET PROFIT)")):
        label(ws, r, name, F_BOLD if name in ("EBITDA", "PAT (pre dep & tax)") else F_NORM)
        c = 2
        for k, _n in ents:
            for p in ("Q1 Total", "Q2 Total", "Apr-Sep Total"):
                num_cell(ws, r, c, cr(ref_or_zero(pl_s.ref(lab, p, k))), NUM_CR)
                c += 1
        marks["ent_" + lab] = r
        r += 1
    label(ws, r, "Gross margin %")
    for c in range(2, 14):
        Lc = col_letter(c)
        num_cell(ws, r, c, f"=IFERROR({Lc}{marks['ent_GROSS PROFIT']}/{Lc}{marks['ent_TOTAL TURNOVER']},\"\")", PCT)
    r += 1
    label(ws, r, "EBITDA margin %")
    for c in range(2, 14):
        Lc = col_letter(c)
        num_cell(ws, r, c, f"=IFERROR({Lc}{marks['ent_EBITDA']}/{Lc}{marks['ent_TOTAL TURNOVER']},\"\")", PCT)
    r += 2
    page_break()

    # ---- 5. Q2 vs Q1 movers -----------------------------------------------------------------------
    band(ws, r, "5. WHY THE Q2 MARGIN FELL - TOP 10 GL LINES vs SALES GROWTH (CFS Grand Total, Rs crore)", 12)
    r += 1
    label(ws, r, "Each line's Q1 amount is scaled to Q2 sales (x Q2/Q1 turnover, cell B). The effect is what the line cost beyond "
                 "sales growth: negative = squeezed the margin, positive = helped it.", F_NOTE)
    r += 1
    label(ws, r, "Q2 turnover / Q1 turnover", F_BOLD)
    k_r = r
    marks["turn_ratio"] = r
    num_cell(ws, r, 2, f"=C{marks['TOTAL TURNOVER']}/B{marks['TOTAL TURNOVER']}", '0.000"x"', F_BOLD)
    r += 1
    for i, h in enumerate(["GL line", "Section", "Q1", "Q2", "Q1 scaled to Q2 sales", "Effect on margin"], start=1):
        hdr_cell(ws, r, i, h)
    r += 1
    marks["movers_first"] = r
    for m in movers:
        label(ws, r, m["gl"])
        ws.cell(r, 2, m["section"]).font = F_NOTE
        num_cell(ws, r, 3, cr(m["q1_ref"]), NUM_CR)
        num_cell(ws, r, 4, cr(m["q2_ref"]), NUM_CR)
        num_cell(ws, r, 5, f"=C{r}*$B${k_r}", NUM_CR)
        num_cell(ws, r, 6, f"=(D{r}-E{r})*{m['sign']}", NUM_CR, F_BOLD)
        r += 1
    marks["movers_last"] = r - 1
    label(ws, r, "Total of the 10 lines", F_BOLD, FILL_SUB)
    num_cell(ws, r, 6, f"=SUM(F{marks['movers_first']}:F{r - 1})", NUM_CR, F_BOLD, FILL_SUB)
    marks["movers_total"] = r
    r += 2
    label(ws, r, "5b. Operating expenses - the sub-groups that grew most, Q2 vs Q1 (Rs crore)", F_BOLD)
    r += 1
    for i, h in enumerate(["Expense sub-group", "", "Q1", "Q2", "Increase", "Share of opex increase"], start=1):
        if h:
            hdr_cell(ws, r, i, h)
    r += 1
    op_tot = (f"((({gt(pl_c, 'TOTAL OPERATING EXPENSES', 'Q2 Total')})-({gt(pl_c, 'TOTAL OPERATING EXPENSES', 'Q1 Total')}))"
              f"/{int(CR)})")
    for m in facts["opex_movers"]:
        label(ws, r, m["name"])
        num_cell(ws, r, 3, cr(m["q1_ref"]), NUM_CR)
        num_cell(ws, r, 4, cr(m["q2_ref"]), NUM_CR)
        num_cell(ws, r, 5, f"=D{r}-C{r}", NUM_CR, F_BOLD)
        num_cell(ws, r, 6, f"=IFERROR(E{r}/{op_tot},\"\")", PCT)
        r += 1
    label(ws, r, "Total operating expenses", F_BOLD, FILL_SUB)
    num_cell(ws, r, 3, cr(gt(pl_c, "TOTAL OPERATING EXPENSES", "Q1 Total")), NUM_CR, F_BOLD, FILL_SUB)
    num_cell(ws, r, 4, cr(gt(pl_c, "TOTAL OPERATING EXPENSES", "Q2 Total")), NUM_CR, F_BOLD, FILL_SUB)
    num_cell(ws, r, 5, f"=D{r}-C{r}", NUM_CR, F_BOLD, FILL_SUB)
    r += 2
    page_break()

    # ---- Annexes A and B: working controls, GROUPED AND HIDDEN (user, 2026-10-24) ------------------
    # They prove the pack - they are not part of the report. Collapsed rows do not print, so the PDF
    # leaves them out; the outline [+] reopens them, and the verifier still reads every check.
    annex_first = r
    band(ws, r, "Annex A - TIE-OUT TO THE EXPENSE REPORT (Expense Sheet 2026-10-10, Rs lakh) - GL file vs Expense Sheet", 12)
    r += 1
    for i, h in enumerate(["Item", "Period", "GL file (this pack)", "Expense Sheet", "Gap", "Reason"], start=1):
        hdr_cell(ws, r, i, h)
    r += 1
    tie_first = r
    for t in tieout:
        label(ws, r, t["item"])
        ws.cell(r, 2, t["period"]).font = F_NORM
        num_cell(ws, r, 3, f"=({t['gl_expr']})/100000", '#,##0.00;[Red](#,##0.00);-')
        num_cell(ws, r, 4, round(t["exp"], 2), '#,##0.00;[Red](#,##0.00);-')
        num_cell(ws, r, 5, f"=C{r}-D{r}", '#,##0.00;[Red](#,##0.00);0.00', F_BOLD)
        ws.cell(r, 6, t["reason"]).font = F_NOTE
        r += 1
    ws.cell(r, 1, "The Expense Sheet is the older Excel method (pulled 10-Oct); the Hub audit of 23-Oct found it reverses the sign of "
                  "328 credit lines and counts non-accounting vouchers. The GL file (Trial Balance based, 16-Oct) is the figure to quote.").font = F_NOTE
    r += 2

    # ---- 7. Checks --------------------------------------------------------------------------------
    band(ws, r, "Annex B - CHECKS - every control in this pack (must read 0.00)", 12)
    r += 1
    hdr_cell(ws, r, 1, "Control")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    hdr_cell(ws, r, 6, "Result (Rs)")
    hdr_cell(ws, r, 7, "Status")
    r += 1
    checks = facts["checks"]
    for name, expr in checks:
        label(ws, r, name)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        num_cell(ws, r, 6, f"=ROUND({expr},2)", '#,##0.00;[Red]-#,##0.00;0.00')
        st = ws.cell(r, 7, f'=IF(ABS(F{r})<1,"OK","CHECK")')
        st.font = F_BOLD
        r += 1
    ws.row_dimensions.group(annex_first, r - 1, outline_level=1, hidden=True)
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.freeze_panes = "A4"
    setup_print(ws, landscape=True, fit_wide=1)
    return ws, marks


def setup_print(ws, landscape=True, fit_wide=1):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = fit_wide
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.75
    ws.page_margins.header = ws.page_margins.footer = 0.3
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.left.text = "&A"
    ws.oddHeader.right.text = "Orange O Tec Group - CFO Pack"


def build_print_statements(wb, L, open_rows):
    """'PDF - Statements': condensed P&L, BS, Cash Flow and Fund Flow, Rs crore, A4 landscape.
    Group (CFS) Q1 | Q2 | Apr-Sep | Sep, then entity Apr-Sep (SFS)."""
    ws = wb.create_sheet("PDF - Statements")
    ws.sheet_properties.tabColor = "C00000"
    ws.column_dimensions["A"].width = 52
    for c in range(2, 10):
        ws.column_dimensions[col_letter(c)].width = 13
    cols = [("CFS", "GT", "Q1 Total", "Group Q1"), ("CFS", "GT", "Q2 Total", "Group Q2"),
            ("CFS", "GT", "Apr-Sep Total", "Group Apr-Sep"), ("CFS", "GT", "Sep-26", "Group Sep-26"),
            ("SFS", "TOT_OTEC", "Apr-Sep Total", "O Tec Apr-Sep (SFS)"), ("SFS", "TOT_ENT", "Apr-Sep Total", "Enterprises Apr-Sep (SFS)"),
            ("SFS", "COLORIX", "Apr-Sep Total", "Colorix Apr-Sep (SFS)"), ("SFS", "GT", "Apr-Sep Total", "Group Apr-Sep (SFS)")]
    r = 1
    ws.cell(r, 1, "CONDENSED STATEMENTS - FY 2026-27 Apr-Sep (Rs crore)").font = F_TITLE
    r += 1
    ws.cell(r, 1, "Group = consolidated (CFS, inter-company eliminated). Entity columns = STANDALONE (SFS). Every figure links to the detail sheets.").font = F_NOTE
    r += 2

    def header():
        nonlocal r
        hdr_cell(ws, r, 1, "Rs crore")
        for i, (_b, _k, _p, h) in enumerate(cols, start=2):
            hdr_cell(ws, r, i, h)
        ws.row_dimensions[r].height = 30
        r += 1

    band(ws, r, "PROFIT & LOSS", 9)
    r += 1
    header()
    for name, lab in (("Turnover", "TOTAL TURNOVER"), ("Cost of sales", "TOTAL COST OF SALES"), ("Gross profit", "GROSS PROFIT"),
                      ("Other income", "TOTAL OTHER INCOME"), ("Operating expenses", "TOTAL OPERATING EXPENSES"),
                      ("EBITDA", "EBITDA"), ("Finance cost", "TOTAL FINANCE COST"),
                      ("Valuation provisions", "Total Valuation Provisions"), ("Depreciation (not booked)", "LESS: DEPRECIATION"),
                      ("PBT", "PBT (PROFIT BEFORE TAX)"), ("PAT (pre tax - not booked)", "PAT (NET PROFIT)")):
        bold = name in ("Gross profit", "EBITDA", "PBT", "Turnover")
        label(ws, r, name, F_BOLD if bold else F_NORM, FILL_KEY if bold else None)
        for i, (b, k, p, _h) in enumerate(cols, start=2):
            lay = L["PL_" + b]
            num_cell(ws, r, i, f"=({ref_or_zero(lay.ref(lab, p, k))})/{int(CR)}", NUM_CR, F_BOLD if bold else F_NORM)
        r += 1
    r += 1
    ws.row_breaks.append(Break(id=r - 1))
    band(ws, r, "BALANCE SHEET", 9)
    r += 1
    # Positions, not flows: group at 31-Mar / 30-Jun / 30-Sep and the change, then each entity at 30-Sep.
    bs_cols = [("OPEN", "GT", None, "Group 31-Mar-26"), ("CFS", "GT", "Jun-26", "Group 30-Jun-26"),
               ("CFS", "GT", "Sep-26", "Group 30-Sep-26"), ("CHG", "GT", None, "Group change Apr-Sep"),
               ("SFS", "TOT_OTEC", "Sep-26", "O Tec 30-Sep (SFS)"), ("SFS", "TOT_ENT", "Sep-26", "Enterprises 30-Sep (SFS)"),
               ("SFS", "COLORIX", "Sep-26", "Colorix 30-Sep (SFS)"), ("SFS", "GT", "Sep-26", "Group 30-Sep (SFS)")]
    hdr_cell(ws, r, 1, "Rs crore")
    for i, (_b, _k, _m, h) in enumerate(bs_cols, start=2):
        hdr_cell(ws, r, i, h)
    ws.row_dimensions[r].height = 30
    r += 1
    op_gt = col_letter(5 + BLOCK8.index("GT"))
    for name, grs in (("Capital + reserves", ["Capital Account", "Reserves & Surplus"]), ("Profit & Loss A/c", [PL_LINE]),
                      ("Secured / term loans", ["Secured Loans"]), ("Unsecured loans", ["Unsecured Loans"]),
                      ("Other loans", ["Loans (Liability)"]), ("Bank OD / cash credit", ["Bank OD A/c"]),
                      ("Sundry creditors", ["Sundry Creditors"]),
                      ("Other current liabilities, duties, provisions", ["Current Liabilities", "Duties & Taxes", "Provisions", "Suspense A/c", "Unadjusted Forex Gain/Loss", "Branch / Divisions"]),
                      ("Fixed assets", ["Fixed Assets"]), ("Investments", ["Investments", "Non-Current Investments"]),
                      ("Inventory", ["STOCK"]), ("Sundry debtors", ["Sundry Debtors"]),
                      ("Loans, advances, deposits, other current assets", ["CA_OTHER", "Loans & Advances (Asset)", "Deposits (Asset)"]),
                      ("Cash & bank", ["Bank Accounts", "Cash-in-hand"])):
        bold = name in ("Sundry debtors", "Bank OD / cash credit", "Cash & bank")
        label(ws, r, name, F_BOLD if bold else F_NORM)
        for i, (b, k, m, _h) in enumerate(bs_cols, start=2):
            if b == "CHG":
                num_cell(ws, r, i, f"={col_letter(4)}{r}-{col_letter(2)}{r}", NUM_CR, F_BOLD)
                continue
            if b == "OPEN":
                parts = []
                for g in grs:
                    if g == "CA_OTHER":
                        a, s = open_rows.get(("CFS", "Current Assets")), open_rows.get(("CFS", "STOCK"))
                        parts.append(f"('Opening BS 31-Mar-26'!{op_gt}{a}-'Opening BS 31-Mar-26'!{op_gt}{s})")
                    elif open_rows.get(("CFS", g)):
                        parts.append(f"'Opening BS 31-Mar-26'!{op_gt}{open_rows[('CFS', g)]}")
                num_cell(ws, r, i, f"=({'+'.join(parts) or '0'})/{int(CR)}", NUM_CR, F_BOLD if bold else F_NORM)
                continue
            lay = L["BS_" + b + "_LAY"]
            groups = bs_group_rows(lay)
            c = lay.col.get((m, k))
            parts = []
            for g in grs:
                if g == "STOCK":
                    rr = lay.row("Subtotal: Closing Stock (Inventory)")
                    parts.append(f"'{lay.title}'!{col_letter(c)}{rr}")
                elif g == "CA_OTHER":
                    rr, rs = groups["Current Assets"][0], lay.row("Subtotal: Closing Stock (Inventory)")
                    parts.append(f"('{lay.title}'!{col_letter(c)}{rr}-'{lay.title}'!{col_letter(c)}{rs})")
                elif g in groups:
                    parts.append(f"'{lay.title}'!{col_letter(c)}{groups[g][0]}")
            num_cell(ws, r, i, f"=({'+'.join(parts) or '0'})/{int(CR)}", NUM_CR, F_BOLD if bold else F_NORM)
        r += 1
    r += 1
    ws.row_breaks.append(Break(id=r - 1))
    # Flows: group columns from the CONSOLIDATED sheets, entity columns from the STANDALONE sheets -
    # the consolidated entity columns carry inter-company re-allocations and would mislead here.
    flow_cols = [("CFS", "GT", "Q1 Total", "Group Q1"), ("CFS", "GT", "Q2 Total", "Group Q2"),
                 ("CFS", "GT", "Apr-Sep Total", "Group Apr-Sep"), ("CFS", "GT", "Sep-26", "Group Sep-26"),
                 ("SFS", "TOT_OTEC", "Apr-Sep Total", "O Tec Apr-Sep (SFS)"), ("SFS", "TOT_ENT", "Apr-Sep Total", "Enterprises Apr-Sep (SFS)"),
                 ("SFS", "COLORIX", "Apr-Sep Total", "Colorix Apr-Sep (SFS)"), ("SFS", "GT", "Apr-Sep Total", "Group Apr-Sep (SFS)")]

    def flow_header():
        nonlocal r
        hdr_cell(ws, r, 1, "Rs crore")
        for i, (_b, _k, _p, h) in enumerate(flow_cols, start=2):
            hdr_cell(ws, r, i, h)
        ws.row_dimensions[r].height = 30
        r += 1

    band(ws, r, "CASH FLOW - flows for the period", 9)
    r += 1
    flow_header()
    for name, key in (("Operating profit before working capital", "opwc"), ("Inventory", "inv"), ("Sundry debtors", "deb"),
                      ("Sundry creditors", "cred"), ("Other working capital", "OWC"),
                      ("NET CASH FROM OPERATIONS (A)", "netA"), ("Capex and investments (B)", "netB"),
                      ("  of which: fixed assets", "fa"), ("Loans and OD, net of finance cost (C)", "netC"),
                      ("  of which: Bank OD / cash credit", "od"), ("  of which: term / secured loans", "sec"),
                      ("P&L A/c vs PAT timing (D)", "D"), ("Branch / inter-company unmatched (E)", "E"),
                      ("NET CHANGE IN CASH & BANK", "net"), ("Check - must be nil", "chk")):
        bold = name.isupper() or name.startswith("NET")
        label(ws, r, name, F_BOLD if bold else (F_NOTE if name.startswith("  ") else F_NORM),
              FILL_CHECK if key == "chk" else (FILL_KEY if bold else None))
        for i, (b, k, p, _h) in enumerate(flow_cols, start=2):
            cf = L["CF_" + b]
            sheet = f"'MOM Cash Flow {b}'"
            colx = col_letter(cf["pos"][(p, k)])
            R = cf["rows"]
            if key == "OWC":
                f = (f"({sheet}!{colx}{R['netA']}-{sheet}!{colx}{R['opwc']}-{sheet}!{colx}{R['inv']}"
                     f"-{sheet}!{colx}{R['deb']}-{sheet}!{colx}{R['cred']})")
            else:
                f = f"{sheet}!{colx}{R[key]}"
            num_cell(ws, r, i, f"=({f})/{int(CR)}", NUM_CR, F_BOLD if bold else F_NORM)
        r += 1
    r += 1
    band(ws, r, "FUND FLOW (CMA style)", 9)
    r += 1
    flow_header()
    for name, key in (("Funds from operations (PAT + depreciation)", "fop"), ("Long-term loans raised", "LTR"),
                      ("Other sources (sale of investments / assets, capital, P&L timing, inter-company)", "OS"),
                      ("TOTAL SOURCES", "TS"), ("Purchase of fixed assets", "faa"), ("Purchase of investments", "inva"),
                      ("Other applications (loans repaid, capital withdrawn)", "OA"),
                      ("TOTAL APPLICATIONS", "TA"), ("INCREASE / (DECREASE) IN NET WORKING CAPITAL", "NWC"),
                      ("  financed by: increase in Bank OD (current liability)", "w_od"), ("Check - must be nil", "chk")):
        bold = name.isupper()
        label(ws, r, name, F_BOLD if bold else (F_NOTE if name.startswith("  ") else F_NORM),
              FILL_CHECK if key == "chk" else (FILL_KEY if bold else None))
        for i, (b, k, p, _h) in enumerate(flow_cols, start=2):
            ff = L["FF_" + b]
            sheet = f"'MOM Fund Flow {b}'"
            colx = col_letter(ff["pos"][(p, k)])
            R = ff["rows"]
            ltr = f"({sheet}!{colx}{R['seci']}+{sheet}!{colx}{R['unsi']}+{sheet}!{colx}{R['othi']})"
            if key == "LTR":
                f = ltr
            elif key == "OS":
                f = f"({sheet}!{colx}{R['TS']}-{sheet}!{colx}{R['fop']}-{ltr})"
            elif key == "OA":
                f = f"({sheet}!{colx}{R['TA']}-{sheet}!{colx}{R['faa']}-{sheet}!{colx}{R['inva']})"
            else:
                f = f"{sheet}!{colx}{R[key]}"
            num_cell(ws, r, i, f"=({f})/{int(CR)}", NUM_CR, F_BOLD if bold else F_NORM)
        r += 1
    ws.cell(r + 1, 1, "Group columns are CONSOLIDATED (inter-company eliminated); entity columns are STANDALONE. Depreciation and tax "
                      "are not booked for Apr-Sep. Full detail: the MOM Cash Flow / Fund Flow sheets.").font = F_NOTE
    setup_print(ws, landscape=True, fit_wide=1)
    return ws


# Month-on-month views: (name, basis, column key). The group and the Orange + Enterprises view are
# CONSOLIDATED; the three entities are STANDALONE, because consolidated entity columns move
# inter-company profit between companies. "OE" = CFS Total of Orange O Tec + Total of Enterprises
# (Colorix left out) - user, 2026-10-24.
MOM_VIEWS = [("Group (CFS)", "CFS", "GT"), ("Orange + Enterprises (CFS)", "CFS", "OE"),
             ("Orange O Tec (SFS)", "SFS", "TOT_OTEC"), ("Enterprises (SFS)", "SFS", "TOT_ENT"),
             ("Colorix (SFS)", "SFS", "COLORIX")]
MOM_PERIODS = [("Apr-26", "Apr-26"), ("May-26", "May-26"), ("Jun-26", "Jun-26"), ("Q1 Total", "Jun-26"),
               ("Jul-26", "Jul-26"), ("Aug-26", "Aug-26"), ("Sep-26", "Sep-26"), ("Q2 Total", "Sep-26"),
               ("Apr-Sep Total", "Sep-26")]
MOM_HEAD = {"Q1 Total": "Q1", "Q2 Total": "Q2", "Apr-Sep Total": "Apr-Sep"}
MOM_DAYS = {"Apr-26": 30, "May-26": 31, "Jun-26": 30, "Q1 Total": 91, "Jul-26": 31, "Aug-26": 31, "Sep-26": 30,
            "Q2 Total": 92, "Apr-Sep Total": 183}
CURRENT_ASSET_GROUPS = ["Current Assets", "Sundry Debtors", "Loans & Advances (Asset)", "Deposits (Asset)",
                        "Bank Accounts", "Cash-in-hand"]
CURRENT_LIAB_GROUPS = ["Sundry Creditors", "Current Liabilities", "Duties & Taxes", "Provisions", "Bank OD A/c"]
# (label, kind, source). kind: pl / bs / cf / band / pct (num, den) / x (num, den) / days (num, den)
MOM_LINES = [
    ("PROFIT & LOSS", "band", None),
    ("Turnover", "pl", "TOTAL TURNOVER"), ("Cost of sales", "pl", "TOTAL COST OF SALES"),
    ("Gross profit", "pl", "GROSS PROFIT"), ("Other income", "pl", "TOTAL OTHER INCOME"),
    ("Operating expenses", "pl", "TOTAL OPERATING EXPENSES"), ("EBITDA", "pl", "EBITDA"),
    ("Finance cost", "pl", "TOTAL FINANCE COST"), ("Valuation provisions", "pl", "Total Valuation Provisions"),
    ("PBT", "pl", "PBT (PROFIT BEFORE TAX)"), ("PAT (before depreciation & tax - not booked)", "pl", "PAT (NET PROFIT)"),
    ("BALANCE SHEET (at period end)", "band", None),
    ("Sundry debtors", "bs", ["Sundry Debtors"]), ("Inventory", "bs", ["STOCK"]),
    ("Sundry creditors", "bs", ["Sundry Creditors"]), ("Bank OD / cash credit", "bs", ["Bank OD A/c"]),
    ("Term / secured loans", "bs", ["Secured Loans"]), ("Cash & bank", "bs", ["Bank Accounts", "Cash-in-hand"]),
    ("Current assets (incl. stock, debtors, cash)", "bs", CURRENT_ASSET_GROUPS),
    ("Current liabilities (incl. Bank OD)", "bs", CURRENT_LIAB_GROUPS),
    ("Gross debt (OD + all loans)", "bs", ["Bank OD A/c", "Secured Loans", "Unsecured Loans", "Loans (Liability)"]),
    ("Equity (capital + reserves + P&L)", "bs", ["Capital Account", "Reserves & Surplus", PL_LINE]),
    ("CASH FLOW (for the period)", "band", None),
    ("Net cash from operations (A)", "cf", "netA"), ("Capex and investments (B)", "cf", "netB"),
    ("Loans and OD, net of finance cost (C)", "cf", "netC"), ("Net change in cash & bank", "cf", "net"),
    ("RATIOS", "band", None),
    ("Gross margin %", "pct", ("Gross profit", "Turnover")),
    ("EBITDA margin %", "pct", ("EBITDA", "Turnover")),
    ("PAT margin %", "pct", ("PAT (before depreciation & tax - not booked)", "Turnover")),
    ("Operating expenses % of sales", "pct", ("Operating expenses", "Turnover")),
    ("Finance cost % of sales", "pct", ("Finance cost", "Turnover")),
    ("Interest cover (x)", "x", ("EBITDA", "Finance cost")),
    ("Debtor days", "days", ("Sundry debtors", "Turnover")),
    ("Inventory days", "days", ("Inventory", "Cost of sales")),
    ("Creditor days", "days", ("Sundry creditors", "Cost of sales")),
    ("Current ratio (x)", "x", ("Current assets (incl. stock, debtors, cash)", "Current liabilities (incl. Bank OD)")),
    ("Gross debt / equity (x)", "x", ("Gross debt (OD + all loans)", "Equity (capital + reserves + P&L)")),
]
MOM_FMT = {"pct": PCT, "x": '0.00"x";[Red]-0.00"x";-', "days": '0;[Red]-0;-'}


def keys_of(key):
    """The sheet columns behind one view key ('OE' is two columns added together)."""
    return ["TOT_OTEC", "TOT_ENT"] if key == "OE" else [key]


# Where each MOM line's GL detail lives. P&L: (section header, section total) on the MOM PL sheets.
# BS: the balance-sheet groups (column B) on the MOM BS sheets; STOCK is the stock sub-group.
MOM_PL_SECTIONS = {
    "Turnover": ("TURNOVER", "TOTAL TURNOVER"),
    "Cost of sales": ("LESS: COST OF SALES", "TOTAL COST OF SALES"),
    "Other income": ("ADD: OTHER INCOME", "TOTAL OTHER INCOME"),
    "Operating expenses": ("LESS: OPERATING EXPENSES", "TOTAL OPERATING EXPENSES"),
    "Finance cost": ("LESS: FINANCE COST", "TOTAL FINANCE COST"),
    "Valuation provisions": ("LESS: VALUATION", "Total Valuation Provisions"),
}
# The two stock lines inside cost of sales read as one movement (user, 2026-10-24).
SUB_RENAME = {"Stock Valuation (Inventory)": "Change in inventory (opening - closing stock)"}


def gl_tree(lay: Layout, line, kind, src):
    """[(sub-group, subtotal row or None, [(gl, row), ...]), ...] under one MOM line, in sheet order."""
    out, cur = [], None
    if kind == "pl" and line in MOM_PL_SECTIONS:
        start, end = (lay.row(x) for x in MOM_PL_SECTIONS[line])
        if not start or not end:
            return []
        sub = None
        for r in range(start + 1, end):
            c, d = lay.cell(r, 3), lay.cell(r, 4)
            if isinstance(c, str) and c.strip():
                sub = SUB_RENAME.get(c.strip(), c.strip())
            if not isinstance(d, str) or not d.strip() or d.startswith("TOTAL"):
                continue
            if d.startswith("Subtotal: "):
                if cur is not None:
                    out.append((cur[0], r, cur[1]))
                cur = None
                continue
            if cur is None:
                cur = (sub or d.strip(), [])
            cur[1].append((d.strip(), r))
        if cur is not None:
            out.append((cur[0], None, cur[1]))
        return out
    if kind == "bs":
        want = set(src)
        group = sub = None
        stop = lay.row("RECONCILIATION") or lay.max_row
        for r in range(5, stop):
            b, c, d = lay.cell(r, 2), lay.cell(r, 3), lay.cell(r, 4)
            if isinstance(b, str) and b.strip() and not (isinstance(d, str) and d.startswith("TOTAL")):
                group = b.strip()
            if isinstance(c, str) and c.strip():
                sub = c.strip()
            if not isinstance(d, str) or d.startswith("TOTAL"):
                continue
            hit = (group in want) or ("STOCK" in want and group == "Current Assets" and sub == "Closing Stock (Inventory)")
            if not hit:
                cur = None
                continue
            if d.startswith("Subtotal: "):
                if cur is not None:
                    out.append((cur[0], r, cur[1]))
                cur = None
                continue
            if cur is None:
                cur = (sub or group, [])
            cur[1].append((d.strip(), r))
    return out


def mom_detail_values(L):
    """{view: {line: [{sub, values, ledgers: [{gl, values}]}]}} - the drill-down data for the Hub,
    from the source's own values (the copied sheets are proved equal to it). Nil ledgers dropped."""
    out = {}
    for name, basis, key in MOM_VIEWS:
        out[name] = {}
        for line, kind, src in MOM_LINES:
            if kind not in ("pl", "bs"):
                continue
            lay = L[f"PL_{basis}"] if kind == "pl" else L[f"BS_{basis}_LAY"]
            cols = [[lay.col.get(((p if kind == "pl" else m), k)) for k in keys_of(key)] for p, m in MOM_PERIODS]
            val = lambda r: [round(sum((lay.cell(r, c) or 0.0) for c in cs if c) / CR, 4) for cs in cols]
            nodes = []
            for sub, sub_r, leds in gl_tree(lay, line, kind, src):
                ls = [{"gl": gl, "values": val(r)} for gl, r in leds]
                ls = [x for x in ls if any(abs(v) >= 1e-6 for v in x["values"])]
                if not ls:
                    continue
                sv = val(sub_r) if sub_r else [round(sum(x["values"][i] for x in ls), 4) for i in range(len(cols))]
                nodes.append({"sub": sub, "values": sv, "ledgers": ls})
            if nodes:
                out[name][line] = nodes
    return out


def mom_expr(L, basis, key, kind, src, period, month):
    """Excel expression (Rs) for one MOM cell, or None."""
    parts = []
    for k in keys_of(key):
        if kind == "pl":
            ref = L["PL_" + basis].ref(src, period, k)
            if ref:
                parts.append(ref)
        elif kind == "bs":
            lay = L[f"BS_{basis}_LAY"]
            groups, c = bs_group_rows(lay), lay.col.get((month, k))
            for g in src:
                r = lay.row("Subtotal: Closing Stock (Inventory)") if g == "STOCK" else groups.get(g, (None,))[0]
                if r and c:
                    parts.append(f"'{lay.title}'!${col_letter(c)}${r}")
        elif kind == "cf":
            cf = L["CF_" + basis]
            parts.append(f"'MOM Cash Flow {basis}'!${col_letter(cf['pos'][(period, k)])}${cf['rows'][src]}")
    if kind not in ("pl", "bs", "cf"):
        return None
    return "+".join(parts) or "0"


def write_mom_block(ws, r, title, cell_expr, detail_fn=None):
    """One block: title band, header, MOM_LINES rows. cell_expr(kind, src, period, month) -> formula.
    detail_fn(line, kind, src) -> [(outline level, label, expr(period, month))]: collapsed GL rows
    written under the line (the Excel 'drop-down'). Returns (next row, {line label: row})."""
    band(ws, r, title, 1 + len(MOM_PERIODS))
    r += 1
    hdr_cell(ws, r, 1, "Rs crore")
    for i, (p, _m) in enumerate(MOM_PERIODS, start=2):
        hdr_cell(ws, r, i, MOM_HEAD.get(p, p), FILL_TOT if p in MOM_HEAD else FILL_HDR)
    r += 1
    at = {}
    for name, kind, src in MOM_LINES:
        if kind == "band":
            label(ws, r, name, F_BOLD, FILL_SUB)
            for c in range(2, 2 + len(MOM_PERIODS)):
                ws.cell(r, c).fill = FILL_SUB
            r += 1
            continue
        at[name] = r
        bold = name in ("Turnover", "Gross profit", "EBITDA", "Net cash from operations (A)", "Net change in cash & bank")
        label(ws, r, name, F_BOLD if bold else F_NORM)
        for i, (p, m) in enumerate(MOM_PERIODS, start=2):
            Lc = col_letter(i)
            if kind in MOM_FMT:
                num, den = src
                f = f"{Lc}{at[num]}/{Lc}{at[den]}" + (f"*{MOM_DAYS[p]}" if kind == "days" else "")
                num_cell(ws, r, i, f"=IFERROR({f},\"\")", MOM_FMT[kind])
            else:
                num_cell(ws, r, i, cell_expr(kind, src, p, m), NUM_CR, F_BOLD if bold else F_NORM)
        r += 1
        for level, text, expr in (detail_fn(name, kind, src) if detail_fn else []):
            label(ws, r, text, F_NORM if level == 1 else F_NOTE, indent=2 * level)
            for i, (p, m) in enumerate(MOM_PERIODS, start=2):
                num_cell(ws, r, i, expr(p, m), NUM_CR, F_NORM if level == 1 else F_NOTE)
            ws.row_dimensions[r].outlineLevel = level
            ws.row_dimensions[r].hidden = True
            r += 1
    return r, at


def write_pct_of_sales(ws, r, at):
    """'% OF SALES' block: every P&L line of the block above over its Turnover, same columns."""
    band(ws, r, "% OF SALES - every P&L line as a share of turnover", 1 + len(MOM_PERIODS))
    r += 1
    for name, kind, _src in MOM_LINES:
        if kind != "pl":
            continue
        label(ws, r, name, F_BOLD if name in ("Turnover", "Gross profit", "EBITDA") else F_NORM)
        for i in range(2, 2 + len(MOM_PERIODS)):
            Lc = col_letter(i)
            num_cell(ws, r, i, f"=IFERROR({Lc}{at[name]}/{Lc}{at['Turnover']},\"\")", PCT)
        r += 1
    return r


def build_mom_views(wb, L):
    """'MOM Summary' - month by month with a view dropdown (five views); 'PDF - MOM' - the same views
    printed one per page (a PDF cannot hold a dropdown)."""
    names = [v[0] for v in MOM_VIEWS]
    ws = wb.create_sheet("MOM Summary")
    ws.sheet_properties.tabColor = "FF6A1F"
    ws.column_dimensions["A"].width = 46
    for c in range(2, 2 + len(MOM_PERIODS)):
        ws.column_dimensions[col_letter(c)].width = 12
    ws["A1"], ws["A1"].font = "MONTH ON MONTH - choose the view in B3", F_TITLE
    ws["A2"] = ("Rs crore. Group and Orange + Enterprises = consolidated (inter-company eliminated); single companies = "
                "standalone. Q1 / Q2 / Apr-Sep are separate Tally pulls proved additive. Balance sheet lines are at the period end.")
    ws["A2"].font = F_NOTE
    ws["A3"], ws["A3"].font = "View:", F_BOLD
    sel = ws["B3"]
    sel.value = names[0]
    sel.font, sel.fill, sel.border = F_BOLD, FILL_CHECK, BOX
    ws.merge_cells("B3:E3")
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1='"' + ",".join(names) + '"', allow_blank=False)
    dv.prompt, dv.promptTitle = "Pick the view", "View"
    ws.add_data_validation(dv)
    dv.add("B3")
    arr = "{" + ",".join(f'"{n}"' for n in names) + "}"

    def chooser(kind, src, period, month):
        exprs = [mom_expr(L, b, k, kind, src, period, month) or "0" for _n, b, k in MOM_VIEWS]
        return f"=CHOOSE(MATCH($B$3,{arr},0),{','.join(exprs)})/{int(CR)}"

    def lay_for(basis, kind):
        return L[f"PL_{basis}"] if kind == "pl" else L[f"BS_{basis}_LAY"]

    def ref(lay, r, key, kind, p, m):
        refs = []
        for k in keys_of(key):
            c = lay.col.get(((p if kind == "pl" else m), k))
            if r and c:
                refs.append(f"'{lay.title}'!${col_letter(c)}${r}")
        return "+".join(refs) or "0"

    def detail(line, kind, src):
        """Union of the views' GL trees; each cell picks its view with CHOOSE, 0 where that view has no
        such row (e.g. a ledger eliminated on consolidation). P&L to ledger level; balance sheet to
        sub-group level only (debtors / creditors alone run to 1,000+ ledgers)."""
        if kind not in ("pl", "bs"):
            return []
        per_view = []
        for _n, b, k in MOM_VIEWS:
            lay = lay_for(b, kind)
            subs, leds = {}, {}
            for sub, sub_r, ledgers in gl_tree(lay, line, kind, src):
                subs.setdefault(sub, []).append(sub_r)
                for gl, r in ledgers:
                    leds.setdefault((sub, gl), r)
            per_view.append((lay, k, subs, leds))
        order_subs, order_leds = [], {}
        for _lay, _k, subs, leds in per_view:
            for sub in subs:
                if sub not in order_subs:
                    order_subs.append(sub)
            for (sub, gl) in leds:
                order_leds.setdefault(sub, [])
                if gl not in order_leds[sub]:
                    order_leds[sub].append(gl)
        rows = []
        for sub in order_subs:
            def sub_expr(p, m, sub=sub):
                exprs = []
                for lay, k, subs, leds in per_view:
                    rs = [x for x in subs.get(sub, []) if x]
                    if rs:
                        exprs.append("+".join(ref(lay, x, k, kind, p, m) for x in rs))
                    else:  # no subtotal row: add its ledgers
                        exprs.append("+".join(ref(lay, r, k, kind, p, m) for (s2, _g), r in leds.items() if s2 == sub) or "0")
                return f"=CHOOSE(MATCH($B$3,{arr},0),{','.join(exprs)})/{int(CR)}"
            rows.append((1, sub, sub_expr))
            if kind == "pl":
                for gl in order_leds.get(sub, []):
                    def led_expr(p, m, sub=sub, gl=gl):
                        exprs = [ref(lay, leds.get((sub, gl)), k, kind, p, m) for lay, k, _s, leds in per_view]
                        return f"=CHOOSE(MATCH($B$3,{arr},0),{','.join(exprs)})/{int(CR)}"
                    rows.append((2, gl, led_expr))
        return rows

    r, at = write_mom_block(ws, 5, "MONTH ON MONTH - selected view", chooser, detail_fn=detail)
    write_pct_of_sales(ws, r + 1, at)
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = True
    # Quarter grouping: the months fold into their quarter total (B:D -> Q1 in E, F:H -> Q2 in I).
    ws.column_dimensions.group("B", "D", outline_level=1, hidden=False)
    ws.column_dimensions.group("F", "H", outline_level=1, hidden=False)
    ws["A4"] = ("Use the [+] at the left of a line to open its sub-groups and GL ledgers; the [-] above the "
                "month columns folds them into the quarter. % of sales and ratios follow below.")
    ws["A4"].font = F_NOTE
    ws.freeze_panes = "B7"
    setup_print(ws, landscape=True, fit_wide=1)

    pdf = wb.create_sheet("PDF - MOM")
    pdf.sheet_properties.tabColor = "C00000"
    pdf.column_dimensions["A"].width = 46
    for c in range(2, 2 + len(MOM_PERIODS)):
        pdf.column_dimensions[col_letter(c)].width = 12
    pdf["A1"], pdf["A1"].font = "MONTH ON MONTH - Group, Orange + Enterprises, Orange O Tec, Enterprises, Colorix (Rs crore)", F_TITLE
    pdf["A2"] = "Group and Orange + Enterprises = consolidated; single companies = standalone. Balance sheet lines at the period end."
    pdf["A2"].font = F_NOTE
    r = 4
    for n, b, k in MOM_VIEWS:
        r, _at = write_mom_block(pdf, r, f"MONTH ON MONTH - {n}",
                                 lambda kind, src, p, m, b=b, k=k: f"=({mom_expr(L, b, k, kind, src, p, m) or '0'})/{int(CR)}")
        pdf.row_breaks.append(Break(id=r))
        r += 2
    setup_print(pdf, landscape=True, fit_wide=1)


def build_index(wb, source_name):
    ws = wb.create_sheet("Index", 0)
    ws["A1"], ws["A1"].font = "INDEX - CFO PACK", F_TITLE
    ws["A2"] = f"Built {dt.datetime.now():%d-%b-%Y %H:%M} from {source_name}. Click a sheet name to jump to it."
    ws["A2"].font = F_NOTE
    desc = {
        "CFO Analysis": "The figures: headline numbers, balance sheet and debt, cash story, entity view, Q2 vs Q1 movers (annexes: expense tie-out and checks, hidden).",
        "PDF - Statements": "Condensed P&L, Balance Sheet, Cash Flow and Fund Flow in Rs crore - the print / PDF page.",
        "MOM Summary": "Month on month P&L, balance sheet, cash flow, % of sales and ratios - pick the view in cell B3 (5 views).",
        "Directors Summary": "START HERE (directors). One-page executive summary; every figure links to CFO Analysis. Also printed as its own PDF.",
        "CFO Briefing": "The CFO's summary, red flags, board questions and data caveats - kept out of the PDF.",
        "Branch Transactions": "Branch sales / purchases and branch balances, per company, month on month (from the consolidation memo).",
        "Related Party": "Related-party sales / purchases and balances, per company, month on month (from the consolidation memo).",
        "Debtors & Creditors": "Sundry debtors and creditors party-wise under their Tally sub-group, at each month end (employees excluded).",
        "Loans & Advances": "Loans & advances, prepaid and deposits, at each month end (employee advances excluded).",
        "Employees": "Employee advances and amounts payable to employees (incl. expense claims, gratuity provision).",
        "Taxes": "GST, TDS / TCS and other taxes - payable and receivable, net, at each month end.",
        "Debtors Ageing": "Debtors party-wise under their group, 0-180 / over 180 days by invoice date - as on the snapshot date.",
        "Creditors Ageing": "Creditors party-wise under their group, 0-180 / over 180 days by invoice date - as on the snapshot date.",
        "Inventory": "Closing stock by category (Machines, Heads, Spare parts, Ink, Paper & packing, Others) and stock group.",
        "Inventory Ageing": "Stock item by item, FIFO on receipt date: 0-90 / 91-180 / over 180 days; filter by category.",
        "PDF - MOM": "The same month-on-month figures for all four views, one per page - the print / PDF page.",
        "MOM PL GL CFS": "CONSOLIDATED Profit & Loss, full GL detail (copied unchanged from the source workbook).",
        "MOM BS GL CFS": "CONSOLIDATED Balance Sheet, full GL detail, as at each month end.",
        "MOM Ratio Analysis CFS": "Ratio analysis on the consolidated figures.",
        "MOM PL GL SFS": "STANDALONE Profit & Loss, full GL detail.",
        "MOM BS GL SFS": "STANDALONE Balance Sheet, full GL detail.",
        "MOM Ratio Analysis SFS": "Ratio analysis on the standalone figures.",
        "MOM Cash Flow CFS": "CONSOLIDATED cash flow (indirect method), same 11 periods x 8 columns as the P&L sheet. Check row must be nil.",
        "MOM Cash Flow SFS": "STANDALONE cash flow, per company. Check row must be nil.",
        "MOM Fund Flow CFS": "CONSOLIDATED fund flow (CMA style) with the schedule of working-capital changes.",
        "MOM Fund Flow SFS": "STANDALONE fund flow, per company.",
        "Opening BS 31-Mar-26": "The 31-Mar-26 position the April / Q1 / Apr-* flows start from, its sources, and the ConnectWave vs Tally cross-check.",
    }
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 130
    r = 4
    for name in [s.title for s in wb.worksheets if s.title != "Index"]:
        c = ws.cell(r, 1, name)
        c.hyperlink = f"#'{name}'!A1"
        c.font = Font(color="0563C1", underline="single", bold=True)
        ws.cell(r, 2, desc.get(name, "")).font = F_NORM
        r += 1
    return ws


# =============================================================================================
# 7. Facts gathered from the source (values) for the narrative, movers and tie-out
# =============================================================================================
def gl_movers(pl_lay: Layout, n=10):
    """The cost and income lines that explain the Q2 margin, on the CFS Grand Total.

    A raw Q2 - Q1 ranking is useless here: Q2 sales were 35% higher, so every purchase ledger
    'moves' by volume alone. Each line is therefore compared with its Q1 amount SCALED to Q2
    sales (Q1 x Q2 turnover / Q1 turnover): the gap is what the line cost (or earned) beyond
    what sales growth explains - i.e. its effect on the margin. Turnover lines are left out (they
    are the base); Opening and Closing Stock are combined into one stock-movement line, because
    either alone is meaningless."""
    q1c, q2c = pl_lay.col.get(("Q1 Total", "GT")), pl_lay.col.get(("Q2 Total", "GT"))
    section, out, stock = None, [], []
    stop = pl_lay.row("ELIMINATED ON CONSOLIDATION") or pl_lay.max_row
    for r in range(5, stop):
        a, d = pl_lay.cell(r, 1), pl_lay.cell(r, 4)
        if isinstance(a, str) and a.strip():
            section = a.strip()
        if not isinstance(d, str) or d.startswith(("Subtotal", "TOTAL", "Total")) or not section:
            continue
        if section.startswith(("TURNOVER", "GROSS PROFIT", "EBITDA", "EBDTA", "PBT", "PAT")):
            continue
        q1, q2 = pl_lay.cell(r, q1c), pl_lay.cell(r, q2c)
        if not isinstance(q1, (int, float)) or not isinstance(q2, (int, float)):
            continue
        ref1, ref2 = f"'{pl_lay.title}'!${col_letter(q1c)}${r}", f"'{pl_lay.title}'!${col_letter(q2c)}${r}"
        sign = 1 if section.startswith("ADD: OTHER") else -1
        sec = section.replace("LESS: ", "").replace("ADD: ", "").split("(")[0].strip().title()[:24]
        if d.strip() in ("Opening Stock", "Closing Stock"):
            stock.append((q1, q2, ref1, ref2, sign, sec))
            continue
        out.append({"gl": d.strip(), "section": sec, "q1": q1, "q2": q2, "sign": sign, "q1_ref": ref1, "q2_ref": ref2})
    if stock:
        out.append({"gl": "Stock movement (opening - closing stock)", "section": stock[0][5],
                    "q1": sum(s[0] for s in stock), "q2": sum(s[1] for s in stock), "sign": -1,
                    "q1_ref": "(" + "+".join(s[2] for s in stock) + ")", "q2_ref": "(" + "+".join(s[3] for s in stock) + ")"})
    to1 = pl_lay.value("TOTAL TURNOVER", "Q1 Total", "GT") or 1.0
    to2 = pl_lay.value("TOTAL TURNOVER", "Q2 Total", "GT") or 0.0
    k = to2 / to1
    for m in out:
        m["effect"] = (m["q2"] - m["q1"] * k) * m["sign"]
    out.sort(key=lambda m: -abs(m["effect"]))
    return out[:n]


def opex_movers(pl_lay: Layout, n=6):
    """Operating-expense sub-groups ranked by their Q2 - Q1 increase (CFS Grand Total), with refs to
    their 'Subtotal:' rows so the table stays live."""
    q1c, q2c = pl_lay.col.get(("Q1 Total", "GT")), pl_lay.col.get(("Q2 Total", "GT"))
    start, end = pl_lay.row("LESS: OPERATING EXPENSES"), pl_lay.row("TOTAL OPERATING EXPENSES")
    out = []
    for r in range(start or 0, end or 0):
        d = pl_lay.cell(r, 4)
        if isinstance(d, str) and d.startswith("Subtotal: "):
            q1, q2 = pl_lay.cell(r, q1c) or 0.0, pl_lay.cell(r, q2c) or 0.0
            out.append({"name": d[len("Subtotal: "):], "q1": q1, "q2": q2, "delta": q2 - q1,
                        "q1_ref": f"'{pl_lay.title}'!${col_letter(q1c)}${r}", "q2_ref": f"'{pl_lay.title}'!${col_letter(q2c)}${r}"})
    out.sort(key=lambda m: -m["delta"])
    return out[:n]


def sales_mix(pl_lay: Layout):
    """Turnover by product, by sales-ledger name (CFS Grand Total, Q1 and Q2). Narrative only."""
    q1c, q2c = pl_lay.col.get(("Q1 Total", "GT")), pl_lay.col.get(("Q2 Total", "GT"))
    end = pl_lay.row("TOTAL TURNOVER")
    mix = collections.defaultdict(lambda: [0.0, 0.0])
    for r in range(5, end or 0):
        d = pl_lay.cell(r, 4)
        if not isinstance(d, str) or d.startswith(("Subtotal", "TOTAL")):
            continue
        n = d.upper()
        seg = ("Machines" if any(x in n for x in ("MACHINE", "HIGH SEAS", "OWNER SHIP")) else "Ink" if "INK" in n
               else "Heads" if "HEAD" in n else "Spares" if "SPARE" in n else "Other")
        mix[seg][0] += pl_lay.cell(r, q1c) or 0.0
        mix[seg][1] += pl_lay.cell(r, q2c) or 0.0
    return dict(mix)


def largest_fa_additions(n=3):
    """Biggest fixed-asset ledger increases 31-Mar -> 30-Apr from the saved Tally pulls (narrative)."""
    out = []
    for key in ENTITY_KEYS:
        a_p, b_p = GL_DATA / f"bs_q4_2526_{key}.csv", SKILL_DIR / "gl_data_mom_detail" / f"bs_apr_{key}.csv"
        if not (a_p.exists() and b_p.exists()):
            continue
        a, b = {}, {}
        for path, d in ((a_p, a), (b_p, b)):
            with open(path, newline="", encoding="utf-8") as f:
                for r in csv.DictReader(f):
                    if r["primary_group"].strip().lower() == "fixed assets":
                        d.setdefault(r["ledger_name"], float(r["closing_balance"]))
        for name in set(a) | set(b):
            add = -(b.get(name, 0.0) - a.get(name, 0.0))  # Dr is negative in the pull
            if add > 0:
                out.append((add, name, SHORT[key]))
    out.sort(reverse=True)
    return out[:n]


def expense_tieout(pl_s: Layout, tmp: Path):
    """GL file vs Expense Sheet 2026-10-10 by main group and quarter (Rs lakh). The Expense Sheet's
    'Summary' pivot holds main groups by month in Rs."""
    if not EXPENSE_SHEET.exists():
        return [], {}
    wb = openpyxl.load_workbook(readable_copy(EXPENSE_SHEET, tmp), data_only=True, read_only=True)
    rows = list(wb["Summary"].iter_rows(values_only=True))
    months = {}
    hdr = next(r for r in rows if r and "APR-26" in [str(x).strip() if x else "" for x in r])
    for i, h in enumerate(hdr):
        if isinstance(h, str) and h.strip().endswith("-26"):
            months[h.strip()] = i
    q_cols = {"Q1": ["APR-26", "MAY-26", "JUN-26"], "Q2": ["JUL-26", "AUG-26", "SEP-26"], "Sep": ["SEP-26"]}
    # the pivot repeats month labels under Q1 / Q2 - the second set of Q2 months is offset; take by position
    order = [h.strip() for h in hdr if isinstance(h, str) and h.strip().endswith("-26")]
    idx = [i for i, h in enumerate(hdr) if isinstance(h, str) and h.strip().endswith("-26")]
    pos_m = dict(zip(order, idx))
    # In the pivot, each quarter block is months + a 'Q? Total' column; values may be None.
    exp = {}
    for r in rows:
        if not r or not isinstance(r[1] if len(r) > 1 else None, str):
            if r and isinstance(r[0], str) and r[0].strip().endswith("Total"):
                key = r[0].strip()
            else:
                continue
        else:
            key = r[1].strip()
        exp[key] = {m: (r[pos_m[m]] or 0.0) if pos_m.get(m) is not None and pos_m[m] < len(r) else 0.0 for m in order}

    def esum(key, q):
        d = exp.get(key, {})
        return sum(d.get(m, 0.0) or 0.0 for m in q_cols[q])

    def gl(label_, q, minus=None):
        per = {"Q1": "Q1 Total", "Q2": "Q2 Total", "Sep": "Sep-26"}[q]
        e = ref_or_zero(pl_s.ref(label_, per, "GT"))
        return e if not minus else f"{e}-{ref_or_zero(pl_s.ref(minus, per, 'GT'))}"

    out = []
    for item, gl_label, exp_key, reason in (
            ("Custom duty", "Subtotal: CUSTOM DUTY EXPENSES", "CUSTOM DUTY EXPENSES", "Same ledgers"),
            ("Exchange gain/(loss), net", "Subtotal: EXCHANGE RATE Gain/(Loss)", "EXCHANGE RATE Gain/(Loss)", "Same ledgers"),
            ("Inventory valuation provisions", "Total Valuation Provisions", "INVENTORY VALUTION", "Provision revised after 10-Oct"),
            ("Finance cost", "TOTAL FINANCE COST", "FINANCE COST", "Postings after 10-Oct"),
    ):
        for q in ("Q1", "Q2"):
            out.append({"item": item, "period": q, "gl_expr": gl(gl_label, q), "exp": esum(exp_key, q) / 1e5, "reason": reason})
    for q in ("Q1", "Q2"):
        ind = esum("INDIRECT Total", q) - esum("FINANCE COST", q)
        out.append({"item": "Operating expenses (indirect excl. finance cost)", "period": q,
                    "gl_expr": gl("TOTAL OPERATING EXPENSES", q), "exp": ind / 1e5,
                    "reason": "Ties in Q1; Q2 gap = entries posted 10-16 Oct"})
    return out, exp


def build_briefing(v):
    """Narrative points, worded from values read at build time (each also appears as a live figure above)."""
    f = lambda x: f"Rs {x:,.2f} cr"
    p = lambda x: f"{x * 100:.1f}%"
    mix = v["mix"]
    mq1, mq2 = mix.get("Machines", [0, 0])
    om = v["opex_movers"][:2]
    fa = v["fa_adds"]
    fa_txt = (f"{fa[0][1]} ({fa[0][2]}) {f(fa[0][0] / CR)} in April" if fa else "a large April addition")
    return [
        ("A. Six headline messages", [
            f"Sales are growing but profit is not: Q2 turnover {f(v['to_q2'])} vs Q1 {f(v['to_q1'])} (+{p(v['to_q2'] / v['to_q1'] - 1)}), "
            f"yet EBITDA fell to {f(v['eb_q2'])} from {f(v['eb_q1'])} and PAT to {f(v['pat_q2'])} from {f(v['pat_q1'])}. "
            f"Gross margin {p(v['gp_q1'] / v['to_q1'])} -> {p(v['gp_q2'] / v['to_q2'])}.",
            f"Why: the growth came from MACHINES (sales {f(mq1 / CR)} in Q1 -> {f(mq2 / CR)} in Q2, by sales-ledger name), the lower-margin "
            f"line, and operating costs rose {f(v['opex_q2'] - v['opex_q1'])} - mainly {om[0]['name']} (+{f(om[0]['delta'] / CR)})"
            + (f" and {om[1]['name']} (+{f(om[1]['delta'] / CR)})" if len(om) > 1 else "") + ". September: biggest sales month "
            f"({f(v['to_sep'])}), weakest profit (EBITDA {f(v['eb_sep'])}, PAT {f(v['pat_sep'])}).",
            f"Apr-Sep group PAT {f(v['pat_h1'])} is BEFORE depreciation and tax, neither of which is booked. The estimate after both is on "
            "this sheet (Adjusted PAT) - quote that, not the unadjusted figure.",
            (f"Operations USED {f(-v['cfo_h1'])} of cash in Apr-Sep" if v["cfo_h1"] < 0 else
             f"Operations generated {f(v['cfo_h1'])} of cash in Apr-Sep") + f" against PAT of {f(v['pat_h1'])}. Debtors "
            f"{f(v['deb_open'])} at 31-Mar -> {f(v['deb_sep'])} at 30-Sep (Sep alone +{f(v['deb_sep'] - v['deb_aug'])}); "
            f"suppliers were paid down {f(v['cred_open'] - v['cred_sep'])}.",
            f"Borrowings more than doubled: gross debt {f(v['debt_open'])} -> {f(v['debt_sep'])} (Bank OD {f(v['od_open'])} -> {f(v['od_sep'])}, "
            f"term loans {f(v['sec_open'])} -> {f(v['sec_sep'])}), funding the working capital and capex of {f(-v['capex_h1'])} "
            f"(incl. {fa_txt}). Cash & bank only {f(v['cash_sep'])} at 30-Sep; interest cover {v['ebitda_h1'] / v['fin_h1']:.1f}x.",
            f"By entity (standalone): Orange O Tec PAT {f(v['pat_otec'])}, Enterprises {f(v['pat_ent'])}, Colorix {f(v['pat_col'])} "
            f"(Colorix Q2 {f(v['pat_col_q2'])}). Enterprises EBITDA fell from {f(v['eb_ent_q1'])} in Q1 to {f(v['eb_ent_q2'])} in Q2.",
        ]),
        ("B. Red flags to have answers for", [
            "Margin compression in Q2 - machine mix and the GL lines in section 5; ask sales for price / margin by segment for September.",
            f"Receivables build-up: debtor days ~{v['debtor_days']:.0f} on Apr-Sep sales. Which customers, how old, what is collectable by 31-Dec?",
            "Working capital and the land purchase are funded largely by short-term OD - headroom against sanctioned limits, the "
            "drawing-power (DP) position, and whether the plot should move to a term loan.",
            f"Commission & brokerage and advertising jumped in Q2 - are they tied to the machine orders (one-off) or a new run-rate?",
            "Monthly profit is distorted by the inventory overhead provision (booked April, reversed June) and an April-June month-allocation "
            "difference in the Surat books (cash flow line D, nil for the quarter) - present quarters, not months.",
        ]),
        ("C. Questions the board is likely to ask - suggested answers", [
            "Why is profit down when sales are up? -> The extra sales were machines at lower margin, and commission / advertising rose; "
            "sections 5 and 5b give the lines and amounts.",
            "Where did the cash go? -> Into debtors, paying suppliers and capex (incl. the plot); funded by OD and term loans (section 3).",
            "What is the real profit for the half-year? -> Adjusted PAT estimate after depreciation and tax (section 1); final once booked.",
            "Which entity is earning the money? -> Use the standalone view (section 4); consolidated entity columns move profit between companies.",
            "Do the reports agree? -> Yes: every control in the pack is nil and it matches the expense report except timing "
            "(Annex A and B of the Excel, grouped and hidden).",
        ]),
        ("D. Data caveats", [
            f"Source: {v['source']} - pulled from live Tally; not refreshed for this pack (user instruction 2026-10-24). Later postings are not in it.",
            "Opening 31-Mar-26: ConnectWave mirror for four companies, saved Tally pull (13-Oct) for Orange O Tec Surat; P&L A/c is the balancing figure.",
            "CFS entity columns are external trade only; depreciation and tax are not booked; quarter and month figures come from separate Tally pulls "
            "and are proved additive in the source's own cross-check.",
        ]),
    ]


# =============================================================================================
# 8. Excel: recalculate, save, export the PDF
# =============================================================================================
def excel_finalise(xlsx: Path, pdf: Path, print_sheets, standalone=(), short_pdf: Path | None = None):
    """Recalculate in Excel, save, export print_sheets merged into `pdf`, and each (sheet, path) in
    `standalone` as a PDF of its own (the directors' summary). `short_pdf`: the standalone PDFs plus the
    first print sheet (CFO Analysis) - the short download."""
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    from pypdf import PdfWriter
    parts = []
    try:
        wb = xl.Workbooks.Open(str(xlsx), UpdateLinks=0)
        xl.CalculateFull()
        # one PDF per print sheet, merged below - grouped-sheet export is unreliable over COM
        for i, name in enumerate(print_sheets):
            part = pdf.with_name(f"{pdf.stem}__part{i}.pdf")
            wb.Worksheets(name).ExportAsFixedFormat(0, str(part), 0, True, False)
            parts.append(part)
        for name, path in standalone:
            wb.Worksheets(name).ExportAsFixedFormat(0, str(path), 0, True, False)
        wb.Worksheets("Index").Activate()
        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        try:
            xl.Quit()
        except Exception:  # Excel already gone after a COM error - nothing to close
            pass
        pythoncom.CoUninitialize()
    writer = PdfWriter()
    for part in parts:
        writer.append(str(part))
    with open(pdf, "wb") as f:
        writer.write(f)
    if short_pdf:
        w2 = PdfWriter()
        for _name, path in standalone:
            w2.append(str(path))
        w2.append(str(parts[0]))
        with open(short_pdf, "wb") as f:
            w2.write(f)
    for part in parts:
        part.unlink()


def relabel(wb, unit):
    """The lakh / absolute versions: change the unit words on the built sheets (never the copied GL sheets)."""
    rules = {"lakh": [(r"\bRs\.? crore\b", "Rs lakh"), (r"\bcrore\b", "lakh"), (r"\bcr\b", "lakh")],
             "abs": [(r"\bRs\.? crore\b", "Rs"), (r"\bcrore\b", "rupees"), (r" ?\bcr\b", "")]}[unit]
    for ws in wb.worksheets:
        if ws.title in KEEP_SHEETS:
            continue
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and ("cr" in c.value):
                    v = c.value
                    for pat, rep in rules:
                        v = re.sub(pat, rep, v)
                    c.value = v


def make_short(full: Path, short: Path):
    """The short Excel: SHORT_SHEETS only, as values (Excel has just saved them calculated)."""
    wb = openpyxl.load_workbook(full, data_only=True)
    for ws in list(wb.worksheets):
        if ws.title not in SHORT_SHEETS:
            wb.remove(ws)
    wb.active = 0
    wb.save(short)


# =============================================================================================
# main
# =============================================================================================
def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("--source", help="PL_BS_GL_Detail workbook (default: newest)")
    ap.add_argument("--no-pdf", action="store_true", help="skip the Excel recalculation and PDF export")
    ap.add_argument("--out-dir", help=f"output folder (default {RESULTS})")
    ap.add_argument("--unit", choices=list(UNIT_DIV), default="cr", help="reporting unit: crore (default), lakh, absolute rupees")
    ap.add_argument("--stamp", help="file stamp YYYY-MM-DD_HHMM (to give the unit versions of one pack the same stamp)")
    args = ap.parse_args()
    global CR, NUM_CR
    CR = UNIT_DIV[args.unit]
    if args.unit == "abs":
        NUM_CR = '#,##0;[Red](#,##0);-'
    unit_sfx = UNIT_SUFFIX[args.unit]
    src = Path(args.source) if args.source else newest_source()
    if not src.is_absolute():
        src = PROJECT / src
    stamp = args.stamp or dt.datetime.now().strftime("%Y-%m-%d_%H%M")
    out_dir = Path(args.out_dir) if args.out_dir else RESULTS
    out_dir.mkdir(parents=True, exist_ok=True)
    out_xlsx = out_dir / f"CFO_Pack_PL_BS_CF_FF_{stamp}{unit_sfx}.xlsx"
    out_pdf = out_xlsx.with_suffix(".pdf")
    if out_xlsx.exists():
        sys.exit(f"{out_xlsx} already exists - not overwriting")
    tmp = Path(tempfile.mkdtemp(prefix="cfo_pack_"))
    say(f"Source: {src}")
    copy = readable_copy(src, tmp)

    # --- companies and opening --------------------------------------------------------------------
    say("1/6 Company pre-flight (company list vs ConnectWave) ...")
    sb = connectwave()
    plan = resolve_companies(tmp, sb)
    for k, p in plan.items():
        say(f"   {SHORT[k]:18s} FY26-27 {p['fy2627']!r}" + ("  ** SPLIT ** prior " + repr(p['fy2526']) if p["split"] else ""))

    say("2/6 Reading the source workbook ...")
    wb_v = openpyxl.load_workbook(copy, data_only=True, read_only=True)
    wb = openpyxl.load_workbook(copy)
    L = {}
    for name, key in (("MOM PL GL CFS", "PL_CFS"), ("MOM BS GL CFS", "BS_CFS_LAY"), ("MOM PL GL SFS", "PL_SFS"),
                      ("MOM BS GL SFS", "BS_SFS_LAY")):
        L[key] = Layout(wb[name], wb_v[name])
    L["BS_CFS"] = L["BS_CFS_LAY"]
    # header text per entity column (exact company names) from the PL sheet
    entity_names = dict(TOTAL_LABEL)
    for (period, key), c in L["PL_SFS"].col.items():
        if period == "Apr-26" and key not in TOTAL_LABEL:
            entity_names[key] = L["PL_SFS"].cache[4][c - 1]
    # the company list and the workbook must name the same 26-27 files
    for k in ENTITY_KEYS:
        if entity_names.get(k) != plan[k]["fy2627"]:
            sys.exit(f"STOP - workbook column for {k} is {entity_names.get(k)!r} but the company list says {plan[k]['fy2627']!r}")

    say("3/6 Opening balances 31-Mar-26 ...")
    opening = {"SFS": {}, "CFS": {}}
    ledgers_by_key, sources, compare = {}, {}, collections.OrderedDict()
    elim = eliminated_ledgers(L["BS_CFS_LAY"])
    for k in ENTITY_KEYS:
        p = plan[k]
        csv_led, csv_time = (opening_from_tally_csv(k) if (GL_DATA / f"bs_q4_2526_{k}.csv").exists() else ({}, None))
        if k in TALLY_CSV_KEYS and not p["split"]:
            led = csv_led
            sources[k] = f"Tally pull {csv_time:%d-%b-%Y} (bs_q4_2526)"
            cw_led = {}
        else:
            try:
                led = opening_from_connectwave(sb, p["tenant"])
            except Exception as exc:  # the mirror's statement timeout on big books
                sys.exit(f"STOP - ConnectWave could not serve {p['fy2627']!r}: {exc}")
            sources[k] = "ConnectWave opening" + (" (FY26-27 file after split)" if p["split"] else "")
            cw_led = led
        ledgers_by_key[k] = led
        # SFS group totals in display sign (liabilities +, assets +)
        sfs = collections.defaultdict(float)
        cfs = collections.defaultdict(float)
        for name, (g, sg, bal) in led.items():
            if g == PL_LINE and name.strip() != "Profit & Loss A/c" and abs(bal) >= 1:
                say(f"   ! {SHORT[k]}: {name!r} ({bal:,.0f}) resolved to the P&L line - it is absorbed by the "
                    "balancing figure; check its group in Tally")
            if g == "STOCK" or g == PL_LINE:
                continue
            if g.startswith("?"):
                say(f"   ! {SHORT[k]}: ledger {name!r} has an unmapped group {g[1:]!r} - carried as Current Assets/Liabilities by sign")
                g = "Current Liabilities" if bal > 0 else "Current Assets"
            disp = bal if g in LIAB_GROUPS else -bal
            sfs[g] += disp
            if not is_intercompany(g, sg, name, elim.get(k, set())):
                cfs[g] += disp
        opening["SFS"][k], opening["CFS"][k] = dict(sfs), dict(cfs)
        # cross-check vs the saved Tally pull (by group, Tally sign)
        if cw_led and csv_led:
            tg_cw, tg_t = group_totals(cw_led), group_totals(csv_led)
            rowsets = []
            for g in sorted(set(tg_cw) | set(tg_t), key=lambda x: (x in (PL_LINE, "STOCK"), x)):
                a, b = tg_cw.get(g, 0.0), tg_t.get(g, 0.0)
                if g == PL_LINE:
                    comment = "Expected to differ: ConnectWave carries FY 25-26 profit into the opening; the TB shows the 1-Apr-25 b/f only."
                elif abs(a - b) < 1:
                    comment = "Ties"
                else:
                    diffs = sorted(((n, cw_led.get(n, (g, '', 0))[2] - csv_led.get(n, (g, '', 0))[2])
                                    for n in set(cw_led) | set(csv_led)
                                    if (cw_led.get(n, (None,))[0] == g or csv_led.get(n, (None,))[0] == g)),
                                   key=lambda x: -abs(x[1]))
                    top = "; ".join(f"{n} {d:,.0f}" for n, d in diffs[:3] if abs(d) >= 1)
                    comment = ("Carry-forward difference (prior-year file closing vs FY26-27 opening): " if p["split"] or k.startswith("ENTER")
                               else "Mirror vs Tally pull: ") + top
                rowsets.append((g, a, b, comment))
            compare[k] = rowsets
        elif p["split"] and p["tenant_prior"]:
            prior = closing_from_connectwave(sb, p["tenant_prior"])
            tg_cw, tg_t = group_totals(led), group_totals(prior)
            compare[k] = [(g, tg_cw.get(g, 0.0), tg_t.get(g, 0.0), "FY26-27 opening vs FY25-26 file closing")
                          for g in sorted(set(tg_cw) | set(tg_t))]
        say(f"   {SHORT[k]:18s} {len(led):5d} ledgers  source: {sources[k]}")

    # stock at 31-Mar-26 = April 'Opening Stock' on the SFS P&L (display sign: positive)
    stock_refs = {k: L["PL_SFS"].ref("Opening Stock", "Apr-26", k) for k in ENTITY_KEYS}
    if any(v is None for v in stock_refs.values()):
        sys.exit("STOP - could not find the April 'Opening Stock' row on MOM PL GL SFS")

    # --- workbook --------------------------------------------------------------------------------
    say("4/6 Building sheets ...")
    for ws in list(wb.worksheets):
        if ws.title not in KEEP_SHEETS:
            wb.remove(ws)
    _op_ws, open_rows = build_opening(wb, L["BS_SFS_LAY"], L["BS_CFS_LAY"], entity_names, opening,
                                      sources, compare, stock_refs)
    cf = {}
    ff = {}
    for basis in ("CFS", "SFS"):
        _ws, rows, pos = build_cash_flow(wb, basis, L[f"BS_{basis}_LAY"], L[f"PL_{basis}"], open_rows, entity_names)
        cf[basis] = {"rows": rows, "pos": pos}
        L["CF_" + basis] = cf[basis]
    for basis in ("CFS", "SFS"):
        _ws, rows, pos = build_fund_flow(wb, basis, L[f"BS_{basis}_LAY"], L[f"PL_{basis}"], open_rows, entity_names)
        ff[basis] = {"rows": rows, "pos": pos}
        L["FF_" + basis] = ff[basis]

    # values for the narrative (read from the source's cached values - the six sheets are unchanged)
    pc, ps, bc = L["PL_CFS"], L["PL_SFS"], L["BS_CFS_LAY"]
    vcr = lambda lay, lab, per, k="GT": (lay.value(lab, per, k) or 0.0) / CR
    bsg = bs_group_rows(bc)

    def bsv(group, month):
        r = bsg[group][0] if group in bsg else None
        c = bc.col.get((month, "GT"))
        return (bc.cache[r][c - 1] or 0.0) / CR if r and c else 0.0

    def openv(group):
        return sum(opening["CFS"][k].get(group, 0.0) for k in ENTITY_KEYS) / CR

    pl_bs_m = {m: bsv(PL_LINE, m) for m in MONTHS}
    pat_m = {m: vcr(pc, "PAT (NET PROFIT)", m) for m in MONTHS}
    v = {
        "to_q1": vcr(pc, "TOTAL TURNOVER", "Q1 Total"), "to_q2": vcr(pc, "TOTAL TURNOVER", "Q2 Total"),
        "gp_q1": vcr(pc, "GROSS PROFIT", "Q1 Total"), "gp_q2": vcr(pc, "GROSS PROFIT", "Q2 Total"),
        "eb_q1": vcr(pc, "EBITDA", "Q1 Total"), "eb_q2": vcr(pc, "EBITDA", "Q2 Total"),
        "pat_q1": vcr(pc, "PAT (NET PROFIT)", "Q1 Total"), "pat_q2": vcr(pc, "PAT (NET PROFIT)", "Q2 Total"),
        "to_sep": vcr(pc, "TOTAL TURNOVER", "Sep-26"), "eb_sep": vcr(pc, "EBITDA", "Sep-26"),
        "pat_sep": vcr(pc, "PAT (NET PROFIT)", "Sep-26"), "opex_sep": vcr(pc, "TOTAL OPERATING EXPENSES", "Sep-26"),
        "opex_apr": vcr(pc, "TOTAL OPERATING EXPENSES", "Apr-26"), "pat_h1": vcr(pc, "PAT (NET PROFIT)", "Apr-Sep Total"),
        "ebitda_h1": vcr(pc, "EBITDA", "Apr-Sep Total"), "fin_h1": vcr(pc, "TOTAL FINANCE COST", "Apr-Sep Total"),
        "deb_open": openv("Sundry Debtors"), "deb_sep": bsv("Sundry Debtors", "Sep-26"), "deb_aug": bsv("Sundry Debtors", "Aug-26"),
        "cred_open": openv("Sundry Creditors"), "cred_sep": bsv("Sundry Creditors", "Sep-26"),
        "od_open": openv("Bank OD A/c"), "od_sep": bsv("Bank OD A/c", "Sep-26"),
        "sec_open": openv("Secured Loans"), "sec_sep": bsv("Secured Loans", "Sep-26"),
        "cash_sep": bsv("Bank Accounts", "Sep-26") + bsv("Cash-in-hand", "Sep-26"),
        "pat_otec": vcr(ps, "PAT (NET PROFIT)", "Apr-Sep Total", "TOT_OTEC"), "pat_ent": vcr(ps, "PAT (NET PROFIT)", "Apr-Sep Total", "TOT_ENT"),
        "pat_col": vcr(ps, "PAT (NET PROFIT)", "Apr-Sep Total", "COLORIX"), "pat_col_q2": vcr(ps, "PAT (NET PROFIT)", "Q2 Total", "COLORIX"),
        "eb_ent_q1": vcr(ps, "EBITDA", "Q1 Total", "TOT_ENT"), "eb_ent_q2": vcr(ps, "EBITDA", "Q2 Total", "TOT_ENT"),
        "opex_q1": vcr(pc, "TOTAL OPERATING EXPENSES", "Q1 Total"), "opex_q2": vcr(pc, "TOTAL OPERATING EXPENSES", "Q2 Total"),
        "source": src.name,
    }
    v["debtor_days"] = v["deb_sep"] / vcr(pc, "TOTAL TURNOVER", "Apr-Sep Total") * 183
    v["debt_open"] = v["od_open"] + v["sec_open"] + openv("Loans (Liability)")
    v["debt_sep"] = v["od_sep"] + v["sec_sep"] + bsv("Loans (Liability)", "Sep-26")
    v["mix"] = sales_mix(pc)
    v["opex_movers"] = opex_movers(pc)
    v["fa_adds"] = largest_fa_additions()
    # operating cash and capex for the narrative: same arithmetic as the CFS cash flow sheet, from values
    open_cfs_gt = lambda g: sum(opening["CFS"][k].get(g, 0.0) for k in ENTITY_KEYS) / CR
    v["capex_h1"] = -(bsv("Fixed Assets", "Sep-26") - open_cfs_gt("Fixed Assets"))
    d_cfs = lambda g: bsv(g, "Sep-26") - open_cfs_gt(g)
    cash_open = open_cfs_gt("Bank Accounts") + open_cfs_gt("Cash-in-hand")
    net_b = -(d_cfs("Fixed Assets") + d_cfs("Investments") + d_cfs("Non-Current Investments"))
    net_c = (sum(d_cfs(g) for g in ("Capital Account", "Reserves & Surplus", "Secured Loans", "Unsecured Loans",
                                    "Loans (Liability)", "Bank OD A/c")) - v["fin_h1"])
    # D and E are nil on the CFS Grand Total for Apr-Sep (proved by the sheet's own lines)
    v["cfo_h1"] = (v["cash_sep"] - cash_open) - net_b - net_c
    movers = gl_movers(pc)
    tieout, _exp = expense_tieout(ps, tmp)

    # checks: (name, expression)
    cfp, ffp = cf["CFS"]["pos"], ff["CFS"]["pos"]
    checks = []
    for basis in ("CFS", "SFS"):
        rows, pos = cf[basis]["rows"], cf[basis]["pos"]
        cols = [pos[(p[0], k)] for p in PERIODS for k in (BLOCK8 if basis == "SFS" else ["GT"])]
        checks.append((f"Cash Flow {basis}: closing cash vs Balance Sheet, sum of |differences| over every period"
                       + (" and company" if basis == "SFS" else " (Grand Total)"),
                       "+".join(f"ABS('MOM Cash Flow {basis}'!{col_letter(c)}{rows['chk']})" for c in cols)))
        rows, pos = ff[basis]["rows"], ff[basis]["pos"]
        cols = [pos[(p[0], k)] for p in PERIODS for k in (BLOCK8 if basis == "SFS" else ["GT"])]
        checks.append((f"Fund Flow {basis}: working-capital schedule vs sources less applications",
                       "+".join(f"ABS('MOM Fund Flow {basis}'!{col_letter(c)}{rows['chk']})" for c in cols)))
    # additivity on the CFS Grand Total cash flow: Apr+May+Jun = Q1, Q1+Q2 = Apr-Sep (net change line)
    rows = cf["CFS"]["rows"]
    g = lambda per: f"'MOM Cash Flow CFS'!{col_letter(cfp[(per, 'GT')])}{rows['net']}"
    checks.append(("Cash Flow CFS additivity: Apr + May + Jun - Q1 (net change)", f"{g('Apr-26')}+{g('May-26')}+{g('Jun-26')}-{g('Q1 Total')}"))
    checks.append(("Cash Flow CFS additivity: Q1 + Q2 - Apr-Sep (net change)", f"{g('Q1 Total')}+{g('Q2 Total')}-{g('Apr-Sep Total')}"))
    ga = lambda per, key: f"'MOM Cash Flow CFS'!{col_letter(cfp[(per, 'GT')])}{rows[key]}"
    checks.append(("Cash Flow CFS additivity: Jul + Aug + Sep - Q2 (operating cash)", f"{ga('Jul-26', 'netA')}+{ga('Aug-26', 'netA')}+{ga('Sep-26', 'netA')}-{ga('Q2 Total', 'netA')}"))
    rows_s, pos_s = cf["SFS"]["rows"], cf["SFS"]["pos"]
    checks.append(("Cash Flow SFS line D for Q1, every company (nil = the 31-Mar opening ties to the books)",
                   "+".join(f"ABS('MOM Cash Flow SFS'!{col_letter(pos_s[('Q1 Total', k)])}{rows_s['D']})" for k in ENTITY_KEYS)))
    sfs_gap = open_rows[("SFS", "GAP")]
    checks.append(("Opening BS SFS: reconciliation gap, all companies", "+".join(
        f"ABS('Opening BS 31-Mar-26'!{col_letter(5 + i)}{sfs_gap})" for i in range(8))))
    checks.append(("Source workbook: MOM BS GL SFS reconciliation gap at 30-Sep (Grand Total)",
                   ref_or_zero(L["BS_SFS_LAY"].ref("RECONCILIATION GAP", "Sep-26", "GT"))))
    facts = {"source_name": src.name, "source_stamp": src.stem.split("_")[-2] + " " + src.stem.split("_")[-1],
             "open_rows": open_rows, "checks": checks, "briefing": build_briefing(v), "opex_movers": v["opex_movers"],
             "dep_half_year_cr": 153.16339 / 100 / 2}  # FY 25-26 CFS depreciation Rs 153.16 lakh (Quarterly file, Q4)

    import cfo_pack_extras as X
    _ca, marks = build_cfo_analysis(wb, L, cf["CFS"]["rows"], cfp, ff["CFS"]["rows"], ffp, facts, tieout, movers)
    build_print_statements(wb, L, open_rows)
    build_mom_views(wb, L)
    note = f"Source: {src.name} (live Tally pull). Depreciation and tax are not booked for Apr-Sep."
    X.build_cfo_briefing(wb, facts["briefing"], note)
    # Ageing and inventory, as on the newest ConnectWave snapshot (fetch_ageing_snapshot.py).
    import ageing_pack as A
    snap, snap_name = A.load_snapshot(RESULTS)
    ageing_hub, ageing_marks = None, None
    if snap:
        submap = {gl: sub for _g, sub, gl, _r in X.bs_ledgers(L["BS_SFS_LAY"])}
        ageing_hub, ageing_marks = A.build_ageing_sheets(wb, snap, snap_name, submap,
                                                         X.EMP_PAYABLE_SUBS | X.EMP_ADVANCE_SUBS)
    else:
        say("   (no Ageing_Snapshot_*.json - run fetch_ageing_snapshot.py for the ageing / inventory sheets)")
    X.build_directors_summary(wb, marks, note, ageing_marks, ageing_hub and ageing_hub["as_of"])
    memo_hub = X.build_memo_sheets(wb, L, entity_names)
    memo_hub.update(X.build_schedule_sheets(wb, L, entity_names))
    build_index(wb, src.name)
    order = ["Index", "Directors Summary", "CFO Briefing", "MOM Summary", "CFO Analysis", "PDF - Statements", "PDF - MOM",
             "Branch Transactions", "Related Party"] + X.SCHEDULE_SHEETS + (
                 ["Debtors Ageing", "Creditors Ageing", "Inventory", "Inventory Ageing"] if snap else []) + KEEP_SHEETS + [
        "MOM Cash Flow CFS", "MOM Cash Flow SFS", "MOM Fund Flow CFS", "MOM Fund Flow SFS", "Opening BS 31-Mar-26"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == "Index"

    if args.unit != "cr":
        relabel(wb, args.unit)
    say(f"5/6 Saving {out_xlsx.name} ...")
    wb.save(out_xlsx)
    # GL drill-down for the Hub page (publish_cfo_pack_hub.py), from the same source values.
    if args.unit == "cr":
        out_xlsx.with_name(out_xlsx.stem + "_detail.json").write_text(
            json.dumps({"mom_detail": mom_detail_values(L), "memo": memo_hub, "ageing": ageing_hub}), encoding="utf-8")
    out_dir_pdf = out_xlsx.with_name(f"Directors_Executive_Summary_{stamp}{unit_sfx}.pdf")
    short_xlsx = out_xlsx.with_name(f"CFO_Pack_PL_BS_CF_FF_{stamp}{unit_sfx}_short.xlsx")
    if not args.no_pdf:
        say("6/6 Excel: full recalculation, save, PDF export ...")
        # Excel automation sometimes refuses a call ("Call was rejected by callee") - try again, fresh.
        for attempt in range(1, 4):
            try:
                excel_finalise(out_xlsx, out_pdf, ["CFO Analysis", "PDF - Statements", "PDF - MOM"],
                               standalone=[("Directors Summary", out_dir_pdf)],
                               short_pdf=short_xlsx.with_suffix(".pdf"))
                break
            except Exception as e:  # pywintypes.com_error and friends
                for part in out_dir.glob(f"{out_pdf.stem}__part*.pdf"):
                    part.unlink()
                if attempt == 3:
                    raise
                say(f"   Excel refused ({e}); retry {attempt + 1}/3 ...")
                time.sleep(10)
        make_short(out_xlsx, short_xlsx)
        say(f"   Short: {short_xlsx.name} + .pdf")
        say(f"   PDF: {out_pdf}")
        say(f"   Directors PDF: {out_dir_pdf}")
    say(f"Done: {out_xlsx}")
    print(f"OUTPUT_XLSX={out_xlsx}")
    print(f"OUTPUT_PDF={out_pdf if not args.no_pdf else ''}")


if __name__ == "__main__":
    main()
