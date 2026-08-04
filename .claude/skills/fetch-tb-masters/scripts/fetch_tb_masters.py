"""
Fetch Chart of Accounts masters from Tally for all open companies.
Outputs Files/TB_Masters_Report.xlsx with 18 columns (A-R) per ledger,
plus a 'Name Variations' sheet for GL codes with dot/space name differences.

Columns A-O mirror TB FROMAT.xlsx structure.
Column P  = Created By       (Tally CREATEDBY field)
Column Q  = Last Modified By  (Tally ALTEREDBY field)
Column R  = Name Variation Flag

Usage:
    python fetch_tb_masters.py
    python fetch_tb_masters.py --output Files/TB_Masters_Report.xlsx
    python fetch_tb_masters.py --host 192.168.33.254 --port 9000
"""

import argparse
import os
import re
import sys
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Load .env: check CWD first (project root), then script dir, then parent
_env_candidates = [
    Path.cwd() / ".env",
    Path(__file__).parent / ".env",
    Path(__file__).parent.parent / ".env",
]
for _env_path in _env_candidates:
    if _env_path.exists():
        for _line in _env_path.read_text().splitlines():
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())
        break

TALLY_HOST = os.environ.get("TALLY_HOST", "localhost")
TALLY_PORT = int(os.environ.get("TALLY_PORT", "9000"))

# ---------------------------------------------------------------------------
# Schedule map: primary_group -> (Main FS, Catgory, BS/PL, Sch Nos, Sub Sch, Cat)
# Seeded from TB FROMAT.xlsx (SURAT O TEC data). These are defaults per
# primary group. Ledgers in complex groups (CURRENT ASSETS / LIABILITIES)
# may have multiple valid mappings — they are flagged "REVIEW" in Sch Nos.
# ---------------------------------------------------------------------------
# Keys match Tally's actual primary group names (Title Case as returned by TallyPrime).
# Each tuple: (Main FS, Catgory, BS/PL, Sch Nos, Sub Sch, Cat)
# Lookup is case-insensitive via _lookup_schedule().
SCHEDULE_MAP = {
    "Capital Account":           ("Share Capital",             "", "BS", "1",   "1",  "Liab"),
    "Profit & Loss A/c":         ("PROFIT & LOSS A/C",         "", "BS", "2",   "2",  "Liab"),
    "Profit & Loss A/C":         ("PROFIT & LOSS A/C",         "", "BS", "2",   "2",  "Liab"),
    "Loans (Liability)":         ("Long Term Liabilities",     "", "BS", "3",   "3",  "Liab"),
    "Current Liabilities":       ("Other Current Liabilities", "", "BS", "8",   "8",  "Liab"),
    "Fixed Assets":              ("Fixed Assets",              "", "BS", "10",  "10", "Asset"),
    "Investments":               ("Investments",               "", "BS", "11",  "11", "Asset"),
    "Stock-in-Hand":             ("Inventory",                 "", "BS", "13",  "13", "Asset"),
    "Current Assets":            ("Other Current Asset",       "", "BS", "18",  "18", "Asset"),
    "Misc. Expenses (ASSET)":    ("Other Current Asset",       "", "BS", "18",  "18", "Asset"),
    "Branch / Divisions":        ("Branch / Divisions",        "", "BS", "19",  "19", "Asset"),
    "Suspense A/c":              ("Branch / Divisions",        "", "BS", "19",  "19", "Asset"),
    "Sales Accounts":            ("Sales Accounts",            "", "PL", "21",  "21", "Sales"),
    "Parts Sales Accounts":      ("Sales Accounts",            "", "PL", "21",  "21", "Sales"),
    "Branch Purchase & Sales":   ("Branch Sales",              "", "PL", "25",  "25", "Sales"),
    "BRANCH PURCHASE & SALES":   ("Branch Sales",              "", "PL", "25",  "25", "Sales"),
    "Direct Incomes":            ("Servcie Income",            "", "PL", "22",  "22", "Sales"),
    "Purchase Accounts":         ("Cost Of Materail",          "", "PL", "23",  "23", "COM"),
    "Direct Expenses":           ("Direct EXP",                "", "PL", "24",  "24", "COM"),
    "Indirect Incomes":          ("Other Income",              "", "PL", "22",  "22", "Sales"),
    "Indirect Expenses":         ("Other Expenses",            "", "PL", "40",  "40", "EXP"),
}


# ---------------------------------------------------------------------------
# Tally connection helpers
# ---------------------------------------------------------------------------

def post_to_tally(xml_body: str, host: str = TALLY_HOST, port: int = TALLY_PORT) -> str:
    url = f"http://{host}:{port}"
    encoded = xml_body.encode("utf-8")
    req = urllib.request.Request(
        url,
        data=encoded,
        headers={"Content-Type": "text/xml; charset=utf-8"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read().decode("utf-8", errors="replace")
    except urllib.error.URLError as e:
        raise RuntimeError(
            f"Cannot connect to Tally at {url}.\n"
            f"  Cause: {e}\n"
            f"  Make sure Tally is running, the HTTP XML server is enabled,\n"
            f"  and at least one company is open."
        ) from e


def _clean_xml(s: str) -> str:
    """
    Tally sometimes emits numeric character references to code points that are
    invalid in XML 1.0 (e.g. &#x1F; &#x3;). Strip them plus bare control chars
    so ElementTree can parse the response without error.
    """
    def _filter_ref(m):
        raw = m.group(0)
        try:
            code = int(m.group(1), 16) if raw.startswith("&#x") else int(m.group(1))
            if code <= 0x08 or code in (0x0B, 0x0C) or (0x0E <= code <= 0x1F):
                return ""
        except ValueError:
            pass
        return raw

    s = re.sub(r'&#x([0-9A-Fa-f]+);', _filter_ref, s)
    s = re.sub(r'&#(\d+);', _filter_ref, s)
    s = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', s)
    return s


def _safe_parse(resp: str):
    """Parse XML response after stripping Tally's invalid character references."""
    try:
        return ET.fromstring(resp)
    except ET.ParseError:
        return ET.fromstring(_clean_xml(resp))


def _get_text(element, tag: str) -> str:
    el = element.find(tag)
    if el is not None and el.text:
        return el.text.strip()
    return ""


# ---------------------------------------------------------------------------
# Fetch open companies
# ---------------------------------------------------------------------------

def get_open_companies(host: str = TALLY_HOST, port: int = TALLY_PORT) -> list:
    xml = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>AllCompanies</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION ISMODIFY="No" NAME="AllCompanies">
            <TYPE>Company</TYPE>
            <FETCH>Name,StartingFrom,EndingAt</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""
    resp = post_to_tally(xml, host, port)
    companies = []
    try:
        root = _safe_parse(resp)
        for elem in root.iter("COMPANY"):
            name = _get_text(elem, "NAME")
            if name and name not in companies:
                companies.append(name)
    except ET.ParseError:
        pass

    if not companies:
        # Fallback regex parse
        found = re.findall(r'<NAME[^>]*>(.*?)</NAME>', resp, re.DOTALL | re.IGNORECASE)
        for m in found:
            name = m.strip()
            if name and name not in companies:
                companies.append(name)

    if not companies:
        print(f"WARNING: No companies found. Raw response (first 600 chars):\n{resp[:600]}",
              file=sys.stderr)
    return companies


# ---------------------------------------------------------------------------
# Fetch group masters and build hierarchy
# ---------------------------------------------------------------------------

def fetch_group_masters(company: str, host: str = TALLY_HOST, port: int = TALLY_PORT) -> dict:
    """Returns {GROUP_NAME_UPPER: {'name': original_str, 'parent': str}}."""
    xml = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>GroupMasters</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION ISMODIFY="No" NAME="GroupMasters">
            <TYPE>Group</TYPE>
            <FETCH>Name,Parent</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""
    resp = post_to_tally(xml, host, port)
    groups = {}
    try:
        root = _safe_parse(resp)
        for elem in root.iter("GROUP"):
            # Tally returns name as an XML attribute, not a child element
            name = elem.get("NAME", "").strip()
            parent = _get_text(elem, "PARENT")
            if name:
                groups[name.upper()] = {"name": name, "parent": parent}
    except ET.ParseError as e:
        print(f"  XML parse error (groups) for {company}: {e}", file=sys.stderr)
    return groups


def get_group_path(group_name: str, group_map: dict) -> list:
    """
    Walk up the parent chain from group_name to the Tally primary group.
    Stops before the synthetic 'Primary' root that TallyPrime uses.
    Returns list ordered from primary_group -> group_name.
    E.g. ['Current Liabilities', 'Sundry Creditors', 'CREDITOR FOR GOODS']
    """
    path = []
    visited = set()
    current = (group_name or "").strip()

    while current:
        key = current.upper()
        if key in visited or key == "PRIMARY":
            break
        visited.add(key)
        path.insert(0, current)

        entry = group_map.get(key)
        if not entry or not entry["parent"] or entry["parent"].upper() == "PRIMARY":
            break
        current = entry["parent"]

    return path


# ---------------------------------------------------------------------------
# Fetch ledger masters (with audit trail)
# ---------------------------------------------------------------------------

def fetch_ledger_masters(company: str, host: str = TALLY_HOST, port: int = TALLY_PORT) -> list:
    """
    Returns list of dicts:
        name, parent, created_by, altered_by
    CREATEDBY / ALTEREDBY are available when Tally's audit feature is active;
    older masters may return blank values — that is expected, not an error.
    """
    xml = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>LedgerMasters</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION ISMODIFY="No" NAME="LedgerMasters">
            <TYPE>Ledger</TYPE>
            <FETCH>Name,Parent,CreatedBy,AlteredBy</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""
    resp = post_to_tally(xml, host, port)
    ledgers = []
    try:
        root = _safe_parse(resp)
        for elem in root.iter("LEDGER"):
            # Tally returns name as an XML attribute, not a child element
            name = elem.get("NAME", "").strip()
            if not name:
                continue
            ledgers.append({
                "name":       name,
                "parent":     _get_text(elem, "PARENT"),
                "created_by": _get_text(elem, "CREATEDBY"),
                "altered_by": _get_text(elem, "ALTEREDBY"),
            })
    except ET.ParseError as e:
        print(f"  XML parse error (ledgers) for {company}: {e}", file=sys.stderr)
    return ledgers


# ---------------------------------------------------------------------------
# Build output rows per company
# ---------------------------------------------------------------------------

# Pre-build a lowercase index for case-insensitive lookups
_SCHEDULE_MAP_LOWER = {k.lower(): v for k, v in SCHEDULE_MAP.items()}


def _lookup_schedule(primary_group: str):
    """Case-insensitive lookup in SCHEDULE_MAP."""
    return _SCHEDULE_MAP_LOWER.get(primary_group.lower())


def build_rows(company: str, ledgers: list, group_map: dict) -> list:
    """One output row dict per ledger, with all 18 column values."""
    rows = []
    for led in ledgers:
        parent_group = led["parent"]

        # Special case: a few Tally system ledgers (e.g. 'PROFIT & LOSS A/C')
        # sit directly under the synthetic root 'Primary' with no real parent group.
        if not parent_group or parent_group.upper() == "PRIMARY":
            primary_group = "Profit & Loss A/c"
            sub_levels = []
            parent_group = "Profit & Loss A/c"
        else:
            path = get_group_path(parent_group, group_map)
            primary_group = path[0] if path else parent_group
            sub_levels    = path[1:] if len(path) > 1 else []

        sched = _lookup_schedule(primary_group)
        if sched:
            main_fs, catgory, bspl, sch_nos, sub_sch, cat = sched
        else:
            main_fs, catgory, bspl, sch_nos, sub_sch, cat = ("UNMAPPED", "", "", "REVIEW", "", "")

        rows.append({
            "company":        company,
            "com":            "",
            "main_nature":    primary_group,
            "b":              sub_levels[0] if len(sub_levels) > 0 else "",
            "c":              sub_levels[1] if len(sub_levels) > 1 else "",
            "d":              sub_levels[2] if len(sub_levels) > 2 else "",
            "e":              sub_levels[3] if len(sub_levels) > 3 else "",
            "particulars":    led["name"],
            "main_fs":        main_fs,
            "sub_group":      parent_group,
            "catgory":        catgory,
            "bspl":           bspl,
            "sch_nos":        sch_nos,
            "sub_sch":        sub_sch,
            "cat":            cat,
            "created_by":     led["created_by"],
            "altered_by":     led["altered_by"],
            "variation_flag": "",
            # internal fields used by variation detection (not written to Excel)
            "_norm_name":     _normalize(led["name"]),
        })
    return rows


# ---------------------------------------------------------------------------
# GL name variation detection (Component B)
# ---------------------------------------------------------------------------

def _normalize(name: str) -> str:
    """Normalize ledger name: uppercase, dots→space, collapse spaces."""
    s = name.upper()
    s = re.sub(r'\.', ' ', s)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()


def detect_name_variations(all_rows: list) -> list:
    """
    Finds ledgers whose normalized names match but differ in exact spelling OR
    parent group. Fills 'variation_flag' on affected rows in-place.
    Returns list of variation groups for the separate Excel sheet.
    """
    # normalized_name -> [(company, exact_name, sub_group, row_index), ...]
    norm_map = defaultdict(list)
    for i, row in enumerate(all_rows):
        norm_map[row["_norm_name"]].append({
            "company":    row["company"],
            "exact_name": row["particulars"],
            "sub_group":  row["sub_group"],
            "idx":        i,
        })

    variation_groups = []
    for norm_name, entries in norm_map.items():
        if len(entries) < 2:
            continue
        exact_names  = {e["exact_name"] for e in entries}
        parent_groups = {e["sub_group"] for e in entries}

        # Only flag if names differ in spelling OR land in different groups
        if len(exact_names) == 1 and len(parent_groups) == 1:
            continue

        for entry in entries:
            others = [e for e in entries if e["idx"] != entry["idx"]]
            desc = "; ".join(
                f"'{e['exact_name']}' [{e['sub_group']}] ({e['company']})"
                for e in others
            )
            all_rows[entry["idx"]]["variation_flag"] = f"VARIATION: also as {desc}"

        variation_groups.append({"norm_name": norm_name, "entries": entries})

    return variation_groups


# ---------------------------------------------------------------------------
# Excel writing
# ---------------------------------------------------------------------------

HEADERS = [
    "COMPANY", "COM", "MAIN NATURE", "B", "C", "D", "E",
    "PARTICULARS NAME", "Main FS", "Sub Group", "Catgory",
    "BS/PL", "Sch Nos", "Sub Sch", "Cat",
    "Created By", "Last Modified By", "Name Variation Flag",
]

COL_KEYS = [
    "company", "com", "main_nature", "b", "c", "d", "e",
    "particulars", "main_fs", "sub_group", "catgory",
    "bspl", "sch_nos", "sub_sch", "cat",
    "created_by", "altered_by", "variation_flag",
]

COL_WIDTHS = [30, 8, 25, 22, 22, 22, 22, 48, 24, 26, 15, 7, 9, 9, 9, 18, 18, 55]

_BOLD   = Font(bold=True,  name="Calibri", size=10)
_NORMAL = Font(bold=False, name="Calibri", size=10)
_RED    = Font(bold=False, name="Calibri", size=10, color="FF0000")
_UNMAPPED_FILL = PatternFill("solid", fgColor="FFF2CC")  # light yellow highlight


def write_excel(all_rows: list, variations: list, output_path: Path):
    wb = openpyxl.Workbook()

    # ── Masters sheet ──────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Masters"

    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = _BOLD

    for col_idx, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    for row_idx, row in enumerate(all_rows, start=2):
        unmapped = row.get("sch_nos") == "REVIEW"
        for col_idx, key in enumerate(COL_KEYS, start=1):
            val  = row.get(key, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            if key == "variation_flag" and val:
                cell.font = _RED
            else:
                cell.font = _NORMAL
            if unmapped:
                cell.fill = _UNMAPPED_FILL

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"

    # ── Name Variations sheet ──────────────────────────────────────────────
    if variations:
        ws2 = wb.create_sheet("Name Variations")
        var_hdrs = ["Normalized Name", "Company", "Exact Ledger Name", "Parent Group"]
        for col_idx, h in enumerate(var_hdrs, start=1):
            cell = ws2.cell(row=1, column=col_idx, value=h)
            cell.font = _BOLD

        for col_idx, w in enumerate([42, 30, 48, 30], start=1):
            ws2.column_dimensions[get_column_letter(col_idx)].width = w

        r = 2
        for var in sorted(variations, key=lambda x: x["norm_name"]):
            for entry in var["entries"]:
                ws2.cell(row=r, column=1, value=var["norm_name"]).font = _NORMAL
                ws2.cell(row=r, column=2, value=entry["company"]).font = _NORMAL
                ws2.cell(row=r, column=3, value=entry["exact_name"]).font = _NORMAL
                ws2.cell(row=r, column=4, value=entry["sub_group"]).font = _NORMAL
                r += 1
            r += 1  # blank separator between variation groups

        ws2.freeze_panes = "A2"

    # ── Summary sheet ──────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Summary")
    # Count by company
    by_company = defaultdict(int)
    unmapped_by_company = defaultdict(int)
    for row in all_rows:
        by_company[row["company"]] += 1
        if row.get("sch_nos") == "REVIEW":
            unmapped_by_company[row["company"]] += 1

    ws3.column_dimensions["A"].width = 35
    ws3.column_dimensions["B"].width = 14
    ws3.column_dimensions["C"].width = 18
    ws3.column_dimensions["D"].width = 18

    for col_idx, h in enumerate(["Company", "Ledger Count", "Unmapped (REVIEW)", "Variations Flagged"], start=1):
        cell = ws3.cell(row=1, column=col_idx, value=h)
        cell.font = _BOLD

    flagged_by_company = defaultdict(int)
    for row in all_rows:
        if row["variation_flag"]:
            flagged_by_company[row["company"]] += 1

    for r_idx, company in enumerate(sorted(by_company.keys()), start=2):
        ws3.cell(row=r_idx, column=1, value=company).font = _NORMAL
        ws3.cell(row=r_idx, column=2, value=by_company[company]).font = _NORMAL
        ws3.cell(row=r_idx, column=3, value=unmapped_by_company[company]).font = _NORMAL
        ws3.cell(row=r_idx, column=4, value=flagged_by_company[company]).font = _NORMAL

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))


# ---------------------------------------------------------------------------
# Period-based company filtering (TALLY Company_List.xlsx)
# ---------------------------------------------------------------------------

_PERIOD_COL_MAP = {"2627": "26-27", "2526": "25-26", "2425": "24-25"}


def _get_period_companies(period_code: str) -> list:
    """Return company names for the given period code from TALLY Company_List.xlsx."""
    col_name = _PERIOD_COL_MAP.get(str(period_code))
    if not col_name:
        raise ValueError(f"Unknown period code '{period_code}'. Use 2627, 2526, or 2425.")
    candidates = [
        Path.cwd() / "Files" / "TALLY Company_List.xlsx",
        Path(__file__).parent.parent.parent.parent / "Files" / "TALLY Company_List.xlsx",
    ]
    co_list_path = next((p for p in candidates if p.exists()), None)
    if co_list_path is None:
        raise FileNotFoundError(
            "Cannot find Files/TALLY Company_List.xlsx. Run from the project root."
        )
    wb2 = openpyxl.load_workbook(str(co_list_path), read_only=True)
    ws2 = wb2.active
    headers = None
    col_idx = None
    result = []
    for row in ws2.iter_rows(values_only=True):
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
            result.append(str(company_name).strip())
    wb2.close()
    return result


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Fetch Chart of Accounts masters from Tally for all open companies."
    )
    parser.add_argument("--output", default="Files/TB_Masters_Report.xlsx",
                        help="Output Excel path (default: Files/TB_Masters_Report.xlsx)")
    parser.add_argument("--host", default=TALLY_HOST,
                        help=f"Tally host (default: {TALLY_HOST})")
    parser.add_argument("--port", type=int, default=TALLY_PORT,
                        help=f"Tally port (default: {TALLY_PORT})")
    parser.add_argument("--period", default=None,
                        help="Period code 2627/2526/2425 — filters companies via TALLY Company_List.xlsx")
    args = parser.parse_args()

    output_path = Path(args.output)

    print(f"Connecting to Tally at {args.host}:{args.port} ...")
    companies = get_open_companies(args.host, args.port)
    if not companies:
        print("No open companies found in Tally. Exiting.", file=sys.stderr)
        sys.exit(1)

    if args.period:
        whitelist = _get_period_companies(args.period)
        whitelist_upper = {w.upper() for w in whitelist}
        companies = [c for c in companies if c.upper() in whitelist_upper]
        print(f"Period filter {args.period}: keeping {len(companies)} company/companies.")
        if not companies:
            print("No matching companies found for this period. Exiting.", file=sys.stderr)
            sys.exit(1)

    print(f"Found {len(companies)} open company/companies:")
    for c in companies:
        print(f"  - {c}")

    all_rows = []

    for company in companies:
        print(f"\n[{company}]")

        print("  Fetching group hierarchy ...")
        group_map = fetch_group_masters(company, args.host, args.port)
        print(f"  -> {len(group_map)} groups")

        print("  Fetching ledger masters ...")
        ledgers = fetch_ledger_masters(company, args.host, args.port)
        print(f"  -> {len(ledgers)} ledgers")

        rows = build_rows(company, ledgers, group_map)
        unmapped = sum(1 for r in rows if r["sch_nos"] == "REVIEW")
        print(f"  -> {len(rows)} rows built  ({unmapped} UNMAPPED primary groups)")
        all_rows.extend(rows)

    print(f"\nTotal rows: {len(all_rows)}")

    print("Running GL name variation detection ...")
    variations = detect_name_variations(all_rows)
    flagged = sum(1 for r in all_rows if r["variation_flag"])
    print(f"  -> {len(variations)} variation groups  |  {flagged} ledgers flagged")

    print(f"\nWriting {output_path} ...")
    write_excel(all_rows, variations, output_path)
    print(f"Done.  Report saved to: {output_path.resolve()}")
    print(f"\nSheets: Masters ({len(all_rows)} rows) | Name Variations ({len(variations)} groups) | Summary")


if __name__ == "__main__":
    main()
