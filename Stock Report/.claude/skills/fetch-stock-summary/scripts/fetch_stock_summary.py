"""
Fetch Stock Summary (closing balances) from Tally Prime for a target company and
rebuild the full stock-group hierarchy, then export to Excel.

Mirrors Tally's Stock Summary screen (F5 Stock Item-wise) but flattens every level
of the group tree so each item carries its complete path.

Output: Inventory/Stock_Summary_<company-slug>_<YYYY-MM-DD_HHMM>.xlsx
        (the previous workbook for that company moves to Inventory/Old Versions/)
  Sheet 'Hierarchy'        - indented, mirrors the Tally screen (groups + items)
  Sheet 'Item List (Flat)' - one row per item, group path split into L1..L6 columns
  Sheet 'Group Summary'    - every group with level, path, item count, rolled-up value
  Sheet 'Cross Check'      - totals reconciliation + data-quality flags

Usage:
    python fetch_stock_summary.py --company "ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)"
    python fetch_stock_summary.py --company "..." --from 2026-04-01 --to 2026-07-30
    python fetch_stock_summary.py --company "..." --expect-total 27099299.17

Default period is 1-Apr of the current financial year through today.
"""

import argparse
import os
import re
import subprocess
import sys
import time
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Layout
#   PL and BS\
#     Inventory\Stock_Summary_<company>_<timestamp>.xlsx   <- default output (visible)
#     Inventory\Old Versions\                              <- superseded workbooks
#     Stock Report\.claude\skills\fetch-stock-summary\scripts\<this file>
# The project root is the folder holding .env; it is located by walking up from
# this script, so the script produces the same output path from any CWD.
# ---------------------------------------------------------------------------

def _find_env() -> Path:
    """CWD first (project root when run per the docs), then walk up from the script."""
    if (Path.cwd() / ".env").is_file():
        return Path.cwd() / ".env"
    for _p in Path(__file__).resolve().parents:
        if (_p / ".env").is_file():
            return _p / ".env"
    # No .env anywhere: fall back to the project root implied by the skill layout.
    return Path(__file__).resolve().parents[5] / ".env"


_env_path = _find_env()
if _env_path.is_file():
    for _line in _env_path.read_text().splitlines():
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip())

PROJECT_ROOT = _env_path.parent
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "Inventory"
ARCHIVE_DIR = DEFAULT_OUTPUT_DIR / "Old Versions"

TALLY_HOST = os.environ.get("TALLY_HOST", "localhost")
TALLY_PORT = int(os.environ.get("TALLY_PORT", "9000"))

# The stock tree for these companies is four groups deep at most, so the flat
# sheet gets exactly four named level columns. Anything deeper is folded into
# the last column rather than dropped -- see the flat-sheet builder.
LEVEL_HEADERS = ["Main Category", "Sub Category", "Product Group", "Sub Group"]
MAX_LEVELS = len(LEVEL_HEADERS)


# ---------------------------------------------------------------------------
# Tally connection helpers
# ---------------------------------------------------------------------------

def post_to_tally(xml_body: str, host: str = TALLY_HOST, port: int = TALLY_PORT,
                  attempts: int = 4) -> str:
    """
    POST a TDL request to Tally, retrying on timeout.

    Tally serialises requests on a single thread, so with several companies
    loaded it will simply stop answering for a while during background work.
    That shows up as a socket timeout, not a refusal -- so back off and retry
    rather than failing a run that would have succeeded a few seconds later.
    """
    url = f"http://{host}:{port}"
    last = None
    for attempt in range(1, attempts + 1):
        req = urllib.request.Request(
            url,
            data=xml_body.encode("utf-8"),
            headers={"Content-Type": "text/xml; charset=utf-8"},
        )
        try:
            with urllib.request.urlopen(req, timeout=90 * attempt) as resp:
                return resp.read().decode("utf-8", errors="replace")
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            last = e
            if attempt < attempts:
                wait = 5 * attempt
                print(f"  Tally did not answer (attempt {attempt}/{attempts}: {e}). "
                      f"Retrying in {wait}s ...", file=sys.stderr)
                time.sleep(wait)

    raise RuntimeError(
        f"Cannot get a response from Tally at {url} after {attempts} attempts.\n"
        f"  Last error: {last}\n"
        f"  Tally is reachable but busy, or the request is too large for it right now.\n"
        f"  Try closing companies you do not need, make sure no modal dialog is open\n"
        f"  in Tally, then re-run."
    ) from last


def _clean_xml(s: str) -> str:
    """Strip numeric char refs / control chars that are invalid in XML 1.0."""
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
    try:
        return ET.fromstring(resp)
    except ET.ParseError:
        return ET.fromstring(_clean_xml(resp))


def _get_text(element, tag: str) -> str:
    el = element.find(tag)
    if el is not None and el.text:
        return el.text.strip()
    return ""


def _to_ddmmmyyyy(iso: str) -> str:
    """2026-04-01 -> 1-Apr-2026 (Tally's SVFROMDATE format)."""
    d = datetime.strptime(iso, "%Y-%m-%d").date()
    return f"{d.day}-{d.strftime('%b')}-{d.year}"


# ---------------------------------------------------------------------------
# Value parsing
# ---------------------------------------------------------------------------

def parse_qty(raw: str):
    """
    '3.000 PCS' -> (3.0, 'PCS'); '18,520.000 KGS' -> (18520.0, 'KGS')
    '9000.000 MTR = 0.000 KGS' (compound) -> first component only.
    Returns (None, '') when there is no quantity.
    """
    if not raw:
        return None, ""
    raw = raw.split("=")[0].strip()
    m = re.match(r'^\s*(-?[\d,]*\.?\d+)\s*(.*)$', raw)
    if not m:
        return None, ""
    try:
        qty = float(m.group(1).replace(",", ""))
    except ValueError:
        return None, ""
    return qty, m.group(2).strip()


def parse_amount(raw: str):
    """'-12,109.60' -> -12109.6. Empty/unparseable -> 0.0"""
    if not raw:
        return 0.0
    raw = raw.replace(",", "").strip()
    m = re.match(r'^-?\d*\.?\d+', raw)
    if not m:
        return 0.0
    try:
        return float(m.group(0))
    except ValueError:
        return 0.0


def parse_rate(raw: str):
    """'4036.53/PCS' -> 4036.53. Returns None when absent."""
    if not raw:
        return None
    m = re.match(r'^\s*(-?[\d,]*\.?\d+)', raw.replace(",", ""))
    if not m:
        return None
    try:
        return float(m.group(1))
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Fetch: stock groups and stock items
# ---------------------------------------------------------------------------

def _collection_xml(company: str, from_date: str, to_date: str,
                    coll_name: str, coll_type: str, fetch: str) -> str:
    return f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>{coll_name}</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
        <SVFROMDATE TYPE="Date">{_to_ddmmmyyyy(from_date)}</SVFROMDATE>
        <SVTODATE TYPE="Date">{_to_ddmmmyyyy(to_date)}</SVTODATE>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION ISMODIFY="No" NAME="{coll_name}">
            <TYPE>{coll_type}</TYPE>
            <FETCH>{fetch}</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""


def fetch_stock_groups(company: str, from_date: str, to_date: str, host, port) -> dict:
    """Returns {GROUP_UPPER: {'name': str, 'parent': str}}."""
    xml = _collection_xml(company, from_date, to_date,
                          "AllStockGroups", "StockGroup", "Name,Parent")
    root = _safe_parse(post_to_tally(xml, host, port))
    groups = {}
    for elem in root.iter("STOCKGROUP"):
        name = (elem.get("NAME") or _get_text(elem, "NAME")).strip()
        if not name:
            continue
        parent = _get_text(elem, "PARENT")
        if parent.lower() in ("primary", "  primary"):
            parent = ""
        groups[name.upper()] = {"name": name, "parent": parent.strip()}
    return groups


def fetch_stock_items(company: str, from_date: str, to_date: str, host, port) -> list:
    """Returns list of dicts with name/parent/qty/unit/rate/value."""
    xml = _collection_xml(
        company, from_date, to_date, "AllStockItems", "StockItem",
        "Name,Parent,BaseUnits,ClosingBalance,ClosingValue,ClosingRate",
    )
    root = _safe_parse(post_to_tally(xml, host, port))
    items = []
    for elem in root.iter("STOCKITEM"):
        name = (elem.get("NAME") or _get_text(elem, "NAME")).strip()
        if not name:
            continue
        parent = _get_text(elem, "PARENT")
        if parent.lower() == "primary":
            parent = ""
        qty, unit = parse_qty(_get_text(elem, "CLOSINGBALANCE"))
        if not unit:
            unit = _get_text(elem, "BASEUNITS")
        items.append({
            "name": name,
            "parent": parent.strip(),
            "qty": qty,
            "unit": unit,
            "rate": parse_rate(_get_text(elem, "CLOSINGRATE")),
            "value": parse_amount(_get_text(elem, "CLOSINGVALUE")),
        })
    return items


# ---------------------------------------------------------------------------
# Hierarchy construction
# ---------------------------------------------------------------------------

def build_paths(groups: dict) -> dict:
    """
    For each group return {'path': [L1, L2, ...], 'level': int}.
    Cycles and missing parents are handled defensively so one bad master
    cannot hang the walk.
    """
    resolved = {}

    def walk(key, seen):
        if key in resolved:
            return resolved[key]
        g = groups.get(key)
        if g is None:
            return None
        if key in seen:              # cycle guard
            resolved[key] = {"path": [g["name"]], "level": 1, "cycle": True}
            return resolved[key]
        seen.add(key)
        parent_key = g["parent"].upper()
        if g["parent"] and parent_key in groups:
            up = walk(parent_key, seen)
            path = (up["path"] if up else []) + [g["name"]]
        else:
            path = [g["name"]]
        resolved[key] = {"path": path, "level": len(path),
                         "cycle": False, "orphan": bool(g["parent"]) and parent_key not in groups}
        return resolved[key]

    for k in groups:
        walk(k, set())
    return resolved


def normalise_sign(items: list) -> int:
    """
    Tally's stock feed can return closing value with the opposite sign
    (credit-side convention). Flip globally when the net total is negative,
    so the workbook reads like the Stock Summary screen. Returns the multiplier.
    """
    total = sum(i["value"] for i in items)
    return -1 if total < 0 else 1


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
GRP_FILLS = ["BDD7EE", "DDEBF7", "EDF3FA", "F5F8FC", "FAFCFE", "FFFFFF"]
TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM = '#,##0.00;[Red]-#,##0.00'
QTY = '#,##0.000;[Red]-#,##0.000'


def _write_header(ws, headers, widths, freeze="A2"):
    for c, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = freeze
    ws.row_dimensions[1].height = 30


def build_workbook(out_path: Path, company: str, from_date: str, to_date: str,
                   items: list, groups: dict, paths: dict, sign: int,
                   expect_total, raw_xml_note: str, all_sheets: bool = False):
    wb = openpyxl.Workbook()

    # ---- index items by owning group, and roll values up the tree ----------
    by_group = {}
    ungrouped = []
    for it in items:
        pk = it["parent"].upper()
        if pk and pk in groups:
            by_group.setdefault(pk, []).append(it)
        else:
            ungrouped.append(it)

    children = {}
    roots = []
    for k, g in groups.items():
        pk = g["parent"].upper()
        if g["parent"] and pk in groups:
            children.setdefault(pk, []).append(k)
        else:
            roots.append(k)

    rollup = {}

    def total_of(key):
        if key in rollup:
            return rollup[key]
        v = sum(i["value"] for i in by_group.get(key, []))
        n = len(by_group.get(key, []))
        for ck in children.get(key, []):
            cv, cn = total_of(ck)
            v += cv
            n += cn
        rollup[key] = (v, n)
        return rollup[key]

    for k in groups:
        total_of(k)

    grand_total = sum(i["value"] for i in items) * sign

    def descendants(key):
        """Every item under a group, at any depth."""
        out = list(by_group.get(key, []))
        for ck in children.get(key, []):
            out.extend(descendants(ck))
        return out

    def unit_rollup(item_list):
        """
        Group quantity, reproducing what Tally's Stock Summary actually prints.

        Verified against the Tally screen for this company:
          - mixed units under the group (e.g. PAPER ROLL = ROLL + MTR) -> blank
          - net quantity of zero (e.g. MACHINERY HEAD, all PCS) -> blank, not "0.000"
        Every unit present counts, including items currently at zero quantity --
        that is what makes PAPER ROLL mixed.

        Top-level group rows are blanked by the caller: Tally does not roll a
        quantity up to the primary group row even when the unit is unambiguous
        (PRINTING INK is all KGS, SOFTWARE all PCS, and both print blank).

        Never sum across different units - the total would be meaningless.
        """
        units = {i["unit"] for i in item_list if i["unit"]}
        if len(units) != 1:
            return None, ""
        u = units.pop()
        q = sum(i["qty"] for i in item_list if i["qty"] is not None and i["unit"] == u)
        if round(q, 3) == 0:
            return None, ""
        return q, u

    # ---- Sheet 1: Hierarchy ----------------------------------------------
    ws = wb.active
    ws.title = "Hierarchy"
    ws["A1"] = "Stock Summary (Closing Balance)"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = f"Company: {company}"
    ws["A3"] = f"Period: {_to_ddmmmyyyy(from_date)} to {_to_ddmmmyyyy(to_date)}"
    ws["A4"] = f"Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}   |   Source: Tally Prime {TALLY_HOST}:{TALLY_PORT}"
    for r in (2, 3, 4):
        ws.cell(row=r, column=1).font = Font(size=9, italic=True, color="595959")

    hdr_row = 6
    headers = ["Particulars", "Type", "Level", "Quantity", "Unit", "Rate", "Value"]
    widths = [62, 10, 7, 16, 10, 14, 18]
    for c, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=hdr_row, column=c, value=h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = f"A{hdr_row + 1}"

    row = hdr_row + 1

    def emit_group(key, level):
        nonlocal row
        g = groups[key]
        val, cnt = rollup[key]
        c = ws.cell(row=row, column=1, value=g["name"])
        c.alignment = Alignment(indent=max(0, level - 1))
        c.font = Font(bold=(level <= 2), size=10)
        ws.cell(row=row, column=2, value="Group")
        ws.cell(row=row, column=3, value=level)
        ws.cell(row=row, column=7, value=round(val * sign, 2)).number_format = NUM
        fill = PatternFill("solid", fgColor=GRP_FILLS[min(level - 1, len(GRP_FILLS) - 1)])
        for cc in range(1, 8):
            ws.cell(row=row, column=cc).fill = fill
            ws.cell(row=row, column=cc).border = BORDER
        row += 1

        for it in sorted(by_group.get(key, []), key=lambda x: x["name"]):
            emit_item(it, level + 1)
        for ck in sorted(children.get(key, []), key=lambda k: groups[k]["name"]):
            emit_group(ck, level + 1)

    def emit_item(it, level):
        nonlocal row
        c = ws.cell(row=row, column=1, value=it["name"])
        c.alignment = Alignment(indent=max(0, level - 1))
        c.font = Font(italic=True, size=10)
        ws.cell(row=row, column=2, value="Item")
        ws.cell(row=row, column=3, value=level)
        if it["qty"] is not None:
            ws.cell(row=row, column=4, value=round(it["qty"], 3)).number_format = QTY
        ws.cell(row=row, column=5, value=it["unit"])
        if it["rate"] is not None:
            ws.cell(row=row, column=6, value=round(it["rate"], 2)).number_format = NUM
        ws.cell(row=row, column=7, value=round(it["value"] * sign, 2)).number_format = NUM
        for cc in range(1, 8):
            ws.cell(row=row, column=cc).border = BORDER
        row += 1

    for rk in sorted(roots, key=lambda k: groups[k]["name"]):
        emit_group(rk, 1)

    if ungrouped:
        c = ws.cell(row=row, column=1, value="(No stock group assigned)")
        c.font = Font(bold=True, size=10, color="C00000")
        ws.cell(row=row, column=2, value="Group")
        ws.cell(row=row, column=3, value=1)
        ws.cell(row=row, column=7,
                value=round(sum(i["value"] for i in ungrouped) * sign, 2)).number_format = NUM
        row += 1
        for it in sorted(ungrouped, key=lambda x: x["name"]):
            emit_item(it, 2)

    ws.cell(row=row, column=1, value="Grand Total").font = Font(bold=True, size=11)
    gt = ws.cell(row=row, column=7, value=round(grand_total, 2))
    gt.number_format = NUM
    gt.font = Font(bold=True, size=11)
    for cc in range(1, 8):
        ws.cell(row=row, column=cc).fill = TOTAL_FILL
        ws.cell(row=row, column=cc).border = BORDER

    # ---- Sheet 2: Item List (Flat) ---------------------------------------
    ws2 = wb.create_sheet("Item List (Flat)")
    headers2 = LEVEL_HEADERS + ["Full Group Path", "Item Name", "Quantity", "Unit",
                                "Rate", "Value", "Depth", "Has Qty", "Zero Value"]
    widths2 = [30, 28, 28, 30] + [58, 46, 15, 9, 13, 17, 8, 9, 11]
    _write_header(ws2, headers2, widths2)

    r2 = 2
    for it in sorted(items, key=lambda x: (x["parent"].upper(), x["name"])):
        pk = it["parent"].upper()
        info = paths.get(pk)
        path = info["path"] if info else (["(No stock group assigned)"] if not it["parent"] else [it["parent"]])
        # Levels shallower than an item's depth stay blank on purpose: an empty
        # cell means that level genuinely does not exist for that item.
        # A path deeper than MAX_LEVELS is folded into the last column so no
        # group name is ever lost; 'Full Group Path' always holds the complete chain.
        for i in range(MAX_LEVELS):
            if i < len(path):
                last = (i == MAX_LEVELS - 1)
                v = " > ".join(path[i:]) if (last and len(path) > MAX_LEVELS) else path[i]
            else:
                v = ""
            ws2.cell(row=r2, column=i + 1, value=v)
        ws2.cell(row=r2, column=MAX_LEVELS + 1, value=" > ".join(path))
        ws2.cell(row=r2, column=MAX_LEVELS + 2, value=it["name"])
        if it["qty"] is not None:
            ws2.cell(row=r2, column=MAX_LEVELS + 3, value=round(it["qty"], 3)).number_format = QTY
        ws2.cell(row=r2, column=MAX_LEVELS + 4, value=it["unit"])
        if it["rate"] is not None:
            ws2.cell(row=r2, column=MAX_LEVELS + 5, value=round(it["rate"], 2)).number_format = NUM
        v = round(it["value"] * sign, 2)
        ws2.cell(row=r2, column=MAX_LEVELS + 6, value=v).number_format = NUM
        ws2.cell(row=r2, column=MAX_LEVELS + 7, value=len(path))
        ws2.cell(row=r2, column=MAX_LEVELS + 8, value="Yes" if it["qty"] not in (None, 0) else "No")
        ws2.cell(row=r2, column=MAX_LEVELS + 9, value="Yes" if v == 0 else "No")
        r2 += 1
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(headers2))}{max(r2 - 1, 1)}"

    # ---- Sheet 3: Group Summary ------------------------------------------
    ws3 = wb.create_sheet("Group Summary")
    headers3 = ["Level", "Group Name", "Parent Group", "Full Path",
                "Direct Items", "Total Items (incl. sub-groups)",
                "Direct Value", "Rolled-up Value", "% of Grand Total"]
    widths3 = [7, 40, 34, 62, 13, 16, 18, 18, 15]
    _write_header(ws3, headers3, widths3)

    r3 = 2
    for k in sorted(groups, key=lambda k: (paths.get(k, {}).get("path") or [groups[k]["name"]])):
        g = groups[k]
        info = paths.get(k, {"path": [g["name"]], "level": 1})
        direct = by_group.get(k, [])
        rolled, tot_n = rollup[k]
        ws3.cell(row=r3, column=1, value=info["level"])
        ws3.cell(row=r3, column=2, value=g["name"]).font = Font(bold=(info["level"] == 1), size=10)
        ws3.cell(row=r3, column=3, value=g["parent"] or "(Primary)")
        ws3.cell(row=r3, column=4, value=" > ".join(info["path"]))
        ws3.cell(row=r3, column=5, value=len(direct))
        ws3.cell(row=r3, column=6, value=tot_n)
        ws3.cell(row=r3, column=7,
                 value=round(sum(i["value"] for i in direct) * sign, 2)).number_format = NUM
        ws3.cell(row=r3, column=8, value=round(rolled * sign, 2)).number_format = NUM
        pct = ws3.cell(row=r3, column=9,
                       value=(rolled * sign / grand_total) if grand_total else 0)
        pct.number_format = '0.00%'
        r3 += 1
    ws3.auto_filter.ref = f"A1:{get_column_letter(len(headers3))}{max(r3 - 1, 1)}"

    # ---- Sheet 4: Cross Check --------------------------------------------
    ws4 = wb.create_sheet("Cross Check")
    _write_header(ws4, ["Check", "Value", "Status", "Note"], [46, 24, 14, 74])

    zero_val = [i for i in items if round(i["value"] * sign, 2) == 0]
    no_qty = [i for i in items if i["qty"] in (None, 0) and round(i["value"] * sign, 2) != 0]
    neg_val = [i for i in items if round(i["value"] * sign, 2) < 0]
    orphan_groups = [groups[k]["name"] for k, v in paths.items() if v.get("orphan")]
    cycle_groups = [groups[k]["name"] for k, v in paths.items() if v.get("cycle")]
    l1_sum = sum(rollup[k][0] for k in roots) * sign + sum(i["value"] for i in ungrouped) * sign

    checks = [
        # Company and period lead the sheet: the Index builder reads them from
        # here, and they make an emailed workbook self-describing.
        ("Company", company, "INFO", "Tally company this workbook was fetched from."),
        ("Period", f"{_to_ddmmmyyyy(from_date)} to {_to_ddmmmyyyy(to_date)}", "INFO",
         "Closing balances as at the end of this period."),
        ("Generated", datetime.now().strftime("%d-%b-%Y %H:%M"), "INFO",
         f"Live from Tally Prime at {TALLY_HOST}:{TALLY_PORT} - not a cached copy."),
        ("Grand Total (sum of all items)", round(grand_total, 2), "INFO",
         "Sum of every item's closing value, sign-normalised."),
        ("Sum of top-level group roll-ups", round(l1_sum, 2),
         "OK" if abs(l1_sum - grand_total) < 0.01 else "MISMATCH",
         "Must equal Grand Total, else an item is orphaned from the tree."),
        ("Total stock items fetched", len(items), "INFO", ""),
        ("Total stock groups fetched", len(groups), "INFO", ""),
        ("Items with no stock group", len(ungrouped),
         "OK" if not ungrouped else "REVIEW",
         "; ".join(i["name"] for i in ungrouped[:10]) if ungrouped else "All items sit under a group."),
        ("Items with zero closing value", len(zero_val), "INFO",
         "Normal - fully consumed or never stocked. Tally hides these on screen."),
        ("Items with value but no quantity", len(no_qty),
         "INFO" if not no_qty else "REVIEW",
         "; ".join(i["name"] for i in no_qty[:10]) if no_qty else ""),
        ("Items with NEGATIVE closing value", len(neg_val),
         "OK" if not neg_val else "REVIEW",
         "; ".join(f"{i['name']} ({round(i['value']*sign,2):,.2f})" for i in neg_val[:10])
         if neg_val else "No negative stock."),
        ("Groups with a missing parent", len(orphan_groups),
         "OK" if not orphan_groups else "REVIEW", "; ".join(orphan_groups[:10])),
        ("Groups in a parent cycle", len(cycle_groups),
         "OK" if not cycle_groups else "REVIEW", "; ".join(cycle_groups[:10])),
        ("Sign normalisation applied", "Yes (flipped)" if sign == -1 else "No", "INFO", raw_xml_note),
    ]
    if expect_total is not None:
        diff = round(grand_total - expect_total, 2)
        checks.insert(1, (
            "Grand Total vs expected (--expect-total)", round(expect_total, 2),
            "OK" if abs(diff) < 0.01 else "MISMATCH",
            f"Difference: {diff:,.2f}",
        ))

    for i, (label, val, status, note) in enumerate(checks, start=2):
        ws4.cell(row=i, column=1, value=label)
        c = ws4.cell(row=i, column=2, value=val)
        if isinstance(val, float):
            c.number_format = NUM
        st = ws4.cell(row=i, column=3, value=status)
        st.font = Font(bold=True, color={"OK": "006100", "MISMATCH": "9C0006",
                                         "REVIEW": "9C5700"}.get(status, "404040"))
        st.alignment = Alignment(horizontal="center")
        ws4.cell(row=i, column=4, value=note).alignment = Alignment(wrap_text=True, vertical="top")

    # ---- Sheet 0: Summary (built last, moved to front) --------------------
    # Depth = top-level groups + what sits directly inside each, collapsible.
    # Zero-value rows are hidden to match the Tally screen, and counted in a
    # footnote so nothing is silently dropped.
    wss = wb.create_sheet("Summary")
    wss.sheet_properties.outlinePr.summaryBelow = False

    wss["A1"] = "STOCK SUMMARY"
    wss["A1"].font = Font(bold=True, size=15, color="1F4E79")
    wss["A2"] = company
    wss["A2"].font = Font(bold=True, size=10)
    wss["A3"] = f"Closing balance for the period {_to_ddmmmyyyy(from_date)} to {_to_ddmmmyyyy(to_date)}"
    wss["A4"] = (f"Generated {datetime.now().strftime('%d-%b-%Y %H:%M')} from Tally Prime "
                 f"({TALLY_HOST}:{TALLY_PORT}) - live data, not a cached copy")
    for r in (3, 4):
        wss.cell(row=r, column=1).font = Font(size=9, italic=True, color="595959")

    s_hdr = 6
    s_headers = ["Particulars", "Quantity", "Unit", "Rate", "Value", "% of Total"]
    s_widths = [50, 15, 9, 13, 18, 11]
    for c, (h, w) in enumerate(zip(s_headers, s_widths), start=1):
        cell = wss.cell(row=s_hdr, column=c, value=h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        wss.column_dimensions[get_column_letter(c)].width = w
    wss.freeze_panes = f"A{s_hdr + 1}"

    L1_FILL = PatternFill("solid", fgColor="BDD7EE")
    hidden_l1, hidden_l2 = [], []
    chart_rows = []       # (name, value) for the bar chart
    r = s_hdr + 1

    def put(row, name, qty, unit, rate, value, *, bold, indent, fill=None):
        c1 = wss.cell(row=row, column=1, value=name)
        c1.alignment = Alignment(indent=indent)
        c1.font = Font(bold=bold, size=10)
        if qty is not None:
            wss.cell(row=row, column=2, value=round(qty, 3)).number_format = QTY
        wss.cell(row=row, column=3, value=unit).alignment = Alignment(horizontal="center")
        if rate is not None:
            wss.cell(row=row, column=4, value=round(rate, 2)).number_format = NUM
        wss.cell(row=row, column=5, value=round(value, 2)).number_format = NUM
        p = wss.cell(row=row, column=6, value=(value / grand_total) if grand_total else 0)
        p.number_format = '0.00%'
        for cc in range(1, 7):
            wss.cell(row=row, column=cc).border = BORDER
            if fill:
                wss.cell(row=row, column=cc).fill = fill

    l1_sorted = sorted(roots, key=lambda k: -rollup[k][0] * sign)
    for rk in l1_sorted:
        val = rollup[rk][0] * sign
        if round(val, 2) == 0:
            hidden_l1.append(groups[rk]["name"])
            continue
        # Top-level rows carry no quantity/rate - see unit_rollup().
        put(r, groups[rk]["name"], None, "", None, val,
            bold=True, indent=0, fill=L1_FILL)
        chart_rows.append((groups[rk]["name"], round(val, 2)))
        r += 1

        # Children = sub-groups + items sitting directly on this group.
        kids = []
        for ck in children.get(rk, []):
            cv = rollup[ck][0] * sign
            kids.append((groups[ck]["name"], cv, unit_rollup(descendants(ck)), False))
        for it in by_group.get(rk, []):
            kids.append((it["name"], it["value"] * sign, (it["qty"], it["unit"]), True))
        kids.sort(key=lambda x: -x[1])

        for name, cv, (q2, u2), is_item in kids:
            if round(cv, 2) == 0:
                hidden_l2.append(name)
                continue
            label = f"{name}  (item)" if is_item else name
            put(r, label, q2, u2, (cv / q2 if q2 else None), cv,
                bold=False, indent=2)
            wss.row_dimensions[r].outlineLevel = 1
            r += 1

    if ungrouped:
        uv = sum(i["value"] for i in ungrouped) * sign
        if round(uv, 2) != 0:
            put(r, "(No stock group assigned)", None, "", None, uv, bold=True, indent=0)
            wss.cell(row=r, column=1).font = Font(bold=True, size=10, color="C00000")
            r += 1

    total_row = r
    put(total_row, "GRAND TOTAL", None, "", None, grand_total, bold=True, indent=0,
        fill=TOTAL_FILL)
    for cc in range(1, 7):
        wss.cell(row=total_row, column=cc).font = Font(bold=True, size=11)
    r += 2

    # ---- Footnotes: what the table deliberately does not show -------------
    notes = [
        f"Rows shown: top-level groups and what sits directly inside each. "
        f"Use the [-] / [+] outline buttons at the left to collapse or expand. "
        f"Full detail for all {len(items):,} items is on the 'Hierarchy' sheet.",
    ]
    if hidden_l1:
        notes.append(f"Hidden - nil value: {len(hidden_l1)} top-level group(s) carry zero value "
                     f"and are omitted, exactly as Tally hides them on screen: "
                     f"{', '.join(sorted(hidden_l1))}.")
    if hidden_l2:
        notes.append(f"Hidden - nil value: {len(hidden_l2)} sub-group(s)/item(s) at the second "
                     f"level carry zero value and are omitted. They remain on the "
                     f"'Group Summary' and 'Item List (Flat)' sheets.")
    notes.append("Quantity and Rate are blank wherever a group's contents use more than one unit "
                 "of measure - a mixed-UOM total would be meaningless, so none is shown.")

    wss.cell(row=r, column=1, value="Notes").font = Font(bold=True, size=10, color="1F4E79")
    r += 1
    for n in notes:
        c = wss.cell(row=r, column=1, value=n)
        c.font = Font(size=9, color="404040")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        wss.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        wss.row_dimensions[r].height = 26
        r += 1
    r += 1

    # ---- Exceptions callout ----------------------------------------------
    exc = []
    for i in sorted(items, key=lambda x: x["value"] * sign):
        v = round(i["value"] * sign, 2)
        if v < 0:
            exc.append(("Negative stock value", i["name"],
                        f"{v:,.2f}" + (f"  (qty {i['qty']:,.3f} {i['unit']})"
                                       if i["qty"] not in (None, 0) else ""),
                        "Issues booked without a matching receipt. Understates closing stock."))
    for i in items:
        v = round(i["value"] * sign, 2)
        if v > 0 and i["qty"] in (None, 0):
            exc.append(("Value but no quantity", i["name"], f"{v:,.2f}",
                        "Value stranded against nothing on hand."))
    for i in ungrouped:
        # `or 0.0` collapses a negative zero, which would print as "-0.00".
        exc.append(("No stock group assigned", i["name"],
                    f"{round(i['value'] * sign, 2) or 0.0:,.2f}",
                    "Master-data gap - falls outside every group report."))

    wss.cell(row=r, column=1, value="EXCEPTIONS TO REVIEW").font = Font(bold=True, size=11,
                                                                       color="9C0006")
    r += 1
    if exc:
        for c, (h, w) in enumerate(zip(["Issue", "Item", "Amount", "Why it matters"],
                                       [26, 42, 30, 60]), start=1):
            cell = wss.cell(row=r, column=c, value=h)
            cell.fill = PatternFill("solid", fgColor="F8CBAD")
            cell.font = Font(bold=True, size=9)
            cell.border = BORDER
        r += 1
        for issue, name, amt, why in exc:
            for c, v in enumerate([issue, name, amt, why], start=1):
                cell = wss.cell(row=r, column=c, value=v)
                cell.font = Font(size=9)
                cell.border = BORDER
                cell.alignment = Alignment(wrap_text=(c == 4), vertical="top")
            r += 1
    else:
        wss.cell(row=r, column=1, value="None - no negative stock, no stranded value, "
                                        "no ungrouped items.").font = Font(size=9, color="006100")
        r += 1

    # ---- Bar chart (fed from a hidden helper block) ------------------------
    if chart_rows:
        hcol = 14  # column N
        wss.cell(row=s_hdr, column=hcol, value="Group")
        wss.cell(row=s_hdr, column=hcol + 1, value="Value")
        for i, (nm, vv) in enumerate(chart_rows, start=1):
            wss.cell(row=s_hdr + i, column=hcol, value=nm)
            wss.cell(row=s_hdr + i, column=hcol + 1, value=vv)
        for cl in (get_column_letter(hcol), get_column_letter(hcol + 1)):
            wss.column_dimensions[cl].hidden = True

        ch = BarChart()
        ch.type = "bar"
        ch.style = 10
        ch.title = "Closing stock value by top-level group"
        ch.y_axis.title = "Value"
        ch.legend = None
        ch.add_data(Reference(wss, min_col=hcol + 1, min_row=s_hdr,
                              max_row=s_hdr + len(chart_rows)), titles_from_data=True)
        ch.set_categories(Reference(wss, min_col=hcol, min_row=s_hdr + 1,
                                    max_row=s_hdr + len(chart_rows)))
        ch.height, ch.width = 7.5, 17
        wss.add_chart(ch, f"H{s_hdr + 1}")

    wb.move_sheet("Summary", offset=-(len(wb.sheetnames) - 1))

    # Default output is the two working sheets only. The Summary / Hierarchy /
    # Group Summary builders above are kept intact -- pass --all-sheets to emit
    # them again rather than re-writing that code.
    if not all_sheets:
        for name in ("Summary", "Hierarchy", "Group Summary"):
            if name in wb.sheetnames:
                del wb[name]
        wb.move_sheet("Item List (Flat)", offset=-wb.sheetnames.index("Item List (Flat)"))

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return grand_total, len(items), len(groups), checks


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def default_fy_start(today: date) -> str:
    """Indian FY: 1-Apr of the current financial year."""
    y = today.year if today.month >= 4 else today.year - 1
    return f"{y}-04-01"


def slug(s: str) -> str:
    s = re.sub(r'[^A-Za-z0-9]+', '_', s).strip('_')
    return s[:48]


def main():
    today = date.today()
    ap = argparse.ArgumentParser(description="Fetch Tally stock summary with full group hierarchy.")
    ap.add_argument("--company", required=True, help="Exact Tally company name")
    ap.add_argument("--from", dest="from_date", default=default_fy_start(today),
                    help="From date YYYY-MM-DD (default: 1-Apr of current FY)")
    ap.add_argument("--to", dest="to_date", default=today.isoformat(),
                    help="To date YYYY-MM-DD (default: today)")
    ap.add_argument("--output", default=None, help="Output .xlsx path")
    ap.add_argument("--expect-total", type=float, default=None,
                    help="Expected grand total for cross-check (positive number)")
    ap.add_argument("--all-sheets", action="store_true",
                    help="Also emit Summary, Hierarchy and Group Summary "
                         "(default is Item List (Flat) + Cross Check only)")
    ap.add_argument("--no-pivots", action="store_true",
                    help="Skip the native PivotTable sheets (needs Excel/xlwings)")
    ap.add_argument("--host", default=TALLY_HOST)
    ap.add_argument("--port", type=int, default=TALLY_PORT)
    args = ap.parse_args()

    print(f"Tally      : {args.host}:{args.port}")
    print(f"Company    : {args.company}")
    print(f"Period     : {_to_ddmmmyyyy(args.from_date)} to {_to_ddmmmyyyy(args.to_date)}")

    print("Fetching stock groups ...")
    groups = fetch_stock_groups(args.company, args.from_date, args.to_date, args.host, args.port)
    print(f"  {len(groups)} groups")
    if not groups:
        print("ERROR: no stock groups returned - check the company name is exact and open.",
              file=sys.stderr)
        return 1

    print("Fetching stock items ...")
    items = fetch_stock_items(args.company, args.from_date, args.to_date, args.host, args.port)
    print(f"  {len(items)} items")
    if not items:
        print("ERROR: no stock items returned.", file=sys.stderr)
        return 1

    paths = build_paths(groups)
    sign = normalise_sign(items)
    note = ("Tally returned closing values on the credit side; all values multiplied by -1 "
            "so the workbook matches the Stock Summary screen."
            if sign == -1 else "Values used exactly as returned by Tally.")

    ts = datetime.now().strftime("%Y-%m-%d_%H%M")
    out = Path(args.output) if args.output else \
        DEFAULT_OUTPUT_DIR / f"Stock_Summary_{slug(args.company)}_{ts}.xlsx"

    total, n_items, n_groups, checks = build_workbook(
        out, args.company, args.from_date, args.to_date,
        items, groups, paths, sign, args.expect_total, note, args.all_sheets)

    # Keep exactly one current workbook per company visible; archive the rest.
    if not args.output:
        prefix = f"Stock_Summary_{slug(args.company)}_"
        superseded = [p for p in DEFAULT_OUTPUT_DIR.glob(f"{prefix}*.xlsx")
                      if p.resolve() != out.resolve()]
        if superseded:
            ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
            moved, locked = 0, []
            for p in superseded:
                try:
                    p.replace(ARCHIVE_DIR / p.name)
                    moved += 1
                except OSError:
                    # Almost always the file is open in Excel. Never let tidy-up
                    # kill a run whose workbook has already been written.
                    locked.append(p.name)
            if moved:
                print(f"Archived {moved} superseded workbook(s) to {ARCHIVE_DIR}")
            if locked:
                print(f"NOTE: could not archive {len(locked)} file(s) - open in Excel? "
                      f"Close and re-run, or move manually: {', '.join(locked)}")

    print(f"\nGrand Total: {total:,.2f}")
    print(f"Items      : {n_items}   Groups: {n_groups}")
    print("\nCross Check:")
    for label, val, status, _ in checks:
        v = f"{val:,.2f}" if isinstance(val, float) else val
        print(f"  [{status:8}] {label}: {v}")
    # Native PivotTables are added in a second pass via Excel COM. A failure
    # there must not invalidate the run -- the data workbook is already saved.
    if not args.no_pivots:
        print("\n--- Building PivotTables ---")
        pivot_script = Path(__file__).with_name("create_stock_pivots.py")
        if pivot_script.is_file():
            rc = subprocess.call([sys.executable, str(pivot_script), str(out.resolve())])
            if rc != 0:
                print("WARNING: PivotTables were not added. The workbook still contains "
                      "'Item List (Flat)' and 'Cross Check' and is fully usable.\n"
                      "         Re-run with the workbook closed, or pass --no-pivots.",
                      file=sys.stderr)
        else:
            print(f"WARNING: {pivot_script.name} not found - skipping pivots.",
                  file=sys.stderr)

    print(f"\nSaved: {out.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
