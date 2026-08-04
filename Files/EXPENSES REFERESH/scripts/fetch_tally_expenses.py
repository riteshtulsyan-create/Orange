"""
Tool: fetch_tally_expenses.py
Fetches Direct and Indirect Expense vouchers from Tally via the TallyXML API.
Returns a list of parsed expense row dicts (23 columns, incl. DEPARTMENT and
COMPONENT CATEGORY) ready for Excel output.

Usage (as a module):
    from fetch_tally_expenses import fetch_expense_rows, get_open_companies
    open_cos = get_open_companies()
    company_cfg = {"company_col": "Otec- Surat", "location": "SURAT", "company_name": "Otec"}
    rows = fetch_expense_rows(company_name, company_cfg, "20250401", "20260331")

Usage (standalone test):
    python fetch_tally_expenses.py --company "COMPANY NAME" --from 20250401 --to 20260331
"""

import os
import re
import sys
import argparse
import requests
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path


def _load_env():
    """Load TALLY_HOST/TALLY_PORT from the project's .env file (per CLAUDE.md -
    Tally always connects via localhost:9010, never hardcode a different port)."""
    env_path = Path(__file__).resolve().parents[3] / ".env"
    values = {}
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            values[k.strip()] = v.strip()
    return values


_ENV = _load_env()
TALLY_HOST    = os.environ.get("TALLY_HOST", _ENV.get("TALLY_HOST", "localhost"))
TALLY_PORT    = os.environ.get("TALLY_PORT", _ENV.get("TALLY_PORT", "9010"))
TALLY_URL     = f"http://{TALLY_HOST}:{TALLY_PORT}"
TALLY_TIMEOUT = 300  # seconds — full-company voucher export can be slow

# Employee -> Department lookup (no Department field exists in Tally's ledger/
# voucher masters for this company; this mapping was extracted from the existing
# Expense Sheet, where it had been maintained manually). Employees not listed
# here get a blank DEPARTMENT, same as before.
EMPLOYEE_DEPARTMENT = {
    "SHIKHA KARAN TOSHNIWAL": "Management", "DHANANJAY PATEL": "Sales",
    "MEHUL GAJJAR": "Sales", "AMIT SHARMA": "Supply Chain",
    "KHUSHBOO AAYUSH RATHI": "Management", "RAKESH KUMAR": "After Sales Service",
    "LAXMIKANT PATIL": "After Sales Service", "VIKAS CHAURASIYA": "After Sales Service",
    "NAKULESHWAR SHARMA": "Sales", "AAYUSHI SHAH": "Administration",
    "KHURSHID ALAM": "Sales", "SNEHA ABHISHEK SONI": "Human Resource",
    "SHIV KISHAN RATHI": "Management", "SAMADHAN PATIL": "Supply Chain",
    "MAYUR ASHOKBHAI PATEL": "After Sales Service",
    "PRATIKSHA SURENDRA HIWASE": "After Sales Service", "SHYAM LAL": "After Sales Service",
    "OM SINGH": "After Sales Service", "VINIT MISHRA": "After Sales Service",
    "ABHISHEK": "M/C Manufacturing", "RAHUL GHANSHYAM PANCHARIYA": "After Sales Service",
    "UMAIR NISAR AHMAD ANSARI": "After Sales Service", "NAND BHANWAR SINGH": "After Sales Service",
    "HIMANSHU SOJITRA": "After Sales Service", "JAYDEEP CHAUHAN": "After Sales Service",
    "KIRIT RAJENDRA PATIL": "After Sales Service", "LOKENDRA SINGH": "After Sales Service",
    "NAHAR SINGH": "After Sales Service", "NIKHIL GHADIYA": "After Sales Service",
    "SANDIPKUMAR PARAGBHAI ROHIT": "After Sales Service", "VINAY SHARMA": "After Sales Service",
    "KAMAL SINGH": "After Sales Service", "SOURABH RAKESH NAGPAL": "After Sales Service",
    "AJAY BALUBHAI SOLANKI": "After Sales Service", "DINANATH MACHHINDRA SONAWANE": "After Sales Service",
    "BHARGAVKUMAR JAYESHBHAI CHAUHAN": "After Sales Service", "SHUBHAM RAJENDRA PATIL": "After Sales Service",
    "VIVEK ASHOK SINGH": "Administration", "SWAPNIL BHARAT BHOITE": "After Sales Service",
    "JAY GULGULIA": "Sales", "MANMOHAN TOTLA": "Sales", "MANOJ PAWAR": "After Sales Service",
    "BHAVIN RAMJIBHAI PATEL": "After Sales Service",
    "PRASHANT GHANSHYAMBHAI PANCHARIYA": "After Sales Service", "DIPAK GUPTA": "Finance",
    "PRAVIN BABUBHAI TANDEL": "After Sales Service", "SUHEL YUSUF TOORAWA": "After Sales Service",
    "Riya Kumari": "Human Resource", "DEEPAK NARAYAN PATIL": "Administration",
    "VIKAS": "Supply Chain", "LALIT SHARMA": "Supply Chain", "Kartik Joshi": "After Sales Service",
    "ROSHAN RATHOD": "After Sales Service", "BHARAT SINGH": "Ink Manufacturing",
    "RAJNEESH KUMAR": "M/C Manufacturing", "ROSHAN KUMAR": "After Sales Service",
    "Jagwinder Singh": "After Sales Service",
}

DIRECT_EXPENSES   = "Direct Expenses"
INDIRECT_EXPENSES = "Indirect Expenses"
EXP_TYPE_GROUPS   = {DIRECT_EXPENSES, INDIRECT_EXPENSES}


def classify_component_category(types_val: str, ledger_val: str) -> str:
    """Classify a DIRECT-expense row into a component category (Head/Ink/Machine/
    Spare/Fabric) based on keywords in the ledger's TYPES/LEDGER NAME - same rule
    used to build the 'Direct by Component' sheet. Blank for INDIRECT rows."""
    text = f"{types_val or ''} {ledger_val or ''}".upper()
    hits = []
    if re.search(r"\bFABRIC\b", text):
        hits.append("FABRIC")
    if re.search(r"\bHEAD\b", text):
        hits.append("HEAD")
    if re.search(r"\bINK\b", text):
        hits.append("INK")
    if re.search(r"\bMACHINE(RY)?\b", text):
        hits.append("MACHINE")
    if re.search(r"\bSPARE\b", text):
        hits.append("SPARE")
    if not hits:
        return "GENERAL / OTHER"
    seen = []
    for h in hits:
        if h not in seen:
            seen.append(h)
    return " & ".join(seen)


# ---------------------------------------------------------------------------
# XML helpers (same pattern as other Tally tools)
# ---------------------------------------------------------------------------

def _clean_xml(raw_bytes: bytes) -> bytes:
    """Fix Tally XML quirks before parsing."""
    text = raw_bytes.decode("utf-8", errors="replace")

    def _strip_invalid_ref(m):
        n = int(m.group(1))
        if n in (9, 10, 13) or (0x20 <= n <= 0xD7FF) or (0xE000 <= n <= 0xFFFD):
            return m.group(0)
        return ""

    text = re.sub(r"&#(\d+);", _strip_invalid_ref, text)
    text = text.replace("<ENVELOPE>", '<ENVELOPE xmlns:UDF="TallyUDF">', 1)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    return text.encode("utf-8")


def _txt(el, tag, default=""):
    val = el.findtext(tag, default)
    return val.strip() if val else default


def _parse_number(s: str) -> float:
    if not s:
        return 0.0
    s = s.strip()

    # Tally foreign-currency format: "-$620.67 @ ? 95.43/$ = -? 59228.37"
    # The INR-equivalent appears after the last '=' sign.
    if "=" in s:
        after_eq = s.split("=")[-1].strip()
        neg = after_eq.startswith("-")
        m = re.search(r"([\d,]+\.?\d*)", after_eq)
        if m:
            try:
                return -float(m.group(1).replace(",", "")) if neg else float(m.group(1).replace(",", ""))
            except ValueError:
                pass

    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _post_tally(xml_body: str, label: str) -> ET.Element:
    """POST to Tally, clean XML, return root element."""
    try:
        resp = requests.post(
            TALLY_URL,
            data=xml_body.encode("utf-8"),
            headers={"Content-Type": "application/xml"},
            timeout=TALLY_TIMEOUT,
        )
        resp.raise_for_status()
    except requests.exceptions.ConnectionError as e:
        raise RuntimeError(f"Cannot connect to Tally at {TALLY_URL}: {e}") from e
    except requests.exceptions.Timeout:
        raise RuntimeError(
            f"Tally timed out after {TALLY_TIMEOUT}s for: {label}. "
            "Restart Tally and try again."
        )
    except requests.exceptions.HTTPError as e:
        raise RuntimeError(f"HTTP error from Tally ({label}): {e}") from e
    try:
        return ET.fromstring(_clean_xml(resp.content))
    except ET.ParseError as e:
        raise RuntimeError(f"XML parse failed ({label}): {e}") from e


# ---------------------------------------------------------------------------
# Open company detection
# ---------------------------------------------------------------------------

_COMPANY_LIST_XML = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>List of Companies</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="List of Companies" ISMODIFY="No">
            <TYPE>Company</TYPE>
            <FETCH>NAME</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""


def get_open_companies() -> list[str]:
    """Return exact names of all companies currently open in Tally."""
    try:
        resp = requests.post(
            TALLY_URL,
            data=_COMPANY_LIST_XML.encode("utf-8"),
            headers={"Content-Type": "application/xml"},
            timeout=60,
        )
        resp.raise_for_status()
    except Exception as e:
        raise RuntimeError(f"Cannot reach Tally: {e}") from e

    root = ET.fromstring(_clean_xml(resp.content))
    names = []
    for c in root.findall(".//DATA/COLLECTION/COMPANY"):
        name = (c.findtext("NAME") or c.get("NAME") or "").strip()
        if name:
            names.append(name)
    return names


# ---------------------------------------------------------------------------
# Phase A: Ledger + Group masters → expense hierarchy map
# ---------------------------------------------------------------------------

def _ledger_master_xml(company: str) -> str:
    return f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE><ID>LedgerMaster</ID></HEADER>
  <BODY><DESC>
    <STATICVARIABLES>
      <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
    </STATICVARIABLES>
    <TDL><TDLMESSAGE>
      <COLLECTION NAME="LedgerMaster" ISMODIFY="No">
        <TYPE>Ledger</TYPE>
        <FETCH>NAME,PARENT</FETCH>
      </COLLECTION>
    </TDLMESSAGE></TDL>
  </DESC></BODY>
</ENVELOPE>"""


def _group_master_xml(company: str) -> str:
    return f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE><ID>GroupMaster</ID></HEADER>
  <BODY><DESC>
    <STATICVARIABLES>
      <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
    </STATICVARIABLES>
    <TDL><TDLMESSAGE>
      <COLLECTION NAME="GroupMaster" ISMODIFY="No">
        <TYPE>Group</TYPE>
        <FETCH>NAME,PARENT</FETCH>
      </COLLECTION>
    </TDLMESSAGE></TDL>
  </DESC></BODY>
</ENVELOPE>"""


def _fetch_ledger_parent(company: str) -> dict[str, str]:
    """Returns {ledger_name: parent_group_name}."""
    root = _post_tally(_ledger_master_xml(company), f"LedgerMaster/{company[:30]}")
    master = {}
    for ledger in root.findall(".//LEDGER"):
        name   = (ledger.get("NAME") or _txt(ledger, "NAME")).strip()
        parent = _txt(ledger, "PARENT")
        if name:
            master[name] = parent
    return master


def _fetch_group_parent(company: str) -> dict[str, str]:
    """Returns {group_name: parent_group_name}."""
    root = _post_tally(_group_master_xml(company), f"GroupMaster/{company[:30]}")
    master = {}
    for group in root.findall(".//GROUP"):
        name   = (group.get("NAME") or _txt(group, "NAME")).strip()
        parent = _txt(group, "PARENT")
        if name:
            master[name] = parent
    return master


def build_expense_ledger_map(
    ledger_parent: dict[str, str],
    group_parent:  dict[str, str],
) -> dict[str, dict]:
    """
    Walk up the group tree for each ledger to find its expense classification.

    Returns only expense ledgers (those that resolve to Direct or Indirect Expenses):
    {
        ledger_name: {
            "exp_type":     "DIRECT" | "INDIRECT",
            "main_group":   str,   # 1st level below EXP_TYPE
            "sub_group":    str,   # 2nd level (or same as main if depth < 2)
            "sub_sub_group": str,  # 3rd level (or same as sub if depth < 3)
        }
    }

    Depth rule for filling the 3 group columns:
      depth 1 (g1 → ledger)          → MAIN=g1, SUB=g1, SUB_SUB=g1
      depth 2 (g1 → g2 → ledger)     → MAIN=g1, SUB=g2, SUB_SUB=g2
      depth 3+ (g1 → g2 → g3 → ...)  → MAIN=g1, SUB=g2, SUB_SUB=g3
    """
    result = {}

    for ledger, direct_parent in ledger_parent.items():
        # Walk up from direct_parent, collecting the group path
        path    = []
        current = direct_parent
        visited = set()

        while current and current not in EXP_TYPE_GROUPS and current not in visited:
            visited.add(current)
            path.append(current)
            current = group_parent.get(current, "")

        if current not in EXP_TYPE_GROUPS:
            continue  # ledger does not belong to Direct or Indirect Expenses

        exp_type = "DIRECT" if current == DIRECT_EXPENSES else "INDIRECT"

        # path is built bottom-up; reverse to get top-down (MAIN GROUP first)
        path.reverse()

        depth = len(path)
        if depth == 0:
            # Ledger sits directly under "Direct/Indirect Expenses" — edge case
            main_group = sub_group = sub_sub_group = direct_parent
        elif depth == 1:
            main_group = sub_group = sub_sub_group = path[0]
        elif depth == 2:
            main_group    = path[0]
            sub_group     = path[1]
            sub_sub_group = path[1]
        else:  # depth >= 3
            main_group    = path[0]
            sub_group     = path[1]
            sub_sub_group = path[2]

        result[ledger] = {
            "exp_type":      exp_type,
            "main_group":    main_group,
            "sub_group":     sub_group,
            "sub_sub_group": sub_sub_group,
        }

    return result


# ---------------------------------------------------------------------------
# Phase B: Voucher fetch
# ---------------------------------------------------------------------------

def _expense_vouchers_xml(company: str, from_date: str = None, to_date: str = None) -> str:
    """
    Fetch ALL vouchers for the company.
    Date filtering and expense classification are done in Python, BUT we also
    pass SVFROMDATE/SVTODATE so Tally scopes the Collection export to the
    requested period explicitly. Without this, Tally falls back to the
    company's own current/default reporting period (set via F2 in the Tally
    UI, or whatever the company was last opened with) - for an actively-used
    current-year company that period is usually wide enough to look correct
    by accident, but for an older "split" company file that's been archived,
    the default period can be narrow enough that almost nothing comes back
    even though the ledgers clearly show activity (verified via Trial Balance,
    which is a separate, period-aware report and unaffected by this).
    SVCURRENTCOMPANY must be an open company or Tally will hang indefinitely.
    """
    period_vars = ""
    if from_date and to_date:
        period_vars = f"""
        <SVFROMDATE>{from_date}</SVFROMDATE>
        <SVTODATE>{to_date}</SVTODATE>"""
    return f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>AllVouchers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>{period_vars}
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="AllVouchers" ISMODIFY="No">
            <TYPE>Voucher</TYPE>
            <FETCH>DATE,VOUCHERTYPENAME,VOUCHERNUMBER,PARTYNAME,NARRATION,ISOPTIONAL,ISVOID,ALLLEDGERENTRIES.LEDGERNAME,ALLLEDGERENTRIES.AMOUNT,ALLLEDGERENTRIES.ISDEEMEDPOSITIVE,ALLLEDGERENTRIES.ISPARTYLEDGER</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""


def _is_sales_voucher(vch_type: str) -> bool:
    """Exclude sales-side vouchers. All other types (payment, purchase, journal, etc.) pass."""
    vt = vch_type.lower()
    return "sale" in vt


def _build_employee_ledger_set(
    ledger_parent: dict[str, str],
    group_parent:  dict[str, str],
) -> set[str]:
    """
    Return all ledger names whose group hierarchy contains 'employ' or 'salary'.
    Used to auto-detect employee name from voucher ledger entries.
    """
    emp_keywords = ("employ", "salary", "salari")
    result = set()

    for ledger, direct_parent in ledger_parent.items():
        current = direct_parent
        visited = set()
        while current and current not in visited:
            if any(kw in current.lower() for kw in emp_keywords):
                result.add(ledger)
                break
            visited.add(current)
            current = group_parent.get(current, "")

    return result


def _parse_narration(narration: str) -> dict[str, str]:
    """
    Split pipe-delimited narration into structured fields.
    Expected format: "Description | SECTIONS | REMARK-1 | REMARK-2 | REMARK-3 | EMPLOYEE NAME"
    Parts that are absent are returned as empty strings.
    """
    parts = [p.strip() for p in narration.split("|")]

    def _get(idx):
        return parts[idx] if idx < len(parts) else ""

    return {
        "narration": _get(0),
        "sections":  _get(1) if "|" in narration else "",
        "remark1":   _get(2) if "|" in narration else "",
        "remark2":   _get(3) if "|" in narration else "",
        "remark3":   _get(4) if "|" in narration else "",
        "employee":  _get(5) if "|" in narration else "",
    }


# ---------------------------------------------------------------------------
# Narration intelligence — auto-extract SECTIONS, REMARKs from free text
# ---------------------------------------------------------------------------

_MON = r'(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)'


def _extract_period(text: str) -> str:
    """Return the most useful date/period string found in narration text."""
    u = text.upper()

    # DD.MM.YYYY TO DD.MM.YYYY date range
    m = re.search(r'(\d{1,2}\.\d{1,2}\.\d{4})\s+TO\s+(\d{1,2}\.\d{1,2}\.\d{4})', u)
    if m:
        return f"{m.group(1)} TO {m.group(2)}"

    # DT: or DT- followed by date
    m = re.search(r'\bDT[:\-]\s*(\d{1,2}[.\-]\d{1,2}[.\-]\d{2,4})', u)
    if m:
        return m.group(1)

    # ON DATED DD-MM-YYYY
    m = re.search(r'DATED\s+(\d{1,2}[.\-]\d{1,2}[.\-]\d{2,4})', u)
    if m:
        return m.group(1)

    # Standalone MMM-YY or MMM-YYYY  (word boundary, not part of a longer word)
    m = re.search(r'\b(' + _MON + r')[.\-](\d{2,4})\b', u)
    if m:
        yr = m.group(2)
        return f"{m.group(1)}-{yr[-2:]}"

    # "MONTH OF APR-26" / "M/O APR26" / "MONTH OF APRIL 2026"
    m = re.search(r'(?:MONTH\s+OF|M/O)\s+(' + _MON + r')[\s\-.]?(\d{2,4})', u)
    if m:
        return f"{m.group(1)}-{m.group(2)[-2:]}"

    return ""


def _clean_types_section(types: str) -> str:
    """Derive a concise SECTIONS label from the Tally ledger name (TYPES column)."""
    s = types.upper()
    s = re.sub(r'\s*@\s*\d+\.?\d*\s*%', '', s)          # remove @18%, @5% etc.
    s = re.sub(r'\s*\(\s*\d+\.?\d*\s*%\s*\)', '', s)     # remove (18%), (5%) etc.
    s = re.sub(r'\s*\(NON-GST\)', '', s)
    s = re.sub(r'\s*\(GST\)', '', s)
    s = re.sub(r'\s*\(RCM\)', '', s)          # remove (RCM) marker
    s = re.sub(r'\s+ACCOUNTS?\s*$', '', s)               # trailing ACCOUNTS / ACCOUNT
    s = re.sub(r'\s+A/C\s*$', '', s)                     # trailing A/C
    s = re.sub(r'\s*-\s*DIRECTOR\s*$', '', s)            # -DIRECTOR qualifier
    s = re.sub(r'\s*-\s*INT(?:ERNATIONAL|\.)\s*$', '', s)
    s = re.sub(r'\bEXPENSE\b', 'EXP', s)                # normalise EXPENSE → EXP
    return s.strip()


_NON_PERSON_WORDS = frozenset({
    'THE', 'A', 'AN', 'IN', 'INV', 'V', 'NO', 'PERIOD', 'MONTH',
    'BEING', 'THIS', 'THAT', 'WHICH', 'ANY', 'DT', 'DATED', 'AGAINST',
    'AGAINTS', 'THEIR', 'ITS', 'OUR', 'YOUR',
})


def _extract_employee_from_narration(narration: str) -> str:
    """
    Extract employee/traveller name from free-text T&C narrations.

    Handles:
    - "BEING TRAVELLING {name} FROM ..."
    - "BEING [AIR] TICKET BOOK... FOR {name} TRAVELLING ..."
    - "... FOR {name}_..." or "... FOR {name}" at end of string
    """
    u = narration.strip().upper()

    # BEING TRAVELLING {name} FROM ...
    m = re.match(r'BEING\s+TRAVELLING\s+(.+?)\s+FROM\s+', u)
    if m:
        return m.group(1).strip()

    # BEING [AIR] TICKET BOOK... FOR {name} TRAVELLING
    m = re.match(r'BEING\s+(?:AIR\s+)?TICKET\s+BOOK\w*\s+FOR\s+(.+?)\s+TRAVELLING\s', u)
    if m:
        return m.group(1).strip()

    # FOR {name}_ or FOR {name} at end — scan all occurrences, pick first valid person
    for m in re.finditer(r'\bFOR\s+([A-Z][A-Z\s()]{2,50}?)(?=\s*(?:_|\Z))', u):
        candidate = m.group(1).strip()
        words = candidate.split()
        if words and words[0] not in _NON_PERSON_WORDS and len(candidate) >= 4:
            return candidate

    return ""


_GROUP_EMPLOYEE_KEYWORDS = (
    "STAFF", "EMPLOYEES", "TEAM MEMBERS", "VARIOUS EMPLOYEE",
    "MULTIPLE EMPLOYEE", "ALL EMPLOYEE", "STAFF MEMBERS",
)


def _is_group_employee(text: str) -> bool:
    """
    Detect narrations/employee fields that reference more than one employee.
    Used to label EMPLOYEE NAME as "GROUP EMPLOYEE" instead of a single name.
    """
    if not text:
        return False
    u = text.strip().upper()
    if any(kw in u for kw in _GROUP_EMPLOYEE_KEYWORDS):
        return True
    # Two+ comma-separated name-like tokens: "RAJESH SHARMA, AMIT PATEL"
    if re.search(r'\b[A-Z][A-Z.]{2,}(?:\s+[A-Z][A-Z.]{2,}){0,3}\s*,\s*[A-Z][A-Z.]{2,}', u):
        return True
    # Two name-like tokens joined by "&" or "AND": "RAJESH & AMIT"
    if re.search(r'\b[A-Z][A-Z.]{2,}(?:\s+[A-Z][A-Z.]{2,}){0,3}\s+(?:&|AND)\s+[A-Z][A-Z.]{2,}(?:\s+[A-Z][A-Z.]{2,}){0,3}\b', u):
        return True
    return False


def _sections_from_keywords(u: str) -> str:
    """Return a keyword-driven SECTIONS label when narration clearly signals one."""
    if any(k in u for k in ('MAGAZINE', 'ADVERTORIAL', 'COLOUR ADVT', 'EDITION')):
        return "MAGAZINE EXP"
    if any(k in u for k in ('EXHIBITION', 'EXPO ')):
        return "EXHIBITION EXP"
    if any(k in u for k in ('STAMP PAPER', 'NOTARY', 'NOTRY', 'LEGAL PRINT', 'ADVOCATE', 'LEGAL EXP')):
        return "LEGAL EXP"
    if any(k in u for k in ('HOTEL', 'RESORT', 'ACCOMMODATION', 'ROOM NO')):
        return "HOTEL EXP"
    if any(k in u for k in ('SALARY BOOKED', 'SALARY FOR', 'SALARY OF')):
        return "SALARY EXP"
    if any(k in u for k in ('BREAKFAST', 'BRAKEFAST', 'LUNCH', 'DINNER', 'FOOD FOR', 'TEA', 'BEVERAGES')):
        return "STAFF WELFARE EXP"
    if any(k in u for k in ('LR.NO', 'LR NO', 'FREIGHT', 'LORRY', 'TRUCK')):
        return "FREIGHT EXP"
    if any(k in u for k in ('PETROL', 'DIESEL', 'FUEL SURCHARGE')):
        return "PETROL EXP"
    return ""


def _enrich_narration(narration_raw: str, types: str) -> dict[str, str]:
    """
    Intelligently derive SECTIONS, REMARK-1, REMARK-2, REMARK-3 from
    a free-text Tally narration. Called only when pipe-split fields are empty.

    Priority for SECTIONS: narration keyword > TYPES-derived label.
    REMARK-3 uses date/period pattern extraction.
    REMARK-1 / REMARK-2 use regex patterns covering the most common narration formats.
    """
    out = {"sections": "", "remark1": "", "remark2": "", "remark3": ""}
    nar = narration_raw.strip()
    if not nar or nar.upper() in ("REVERSED", "0", "-"):
        out["sections"] = _clean_types_section(types)
        return out

    u = nar.upper()

    # SECTIONS
    out["sections"] = _sections_from_keywords(u) or _clean_types_section(types)

    # REMARK-3: date / period
    out["remark3"] = _extract_period(nar)

    # ── Pattern A: BEING TRAVELLING {name} FROM {src} TO {dst} [_DT:...] ──
    m = re.match(
        r'BEING\s+TRAVELLING\s+(.+?)\s+FROM\s+(.+?)\s+TO\s+([^_,\n]+?)(?:\s*_DT[:\-]|\s*,\s*DT|$)',
        u)
    if m:
        out["remark1"] = m.group(1).strip()
        out["remark2"] = f"FROM {m.group(2).strip()} TO {m.group(3).strip().rstrip('_.')}"
        return out

    # ── Pattern B: BEING [AIR] TICKET BOOK[ED/ING] FOR {person} TRAVELLING FROM {src} TO {dst} ──
    m = re.match(
        r'BEING\s+(?:AIR\s+)?TICKET\s+BOOK\w*\s+FOR\s+(.+?)\s+TRAVELLING\s+FROM\s+(.+?)\s+TO\s+([^_,\n]+?)(?:\s*_DT[:\-]|$)',
        u)
    if m:
        out["remark1"] = m.group(1).strip()
        out["remark2"] = f"FROM {m.group(2).strip()} TO {m.group(3).strip().rstrip('_.')}"
        return out

    # ── Pattern B2: BEING [AIR] TICKET BOOK[ED/ING] FOR {person} TRAVELLING FROM {route} ──
    m = re.match(
        r'BEING\s+(?:AIR\s+)?TICKET\s+BOOK\w*\s+FOR\s+(.+?)\s+TRAVELLING\s+FROM\s+(.+?)(?:\s*_DT[:\-]|$)',
        u)
    if m:
        out["remark1"] = m.group(1).strip()
        out["remark2"] = f"FROM {m.group(2).strip().rstrip('_.')}"
        return out

    # ── Pattern B3: BEING TICKET BOOK FOR {airline} ... AS PER {person} ──
    m = re.match(r'BEING\s+TICKET\s+BOOK\s+FOR\s+(.+?)\s+AS\s+PER\s+(.+)', u)
    if m:
        out["remark1"] = m.group(2).strip()   # person
        out["remark2"] = m.group(1).strip()[:60]  # airline / details
        return out

    # ── Pattern C: BEING HOTEL {name} BOOKED FOR {person} [FOR THE DT:...] ──
    m = re.match(
        r'BEING\s+HOTEL\s+(.+?)\s+BOOKED\s+FOR\s+(.+?)(?:\s+FOR\s+THE\s+DT|\s*_DT[:\-]|$)',
        u)
    if m:
        out["remark1"] = m.group(2).strip()   # person
        out["remark2"] = m.group(1).strip()   # hotel name
        return out

    # ── Pattern D: LR freight — "LR[.]NO[.]:N _DT:date _ROUTE _VEHICLE _WEIGHT" ──
    m = re.match(r'LR\.?\s*NO\.?:?\s*(\S+)', u)
    if m:
        lr_no = m.group(1).split("_")[0].rstrip(",.")   # stop at first underscore
        parts = re.split(r'[_|]', nar)
        route = next(
            (p.strip() for p in parts if " TO " in p.upper() and not p.upper().strip().startswith("DT")),
            "")
        out["remark1"] = route.upper() if route else f"LR NO: {lr_no}"
        out["remark2"] = f"LR NO: {lr_no}"
        return out

    # ── Pattern E: BEING [PRE|RE]PAID TO {vendor} _ {exp_type} _ {date} ──
    m = re.match(r'BEING\s+(?:PRE|RE)PAID\s+TO\s+(.+?)(?:_|$)', u)
    if m:
        parts = [p.strip() for p in re.split(r'_', nar)]
        vendor_m = re.match(r'BEING\s+(?:PRE|RE)PAID\s+TO\s+(.+)', parts[0], re.I)
        out["remark1"] = vendor_m.group(1).strip().upper() if vendor_m else ""
        out["remark2"] = parts[1].strip().upper() if len(parts) > 1 else ""
        if len(parts) > 2 and not out["remark3"]:
            out["remark3"] = parts[2].strip().upper()
        return out

    # ── Pattern F: BEING AMT PAID {AGAINST|TO} {subject} FOR {purpose} ──
    m = re.match(r'BEING\s+AMT\s+PAID\s+(?:AGAINST|TO)\s+(.+?)\s+FOR\s+(.+)', u)
    if m:
        out["remark1"] = m.group(1).strip()
        out["remark2"] = m.group(2).strip()
        return out

    # ── Pattern G: BEING PAID TO {recipient} [_ / FOR ] {purpose} ──
    m = re.match(r'BEING\s+PAID\s+TO\s+(.+?)(?:_|\s+FOR\s+)(.*)', u)
    if m:
        out["remark1"] = m.group(1).strip()
        out["remark2"] = m.group(2).strip()
        return out

    # ── Pattern H: RTGS / NEFT bank reference ──
    m = re.match(r'(?:RTGS|NEFT)/[^/]*/[^/]*/([^/\n]+?)(?:/[^\s(]+)?\s*(?:\((.+?)\))?$', u)
    if m:
        out["remark1"] = m.group(1).strip()
        out["remark2"] = m.group(2).strip() if m.group(2) else ""
        return out

    # ── Pattern I: SALARY ──
    if re.match(r'SALARY\b', u):
        out["remark1"] = "SALARY"
        return out

    # ── Pattern J: plain MMM-YY month reference ──
    if re.match(r'^' + _MON + r'[.\-]\d{2,4}$', u.strip()):
        out["remark3"] = nar.strip().upper()
        return out

    # ── Pattern K: generic BEING {subject} [_ {detail}] ──
    m = re.match(r'BEING\s+(.+)', u)
    if m:
        parts = re.split(r'_', m.group(1), maxsplit=1)
        out["remark1"] = parts[0].strip()[:80]
        out["remark2"] = parts[1].strip()[:80] if len(parts) > 1 else ""
        return out

    # ── Fallback: short narration as REMARK-1 ──
    if len(nar) <= 80:
        out["remark1"] = u
    return out


# ---------------------------------------------------------------------------
# Main fetch function
# ---------------------------------------------------------------------------

def fetch_expense_rows(
    company:     str,
    company_cfg: dict,
    from_date:   str,
    to_date:     str,
) -> list[dict]:
    """
    Fetch all Direct + Indirect expense rows from Tally for one company.

    Args:
        company:     Exact Tally company name (must be confirmed open).
        company_cfg: {"company_col": str, "location": str, "company_name": str}
        from_date:   YYYYMMDD start date (inclusive).
        to_date:     YYYYMMDD end date (inclusive).

    Returns:
        List of row dicts with 23 keys matching the Expense Sheet column headers
        (includes DEPARTMENT and COMPONENT CATEGORY).
        DATE values are Python datetime objects (required for Excel MONTH formula).

    IMPORTANT: Only call for companies confirmed open in Tally.
    Sending SVCURRENTCOMPANY for a non-open company causes Tally to hang indefinitely.
    """
    try:
        dt_from = datetime.strptime(from_date, "%Y%m%d")
        dt_to   = datetime.strptime(to_date,   "%Y%m%d")
    except ValueError as e:
        raise ValueError(f"Invalid date format (expected YYYYMMDD): {e}") from e

    # Phase A: build expense hierarchy map + employee ledger set
    print(f"  Fetching ledger master...")
    ledger_parent = _fetch_ledger_parent(company)
    print(f"  Fetching group master...")
    group_parent  = _fetch_group_parent(company)
    exp_ledger_map = build_expense_ledger_map(ledger_parent, group_parent)
    emp_ledger_set = _build_employee_ledger_set(ledger_parent, group_parent)
    print(f"  Expense ledgers: {len(exp_ledger_map)}, Employee ledgers: {len(emp_ledger_set)}")

    # Phase B: fetch all vouchers
    print(f"  Fetching all vouchers (may take up to {TALLY_TIMEOUT}s)...")
    root = _post_tally(_expense_vouchers_xml(company, from_date, to_date), f"AllVouchers/{company[:30]}")
    all_vouchers = root.findall(".//VOUCHER")
    print(f"  Total vouchers in company: {len(all_vouchers)}")

    rows             = []
    skipped_optional = 0
    skipped_date     = 0
    skipped_sales    = 0
    skipped_no_exp   = 0

    for el in all_vouchers:
        # Skip optional (provisional) and voided vouchers
        if (_txt(el, "ISOPTIONAL").upper() == "YES" or
                _txt(el, "ISVOID").upper() == "YES"):
            skipped_optional += 1
            continue

        # Date filter
        date_raw = _txt(el, "DATE")
        try:
            dt = datetime.strptime(date_raw, "%Y%m%d")
        except ValueError:
            skipped_date += 1
            continue
        if not (dt_from <= dt <= dt_to):
            skipped_date += 1
            continue

        # Exclude sales vouchers
        vch_type = _txt(el, "VOUCHERTYPENAME")
        if _is_sales_voucher(vch_type):
            skipped_sales += 1
            continue

        vch_no        = _txt(el, "VOUCHERNUMBER")
        party         = _txt(el, "PARTYNAME")
        narration_raw = _txt(el, "NARRATION")
        nar           = _parse_narration(narration_raw)

        # Fallback party: first ledger entry with ISPARTYLEDGER=Yes
        if not party:
            for le in el.findall(".//ALLLEDGERENTRIES.LIST"):
                if _txt(le, "ISPARTYLEDGER").upper() == "YES":
                    candidate = _txt(le, "LEDGERNAME")
                    if candidate:
                        party = candidate
                        break

        # Employee name — priority order:
        # 1. PARTY NAME ends with "(EXP)" → it's an employee expense ledger, use party
        # 2. Any ledger entry name ends with "(EXP)" → use that ledger name
        # 3. PARTY NAME is in the employee ledger set (e.g. DIPAK VEDU PATIL) → use party
        # 4. Narration pipe position 5 fallback
        # 5. Free-text extraction from T&C-style narrations ("BEING TRAVELLING {name} FROM...", "FOR {name}_...")
        # 6. If the narration/resolved name references more than one employee, use "GROUP EMPLOYEE"
        employee_name = ""
        if party and party.lower().endswith("(exp)"):
            employee_name = party
        else:
            for le in el.findall(".//ALLLEDGERENTRIES.LIST"):
                lname = _txt(le, "LEDGERNAME")
                if lname and lname.lower().endswith("(exp)"):
                    employee_name = lname
                    break
            if not employee_name and party and party in emp_ledger_set:
                employee_name = party
        if not employee_name:
            employee_name = nar["employee"]
        if not employee_name:
            employee_name = _extract_employee_from_narration(narration_raw)
        if _is_group_employee(narration_raw) or (employee_name and _is_group_employee(employee_name)):
            employee_name = "GROUP EMPLOYEE"

        # Tally sign convention differs by voucher family:
        #
        #   DIRECT expenses  (Purchase / Journal vouchers):
        #       debit entry = POSITIVE amount in XML  →  use amount as-is
        #       credit/reversal = NEGATIVE amount     →  also as-is (shows negative in report)
        #
        #   INDIRECT expenses  (Payment / GST Inward Service vouchers):
        #       debit entry = NEGATIVE amount in XML  →  negate to get positive
        #       credit/reversal = POSITIVE amount     →  negate to get negative
        #
        # This matches Tally P&L exactly: debit-balance → positive, credit-balance → negative.

        voucher_exp_rows = []
        for le in el.findall(".//ALLLEDGERENTRIES.LIST"):
            lname = _txt(le, "LEDGERNAME")
            if not lname or lname not in exp_ledger_map:
                continue
            amount = _parse_number(_txt(le, "AMOUNT"))
            if amount == 0:
                continue

            # ISDEEMEDPOSITIVE: "No" = debit side (expense increases) → positive
            #                   "Yes" = credit side (expense decreases) → negative
            idp_raw    = _txt(le, "ISDEEMEDPOSITIVE").strip().upper()
            idp_is_no  = (idp_raw == "NO" or idp_raw == "")   # missing → treat as debit

            if idp_raw in ("YES", "NO"):
                # Use IDP directly: debit (No) → positive, credit (Yes) → negative
                net_amount = abs(amount) if idp_is_no else -abs(amount)
            else:
                # Fallback: ISDEEMEDPOSITIVE not returned by this Tally version
                info_fb    = exp_ledger_map[lname]
                net_amount = amount if info_fb["exp_type"] == "DIRECT" else -amount

            if net_amount == 0:
                continue

            info   = exp_ledger_map[lname]
            enrich = _enrich_narration(narration_raw, lname)

            # Pipe-split fields take priority; enriched fields fill in the blanks;
            # anything still blank falls back to the ledger name (TYPES) so no cell is empty
            row_ledger_name = party or lname
            row_narration   = nar["narration"] or lname
            row_sections    = nar["sections"] or enrich["sections"] or lname
            row_remark1     = nar["remark1"]  or enrich["remark1"]  or lname
            row_remark2     = nar["remark2"]  or enrich["remark2"]  or lname
            row_remark3     = nar["remark3"]  or enrich["remark3"]  or lname

            component_category = (
                classify_component_category(lname, row_ledger_name)
                if info["exp_type"] == "DIRECT" else None
            )

            voucher_exp_rows.append({
                "Company":                    company_cfg["company_col"],
                "EXP TYPE":                   info["exp_type"],
                "MAIN GROUP":                 info["main_group"],
                "SUB-GROUP":                  info["sub_group"],
                "SUB-SUB GROUP":              info["sub_sub_group"],
                "TYPES":                      lname,
                "DATE":                       dt,
                "LEDGER NAME":                row_ledger_name,
                "VCH. TYPE":                  vch_type,
                "VCH NO.":                    vch_no,
                "NET AMOUNT":                 net_amount,
                "MONTH":                      None,
                "LOCATION":                   company_cfg["location"],
                "NARRATION":                  row_narration,
                "SECTIONS":                   row_sections,
                "REMARK-1":                   row_remark1,
                "REMARK-2":                   row_remark2,
                "REMARK-3 (Specified period)": row_remark3,
                "EMPLOYEE NAME":              employee_name,
                "DEPARTMENT":                 EMPLOYEE_DEPARTMENT.get(employee_name, ""),
                "Financial Quarter":          None,
                "Company Name":               company_cfg["company_name"],
                "COMPONENT CATEGORY":         component_category,
            })

        if voucher_exp_rows:
            rows.extend(voucher_exp_rows)
        else:
            skipped_no_exp += 1

    print(f"  Expense rows: {len(rows)} "
          f"(skipped: optional/void={skipped_optional}, out-of-range={skipped_date}, "
          f"sales={skipped_sales}, no-exp-entry={skipped_no_exp})")
    return rows


# ---------------------------------------------------------------------------
# CLI test
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fetch Tally expense vouchers (test mode)")
    parser.add_argument("--company", required=True, help="Exact company name in Tally")
    parser.add_argument("--from", dest="from_date", required=True, help="YYYYMMDD")
    parser.add_argument("--to",   dest="to_date",   required=True, help="YYYYMMDD")
    args = parser.parse_args()

    print(f"Open companies: {get_open_companies()}\n")
    print(f"Fetching: {args.company}")
    print(f"Period  : {args.from_date} to {args.to_date}\n")

    cfg = {"company_col": "Test", "location": "SURAT", "company_name": "Test"}
    rows = fetch_expense_rows(args.company, cfg, args.from_date, args.to_date)

    print(f"\nTotal expense rows: {len(rows)}")
    for r in rows[:10]:
        print(f"  {r['DATE'].strftime('%d-%m-%Y')} | {r['EXP TYPE']:8s} | "
              f"{r['MAIN GROUP'][:30]:30s} | {r['TYPES'][:30]:30s} | "
              f"{r['NET AMOUNT']:>12,.2f}")
