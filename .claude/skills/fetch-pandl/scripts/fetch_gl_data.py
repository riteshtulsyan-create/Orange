"""
Fetches the ledger-level P&L / Balance Sheet CSVs that build_gl_detail_sidebyside.py and
build_mom_gl_detail.py consume, straight from Tally Prime over HTTP-XML.

WHY THIS EXISTS
---------------
The fetch-pandl SKILL.md originally documented this refresh as a sequence of tally-prime **MCP**
tool calls (set-company / set-period / chart-of-accounts / profit-loss / balance-sheet /
query-database). That MCP server is not always connected to a given session - when it isn't, the
documented Step A is simply not runnable and the whole CFO-mode procedure stalls. This script is
the no-MCP path: it speaks Tally's native XML gateway directly (same transport fetch_pandl.py
uses) and writes byte-compatible CSVs into gl_data/ and gl_data_mom_detail/.

PERFORMANCE - READ BEFORE CHANGING THE QUERY STRATEGY
----------------------------------------------------
Do NOT fetch balances via a `<COLLECTION><TYPE>Ledger</TYPE>` with a ClosingBalance
NATIVEMETHOD. That asks Tally to walk every voucher for every ledger from the company's own book
start (Colorix's books open 01-Apr-2020), and it is pathologically slow: a single such call for
Colorix - the SMALLEST of the five companies - ran >20 minutes at 100% CPU without returning, and
because Tally serialises XML requests on one thread it blocked every subsequent request until the
process was restarted (hit for real 2026-08-03). Use the native **Trial Balance** report instead
(see trial_balance_xml below): Tally's own report engine aggregates this efficiently.

Master data (group names/parents, ledger names/parents) has no balance computation in it and IS
cheap to pull as a plain collection - that part is fine.

CSV CONTRACT (must match what load_csv() in build_gl_detail_sidebyside.py expects)
----------------------------------------------------------------------------------
Columns, in order: primary_group, group_name, ledger_name, closing_balance

  primary_group  the ledger's nearest RESERVED (Tally-predefined) ancestor group, found by
                 walking up the Parent chain - e.g. a ledger under 'CREDITOR FOR OTHER' (user
                 group) under 'Sundry Creditors' (reserved) gets 'Sundry Creditors'. Tally's own
                 $PrimaryGroup method returns empty over XML, hence the manual walk. Ledgers with
                 no reserved ancestor at all (notably 'Profit & Loss A/c') get "" - which
                 BS_LIABILITY_ORDER deliberately contains, mapping to 'Profit & Loss A/c
                 (current)' via BS_GROUP_LABEL.
  group_name     the ledger's IMMEDIATE parent group, as Tally stores it (e.g. 'CREDITOR FOR
                 OTHER', 'RELATED PARTY PURCHASE'). Feeds the 'Sub Group' column and the Related
                 Party / Finance Cost detection.
  closing_balance  Tally sign convention, preserved as-is: assets and expenses NEGATIVE,
                 liabilities and income POSITIVE. Trial Balance returns separate Dr/Cr amounts;
                 this script emits Cr - Dr, which reproduces that convention.

Rows with closing_balance == 0 are dropped (the original query had WHERE closing_balance <> 0).

USAGE
-----
    python fetch_gl_data.py --probe                 # dump raw TB XML for one company, no writes
    python fetch_gl_data.py --validate              # refetch Q1 and diff vs existing CSVs
    python fetch_gl_data.py --periods qtr,jul,ytd   # fetch specific periods, all 5 companies
    python fetch_gl_data.py --all                   # everything: apr may jun qtr jul ytd
    python fetch_gl_data.py --company COLORIX --periods jul

Periods are defined in PERIODS below. Anything writing into gl_data_mom_detail/ is a monthly
flow (P&L) or an as-at position (BS); 'qtr' and 'ytd' write the cumulative files in gl_data/.
"""

import argparse
import csv
import os
import re
import sys
import time
import urllib.error
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
SKILL_DIR = SCRIPT_DIR.parent
GL_DATA_DIR = SKILL_DIR / "gl_data"
MOM_DIR = SKILL_DIR / "gl_data_mom_detail"

# ---------------------------------------------------------------------------
# Connection - .env at the project root is the single source of truth (CLAUDE.md).
# ---------------------------------------------------------------------------
def _load_env():
    for cand in (Path.cwd() / ".env", SCRIPT_DIR / ".env", SKILL_DIR / ".env",
                 SKILL_DIR.parent.parent / ".env"):
        if cand.exists():
            for line in cand.read_text().splitlines():
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip())
            return cand
    return None


_ENV_PATH = _load_env()
TALLY_HOST = os.environ.get("TALLY_HOST", "localhost")
TALLY_PORT = int(os.environ.get("TALLY_PORT", "9000"))

# ---------------------------------------------------------------------------
# The 5 companies - exact Tally names. Keep in lockstep with
# build_gl_detail_sidebyside.ALL_COMPANIES; a shortened name can silently resolve to a
# different company variant in Tally (see SKILL.md warning).
# ---------------------------------------------------------------------------
COMPANIES = {
    "ORANGE_O_TEC_PRIVATE_LIMITED": "ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)",
    "PRIVATE_LIMITED_NOIDA": "ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)",
    "ENTERPRISES_PVT_LTD": "ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)",
    "ENTERPRISES_NOIDA": "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27",
    "COLORIX": "COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)",
}

FY_START = "20260401"

# period key -> (pl_from, pl_to, bs_from, bs_to, destination dir, filename infix)
# P&L is a FLOW: scoped to the period itself. BS is a POSITION: always from FY start to the
# period end, so the embedded 'Profit & Loss A/c (current)' line is cumulative (SKILL.md, "MOM
# file naming convention").
PERIODS = {
    "apr": (("20260401", "20260430"), (FY_START, "20260430"), MOM_DIR, "apr"),
    "may": (("20260501", "20260531"), (FY_START, "20260531"), MOM_DIR, "may"),
    "jun": (("20260601", "20260630"), (FY_START, "20260630"), MOM_DIR, "jun"),
    "jul": (("20260701", "20260731"), (FY_START, "20260731"), MOM_DIR, "jul"),
    # Cumulative files consumed as "the quarter" / "the YTD period" by the build scripts.
    "qtr": ((FY_START, "20260630"), (FY_START, "20260630"), GL_DATA_DIR, None),
    "ytd": ((FY_START, "20260731"), (FY_START, "20260731"), GL_DATA_DIR, "ytd"),
}

PL_PRIMARY_GROUPS = {
    "Sales Accounts", "Purchase Accounts", "Direct Expenses",
    "Direct Incomes", "Indirect Expenses", "Indirect Incomes",
}


# ---------------------------------------------------------------------------
# Transport
# ---------------------------------------------------------------------------
def _ok_cp(cp: int) -> bool:
    return cp in (0x9, 0xA, 0xD) or 0x20 <= cp <= 0xD7FF or 0xE000 <= cp <= 0xFFFD


def sanitize(s: str) -> str:
    """Tally emits raw control characters and invalid numeric character references (e.g. &#4;)
    inside ledger names, plus bare '&'. All three make a standards-compliant XML parser throw
    'reference to invalid character number' - strip/repair them before ET.fromstring."""
    s = re.sub(r"&#x([0-9a-fA-F]+);", lambda m: m.group(0) if _ok_cp(int(m.group(1), 16)) else "", s)
    s = re.sub(r"&#(\d+);", lambda m: m.group(0) if _ok_cp(int(m.group(1))) else "", s)
    s = re.sub(r"&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)", "&amp;", s)
    return "".join(c for c in s if _ok_cp(ord(c)))


def decode(raw: bytes) -> str:
    """Tally replies UTF-16LE or UTF-8 depending on its own export settings."""
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16", errors="replace")
    if len(raw) > 4 and raw[1] == 0 and raw[3] == 0:
        return raw.decode("utf-16-le", errors="replace")
    return raw.decode("utf-8", errors="replace")


def post(xml: str, timeout: int = 240) -> str:
    req = urllib.request.Request(
        f"http://{TALLY_HOST}:{TALLY_PORT}",
        data=xml.encode("utf-8"),
        headers={"Content-Type": "text/xml; charset=utf-8"},
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return sanitize(decode(r.read()))
    except urllib.error.URLError as e:
        raise RuntimeError(
            f"Cannot reach Tally at {TALLY_HOST}:{TALLY_PORT} ({e}).\n"
            f"  - Is Tally Prime running with all 5 companies open?\n"
            f"  - Check F1 > Settings > Connectivity for the actual port, and that it matches\n"
            f"    TALLY_PORT in {_ENV_PATH}."
        ) from e


# ---------------------------------------------------------------------------
# Requests
# ---------------------------------------------------------------------------
def _collection(name: str, ctype: str, methods, company: str) -> str:
    fetch = "\n".join(f"        <NATIVEMETHOD>{m}</NATIVEMETHOD>" for m in methods)
    return f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>{name}</ID></HEADER>
  <BODY><DESC>
    <STATICVARIABLES>
      <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
    </STATICVARIABLES>
    <TDL><TDLMESSAGE>
      <COLLECTION ISMODIFY="No" NAME="{name}">
        <TYPE>{ctype}</TYPE>
{fetch}
      </COLLECTION>
    </TDLMESSAGE></TDL>
  </DESC></BODY>
</ENVELOPE>"""


def groups_xml(company: str) -> str:
    return _collection("GrpMast", "Group", ("Name", "Parent", "IsReserved"), company)


def ledgers_xml(company: str) -> str:
    # Name + Parent only. Adding ClosingBalance here is what makes it unusably slow - see the
    # PERFORMANCE note in this module's docstring.
    return _collection("LedMast", "Ledger", ("Name", "Parent"), company)


def trial_balance_xml(company: str, from_d: str, to_d: str) -> str:
    """Tally's native Trial Balance at LEDGER level.

    **ISLEDGERWISE is the flag that matters** - without it Tally returns only the ~32 group
    totals, and none of the ledger rows this project needs. EXPLODEFLAG / SVEXPANDALLLEVELS /
    SVDETAILED / a REPORT ISMODIFY 'ExplodeFlag: Yes' were all tried and all silently returned
    group level anyway (measured 2026-08-03); ISLEDGERWISE was the only one that worked. It is
    also fast - ~0.3s per company/period vs >20 minutes for the ledger-collection approach the
    module docstring warns about."""
    return f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Data</TYPE><ID>Trial Balance</ID></HEADER>
  <BODY><DESC>
    <STATICVARIABLES>
      <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
      <SVFROMDATE TYPE="Date">{from_d}</SVFROMDATE>
      <SVTODATE TYPE="Date">{to_d}</SVTODATE>
      <ISLEDGERWISE>Yes</ISLEDGERWISE>
      <EXPLODEFLAG>Yes</EXPLODEFLAG>
      <ISITEMWISE>No</ISITEMWISE>
    </STATICVARIABLES>
  </DESC></BODY>
</ENVELOPE>"""


# ---------------------------------------------------------------------------
# Master data -> primary_group resolution
# ---------------------------------------------------------------------------
def fetch_group_tree(company: str):
    """Returns {group_name: (parent, is_reserved)}."""
    root = ET.fromstring(post(groups_xml(company), timeout=120))
    tree = {}
    for g in root.iter("GROUP"):
        name = (g.get("NAME") or g.findtext("NAME", "") or "").strip()
        if not name:
            continue
        tree[name] = (
            (g.findtext("PARENT", "") or "").strip(),
            (g.findtext("ISRESERVED", "") or "").strip().lower() in ("yes", "1", "true"),
        )
    return tree


def fetch_ledger_parents(company: str):
    """Returns {ledger_name: parent_group}."""
    root = ET.fromstring(post(ledgers_xml(company), timeout=240))
    out = {}
    for l in root.iter("LEDGER"):
        name = (l.get("NAME") or l.findtext("NAME", "") or "").strip()
        if name:
            out[name] = (l.findtext("PARENT", "") or "").strip()
    return out


def resolve_primary_group(parent: str, group_tree: dict) -> str:
    """Walk up the group Parent chain to the nearest RESERVED (Tally-predefined) group.

    Validated against the known-good gl_data/bs_COLORIX.csv:
      BANK CHARGES -> FINANCE COST (user) -> Indirect Expenses (reserved)  => 'Indirect Expenses'
      EXHIBITION EXP -> PRELIMINARY ASSETS (user) -> Current Assets (res)  => 'Current Assets'
      AXIS BANK -> Bank Account (reserved, note THIS Tally names it singular) => 'Bank Account'
    Returns "" when no reserved ancestor exists (e.g. 'Profit & Loss A/c', whose parent is
    'Primary'); BS_LIABILITY_ORDER contains "" for exactly this case.
    """
    seen = set()
    cur = parent
    while cur and cur not in seen and cur != "Primary":
        seen.add(cur)
        node = group_tree.get(cur)
        if node is None:
            return cur  # unknown group: best effort, keep the name we have
        par, reserved = node
        if reserved:
            return cur
        cur = par
    return ""


# ---------------------------------------------------------------------------
# Trial Balance parsing
# ---------------------------------------------------------------------------
def _amt(text: str) -> float:
    text = (text or "").strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def parse_trial_balance(xml_text: str):
    """Returns {ledger_name: closing_balance} using Tally's own sign convention
    (Cr positive, Dr negative), matching the existing CSVs.

    Tally's Trial Balance XML nests ledger rows as DSPACCNAME/DSPACCINFO pairs. Debit and credit
    come back in separate elements (DSPCLDRAMT / DSPCLCRAMT); some builds instead emit a single
    signed DSPCLAMTA. Both shapes are handled."""
    root = ET.fromstring(xml_text)
    out = {}
    # DSPACCNAME and DSPACCINFO are FLAT SIBLINGS in document order, not nested in a per-row
    # container:  <DSPACCNAME/><DSPACCINFO/><DSPACCNAME/><DSPACCINFO/>...
    # So each name must be paired with the DSPACCINFO that immediately FOLLOWS it. Using
    # find("DSPACCINFO") on the shared parent instead returns the document's first DSPACCINFO
    # for every ledger - which silently assigns one identical balance to all of them (caught by
    # --validate on 2026-08-03: every Colorix ledger came back as 105,648.00).
    for container in root.iter():
        children = list(container)
        for i, child in enumerate(children):
            if child.tag != "DSPACCNAME":
                continue
            name = (child.findtext("DSPDISPNAME", "") or "").strip()
            if not name:
                continue
            info = children[i + 1] if i + 1 < len(children) and children[i + 1].tag == "DSPACCINFO" else None
            if info is None:
                continue
            # Tally reports the DEBIT amount ALREADY NEGATIVE in DSPCLDRAMTA and the credit
            # positive in DSPCLCRAMTA, so the closing balance is their plain SUM - not cr - dr.
            # (Seen directly in the group-level TB: Capital Account = Dr -123428.38 + Cr 50000.00
            # = -73428.38, which is Reserves & Surplus plus Share Capital.) Summing them this way
            # yields Tally's native convention that the CSV contract requires: assets/expenses
            # negative, liabilities/income positive.
            dr = _amt(info.findtext(".//DSPCLDRAMTA", ""))
            cr = _amt(info.findtext(".//DSPCLCRAMTA", ""))
            bal = (dr + cr) if (dr or cr) else _amt(info.findtext(".//DSPCLAMTA", ""))
            # A ledger can repeat across groups in an exploded TB; keep the first occurrence,
            # matching load_csv()'s own duplicate guard.
            if name not in out:
                out[name] = bal
    return out


# ---------------------------------------------------------------------------
# CSV writing
# ---------------------------------------------------------------------------
# Tally's own inventory pseudo-ledgers. They come back with NO parent group at all, so they
# cannot be recognised by group; match them by name and stamp the group_name that load_csv()
# already filters on. This project injects its own stock figures from the STOCK / STOCK_MOM
# dicts instead (stock is not a ledger - see SKILL.md "Stock, and the manually-confirmed
# adjustments"), so letting these through would double-count.
STOCK_PSEUDO_LEDGERS = {"opening stock", "closing stock"}
STOCK_IN_HAND_GROUP = "Stock-in-Hand"

PL_LEDGER_NAME = "Profit & Loss A/c"

# Tally's COMPUTED Balance Sheet lines. A ledger-wise Trial Balance surfaces these even though
# they are not ledger masters at all (confirmed: absent from the Ledger collection, and they
# carry no parent group), which is why SKILL.md describes them as visible only on Tally's own
# report screens. They must NOT go into the CSVs:
#   * with no parent group they resolve to primary_group "" and would be filed under
#     'Profit & Loss A/c (current)' - the wrong bucket entirely; and
#   * load_company_rows() already injects each one into the correct bucket on BOTH statements
#     from SALES_BILLS_TO_MAKE / PURCHASE_BILLS_TO_COME / UNADJUSTED_FOREX_GAIN_LOSS, whereas
#     the raw row only ever carries the Balance Sheet side. Letting them through would
#     double-count the BS side and misplace it.
# Their live values ARE worth having though - report_computed_lines() below prints them, which
# is how the per-period constants should now be set (read them off Tally directly) instead of
# back-solving from a Reconciliation Gap or reading a screenshot.
COMPUTED_BS_LINES = {
    "sales bills to make",
    "purchase bills to come",
    "unadjusted forex gain/loss",
}


def build_rows(balances: dict, ledger_parents: dict, group_tree: dict):
    """Splits ledger balances into (pl_rows, bs_rows) in CSV column order.

    Two semantics are reconciled here so the output matches what the MCP profit-loss /
    balance-sheet tools used to produce:

    * Stock pseudo-ledgers are re-tagged into the 'Stock-in-Hand' group so load_csv() drops them.
    * 'Profit & Loss A/c' is rewritten from its Trial Balance value (which is only the OPENING
      balance brought forward) to its Balance Sheet value (opening + profit earned so far in the
      period). Tally's BS shows the latter, and the workbook's verification checklist compares
      the P&L schedule's profit against this very line, so it has to be the BS figure. Verified
      on Colorix Q1: TB opening -517,012.98 + period P&L 4,372,217.31 = 3,855,204.33.
    """
    pl_rows, bs_rows = [], []
    period_pl_total = 0.0

    for ledger, bal in balances.items():
        if ledger.strip().lower() in COMPUTED_BS_LINES:
            continue  # not a ledger - injected by load_company_rows() instead, see above
        parent = ledger_parents.get(ledger, "")
        if ledger.strip().lower() in STOCK_PSEUDO_LEDGERS and not parent:
            parent = STOCK_IN_HAND_GROUP
        primary = resolve_primary_group(parent, group_tree)
        is_pl = primary in PL_PRIMARY_GROUPS
        if is_pl:
            period_pl_total += bal
        if abs(bal) < 0.005:
            continue
        row = (primary, parent, ledger, round(bal, 2))
        (pl_rows if is_pl else bs_rows).append(row)

    # Restate the P&L A/c line onto the Balance Sheet basis (see docstring). This row must be
    # emitted even when the Trial Balance has no P&L A/c entry - a company whose accumulated P&L
    # brought forward is zero (or whose row was dropped by the zero-balance filter above) still
    # has to carry the CURRENT period's profit on its Balance Sheet. Skipping it silently drops
    # the entire period profit from the BS: caught on 2026-08-03 when Enterprises Noida's
    # Reconciliation Gap came out at +77,39,626.16, exactly equal and opposite to its own Q1
    # profit of -77,39,626.16.
    for i, (primary, group_name, ledger, bal) in enumerate(bs_rows):
        if ledger.strip() == PL_LEDGER_NAME:
            bs_rows[i] = (primary, group_name, ledger, round(bal + period_pl_total, 2))
            break
    else:
        opening = balances.get(PL_LEDGER_NAME, 0.0)
        bs_rows.append(("", "", PL_LEDGER_NAME, round(opening + period_pl_total, 2)))

    pl_rows.sort(key=lambda r: r[2])
    bs_rows.sort(key=lambda r: r[2])
    return pl_rows, bs_rows


def write_csv(path: Path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["primary_group", "group_name", "ledger_name", "closing_balance"])
        for r in rows:
            w.writerow(r)


def csv_paths(period: str, key: str):
    _, _, dest, infix = PERIODS[period]
    if infix is None:
        return dest / f"pl_{key}.csv", dest / f"bs_{key}.csv"
    if dest is GL_DATA_DIR:
        return dest / f"pl_{infix}_{key}.csv", dest / f"bs_{infix}_{key}.csv"
    return dest / f"pl_{infix}_{key}.csv", dest / f"bs_{infix}_{key}.csv"


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------
def fetch_company_period(key: str, period: str, masters: dict, dry_run: bool = False):
    company = COMPANIES[key]
    (pl_from, pl_to), (bs_from, bs_to), _, _ = PERIODS[period]
    group_tree, ledger_parents = masters[key]

    t0 = time.time()
    pl_bal = parse_trial_balance(post(trial_balance_xml(company, pl_from, pl_to)))
    pl_rows, _ = build_rows(pl_bal, ledger_parents, group_tree)

    if (bs_from, bs_to) == (pl_from, pl_to):
        bs_rows = build_rows(pl_bal, ledger_parents, group_tree)[1]
    else:
        bs_bal = parse_trial_balance(post(trial_balance_xml(company, bs_from, bs_to)))
        bs_rows = build_rows(bs_bal, ledger_parents, group_tree)[1]

    pl_path, bs_path = csv_paths(period, key)
    if not dry_run:
        write_csv(pl_path, pl_rows)
        write_csv(bs_path, bs_rows)
    print(f"    {period:<4} {key:<32} pl={len(pl_rows):>4} bs={len(bs_rows):>4}  ({time.time()-t0:.1f}s)")
    return pl_rows, bs_rows


def load_masters(keys):
    masters = {}
    for key in keys:
        t0 = time.time()
        gt = fetch_group_tree(COMPANIES[key])
        lp = fetch_ledger_parents(COMPANIES[key])
        masters[key] = (gt, lp)
        print(f"    masters {key:<32} groups={len(gt):>4} ledgers={len(lp):>5}  ({time.time()-t0:.1f}s)")
    return masters


def _ddmmmyyyy(yyyymmdd: str) -> str:
    """20260731 -> 31-Jul-2026, the date form Tally's StockItem collection expects."""
    from datetime import datetime
    d = datetime.strptime(yyyymmdd, "%Y%m%d").date()
    return f"{d.day}-{d.strftime('%b')}-{d.year}"


def stock_value(company: str, as_at: str) -> float:
    """Total closing stock VALUE for a company as at a date.

    Closing stock is not a ledger (SKILL.md), so it never appears in the Trial Balance pull -
    it comes from the inventory valuation instead, and the build scripts carry it in the hardcoded
    STOCK / STOCK_MOM dicts. This reads the same figure Tally's Stock Summary screen shows, so
    those dicts can be refreshed from live data rather than left stale."""
    xml = f"""<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>AllStockItems</ID></HEADER>
  <BODY><DESC>
    <STATICVARIABLES>
      <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
      <SVFROMDATE TYPE="Date">{_ddmmmyyyy(FY_START)}</SVFROMDATE>
      <SVTODATE TYPE="Date">{_ddmmmyyyy(as_at)}</SVTODATE>
    </STATICVARIABLES>
    <TDL><TDLMESSAGE>
      <COLLECTION ISMODIFY="No" NAME="AllStockItems">
        <TYPE>StockItem</TYPE>
        <NATIVEMETHOD>Name</NATIVEMETHOD>
        <NATIVEMETHOD>ClosingValue</NATIVEMETHOD>
      </COLLECTION>
    </TDLMESSAGE></TDL>
  </DESC></BODY>
</ENVELOPE>"""
    root = ET.fromstring(post(xml, timeout=240))
    total = 0.0
    for it in root.iter("STOCKITEM"):
        total += _amt(it.findtext("CLOSINGVALUE", ""))
    return total


def report_stock(keys, periods):
    """Closing stock value per company per period, next to whatever STOCK_MOM currently holds.
    Apr/May/Jun are the control: they were confirmed previously, so if live matches there, the
    July figure from the same query can be trusted too."""
    sys.path.insert(0, str(SCRIPT_DIR))
    import build_mom_gl_detail as m

    print("\nClosing stock value (live) vs STOCK_MOM currently in build_mom_gl_detail.py:")
    print(f"  {'company':<32} {'period':<6} {'LIVE':>18} {'IN SCRIPT':>18}   delta")
    print("  " + "-" * 84)
    for key in keys:
        for period in periods:
            _, (_, bs_to), _, _ = PERIODS[period]
            live = stock_value(COMPANIES[key], bs_to)
            cur = m.STOCK_MOM.get(key, {}).get(period)
            if cur is None:
                print(f"  {key:<32} {period:<6} {live:>18,.2f} {'-':>18}")
            else:
                d = live - cur
                flag = "OK" if abs(d) < 1 else f"{d:+,.2f}"
                print(f"  {key:<32} {period:<6} {live:>18,.2f} {cur:>18,.2f}   {flag}")


def report_computed_lines(keys, periods):
    """Prints Tally's own computed Balance Sheet lines (Sales Bills to Make / Purchase Bills to
    Come / Unadjusted Forex Gain/Loss) per company/period.

    These are the figures behind SALES_BILLS_TO_MAKE / PURCHASE_BILLS_TO_COME /
    UNADJUSTED_FOREX_GAIN_LOSS in build_gl_detail_sidebyside.py. SKILL.md documents them as
    obtainable only from a Tally drill-down screenshot or by back-solving a Reconciliation Gap;
    a ledger-wise Trial Balance exposes them directly, so a new period's constants can simply be
    read off here.

    Values print SIGNED, exactly as Tally reports them. Do not take the magnitude: these items
    genuinely change direction between periods (Colorix's Unadjusted Forex Gain/Loss is
    +44,266.20 in April but -12,987.45 in July - a gain then a loss), and an earlier version of
    this report printed abs(), which put July on the wrong side of the balance sheet and left a
    25,974.90 Reconciliation Gap. Sign conventions when copying into the dicts:
      * UNADJUSTED_FOREX_GAIN_LOSS  - store the sign AS SHOWN (written to liabilities as-is).
      * the other three dicts       - store the POSITIVE magnitude; the build code applies the
                                      negation itself."""
    print("\nTally computed BS lines (use these to set the per-period adjustment constants):")
    header = f"  {'company':<32} {'period':<6} " + " ".join(f"{n:>22}" for n in
                                                            ("Sales Bills to Make",
                                                             "Purchase Bills to Come",
                                                             "Unadj Forex Gain/Loss"))
    print(header)
    print("  " + "-" * (len(header) - 2))
    for key in keys:
        co = COMPANIES[key]
        for period in periods:
            _, (bs_from, bs_to), _, _ = PERIODS[period]
            bal = parse_trial_balance(post(trial_balance_xml(co, bs_from, bs_to)))
            lower = {k.strip().lower(): v for k, v in bal.items()}
            cells = []
            for name in ("sales bills to make", "purchase bills to come", "unadjusted forex gain/loss"):
                v = lower.get(name)
                cells.append(f"{v:>22,.2f}" if v is not None else f"{'-':>22}")
            print(f"  {key:<32} {period:<6} " + " ".join(cells))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--company", action="append", help="short key; repeatable. Default: all 5")
    ap.add_argument("--periods", default="", help="comma list: apr,may,jun,jul,qtr,ytd")
    ap.add_argument("--all", action="store_true", help="all periods")
    ap.add_argument("--probe", action="store_true", help="dump raw Trial Balance XML, no writes")
    ap.add_argument("--validate", action="store_true", help="refetch qtr and diff vs existing CSVs")
    ap.add_argument("--dry-run", action="store_true", help="fetch but do not write CSVs")
    ap.add_argument("--computed-lines", action="store_true",
                    help="report Tally's computed BS lines per period, for setting the constants")
    ap.add_argument("--stock", action="store_true",
                    help="report live closing stock value per period vs STOCK_MOM")
    args = ap.parse_args()

    keys = args.company or list(COMPANIES)
    for k in keys:
        if k not in COMPANIES:
            ap.error(f"unknown company {k!r}; choose from {list(COMPANIES)}")

    print(f"Tally {TALLY_HOST}:{TALLY_PORT}   (.env: {_ENV_PATH})")

    if args.probe:
        key = keys[0]
        xml = post(trial_balance_xml(COMPANIES[key], FY_START, "20260630"))
        out = SCRIPT_DIR / f"_probe_tb_{key}.xml"
        out.write_text(xml, encoding="utf-8")
        print(f"raw TB XML for {key}: {len(xml)} chars -> {out}")
        print(xml[:3000])
        return

    if args.stock:
        wanted = [p.strip() for p in args.periods.split(",") if p.strip()] or \
                 ["apr", "may", "jun", "jul"]
        for p in wanted:
            if p not in PERIODS:
                ap.error(f"unknown period {p!r}; choose from {list(PERIODS)}")
        report_stock(keys, wanted)
        return

    if args.computed_lines:
        wanted = [p.strip() for p in args.periods.split(",") if p.strip()] or \
                 ["apr", "may", "jun", "qtr", "jul", "ytd"]
        for p in wanted:
            if p not in PERIODS:
                ap.error(f"unknown period {p!r}; choose from {list(PERIODS)}")
        report_computed_lines(keys, wanted)
        return

    if args.validate:
        periods = ["qtr"]
    elif args.all:
        periods = ["apr", "may", "jun", "jul", "qtr", "ytd"]
    else:
        periods = [p.strip() for p in args.periods.split(",") if p.strip()]
    if not periods:
        ap.error("nothing to do: pass --periods, --all, --probe or --validate")
    for p in periods:
        if p not in PERIODS:
            ap.error(f"unknown period {p!r}; choose from {list(PERIODS)}")

    print("Loading master data (cheap - no balance computation)...")
    masters = load_masters(keys)

    print("Fetching Trial Balance per company/period...")
    results = {}
    for period in periods:
        for key in keys:
            results[(period, key)] = fetch_company_period(
                key, period, masters, dry_run=args.dry_run or args.validate)

    if args.validate:
        validate(results, keys)


def validate(results, keys):
    """Diffs a freshly-fetched 'qtr' against the existing known-good gl_data/ CSVs. A clean diff
    proves the direct-HTTP path reproduces what the MCP path produced, which is the gate before
    trusting it for July."""
    sys.path.insert(0, str(SCRIPT_DIR))
    import build_gl_detail_sidebyside as b

    print("\nValidation vs existing gl_data/ CSVs (via the real load_csv, so dedup/"
          "Stock-in-Hand/casing guards apply to both sides):")
    ok = True
    for key in keys:
        for stat, idx in (("pl", 0), ("bs", 1)):
            path = GL_DATA_DIR / f"{stat}_{key}.csv"
            if not path.exists():
                print(f"  {stat} {key}: no existing file to compare")
                continue
            old = {r[2]: r[3] for r in b.load_csv(path)}
            fresh_rows = results[("qtr", key)][idx]
            tmp = SCRIPT_DIR / f"_tmp_{stat}_{key}.csv"
            write_csv(tmp, fresh_rows)
            new = {r[2]: r[3] for r in b.load_csv(tmp)}
            tmp.unlink()

            only_old = set(old) - set(new)
            only_new = set(new) - set(old)
            changed = {k for k in set(old) & set(new) if abs(old[k] - new[k]) > 1.0}
            status = "OK" if not (only_old or only_new or changed) else "DIFF"
            if status == "DIFF":
                ok = False
            print(f"  [{status}] {stat} {key}: old={len(old)} new={len(new)} "
                  f"missing={len(only_old)} added={len(only_new)} changed={len(changed)}")
            for n in list(only_old)[:5]:
                print(f"        only in OLD: {n!r} = {old[n]:,.2f}")
            for n in list(only_new)[:5]:
                print(f"        only in NEW: {n!r} = {new[n]:,.2f}")
            for n in list(changed)[:5]:
                print(f"        changed: {n!r} {old[n]:,.2f} -> {new[n]:,.2f}")
    print("\nVALIDATION PASSED" if ok else "\nVALIDATION FAILED - do not trust this fetcher yet")


if __name__ == "__main__":
    main()
