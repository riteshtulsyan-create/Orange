"""Fetch the raw data for the GP Segment MoM report and cache it to JSON.

What is pulled per company, per period:

  1. StockItem -> opening / inward / outward / closing VALUE, kept at LEAF
     stock-group level (not pre-aggregated to a segment) so that mapping
     changes never require another Tally round-trip.
  2. Ledger    -> sales, purchase, discount and direct-cost ledgers, each
     classified as segment-tagged / untagged / inter-branch / related-party.

Why sales come from ledgers and not from items: item OutwardValue includes
branch transfers, godown moves and manufacturing issues. Measured against
Orange O Tec HO Q1 FY26-27 it overstates real sales by roughly 1.8x, so it is
used only for stock balances and mix, never as the revenue line.

Run:
    python .claude\\skills\\gp-segment-mom\\scripts\\fetch_gp_data.py
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parents[3]

sys.path.insert(0, str(PROJECT_ROOT / "Stock Report" / ".claude" / "skills" /
                      "fetch-stock-summary" / "scripts"))
from fetch_stock_summary import (  # noqa: E402
    _collection_xml, _safe_parse, post_to_tally, TALLY_HOST, TALLY_PORT,
)

sys.path.insert(0, str(HERE))
from segment_map import COMPANY_ENTITY  # noqa: E402


# ---------------------------------------------------------------------------
# parsing
# ---------------------------------------------------------------------------

def parse_value(raw) -> float:
    """Tally monetary field -> float.

    Handles the three shapes Tally actually returns:
      '-4140.00'                                  -> -4140.0
      ''            (item fully sold, means zero) ->     0.0
      '-$15500.00 @ ? 95/$ = -? 1472500.00'       -> -1472500.0   (forex)

    The forex form is a formatted string, not a number - the base-currency
    amount is whatever follows the final '='. Missing this made an early test
    report a 171% gross margin.
    """
    if raw is None:
        return 0.0
    s = str(raw).strip()
    if not s:
        return 0.0
    if "=" in s:
        s = s.rsplit("=", 1)[1]
    s = re.sub(r"[^0-9.\-]", "", s.replace(",", ""))
    if not s or s in ("-", ".", "-."):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _text(elem, tag: str) -> str:
    node = elem.find(tag)
    return (node.text or "").strip() if node is not None else ""


def _name(elem) -> str:
    return (elem.get("NAME") or _text(elem, "NAME")).strip()


# ---------------------------------------------------------------------------
# stock groups + items
# ---------------------------------------------------------------------------

def fetch_group_parents(company, from_date, to_date, host, port) -> dict:
    xml = _collection_xml(company, from_date, to_date,
                          "GPStockGroups", "StockGroup", "Name,Parent")
    root = _safe_parse(post_to_tally(xml, host, port))
    parents = {}
    for elem in root.iter("STOCKGROUP"):
        name = _name(elem)
        if not name:
            continue
        parent = _text(elem, "PARENT")
        if parent.lower() in ("primary", ""):
            parent = ""
        parents[name.upper()] = parent
    return parents


ITEM_FETCH = "Name,Parent,OpeningValue,ClosingValue,InwardValue,OutwardValue"


def fetch_period_items(company, from_date, to_date, host, port) -> dict:
    """Aggregate one company/period to LEAF stock group level."""
    xml = _collection_xml(company, from_date, to_date,
                          "GPItems", "StockItem", ITEM_FETCH)
    root = _safe_parse(post_to_tally(xml, host, port))

    by_group = {}
    count = 0
    for elem in root.iter("STOCKITEM"):
        if not _name(elem):
            continue
        count += 1
        leaf = _text(elem, "PARENT") or "(No stock group assigned)"
        g = by_group.setdefault(leaf, dict(opening=0.0, inward=0.0,
                                           outward=0.0, closing=0.0))
        # Sign convention confirmed live: inward/opening/closing are negative
        # (debit), outward positive (credit). Negate the former; keep outward
        # signed so a return-heavy month stays negative instead of being
        # flipped positive by a blanket abs().
        g["opening"] += -parse_value(_text(elem, "OPENINGVALUE"))
        g["closing"] += -parse_value(_text(elem, "CLOSINGVALUE"))
        g["inward"] += -parse_value(_text(elem, "INWARDVALUE"))
        g["outward"] += parse_value(_text(elem, "OUTWARDVALUE"))

    return {"by_group": by_group, "item_count": count}


# ---------------------------------------------------------------------------
# ledgers
# ---------------------------------------------------------------------------

_DISCOUNT_RE = re.compile(r"DISCOUNT|RATE DIFFERENCE", re.I)
_DIRECT_COST_RE = re.compile(
    r"CUSTOM DUTY|FREIGHT|CLEAR(ING|NING)|CARRIAGE|OCEAN|PACKING|"
    r"TOOLS & SPARE|LABOUR|SERVICE CHARGES|LC COST|DDP SHIPMENT", re.I)


def fetch_group_roots(company, from_date, to_date, host, port) -> dict:
    """{GROUP_UPPER: root group name} by walking Group.Parent to the top.

    Classifying by the group tree rather than by ledger name matters: a
    name-matching rule captured only 3.11 crore of direct expenses against the
    P&L's 4.15 crore, because it cannot know about ledgers whose names carry no
    recognisable keyword.
    """
    root = _safe_parse(post_to_tally(_collection_xml(
        company, from_date, to_date, "GPGroups", "Group", "Name,Parent"),
        host, port))
    parent = {}
    for elem in root.iter("GROUP"):
        name = _name(elem)
        if name:
            parent[name.upper()] = _text(elem, "PARENT").strip()

    def walk(group: str) -> str:
        seen = set()
        current = group
        while current:
            key = current.upper()
            if key in seen:
                return key
            seen.add(key)
            nxt = parent.get(key, "")
            if not nxt or nxt.upper() == key:
                return key
            current = nxt
        return (group or "").upper()

    return {g: walk(g) for g in parent}


def classify_ledger(name: str, parent: str, group_roots: dict = None) -> str | None:
    """Bucket a ledger: sales / purchase / discount / direct_cost / None.

    Discount and rate-difference ledgers sit UNDER Sales Accounts in Tally but
    behave as a reduction of revenue, so the name test runs first.
    """
    n = name.upper()
    if _DISCOUNT_RE.search(n):
        return "discount"

    root = (group_roots or {}).get(parent.strip().upper(), "")
    if root:
        if "SALES" in root:
            return "sales"
        if "PURCHASE" in root:
            return "purchase"
        if "DIRECT EXPENSES" in root:
            return "direct_cost"
        return None

    # fallback when the group tree is unavailable
    p = parent.upper()
    if "SALES ACCOUNTS" in p or "SALES" in p or "SALE ACCOUNT" in p:
        return "sales"
    if "PURCHASE" in p:
        return "purchase"
    if _DIRECT_COST_RE.search(n):
        return "direct_cost"
    return None


def fetch_ledgers(company, from_date, to_date, host, port,
                  group_roots=None) -> list:
    """Return every P&L-relevant ledger with its period movement."""
    xml = _collection_xml(company, from_date, to_date, "GPLedgers",
                          "Ledger", "Name,Parent,ClosingBalance")
    root = _safe_parse(post_to_tally(xml, host, port))

    rows = []
    for elem in root.iter("LEDGER"):
        name = _name(elem)
        if not name:
            continue
        parent = _text(elem, "PARENT")
        bucket = classify_ledger(name, parent, group_roots)
        if bucket is None:
            continue
        amount = parse_value(_text(elem, "CLOSINGBALANCE"))
        if not amount:
            continue
        upper = name.upper()
        rows.append({
            "ledger": name,
            "parent": parent,
            "bucket": bucket,
            "amount": amount,
            "is_branch": "BRANCH" in upper,
            "is_related_party": "RELATED PARTY" in upper,
            "is_return": "RETURN" in upper,
        })
    return rows


# ---------------------------------------------------------------------------

def build_periods(fy_start: date, through: date) -> list:
    """Monthly periods from FY start through `through`, plus a Q1 aggregate.

    Q1 is fetched in its own right rather than summed, so the workbook can
    cross-check it against Apr+May+Jun and catch period-boundary errors.
    """
    periods = []
    y, m = fy_start.year, fy_start.month
    while (y, m) <= (through.year, through.month):
        start = date(y, m, 1)
        nxt = date(y + (m == 12), (m % 12) + 1, 1)
        periods.append({
            "key": start.strftime("%b-%y").upper(),
            "kind": "month",
            "from": start.isoformat(),
            "to": min(date.fromordinal(nxt.toordinal() - 1), through).isoformat(),
        })
        y, m = nxt.year, nxt.month

    q1_end = date(fy_start.year, 6, 30)
    if through >= q1_end:
        periods.append({"key": "Q1 (APR-JUN)", "kind": "quarter",
                        "from": fy_start.isoformat(), "to": q1_end.isoformat()})
    return periods


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default=TALLY_HOST)
    ap.add_argument("--port", type=int, default=TALLY_PORT)
    ap.add_argument("--through", default=None)
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    through = date.fromisoformat(args.through) if args.through else date.today()
    fy_start = date(through.year if through.month >= 4 else through.year - 1, 4, 1)
    periods = build_periods(fy_start, through)

    companies = list(COMPANY_ENTITY)
    print(f"Companies: {len(companies)}   Periods: {len(periods)}")

    data = {"generated": date.today().isoformat(),
            "fy_start": fy_start.isoformat(), "through": through.isoformat(),
            "periods": periods, "companies": {}}

    for company in companies:
        entity, location = COMPANY_ENTITY[company]
        print(f"\n=== {company}  [{entity} / {location}]", flush=True)
        try:
            parents = fetch_group_parents(company, periods[0]["from"],
                                          periods[-1]["to"], args.host, args.port)
        except Exception as exc:                       # company not open
            print(f"    SKIPPED - {exc}", flush=True)
            continue
        group_roots = fetch_group_roots(company, periods[0]["from"],
                                        periods[-1]["to"], args.host, args.port)
        print(f"    stock groups: {len(parents)}   "
              f"account groups: {len(group_roots)}", flush=True)

        cdata = {"entity": entity, "location": location,
                 "parents": parents, "periods": {}, "ledgers": {}}
        for p in periods:
            items = fetch_period_items(company, p["from"], p["to"],
                                       args.host, args.port)
            ledgers = fetch_ledgers(company, p["from"], p["to"],
                                    args.host, args.port, group_roots)
            cdata["periods"][p["key"]] = items
            cdata["ledgers"][p["key"]] = ledgers

            sales = sum(r["amount"] for r in ledgers
                        if r["bucket"] == "sales" and not r["is_branch"]
                        and not r["is_related_party"])
            branch = sum(r["amount"] for r in ledgers
                         if r["bucket"] == "sales" and r["is_branch"])
            print(f"    {p['key']:<14} items={items['item_count']:<6} "
                  f"ledgers={len(ledgers):<4} sales={sales:>15,.0f} "
                  f"branch={branch:>14,.0f}", flush=True)

        data["companies"][company] = cdata

    out = Path(args.out) if args.out else (
        PROJECT_ROOT / "Files" / "gp_segment_cache.json")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data, indent=1), encoding="utf-8")
    print(f"\nCached -> {out}")


if __name__ == "__main__":
    main()
