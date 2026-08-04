"""Voucher-level extract for true item-wise gross profit.

This is the accurate route. Instead of inferring segment from ledger names
(which only Orange O Tec tags), it reads every sales and purchase INVENTORY
LINE and takes the segment from the STOCK ITEM's own group. That works for all
five companies and gives sale quantity as well.

Per item, per period, it accumulates:
    sales qty / value      (Sales vouchers, net of Credit Notes)
    purchase qty / value   (Purchase vouchers, net of Debit Notes)
plus opening and closing quantity and value from the StockItem collection, so
cost can be struck on a weighted-average basis:

    cost rate = (opening value + purchase value)
                / (opening qty + purchase qty)
    COGS      = sales qty * cost rate
    GP        = sales value - COGS

Voucher types are classified by walking their PARENT chain to a Tally base
type. This is essential: PROFORMA INVOICE sits under "Sales Order" and
DELIVERY CHALLAN under "Delivery Note". Counting those as sales would have
overstated Orange O Tec HO April revenue by about 8.8 crore.

Run:
    python .claude\\skills\\gp-segment-mom\\scripts\\fetch_gp_vouchers.py
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path
import xml.etree.ElementTree as ET

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parents[3]
sys.path.insert(0, str(PROJECT_ROOT / "Stock Report" / ".claude" / "skills" /
                      "fetch-stock-summary" / "scripts"))
from fetch_stock_summary import (  # noqa: E402
    _clean_xml, _collection_xml, _safe_parse, post_to_tally,
    TALLY_HOST, TALLY_PORT, _to_ddmmmyyyy,
)
sys.path.insert(0, str(HERE))
from segment_map import COMPANY_ENTITY  # noqa: E402
from fetch_gp_data import build_periods, parse_value  # noqa: E402


# ---------------------------------------------------------------------------
# parsing helpers
# ---------------------------------------------------------------------------

def parse_qty(raw) -> tuple[float, str]:
    """' 19200.000 KGS' -> (19200.0, 'KGS');  ' 1 NOS' -> (1.0, 'NOS')."""
    if not raw:
        return 0.0, ""
    s = str(raw).strip()
    m = re.match(r"^(-?[\d,]*\.?\d+)\s*(.*)$", s)
    if not m:
        return 0.0, ""
    try:
        qty = float(m.group(1).replace(",", ""))
    except ValueError:
        return 0.0, ""
    return qty, m.group(2).strip().upper()


def _text(elem, tag) -> str:
    node = elem.find(tag)
    return (node.text or "").strip() if node is not None else ""


def _name(elem) -> str:
    """Tally puts the name on an attribute for some collections, a child for others."""
    return (elem.get("NAME") or _text(elem, "NAME")).strip()


def _parse_vouchers(raw: str):
    """Tally emits undeclared namespace prefixes on some UDF tags."""
    cleaned = re.sub(r"<(/?)[A-Za-z0-9_]+:", r"<\1", _clean_xml(raw))
    return ET.fromstring(cleaned)


# ---------------------------------------------------------------------------
# voucher type classification
# ---------------------------------------------------------------------------

BASE_SALES = "SALES"
BASE_PURCHASE = "PURCHASE"
BASE_CREDIT = "CREDIT NOTE"
BASE_DEBIT = "DEBIT NOTE"

# Only these four base types move goods at a transacted price. Everything else
# - Sales Order (proforma), Delivery Note (challan), Receipt Note, Stock
# Journal, Material In/Out, Rejections, Journal, Physical Stock - is either an
# order, a transfer, or a non-priced movement and must NOT count as a sale.
COUNTED = {BASE_SALES, BASE_PURCHASE, BASE_CREDIT, BASE_DEBIT}


def fetch_voucher_type_base(company, from_date, to_date, host, port) -> dict:
    """{voucher type name -> base Tally type}, resolved through the parent chain."""
    root = _safe_parse(post_to_tally(
        _collection_xml(company, from_date, to_date, "GPVT",
                        "VoucherType", "Name,Parent"), host, port))
    parent = {}
    for elem in root.iter("VOUCHERTYPE"):
        name = (elem.get("NAME") or _text(elem, "NAME")).strip()
        if name:
            parent[name.upper()] = _text(elem, "PARENT").strip()

    def root_of(name: str) -> str:
        """Walk to the root of the parent chain.

        A voucher type that is its OWN parent is a root. Tally lets the default
        types be renamed, so Orange O Tec HO has the base Sales type renamed to
        "Sales Accounts-HSS" and Purchase renamed to "GST PURCHASE" - the root
        name cannot be assumed to be literally "Sales".
        """
        seen = set()
        current = name
        while current:
            key = current.upper()
            if key in seen:
                return key
            seen.add(key)
            nxt = parent.get(key, "")
            if not nxt or nxt.upper() == key:
                return key
            current = nxt
        return (name or "").upper()

    return {k: classify_root(root_of(k)) for k in parent}


def classify_root(root: str) -> str:
    """Map a root voucher type name to the base category we count.

    Order matters: "Sales Order" and "Delivery Note" must be caught before the
    bare SALES / PURCHASE tests, or proforma invoices and delivery challans
    would be booked as revenue.
    """
    r = (root or "").upper()
    if "ORDER" in r:
        return "ORDER"
    if "DELIVERY NOTE" in r or "RECEIPT NOTE" in r or "CHALLAN" in r:
        return "NOTE"
    if "CREDIT NOTE" in r:
        return BASE_CREDIT
    if "DEBIT NOTE" in r:
        return BASE_DEBIT
    if "SALES" in r or "SALE" in r:
        return BASE_SALES
    if "PURCHASE" in r:
        return BASE_PURCHASE
    return r


def fetch_branch_ledgers(company, from_date, to_date, host, port) -> set:
    """Ledgers sitting under Tally's 'Branch / Divisions' group.

    Inter-branch sales are posted to a single untagged 'GST BRANCH SALE
    ACCOUNTS' ledger, so they cannot be segmented from the ledger side at all.
    They ARE ordinary sales vouchers with inventory lines though, so flagging
    the voucher by its party lets the stock item supply the segment.
    """
    groups = _safe_parse(post_to_tally(_collection_xml(
        company, from_date, to_date, "GPGrp", "Group", "Name,Parent"),
        host, port))
    parent = {}
    for elem in groups.iter("GROUP"):
        nm = _name(elem)
        if nm:
            parent[nm.upper()] = _text(elem, "PARENT").strip()

    def is_branch(group: str) -> bool:
        seen = set()
        cur = group
        while cur:
            key = cur.upper()
            if key in seen:
                return False
            if "BRANCH" in key and "DIVISION" in key:
                return True
            seen.add(key)
            nxt = parent.get(key, "")
            if not nxt or nxt.upper() == key:
                return False
            cur = nxt
        return False

    branch_groups = {g for g in parent if is_branch(g)}
    ledgers = _safe_parse(post_to_tally(_collection_xml(
        company, from_date, to_date, "GPLed", "Ledger", "Name,Parent"),
        host, port))
    out = set()
    for elem in ledgers.iter("LEDGER"):
        nm = _name(elem)
        if nm and _text(elem, "PARENT").strip().upper() in branch_groups:
            out.add(nm.upper())
    return out


# ---------------------------------------------------------------------------
# stock items: group, opening / closing qty and value
# ---------------------------------------------------------------------------

def fetch_items(company, from_date, to_date, host, port) -> dict:
    root = _safe_parse(post_to_tally(_collection_xml(
        company, from_date, to_date, "GPItemDetail", "StockItem",
        "Name,Parent,OpeningBalance,OpeningValue,ClosingBalance,ClosingValue"),
        host, port))
    items = {}
    for elem in root.iter("STOCKITEM"):
        name = (elem.get("NAME") or _text(elem, "NAME")).strip()
        if not name:
            continue
        oq, unit = parse_qty(_text(elem, "OPENINGBALANCE"))
        cq, unit2 = parse_qty(_text(elem, "CLOSINGBALANCE"))
        items[name.upper()] = {
            "name": name,
            "group": _text(elem, "PARENT") or "(No stock group assigned)",
            "open_qty": oq,
            "close_qty": cq,
            # debit balances come back negative; stock is an asset
            "open_val": -parse_value(_text(elem, "OPENINGVALUE")),
            "close_val": -parse_value(_text(elem, "CLOSINGVALUE")),
            "unit": unit or unit2,
        }
    return items


# ---------------------------------------------------------------------------
# vouchers
# ---------------------------------------------------------------------------

VOUCHER_XML = """<ENVELOPE>
  <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>
  <TYPE>Collection</TYPE><ID>GPVouchers</ID></HEADER>
  <BODY><DESC><STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    <SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
    <SVFROMDATE TYPE="Date">{from_date}</SVFROMDATE>
    <SVTODATE TYPE="Date">{to_date}</SVTODATE>
  </STATICVARIABLES><TDL><TDLMESSAGE>
    <COLLECTION ISMODIFY="No" NAME="GPVouchers">
      <TYPE>Voucher</TYPE>
      <FETCH>Date,VoucherTypeName,VoucherNumber,IsCancelled,IsOptional,PartyLedgerName</FETCH>
      <FETCH>AllInventoryEntries.StockItemName,AllInventoryEntries.BilledQty,AllInventoryEntries.Amount</FETCH>
    </COLLECTION>
  </TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"""


def fetch_period_vouchers(company, from_date, to_date, vt_base, host, port,
                          branch_ledgers=None):
    """Accumulate sales / purchase quantity and value per stock item."""
    raw = post_to_tally(VOUCHER_XML.format(
        company=company, from_date=_to_ddmmmyyyy(from_date),
        to_date=_to_ddmmmyyyy(to_date)), host, port)
    root = _parse_vouchers(raw)

    branch_ledgers = branch_ledgers or set()
    per_item = defaultdict(lambda: dict(sales_qty=0.0, sales_val=0.0,
                                        purch_qty=0.0, purch_val=0.0,
                                        ret_qty=0.0, ret_val=0.0,
                                        br_sales_qty=0.0, br_sales_val=0.0,
                                        br_purch_qty=0.0, br_purch_val=0.0))
    skipped = defaultdict(float)
    counted_types = defaultdict(float)

    for v in root.iter("VOUCHER"):
        if (_text(v, "ISCANCELLED") or "").lower() in ("yes", "true"):
            continue
        if (_text(v, "ISOPTIONAL") or "").lower() in ("yes", "true"):
            continue
        vtype = _text(v, "VOUCHERTYPENAME")
        base = vt_base.get(vtype.upper(), vtype.upper())
        to_branch = _text(v, "PARTYLEDGERNAME").strip().upper() in branch_ledgers

        for e in v.findall("ALLINVENTORYENTRIES.LIST"):
            item = _text(e, "STOCKITEMNAME")
            if not item:
                continue
            amount = parse_value(_text(e, "AMOUNT"))
            qty, _unit = parse_qty(_text(e, "BILLEDQTY"))
            if base not in COUNTED:
                skipped[base] += amount
                continue
            counted_types[vtype] += amount
            rec = per_item[item.upper()]
            # Tally signs AMOUNT on returns but NOT BilledQty: a credit note
            # comes back as qty '2.000 PCS' with amount -676700. Left as-is a
            # return would ADD to quantity sold and so add COGS, which is what
            # drove spare parts to a negative margin. Negate quantity on the
            # return voucher types to match the signed amount.
            if base == BASE_SALES:
                rec["sales_val"] += amount
                rec["sales_qty"] += qty
                if to_branch:
                    rec["br_sales_val"] += amount
                    rec["br_sales_qty"] += qty
            elif base == BASE_CREDIT:
                rec["sales_val"] += amount
                rec["sales_qty"] -= abs(qty)
                rec["ret_val"] += amount
                rec["ret_qty"] += abs(qty)
                if to_branch:
                    rec["br_sales_val"] += amount
                    rec["br_sales_qty"] -= abs(qty)
            elif base == BASE_PURCHASE:
                # purchases are debits (negative); flip to positive cost
                rec["purch_val"] += -amount
                rec["purch_qty"] += qty
                if to_branch:
                    rec["br_purch_val"] += -amount
                    rec["br_purch_qty"] += qty
            else:                                     # BASE_DEBIT
                rec["purch_val"] += -amount
                rec["purch_qty"] -= abs(qty)
                if to_branch:
                    rec["br_purch_val"] += -amount
                    rec["br_purch_qty"] -= abs(qty)

    return {"per_item": {k: v for k, v in per_item.items()
                         if any(abs(x) > 1e-9 for x in v.values())},
            "skipped": dict(skipped),
            "counted_types": dict(counted_types)}


# ---------------------------------------------------------------------------

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

    data = {"generated": date.today().isoformat(),
            "fy_start": fy_start.isoformat(), "through": through.isoformat(),
            "periods": periods, "companies": {}}

    for company in COMPANY_ENTITY:
        entity, location = COMPANY_ENTITY[company]
        print(f"\n=== {company}  [{entity} / {location}]", flush=True)
        try:
            vt_base = fetch_voucher_type_base(
                company, fy_start.isoformat(), through.isoformat(),
                args.host, args.port)
        except Exception as exc:
            print(f"    SKIPPED - {exc}", flush=True)
            continue
        branch_ledgers = fetch_branch_ledgers(
            company, fy_start.isoformat(), through.isoformat(),
            args.host, args.port)
        print(f"    voucher types: {len(vt_base)}   "
              f"branch ledgers: {len(branch_ledgers)}", flush=True)

        # FY-to-date item snapshot: supplies the fallback cost rate for items
        # sold out of stock bought before the period being reported.
        ytd_items = fetch_items(company, fy_start.isoformat(),
                                through.isoformat(), args.host, args.port)
        ytd_vch = fetch_period_vouchers(company, fy_start.isoformat(),
                                        through.isoformat(), vt_base,
                                        args.host, args.port, branch_ledgers)
        print(f"    items: {len(ytd_items)}   YTD active items: "
              f"{len(ytd_vch['per_item'])}", flush=True)

        cdata = {"entity": entity, "location": location,
                 "ytd_items": ytd_items, "ytd_vouchers": ytd_vch["per_item"],
                 "periods": {}}

        for p in periods:
            items = fetch_items(company, p["from"], p["to"], args.host, args.port)
            vch = fetch_period_vouchers(company, p["from"], p["to"], vt_base,
                                        args.host, args.port, branch_ledgers)
            # keep only items with activity or stock, to hold the cache down
            keep = {k: v for k, v in items.items()
                    if k in vch["per_item"] or abs(v["open_val"]) > 1e-9
                    or abs(v["close_val"]) > 1e-9}
            cdata["periods"][p["key"]] = {"items": keep,
                                          "vouchers": vch["per_item"],
                                          "skipped": vch["skipped"]}
            sales = sum(r["sales_val"] for r in vch["per_item"].values())
            purch = sum(r["br_sales_val"] for r in vch["per_item"].values())
            print(f"    {p['key']:<14} lines_items={len(vch['per_item']):<5} "
                  f"sales={sales:>15,.0f}  branch={purch:>14,.0f}", flush=True)

        data["companies"][company] = cdata

    out = Path(args.out) if args.out else (
        PROJECT_ROOT / "Files" / "gp_voucher_cache.json")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(data), encoding="utf-8")
    print(f"\nCached -> {out}  ({out.stat().st_size/1e6:.1f} MB)")


if __name__ == "__main__":
    main()
