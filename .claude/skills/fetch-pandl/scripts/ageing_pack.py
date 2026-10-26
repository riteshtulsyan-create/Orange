"""
ageing_pack.py - the CFO Pack's ageing and inventory sheets (user, 2026-10-24), built from the newest
Files/Results/Ageing_Snapshot_*.json (fetch_ageing_snapshot.py - ConnectWave, AS ON THE FETCH DATE).

  Debtors Ageing     company | group | party | 0-180 days | over 180 days | total   (age = as-on date - invoice date)
  Creditors Ageing   the same for creditors (employees left out - they are on the Employees sheet)
  Inventory          closing stock by category and stock group
  Inventory Ageing   item by item, FIFO on receipt date: 0-90 / 91-180 / over 180 days, by category

The figures are the snapshot's, as on its date - not the pack's 30-Sep balance sheet. Every sheet says so.

FIFO ageing: an item's closing quantity is matched to its most recent receipts, newest first. Receipts are
purchases, receipt notes, sales returns / rejections in and production journals; internal transfers and other
stock journals only move stock, so they are not receipts. Quantity older than the book's history (Orange O Tec
from 1-Apr-25, Enterprises from 1-Apr-26) is dated at the history start - over 180 days either way.
"""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path

from openpyxl.utils import get_column_letter as col_letter

import build_cfo_pack as B

CATEGORIES = ["Machines", "Heads", "Spare parts", "Ink", "Paper & packing material", "Others"]
STOCK_BUCKETS = ["0-90 days", "91-180 days", "Over 180 days"]
BILL_BUCKETS = ["0-180 days", "Over 180 days"]
COMPANY_KEYS = ["ORANGE_O_TEC_PRIVATE_LIMITED", "PRIVATE_LIMITED_NOIDA", "ENTERPRISES_PVT_LTD", "ENTERPRISES_NOIDA", "COLORIX"]
COMPANY_SHORT = {"ORANGE_O_TEC_PRIVATE_LIMITED": "O Tec Surat", "PRIVATE_LIMITED_NOIDA": "O Tec Noida",
                 "ENTERPRISES_PVT_LTD": "Enterprises Surat", "ENTERPRISES_NOIDA": "Enterprises Noida", "COLORIX": "Colorix"}


def category(path: str) -> str:
    """Stock group path -> one of CATEGORIES, by its top (primary) group."""
    top = (path or "").split(" > ")[0].upper()
    if "HEAD" in top:
        return "Heads"
    if "PARTS" in top or "SPARE" in top:
        return "Spare parts"
    if "MACHIN" in top:
        return "Machines"
    if "INK" in top or top in ("FINISHED GOODS", "MANUFACTURING STOCK"):
        return "Ink"
    if "PAPER" in top or "PACKING" in top:
        return "Paper & packing material"
    return "Others"


def is_receipt(vtype: str, tracking) -> bool:
    v = (vtype or "").upper()
    if "PRODUCTION" in v and "JOURNAL" in v:
        return True  # manufactured: new stock
    if "TRANSFER" in v or "JOURNAL" in v or "MATERIAL IN" in v:
        return False  # moved, not received
    # An invoice tracked to a receipt note: the note already brought the stock in.
    return not (tracking and "NOTE" not in v)


def load_snapshot(results: Path):
    snaps = sorted(results.glob("Ageing_Snapshot_*.json"))
    return (json.loads(snaps[-1].read_text(encoding="utf-8")), snaps[-1].name) if snaps else (None, None)


def _date(s):
    return dt.date(int(s[:4]), int(s[4:6]), int(s[6:8])) if s else None


def party_rows(snap, kind, submap, skip_subs):
    """[{e, g, p, b: [0-180, >180]}] in rupees; payables turned positive."""
    as_of = dt.date.fromisoformat(snap["as_of"])
    sign = 1 if kind == "receivables" else -1
    out = {}
    for e in COMPANY_KEYS:
        for b in snap["companies"][e][kind]:
            p = b["ledger"]
            g = submap.get(p, "Not in the 30-Sep balance sheet")
            if g in skip_subs:
                continue
            d = _date(b["bill_date"])
            i = 0 if d is None or (as_of - d).days <= 180 else 1
            if d is None:
                i = 1  # undated (opening on-account) - treat as old
            row = out.setdefault((e, g, p), [0.0, 0.0])
            row[i] += sign * (b["pending_amount"] or 0.0)
    return [{"e": e, "g": g, "p": p, "b": [round(x, 2) for x in v]} for (e, g, p), v in sorted(out.items())
            if abs(v[0]) + abs(v[1]) >= 1]


def stock_rows(snap):
    """[{e, c, g, i, q, v, b: [0-90, 91-180, >180]}] - FIFO on receipt date, value split pro rata to quantity."""
    as_of = dt.date.fromisoformat(snap["as_of"])
    out = []
    for e in COMPANY_KEYS:
        comp = snap["companies"][e]
        receipts, seen, first = {}, set(), None
        for ln in comp["inward"]:
            key = (ln["voucher_guid"], ln["line_no"], ln["batch_no"])  # split books repeat vouchers
            if key in seen or not is_receipt(ln["voucher_type"], ln["tracking_number"]):
                continue
            seen.add(key)
            d = _date(ln["vch_date"])
            first = d if first is None or d < first else first
            receipts.setdefault(ln["stock_item"], []).append((d, ln["qty"] or 0.0))
        start = first or as_of
        for s in comp["stock"]:
            q, v = s["closing_qty"] or 0.0, s["closing_value"] or 0.0
            buckets = [0.0, 0.0, 0.0]
            if q <= 0:
                buckets[2] = v  # no quantity to age (value only, or negative stock): shown as old
            else:
                left = q
                for d, rq in sorted(receipts.get(s["item"], []), reverse=True):
                    if left <= 0:
                        break
                    take = min(left, rq)
                    age = (as_of - d).days
                    buckets[0 if age <= 90 else 1 if age <= 180 else 2] += take / q * v
                    left -= take
                if left > 0:  # older than the book's history
                    age = (as_of - start).days
                    buckets[0 if age <= 90 else 1 if age <= 180 else 2] += left / q * v
            out.append({"e": e, "c": category(s["group_path"]), "g": s["stock_group"] or "(Ungrouped)",
                        "pg": (s["group_path"] or "").split(" > ")[0], "i": s["item"], "u": s["base_unit"],
                        "q": round(q, 3), "v": round(v, 2), "b": [round(x, 2) for x in buckets]})
    return out


def _sheet(wb, title, note, headers, widths):
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = "BF8F00"
    ws["A1"], ws["A1"].font = title.upper(), B.F_TITLE
    ws["A2"], ws["A2"].font = note, B.F_NOTE
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        B.hdr_cell(ws, 4, i, h)
        ws.column_dimensions[col_letter(i)].width = w
    ws.freeze_panes = "A5"
    return ws


def _total_row(ws, r, first, last, cols, label="TOTAL"):
    B.label(ws, r, label, B.F_BOLD, B.FILL_TOT)
    for c in cols:
        B.num_cell(ws, r, c, f"=SUBTOTAL(9,{col_letter(c)}{first}:{col_letter(c)}{last})", B.NUM_CR, B.F_BOLD, B.FILL_TOT)


def build_ageing_sheets(wb, snap, snap_name, submap, emp_subs):
    """Writes the four sheets; returns (hub data, marks) - marks give the Directors Summary its cells."""
    as_of = dt.date.fromisoformat(snap["as_of"]).strftime("%d-%b-%Y")
    src = f"As on {as_of} - ConnectWave snapshot {snap_name} (not the 30-Sep balance sheet). Rs crore."
    hub = {"as_of": as_of, "source": snap_name, "debtors": party_rows(snap, "receivables", submap, set()),
           "creditors": party_rows(snap, "payables", submap, emp_subs), "stock": stock_rows(snap)}
    marks = {}
    for kind, title in (("debtors", "Debtors Ageing"), ("creditors", "Creditors Ageing")):
        ws = _sheet(wb, title, src + " Age = as-on date less invoice date. Filter any column.",
                    ["Company", "Group", "Party", *BILL_BUCKETS, "Total"], [18, 34, 44, 15, 15, 15])
        r = 5
        for row in hub[kind]:
            ws.cell(r, 1, COMPANY_SHORT[row["e"]])
            ws.cell(r, 2, row["g"])
            ws.cell(r, 3, row["p"])
            B.num_cell(ws, r, 4, row["b"][0] / B.CR, B.NUM_CR)
            B.num_cell(ws, r, 5, row["b"][1] / B.CR, B.NUM_CR)
            B.num_cell(ws, r, 6, f"=D{r}+E{r}", B.NUM_CR)
            r += 1
        ws.auto_filter.ref = f"A4:F{r - 1}"
        _total_row(ws, r, 5, r - 1, (4, 5, 6))
        marks[kind] = (title, r)
        B.setup_print(ws, landscape=False, fit_wide=1)

    ws = _sheet(wb, "Inventory", src + " Closing stock by category and stock group.",
                ["Company", "Category", "Primary group", "Stock group", "Items", "Value"], [18, 24, 30, 34, 8, 15])
    agg = {}
    for s in hub["stock"]:
        k = (s["e"], s["c"], s["pg"], s["g"])
        a = agg.setdefault(k, [0, 0.0])
        a[0] += 1
        a[1] += s["v"]
    r = 5
    for (e, c, pg, g), (n, v) in sorted(agg.items(), key=lambda kv: (COMPANY_KEYS.index(kv[0][0]), CATEGORIES.index(kv[0][1]), -kv[1][1])):
        for i, x in enumerate((COMPANY_SHORT[e], c, pg, g, n), start=1):
            ws.cell(r, i, x)
        B.num_cell(ws, r, 6, v / B.CR, B.NUM_CR)
        r += 1
    ws.auto_filter.ref = f"A4:F{r - 1}"
    _total_row(ws, r, 5, r - 1, (6,))
    B.setup_print(ws, landscape=False, fit_wide=1)

    ws = _sheet(wb, "Inventory Ageing", src + " FIFO on receipt date; value split pro rata to quantity. Filter Category.",
                ["Company", "Category", "Stock group", "Item", "Unit", "Qty", *STOCK_BUCKETS, "Value"],
                [18, 22, 30, 44, 8, 11, 14, 14, 14, 14])
    r = 5
    for s in sorted(hub["stock"], key=lambda s: (COMPANY_KEYS.index(s["e"]), CATEGORIES.index(s["c"]), s["g"], -s["v"])):
        for i, x in enumerate((COMPANY_SHORT[s["e"]], s["c"], s["g"], s["i"], s["u"]), start=1):
            ws.cell(r, i, x)
        B.num_cell(ws, r, 6, s["q"], "#,##0.00")
        for j, b in enumerate(s["b"]):
            B.num_cell(ws, r, 7 + j, b / B.CR, B.NUM_CR)
        B.num_cell(ws, r, 10, f"=SUM(G{r}:I{r})", B.NUM_CR)
        r += 1
    ws.auto_filter.ref = f"A4:J{r - 1}"
    _total_row(ws, r, 5, r - 1, (7, 8, 9, 10))
    B.setup_print(ws, landscape=True, fit_wide=1)
    return hub, marks
