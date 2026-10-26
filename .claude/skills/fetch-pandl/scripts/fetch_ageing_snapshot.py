"""
fetch_ageing_snapshot.py - debtor / creditor bills and stock, as on today, from ConnectWave (read-only).

For the CFO Pack's ageing and inventory tabs (user, 2026-10-24: "develop the ageing report as on date").
ConnectWave holds TODAY's position only (its bill and stock tables are rebuilt nightly), so these tabs are
"as on <date>", not at the pack's 30-Sep balance-sheet date - the pack says so wherever they appear.

Per company (the five FY 26-27 books, bare tenants - never the ~YYYYMMDD archive tenants):
  receivables   rpt_receivables_bill   ledger, bill_ref, bill_date, pending_amount (Dr +)
  payables      rpt_payables_bill      ledger, bill_ref, bill_date, pending_amount (Cr -)
  stock         rpt_stock_summary_item item, stock group path, closing qty / value (items holding stock)
  inward        rpt_batch_line         stock receipts (movement in, affects stock) - dates for FIFO ageing

Output: Files/Results/Ageing_Snapshot_<YYYY-MM-DD_HHMM>.json, read by build_cfo_pack.py (newest one).

    python fetch_ageing_snapshot.py
"""
from __future__ import annotations

import datetime as dt
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(ROOT.parent / "Orange One Supabase Connect" / "tools"))
from supabase_client import select  # noqa: E402  (read-only client: GET only)

RESULTS = ROOT / "Files" / "Results"
# Pack entity key -> ConnectWave company guid (FY 26-27 books).
COMPANIES = {
    "ORANGE_O_TEC_PRIVATE_LIMITED": "a4e100d1-3b6f-4193-876a-c754f1a74552",
    "PRIVATE_LIMITED_NOIDA": "53d35745-5246-4e1a-a27a-d4769f245b50",
    "ENTERPRISES_PVT_LTD": "59a6c2d9-0c5a-4fc5-b8c5-3be6fec3289e",
    "ENTERPRISES_NOIDA": "779c26f4-3fd8-46bd-9995-4f9916c98856",
    "COLORIX": "393ee4bd-4fdc-4aed-ae88-e2ff1394927a",
}
BILL_COLS = "ledger,bill_ref,bill_date,due_date,bill_amount,pending_amount,built_at"
STOCK_COLS = "item,base_unit,stock_group,primary_group,group_path,closing_qty,closing_value,last_move_date,built_at"
INWARD_COLS = "fy,vch_date,voucher_guid,line_no,batch_no,voucher_type,stock_item,qty,amount,tracking_number"


def fetch(guid: str) -> dict:
    t = {"tenant_id": f"eq.acct_orange::{guid}"}
    out = {}
    for name, table in (("receivables", "rpt_receivables_bill"), ("payables", "rpt_payables_bill")):
        out[name] = select(table, columns=BILL_COLS, filters=t, order="ledger,bill_ref,bill_date,pending_amount,due_date")
    out["stock"] = [s for s in select("rpt_stock_summary_item", columns=STOCK_COLS, filters={**t, "fy": "eq.2026-27"},
                                      order="item")
                    if (s["closing_qty"] or 0) > 0 or abs(s["closing_value"] or 0) > 0.5]
    out["inward"] = select("rpt_batch_line", columns=INWARD_COLS,
                           filters={**t, "movement": "eq.in", "affects_stock": "is.true"},
                           order="fy,voucher_guid,line_no,batch_no")
    return out


def main() -> int:
    today = dt.date.today()
    snap = {"as_of": today.isoformat(), "fetched": dt.datetime.now().isoformat(timespec="minutes"), "companies": {}}
    for key, guid in COMPANIES.items():
        print(f"  {key} ...", flush=True)
        d = fetch(guid)
        snap["companies"][key] = d
        print(f"    bills receivable {len(d['receivables'])}, payable {len(d['payables'])}, "
              f"stock items {len(d['stock'])}, receipt lines {len(d['inward'])}", flush=True)
    out = RESULTS / f"Ageing_Snapshot_{dt.datetime.now():%Y-%m-%d_%H%M}.json"
    out.write_text(json.dumps(snap), encoding="utf-8")
    print(f"OUTPUT_JSON={out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
