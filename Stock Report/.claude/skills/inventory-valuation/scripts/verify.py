"""Independent verification of the published workbook.

Re-derives the top gaps straight from the cached raw Tally XML — NOT from the parquet the
workbook was built from — and compares against what the workbook actually says.
Also checks every hyperlink resolves.
"""
import re
from pathlib import Path
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from tally_lib import SCRATCH, clean_xml, num, qty, rate

from tally_lib import ROOT  # noqa: E402

# always verify the newest workbook sitting directly in Inventory\
WB = max((ROOT / "Inventory").glob("Stock_Valuation_Check_*.xlsx"),
         key=lambda p: p.stat().st_mtime)
print(f"verifying {WB.name}\n")

# ------------------------------------------------ 1. hyperlinks
wb = load_workbook(WB)
names = set(wb.sheetnames)
tot = bad = 0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if not c.hyperlink:
                continue
            tot += 1
            t = str(c.hyperlink.location or c.hyperlink.target or "")
            sh = t.lstrip("#").split("!")[0].strip("'")
            if sh not in names:
                bad += 1
                print("  BROKEN", ws.title, c.coordinate, t)
print(f"[links] {tot} hyperlinks, {bad} broken")

# ------------------------------------------------ 2. re-derive from raw XML
raw = SCRATCH / "raw" / ("full_ORANGE_O_TEC_PRIVATE_LIMITED_01_04_25TO31_03_27_.xml")
if not raw.exists():
    cands = list((SCRATCH / "raw").glob("full_ORANGE_O_TEC_PRIVATE_LIMITED_01*.xml"))
    raw = cands[0]
root = ET.fromstring(clean_xml(raw.read_text(encoding="utf-8", errors="replace")))

fresh = {}
for it in root.findall(".//STOCKITEM"):
    nm = it.get("NAME")
    if not nm:
        continue

    def g(t):
        e = it.find(t)
        return e.text if e is not None else None

    cq, _ = qty(g("CLOSINGBALANCE"))
    iq, _ = qty(g("CREDITTOTALS"))
    cv = -num(g("CLOSINGVALUE"))
    iv = -num(g("INWARDVALUE"))
    fresh[nm] = dict(cq=cq, cv=cv, iq=iq, iv=iv,
                     ir=(iv / iq if iq else None),
                     gap=(cv - cq * (iv / iq)) if iq else None,
                     cr=rate(g("CLOSINGRATE")))

# ------------------------------------------------ 3. compare with workbook
ws = wb["Orange O Tec"]
hdr_row = None
for r in range(1, 40):
    if ws.cell(row=r, column=1).value == "Rank":
        hdr_row = r
        break
hdrs = [ws.cell(row=hdr_row, column=c).value for c in range(1, ws.max_column + 1)]
ix = {h: i + 1 for i, h in enumerate(hdrs)}

# The gap columns are live formulas (verified separately by verify_formulas.py), so compare
# the INPUTS the formulas consume — closing value, closing qty and inward rate — against a
# fresh parse of the raw Tally XML.
print(f"\n[recalc] re-proving detail inputs from raw XML ({raw.name})\n")
print(f"{'item':<44}{'wb close val':>16}{'raw close val':>16}{'rate':>10}  verdict")
checked = ok = 0
r = hdr_row + 1
while checked < 15 and r <= ws.max_row:
    name = ws.cell(row=r, column=ix["Item Name"]).value
    unit = ws.cell(row=r, column=ix["Unit"]).value
    wbval = ws.cell(row=r, column=ix["Tally Value"]).value
    wbqty = ws.cell(row=r, column=ix["Closing Qty"]).value
    wbir = ws.cell(row=r, column=ix["Inward Rate"]).value
    r += 1
    if unit != "HO" or name not in fresh or not isinstance(wbval, (int, float)):
        continue
    f = fresh[name]
    checked += 1
    okv = abs(f["cv"] - wbval) < 0.01 and abs(f["cq"] - (wbqty or 0)) < 0.001
    okr = (f["ir"] is None and wbir is None) or (
        f["ir"] is not None and wbir is not None and abs(f["ir"] - wbir) < 0.01)
    ok += okv and okr
    print(f"{name[:42]:<44}{wbval:>16,.2f}{f['cv']:>16,.2f}"
          f"{('n/a' if wbir is None else f'{wbir:,.2f}'):>10}  "
          f"{'MATCH' if okv and okr else '*** MISMATCH ***'}")

print(f"\n[recalc] {ok}/{checked} re-proved from raw XML")

# ------------------------------------------------ 4. totals tie
tots = wb["Summary"]
print("\n[summary] group row as published:")
for r in range(1, tots.max_row + 1):
    if str(tots.cell(row=r, column=1).value or "").startswith("GROUP TOTAL"):
        for c in range(1, 11):
            h = tots.cell(row=r - (r - 6), column=c).value
            v = tots.cell(row=r, column=c).value
            if isinstance(v, (int, float)):
                print(f"    col{c}: {v:,.2f}")
            else:
                print(f"    col{c}: {v}")
        break
