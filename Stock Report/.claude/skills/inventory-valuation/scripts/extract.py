"""Phase B2 — full per-item extract for 1-Apr-26..30-Jun-26, all 5 files, and rebuild
closing stock on independent cost bases.

Tally method-name notes (established by probe, not assumed):
    CREDITTOTALS = inward quantity      (Tally's stock debit/credit runs opposite intuition)
    DEBITTOTALS  = outward quantity
    INWARDVALUE / OUTWARDVALUE = period values (credit-side, so negated)
    STANDARDCOST = the Std. Cost rate on the item master
"""
import re
import pandas as pd
from tally_lib import FILES, collection, parse, num, qty, rate, SCRATCH

RAW = SCRATCH / "raw"
RAW.mkdir(exist_ok=True)

FIELDS = ["Name", "Parent", "BaseUnits", "CostingMethod", "ValuationMethod",
          "StandardCost", "OpeningBalance", "OpeningValue", "OpeningRate",
          "ClosingBalance", "ClosingValue", "ClosingRate",
          "CreditTotals", "DebitTotals", "InwardValue", "OutwardValue"]

BODY = "<TYPE>StockItem</TYPE>\n" + \
       "\n".join(f"<NATIVEMETHOD>{f}</NATIVEMETHOD>" for f in FIELDS)


def slug(s):
    return re.sub(r"[^A-Za-z0-9]+", "_", s)[:60]


def g(n, t):
    e = n.find(t)
    return e.text if e is not None else None


rows = []
for company, entity, unit in FILES:
    cache = RAW / f"full_{slug(company)}.xml"
    if cache.exists() and cache.stat().st_size > 5000:
        txt, src = cache.read_text(encoding="utf-8", errors="replace"), "cache"
    else:
        txt = collection(company, "FullExtract", BODY)
        if "<STOCKITEM" not in txt:
            raise RuntimeError(f"No stock items for {company!r} ({len(txt)} bytes)")
        cache.write_text(txt, encoding="utf-8")
        src = "tally"

    nodes = parse(txt).findall(".//STOCKITEM")
    print(f"{entity:<14} {unit:<6} {len(nodes):>6} items  {len(txt)/1e6:>5.1f} MB  [{src}]")

    for it in nodes:
        name = it.get("NAME")
        if not name:
            continue
        ob, u1 = qty(g(it, "OPENINGBALANCE"))
        cb, u2 = qty(g(it, "CLOSINGBALANCE"))
        inq, _ = qty(g(it, "CREDITTOTALS"))     # inward qty
        outq, _ = qty(g(it, "DEBITTOTALS"))     # outward qty
        rows.append({
            "Entity": entity, "Unit": unit, "Company": company,
            "Item Name": name,
            "Stock Group": g(it, "PARENT") or "",
            "UOM": (g(it, "BASEUNITS") or u1 or u2 or "").strip(),
            "Costing Method": (g(it, "COSTINGMETHOD") or "").strip(),
            "Valuation Method": (g(it, "VALUATIONMETHOD") or "").strip(),
            "Std Cost Rate": rate(g(it, "STANDARDCOST")),
            "Opening Qty": ob,
            "Opening Value": -num(g(it, "OPENINGVALUE")),
            "Opening Rate": rate(g(it, "OPENINGRATE")),
            "Inward Qty": inq,
            "Inward Value": -num(g(it, "INWARDVALUE")),
            "Outward Qty": outq,
            "Outward Value": num(g(it, "OUTWARDVALUE")),
            "Closing Qty": cb,
            "Closing Value": -num(g(it, "CLOSINGVALUE")),
            "Closing Rate": rate(g(it, "CLOSINGRATE")),
        })

df = pd.DataFrame(rows)

# ---------------------------------------------------------------- derived bases
df["Roll-fwd Qty"] = df["Opening Qty"] + df["Inward Qty"] - df["Outward Qty"]
df["Roll-fwd Diff"] = (df["Roll-fwd Qty"] - df["Closing Qty"]).round(3)

# actual cost of goods received in the period
df["Inward Rate"] = (df["Inward Value"] / df["Inward Qty"]).where(df["Inward Qty"] > 0)
# weighted average over opening + period inwards
den = df["Opening Qty"] + df["Inward Qty"]
df["Wtd Avg Rate"] = ((df["Opening Value"] + df["Inward Value"]) / den).where(den > 0)

df["Value @ Inward Rate"] = df["Closing Qty"] * df["Inward Rate"]
df["Value @ Wtd Avg"] = df["Closing Qty"] * df["Wtd Avg Rate"]
df["Gap vs Inward Rate"] = df["Closing Value"] - df["Value @ Inward Rate"]
df["Gap vs Wtd Avg"] = df["Closing Value"] - df["Value @ Wtd Avg"]
df["Gap %"] = (df["Gap vs Inward Rate"] / df["Closing Value"]).where(
    df["Closing Value"] != 0)

# can period purchases alone cover the closing quantity?
df["Cover"] = "no period purchases"
df.loc[df["Inward Qty"] > 0, "Cover"] = "partial - opening layers remain"
df.loc[df["Inward Qty"] >= df["Closing Qty"], "Cover"] = "fully covered by period purchases"

df.to_parquet(SCRATCH / "phase_b2_full.parquet", index=False)
print(f"\nTOTAL {len(df):,} item rows\n")

# ---------------------------------------------------------------- verification
bad = df[df["Roll-fwd Diff"].abs() > 0.01]
print(f"CHECK roll-forward (open + in - out = close): "
      f"{len(df) - len(bad):,} tie / {len(bad):,} differ")
if len(bad):
    print(bad[["Entity", "Unit", "Item Name", "Opening Qty", "Inward Qty",
               "Outward Qty", "Closing Qty", "Roll-fwd Diff"]].head(10).to_string(index=False))

live = df[(df["Closing Qty"] != 0) | (df["Closing Value"] != 0)].copy()
print(f"\nItems with closing stock: {len(live):,}")
print(f"Group closing value: {live['Closing Value'].sum():,.2f}")

print("\n=== Coverage of closing qty by period purchases ===")
print(live.groupby("Cover").agg(Items=("Item Name", "size"),
                                Closing_Value=("Closing Value", "sum")).to_string(
    float_format=lambda v: f"{v:,.2f}"))

# only items where a benchmark exists
bench = live[live["Inward Qty"] > 0].copy()
print(f"\n=== Items with a period-purchase benchmark: {len(bench):,} "
      f"| closing value {bench['Closing Value'].sum():,.2f} ===")
print(f"Gap vs inward rate : {bench['Gap vs Inward Rate'].sum():,.2f}")
print(f"Gap vs wtd avg     : {bench['Gap vs Wtd Avg'].sum():,.2f}")

print("\n=== Gap vs inward rate by entity ===")
print(bench.groupby(["Entity", "Unit"]).agg(
    Items=("Item Name", "size"),
    Closing_Value=("Closing Value", "sum"),
    Gap_Inward=("Gap vs Inward Rate", "sum"),
    Gap_WtdAvg=("Gap vs Wtd Avg", "sum")).to_string(float_format=lambda v: f"{v:,.2f}"))

print("\n=== Gap by costing method ===")
print(bench.groupby("Costing Method").agg(
    Items=("Item Name", "size"),
    Closing_Value=("Closing Value", "sum"),
    Gap_Inward=("Gap vs Inward Rate", "sum")).sort_values(
    "Gap_Inward", ascending=False).to_string(float_format=lambda v: f"{v:,.2f}"))

print("\n=== Top 25 overstatements (closing carried above period purchase cost) ===")
top = bench.nlargest(25, "Gap vs Inward Rate")
print(top[["Entity", "Unit", "Item Name", "UOM", "Costing Method", "Closing Qty",
           "Closing Rate", "Inward Rate", "Closing Value", "Value @ Inward Rate",
           "Gap vs Inward Rate", "Gap %"]].to_string(
    index=False, float_format=lambda v: f"{v:,.2f}"))
