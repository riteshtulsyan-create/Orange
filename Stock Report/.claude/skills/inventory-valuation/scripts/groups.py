"""Fetch the stock-group tree per company and resolve each item's full hierarchy.

An item's PARENT is only its immediate group. Walking the group->parent chain up to the root
gives the segment view: Main Category / Sub Category / Product Group / Sub Group.
"""
import re
import pandas as pd
from tally_lib import FILES, collection, parse, SCRATCH

RAW = SCRATCH / "raw"
RAW.mkdir(parents=True, exist_ok=True)

BODY = """<TYPE>StockGroup</TYPE>
<NATIVEMETHOD>Name</NATIVEMETHOD>
<NATIVEMETHOD>Parent</NATIVEMETHOD>"""


def slug(s):
    return re.sub(r"[^A-Za-z0-9]+", "_", s)[:60]


def chain(group, parent_of, seen=None):
    """Root-first list of group names, e.g. ['PRINTING INK','REACTIVE H SERIES']."""
    out, seen = [], seen or set()
    g = (group or "").strip()
    while g and g.lower() not in ("primary", "") and g not in seen:
        seen.add(g)
        out.append(g)
        g = (parent_of.get(g) or "").strip()
    return list(reversed(out))


maps = {}
for company, entity, unit in FILES:
    cache = RAW / f"groups_{slug(company)}.xml"
    if cache.exists() and cache.stat().st_size > 500:
        txt = cache.read_text(encoding="utf-8", errors="replace")
        src = "cache"
    else:
        txt = collection(company, "GroupTree", BODY)
        if "<STOCKGROUP" not in txt:
            raise RuntimeError(f"No stock groups for {company!r} ({len(txt)} bytes)")
        cache.write_text(txt, encoding="utf-8")
        src = "tally"

    parent_of = {}
    for g in parse(txt).findall(".//STOCKGROUP"):
        nm = g.get("NAME")
        if not nm:
            continue
        pe = g.find("PARENT")
        parent_of[nm] = (pe.text or "") if pe is not None else ""
    maps[company] = parent_of
    print(f"{entity:<14} {unit:<6} {len(parent_of):>5} stock groups  [{src}]")

df = pd.read_parquet(SCRATCH / "phase_b3_classified.parquet")

lvl1, lvl2, lvl3, lvl4, full, depth = [], [], [], [], [], []
for _, r in df.iterrows():
    ch = chain(r["Stock Group"], maps[r["Company"]])
    lvl1.append(ch[0] if len(ch) > 0 else "")
    lvl2.append(ch[1] if len(ch) > 1 else "")
    lvl3.append(ch[2] if len(ch) > 2 else "")
    # anything deeper than 4 folds into level 4 so the full chain is never lost
    lvl4.append(" > ".join(ch[3:]) if len(ch) > 3 else "")
    full.append(" > ".join(ch))
    depth.append(len(ch))

df["Main Category"] = lvl1
df["Sub Category"] = lvl2
df["Product Group"] = lvl3
df["Sub Group"] = lvl4
df["Full Group Path"] = full
df["Group Depth"] = depth

df.to_parquet(SCRATCH / "phase_b4_grouped.parquet", index=False)
print(f"\nresolved hierarchy for {len(df):,} items")
print("max nesting depth:", df["Group Depth"].max())

print("\n=== Closing stock by MAIN CATEGORY (segment) ===")
seg = df.groupby("Main Category").agg(
    Items=("Item Name", "size"),
    Closing_Value=("Closing Value", "sum"),
    Gap=("Gap vs Inward Rate", "sum")).sort_values("Closing_Value", ascending=False)
seg["% of stock"] = seg["Closing_Value"] / seg["Closing_Value"].sum() * 100
print(seg.head(20).to_string(float_format=lambda v: f"{v:,.2f}"))

blank = int((df["Main Category"] == "").sum())
print(f"\nitems with no resolvable main category: {blank}")
