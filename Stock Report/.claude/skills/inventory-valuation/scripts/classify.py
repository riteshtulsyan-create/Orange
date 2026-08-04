"""Phase B3 — clean the benchmark, then quantify.

A period-purchase rate is only a valid cost benchmark when goods actually arrived WITH value.
Items received at nil/near-nil value (inter-branch transfers booked at zero, stock journals,
free receipts) produce a fake 100% "overstatement" and must be excluded from the gap and
reported as a data-quality exception instead.
"""
import pandas as pd
from tally_lib import SCRATCH

pd.set_option("display.width", 250)
df = pd.read_parquet(SCRATCH / "phase_b2_full.parquet")

live = df[(df["Closing Qty"] != 0) | (df["Closing Value"] != 0)].copy()

# ---------------------------------------------------------------- classify
live["Benchmark"] = "none - no period purchases"
live.loc[live["Inward Qty"] > 0, "Benchmark"] = "invalid - received at nil value"
live.loc[(live["Inward Qty"] > 0) & (live["Inward Value"] > 1) &
         (live["Inward Rate"] > 0.01), "Benchmark"] = "valid"

# negative stock / stranded value are separate defects, never a valuation gap
live.loc[live["Closing Qty"] < 0, "Benchmark"] = "excluded - negative stock"
live.loc[(live["Closing Qty"] == 0) & (live["Closing Value"] != 0),
         "Benchmark"] = "excluded - value with no quantity"

print("=== Population split (items with closing stock) ===")
print(live.groupby("Benchmark").agg(
    Items=("Item Name", "size"),
    Closing_Value=("Closing Value", "sum")).to_string(float_format=lambda v: f"{v:,.2f}"))

valid = live[live["Benchmark"] == "valid"].copy()
print(f"\n=== VALID benchmark: {len(valid):,} items, "
      f"closing value {valid['Closing Value'].sum():,.2f} ===")
print(f"Gap vs period inward rate : {valid['Gap vs Inward Rate'].sum():>16,.2f}")
print(f"Gap vs weighted average   : {valid['Gap vs Wtd Avg'].sum():>16,.2f}")

over = valid[valid["Gap vs Inward Rate"] > 0]
under = valid[valid["Gap vs Inward Rate"] < 0]
print(f"\n  overstated : {len(over):>5,} items  {over['Gap vs Inward Rate'].sum():>16,.2f}")
print(f"  understated: {len(under):>5,} items  {under['Gap vs Inward Rate'].sum():>16,.2f}")

print("\n=== By entity (valid benchmark only) ===")
print(valid.groupby(["Entity", "Unit"]).agg(
    Items=("Item Name", "size"),
    Closing_Value=("Closing Value", "sum"),
    Gap_Inward=("Gap vs Inward Rate", "sum"),
    Gap_WtdAvg=("Gap vs Wtd Avg", "sum")).to_string(float_format=lambda v: f"{v:,.2f}"))

print("\n=== By costing method (valid benchmark only) ===")
print(valid.groupby("Costing Method").agg(
    Items=("Item Name", "size"),
    Closing_Value=("Closing Value", "sum"),
    Gap_Inward=("Gap vs Inward Rate", "sum")).sort_values(
    "Gap_Inward", ascending=False).to_string(float_format=lambda v: f"{v:,.2f}"))

print("\n=== Top 20 overstatements (valid benchmark) ===")
cols = ["Entity", "Unit", "Item Name", "UOM", "Costing Method", "Closing Qty",
        "Closing Rate", "Inward Rate", "Closing Value", "Gap vs Inward Rate", "Gap %"]
print(valid.nlargest(20, "Gap vs Inward Rate")[cols].to_string(
    index=False, float_format=lambda v: f"{v:,.2f}"))

print("\n=== Top 10 understatements (carried BELOW purchase cost) ===")
print(valid.nsmallest(10, "Gap vs Inward Rate")[cols].to_string(
    index=False, float_format=lambda v: f"{v:,.2f}"))

# ---------------------------------------------------------------- exceptions
print("\n\n############ DATA-QUALITY EXCEPTIONS ############")

nil = live[live["Benchmark"] == "invalid - received at nil value"]
print(f"\n--- Goods received at nil value: {len(nil):,} items, "
      f"closing value {nil['Closing Value'].sum():,.2f} ---")
print(nil.nlargest(12, "Closing Value")[
    ["Entity", "Unit", "Item Name", "UOM", "Inward Qty", "Inward Value",
     "Closing Qty", "Closing Value"]].to_string(
    index=False, float_format=lambda v: f"{v:,.2f}"))

neg = live[live["Closing Qty"] < 0]
print(f"\n--- Negative closing stock: {len(neg):,} items, "
      f"value {neg['Closing Value'].sum():,.2f} ---")
print(neg.nsmallest(12, "Closing Qty")[
    ["Entity", "Unit", "Item Name", "UOM", "Opening Qty", "Inward Qty",
     "Outward Qty", "Closing Qty", "Closing Value"]].to_string(
    index=False, float_format=lambda v: f"{v:,.2f}"))

stranded = live[(live["Closing Qty"] == 0) & (live["Closing Value"] != 0)]
print(f"\n--- Value with zero quantity: {len(stranded):,} items, "
      f"value {stranded['Closing Value'].sum():,.2f} ---")
print(stranded.nlargest(12, "Closing Value")[
    ["Entity", "Unit", "Item Name", "UOM", "Closing Qty", "Closing Value"]].to_string(
    index=False, float_format=lambda v: f"{v:,.2f}"))

nogrp = live[live["Stock Group"].isin(["", "Primary"])]
print(f"\n--- Items with no stock group: {len(nogrp):,} ---")

live.to_parquet(SCRATCH / "phase_b3_classified.parquet", index=False)
print("\nsaved phase_b3_classified.parquet")
