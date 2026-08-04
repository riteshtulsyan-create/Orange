---
name: inventory-valuation
description: Run the Inventory Valuation check — rebuild closing stock on FIFO / weighted-average / last-purchase bases from Tally voucher data, compare against Tally's own carrying value, and report the gap at legal-entity level (Orange O Tec, Enterprises, Colorix) in a cross-linked Excel workbook. Use whenever the user asks to check stock valuation, verify closing stock value, test whether inventory is over/under-stated, audit costing methods, check landed cost of imports, run an NRV test on stock, or quantify unrealised profit on inter-branch stock transfers.
---

# Inventory Valuation

Answers one question: **is closing stock carried at the right value?** — and if not, by how much,
for which items, and why.

Produces `Inventory\Stock_Valuation_Check_<YYYY-MM-DD_HHMM>.xlsx`, a 10-sheet cross-linked
workbook presented at **legal entity** level.

## Build status

| Component | Status |
|-----------|--------|
| Design, sheet layout, column set, linking | Approved by the user 2026-07-31 |
| `scripts\tally_lib.py` — Tally access, retry, XML cleaning, entity map | **Working** |
| `scripts\extract.py` — per-item extract, all 5 files | **Working** |
| `scripts\classify.py` — benchmark validity + exception classification | **Working** |
| `scripts\build_workbook.py` — the 8-sheet workbook | **Working** |
| `scripts\verify.py` — link check + re-proof from raw XML | **Working** |
| `scripts\build_sample_layout.py` — design sample (illustrative figures) | Working |
| Landed cost / NRV / inter-branch elimination | **NOT BUILT** — see "Not yet covered" |

First live run: 2026-07-31, period 1-Apr-26 to 30-Jun-26. 14,090 masters, 5,216 items with
stock, roll-forward tied on every item, 15/15 top gaps re-proved from raw XML, 77/77 links good.

## Run

```powershell
cd "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS\Stock Report\.claude\skills\inventory-valuation\scripts"
python extract.py          # hits Tally; caches raw XML in _work\raw\
python classify.py         # benchmark validity + exceptions
python groups.py           # stock-group tree -> Main/Sub/Product/Sub-Group hierarchy
python build_workbook.py   # writes Inventory\Stock_Valuation_Check_<stamp>.xlsx
python verify.py           # links + re-proof of detail inputs from raw XML
python verify_formulas.py  # EVALUATES the formulas — MUST pass before quoting any figure
```

`extract.py` reuses `_work\raw\*.xml` if present — **delete that folder to force a fresh pull**.
Set the period via `FROM_DATE` / `TO_DATE` in `tally_lib.py`, and `IV_PERIOD` / `IV_STAMP`
environment variables to label the workbook.

### Tally stability — learned the hard way

Large back-to-back requests made Tally **close 4 of its 5 companies mid-run** on 2026-07-31.
`tally_lib.post()` therefore uses a fresh socket per request, `Connection: close`, retry with
backoff, and a 1.5 s pause between calls. Do not remove these. Do not pull whole vouchers with
`AllInventoryEntries.*` across a quarter — that is ~40 MB per company and is what triggered the
drop. Always check the open-company list before and after a run.

### Tally field names — established by probe, not assumed

| Need | Method | Note |
|---|---|---|
| Inward quantity | `CreditTotals` | Tally's stock debit/credit runs **opposite** to intuition |
| Outward quantity | `DebitTotals` | |
| Inward / outward value | `InwardValue` / `OutwardValue` | credit-side, negate |
| Std. Cost rate | `StandardCost` | |
| Costing method | `CostingMethod` | not exposed by the MCP tools |
| Opening/closing | `Opening/ClosingBalance`, `…Value`, `…Rate` | period-scoped by `SVFROMDATE`/`SVTODATE` |

`InwardsQty`, `OutwardsQty`, `InwardQty`, `PurchaseQty` and similar do **not** exist — they return
silently empty, which looks like zero movement. Never trust a zero without checking the tag came
back at all. Tally also emits undeclared `UDF:` namespace prefixes and bare `&`; `clean_xml()`
handles both.

## Folder layout

```
PL and BS\
  Inventory\                                        <- ALL stock workbooks (top level)
      Stock_Valuation_Check_<timestamp>.xlsx            <- current
      SAMPLE LAYOUT - Stock_Valuation_Check.xlsx        <- approved design sample (illustrative figures)
      Old Versions\                                     <- superseded, auto-archived each run
  Stock Report\.claude\skills\inventory-valuation\
      SKILL.md                                      <- this file
      scripts\build_sample_layout.py                <- rebuilds the design sample
      scripts\fetch_inventory_valuation.py          <- the live extract (to be built)
```

Same rule as `fetch-stock-summary`: output goes in **`Inventory\`**, never `Files\`, never inside
`Stock Report\`. The skill and its output stay separate.

## Scope rules (standing)

| Rule | Value |
|------|-------|
| Period | **1-Apr of current FY → today.** Default; only override if the user names a period. |
| Companies | Period code **2627** by default — read `Files\TALLY Company_List.xlsx` and take only rows marked `1` in the `26-27` column. Re-read the file every run; the user updates it. Rows marked "Not To USE" are always excluded. |
| Connection | `localhost:9000`, from `.env` (`TALLY_HOST` / `TALLY_PORT`). Never the old network server `192.168.33.254`. |

### 5 Tally files → 3 legal entities

Detail is reported per **legal entity**, not per Tally file. HO and its Noida branch are the same
legal company and must be combined.

| Legal entity | Tally files (period 2627) |
|---|---|
| **Orange O Tec** | `ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)` [HO] + `ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)` [Noida] |
| **Orange O Tec Enterprises** | `ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)` [HO] + `ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27` [Noida] |
| **Colorix** | `COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)` [standalone] |

**Consequence for elimination — do not get this wrong.** An HO→Noida stock transfer is
*intra-entity*: same legal company. It is not revenue and the margin on unsold transferred stock is
not profit. It eliminates **inside the entity sheet**, one level before group consolidation.
Elimination is therefore two-tier: intra-entity (HO↔Noida) first, then inter-entity.

## Decisions locked with the user (2026-07-31)

| Decision | Locked answer | Why |
|---|---|---|
| Same item held in both HO and Noida | **Separate rows**, distinguished by a `Unit` column | Costing method and carrying rate genuinely differ between the two files; merging hides the very problem being hunted |
| Primary comparison basis | **FIFO leads**; weighted-average shown alongside in every row | Both always present, so either can be argued from |
| Materiality | **Show all items**, flag gaps above ₹5,000 | Ranked + autofiltered, so immaterial items sort to the bottom rather than being hidden |

## Output — 10 sheets, all cross-linked

Every sheet carries a nav bar in row 2 (current sheet in `[brackets]`, all others clickable).

| # | Sheet | Level | Content |
|---|---|---|---|
| 1 | `Summary` | Group | 3 entity rows → group total, all live references to the entity sheets; plus how the items-with-stock population splits |
| 2 | `Segment Summary` | Group | Closing stock by **main stock group**, split across the three entities, with net gap and gap % per segment |
| 3 | `Orange O Tec` | Entity | HO / Noida / Combined roll-up, **by-segment block**, then item detail |
| 4 | `Enterprises` | Entity | as above |
| 5 | `Colorix` | Entity | standalone roll-up, by-segment block, item detail |
| 6 | `Costing Method Audit` | Cross-cutting | **Root cause sheet.** Costing method per item across all 5 files |
| 7 | `Exception Register` | Cross-cutting | Zero-rated stock, nil-value receipts, negative stock, stranded value |
| 8 | `Cross Check` | Cross-cutting | 10 reconciliation tests |
| 9 | `Scope & Method` | Cross-cutting | What was done and — just as important — what was not |

`Landed Cost`, `NRV Exceptions` and `Inter-Branch` from the original design are **not built** —
see "Not yet covered".

### Stock hierarchy (segment view)

`groups.py` fetches the StockGroup tree per company and walks each item's parent chain to the root,
giving `Main Category` / `Sub Category` / `Product Group` / `Sub Group` / `Full Group Path`.
Nesting runs up to **5 levels deep** in these companies; anything past level 4 folds into
`Sub Group` while `Full Group Path` keeps the complete chain. This is what makes the segment view
possible — an item's `Parent` alone is only its immediate group.

### Linking map

| From | To |
|---|---|
| Row 2 nav bar, every sheet | every other sheet |
| `Summary` entity name + "open →" | that entity's sheet |
| `Summary` "Where to look" block | all 9 other sheets |
| Entity sheet `See Also` column | Inter-Branch / NRV / Exception Register |
| Entity sheet "Less: inter-branch…" row | `Inter-Branch` |
| Support sheet `Entity` column | owning entity sheet |
| `Cross Check` `Sheet` column | the sheet each test covers |

### Entity sheet — roll-up block

Columns: `Measure` · `HO` · `Noida Branch` · `Entity Combined`. Rows: Items with stock · Tally
carrying value · FIFO value · Wtd avg value · Gap vs FIFO · Var % · Items flagged ·
*Less: inter-branch unrealised profit* · **ENTITY NET EFFECT**.

### Entity sheet — item detail columns

`Rank` · `Unit` (HO/Noida) · `Main Category` · `Sub Category` · `Item Name` · `UOM` ·
`Closing Qty` · `Tally Rate` · `Tally Value` · `FIFO Rate` · `FIFO Value` · `Wtd Avg Rate` ·
`Wtd Avg Value` · `Last Pur Rate` · `Gap vs FIFO` · `Gap %` · `Costing Method` · `Flag` · `See Also`

Ranked by gap descending, autofiltered, frozen panes below the header.

### Support sheet columns

| Sheet | Columns |
|---|---|
| `Costing Method Audit` | Entity · Unit · Stock Group · Item Name · **Costing Method** · Market Valuation Method · Std Cost Defined? · Std Rate · Closing Qty · Tally Closing Rate · Last Purchase Rate · **Rate Moved in Period?** · Flag |
| `Landed Cost` | Entity · Date · Supplier · Vch No · Item · Qty · UOM · Invoice Value · Invoice Rate · Customs Duty · Freight · CHA/Clearing · Forex Diff · **Landed Value** · **Landed Rate** · Booked to Stock? · Cost Not Capitalised |
| `NRV Exceptions` | Entity · Unit · Item · Closing Qty · UOM · Carrying Rate · Carrying Value · Lowest Arm's-Length Sale Rate · Sale Date · Party · Est. Selling Cost % · NRV per Unit · Shortfall per Unit · Write-down Required · Status |
| `Inter-Branch` | Entity · From (Tally file) · To (Tally file) · Item · Qty Transferred · UOM · Transfer Rate · Transfer Value · Qty Received · **Qty Still in Stock at period end** · Group Cost Rate · Margin per Unit · **Unrealised Profit in Stock** · Status |
| `Exception Register` | Entity · Unit · Date · Voucher Type · Vch No · Item · Party · Qty · Rate · Value · Exception Type · **Why Flagged** · Est. ₹ Impact · **Suggested Action** |
| `Cross Check` | # · Check · Scope · Expected · Computed · Difference · Status · Note · Sheet |

Flag vocabulary drives conditional formatting — red: `FROZEN` `OVERSTATED` `FAIL` `WRITE-DOWN`
`NOT CAPITALISED` `ELIMINATE`; green: `OK` `TRACKS COST` `PASS` `CAPITALISED` `NONE`;
amber: `REVIEW` `PARTIAL` `CHECK`.

## Procedure

### Phase A — Diagnose (root cause)

1. Read every stock item master across all 5 files: **Costing Method**, Market Valuation Method,
   Standard Cost/Price table, opening balance qty/value/rate.
2. Classify each item **fixed-rate vs cost-tracking** by comparing the closing rate against
   purchase rates in the period. Write `Costing Method Audit`.

**This phase decides whether there is a problem at all:**

| If the frozen rate is… | Then | Impact |
|---|---|---|
| the report's *Basis of Values* display setting | books are fine, screen rendering only | nil |
| the item's own **Costing Method** (standard/fixed) | every purchase is written up to the fixed rate in stock | real — inventory **and** P&L overstated |

Do not quote a rupee impact before this is settled.

### Phase B — Rebuild

3. Pull opening balances plus every inventory voucher line for the period, per file.
4. Independently recompute closing value on **three bases** — FIFO, weighted-average,
   last-purchase — in Python. Never read Tally's own rate and call it verified.
5. Diff each basis against Tally's carrying value per item; rank by ₹ gap. Roll HO + Noida into
   entity totals.

### Phase C — Adjust for reality

6. **Landed cost trace.** For import purchases, find customs duty / freight / CHA / forex sitting
   in expense ledgers rather than in stock value. This *raises* true cost and therefore *shrinks*
   any apparent overstatement — it must be done before concluding.
7. **NRV test.** Lowest arm's-length selling rate per item (record date + party so it is auditable)
   less estimated selling cost, vs carrying rate. Lower-of-cost-or-NRV exceptions only.

### Phase D — Eliminate

8. Two-tier: intra-entity HO↔Noida first, then inter-entity. Match outward against the receiving
   file's inward, and measure how much is **still in stock at period end** — only that portion
   carries unrealised profit.

### Phase E — Exceptions

9. Sweep for: near-zero-value issues (Re 1 challans), returns booked at rates unrelated to the
   original sale, stock journals priced at neither cost nor any sale rate, negative stock,
   value-with-zero-quantity, items with no stock group.

### Phase F — Report

10. Build the workbook, archive the previous one to `Inventory\Old Versions\`, then run
    verification below.

## Verification — before quoting any figure

1. `Cross Check` must show **PASS** on tests 1–5 (ties to Tally, quantity roll-forward, FIFO layer
   integrity, entity roll-up = sum of its files).
2. Test 10 is mandatory: **re-prove the top-20 gaps from raw voucher data** by an independent
   recomputation, not by re-reading the same intermediate.
3. Tests 6–7 (negative stock, stranded value) are *expected* to fail where the data is genuinely
   defective — report them, do not suppress them.
4. If the user supplied a Tally screenshot, reconcile at least the closing qty and value of the
   item shown before quoting anything.

Report honestly which tests failed. A failing test is a finding, not something to hide.

### Known behaviours — not bugs

- **Sign flip.** Tally returns stock closing values credit-side (negative). Detect a negative net
  total and multiply by -1 so the workbook reads like the screen. Record it in `Cross Check`.
- **Zero-value items.** Tally hides them on screen; keep them, flagged.
- **Blank group quantities** under mixed UOM, net-zero, or top-level group rows — do not "fix" by
  summing. See the `fetch-stock-summary` skill for the full list.

## Live formulas — and the two traps

Every derived figure in the workbook is a **live Excel formula**, not a pasted value. Item detail
computes `Value @ rate`, both gaps and `Gap %`; entity roll-ups are SUMIFS over the detail on the
same sheet; Summary, Segment Summary and Costing Method Audit reference the entity sheets. Change a
rate in the detail and everything above it recalculates.

`verify_formulas.py` **evaluates** the workbook and compares each roll-up against the same figure
computed in pandas. openpyxl writes formulas but never calculates them, so a workbook can look
perfect and be wrong. This is not optional — it caught both of the following on the first run:

| Trap | Symptom | Fix |
|---|---|---|
| `COUNTIFS(range,"<>")` / `SUMIFS(v,v,"<>")` as a "count non-blank" idiom | Evaluated to **0**; the two headline rows read zero while every other row was right | Count over the always-populated `Unit` column: `COUNTIFS(unit,"*")` per unit, `COUNTA`/`SUM` for the combined column |
| Ungrouped items relabelled only in the segment sheet | Segment total was short by ₹3.69 lakh — the SUMIFS criteria said `(no group assigned)` but the data still held `""` | Write the label into the **data** (`build_workbook.py` sets it before the detail is written), never only into the criteria cell |

General rule: if a label is used as SUMIFS criteria, the identical string must exist in the data.

## The cost benchmark — and its one trap

Closing stock is compared against two independently computed bases:

| Basis | Formula | Meaning |
|---|---|---|
| **Inward Rate** | period inward value / period inward qty | what was actually paid for goods received in the period |
| **Wtd Avg Rate** | (opening value + inward value) / (opening qty + inward qty) | blended cost including opening stock |

**The trap.** A benchmark is only valid where goods arrived *with value*. Items received at nil
value (transfers booked at zero, stock journals, free receipts) give an inward rate of 0, which
makes every rupee of carrying value look like a 100% overstatement. On the first live run this
inflated the apparent gap from ₹28 lakh to ₹5.22 **crore**. `classify.py` therefore requires
`Inward Qty > 0 AND Inward Value > 1 AND Inward Rate > 0.01`; everything else is reported as an
exception, never as a gap. **Do not relax this.**

Read gross over/understatement, not the net — large errors in both directions offset, and a small
net figure hides them.

## Not yet covered

| Gap | Why it matters | What it needs |
|---|---|---|
| **True FIFO** | No FIFO figure is claimed anywhere in the workbook | Lot-level purchase layers from voucher detail (~40 MB/company/quarter — the load that destabilised Tally) |
| **Landed cost** | Would RAISE true cost and REDUCE apparent overstatement | GL-detail extract of duty / freight / CHA / forex |
| **NRV test** | Write-down candidates | Sales voucher detail for arm's-length rates |
| **Inter-branch elimination** | INTRA-entity, so eliminates inside the entity sheet | Voucher party detail to identify transfers |

The workbook's `Scope & Method` sheet states all of this to the reader — keep the two in sync.

## Origin

Created 2026-07-31 from a user review of a Tally *Stock Item Vouchers* screen for
`REACTIVE INK H-SERIES BLACK` (Orange O Tec HO, Q1 FY26-27), which showed the closing rate pinned
at **393.58 on every row** while the period's imports from Han Glory Hong Kong landed at
**306–316/kg**. Quantity flow on that screen reconciled cleanly; only value was suspect. That single
observation is what this skill generalises across every item and entity.

Other observations from the same screen, to be confirmed by the live run — a 10 kg delivery challan
to Bhoomi Fashion at Re 1.00, a sales return booked at 690 against an original sale at 650 to O FAB,
a stock journal priced at 501.03, and 5,100 kgs transferred HO→Noida at 380 (intra-entity, so
eliminable).
