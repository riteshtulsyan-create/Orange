---
name: gp-segment-mom
description: Build or refresh the month-on-month Gross Profit by segment report (Machines / Heads / Ink / Spare Parts / Other) for all five Tally companies, straight from Tally Prime. Use this skill whenever the user asks for GP by segment, segment-wise gross profit, month-on-month GP from Tally, or says they want the GP report rebuilt from live Tally data rather than from the hand-maintained GP REPORT.xlsx.
---

# GP by Segment — Month on Month

Produces `Files\GP_Segment_ItemWise_<YYYY-MM-DD_HHMM>.xlsx` from live Tally, split into
five segments: **Machines, Heads, Ink, Spare Parts, Other**.

## How to run — use the item-wise route

```powershell
cd "C:\Users\account2\OneDrive - Orange O Tec Private Limited\00 Claude Project\PL and BS"
python ".claude\skills\gp-segment-mom\scripts\fetch_gp_data.py"      # ledgers + stock groups
python ".claude\skills\gp-segment-mom\scripts\fetch_gp_vouchers.py"  # voucher inventory lines
python ".claude\skills\gp-segment-mom\scripts\build_gp_itemwise.py"  # the workbook
```

The build auto-reconciles against the **newest** `Files\PL_BS_GL_Detail_*.xlsx`; override with
`--pl <path>`. Sheet order is Index, Basis & Method, Consolidated, by Entity, by Company,
Segment Detail, Board Reconciliation, Cross Check.

**Figures are linked, not retyped.** The by-Entity sheet is written first and the Consolidated
sheet is built as live cross-sheet sums of it, so Consolidated always equals the sum of the
three entities. The Index headline links to Consolidated, and the Board Reconciliation
sub-totals are formulas.

Never start a label cell with `=`. Excel parses it as a formula and renders `#NAME?` — this
bit the reconciliation sub-total labels once already. Verify with the `formulas` engine, not
cached values.

All five FY26-27 companies must be **open in Tally** first. The fetches take a few
minutes; both cache to JSON (`Files\gp_segment_cache.json`, `Files\gp_voucher_cache.json`)
so the workbook can be rebuilt without re-querying Tally. Options: `--through YYYY-MM-DD`
(default today), `--host`, `--port`.

`build_gp_workbook.py` is the older **ledger-based** build. It is superseded — it could
not segment Enterprises or Colorix and produced negative Machines margins. Keep it only
for reference.

## Method

Segment comes from each **stock item's own group**, not from the ledger name. That is the
whole point: only Orange O Tec tags its sales ledgers by segment, so a ledger-driven
build leaves half of group revenue unsegmentable.

```
cost rate = (opening value + purchase value) / (opening qty + purchase qty)
COGS      = sales qty * cost rate
GP        = sales value - COGS - segment discounts
```

Fallbacks when a period has no opening stock and no purchases: year-to-date weighted
average, then closing rate. The Cross Check sheet reports the share of sales resting on
each basis (currently 98.1% on the period average).

Segment-tagged discount and rate-difference ledgers are netted off revenue — they never
appear on an inventory line. This matters: Q1 head discounts are about ₹63 lakh, and
applying them moved Heads materially.

Quarter columns are the **sum of their months**, not the independently fetched quarter. A
quarterly weighted-average cost rate is not the sum of three monthly ones (~2% gap on GP);
summing keeps the columns adding up. The independent quarterly fetch is retained purely as
a cross-check.

## Five traps that will silently corrupt this

1. **Credit notes carry a positive quantity with a negative amount** — `2.000 PCS` against
   `-676700`. Left alone, a sales return ADDS to quantity sold and so adds COGS. This drove
   Spare Parts to −7.7% when the true figure is +21.9%. Negate quantity on Credit Note and
   Debit Note voucher types.
2. **Proforma invoices and delivery challans are not sales.** Their voucher types sit under
   `Sales Order` and `Delivery Note`. Counting them overstated Orange O Tec HO April revenue
   by about ₹8.8 crore. Classify by walking the voucher type's PARENT chain to its root.
3. **Tally's default voucher types have been renamed.** In Orange O Tec HO the base Sales
   type is `Sales Accounts-HSS` and Purchase is `GST PURCHASE`, and **each is its own
   parent**. A resolver that assumes the root is literally "Sales", or that does not detect
   the self-parent cycle, silently classifies every sale as unknown.
4. **Never use item `OutwardValue` as sales.** It includes branch transfers, godown moves
   and manufacturing issues — about 1.8x real sales for Orange O Tec HO.
5. **Forex purchases return a formatted string**, not a number:
   `-$15500.00 @ ₹95/$ = -₹1472500.00`. Parse what follows the final `=`. Missing this
   produced a 171% gross margin in testing. Tally's signs are also inverted — inward,
   opening and closing are negative (debit), outward positive.

`fetch_ledger_balances` in the `fetch-pandl` skill does **not** work — Tally rejects its
`FORMULACOLLECTION` shape with "Unknown Request". Use `_collection_xml` from
`fetch-stock-summary`, which is what this skill does.

## Always check before reporting

Open the **Cross Check** sheet. It verifies that quarter columns equal their months, that
each company's voucher sales reconcile to its sales ledgers, how much of COGS rests on a
fallback rate, and whether each segment's GP% is credible.

Two entries are expected to stay flagged:

| Flag | Why |
|---|---|
| OOT Enterprises HO: voucher vs ledger sales | Non-inventory revenue (software, services) never appears on an inventory line |
| OTHER: GP% credible? | Paper, packing and software carry cost with little matched revenue. Only ~0.1% of group sales — immaterial, but do not quote it |

## Known limitation

Only **Orange O Tec** tags sales ledgers by segment (HO 77%, Noida 100%); Enterprises and
Colorix tag none. The item-wise route works around this entirely by reading the stock item,
so segment figures are now **measured, not allocated** — but the underlying Tally hygiene
gap remains. Creating segment-suffixed sales ledgers in Enterprises and Colorix would make
the ledger and item routes reconcile directly. See the `project_gp_segment_tagging_gap`
memory.

Quantity is summed in each item's own billing unit. Units differ between items (NOS, PCS,
KGS, MTR), so a segment quantity total mixes units — use it for trend, not level.

## Editing the segment mapping

`scripts\segment_map.py` holds stock-group → segment, ledger → segment, and company-specific
overrides (Enterprises' "Finished Goods" and "Manufacturing Stock" are finished ink and ink
WIP, not machines). Items resolve leaf → root, stopping at the first explicit match, so
container groups like Orange O Tec Noida's "MACHINERY & MACHINERY PARTS" don't collapse two
segments into one.
