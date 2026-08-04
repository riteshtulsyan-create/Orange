---
name: inventory-ageing
description: Build the Inventory Ageing report — split closing stock into five age buckets by date of receipt (FIFO), add last-movement and months-of-cover analysis, reported at legal-entity level (Orange O Tec, Enterprises, Colorix). Use whenever the user asks for stock ageing, inventory ageing analysis, how old the stock is, non-moving or slow-moving stock, dead stock, months of cover, or says "age the inventory".
---

# Inventory Ageing

Answers: **how old is the stock we are carrying, and is any of it dead?**

Produces `Inventory\Stock_Ageing_Analysis_<YYYY-MM-DD_HHMM>.xlsx`, a 10-sheet cross-linked
workbook at **legal entity** level.

## Two decisions locked with the user (2026-08-01) — do not reopen

| Decision | Rule |
|---|---|
| **No provision modelling** | The report measures age. It does not price it. No provision %, no write-down, no P&L impact anywhere in the workbook. |
| **No "unaged" bucket** | Stock with no purchase layer in the file is dated at the **start of the period that file makes available** and labelled **`Opening`** in the adjacent `Receipt Source` column. Every item lands in a real bucket; the Source column is what stops an assumed date being read as a measured one. |

## Build status

| Component | Status |
|-----------|--------|
| Design, sheet layout, column set, Opening treatment | **Awaiting user approval** — see the sample (draft v2) |
| `scripts\build_sample_layout.py` — design sample (illustrative figures) | **Working** |
| `scripts\extract.py` — the live extract | **NOT BUILT** |
| Everything downstream of the extract | **NOT BUILT** |

Sample: `Inventory\SAMPLE LAYOUT - Stock_Ageing_Analysis.xlsx`. Every figure in it is invented.
Edit the builder, re-run it, and the sample refreshes in place.

## Data sourcing — what was PROVED by probe on 2026-07-31

Probed against **one company only** (`ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)`) to avoid
destabilising Tally. Do not assume any of this; it was measured.

| Route | Request | Result | Verdict |
|---|---|---|---|
| Tally's own ageing screen | `TALLYREQUEST=Export, TYPE=Data, ID=Stock Ageing Analysis` | no response in **>10 min**, connection reset, Tally unresponsive for ~1 min after | **Unusable** |
| Batch collection | `<TYPE>Batch</TYPE>` with `ClosingBalance` / `MfdOn` | returned **0 batch records** (1.4 KB, all CMPINFO zeros) | **Unusable** — items are not batch-tracked, so the "All Batches" ageing on screen is derived internally by Tally, not stored anywhere we can read |
| StockItem `BatchAllocations` | filtered to a single item | **timed out at 90 s** — the filter is applied after the collection is built, so it walks every item anyway | **Unusable as-is** |
| Voucher lines, `AllInventoryEntries.*`, one month, `$$IsInward` filter | Apr-2026 only | **timed out at 240 s** | Feasible only if re-engineered — see below |
| `ID=Voucher Register` report | with `VOUCHERTYPENAME` | returns company info only, **0 vouchers** | Wrong report name; not a valid route |

**Conclusion: there is no cheap path.** The ageing must be rebuilt from inward voucher lines, and
the extract has to be engineered for it — month-sliced, cached to disk, resumable, run in the
background. Budget a slow first run. This is the single biggest reason this skill is more work than
`fetch-stock-summary`.

Incidental finding: Tally reported **6 companies open** after the probe (the 5 expected plus one
more) — check the open-company list before and after every run, as the `inventory-valuation` skill
already warns.

### Each file can only see back as far as it starts — THE hard limit of this report

Ageing is capped by the history the file itself holds. Derive the start date from Tally's company
info (`BooksFrom` / `StartingFrom`), **never by parsing the company name**.

| Legal entity | History from | Max age reachable as at 01-Aug-2026 | Consequence |
|---|---|---|---|
| Orange O Tec | 01-Apr-2025 | 487 days | Cannot reach `Over 2 years`. That bucket is structurally nil. |
| **Enterprises** | 01-Apr-2026 | **122 days** | **Cannot show anything past `91-180 days`.** Three of its five buckets are structurally zero. Stock that has sat for years still reads as 122 days old. |
| Colorix | 01-Apr-2020 | 2,313 days | The only entity where `Over 2 years` is genuinely measured. |

**The three entities' bucket profiles are therefore not comparable.** Say so every time figures are
quoted. For Enterprises, read `Slow & Non-Moving` instead — last-outward-date analysis is unaffected
by how far back the file goes.

`Cross Check` test 12 enforces this: no entity may show value in a bucket its file cannot reach.
The sample builder asserts the same thing before writing.

**The fix, if the user wants it.** `REMOTECMPINFO` shows prior-year files on disk — Orange O Tec
FY2022-23 / 2017-18 / 2015-16, and Enterprises has a `F.Y.2024-26` file that would extend its
history to 01-Apr-2024. Chaining those in converts `Opening`-dated stock into genuinely measured
dates, at the cost of a much longer run. Offer it; do not assume it.

## The three things this report must get right

### 1. The as-at date

The user's Tally screen was run **"as at 31-Mar-27"** — the company file's period end, which is a
**future date**. Ageing from a future date adds ~8 months to every item and inflates every old
bucket. **Always run as at today** (or an explicit period end the user names), never the file's
period end. `Cross Check` test 10 exists solely to enforce this.

### 2. Quantity without value

Straight off the user's screen: `ACID INK BLACK-D` shows **10.000 KGS in the >180 day bucket with
no value**. So do ACID INK BLUE/CYAN/MAGENTA/ORANGE/YELLOW, `50X50 PADDING CHEMICAL-D`,
`BLACK FOILJET INK-D` and many more. "Valued based on actual purchase" means Tally found no
purchase to price them from.

This is a **different problem from an unknown date**. The date is solved by the `Opening` treatment
above. The missing *value* is not — the ageing simply cannot price this stock, so the quantity sits
in its bucket contributing ₹0. The `Unvalued Ageing` sheet lists these items and reports the
coverage ratio (priced value ÷ carrying value), which is the figure that decides how much of the
ageing is real.

### 3. Purchase age is not the same question as movement

An item bought 400 days ago that sells every month is fine. An item bought last month that has
never issued is not. Tally's screen only answers the first. `Slow & Non-Moving` answers the second
using **last outward date** and **months of cover** (closing qty ÷ average monthly consumption).
Both belong in the pack; neither replaces the other.

## Buckets

Tally's screen stops at `>180 days`, which cannot tell 7-month-old stock from 3-year-old stock —
the whole point of the exercise. **Five** buckets, no catch-all:

`0 - 90 days` · `91 - 180 days` · `181 - 365 days` · `1 - 2 years` · `Over 2 years`

Every item lands in one of these. `Receipt Source` (`Purchase` or `Opening`) says whether the date
behind it was traced or assumed. Opening-dated cells are shaded amber on every sheet so an assumed
date is never mistaken for a measured one.

## Scope rules (standing)

| Rule | Value |
|------|-------|
| As-at date | **Today**, unless the user names a period end. Never the file's period end. |
| Companies | Period code **2627** by default — read `Files\TALLY Company_List.xlsx`, take only rows marked `1` in the 26-27 column. Re-read every run. Rows marked "Not To USE" always excluded. |
| Entities | 5 Tally files → 3 legal entities. HO + Noida combined, `Unit` column keeps them separable. |
| Connection | `localhost:9000` from `.env`. Never the old network server `192.168.33.254`. |
| Output | `Inventory\` — never `Files\`, never inside `Stock Report\`. Previous run auto-archived to `Inventory\Old Versions\`. |

## Output — 10 sheets, all cross-linked

Nav bar in row 2 of every sheet (current sheet in `[brackets]`, all others clickable).

| # | Sheet | Content |
|---|---|---|
| 1 | `Summary` | 3 entity rows over 5 buckets, % older than 1 year, value dated from Opening, **history available from**; plus a "read this before reading the buckets" block |
| 2 | `Ageing by Segment` | Main stock group by bucket, split across the three entities |
| 3 | `Orange O Tec` | HO / Noida / Combined roll-up, then item detail |
| 4 | `Enterprises` | as above |
| 5 | `Colorix` | standalone |
| 6 | `Slow & Non-Moving` | Aged by **last outward movement**, with months of cover |
| 7 | `Unvalued Ageing` | Quantity-without-value, and the coverage ratio |
| 8 | `Exception Register` | Dead (`-D`) items, negative stock, stranded value, ungrouped items |
| 9 | `Cross Check` | 12 reconciliation tests |
| 10 | `Scope & Method` | How far back each file sees, what was done, and what was not |

### Item detail columns (entity sheets)

`Rank` · `Unit` · `Main Category` · `Sub Category` · `Item Name` · `UOM` · `Closing Qty` ·
`Closing Value` · `Avg Rate` · then **Qty + Value for each of the five buckets** ·
`Oldest Receipt` · **`Receipt Source`** · `Age of Oldest (days)` · `Last Inward` · `Last Outward` ·
`Days Since Movement` · `Months of Cover` · `Flag` · `See Also`

`Closing Value` is a live formula summing the five bucket values. Ranked by value over one year
old, autofiltered, frozen panes.

Flag vocabulary drives conditional formatting — red: `DEAD` `NON-MOVING` `NEGATIVE` `NO VALUE`
`FAIL` `OVERSTATED`; green: `OK` `PASS` `FRESH` `MOVING` `NONE` `PURCHASE`; amber: `SLOW`
`REVIEW` `CHECK` `PARTIAL` `OPENING`.

### Internal consistency the builder enforces

`build_sample_layout.py` asserts these before writing, and the live build must too. Each one
already caught a real error in the sample:

| Rule | What it caught |
|---|---|
| No item's quantity may sit in a bucket older than its own oldest receipt date | 3 rows placing stock in buckets the item's date could not reach |
| No entity may show value in a bucket its file cannot reach | Enterprises carrying `181-365` / `1-2 yr` / `>2 yr` value from a file starting 01-Apr-26 |
| Segment matrix must tie on **both** margins — row = segment total, column = group bucket total | hand-written segment figures that tied on neither |
| Stated age in days must equal `as-at − oldest receipt` | — |
| `Receipt Source` is a closed vocabulary: `Purchase` or `Opening` only | — |

The segment matrix is fitted by iterative proportional fitting, so each segment keeps a believable
ageing shape while both margins stay exact.

## Procedure (once the extract is built)

1. Read `Files\TALLY Company_List.xlsx`, resolve the 2627 company set, confirm all are open.
2. **Read each company's own start date from Tally** (`BooksFrom` / `StartingFrom`). This is the
   Opening date and the ageing cap for that file. Never parse it from the company name.
3. Pull closing quantity and value per item per company (reuse `inventory-valuation\extract.py`
   field names — they were established by probe, see that skill).
4. Pull inward voucher lines **month by month**, cache each slice to disk, resume on failure.
5. Build FIFO layers per item; consume closing quantity against layers **newest-first**, since
   what remains in stock under FIFO is the most recent receipts. Any residual quantity not covered
   by a layer is dated at the file's start with `Receipt Source = Opening`.
6. Pull last outward date and 12-month issue quantity per item for the movement analysis.
7. Walk the stock-group tree for the segment view (reuse `inventory-valuation\groups.py`).
8. Build the workbook, archive the previous one, then verify.

## Verification — before quoting any figure

1. `Cross Check` tests 1-5 must **PASS**: the five buckets sum to Tally closing value; ties to the
   latest `Stock_Summary_ALL_COMPANIES` workbook; ties to the Balance Sheet inventory line; bucket
   quantities sum to closing quantity on every item; entity = sum of its files.
2. Tests 11-12 must **PASS**: no date precedes its own file's start, and no entity shows value in a
   bucket its file cannot reach.
3. Test 9 is mandatory: **re-prove the top-20 ages from raw voucher XML** by independent
   recomputation, not by re-reading the intermediate.
4. Tests 6-8 (negative stock, stranded value, ungrouped items) are *expected* to fail where the
   data is genuinely defective. Report them; do not suppress them.
5. Run a formula-evaluation pass (as `inventory-valuation\verify_formulas.py` does). openpyxl
   writes formulas but never calculates them — a workbook can look perfect and be wrong.
6. Reconcile at least one item against the user's Tally screen before quoting anything.
7. **Never quote a cross-entity bucket comparison without stating the history caps.** Enterprises
   showing 0% over one year is an artefact of its file, not a fact about its stock.

## Not yet covered

| Gap | Why it matters |
|---|---|
| **Provision / write-down** | Excluded at the user's instruction. This report measures age; it does not price it. NRV work lives in `inventory-valuation`. |
| **Inter-branch transfers** | An HO→Noida transfer resets the receipt date at Noida, so transferred stock looks younger than it is. Quantified in `inventory-valuation`. |
| **Landed cost** | Bucket values are invoice cost, not landed cost — duty / freight / CHA sit in expense ledgers, so bucket values understate true carrying cost. |
| **Layers predating each file** | Recoverable only by opening the archived prior-year files. Until then that stock is dated `Opening`. Enterprises is the acute case: its FY2024-26 file would take history back two more years. |

Keep this list and the workbook's `Scope & Method` sheet in sync.

## Origin

Created 2026-07-31 from a user screenshot of Tally's *Stock Ageing Analysis* for
`ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)`, items under `PRINTING INK`, as at 31-Mar-27.
That screen showed the three problems this skill exists to fix: a future as-at date, three coarse
buckets, and roughly two-thirds of the visible rows carrying quantity with **no value at all**.
