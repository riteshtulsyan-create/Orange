# CFO Pack: 6 MOM sheets + Cash Flow + Fund Flow + CFO Analysis (full re-audit, v3)

## Context
Management wants a CFO-ready workbook for the coming meeting. It is built from the live-Tally workbook
`Files\Orange FS\PL_BS_GL_Detail_2026-10-16_1523.xlsx` and holds **only** these six sheets:
MOM PL GL CFS, MOM BS GL CFS, MOM Ratio Analysis CFS, MOM PL GL SFS, MOM BS GL SFS and MOM Ratio Analysis SFS.
It adds **Cash Flow** and **Fund Flow** sheets (CFS and SFS) with the same column and row grouping, plus **one analytical
sheet** with the major points the CFO needs. The row design follows the CFS Cash Flow / Fund Flow sheets in
`Files\Results\Quarterly_PL_BS_FY25-26_vs_FY26-27_2026-10-17_1456.xlsx`, a format management already uses.

User decisions so far (2026-10-24):
- Use the 16-Oct file **as-is**. No full refresh; this is a stated scope-down from CFO mode.
- Opening balances at 31-Mar-26 come **from ConnectWave** (the read-only Supabase copy of Tally). Surat comes from the
  saved Tally pull of 13-Oct, because ConnectWave can't serve it.
- **New file**, with CF and FF for **CFS and SFS**.
- Copy only the needed scripts from the stash; Ritesh-Fund-Flow stays untouched.
- **New (v3):** be ready for the Orange O Tec Tally company being **split into a FY 25-26 file and a FY 26-27 file**, and
  check the figures against the **Ritesh-EXP expense report**.

---

## A. Re-audit 1: is it useful for management and the CFO?
The first draft was too heavy to brief from: 88 columns of cash flow and nothing on one page. These changes are now in the plan:
1. **One-page CFO summary first** (top of CFO Analysis, ₹ crore): 6 headline numbers, 5 red flags, 5 likely board
   questions with suggested answers, and 3 data caveats. Detail sits below.
2. **Adjusted PAT (labelled estimate).** Depreciation and tax are not booked for Apr-Sep, so the ₹12.44 cr PAT is
   overstated. The estimate is half of FY 25-26 depreciation plus tax at 25.17%, and the reasoning is shown.
3. **SFS for entity performance, CFS for the group.** CFS entity columns shift profit between entities, so they are
   never quoted as entity results. A note on the analysis sheet says so.
4. **Cash conversion story.** Operating cash flow against PAT, showing where the cash went (debtors, supplier paydown, stock, capex) and how it
   was funded (OD and loans).
5. **Q2 vs Q1 bridge plus the top 10 GL movers**, so the margin drop can be explained line by line.
6. **Distorted months flagged**: the inventory provision swings April and June.
7. **Checks block**: every control is shown as nil or not nil, so the CFO can see the pack ties.
8. **Expense tie-out to the Ritesh-EXP report** (section D), so two reports in the same meeting cannot disagree unexplained.

## B. Re-audit 2: cross-check of every finding (re-verified from the workbook; ₹ crore)
| # | Finding | How checked | Result |
|---|---------|-------------|--------|
| 1 | Q2 turnover ₹74.46 cr against Q1 ₹55.14 cr (+35%); EBITDA ₹6.08 cr against ₹8.81 cr (−31%); PAT ₹4.92 cr against ₹7.52 cr; GP% 20.5 against 27.2 | PL CFS Grand Total, Q1/Q2 | ✔ |
| 1b | Sep has the highest sales (₹32.57 cr) and the lowest EBITDA (₹1.44 cr) and PAT (₹0.27 cr); opex ₹1.84 cr in Apr, ₹3.40 cr in Sep | PL CFS monthly | ✔ |
| 2 | Apr-Sep: turnover ₹129.6 cr, GP ₹30.28 cr (23.4%), EBITDA ₹14.89 cr, finance cost ₹1.97 cr, PAT ₹12.44 cr, with no depreciation or tax booked | PL CFS totals; the depreciation and tax rows are 0 | ✔ |
| 3 | CFS entity PAT (O Tec −₹11.13 cr, Colorix GP 90%) is an elimination effect. SFS PAT: O Tec ₹3.44 cr, Ent ₹8.77 cr, Colorix ₹0.23 cr (Q2 −₹0.20 cr). Ent EBITDA ₹6.47 cr → ₹2.22 cr; Ent GP 47% against O Tec 14% | PL SFS against CFS; SFS Grand Total equals CFS Grand Total | ✔ |
| 4 | The BS P&L line moved **less** than PAT in May (₹0.08 cr) and June (₹1.62 cr); Jul to Sep tie. Last year's P&L appears to have been reduced after year end | Month-on-month change in `TOTAL: Profit & Loss A/c` against PAT | ✔ (corrected from v1) |
| 5 | Debtors ₹59.3 cr → ₹77.2 cr (₹13.2 cr in Sep); creditors ₹52.4 cr → ₹40.7 cr; stock ₹43.7 cr; debtor days about 109 | BS CFS totals | ✔ |
| 6 | OD ₹31.7 cr → ₹44.9 cr; secured loans ₹3.1 cr → ₹16.3 cr; gross debt ₹66.6 cr against equity ₹69.3 cr (D/E about 0.96); cash ₹1.56 cr; interest cover 7.6x | BS CFS totals | ✔ |
| 7 | Inventory overhead provision ₹5.08 cr in Apr, −₹5.54 cr in Jun | PL rows 442–446 | ✔ |
| 8 | SFS additivity and BS gap nil; CFS BS gap ₹9,240; inter-company profit eliminated ₹0.8 lakh | The workbook's Cross Check | ✔ |
| 9 | The quarterly workbook's saved totals read 0.0 in places, so recalculate before quoting it | data_only read | ✔ |
| 10 | ConnectWave 1-Apr-26 opening agrees with the Tally 31-Mar-26 pull for 95–99% of ledgers; it can't serve Surat | Live read-only queries | ✔ |
| 11 | **New:** the expense report ties at group level, with small Q2 timing gaps (section D) | Expense Sheet 2026-10-10 against the GL file | ✔ |

## C. Re-audit 3: Orange O Tec Tally company split (FY 25-26 file / FY 26-27 file)
Today Orange O Tec Surat is one file spanning two years, `ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)`, and Noida
is `...-NOIDA-(from 1-Apr-25)`. If either is split the way Enterprises already was, a wrong pick **fails silently**: the wrong file
returns figures without any error (CLAUDE.md, the 2026-08-04 incident). The build will guard against this:
1. **Company names are resolved, never hard-coded.** The FY 26-27 file per entity comes from `Files\TALLY Company_List.xlsx`, column
   **2627** (the source of truth), and FY 25-26 from column **2526**. They are matched by exact name against ConnectWave `v_company`, where a
   prior-year file shows as a tenant with a `~2025…` suffix, as Enterprises does today.
2. **Opening 31-Mar-26 for a split entity** = the **FY 26-27 file's `opening`**, its carried-forward balance.
   **Check:** compare it ledger by ledger with the FY 25-26 file's **closing at 31-Mar-26**. Any difference is a carry-forward
   error. It is listed on the Opening sheet with its value and **never plugged**.
3. **No double counting.** FY 26-27 months come only from the FY 26-27 file, even when both files are open.
4. **Not split yet (today):** Surat opening comes from the saved Tally pull of 13-Oct (`bs_q4_2526_ORANGE_O_TEC_PRIVATE_LIMITED.csv`),
   and Noida, Ent, Ent Noida and Colorix from ConnectWave, as agreed. The script detects a split automatically and switches to rule 2.
5. **Pre-flight stop.** If the 2627 company list doesn't match what ConnectWave holds exactly, for example a new
   name or a missing file, the build stops and names the mismatch rather than guessing.
6. **The 6 copied sheets are unaffected.** They are copied from the 16-Oct file. The split only matters for the opening
   balances and any future refresh; the fetch-pandl skill's `company_for()` / `PRIOR_YEAR_COMPANIES` handles that the same way,
   with a note added there.

## D. Re-audit 4: tie-out to the Ritesh-EXP expense report
The git branch `Ritesh-EXP` has no commits of its own (same commit as `main`). The **report** is
`Files\EXPENSES REFERESH\Expense Sheet 2026-10-10 12-40-56.xlsx` (Apr-Sep, pulled 10-Oct). The TB-proven Hub audit is
`Hub_vs_Tally_Audit_2026-10-23_1335.xlsx`. Comparison with the GL file (pulled 16-Oct), in ₹ lakh, Apr-Sep, group:

| Item | Expense Sheet (10-Oct) | GL file (16-Oct) | Gap | Comment |
|------|-----------------------:|-----------------:|----:|---------|
| Custom duty | 318.77 | 318.77 | 0 | Ties exactly |
| Exchange gain/(loss), net | −0.78 | −0.78 | 0 | Ties exactly |
| Operating expenses (indirect excl. finance), **Q1** | 619.34 | 619.3 | 0 | Ties |
| Operating expenses, Q2 | 925.9 | 928.7 | +2.8 | Mostly Sep (+2.3); entries posted 10–16 Oct |
| Finance cost, Sep | 35.46 | 28.8 | −6.6 | Changed after 10-Oct (e.g. NBFC interest −1.6 in Sep in GL); to be named in the build |
| Inventory valuation, Sep | 96.71 | 88.8 | −7.9 | Provision revised after 10-Oct |
| Entity split, opex (O Tec / Ent) | 917.5 / 626.3 | 908.1 / 638.5 | −9.4 / +12.3 | Offsetting (Noida mapping?); group nets to +2.8 |

**Advice:** at group level the two reports **match, apart from timing**. Q1 ties to the rupee, and Q2 differs by about ₹2.8 lakh of opex,
₹6.6 lakh of finance cost and ₹7.9 lakh of valuation, all from postings made between the 10-Oct and 16-Oct pulls. Two cautions for the CFO:
(a) the Hub audit shows the old Expense Sheet method has known defects: it reverses the sign of 328 credit lines and counts
non-accounting vouchers. **The TB-based GL file is the one to quote.** (b) The entity-level split differs by about ₹10–12 lakh each way;
the build will trace it, most likely to how Noida ledgers are mapped. The build adds an **"Expense tie-out" block** to CFO Analysis
(main group × month: GL against Expense Sheet, with the gap and reason). Values are read from the Expense Sheet and labelled with its date.

---

## Output: two download files, Excel and PDF (user, 2026-10-24)
1. **Excel:** `Files\Results\CFO_Pack_PL_BS_CF_FF_<YYYY-MM-DD_HHMM>.xlsx`, the full working pack (sheets below).
2. **PDF:** `Files\Results\CFO_Pack_PL_BS_CF_FF_<same stamp>.pdf`, the **board-ready print**. The 124-column GL sheets
   don't print legibly, so the PDF holds the pages a CFO reads. Each page is A4 landscape, fit to width, with a header (entity, period,
   "₹ crore", source and pull date) and page numbers in the footer:
   - p1: one-page CFO summary (headline numbers, red flags, board questions and answers, caveats)
   - KPI strip; Q2 vs Q1 bridge with the top 10 GL movers; entity view (SFS)
   - Condensed P&L and BS: group lines × (Q1, Q2, Apr-Sep, and Sep month) for Total O Tec, Total Ent, Colorix and Grand Total
   - Condensed Cash Flow and Fund Flow (CFS and SFS): same columns, with the check rows shown at nil
   - Expense tie-out to the Ritesh-EXP report; Checks block; Opening BS reconciliation summary
   These pages are built as **print sheets inside the Excel file** (`PDF - …`, formulas linked to the detail sheets), so the PDF
   and the Excel can never disagree. The PDF is exported **by Excel itself** (COM `ExportAsFixedFormat` on the print sheets,
   after a full recalculation), so every figure in it is calculated. If Excel automation is unavailable, the fallback is LibreOffice headless, and the
   report says which one was used.
3. After export, verify: page count, no `#REF!` / `#VALUE!` / `####` on any page (text extracted from the PDF), and the p1 headline
   numbers match the Excel values.

The standard `Files\Orange FS\` file is not touched.
1. **Index**
2. **CFO Analysis**: one-page summary, KPIs, Q2/Q1 bridge, entity view (SFS), cash story, expense tie-out, Checks
3. The **6 MOM sheets**, copied as-is. They reference only each other, which I checked.
4. **MOM Cash Flow CFS / SFS**, **MOM Fund Flow CFS / SFS**
5. **Opening BS 31-Mar-26**: SFS and CFS in the same 8 columns, with a source column, the ConnectWave-vs-Tally differences and the split-company check

### Cash Flow and Fund Flow layout
- There are 11 period blocks (Apr, May, Jun, Q1, Jul, Apr-Jul, Aug, Apr-Aug, Sep, Q2, Apr-Sep), each with 8 columns (Surat, Noida, Total O Tec,
  Ent, Ent Noida, Total Ent, Colorix, Grand Total). Headers are copied from the PL sheet's rows 3 and 4.
- Flow = BS at period end − BS at period start. The start is the previous month, or the Opening sheet for Apr, Q1 and Apr-*; Q2 starts at Jun.
  PBT comes from the matching PL column. All cells are live formulas.
- **Cash Flow:** A Operating, B Investing, C Financing (as in the quarterly file), **D prior-year P&L adjustments** (finding 4),
  **E unmatched inter-company** (CFS entity columns), then net change, opening and closing cash, and a **check (must be nil)**.
- **Fund Flow:** Sources / Applications / working-capital schedule (OD as a current liability, CMA style), with a **check (must be nil)**.
- BS rows are found **by label**, never by row number, because SFS has Unsecured Loans and Branch / Divisions and CFS doesn't.

### Opening BS 31-Mar-26: sources
- **Noida, Ent, Ent Noida, Colorix:** ConnectWave `v_ledger_detail.opening` via the read-only
  `..\Orange One Supabase Connect\tools\supabase_client.py` (`select`, ordered by the unique `guid`), mapped to the MOM BS group labels.
- **Surat:** `gl_data\bs_q4_2526_ORANGE_O_TEC_PRIVATE_LIMITED.csv` (Tally, 13-Oct). If split, the rule in section C applies.
- **Opening stock:** the April "Opening Stock" row of MOM PL GL SFS, already in the file.
- **b/f P&L:** the balancing figure, shown beside the ConnectWave and Tally values.
- **CFS opening:** SFS less the inter-company and branch ledgers in the BS CFS memo block, at their 31-Mar values. If that list is incomplete,
  fall back to `build_mom_gl_detail.py` restored from the stash. Any residue goes to line E.

## Steps
1. Copy from `stash@{0}` only what the build turns out to import (expected: nothing; `build_mom_gl_detail.py` as the fallback).
2. New `.claude\skills\fetch-pandl\scripts\build_cfo_pack.py`:
   - Company pre-flight (section C).
   - Read a copy of the 16-Oct workbook and keep the 6 sheets.
   - Build Opening, the 4 flow sheets, and CFO Analysis with the expense tie-out, using PL-sheet styles. Save with a timestamp, never overwriting.
3. New `verify_cfo_pack.py`, using the `formulas` engine, not saved values:
   - Cash check is nil (every SFS column, CFS Grand Total).
   - The FF working-capital schedule equals Sources − Applications.
   - Months add to Q1, Q2 and the cumulative columns.
   - The copied sheets recalculate to the source values.
   - Opening + April movement equals the 30-Apr BS.
   - Each CFO Analysis figure equals its source cell.
   - The expense tie-out gaps are explained.
4. Add a short note to the fetch-pandl SKILL.md on the company-split rule (section C).
5. Report: file path, check matrix, findings, the one-page CFO summary in chat, ConnectWave-vs-Tally differences and the expense
   tie-out. No commit unless asked.

## Critical files
- Source (read a copy): `Files\Orange FS\PL_BS_GL_Detail_2026-10-16_1523.xlsx`
- Layout reference: `Files\Results\Quarterly_PL_BS_FY25-26_vs_FY26-27_2026-10-17_1456.xlsx`
- Company list: `Files\TALLY Company_List.xlsx` (columns 2627 / 2526)
- ConnectWave: `..\Orange One Supabase Connect\tools\supabase_client.py`
- Surat opening and cross-check: `.claude\skills\fetch-pandl\gl_data\bs_q4_2526_*.csv`
- Expense report: `Files\EXPENSES REFERESH\Expense Sheet 2026-10-10 12-40-56.xlsx`, `Hub_vs_Tally_Audit_2026-10-23_1335.xlsx`
- New: `.claude\skills\fetch-pandl\scripts\build_cfo_pack.py`, `verify_cfo_pack.py`

---

## v4 - changes after the first Hub review (user, 2026-10-24)

### Already delivered (v3.1 - v3.3)
- Hub page `/reports/finance-cfo-pack` on Orange Hub branch `Ritesh-PL` (worktree folder `Orange Hub-Ritesh-PL`),
  local server `http://127.0.0.1:5174` (bound to IPv4 - the IPv6-only bind did not open in the user's browser).
- Files published by `publish_cfo_pack_hub.py` into the gitignored `frontend/public/reports/cfo-pack/`.
- MOM figures with a view selector (Group CFS / Orange O Tec / Enterprises / Colorix), quarter column folding,
  GL drill-down (line -> sub-group -> ledger), Excel MOM Summary with the B3 dropdown and outline grouping.
- Expense tie-out and checks grouped and hidden as Annex A / B; expense links removed (user).

### v4 requests
| # | Request | Where | How |
|---|---------|-------|-----|
| 1 | Freeze the headings while scrolling | Hub | Sticky table header (and first column) inside a scroll box |
| 2 | Summary and red flags BELOW the financials | Hub | Financials tab first; briefing on its own tab (see 6) |
| 3 | One more view: **Orange + Enterprises (CFS)** | Hub, Excel, PDF | 5th MOM view = CFS Total O Tec + Total Enterprises columns (Colorix excluded) |
| 4 | Stock lines -> **Change in inventory** | Hub, Excel | Rename the drill-down sub-group "Stock Valuation (Inventory)" to "Change in inventory" (opening - closing) |
| 5 | **Expand all** rows button | Hub | Next to "Collapse all rows" |
| 6 | Separate tabs: **Branch** (branch sales / purchases + branch BS balances) and **Related party** (RP P&L flows + RP BS balances) | Hub, Excel | From the SFS P&L branch section, the CFS P&L / BS "ELIMINATED ON CONSOLIDATION" memo blocks - per company, per month |
| 7 | CFO summary + point 6 (board Q&A, caveats) in **one tab only**, and **not in the PDF** | Excel, Hub | New sheet `CFO Briefing`; removed from CFO Analysis; PDF prints only the financial pages |
| 8 | **Directors' executive summary** - a separate report linked to the figures | Excel, PDF, Hub | Sheet `Directors Summary` (one page, every number a formula), its own PDF `Directors_Executive_Summary_<ts>.pdf`, and a Hub tab |
| 9 | **% of sales** for each month and quarter, with a toggle at the top | Hub, Excel | Hub: "Amounts / % of sales" toggle on the MOM card; Excel: % of sales block under the MOM Summary |
| 10 | **Ratios** month by month | Hub, Excel, PDF | GP%, EBITDA%, PAT%, opex % sales, finance % sales, interest cover, debtor / inventory / creditor days, current ratio, gross debt / equity - per view |
| 11 | Save the plan | - | This file; copy kept at `.claude\skills\fetch-pandl\CFO_PACK_PLAN.md` |

### Rules kept
- Every Excel figure stays a live formula; verify_cfo_pack.py must still pass, with new checks: the Branch tab nets to
  the SFS branch section, the Related-party tab equals the CFS memo blocks, and the Orange + Enterprises view equals
  CFS Total O Tec + Total Enterprises.
- Hub data stays in the gitignored public folder; nothing committed until the user asks.

### v4 status - DELIVERED 2026-10-24 16:12
Pack `Files\Results\CFO_Pack_PL_BS_CF_FF_2026-10-24_1612.xlsx` + `.pdf` (16 pages) + `Directors_Executive_Summary_2026-10-24_1612.pdf`
(1 page). verify_cfo_pack.py: ALL CHECKS PASSED, incl. the new ones (branch and related-party P&L net to nil; Orange +
Enterprises = CFS Total O Tec + Total Enterprises; directors' PDF one page). Published to the Hub; tsc clean.
New code: `cfo_pack_extras.py` (CFO Briefing, Directors Summary, Branch Transactions, Related Party).
Note for the CFO: the Orange + Enterprises (CFS) view shows external trade only - profit on goods O Tec sells to Colorix
(e.g. high-seas sales) sits with Colorix in consolidated columns, so its EBITDA (Rs 3.15 cr) is well below the two
companies' standalone EBITDA (Rs 13.70 cr).

---

## v5 - balance-sheet schedules, ageing and inventory (user, 2026-10-24 evening)

### Delivered 2026-10-24 17:44 (pack `CFO_Pack_PL_BS_CF_FF_2026-10-24_1744`, ALL CHECKS PASSED)
| # | Request | Done |
|---|---------|------|
| 1 | Branch / Related party: net impact nil, green / red | Net row + badge; Surat / Noida buttons removed |
| 2 | O Tec vs Enterprises tie-up | Sub-tab under Related party: sales, purchases, balances, difference. Sep balance gap Rs 9,240 (reimbursement posted by O Tec only) |
| 3 | Rows like Financials; Amount + %; period total; one-line heading | Collapsible rows, 3-way mode, "Period total" picker (flows add up, balances = last month) |
| 4 | Analysis by view: CFS, Enterprises, O Tec, then company-wise | View buttons on the Analysis tab (Hub) |
| 5 | Debtors & Creditors tab (party-wise under group) | Sheet + Hub tab, ties to BS (82.73 / 40.20 cr) - **ageing columns pending (phase 2)** |
| 6 | Taxes tab: GST, TDS, others | Sheet + Hub tab; payable +, receivable -, net |
| 7 | Employees tab: advances and payable, not in creditors / L&A | Sheet + Hub tab (employee advances, creditor for employee / expense claims, gratuity) |
| 8 | Loans & Advances tab | Sheet + Hub tab (L&A excl. employees, prepaid, other L&A, partner, deposits) |
| 9 | Month close: roll forward, reasons for increases | "Change in month" toggle + "Biggest movements, Sep against Aug" list on every balance table |
| - | Bug: "Estimated tax" row showed 293.7% | Fixed - only "...%" / "% of" labels are percentages |

### Phase 2 - needs the five FY 26-27 companies OPEN in Tally (localhost:9000)
1. **Debtor / creditor ageing at 30-Sep by invoice date**: 0-180 and >180 days per party (bill-wise outstanding from Tally,
   as at 30-Sep); columns added to the Debtors & Creditors tab; ties to the party balance.
2. **Directors' summary**: debtors < 180 and > 180 days.
3. **Inventory tab**: closing stock at 30-Sep by stock group (fetch_stock_summary.py, `--to 2026-09-30`), ties to the BS stock.
4. **Inventory ageing tab** by category - Machines, Heads, Spare parts, Ink, Paper & packing material, Others - FIFO on
   receipt date (stock-group -> category map to be confirmed with the user).
ConnectWave cannot give 30-Sep figures (snapshots are "today"), so Tally is the source.

### Month-end roll-forward (standing)
When a month closes: refresh the GL workbook (fetch-pandl CFO mode), rebuild the pack, and the period columns move on;
the "Biggest movements" list then explains the new month against the previous one. The pack's period list (Apr-Sep)
is fixed in the builder today and needs extending when October is added.

### v5.1 delivered 2026-10-24 18:14 (pack CFO_Pack_PL_BS_CF_FF_2026-10-24_1814, ALL CHECKS PASSED)
- Units switch on every Hub tab: Crores / Lakhs / Absolute.
- Ageing AS ON DATE (user: "develop the ageing report as on date") from ConnectWave via new fetch_ageing_snapshot.py (read-only, ~70 s) -> Files/Results/Ageing_Snapshot_<ts>.json; ageing_pack.py builds sheets Debtors Ageing, Creditors Ageing (0-180 / over 180 by invoice date, party-wise under group), Inventory (category / stock group), Inventory Ageing (FIFO on receipt date, 0-90 / 91-180 / over 180, filter by category).
- Hub: Debtors & Creditors tab = Ageing + Month-end balances; new tabs Inventory and Inventory ageing (category buttons).
- Directors' summary: debtors 0-180 / over 180 (24-Oct: 82.37 / 18.54 of 100.90 cr, 18.4%%) - still one page.
- Category map: Machines = MACHINERY, DIGITAL PRINTING MACHINES STOCK; Heads = MACHINERY HEAD; Spare parts = ...PARTS / SPARE; Ink = PRINTING INK, FINISHED GOODS, MANUFACTURING STOCK; Paper & packing = PAPER ROLL, PACKING MATERIAL STOCK; Others = rest. Awaiting user confirmation.
- Before each rebuild: run fetch_ageing_snapshot.py so the as-on date is current.
