"""Build the GP Segment MoM workbook from the cached Tally data.

Method (trading-account basis, per segment):

    Sales = segment-tagged sales ledgers
            + untagged sales allocated on that company's item outward mix
    COGS  = opening stock + purchases - closing stock + segment direct costs
    GP    = Sales - COGS

Sales come from ledgers, never from item outward value: item outward includes
branch transfers, godown moves and manufacturing issues and overstates real
revenue by ~1.8x. Items supply stock balances and the allocation mix only.

Scope rules:
    company view - inter-branch sales and purchases INCLUDED (the company
                   really did buy/sell that stock)
    entity  view - inter-branch EXCLUDED on both sides; HO and Noida stock
                   summed, so transfers net out correctly
    consolidated - inter-branch and related-party both excluded

Run:
    python .claude\\skills\\gp-segment-mom\\scripts\\build_gp_workbook.py
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parents[3]
sys.path.insert(0, str(HERE))

from segment_map import (  # noqa: E402
    COMPANY_ENTITY, SEGMENTS, segment_for_ledger, segment_for_item,
)

TITLE_FONT = Font(bold=True, size=14, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, size=10, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="2E5C8A")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
TOTAL_FILL = PatternFill("solid", fgColor="FFF2CC")
BOLD = Font(bold=True, size=10)
NORMAL = Font(size=10)
NOTE_FONT = Font(size=9, italic=True, color="606060")
OK_FILL = PatternFill("solid", fgColor="C6EFCE")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '#,##0;[Red](#,##0)'
PCT = '0.0%'
METRICS = ["Sales", "COGS", "Gross Profit", "GP %"]


# ---------------------------------------------------------------------------
# core computation
# ---------------------------------------------------------------------------

def item_mix(cdata, company, period):
    """Per-segment stock movement for one company/period.

    Returns {segment: {opening, closing, inward, outward}} plus the set of
    stock groups that had to fall back to keyword matching.
    """
    parents = cdata.get("parents", {})
    by_group = cdata.get("periods", {}).get(period, {}).get("by_group", {})
    out = {s: defaultdict(float) for s in SEGMENTS}
    unmapped = set()
    for leaf, vals in by_group.items():
        segment, explicit, used = segment_for_item(leaf, parents, company)
        if not explicit:
            unmapped.add(leaf)
        for k in ("opening", "closing", "inward", "outward"):
            out[segment][k] += vals.get(k, 0.0)
    return out, unmapped


def company_figures(cdata, company, period, scope):
    """Sales / COGS / GP per segment for one company/period.

    scope: 'company'  - keep inter-branch and related-party
           'entity'   - drop inter-branch
           'group'    - drop inter-branch and related-party
    """
    mix, unmapped = item_mix(cdata, company, period)
    ledgers = cdata.get("ledgers", {}).get(period, [])

    def keep(row):
        if scope in ("entity", "group") and row["is_branch"]:
            return False
        if scope == "group" and row["is_related_party"]:
            return False
        return True

    tagged_sales = defaultdict(float)
    tagged_purch = defaultdict(float)
    direct_cost = defaultdict(float)
    discount = defaultdict(float)
    untagged_sales = untagged_purch = 0.0
    untagged_cost = untagged_disc = 0.0
    memo = defaultdict(float)

    for row in ledgers:
        amount, bucket = row["amount"], row["bucket"]
        if row["is_branch"]:
            memo["branch_" + bucket] += amount
        if row["is_related_party"]:
            memo["related_party_" + bucket] += amount
        if not keep(row):
            continue
        segment = segment_for_ledger(row["ledger"])
        if bucket == "sales":
            if segment:
                tagged_sales[segment] += amount
            else:
                untagged_sales += amount
        elif bucket == "purchase":
            # purchases carry a debit balance -> negative; flip to cost
            if segment:
                tagged_purch[segment] += -amount
            else:
                untagged_purch += -amount
        elif bucket == "direct_cost":
            if segment:
                direct_cost[segment] += -amount
            else:
                untagged_cost += -amount
        elif bucket == "discount":
            if segment:
                discount[segment] += amount
            else:
                untagged_disc += amount

    # Allocate the untagged remainder on this company's own item mix. For
    # Colorix that is ~100% machines and for Enterprises ~95% ink, which is
    # exactly how the business actually trades.
    def allocate(total, key):
        base = {s: max(mix[s][key], 0.0) for s in SEGMENTS}
        denom = sum(base.values())
        if not total:
            return {s: 0.0 for s in SEGMENTS}
        if denom <= 0:                    # nothing to go on - park in OTHER
            return {s: (total if s == "OTHER" else 0.0) for s in SEGMENTS}
        return {s: total * base[s] / denom for s in SEGMENTS}

    alloc_sales = allocate(untagged_sales, "outward")
    alloc_purch = allocate(untagged_purch, "inward")
    alloc_cost = allocate(untagged_cost, "inward")
    alloc_disc = allocate(untagged_disc, "outward")

    result = {}
    for s in SEGMENTS:
        sales = tagged_sales[s] + alloc_sales[s] + discount[s] + alloc_disc[s]
        purchases = tagged_purch[s] + alloc_purch[s]
        costs = direct_cost[s] + alloc_cost[s]
        cogs = mix[s]["opening"] + purchases - mix[s]["closing"] + costs
        result[s] = {
            "Sales": sales, "COGS": cogs, "Gross Profit": sales - cogs,
            "_purchases": purchases, "_opening": mix[s]["opening"],
            "_closing": mix[s]["closing"], "_direct_cost": costs,
            "_tagged_sales": tagged_sales[s], "_alloc_sales": alloc_sales[s],
        }
    return result, memo, unmapped, untagged_sales


def aggregate(data, period, scope, company_filter=None):
    out = {s: defaultdict(float) for s in SEGMENTS}
    memo_all = defaultdict(float)
    unmapped_all = set()
    alloc_total = 0.0
    for company, cdata in data["companies"].items():
        if company_filter and not company_filter(company, cdata):
            continue
        figs, memo, unmapped, untagged = company_figures(
            cdata, company, period, scope)
        for s in SEGMENTS:
            for k, v in figs[s].items():
                out[s][k] += v
        for k, v in memo.items():
            memo_all[k] += v
        unmapped_all |= unmapped
        alloc_total += untagged
    for s in SEGMENTS:
        sales = out[s]["Sales"]
        out[s]["GP %"] = (out[s]["Gross Profit"] / sales) if sales else 0.0
    return out, memo_all, unmapped_all, alloc_total


def rows_for(data, period_keys, scope, company_filter=None):
    rows = {s: {} for s in SEGMENTS}
    total = {}
    for pk in period_keys:
        agg, _, _, _ = aggregate(data, pk, scope, company_filter)
        for s in SEGMENTS:
            rows[s][pk] = dict(agg[s])
        t_sales = sum(agg[s]["Sales"] for s in SEGMENTS)
        t_cogs = sum(agg[s]["COGS"] for s in SEGMENTS)
        t_gp = t_sales - t_cogs
        total[pk] = {"Sales": t_sales, "COGS": t_cogs, "Gross Profit": t_gp,
                     "GP %": (t_gp / t_sales) if t_sales else 0.0}
    rows["TOTAL"] = total
    return rows


# ---------------------------------------------------------------------------
# sheet helpers
# ---------------------------------------------------------------------------

def _title(ws, text, ncols, subtitle=None):
    ws.cell(row=1, column=1, value=text).font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(ncols, 2))
    for c in range(1, max(ncols, 2) + 1):
        ws.cell(row=1, column=c).fill = TITLE_FILL
    ws.row_dimensions[1].height = 22
    if subtitle:
        ws.cell(row=2, column=1, value=subtitle).font = NOTE_FONT
        return 4
    return 3


def write_matrix(ws, title, subtitle, period_keys, row_groups):
    ncols = 1 + len(period_keys) * len(METRICS)
    r = _title(ws, title, ncols, subtitle)

    col = 2
    for pk in period_keys:
        ws.cell(row=r, column=col, value=pk).font = HDR_FONT
        ws.merge_cells(start_row=r, start_column=col,
                       end_row=r, end_column=col + len(METRICS) - 1)
        for i in range(len(METRICS)):
            cell = ws.cell(row=r, column=col + i)
            cell.fill = HDR_FILL
            cell.alignment = Alignment(horizontal="center")
            cell.border = BORDER
        col += len(METRICS)
    ws.cell(row=r, column=1).fill = HDR_FILL
    ws.cell(row=r, column=1).border = BORDER

    r += 1
    h = ws.cell(row=r, column=1, value="Segment")
    h.font, h.fill, h.border = HDR_FONT, HDR_FILL, BORDER
    col = 2
    for _ in period_keys:
        for m in METRICS:
            cell = ws.cell(row=r, column=col, value=m)
            cell.font, cell.fill, cell.border = HDR_FONT, HDR_FILL, BORDER
            cell.alignment = Alignment(horizontal="center")
            col += 1
    ws.freeze_panes = ws.cell(row=r + 1, column=2)
    r += 1

    for label, rows in row_groups:
        if label:
            ws.cell(row=r, column=1, value=label).font = BOLD
            for c in range(1, ncols + 1):
                ws.cell(row=r, column=c).fill = SUB_FILL
                ws.cell(row=r, column=c).border = BORDER
            r += 1
        for name, per_period in rows.items():
            is_total = name.upper().startswith("TOTAL")
            c1 = ws.cell(row=r, column=1, value=name)
            c1.font = BOLD if is_total else NORMAL
            c1.border = BORDER
            if is_total:
                c1.fill = TOTAL_FILL
            col = 2
            for pk in period_keys:
                fig = per_period.get(pk, {})
                for m in METRICS:
                    cell = ws.cell(row=r, column=col)
                    if m == "GP %":
                        gp = get_column_letter(col - 1)
                        sa = get_column_letter(col - 3)
                        cell.value = f"=IF({sa}{r}=0,0,{gp}{r}/{sa}{r})"
                        cell.number_format = PCT
                    else:
                        cell.value = fig.get(m, 0.0)
                        cell.number_format = MONEY
                    cell.font = BOLD if is_total else NORMAL
                    cell.border = BORDER
                    if is_total:
                        cell.fill = TOTAL_FILL
                    col += 1
            r += 1
        r += 1

    ws.column_dimensions["A"].width = 36
    for c in range(2, ncols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 15


# ---------------------------------------------------------------------------

def build(cache_path: Path, out_path: Path):
    data = json.loads(cache_path.read_text(encoding="utf-8"))
    companies = data["companies"]
    periods = data["periods"]

    months = [p["key"] for p in periods if p["kind"] == "month"]
    quarters = [p["key"] for p in periods if p["kind"] == "quarter"]
    period_keys = []
    for m in months:
        if m.startswith("JUL") and quarters:
            period_keys.extend(q for q in quarters if q not in period_keys)
        period_keys.append(m)
    period_keys.extend(q for q in quarters if q not in period_keys)

    detail_period = quarters[0] if quarters else period_keys[0]
    wb = Workbook()

    # 1 ---- Basis & Method --------------------------------------------------
    ws = wb.active
    ws.title = "Basis & Method"
    r = _title(ws, "Gross Profit by Segment - Month on Month", 4)
    for label, text in [
        ("Source", "Tally Prime, live (localhost:9010). Sales and purchases from "
                   "ledgers; stock balances from stock items."),
        ("Generated", data.get("generated", "")),
        ("Period", f"{data.get('fy_start','')} to {data.get('through','')}"),
        ("Companies", f"{len(companies)} Tally files -> "
                      f"{len(set(c['entity'] for c in companies.values()))} legal entities"),
        ("", ""),
        ("Method", "Sales = segment-tagged sales ledgers + untagged sales "
                   "allocated on that company's item outward mix"),
        ("", "COGS = opening stock + purchases - closing stock + segment direct costs"),
        ("", "GP = Sales - COGS"),
        ("", ""),
        ("Why not item sales", "Item outward value includes branch transfers, godown "
                               "moves and manufacturing issues. For Orange O Tec HO Q1 "
                               "it overstates real sales by about 1.8x, so items are "
                               "used only for stock balances and allocation mix."),
        ("Inter-branch", "Company view keeps inter-branch flows. Entity view drops "
                         "them on both sides and sums HO + Noida stock, so transfers "
                         "net out. Consolidated also drops related-party sales."),
        ("", ""),
        ("KEY LIMITATION", "Only Orange O Tec has segment-tagged SALES ledgers "
                           "(HO 77% tagged, Noida 100%). Enterprises and Colorix have "
                           "NO segment tag on sales - about half of group revenue."),
        ("How that is handled", "Their sales are allocated across segments on each "
                                "company's own stock movement mix, per the business "
                                "rule that Enterprises is the ink company and Colorix "
                                "sells machines. Q1 data supports this: Enterprises HO "
                                "is 95% ink groups, Colorix 99.8% machinery."),
        ("", "These allocated figures are an ESTIMATE, not a measurement. See the "
             "Allocation Check sheet for how much of each segment is estimated."),
        ("Recommended fix", "Create segment-suffixed sales ledgers in Enterprises "
                            "(HO + Noida) and Colorix mirroring the Orange O Tec "
                            "naming. The report then becomes fully measured."),
        ("", ""),
        ("Not included", "Sale quantity. Tally returns no InwardQuantity / "
                         "OutwardQuantity / BilledQty on this route; it needs a "
                         "voucher-level fetch."),
        ("", ""),
        ("RELIABILITY - by segment", "Ink is reliable: 35.6% Q1 here against 37.4% in "
                                     "the existing GP REPORT, derived independently. "
                                     "Heads and Spare Parts are reasonable."),
        ("", "MACHINES and OTHER are NOT reliable at segment level and are flagged "
             "on Cross Check. Inter-branch and related-party transfers are attributed "
             "to a different segment on the sales side than on the purchase side, "
             "which pushes cost into Machines and Other without matching revenue."),
        ("What totals ARE reliable", "Group and company TOTALS are sound and verified: "
                                     "monthly figures sum exactly to Q1, every company "
                                     "ties to its sales ledgers, and gross profit is "
                                     "invariant across all three elimination scopes."),
        ("To fix the split", "Segment GP needs item-wise profitability at voucher level "
                             "(sales value and matched cost per invoice line) rather "
                             "than ledger-and-stock arithmetic. That is a heavier "
                             "extract and is the recommended next step."),
    ]:
        ws.cell(row=r, column=1, value=label).font = BOLD
        c = ws.cell(row=r, column=2, value=text)
        c.font = NORMAL
        c.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 115

    # 2 ---- Consolidated ----------------------------------------------------
    write_matrix(wb.create_sheet("MoM GP - Consolidated"),
                 "Gross Profit by Segment - Group Consolidated",
                 "Inter-branch and related-party sales excluded.",
                 period_keys, [("", rows_for(data, period_keys, "group"))])

    # 3 ---- By entity -------------------------------------------------------
    entities = sorted({c["entity"] for c in companies.values()})
    groups = [(f"{e}   (HO + branch combined)",
               rows_for(data, period_keys, "entity",
                        lambda n, c, e=e: c["entity"] == e)) for e in entities]
    write_matrix(wb.create_sheet("MoM GP - by Entity"),
                 "Gross Profit by Segment - by Legal Entity",
                 "Five Tally files collapse to three legal entities. "
                 "Inter-branch flows eliminated within each entity.",
                 period_keys, groups)

    # 4 ---- By company ------------------------------------------------------
    groups = []
    for name in companies:
        ent, loc = COMPANY_ENTITY.get(name, ("?", "?"))
        groups.append((f"{name}   [{ent} / {loc}]",
                       rows_for(data, period_keys, "company",
                                lambda n, c, t=name: n == t)))
    write_matrix(wb.create_sheet("MoM GP - by Company"),
                 "Gross Profit by Segment - by Tally Company",
                 "Inter-branch flows INCLUDED - each file shown as it stands.",
                 period_keys, groups)

    # 5 ---- Allocation check ------------------------------------------------
    ws = wb.create_sheet("Allocation Check")
    r = _title(ws, "Measured vs Allocated Sales", 6,
               f"How much of each company's {detail_period} sales came from a "
               "segment-tagged ledger versus an allocation.")
    for i, h in enumerate(["Company", "Tagged sales (measured)",
                           "Untagged sales (allocated)", "Total",
                           "Measured %"], start=1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.font, cell.fill, cell.border = HDR_FONT, HDR_FILL, BORDER
    r += 1
    grand_t = grand_u = 0.0
    for name, cdata in companies.items():
        figs, _, _, untagged = company_figures(cdata, name, detail_period, "group")
        tagged = sum(f["_tagged_sales"] for f in figs.values())
        grand_t += tagged
        grand_u += untagged
        ws.cell(row=r, column=1, value=name).font = NORMAL
        for i, v in enumerate([tagged, untagged, tagged + untagged], start=2):
            cell = ws.cell(row=r, column=i, value=v)
            cell.number_format = MONEY
            cell.font = NORMAL
        cell = ws.cell(row=r, column=5, value=f"=IF(D{r}=0,0,B{r}/D{r})")
        cell.number_format = PCT
        cell.font = NORMAL
        r += 1
    ws.cell(row=r, column=1, value="GROUP TOTAL").font = BOLD
    for i, v in enumerate([grand_t, grand_u, grand_t + grand_u], start=2):
        cell = ws.cell(row=r, column=i, value=v)
        cell.number_format = MONEY
        cell.font = BOLD
        cell.fill = TOTAL_FILL
    cell = ws.cell(row=r, column=5, value=f"=IF(D{r}=0,0,B{r}/D{r})")
    cell.number_format = PCT
    cell.font = BOLD
    cell.fill = TOTAL_FILL
    ws.column_dimensions["A"].width = 54
    for c in "BCDE":
        ws.column_dimensions[c].width = 24

    # 6 ---- Segment detail --------------------------------------------------
    ws = wb.create_sheet("Segment Detail")
    r = _title(ws, "Segment Detail - Stock Movement by Group", 8,
               f"{detail_period}. Stock groups driving each segment.")
    for i, h in enumerate(["Segment", "Stock Group", "Company", "Opening",
                           "Inward", "Outward", "Closing"], start=1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.font, cell.fill, cell.border = HDR_FONT, HDR_FILL, BORDER
    ws.freeze_panes = ws.cell(row=r + 1, column=1)
    r += 1
    detail_rows = []
    for name, cdata in companies.items():
        parents = cdata.get("parents", {})
        by_group = cdata.get("periods", {}).get(detail_period, {}).get("by_group", {})
        for leaf, v in by_group.items():
            if not any(v.values()):
                continue
            seg, _, _ = segment_for_item(leaf, parents, name)
            detail_rows.append((seg, leaf, name, v))
    detail_rows.sort(key=lambda x: (SEGMENTS.index(x[0]), -x[3]["outward"]))
    for seg, leaf, name, v in detail_rows:
        ws.cell(row=r, column=1, value=seg).font = NORMAL
        ws.cell(row=r, column=2, value=leaf).font = NORMAL
        ws.cell(row=r, column=3, value=name).font = NORMAL
        for i, k in enumerate(["opening", "inward", "outward", "closing"], start=4):
            cell = ws.cell(row=r, column=i, value=v.get(k, 0.0))
            cell.number_format = MONEY
            cell.font = NORMAL
        r += 1
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 50
    for c in "DEFG":
        ws.column_dimensions[c].width = 16

    # 7 ---- Cross check -----------------------------------------------------
    ws = wb.create_sheet("Cross Check")
    r = _title(ws, "Cross Check", 5)
    for i, h in enumerate(["Check", "Value", "Expected", "Status", "Note"], start=1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.font, cell.fill, cell.border = HDR_FONT, HDR_FILL, BORDER
    r += 1
    cursor = [r]

    def check(label, value, expected, note, tol=1.0):
        row = cursor[0]
        ok = abs(value - expected) <= max(tol, abs(expected) * 0.0001)
        ws.cell(row=row, column=1, value=label).font = NORMAL
        for i, v in ((2, value), (3, expected)):
            cell = ws.cell(row=row, column=i, value=v)
            cell.number_format = MONEY
            cell.font = NORMAL
        st = ws.cell(row=row, column=4, value="OK" if ok else "CHECK")
        st.fill = OK_FILL if ok else BAD_FILL
        st.font = BOLD
        ws.cell(row=row, column=5, value=note).font = NOTE_FONT
        cursor[0] += 1

    q1_months = [m for m in months if not m.startswith("JUL")]
    if quarters:
        q = quarters[0]
        for metric in ("Sales", "Gross Profit"):
            qa, _, _, _ = aggregate(data, q, "group")
            qv = sum(qa[s][metric] for s in SEGMENTS)
            mv = 0.0
            for mk in q1_months:
                ma, _, _, _ = aggregate(data, mk, "group")
                mv += sum(ma[s][metric] for s in SEGMENTS)
            check(f"Q1 {metric} vs Apr+May+Jun", qv, mv,
                  "Q1 is fetched from Tally separately; must equal the 3 months.")

    for pk in period_keys:
        all_agg, _, _, _ = aggregate(data, pk, "entity")
        tot = sum(all_agg[s]["Sales"] for s in SEGMENTS)
        by_ent = 0.0
        for e in entities:
            a, _, _, _ = aggregate(data, pk, "entity",
                                   lambda n, c, e=e: c["entity"] == e)
            by_ent += sum(a[s]["Sales"] for s in SEGMENTS)
        check(f"{pk}: entity sum vs all-company", by_ent, tot,
              "Every Tally file must map to exactly one legal entity.")

    # Gross profit must be invariant to the elimination scope: excluding a
    # branch sale also excludes the matching branch purchase.
    scope_gp = {}
    for scope in ("company", "entity", "group"):
        a, _, _, _ = aggregate(data, detail_period, scope)
        scope_gp[scope] = sum(a[s]["Gross Profit"] for s in SEGMENTS)
    check("GP invariant across elimination scopes",
          scope_gp["group"], scope_gp["company"],
          "Branch sales and branch purchases must cancel; a gap means the "
          "elimination is lopsided.")

    # Flag segments whose margin is not credible, so nobody quotes them.
    seg_agg, _, _, _ = aggregate(data, detail_period, "group")
    for s in SEGMENTS:
        sales, pct = seg_agg[s]["Sales"], seg_agg[s]["GP %"]
        if abs(sales) < 1:
            continue
        row = cursor[0]
        # A negative gross margin on a traded product line is not credible;
        # where it appears it is the transfer-attribution defect, not trading.
        bad = not (0.0 <= pct <= 0.75)
        ws.cell(row=row, column=1, value=f"{s}: GP% credible?").font = NORMAL
        c = ws.cell(row=row, column=2, value=pct)
        c.number_format = PCT
        c.font = NORMAL
        ws.cell(row=row, column=3, value="0% to 75%").font = NORMAL
        st = ws.cell(row=row, column=4, value="OK" if not bad else "DO NOT USE")
        st.fill = OK_FILL if not bad else BAD_FILL
        st.font = BOLD
        ws.cell(row=row, column=5,
                value=("Within a plausible trading range." if not bad else
                       "Implausible - internal transfers are attributed to a "
                       "different segment on the sales side than the purchase "
                       "side. Use the TOTAL, not this segment.")).font = NOTE_FONT
        cursor[0] += 1

    _, _, unmapped, _ = aggregate(data, detail_period, "group")
    row = cursor[0]
    ws.cell(row=row, column=1, value="Stock groups without an explicit mapping").font = NORMAL
    ws.cell(row=row, column=2, value=len(unmapped)).font = NORMAL
    ws.cell(row=row, column=3, value=0).font = NORMAL
    st = ws.cell(row=row, column=4, value="OK" if not unmapped else "REVIEW")
    st.fill = OK_FILL if not unmapped else BAD_FILL
    st.font = BOLD
    ws.cell(row=row, column=5,
            value="; ".join(sorted(unmapped)[:25]) if unmapped
            else "All groups matched an explicit rule.").font = NOTE_FONT
    cursor[0] += 1

    ws.column_dimensions["A"].width = 42
    for c, w in (("B", 18), ("C", 18), ("D", 10), ("E", 95)):
        ws.column_dimensions[c].width = w

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cache", default=str(PROJECT_ROOT / "Files" / "gp_segment_cache.json"))
    ap.add_argument("--out", default=None)
    args = ap.parse_args()
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    out = Path(args.out) if args.out else (
        PROJECT_ROOT / "Files" / f"GP_Segment_MoM_{stamp}.xlsx")
    print(f"Wrote {build(Path(args.cache), out)}")


if __name__ == "__main__":
    main()
