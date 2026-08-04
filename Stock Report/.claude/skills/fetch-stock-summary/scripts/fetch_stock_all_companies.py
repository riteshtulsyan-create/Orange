"""
Fetch the Stock Summary for EVERY open Tally company into ONE workbook.

All companies land in a single 'Item List (Flat)' sheet with Company as column A,
so the pivots can slice across companies or drill into any one of them.

Output: Inventory/Stock_Summary_ALL_COMPANIES_<YYYY-MM-DD_HHMM>.xlsx
  Item List (Flat) - every item, every company, one row each
  Cross Check      - per-company reconciliation + combined total + name variations
  Index / Summary / Group / Sub Group - native PivotTables (added by
                     create_all_pivots.py, which this script calls)

Usage:
    python fetch_stock_all_companies.py
    python fetch_stock_all_companies.py --from 2026-04-01 --to 2026-07-31
    python fetch_stock_all_companies.py --company "A" --company "B"
"""

import argparse
import re
import subprocess
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).parent))
import fetch_stock_summary as F           # noqa: E402  (path set above)

# Five levels: one company (ORANGE O TEC PVT LTD) nests 5 groups deep.
LEVEL_HEADERS = ["Main Category", "Sub Category", "Product Group",
                 "Sub Group", "Sub Group 2"]
MAX_LEVELS = len(LEVEL_HEADERS)

# Short, tab-friendly company labels. Anything not listed falls back to a
# trimmed version of the Tally name.
SHORT_NAMES = {
    "COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)": "Colorix",
    "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27": "OOT Ent Noida",
    "ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)": "OOT Ent",
    "ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)": "OOT Pvt Ltd",
    "ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)": "OOT Noida",
}

# The 5 Tally files are only 3 legal companies: OOT and Enterprises each keep a
# separate Tally file for their Noida branch. Confirmed with the user 2026-07-31.
# (tally company name) -> (Main Company, Location)
# Short codes set by the user 2026-07-31: OTEC / ENTER / COLORIX.
COMPANY_STRUCTURE = {
    "ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)":
        ("OTEC", "Surat"),
    "ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)":
        ("OTEC", "Noida"),
    "ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)":
        ("ENTER", "Surat"),
    "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27":
        ("ENTER", "Noida"),
    "COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)":
        ("COLORIX", "-"),
}

# Groups that could double-count on consolidation if they ever carry value:
# goods in transit between HO and branch, or stock physically held elsewhere.
IN_TRANSIT_MARKERS = ("GIT", "STOCK LYING AT THIRD PARTY", "AT JOB WORKERS")

# Main Group consolidation. Original Tally names are ALWAYS kept in
# 'Main Category'; this only fills the extra 'Main Group (Mapped)' column.
# Confirmed with the user 2026-07-31.
MAIN_GROUP_MAP = {
    "MACHINERY PARTS": "MACHINERY PARTS",
    "MACHINERY & MACHINERY PARTS": "MACHINERY PARTS",
    "MACHINERY": "MACHINES",
    "DIGITAL PRINTING MACHINES STOCK": "MACHINES",
}

# Shown on Cross Check so the mapping and its caveat travel with the workbook.
NAME_VARIATION_NOTES = [
    ("MACHINERY PARTS", "MACHINERY PARTS  +  MACHINERY & MACHINERY PARTS",
     "Same concept, named differently across companies."),
    ("MACHINES", "MACHINERY  +  DIGITAL PRINTING MACHINES STOCK",
     "Finished machines, as distinct from spare parts."),
    ("MACHINERY HEAD", "left separate - NOT merged into MACHINERY PARTS",
     "Print heads are a distinct class from spares."),
    ("STRUCTURAL WARNING", "OOT Noida nests MACHINERY HEAD inside "
                           "MACHINERY & MACHINERY PARTS",
     "The other companies hold MACHINERY HEAD at top level. Mapped totals for "
     "MACHINERY PARTS therefore include heads for OOT Noida but not for the "
     "others - not like-for-like. Fix at source in Tally for a clean compare."),
]

NUM = '#,##0.00;[Red]-#,##0.00'
QTY = '#,##0.000;[Red]-#,##0.000'
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
STAMP_FONT = Font(size=9, italic=True, color="595959")
SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
TOT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def get_open_companies(host, port) -> list:
    """Names of every company currently open in Tally, in Tally's own order."""
    xml = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>AllCompanies</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION ISMODIFY="No" NAME="AllCompanies">
            <TYPE>Company</TYPE>
            <FETCH>Name</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""
    root = F._safe_parse(F.post_to_tally(xml, host, port))
    out = []
    for elem in root.iter("COMPANY"):
        name = (elem.get("NAME") or F._get_text(elem, "NAME")).strip()
        if name and name not in out:
            out.append(name)
    return out


def short_name(company: str) -> str:
    if company in SHORT_NAMES:
        return SHORT_NAMES[company]
    s = re.sub(r"\s*\(.*?\)\s*", " ", company).strip()
    return (s[:28] or company[:28])


def month_ends(from_iso: str, to_iso: str):
    """Every month-end date from `from` through `to`, inclusive of `to` itself."""
    start = datetime.strptime(from_iso, "%Y-%m-%d").date()
    end = datetime.strptime(to_iso, "%Y-%m-%d").date()
    out, y, m = [], start.year, start.month
    while True:
        ny, nm = (y + 1, 1) if m == 12 else (y, m + 1)
        last = date(ny, nm, 1) - timedelta(days=1)
        if last > end:
            break
        if last >= start:
            out.append(last)
        y, m = ny, nm
    if not out or out[-1] != end:
        out.append(end)          # include the as-on date even if mid-month
    return out


def stamp_text(from_date, to_date):
    return (f"Last refreshed {datetime.now().strftime('%d-%b-%Y %H:%M')}  |  "
            f"Live from Tally Prime {F.TALLY_HOST}:{F.TALLY_PORT}  |  "
            f"Stock as at {F._to_ddmmmyyyy(to_date)}  "
            f"(period {F._to_ddmmmyyyy(from_date)} to {F._to_ddmmmyyyy(to_date)})")


def structure_of(company: str):
    """(Main Company, Location) for a Tally company name."""
    if company in COMPANY_STRUCTURE:
        return COMPANY_STRUCTURE[company]
    # Unmapped company: infer rather than drop it out of the consolidation.
    # Flagged on Cross Check so an unrecognised file cannot pass silently.
    up = company.upper()
    loc = "Noida" if "NOIDA" in up else "Surat"
    main = re.sub(r"\s*[-(].*$", "", company).strip() or company
    return main.upper(), loc


def fetch_one(company, from_date, to_date, host, port):
    """Return (rows, stats) for one company. rows carry the full group path."""
    groups = F.fetch_stock_groups(company, from_date, to_date, host, port)
    items = F.fetch_stock_items(company, from_date, to_date, host, port)
    if not items:
        return [], {"items": 0, "groups": len(groups), "total": 0.0,
                    "sign": 1, "ungrouped": 0, "negative": 0, "zero": 0,
                    "max_depth": 0}

    sign = F.normalise_sign(items)
    paths = F.build_paths(groups)

    rows, ungrouped, negative, zero, max_depth = [], 0, 0, 0, 0
    for it in items:
        pk = it["parent"].upper()
        info = paths.get(pk)
        if info:
            path = info["path"]
        else:
            path = ["(No stock group assigned)"] if not it["parent"] else [it["parent"]]
            if not it["parent"]:
                ungrouped += 1
        val = round(it["value"] * sign, 2)
        if val < 0:
            negative += 1
        if val == 0:
            zero += 1
        max_depth = max(max_depth, len(path))
        rows.append({"company": company, "path": path, "name": it["name"],
                     "qty": it["qty"], "unit": it["unit"],
                     "rate": it["rate"], "value": val})

    return rows, {"items": len(items), "groups": len(groups),
                  "total": round(sum(r["value"] for r in rows), 2), "sign": sign,
                  "ungrouped": ungrouped, "negative": negative, "zero": zero,
                  "max_depth": max_depth}


def write_workbook(out, rows, stats, from_date, to_date, mom=None):
    wb = openpyxl.Workbook()

    # ---- Item List (Flat) -------------------------------------------------
    ws = wb.active
    ws.title = "Item List (Flat)"
    headers = (["Main Company", "Location", "Company", "Company (Full)"] + LEVEL_HEADERS +
               ["Main Group (Mapped)", "Full Group Path", "Item Name",
                "Quantity", "Unit", "Rate", "Value", "Depth", "Has Qty", "Zero Value"])
    widths = ([26, 10, 16, 46] + [30, 28, 28, 28, 24] +
              [24, 60, 46, 15, 9, 13, 17, 8, 9, 11])
    # Row 1 carries the refresh stamp on every sheet; headers sit on row 2 and
    # data starts on row 3. The pivot source range in create_all_pivots.py
    # starts at row 2 to match -- change one and you must change the other.
    ws.cell(row=1, column=1, value=stamp_text(from_date, to_date)).font = STAMP_FONT
    for c, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "E3"
    ws.row_dimensions[2].height = 30

    r = 3
    for row in rows:
        path = row["path"]
        main_co, loc = structure_of(row["company"])
        ws.cell(row=r, column=1, value=main_co)
        ws.cell(row=r, column=2, value=loc)
        ws.cell(row=r, column=3, value=short_name(row["company"]))
        ws.cell(row=r, column=4, value=row["company"])
        for i in range(MAX_LEVELS):
            if i < len(path):
                last = (i == MAX_LEVELS - 1)
                v = " > ".join(path[i:]) if (last and len(path) > MAX_LEVELS) else path[i]
            else:
                v = ""
            ws.cell(row=r, column=5 + i, value=v)
        main = path[0] if path else ""
        ws.cell(row=r, column=10, value=MAIN_GROUP_MAP.get(main, main))
        ws.cell(row=r, column=11, value=" > ".join(path))
        ws.cell(row=r, column=12, value=row["name"])
        if row["qty"] is not None:
            ws.cell(row=r, column=13, value=round(row["qty"], 3)).number_format = QTY
        ws.cell(row=r, column=14, value=row["unit"])
        if row["rate"] is not None:
            ws.cell(row=r, column=15, value=round(row["rate"], 2)).number_format = NUM
        ws.cell(row=r, column=16, value=row["value"]).number_format = NUM
        ws.cell(row=r, column=17, value=len(path))
        ws.cell(row=r, column=18, value="Yes" if row["qty"] not in (None, 0) else "No")
        ws.cell(row=r, column=19, value="Yes" if row["value"] == 0 else "No")
        r += 1
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{max(r - 1, 2)}"

    # ---- Cross Check ------------------------------------------------------
    ws2 = wb.create_sheet("Cross Check")
    ws2.cell(row=1, column=1, value=stamp_text(from_date, to_date)).font = STAMP_FONT
    for c, (h, w) in enumerate(zip(["Check", "Value", "Status", "Note"],
                                   [46, 30, 14, 78]), start=1):
        cell = ws2.cell(row=2, column=c, value=h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws2.column_dimensions[get_column_letter(c)].width = w
    ws2.freeze_panes = "A3"

    grand = round(sum(x["total"] for x in stats.values()), 2)
    sheet_sum = round(sum(r_["value"] for r_ in rows), 2)

    checks = [
        ("Period", f"{F._to_ddmmmyyyy(from_date)} to {F._to_ddmmmyyyy(to_date)}", "INFO",
         "Closing balances as at the end of this period."),
        ("Generated", datetime.now().strftime("%d-%b-%Y %H:%M"), "INFO",
         f"Live from Tally Prime at {F.TALLY_HOST}:{F.TALLY_PORT} - not a cached copy."),
        ("Companies included", len(stats), "INFO",
         "; ".join(short_name(c) for c in stats)),
        ("COMBINED closing stock value", grand, "INFO",
         "Sum across every company below."),
        ("Item List sheet total", sheet_sum,
         "OK" if abs(sheet_sum - grand) < 0.01 else "MISMATCH",
         "Must equal the combined total, else a row was lost writing the sheet."),
        ("Total items (all companies)", sum(x["items"] for x in stats.values()), "INFO", ""),
        ("", "", "", ""),
    ]
    for comp, st in sorted(stats.items(), key=lambda kv: -kv[1]["total"]):
        main_co, loc = structure_of(comp)
        checks.append((f"  {short_name(comp)}", st["total"], "INFO",
                       f"{main_co} / {loc} | {st['items']:,} items | "
                       f"{st['groups']} groups | depth {st['max_depth']} | "
                       f"{st['zero']:,} nil-value | {st['negative']} negative | "
                       f"{st['ungrouped']} ungrouped"
                       + ("  [sign flipped]" if st["sign"] == -1 else "")))
    checks.append(("", "", "", ""))

    # ---- consolidation by legal entity ------------------------------------
    by_main, by_loc, unmapped = {}, {}, []
    for comp, st in stats.items():
        main_co, loc = structure_of(comp)
        by_main[main_co] = by_main.get(main_co, 0.0) + st["total"]
        by_loc[loc] = by_loc.get(loc, 0.0) + st["total"]
        if comp not in COMPANY_STRUCTURE:
            unmapped.append(comp)

    checks.append(("BY MAIN COMPANY (legal entity)", len(by_main), "INFO",
                   f"{len(stats)} Tally files consolidate into {len(by_main)} companies. "
                   f"OOT and Enterprises each keep a separate Tally file for Noida."))
    for k, v in sorted(by_main.items(), key=lambda kv: -kv[1]):
        checks.append((f"  {k}", round(v, 2), "INFO", ""))
    consolidated = round(sum(by_main.values()), 2)
    checks.append(("  Consolidated total", consolidated,
                   "OK" if abs(consolidated - grand) < 0.01 else "MISMATCH",
                   "Must equal the combined total above."))
    checks.append(("", "", "", ""))

    checks.append(("BY LOCATION", len(by_loc), "INFO",
                   "Stock held at each site, across all companies."))
    for k, v in sorted(by_loc.items(), key=lambda kv: -kv[1]):
        checks.append((f"  {k}", round(v, 2), "INFO", ""))
    if unmapped:
        checks.append(("  Unmapped companies", len(unmapped), "REVIEW",
                       "Not in the known structure; Main Company/Location were inferred "
                       "from the name: " + "; ".join(unmapped)))
    checks.append(("", "", "", ""))

    # ---- in-transit watch -------------------------------------------------
    transit = {}
    for row_ in rows:
        p0 = (row_["path"][0] if row_["path"] else "").upper()
        if any(m in p0 for m in IN_TRANSIT_MARKERS):
            transit[row_["path"][0]] = transit.get(row_["path"][0], 0.0) + row_["value"]
    transit_val = round(sum(transit.values()), 2)
    checks.append(("Goods-in-transit / third-party stock", transit_val,
                   "OK" if transit_val == 0 else "REVIEW",
                   ("Nil - nothing to double-count, so HO + Noida simply add up."
                    if transit_val == 0 else
                    "CARRIES VALUE: " + "; ".join(f"{k} {v:,.2f}" for k, v in transit.items())
                    + ". Consolidated totals may double-count stock in transit between "
                      "HO and branch. Verify before quoting the consolidated figure.")))
    checks.append(("", "", "", ""))

    neg_total = sum(x["negative"] for x in stats.values())
    ung_total = sum(x["ungrouped"] for x in stats.values())
    checks += [
        ("Items with NEGATIVE closing value", neg_total,
         "OK" if not neg_total else "REVIEW",
         "Stock issued without a matching receipt. See the Company column to locate."),
        ("Items with no stock group", ung_total,
         "OK" if not ung_total else "REVIEW", "Master-data gap in Tally."),
        ("Deepest group nesting found", max(x["max_depth"] for x in stats.values()), "INFO",
         f"The flat sheet provides {MAX_LEVELS} level columns; anything deeper is "
         f"folded into '{LEVEL_HEADERS[-1]}'."),
        ("", "", "", ""),
        ("MAIN GROUP NAME VARIATIONS", "", "REVIEW",
         "'Main Category' holds the original Tally name. "
         "'Main Group (Mapped)' applies the consolidation below."),
    ]
    for mapped, sources, why in NAME_VARIATION_NOTES:
        checks.append((f"  {mapped}", sources, "", why))

    for i, (label, val, status, note) in enumerate(checks, start=3):
        ws2.cell(row=i, column=1, value=label)
        c = ws2.cell(row=i, column=2, value=val)
        if isinstance(val, float):
            c.number_format = NUM
        st = ws2.cell(row=i, column=3, value=status)
        st.font = Font(bold=True, color={"OK": "006100", "MISMATCH": "9C0006",
                                         "REVIEW": "9C5700"}.get(status, "404040"))
        st.alignment = Alignment(horizontal="center")
        ws2.cell(row=i, column=4, value=note).alignment = Alignment(wrap_text=True,
                                                                   vertical="top")
        if str(label).isupper() and label:
            ws2.cell(row=i, column=1).font = Font(bold=True)

    # ---- Month on Month ---------------------------------------------------
    if mom:
        write_mom_data_sheet(wb, mom, from_date, to_date)
        write_mom_sheet(wb, mom, from_date, to_date)

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return grand, sheet_sum


def fetch_mom(companies, from_date, to_date, host, port):
    """
    Closing stock value at each month end in the period.

    One extra fetch per company per month end -- Tally has no 'give me every
    month end at once' call, so this is unavoidable. Aggregated to company and
    Main Category level only; item-level month-on-month would be 14,000 rows
    times N months and is not what the analysis needs.
    """
    ends = month_ends(from_date, to_date)
    labels = [d.strftime("%b-%y") for d in ends]
    print(f"\n--- Month on Month: {len(ends)} month end(s) x {len(companies)} companies ---")
    print(f"    {', '.join(labels)}")

    as_on = datetime.strptime(to_date, "%Y-%m-%d").date()
    by_company, by_category, flat = {}, {}, {}
    for comp in companies:
        groups = None
        main_co, loc = structure_of(comp)
        for d in ends:
            iso = d.isoformat()
            try:
                if groups is None:
                    groups = F.fetch_stock_groups(comp, from_date, iso, host, port)
                items = F.fetch_stock_items(comp, from_date, iso, host, port)
            except Exception as e:
                print(f"    {short_name(comp)} @ {d:%b-%y}: FAILED - {str(e)[:70]}")
                continue
            if not items:
                continue
            sign = F.normalise_sign(items)
            paths = F.build_paths(groups)
            tot = 0.0
            for it in items:
                v = round(it["value"] * sign, 2)
                tot += v
                info = paths.get(it["parent"].upper())
                path = (info["path"] if info else
                        (["(No stock group assigned)"] if not it["parent"] else [it["parent"]]))
                main = MAIN_GROUP_MAP.get(path[0], path[0])
                sub = path[1] if len(path) > 1 else ""
                by_category.setdefault(main, {})[d] = \
                    round(by_category.setdefault(main, {}).get(d, 0.0) + v, 2)
                # Long-form key for the pivot source. Aggregated to
                # category/sub-category so the sheet stays a few thousand rows
                # rather than 14,000 items x every period.
                k = (d, comp, path[0], main, sub)
                flat[k] = round(flat.get(k, 0.0) + v, 2)
            by_company.setdefault(comp, {})[d] = round(tot, 2)
            print(f"    {short_name(comp):<14} {d:%b-%y}  {tot:>18,.2f}")

    rows = []
    for (d, comp, cat, mapped, sub), v in flat.items():
        main_co, loc = structure_of(comp)
        if d == as_on:
            ptype = "Month End + As On Date" if d == ends[-1] and _is_month_end(d) \
                else "As On Date"
        else:
            ptype = "Month End"
        rows.append({
            "period": d.strftime("%b-%y"), "date": d, "ptype": ptype,
            "main_co": main_co, "company": short_name(comp), "location": loc,
            "category": cat, "mapped": mapped, "sub": sub, "value": v,
        })
    rows.sort(key=lambda r: (r["date"], r["main_co"], r["company"],
                             -r["value"]))
    print(f"    MOM Data rows: {len(rows):,}")

    return {"ends": ends, "labels": labels, "as_on": as_on,
            "by_company": by_company, "by_category": by_category, "rows": rows}


def _is_month_end(d):
    nd = date(d.year + 1, 1, 1) if d.month == 12 else date(d.year, d.month + 1, 1)
    return (nd - timedelta(days=1)) == d


def write_mom_data_sheet(wb, mom, from_date, to_date):
    """
    Long-form month-on-month data: one row per period x company x category x
    sub-category. This is the source for the 'MOM Pivot' sheet, so the user can
    reshape the analysis rather than being stuck with a fixed layout.

    Same convention as the other data sheets: stamp on row 1, header on row 2,
    data from row 3 -- create_all_pivots.py points its cache at row 2.
    """
    ws = wb.create_sheet("MOM Data")
    ws.cell(row=1, column=1, value=stamp_text(from_date, to_date)).font = STAMP_FONT

    headers = ["Period", "Period End Date", "Period Type", "Main Company",
               "Company", "Location", "Main Category", "Main Group (Mapped)",
               "Sub Category", "Value"]
    widths = [11, 16, 22, 16, 16, 10, 30, 26, 28, 18]
    for c, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=2, column=c, value=h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[2].height = 28
    ws.freeze_panes = "A3"

    r = 3
    for row in mom["rows"]:
        ws.cell(row=r, column=1, value=row["period"])
        d = ws.cell(row=r, column=2, value=row["date"])
        d.number_format = "dd-mmm-yyyy"
        ws.cell(row=r, column=3, value=row["ptype"])
        ws.cell(row=r, column=4, value=row["main_co"])
        ws.cell(row=r, column=5, value=row["company"])
        ws.cell(row=r, column=6, value=row["location"])
        ws.cell(row=r, column=7, value=row["category"])
        ws.cell(row=r, column=8, value=row["mapped"])
        ws.cell(row=r, column=9, value=row["sub"])
        ws.cell(row=r, column=10, value=row["value"]).number_format = NUM
        r += 1
    ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{max(r - 1, 2)}"
    return ws


def write_mom_sheet(wb, mom, from_date, to_date):
    ends, labels = mom["ends"], mom["labels"]
    ws = wb.create_sheet("Month on Month")

    ws.cell(row=1, column=1, value=stamp_text(from_date, to_date)).font = STAMP_FONT
    ws.cell(row=2, column=1, value="STOCK MONTH ON MONTH - closing value at each month end")
    ws.cell(row=2, column=1).font = Font(bold=True, size=13, color="1F4E79")
    ws.cell(row=3, column=1,
            value="Each column is a separate live fetch from Tally as at that date, "
                  "not an interpolation.").font = STAMP_FONT

    n = len(labels)
    widths = [26, 16, 10] + [17] * n + [17, 11]
    hdr = ["Main Company", "Company", "Location"] + labels + \
          [f"Change {labels[-2]}>{labels[-1]}" if n >= 2 else "Change", "%"]
    r = 5
    for c, (h, w) in enumerate(zip(hdr, widths), start=1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[r].height = 28
    ws.freeze_panes = f"D{r + 1}"
    r += 1

    def money(row, col, val):
        c = ws.cell(row=row, column=col, value=round(val, 2))
        c.number_format = NUM
        c.border = BORDER
        return c

    def change_cols(row, series):
        if n >= 2:
            delta = series.get(ends[-1], 0.0) - series.get(ends[-2], 0.0)
            money(row, 4 + n, delta)
            prev = series.get(ends[-2], 0.0)
            p = ws.cell(row=row, column=5 + n,
                        value=(delta / prev) if prev else 0)
            p.number_format = '0.0%'
            p.border = BORDER

    # by legal entity, then the Tally files inside it
    entity_series, grand_series = {}, {}
    for comp, series in mom["by_company"].items():
        ent, _ = structure_of(comp)
        for d, v in series.items():
            entity_series.setdefault(ent, {})[d] = entity_series.setdefault(ent, {}).get(d, 0.0) + v
            grand_series[d] = grand_series.get(d, 0.0) + v

    for ent in sorted(entity_series, key=lambda e: -entity_series[e].get(ends[-1], 0)):
        ws.cell(row=r, column=1, value=ent).font = Font(bold=True, size=10)
        for i, d in enumerate(ends):
            money(r, 4 + i, entity_series[ent].get(d, 0.0)).font = Font(bold=True, size=10)
        change_cols(r, entity_series[ent])
        for cc in range(1, 6 + n):
            ws.cell(row=r, column=cc).fill = SUB_FILL
            ws.cell(row=r, column=cc).border = BORDER
        r += 1
        for comp, series in sorted(mom["by_company"].items(),
                                   key=lambda kv: -kv[1].get(ends[-1], 0)):
            e2, loc = structure_of(comp)
            if e2 != ent:
                continue
            ws.cell(row=r, column=2, value=short_name(comp)).alignment = Alignment(indent=1)
            ws.cell(row=r, column=3, value=loc)
            for i, d in enumerate(ends):
                money(r, 4 + i, series.get(d, 0.0))
            change_cols(r, series)
            for cc in range(1, 6 + n):
                ws.cell(row=r, column=cc).border = BORDER
            r += 1

    ws.cell(row=r, column=1, value="CONSOLIDATED").font = Font(bold=True, size=11)
    for i, d in enumerate(ends):
        money(r, 4 + i, grand_series.get(d, 0.0)).font = Font(bold=True, size=11)
    change_cols(r, grand_series)
    for cc in range(1, 6 + n):
        ws.cell(row=r, column=cc).fill = TOT_FILL
        ws.cell(row=r, column=cc).border = BORDER
    r += 3

    # by Main Category, consolidated across all companies
    ws.cell(row=r, column=1, value="BY MAIN CATEGORY (all companies)").font = \
        Font(bold=True, size=12, color="1F4E79")
    r += 1
    hdr2 = ["Main Category", "", ""] + labels + \
           [f"Change {labels[-2]}>{labels[-1]}" if n >= 2 else "Change", "%"]
    for c, h in enumerate(hdr2, start=1):
        cell = ws.cell(row=r, column=c, value=h)
        cell.fill, cell.font = HDR_FILL, HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[r].height = 28
    r += 1

    cats = mom["by_category"]
    for cat in sorted(cats, key=lambda k: -cats[k].get(ends[-1], 0)):
        if all(round(cats[cat].get(d, 0.0), 2) == 0 for d in ends):
            continue
        ws.cell(row=r, column=1, value=cat)
        for i, d in enumerate(ends):
            money(r, 4 + i, cats[cat].get(d, 0.0))
        change_cols(r, cats[cat])
        for cc in range(1, 6 + n):
            ws.cell(row=r, column=cc).border = BORDER
        r += 1

    ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True, size=11)
    for i, d in enumerate(ends):
        money(r, 4 + i, sum(c.get(d, 0.0) for c in cats.values())).font = Font(bold=True, size=11)
    change_cols(r, {d: sum(c.get(d, 0.0) for c in cats.values()) for d in ends})
    for cc in range(1, 6 + n):
        ws.cell(row=r, column=cc).fill = TOT_FILL
        ws.cell(row=r, column=cc).border = BORDER
    r += 3

    # ---- Option C cross-tab: category rows, Main Company x month columns ---
    # Layout chosen by the user 2026-08-01. Main Company level only.
    ents = []
    grid = {}
    for row_ in mom.get("rows", []):
        e = row_["main_co"]
        if e not in ents:
            ents.append(e)
        k = (row_["mapped"], e, row_["date"])
        grid[k] = round(grid.get(k, 0.0) + row_["value"], 2)
    ents.sort(key=lambda e: -sum(v for k, v in grid.items()
                                 if k[1] == e and k[2] == ends[-1]))

    if grid:
        ws.cell(row=r, column=1,
                value="BY MAIN CATEGORY x COMPANY (cross-tab)").font = \
            Font(bold=True, size=12, color="1F4E79")
        r += 1
        ws.cell(row=r, column=1,
                value="Each block is one company; columns within it are the "
                      "month ends. The right-hand TOTAL block is all companies "
                      "combined.").font = STAMP_FONT
        r += 2

        # two header rows: company band, then months under each
        band = r
        ws.cell(row=band, column=1, value="Main Category")
        c1 = ws.cell(row=band, column=1)
        c1.fill, c1.font = HDR_FILL, HDR_FONT
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = BORDER
        ws.merge_cells(start_row=band, start_column=1, end_row=band + 1, end_column=1)
        ws.column_dimensions["A"].width = 32

        col = 2
        for label in ents + ["TOTAL (all companies)"]:
            ws.merge_cells(start_row=band, start_column=col,
                           end_row=band, end_column=col + n - 1)
            cell = ws.cell(row=band, column=col, value=label)
            cell.fill, cell.font = HDR_FILL, HDR_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")
            for k in range(n):
                ws.cell(row=band, column=col + k).border = BORDER
            col += n
        for k in range(2, col):
            m = labels[(k - 2) % n]
            cell = ws.cell(row=band + 1, column=k, value=m)
            cell.fill, cell.font = HDR_FILL, HDR_FONT
            cell.alignment = Alignment(horizontal="center")
            cell.border = BORDER
            ws.column_dimensions[get_column_letter(k)].width = 17
        ws.row_dimensions[band].height = 20
        r = band + 2

        cat_names = sorted({k[0] for k in grid},
                           key=lambda c: -sum(grid.get((c, e, ends[-1]), 0.0)
                                              for e in ents))
        for cat in cat_names:
            tot_by_month = [sum(grid.get((cat, e, d), 0.0) for e in ents) for d in ends]
            if all(round(v, 2) == 0 for v in tot_by_month):
                continue
            ws.cell(row=r, column=1, value=cat).border = BORDER
            col = 2
            for e in ents:
                for d in ends:
                    money(r, col, grid.get((cat, e, d), 0.0))
                    col += 1
            for i, d in enumerate(ends):
                money(r, col, tot_by_month[i]).font = Font(bold=True, size=10)
                ws.cell(row=r, column=col).fill = SUB_FILL
                col += 1
            r += 1

        ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True, size=11)
        col = 2
        for e in ents:
            for d in ends:
                money(r, col, sum(grid.get((c, e, d), 0.0)
                                  for c in cat_names)).font = Font(bold=True, size=11)
                col += 1
        for d in ends:
            money(r, col, sum(grid.get((c, e, d), 0.0)
                              for c in cat_names
                              for e in ents)).font = Font(bold=True, size=11)
            col += 1
        for cc in range(1, col):
            ws.cell(row=r, column=cc).fill = TOT_FILL
            ws.cell(row=r, column=cc).border = BORDER

    return ws


def main():
    today = date.today()
    ap = argparse.ArgumentParser(description="Stock summary for all open Tally companies.")
    ap.add_argument("--company", action="append", dest="companies",
                    help="Repeatable. Omit to use every open company.")
    ap.add_argument("--from", dest="from_date", default=F.default_fy_start(today))
    ap.add_argument("--to", dest="to_date", default=today.isoformat())
    ap.add_argument("--output", default=None)
    ap.add_argument("--no-pivots", action="store_true")
    ap.add_argument("--no-mom", action="store_true",
                    help="Skip the Month on Month sheet (saves one Tally fetch "
                         "per company per month end)")
    ap.add_argument("--host", default=F.TALLY_HOST)
    ap.add_argument("--port", type=int, default=F.TALLY_PORT)
    args = ap.parse_args()

    companies = args.companies or get_open_companies(args.host, args.port)
    if not companies:
        print("ERROR: no open companies found in Tally.", file=sys.stderr)
        return 1

    print(f"Tally   : {args.host}:{args.port}")
    print(f"Period  : {F._to_ddmmmyyyy(args.from_date)} to {F._to_ddmmmyyyy(args.to_date)}")
    print(f"Companies ({len(companies)}):")
    for c in companies:
        print(f"   - {c}")

    all_rows, stats = [], {}
    for i, comp in enumerate(companies, 1):
        print(f"\n[{i}/{len(companies)}] {short_name(comp)} ...")
        try:
            rows, st = fetch_one(comp, args.from_date, args.to_date, args.host, args.port)
        except Exception as e:
            print(f"    FAILED: {e}", file=sys.stderr)
            continue
        all_rows.extend(rows)
        stats[comp] = st
        print(f"    {st['items']:,} items | {st['groups']} groups | "
              f"depth {st['max_depth']} | total {st['total']:,.2f}")

    if not all_rows:
        print("ERROR: nothing fetched.", file=sys.stderr)
        return 1

    ts = datetime.now().strftime("%Y-%m-%d_%H%M")
    out = Path(args.output) if args.output else \
        F.DEFAULT_OUTPUT_DIR / f"Stock_Summary_ALL_COMPANIES_{ts}.xlsx"

    mom = None
    if not args.no_mom:
        mom = fetch_mom(list(stats), args.from_date, args.to_date, args.host, args.port)

    grand, sheet_sum = write_workbook(out, all_rows, stats,
                                      args.from_date, args.to_date, mom)

    print(f"\n{'=' * 62}")
    print(f"Companies      : {len(stats)}")
    print(f"Rows written   : {len(all_rows):,}")
    print(f"COMBINED TOTAL : {grand:,.2f}")
    print(f"Sheet total    : {sheet_sum:,.2f}  "
          f"[{'OK' if abs(grand - sheet_sum) < 0.01 else 'MISMATCH'}]")
    print("=" * 62)

    if not args.no_pivots:
        print("\n--- Building PivotTables ---")
        ps = Path(__file__).with_name("create_all_pivots.py")
        if ps.is_file():
            if subprocess.call([sys.executable, str(ps), str(out.resolve())]) != 0:
                print("WARNING: pivots not added; the data workbook is still valid.",
                      file=sys.stderr)

    # Archive older combined workbooks so only the current one sits in Inventory.
    if not args.output:
        F.ARCHIVE_DIR.mkdir(parents=True, exist_ok=True)
        for p in F.DEFAULT_OUTPUT_DIR.glob("Stock_Summary_ALL_COMPANIES_*.xlsx"):
            if p.resolve() != out.resolve():
                try:
                    p.replace(F.ARCHIVE_DIR / p.name)
                except OSError:
                    print(f"NOTE: could not archive {p.name} - open in Excel?")

    print(f"\nSaved: {out.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
