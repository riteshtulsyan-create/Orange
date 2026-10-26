"""
cfo_pack_extras.py - the CFO Pack sheets added in v4 (user, 2026-10-24). Called by build_cfo_pack.py.

  CFO Briefing        the summary, red flags, board questions and caveats - ONE tab, and never in the PDF
  Directors Summary   a one-page executive summary for the directors; every figure (and every number
                      inside the key-message sentences) is a formula on CFO Analysis, so it cannot drift
                      from the pack. Exported as its own PDF.
  Branch Transactions branch sales / purchases (P&L) and branch balances (BS), per company, per period
  Related Party       inter-company sales / purchases (P&L) and related-party balances (BS)

The Branch and Related Party sheets read the CFS sheets' "ELIMINATED ON CONSOLIDATION" memo blocks -
the exact amounts consolidation removes - so they agree with the consolidation by construction.
"""
from __future__ import annotations

from openpyxl.styles import Alignment
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter as col_letter

import build_cfo_pack as B

# Period blocks of the memo sheets, and the balance-sheet month each one reads.
MEMO_PERIODS = B.MOM_PERIODS
MEMO_GROUPS = {
    "Branch Transactions": [("BRANCH SALES AND PURCHASES (P&L) - Rs crore", "pl", {"Branch Sales", "Branch Purchase"}),
                            ("BRANCH BALANCES (Balance Sheet, at period end) - Rs crore", "bs", {"Branch balances"})],
    "Related Party": [("RELATED-PARTY SALES AND PURCHASES (P&L) - Rs crore", "pl",
                       {"Intercompany Sales", "Intercompany Purchases"}),
                      ("RELATED-PARTY BALANCES (Balance Sheet, at period end) - Rs crore", "bs", {"Related party balances"})],
}


# =============================================================================================
# CFO Briefing
# =============================================================================================
def build_cfo_briefing(wb, briefing, source_note):
    ws = wb.create_sheet("CFO Briefing")
    ws.sheet_properties.tabColor = "C00000"
    ws.column_dimensions["A"].width = 160
    ws["A1"], ws["A1"].font = "CFO BRIEFING - summary, red flags, board questions and data caveats", B.F_TITLE
    ws["A2"], ws["A2"].font = source_note + " The figures behind every point are on CFO Analysis and MOM Summary.", B.F_NOTE
    r = 4
    for head, items in briefing:
        B.band(ws, r, head, 1)
        r += 1
        for it in items:
            cell = ws.cell(r, 1, "•  " + it)
            cell.font, cell.alignment = B.F_NORM, B.WRAP
            ws.row_dimensions[r].height = 15 * (1 + len(it) // 170)
            r += 1
        r += 1
    return ws


# =============================================================================================
# Directors Summary
# =============================================================================================
def build_directors_summary(wb, marks, source_note, ageing_marks=None, ageing_as_of=None):
    """Every number links to CFO Analysis (sections 1, 2, 3, 4)."""
    ws = wb.create_sheet("Directors Summary")
    ws.sheet_properties.tabColor = "1F3864"
    ca = "'CFO Analysis'"
    ref = lambda col, key: f"{ca}!${col}${marks[key]}"
    widths = {"A": 44, "B": 15, "C": 15, "D": 15, "E": 15}
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    ws["A1"], ws["A1"].font = "EXECUTIVE SUMMARY FOR THE DIRECTORS", B.F_TITLE
    ws["A2"] = "Orange O Tec Group - FY 2026-27, April to September. Rs crore."
    ws["A2"].font = B.F_BOLD
    ws["A3"], ws["A3"].font = source_note, B.F_NOTE
    r = 5

    def table(title, header, rows):
        nonlocal r
        B.band(ws, r, title, len(header))
        r += 1
        for i, h in enumerate(header, start=1):
            B.hdr_cell(ws, r, i, h)
        r += 1
        for name, cells, fmt, bold in rows:
            B.label(ws, r, name, B.F_BOLD if bold else B.F_NORM)
            for i, f in enumerate(cells, start=2):
                if f is not None:
                    B.num_cell(ws, r, i, f, fmt, B.F_BOLD if bold else B.F_NORM)
            r += 1
        r += 1

    q = lambda key: [f"={ref('B', key)}", f"={ref('C', key)}", f"={ref('E', key)}"]
    table("1. PERFORMANCE (consolidated)", ["Rs crore", "Q1 (Apr-Jun)", "Q2 (Jul-Sep)", "Apr-Sep"], [
        ("Turnover", q("TOTAL TURNOVER"), B.NUM_CR, True),
        ("Gross profit", q("GROSS PROFIT"), B.NUM_CR, False),
        ("Gross margin %", q("Gross margin %"), B.PCT, False),
        ("EBITDA", q("EBITDA"), B.NUM_CR, True),
        ("EBITDA margin %", q("EBITDA margin %"), B.PCT, False),
        ("Finance cost", q("TOTAL FINANCE COST"), B.NUM_CR, False),
        ("Profit before depreciation & tax", q("PAT (NET PROFIT)"), B.NUM_CR, True),
        ("Profit after depreciation & tax (estimate)", [None, None, f"={ref('E', 'adj_pat')}"], B.NUM_CR, True),
        ("Net cash from operations (break-up in 1b)", [f"={ref('B', 'cf_netA')}", f"={ref('C', 'cf_netA')}",
                                                         f"={ref('D', 'cf_netA')}"], B.NUM_CR, False),
    ])
    # Where the operating cash went - the break-up of "Net cash from operations" (working: CFO Analysis section 3).
    cq = lambda *keys: [("=" + "+".join(ref(col, k) for k in keys)) for col in ("B", "C", "D")]
    other = ["cf_oca", "cf_ocl", "cf_dt", "cf_prov"]
    table("1b. NET CASH FROM OPERATIONS - BREAK-UP (consolidated)", ["Rs crore", "Q1 (Apr-Jun)", "Q2 (Jul-Sep)", "Apr-Sep"], [
        ("Operating profit before working capital", cq("cf_opwc"), B.NUM_CR, False),
        ("Debtors (increase) / decrease", cq("cf_deb"), B.NUM_CR, False),
        ("Inventory (increase) / decrease", cq("cf_inv"), B.NUM_CR, False),
        ("Creditors increase / (decrease)", cq("cf_cred"), B.NUM_CR, False),
        ("Loans, advances & deposits", cq("cf_lna", "cf_dep_a"), B.NUM_CR, False),
        ("Other current assets / liabilities, taxes, provisions", cq(*other), B.NUM_CR, False),
        ("Net cash from operations", cq("cf_netA"), B.NUM_CR, True),
    ])
    pos = lambda key: [f"={ref('B', key)}", f"={ref('D', key)}", f"={ref('E', key)}"]
    table("2. FINANCIAL POSITION (consolidated)", ["Rs crore", "31-Mar-26", "30-Sep-26", "Change"], [
        ("Sundry debtors", pos("Sundry debtors"), B.NUM_CR, True),
        ("Inventory", pos("Inventory"), B.NUM_CR, False),
        ("Sundry creditors", pos("Sundry creditors"), B.NUM_CR, False),
        ("Gross debt (OD + loans)", pos("debt"), B.NUM_CR, True),
        ("Cash & bank", pos("Cash & bank"), B.NUM_CR, False),
        ("Net debt", pos("netdebt"), B.NUM_CR, True),
        ("Equity", pos("Equity (capital + reserves + P&L)"), B.NUM_CR, False),
        ("Debtor days (on Apr-Sep sales)", [None, f"={ref('D', 'debtor_days')}", None], "0", False),
    ])
    ent = lambda key: [f"={ca}!$D${marks[key]}", f"={ca}!$G${marks[key]}", f"={ca}!$J${marks[key]}", f"={ca}!$M${marks[key]}"]
    table("3. BY COMPANY - standalone, Apr-Sep", ["Rs crore", "Orange O Tec", "Enterprises", "Colorix", "Group"], [
        ("Turnover", ent("ent_TOTAL TURNOVER"), B.NUM_CR, False),
        ("EBITDA", ent("ent_EBITDA"), B.NUM_CR, False),
        ("Profit before depreciation & tax", ent("ent_PAT (NET PROFIT)"), B.NUM_CR, True),
    ])
    # Key messages: sentences assembled by formula, so the numbers in them are live too.
    T = lambda col, key, fmt="0.00": f'TEXT({ref(col, key)},"{fmt}")'
    msgs = [
        f'="Sales of Rs "&{T("E", "TOTAL TURNOVER")}&" cr for the half-year; Q2 was "&TEXT({ref("C", "TOTAL TURNOVER")}/'
        f'{ref("B", "TOTAL TURNOVER")}-1,"+0%;-0%")&" on Q1."',
        f'="Margins fell in Q2: EBITDA margin "&{T("B", "EBITDA margin %", "0.0%")}&" in Q1, "&{T("C", "EBITDA margin %", "0.0%")}'
        f'&" in Q2 - more lower-margin machine sales and higher commission / advertising."',
        f'="Profit before depreciation and tax Rs "&{T("E", "PAT (NET PROFIT)")}&" cr; after both, an estimated Rs "'
        f'&{T("E", "adj_pat")}&" cr (neither is booked yet)."',
        f'="Cash: operations used Rs "&TEXT(-{ref("D", "cf_netA")},"0.00")&" cr - debtors up Rs "&{T("E", "Sundry debtors")}'
        f'&" cr; gross debt up Rs "&{T("E", "debt")}&" cr to Rs "&{T("D", "debt")}&" cr."',
        f'="Debt / equity "&TEXT({ref("D", "debt")}/{ref("D", "Equity (capital + reserves + P&L)")},"0.00")&"x and cash only Rs "'
        f'&{T("D", "Cash & bank")}&" cr - collecting debtors is the priority."',
    ]
    if ageing_marks:
        # Debtors by invoice age - from the Debtors Ageing sheet (as on its snapshot date, not 30-Sep).
        sh, tr = ageing_marks["debtors"]
        a = lambda col: f"='{sh}'!${col}${tr}"
        table(f"4. DEBTORS AGEING - as on {ageing_as_of}, by invoice date", ["Rs crore", "0-180 days", "Over 180 days", "Total",
                                                                            "% over 180"], [
            ("Sundry debtors (all companies)", [a("D"), a("E"), a("F"), f"=IFERROR('{sh}'!$E${tr}/'{sh}'!$F${tr},0)"],
             B.NUM_CR, True),
        ])
        ws.cell(r - 2, 5).number_format = B.PCT  # the one data row (table() leaves a blank row after it)
    B.band(ws, r, "5. KEY MESSAGES" if ageing_marks else "4. KEY MESSAGES", 5)
    r += 1
    for m in msgs:
        cell = ws.cell(r, 1, m)
        cell.font, cell.alignment = B.F_NORM, Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 30
        r += 1
    # ---- Page 2: why the margin fell, and Orange O Tec / Enterprises side by side ------------------
    def page_break():
        nonlocal r
        ws.row_breaks.append(Break(id=r - 1))

    page_break()
    first, last, tot = marks["movers_first"], marks["movers_last"], marks["movers_total"]
    k = f"{ca}!$B${marks['turn_ratio']}"
    rows = [(f"={ca}!$A${i}", [f"={ca}!$C${i}", f"={ca}!$D${i}", f"={ca}!$E${i}", f"={ca}!$F${i}"], B.NUM_CR, False)
            for i in range(first, last + 1)]
    pat = lambda col: f"{ca}!${col}${marks['PAT (NET PROFIT)']}"
    beyond = f"={pat('C')}-{pat('B')}*{k}"  # Q2 profit less Q1 profit grown with sales
    rows += [
        ("Total of these 10 lines", [None, None, None, f"={ca}!$F${tot}"], B.NUM_CR, True),
        ("All other lines", [None, None, None, f"{beyond}-{ca}!$F${tot}"], B.NUM_CR, False),
        ("Profit change beyond sales growth", [f"={pat('B')}", f"={pat('C')}", f"={pat('B')}*{k}", beyond], B.NUM_CR, True),
    ]
    table("6. WHY THE Q2 MARGIN FELL - biggest GL lines against sales growth (consolidated)",
          ["Rs crore", "Q1", "Q2", "Q1 at Q2 sales", "Effect on profit"], rows)
    ws.cell(r - 1, 1, "Q1 amount x Q2/Q1 sales; effect = what the line cost beyond sales growth (- squeezed the margin). "
                      "Last row: profit before depreciation & tax.").font = B.F_NOTE

    # Standalone entity figures: CFO Analysis section 4 (O Tec B-D, Enterprises E-G, Colorix H-J; Q1, Q2, Apr-Sep).
    lines = [("Turnover", "ent_TOTAL TURNOVER", None), ("Gross profit", "ent_GROSS PROFIT", None),
             ("Gross margin %", "ent_GROSS PROFIT", "ent_TOTAL TURNOVER"), ("EBITDA", "ent_EBITDA", None),
             ("EBITDA margin %", "ent_EBITDA", "ent_TOTAL TURNOVER"), ("Finance cost", "ent_TOTAL FINANCE COST", None),
             ("Profit before depreciation & tax", "ent_PAT (NET PROFIT)", None),
             ("Profit margin %", "ent_PAT (NET PROFIT)", "ent_TOTAL TURNOVER")]

    def cell(col, key):
        return f"{ca}!${col}${marks[key]}"

    def entity_rows(cols_list):
        """cols_list: one list of columns per output column (summed, e.g. O Tec + Enterprises)."""
        out = []
        for name, num, den in lines:
            cells = []
            for cols in cols_list:
                n = "+".join(cell(c, num) for c in cols)
                cells.append(f"=IFERROR(({n})/({'+'.join(cell(c, den) for c in cols)}),0)" if den else f"={n}")
            out.append((name, cells, B.PCT if den else B.NUM_CR, name in ("Turnover", "EBITDA", "Profit before depreciation & tax")))
        return out

    for title, (o, e) in (("Apr-Sep", ("D", "G")), ("Q2 (Jul-Sep)", ("C", "F")), ("Q1 (Apr-Jun)", ("B", "E"))):
        table(f"7. ORANGE O TEC AND ENTERPRISES - standalone, {title}",
              ["Rs crore", "Orange O Tec", "Enterprises", "Total"], entity_rows([[o], [e], [o, e]]))

    # ---- Page 3: Colorix ---------------------------------------------------------------------------
    page_break()
    table("8. COLORIX - standalone", ["Rs crore", "Q1 (Apr-Jun)", "Q2 (Jul-Sep)", "Apr-Sep"], entity_rows([["H"], ["I"], ["J"]]))

    B.setup_print(ws, landscape=False, fit_wide=1)
    # Excel ignores manual page breaks under "fit to page", so a fixed scale keeps the three pages as laid out.
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.page_setup.scale = 78
    ws.oddHeader.right.text = "Orange O Tec Group - Directors' executive summary"
    return ws


# =============================================================================================
# Branch Transactions / Related Party
# =============================================================================================
def memo_rows(lay, groups):
    """[(kind, label, row)] from a CFS sheet's 'ELIMINATED ON CONSOLIDATION' block, keeping only the
    column-B groups asked for. kind: group / sub / ledger / subtotal / total."""
    start = lay.row("ELIMINATED ON CONSOLIDATION")
    out, group = [], None
    if not start:
        return out
    for r in range(start + 1, lay.max_row + 1):
        b, c, d = lay.cell(r, 2), lay.cell(r, 3), lay.cell(r, 4)
        if isinstance(d, str) and d.startswith("TOTAL ELIMINATED"):
            break
        if isinstance(b, str) and b.strip():
            if b.strip() != group and not (isinstance(d, str) and d.startswith("TOTAL")):
                group = b.strip()
                if group in groups:
                    out.append(("group", group, None))
        if group not in groups or not isinstance(d, str) or not d.strip():
            continue
        if isinstance(c, str) and c.strip() and not d.startswith(("Subtotal", "TOTAL")):
            out.append(("sub", c.strip(), None))
        kind = "total" if d.startswith("TOTAL") else "subtotal" if d.startswith("Subtotal") else "ledger"
        out.append((kind, d.strip(), r))
    return out


def build_memo_sheets(wb, L, entity_names):
    """Writes the two sheets; returns their Hub data {sheet: [{title, kind, rows}]}."""
    hub = {}
    for sheet, parts in MEMO_GROUPS.items():
        ws = wb.create_sheet(sheet)
        ws.sheet_properties.tabColor = "7030A0"
        ws.column_dimensions["A"].width = 50
        ws["A1"], ws["A1"].font = sheet.upper() + " - per company, month on month", B.F_TITLE
        ws["A2"] = ("Rs crore. Read from the consolidation's own 'ELIMINATED ON CONSOLIDATION' memo on the CFS sheets - "
                    "the amounts removed from the group figures. P&L: sales +, purchases -. Balance sheet lines at the period end.")
        ws["A2"].font = B.F_NOTE
        pos, c = {}, 2
        for p, _m in MEMO_PERIODS:
            B.hdr_cell(ws, 3, c, B.MOM_HEAD.get(p, p), B.FILL_TOT if p in B.MOM_HEAD else B.FILL_HDR)
            ws.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c + 7)
            for i, key in enumerate(B.BLOCK8):
                B.hdr_cell(ws, 4, c + i, entity_names[key], B.FILL_TOT if key in B.TOTAL_LABEL else B.FILL_HDR)
                ws.column_dimensions[col_letter(c + i)].width = 13
                pos[(p, key)] = c + i
            c += 8
        ws.row_dimensions[4].height = 45
        ws.freeze_panes = "B5"
        r = 6
        hub[sheet] = []
        for title, kind, groups in parts:
            lay = L["PL_CFS"] if kind == "pl" else L["BS_CFS_LAY"]
            B.band(ws, r, title, c - 1)
            r += 1
            sec = {"title": title, "kind": kind, "rows": []}
            for rk, text, src_r in memo_rows(lay, groups):
                bold = rk in ("group", "subtotal", "total")
                B.label(ws, r, text, B.F_BOLD if bold else B.F_NORM,
                        B.FILL_SUB if rk == "total" else None, indent={"group": 0, "sub": 1}.get(rk, 2))
                hub_vals = {}
                if src_r:
                    for p, m in MEMO_PERIODS:
                        for key in B.BLOCK8:
                            sc = lay.col.get(((p if kind == "pl" else m), key))
                            cell_c = pos[(p, key)]
                            if sc:
                                B.num_cell(ws, r, cell_c, f"='{lay.title}'!${col_letter(sc)}${src_r}/{int(B.CR)}", B.NUM_CR,
                                           B.F_BOLD if bold else B.F_NORM)
                            hub_vals.setdefault(key, []).append(round((lay.cell(src_r, sc) or 0.0) / B.CR, 7) if sc else 0.0)
                sec["rows"].append({"type": rk, "label": text, "values": hub_vals})
                r += 1
            hub[sheet].append(sec)
            r += 1
        B.setup_print(ws, landscape=True, fit_wide=1)
    return hub


# =============================================================================================
# Balance-sheet schedules (user, 2026-10-24): Taxes, Employees, Loans & Advances, Debtors & Creditors
# =============================================================================================
# Read ledger by ledger from the standalone balance sheet (MOM BS GL SFS), so each schedule ties to
# the balance sheet. That sheet shows liabilities + on the liabilities side and assets + on the assets
# side; a schedule mixing the two shows the asset side negative, so its net row reads
# "payable / (receivable)". Employee ledgers are taken out of creditors and loans & advances.
EMP_ADVANCE_SUBS = {"EMPLOYEE ADVANCES ACCOUNTS"}
EMP_PAYABLE_SUBS = {"CREDITOR FOR EMPLOYEE", "CREDITOR FOR EMPLOYEE (EXP)"}
TAX_ASSET_SUBS = {"Advance Tax", "TDS TCS Receivable", "TDS PAID AMT RECEIVABLE", "Balance with  Revenue Authorities",
                  "Advance Tax/Refund Due"}
LA_CURRENT_SUBS = {"Other Loans & Advances", "Advance to Partner"}
EMP_SECTION = "EMPLOYEES - payable and advances (at period end) - Rs crore"
LA_SECTION = "LOANS & ADVANCES (excl. employees) - at period end - Rs crore"


def _tax_kind(sub, gl):
    t = f"{sub} {gl}".upper()
    if "TDS" in t or "TCS" in t:
        return "TDS / TCS"
    if "GST" in t or "CESS" in t:
        return "GST"
    return "OTHERS (custom duty, ESIC / PF / PT, income tax)"


def _is_tax_asset(group, sub, gl):
    if group != "Current Assets":
        return False
    return sub in TAX_ASSET_SUBS or (sub == "Current Assets" and any(
        k in gl.upper() for k in ("TDS", "TCS", "GST", "ADVANCE TAX", "REFUND")))


def _classify(group, sub, gl):
    """(sheet, section, category, sign) for one balance-sheet ledger, or None."""
    up = gl.upper()
    if sub in EMP_ADVANCE_SUBS:
        return "Employees", EMP_SECTION, "Advances to employees (shown -)", -1
    if sub in EMP_PAYABLE_SUBS or (group == "Provisions" and "GRATUITY" in up):
        return "Employees", EMP_SECTION, "Payable to employees", 1
    if group == "Duties & Taxes" or (group == "Provisions" and "TAX" in up):
        k = _tax_kind(sub, gl)
        return "Taxes", f"{k} - at period end - Rs crore", "Payable (net of credits in the liability ledgers)", 1
    if _is_tax_asset(group, sub, gl):
        k = _tax_kind(sub, gl)
        return "Taxes", f"{k} - at period end - Rs crore", "Receivable / paid in advance (shown -)", -1
    if group == "Loans & Advances (Asset)" or (group == "Current Assets" and sub in LA_CURRENT_SUBS):
        return "Loans & Advances", LA_SECTION, sub, 1
    if group == "Deposits (Asset)":
        return "Loans & Advances", LA_SECTION, "Deposits", 1
    if group == "Sundry Debtors":
        return "Debtors & Creditors", "SUNDRY DEBTORS - party-wise, at period end - Rs crore", "Sundry debtors", 1
    if group == "Sundry Creditors":
        return "Debtors & Creditors", "SUNDRY CREDITORS - party-wise, at period end - Rs crore", "Sundry creditors", 1
    return None


SCHEDULE_NET = {"Employees": "NET PAYABLE / (ADVANCE) TO EMPLOYEES", "Taxes": "NET PAYABLE / (RECEIVABLE)",
                "Loans & Advances": "TOTAL LOANS & ADVANCES", "Debtors & Creditors": None}
SCHEDULE_SHEETS = ["Debtors & Creditors", "Loans & Advances", "Employees", "Taxes"]


def bs_ledgers(lay):
    """(group, sub, ledger, row) for every ledger row of MOM BS GL SFS (group / sub carried down)."""
    group = sub = None
    for r in range(5, lay.max_row + 1):
        b, c, d = lay.cell(r, 2), lay.cell(r, 3), lay.cell(r, 4)
        if isinstance(b, str) and b.strip():
            group = b.strip()
        if isinstance(c, str) and c.strip():
            sub = c.strip()
        if isinstance(d, str) and d.strip() and not d.startswith(("Subtotal", "TOTAL")) and group:
            yield group, sub or group, d.strip(), r


def _add(a, b):
    out = dict(a)
    for k, vs in b.items():
        out[k] = [x + y for x, y in zip(out.get(k, [0.0] * len(vs)), vs)]
    return out


def build_schedule_sheets(wb, L, entity_names):
    """Writes the four schedule sheets; returns their Hub data in the memo format
    {sheet: [{title, kind: "bs", net_label, rows: [group / sub / ledger / subtotal / total]}]}."""
    lay = L["BS_SFS_LAY"]
    tree = {}  # sheet -> section -> category -> sub -> [(ledger, row, sign)]
    for group, sub, gl, r in bs_ledgers(lay):
        hit = _classify(group, sub, gl)
        if hit:
            sheet, section, cat, sign = hit
            tree.setdefault(sheet, {}).setdefault(section, {}).setdefault(cat, {}).setdefault(sub, []).append((gl, r, sign))

    hub = {}
    for sheet in SCHEDULE_SHEETS:
        ws = wb.create_sheet(sheet)
        ws.sheet_properties.tabColor = "548235"
        ws.sheet_properties.outlinePr.summaryBelow = True
        ws.column_dimensions["A"].width = 55
        ws["A1"], ws["A1"].font = sheet.upper() + " - per company, at each month end", B.F_TITLE
        ws["A2"] = ("Rs crore. Ledger by ledger from MOM BS GL SFS, so it ties to the balance sheet. Where a schedule mixes "
                    "payables and receivables the asset side is shown -. Employee ledgers are kept out of creditors and "
                    "loans & advances. Ledger rows are grouped: use the outline buttons to open them.")
        ws["A2"].font = B.F_NOTE
        pos, c = {}, 2
        for p, _m in MEMO_PERIODS:
            B.hdr_cell(ws, 3, c, B.MOM_HEAD.get(p, p), B.FILL_TOT if p in B.MOM_HEAD else B.FILL_HDR)
            ws.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c + 7)
            for i, key in enumerate(B.BLOCK8):
                B.hdr_cell(ws, 4, c + i, entity_names[key], B.FILL_TOT if key in B.TOTAL_LABEL else B.FILL_HDR)
                ws.column_dimensions[col_letter(c + i)].width = 13
                pos[(p, key)] = c + i
            c += 8
        ws.row_dimensions[4].height = 45
        ws.freeze_panes = "B5"
        r = 6

        def put(text, font, fill, indent, cells, outline=0):
            nonlocal r
            B.label(ws, r, text, font, fill, indent=indent)
            for pk, f in cells.items():
                B.num_cell(ws, r, pos[pk], f, B.NUM_CR, font)
            if outline:
                ws.row_dimensions[r].outlineLevel = outline
                ws.row_dimensions[r].hidden = True
            r += 1
            return r - 1

        rnd = lambda vals: {k: [round(x, 7) for x in v] for k, v in vals.items()}
        hub[sheet] = []
        # Debtors before creditors; otherwise the balance sheet's own order.
        sections = sorted(tree.get(sheet, {}).items(), key=lambda kv: "DEBTORS" not in kv[0])
        for section, cats in sections:
            B.band(ws, r, section, c - 1)
            r += 1
            sec = {"title": section, "kind": "bs", "net_label": SCHEDULE_NET[sheet], "rows": []}
            total_rows = []
            for cat, subs in cats.items():
                put(cat, B.F_BOLD, None, 0, {})
                sec["rows"].append({"type": "group", "label": cat, "values": {}})
                sub_rows, cat_vals = [], {}
                for sub, ledgers in subs.items():
                    sec["rows"].append({"type": "sub", "label": sub, "values": {}})
                    first, sub_vals = r, {}
                    for gl, sr, sign in ledgers:
                        cells, vals = {}, {}
                        for p, m in MEMO_PERIODS:
                            for key in B.BLOCK8:
                                sc = lay.col.get((m, key))
                                v = (lay.cell(sr, sc) or 0.0) if sc else 0.0
                                if sc:
                                    cells[(p, key)] = f"={'-' if sign < 0 else ''}'{lay.title}'!${col_letter(sc)}${sr}/{int(B.CR)}"
                                vals.setdefault(key, []).append(round(sign * v / B.CR, 7))
                        put(gl, B.F_NORM, None, 2, cells, outline=1)
                        sec["rows"].append({"type": "ledger", "label": gl, "values": vals})
                        sub_vals = _add(sub_vals, vals)
                    last = r - 1
                    cells = {pk: f"=SUM({col_letter(cc)}{first}:{col_letter(cc)}{last})" for pk, cc in pos.items()}
                    sub_rows.append(put(sub, B.F_BOLD, None, 1, cells))
                    sec["rows"].append({"type": "subtotal", "label": f"Subtotal: {sub}", "values": rnd(sub_vals)})
                    cat_vals = _add(cat_vals, sub_vals)
                cells = {pk: "=" + "+".join(f"{col_letter(cc)}{sr}" for sr in sub_rows) for pk, cc in pos.items()}
                total_rows.append(put(f"TOTAL: {cat}", B.F_BOLD, B.FILL_SUB, 0, cells))
                sec["rows"].append({"type": "total", "label": f"TOTAL: {cat}", "values": rnd(cat_vals)})
            if SCHEDULE_NET[sheet] and len(total_rows) > 1:
                cells = {pk: "=" + "+".join(f"{col_letter(cc)}{tr}" for tr in total_rows) for pk, cc in pos.items()}
                put(SCHEDULE_NET[sheet], B.F_BOLD, B.FILL_TOT, 0, cells)
            hub[sheet].append(sec)
            r += 1
        B.setup_print(ws, landscape=True, fit_wide=1)
    return hub
