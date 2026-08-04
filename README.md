# PL and BS — Tally Prime Financial Reporting Toolkit

Automation for building month-on-month **Profit & Loss** and **Balance Sheet** reporting across
five Tally Prime company files, plus the surrounding CFO-level analysis packs (gross-profit
segmentation, working capital, fund flow, inventory valuation and ageing, quarterly board review).

Everything here talks to a locally running **Tally Prime** instance over its XML/HTTP gateway,
pulls general-ledger data, and renders formatted Excel deliverables.

> **Note on data:** this repository contains **code and documentation only**. All financial
> outputs — general ledger extracts, balance sheets, bank vouchers, stock and fund-flow workbooks —
> are deliberately excluded by a deny-by-default [`.gitignore`](.gitignore) and never leave the
> local machine.

---

## Architecture

The toolkit is built as a set of **Claude Code skills**. Each skill is a self-contained folder with
a `SKILL.md` specification and a `scripts/` directory holding its extract and build steps.

```
.claude/skills/
├── fetch-pandl/               P&L + Balance Sheet extract and month-on-month workbook
├── fetch-tb-masters/          Chart of Accounts / trial balance masters
├── fetch-bank-transactions/   Bank payment & receipt vouchers → flat + pivot reports
├── company-review/            Quarterly board pack, action plan, HTML dashboard
├── gp-segment-mom/            Gross profit by segment, month on month, item-wise
├── gp-executive-summary/      Executive summary layer over the GP segmentation
├── working-capital-review/    Working capital analysis workbook
└── skill-creator/             Tooling to author, evaluate and benchmark the skills above

Stock Report/.claude/skills/
├── fetch-stock-summary/       Closing inventory with full stock-group hierarchy
├── inventory-valuation/       FIFO / weighted-average / last-purchase rebuild vs Tally carrying value
└── inventory-ageing/          Six-bucket FIFO ageing, months-of-cover, slow-moving provision

FUND FLOW/                     Fund flow analysis builders (monthly and month-on-month)
Files/EXPENSES REFERESH/       Expense sheet extract, refresh and gap-check scripts
```

## Core scripts

| Script | Skill | Purpose |
|---|---|---|
| `fetch_gl_data.py` | `fetch-pandl` | Pull GL data per company; `--probe` lists which Tally companies are open |
| `fetch_pandl.py` | `fetch-pandl` | Simple Profit & Loss extract |
| `build_gl_detail_sidebyside.py` | `fetch-pandl` | Side-by-side GL detail workbook across entities |
| `build_mom_pl_bs_26_27.py` | `fetch-pandl` | Month-on-month P&L/BS deliverable (Main / Sub / Full-GL levels) |
| `build_gl_reconciliation_report.py` | `fetch-pandl` | Change reconciliation vs the previous refresh |
| `fetch_tb_masters.py` | `fetch-tb-masters` | Chart of Accounts masters for all open companies |
| `fetch-bank-transactions.mjs` | `fetch-bank-transactions` | Bank voucher extract (Node) |
| `build_board_pack.py` | `company-review` | Quarterly board pack workbook |
| `fetch_stock_summary.py` | `fetch-stock-summary` | Closing stock with group hierarchy |
| `build_fund_flow_mom.py` | `FUND FLOW` | Month-on-month fund flow |

## Reporting entities

Five Tally company files consolidating to three legal entities:

- **Orange O Tec Private Limited** — HO + Noida
- **Orange O Tec Enterprises Private Limited** — HO + Noida
- **Colorix Digital Printing Solutions LLP** — standalone

Because HO and Noida collapse into a single legal entity, inter-branch stock transfers eliminate
*within* the entity rather than at group consolidation.

Scripts accept `--period 2627`, `2526` or `2425` to restrict the run to the company files valid for
that fiscal year.

## Setup

**Requirements:** Python 3.12+, Node.js, and Tally Prime running locally with its HTTP gateway
enabled (`F1 > Settings > Connectivity`).

```bash
pip install openpyxl pandas requests python-dotenv formulas
npm install
```

Create a `.env` file in the project root — it is the single source of truth for the connection and
is not committed:

```ini
TALLY_HOST=localhost
TALLY_PORT=9000
```

## Usage

Run any script from the project root so `.env` resolves:

```bash
# Check which companies Tally currently has open — do this before trusting any refresh
python .claude/skills/fetch-pandl/scripts/fetch_gl_data.py --probe

# Full P&L / Balance Sheet refresh
python .claude/skills/fetch-pandl/scripts/fetch_pandl.py

# Chart of Accounts masters
python .claude/skills/fetch-tb-masters/scripts/fetch_tb_masters.py

# Closing stock for one company
python "Stock Report/.claude/skills/fetch-stock-summary/scripts/fetch_stock_summary.py" --company "<exact Tally company name>"

# Bank vouchers
node .claude/skills/fetch-bank-transactions/scripts/fetch-bank-transactions.mjs
```

## Operating notes

- **Always verify which companies are open before a refresh.** A wrong company file fails silently —
  it returns plausible figures for the wrong entity rather than erroring. Run `--probe` and diff the
  result against the expected period before fetching.
- **Connect on `localhost`**, not the network server IP. The two share a port number but are not the
  same target and the network server may not have every company open.
- **Refreshes are dated, not overwritten.** Each run produces `PL_BS_GL_Detail_<YYYY-MM-DD_HHMM>.xlsx`
  plus a matching `Profit_Change_Reconciliation_*.xlsx`; the superseded version is archived once the
  new one is confirmed good.
- Cross-check sheets must be validated with the `formulas` engine rather than cached Excel values.

See [`CLAUDE.md`](CLAUDE.md) for the full working agreement, and each skill's `SKILL.md` for its
detailed specification.
