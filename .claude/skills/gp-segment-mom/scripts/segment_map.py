"""Segment mapping for the GP Segment MoM report.

Maps Tally stock groups -> the five reporting segments the board uses:
    MACHINES, HEADS, INK, SPARE PARTS, OTHER

and Tally ledger names -> the same five segments, for the discount /
direct-cost overlay that sits on top of the item-based gross profit.

Both maps are exported to the workbook's "Mapping" sheet so the user can see
exactly how every group and ledger was classified.
"""

from __future__ import annotations

import re

SEGMENTS = ["MACHINES", "HEADS", "INK", "SPARE PARTS", "OTHER"]

# ---------------------------------------------------------------------------
# Stock group -> segment
#
# Keys are the top-level stock group as it appears in Tally (upper-cased).
# Every group observed across the five companies as at 30-Jun-2026 is listed
# explicitly so that nothing silently falls into OTHER without us knowing.
# ---------------------------------------------------------------------------
STOCK_GROUP_SEGMENT = {
    # --- Machines -----------------------------------------------------------
    "MACHINES": "MACHINES",
    "MANUFACTURING STOCK": "MACHINES",
    "FINISHED GOODS": "MACHINES",
    "SHENZHEN RUNTIANZHI DIGITAL EQUIPMENT CO., LTD": "MACHINES",
    # --- Heads --------------------------------------------------------------
    "MACHINERY HEAD": "HEADS",
    "KYOCERA PRINTHEADS": "HEADS",
    # --- Ink ----------------------------------------------------------------
    "PRINTING INK": "INK",
    "DIGITAL PRINTING INK": "INK",
    "OTHER INK": "INK",
    "PROVISION INK": "INK",
    "PADDING CHEMICAL": "INK",
    "DIGISTAR (BIB)": "INK",
    "DRUMS": "INK",
    # --- Spare parts --------------------------------------------------------
    "MACHINERY PARTS": "SPARE PARTS",
    "SUBPRO-2 SPARE PARTS": "SPARE PARTS",
    "HANGLORY SPARE PARTS": "SPARE PARTS",
    "HOMER SPARE PARTS": "SPARE PARTS",
    "S3200  EPSON SPARE PARTS": "SPARE PARTS",
    "S3200 EPSON SPARE PARTS": "SPARE PARTS",
    "FABRO PARTS": "SPARE PARTS",
    "MT DRIVES": "SPARE PARTS",
    "LIGHT EXPENSE": "SPARE PARTS",
    # --- Other --------------------------------------------------------------
    "PAPER ROLL": "OTHER",
    "SUBLIMATION PAPER": "OTHER",
    "FABRIC": "OTHER",
    "SAMPLE FABRIC": "OTHER",
    "POWDER": "OTHER",
    "SOFTWARE": "OTHER",
    "PACKING MATERIAL": "OTHER",
    "PACKING MATERIAL STOCK": "OTHER",
    "LABORATORY STOCK": "OTHER",
    "STOCK LYING AT THIRD PARTY": "OTHER",
    "AT JOB WORKERS PREMISES": "OTHER",
    "GIT-2526": "OTHER",
    "(NO STOCK GROUP ASSIGNED)": "OTHER",
}

# Fallback keyword rules, applied in order, when a group is not in the table
# above (protects against new groups being created in Tally after this was
# written). Anything still unmatched lands in OTHER and is reported as
# "unmapped" on the Cross Check sheet.
_KEYWORD_RULES = [
    (r"\bHEAD", "HEADS"),
    (r"PRINTHEAD", "HEADS"),
    (r"\bINK\b", "INK"),
    (r"CHEMICAL", "INK"),
    (r"SPARE|PARTS\b|DRIVES", "SPARE PARTS"),
    (r"MACHINE|MANUFACTUR|FINISHED GOODS", "MACHINES"),
    (r"PAPER|FABRIC|POWDER|SOFTWARE|PACKING|LABORATORY", "OTHER"),
]


# ---------------------------------------------------------------------------
# Company-specific overrides.
#
# The same group name means different things in different companies. Orange O
# Tec Enterprises is the group's INK MANUFACTURER, so its "Finished Goods" and
# "Manufacturing Stock" are finished ink and ink WIP - not machines. Confirmed
# with the user and corroborated by Q1 item mix (95% of Enterprises HO outward
# sits in those two groups, and Enterprises sells ink).
# ---------------------------------------------------------------------------
COMPANY_GROUP_OVERRIDES = {
    "ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)": {
        "FINISHED GOODS": "INK",
        "MANUFACTURING STOCK": "INK",
        "LABORATORY STOCK": "INK",
        "DRUMS": "INK",
        "PROVISION INK": "INK",
    },
    "ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2024-26)": {
        "FINISHED GOODS": "INK",
        "MANUFACTURING STOCK": "INK",
        "LABORATORY STOCK": "INK",
        "DRUMS": "INK",
        "PROVISION INK": "INK",
    },
    "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27": {
        "FINISHED GOODS": "INK",
        "MANUFACTURING STOCK": "INK",
        # Noida genuinely sells machines - ~46% of its Q1 outward - so this
        # group is deliberately NOT folded into ink.
        "DIGITAL PRINTING MACHINES STOCK": "MACHINES",
    },
    "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA - FY25-26": {
        "FINISHED GOODS": "INK",
        "MANUFACTURING STOCK": "INK",
        "DIGITAL PRINTING MACHINES STOCK": "MACHINES",
    },
}

# Groups seen in Tally that the base table did not cover.
STOCK_GROUP_SEGMENT.update({
    "MACHINERY": "MACHINES",
    "DIGITAL PRINTING MACHINES STOCK": "MACHINES",
    "MACHINERY & MACHINERY PARTS": "MACHINES",
    "REACTIVE INK (FLOTEC)": "INK",
    "SUBLIMATION INK": "INK",
    "PROVISION INK-PAYABLE-FOC": "INK",
    "PROVISION INK-RECEIVABLE": "INK",
    "PRINTHEAD": "HEADS",
    "PRIMARY": "OTHER",
})


def segment_for_stock_group(group: str | None,
                            company: str | None = None) -> tuple[str, bool]:
    """Return (segment, was_explicitly_mapped) for a single Tally stock group."""
    key = (group or "(No stock group assigned)").strip().upper()
    if company:
        override = COMPANY_GROUP_OVERRIDES.get(company, {})
        if key in override:
            return override[key], True
    if key in STOCK_GROUP_SEGMENT:
        return STOCK_GROUP_SEGMENT[key], True
    for pattern, segment in _KEYWORD_RULES:
        if re.search(pattern, key):
            return segment, False
    return "OTHER", False


def segment_for_item(leaf_group: str | None, parents: dict,
                     company: str | None = None) -> tuple[str, bool, str]:
    """Resolve an item's segment by walking leaf -> root.

    Returns (segment, was_explicitly_mapped, group_used).

    Walking upward and stopping at the FIRST explicit match preserves
    granularity: Orange O Tec Noida nests both machines and spares under a
    container group called "MACHINERY & MACHINERY PARTS", so collapsing
    straight to the top level would merge two segments that must stay apart.
    """
    seen = set()
    current = (leaf_group or "").strip()
    fallback = current or "(No stock group assigned)"
    while current:
        key = current.upper()
        if key in seen:
            break
        seen.add(key)
        segment, explicit = segment_for_stock_group(current, company)
        if explicit:
            return segment, True, current
        parent = parents.get(key, "")
        if not parent:
            break
        current = parent
    segment, explicit = segment_for_stock_group(fallback, company)
    return segment, explicit, fallback


# ---------------------------------------------------------------------------
# Ledger -> segment
#
# Tally suffixes segment onto sales, discount and direct-cost ledgers, e.g.
#   GST SALES ACCOUNTS - MACHINE / DISCOUNT & RATE DIFFERENCE@18%-HEAD
#   CUSTOM DUTY EXP ACCOUNTS-INK / OCEAN FREIGHT CHARGES @ 5% (SPARE)
# The suffix is not perfectly consistent, so match on token rather than an
# exact trailing string.
# ---------------------------------------------------------------------------
_LEDGER_TOKENS = [
    # order matters: check the more specific tokens first
    ("SPAREPARTS", "SPARE PARTS"),
    ("SPARE PARTS", "SPARE PARTS"),
    ("SPARE", "SPARE PARTS"),
    ("MACHINERY", "MACHINES"),
    ("MACHINE", "MACHINES"),
    ("HEAD", "HEADS"),
    ("INK", "INK"),
    ("PAPER", "OTHER"),
    ("FABRIC", "OTHER"),
    ("POWDER", "OTHER"),
    ("DRYER", "OTHER"),
    ("PACKING MATERIAL", "OTHER"),
]


def segment_for_ledger(ledger: str | None) -> str | None:
    """Return the segment a ledger is tagged to, or None if it carries no tag.

    None means the ledger is segment-agnostic (e.g. a plain "GST PURCHASE"
    used by Colorix and Enterprises Noida) and must be reported as
    unallocated rather than guessed at.
    """
    if not ledger:
        return None
    name = ledger.strip().upper()
    for token, segment in _LEDGER_TOKENS:
        # token must appear as a suffix-style tag, not incidentally mid-word
        if re.search(rf"[-(\s]{re.escape(token)}\b", name):
            return segment
    return None


# ---------------------------------------------------------------------------
# Tally company file -> legal entity / location
# Five Tally files collapse to three legal entities; HO<->Noida transfers
# eliminate inside the entity, not at group level.
# ---------------------------------------------------------------------------
COMPANY_ENTITY = {
    "ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)": ("ORANGE O TEC", "HO"),
    "ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)": ("ORANGE O TEC", "Noida"),
    "ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)": ("OOT ENTERPRISES", "HO"),
    "ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27": ("OOT ENTERPRISES", "Noida"),
    "COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)": ("COLORIX", "-"),
}
