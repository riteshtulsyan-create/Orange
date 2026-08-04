"""
Turns a vendor's raw Tally narrations into one short, concise Remarks line -
clustering near-duplicate narrations (e.g. the same wording repeated with a
different month) into a single theme, instead of picking "most frequent + a
count of how many others exist".
"""
import re

MONTH_TOKENS = r"(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|SEPT|OCT|NOV|DEC)[A-Z]*[-\s]?\d{2,4}"

def normalize(text):
    t = text.upper()
    t = re.sub(MONTH_TOKENS, " ", t)
    t = re.sub(r"\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b", " ", t)          # dates dd.mm.yyyy
    t = re.sub(r"\bFY\s?\d{2,4}[-/]\d{2,4}\b", " ", t)                 # FY 2025-26
    t = re.sub(r"[\d,]+\.?\d*\s*%?", " ", t)                          # bare numbers/amounts/%
    t = re.sub(r"[^A-Z& ]+", " ", t)                                   # punctuation
    t = re.sub(r"\s+", " ", t).strip()
    return t

def cluster_narrations(narrations):
    """narrations: list of raw strings.
    Returns distinct underlying themes, ordered by first appearance, longest
    original text kept per cluster."""
    clusters = {}
    order = []
    for n in narrations:
        core = normalize(n)
        if not core:
            continue
        if core not in clusters:
            clusters[core] = n
            order.append(core)
        elif len(n) > len(clusters[core]):
            clusters[core] = n
    return [clusters[c] for c in order]

def title_clean(s, max_len=90):
    s = s.strip()
    if len(s) > max_len:
        s = s[:max_len - 3].rsplit(" ", 1)[0] + "..."
    return s

def draft_remark(narrations, fallback):
    """narrations: list of raw narration strings for this vendor (all months).
    Produces a short remark covering all distinct reasons, no narration counts."""
    narrations = [n.strip() for n in narrations if n and n.strip() and n.strip().upper() != "ROUND OFF"]
    if not narrations:
        return fallback
    themes = cluster_narrations(narrations)
    if not themes:
        # every narration stripped down to nothing (e.g. just a month name like
        # "APRIL-26") - that's not an explanation, so say so plainly instead of
        # echoing a bare month/date string (Excel would also auto-convert a bare
        # date-like string to a real date if written as-is)
        return "Recurring monthly charge - no descriptive narration in ledger"
    if len(themes) == 1:
        return title_clean(themes[0], 140)
    parts = []
    total_len = 0
    for th in themes:
        piece = title_clean(th, 70)
        if total_len + len(piece) > 190 and parts:
            break
        parts.append(piece)
        total_len += len(piece)
    return "; ".join(parts)
