"""Shared Tally access + entity mapping for the Inventory Valuation check."""
import os
from pathlib import Path
import re
import time
import xml.etree.ElementTree as ET
import requests

# Connection comes from the project .env (TALLY_HOST / TALLY_PORT), which is the single
# source of truth -- see CLAUDE.md "Tally Prime Connection". Real environment variables win.
# This file sits at .../PL and BS/Stock Report/.claude/skills/<skill>/scripts/tally_lib.py,
# so the project root is 5 levels up.
_ENV_FILE = Path(__file__).resolve().parents[5] / ".env"
if _ENV_FILE.exists():
    for _line in _ENV_FILE.read_text(encoding="utf-8").splitlines():
        _line = _line.strip()
        if not _line or _line.startswith("#") or "=" not in _line:
            continue
        _k, _v = _line.split("=", 1)
        os.environ.setdefault(_k.strip(), _v.strip())

HOST = os.environ.get("TALLY_HOST", "localhost")
PORT = int(os.environ.get("TALLY_PORT", "9000"))
URL = f"http://{HOST}:{PORT}"

FROM_DATE = "20260401"
TO_DATE = "20260630"

# 5 Tally files -> 3 legal entities
FILES = [
    ("ORANGE O TEC PRIVATE LIMITED (01-04-25TO31-03-27)", "Orange O Tec", "HO"),
    ("ORANGE O TEC PRIVATE LIMITED-NOIDA-(from 1-Apr-25)", "Orange O Tec", "Noida"),
    ("ORANGE O TEC ENTERPRISES PVT LTD(F.Y.2026-27)", "Enterprises", "HO"),
    ("ORANGE O TEC ENTERPRISES PRIVATE LIMITED-NOIDA -FY 26-27", "Enterprises", "Noida"),
    ("COLORIX DIGITAL PRINTING SOLUTIONS LLP - (from 1-Apr-20)", "Colorix", "-"),
]

def project_root() -> Path:
    """Walk up from this script until the folder holding .env is found."""
    here = Path(__file__).resolve()
    for parent in here.parents:
        if (parent / ".env").is_file():
            return parent
    return here.parents[5]


ROOT = project_root()
# intermediate extracts + raw XML cache live beside the skill, not in Inventory\
SCRATCH = Path(__file__).parent / "_work"
SCRATCH.mkdir(exist_ok=True)


def post(xml: str, timeout=900, retries=5) -> str:
    """Tally drops the connection under rapid large requests — retry with backoff.

    A fresh session per attempt avoids reusing a socket Tally has already closed.
    """
    last = None
    for attempt in range(retries):
        if attempt:
            time.sleep(3 * attempt)
        try:
            with requests.Session() as s:
                s.headers.update({"Connection": "close"})
                r = s.post(URL, data=xml.encode("utf-8"),
                           headers={"Content-Type": "text/xml;charset=utf-8"},
                           timeout=timeout)
            r.raise_for_status()
            if len(r.text) < 200 and "LINEERROR" in r.text.upper():
                raise RuntimeError(f"Tally returned an error: {r.text[:200]}")
            time.sleep(1.5)          # let Tally settle before the next request
            return r.text
        except Exception as e:                      # noqa: BLE001 - retry anything
            last = e
            print(f"    retry {attempt + 1}/{retries} after {type(e).__name__}")
    raise RuntimeError(f"Tally request failed after {retries} attempts") from last


def clean_xml(txt: str) -> str:
    """Tally emits raw control chars, bare ampersands, and undeclared namespace
    prefixes (e.g. <UDF:FOO>), any of which break a strict XML parser."""
    txt = re.sub(r"&#[0-9]+;", "", txt)
    txt = re.sub(r"&(?!(amp|lt|gt|quot|apos);)", "&amp;", txt)
    # drop undeclared prefixes on both open and close tags
    txt = re.sub(r"<(/?)[A-Za-z_][\w.-]*:", r"<\1", txt)
    # ...and any xmlns-less prefixed attributes they carry
    txt = re.sub(r'\s[A-Za-z_][\w.-]*:[A-Za-z_][\w.-]*\s*=\s*"[^"]*"', "", txt)
    txt = "".join(ch for ch in txt if ch >= " " or ch in "\t\n\r")
    return txt


def parse(txt: str) -> ET.Element:
    return ET.fromstring(clean_xml(txt))


def num(s):
    """'-1182707.90' -> float; '' / None -> 0.0"""
    if s is None:
        return 0.0
    s = str(s).strip().replace(",", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def qty(s):
    """' 3005.000 KGS' -> (3005.0, 'KGS'). Handles '1 PCS = 4 NOS' style by taking first."""
    if s is None:
        return 0.0, ""
    s = str(s).strip().replace(",", "")
    if not s:
        return 0.0, ""
    s = s.split("=")[0].strip()
    m = re.match(r"^(-?[\d.]+)\s*(.*)$", s)
    if not m:
        return 0.0, ""
    return num(m.group(1)), m.group(2).strip()


def rate(s):
    """'393.58/KGS' -> 393.58"""
    if s is None:
        return 0.0
    return num(str(s).split("/")[0])


COLL = """<ENVELOPE>
<HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>{cid}</ID></HEADER>
<BODY><DESC><STATICVARIABLES>
<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
<SVCURRENTCOMPANY>{company}</SVCURRENTCOMPANY>
<SVFROMDATE TYPE="Date">{frm}</SVFROMDATE>
<SVTODATE TYPE="Date">{to}</SVTODATE>
</STATICVARIABLES>
<TDL><TDLMESSAGE>
<COLLECTION NAME="{cid}" ISMODIFY="No">
{body}
</COLLECTION>
</TDLMESSAGE></TDL>
</DESC></BODY></ENVELOPE>"""


def collection(company, cid, body, frm=FROM_DATE, to=TO_DATE, timeout=900):
    return post(COLL.format(cid=cid, company=company, frm=frm, to=to, body=body),
                timeout=timeout)
