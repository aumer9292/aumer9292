"""Which set of financial statements is this? Classification of the entity under the
Third Schedule to the Companies Act 2017 (Pakistan), and the reporting framework
that goes with it.

The tool suggests a category from the figures and wording it finds, and the
user confirms it before any disclosure checking is done. The thresholds below
are defaults: check them against the current SECP notifications and change
them in ``~/.fsreview/thresholds.json`` if they have been revised.
"""

import json
import os
import re
from decimal import Decimal

CATEGORIES = {
    "PIC": {
        "name": "Public Interest Company",
        "framework": "IFRS as notified under the Companies Act 2017",
        "schedule": "Fourth Schedule to the Companies Act 2017",
        "who": "Listed companies; non-listed financial institutions, insurers, public sector companies, public "
               "utilities and companies holding assets in a fiduciary capacity for a broad group of outsiders.",
    },
    "LSC": {
        "name": "Large Sized Company",
        "framework": "IFRS as notified under the Companies Act 2017",
        "schedule": "Fifth Schedule to the Companies Act 2017",
        "who": "Non-listed company with paid-up capital, turnover or employees at or above the large thresholds.",
    },
    "MSC": {
        "name": "Medium Sized Company",
        "framework": "IFRS for SMEs",
        "schedule": "Fifth Schedule to the Companies Act 2017",
        "who": "Non-listed company that is neither large nor small.",
    },
    "SSC": {
        "name": "Small Sized Company",
        "framework": "Revised AFRS for Small-Sized Entities (ICAP)",
        "schedule": "Fifth Schedule to the Companies Act 2017",
        "who": "Private company within all of the small thresholds.",
    },
}

DEFAULT_THRESHOLDS = {
    "note": "Defaults from the Third Schedule to the Companies Act 2017 as commonly summarised. Verify against the "
            "current SECP notification and edit this file if they have been revised. Amounts in rupees.",
    "large": {"paid_up_capital": 200_000_000, "turnover": 1_000_000_000, "employees": 750},
    "small": {"paid_up_capital": 10_000_000, "turnover": 100_000_000, "employees": 250},
}

HOME = os.path.join(os.path.expanduser("~"), ".fsreview")


def thresholds():
    path = os.path.join(HOME, "thresholds.json")
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return {**DEFAULT_THRESHOLDS, **data}
    except (OSError, ValueError):
        return DEFAULT_THRESHOLDS


def unit_multiplier(doc, settings=None):
    """1, 1,000 or 1,000,000 depending on how the statements are presented."""
    texts = [(settings or {}).get("unit") or ""]
    texts += [t.unit_label for t in doc.tables() if t.unit_label]
    texts += [p.text for p in doc.paragraphs() if p.role == "unit"]
    for t in doc.tables():
        for r in t.rows:
            if r.kind == "header":
                texts.append(" ".join(r.cells))
    joined = " ".join(texts).lower()
    if re.search(r"million|\bmn\b|\(m\)|'m\b", joined):
        return Decimal(1_000_000)
    if re.search(r"['’]000|thousand", joined):
        return Decimal(1000)
    return Decimal(1)


def _employees(text):
    m = re.search(r"(?:total )?number of (?:persons |permanent )?employees[^.\d]{0,80}?(\d{1,3}(?:,\d{3})*|\d+)",
                  text, re.I)
    if m:
        return int(m.group(1).replace(",", ""))
    m = re.search(r"(\d{1,3}(?:,\d{3})*|\d+)\s+(?:permanent )?employees", text, re.I)
    return int(m.group(1).replace(",", "")) if m else None


def suggest(doc, figures, settings=None):
    """Return the suggested category, the reasons, and the facts used."""
    text = doc.text()
    mult = unit_multiplier(doc, settings)
    turnover = figures.get("revenue", (None, None))
    capital = figures.get("share_capital", (None, None))
    # classification uses the preceding year's audited figures; the comparative is that year
    def amt(pair):
        v = pair[1] if pair[1] is not None else pair[0]
        return abs(v) * mult if v is not None else None
    t, c = amt(turnover), amt(capital)
    emp = _employees(text)
    listed = bool(re.search(r"listed on (?:the )?(?:pakistan|karachi|lahore|islamabad) stock exchange|"
                            r"shares (?:are|of the company are) (?:listed|quoted)|\blisted company\b", text, re.I))
    private = bool(re.search(r"\(private\) limited|private limited company|\(pvt\.?\) ltd|"
                             r"incorporated .{0,60}as a private", text, re.I)) or \
        bool(re.search(r"\(private\)|\(pvt", (settings or {}).get("entity_name") or "", re.I))
    regulated = re.search(r"\b(banking company|insurance company|modaraba|non-banking finance|nbfc|"
                          r"asset management company|public sector company|public utility|mutual fund)\b", text, re.I)
    th = thresholds()
    reasons = []
    if listed:
        reasons.append("The statements say the company is listed.")
        cat = "PIC"
    elif regulated:
        reasons.append(f"The statements describe a '{regulated.group(1)}', which is a public interest company "
                       "even if not listed.")
        cat = "PIC"
    else:
        big = []
        if c is not None and c >= th["large"]["paid_up_capital"]:
            big.append(f"paid-up capital Rs {c:,.0f}")
        if t is not None and t >= th["large"]["turnover"]:
            big.append(f"turnover Rs {t:,.0f}")
        if emp is not None and emp >= th["large"]["employees"]:
            big.append(f"{emp} employees")
        small = (private or c is not None) and \
            (c is None or c <= th["small"]["paid_up_capital"]) and \
            (t is None or t <= th["small"]["turnover"]) and \
            (emp is None or emp <= th["small"]["employees"]) and (c is not None or t is not None)
        if big:
            cat = "LSC"
            reasons.append("At or above a large-company threshold: " + ", ".join(big) + ".")
        elif small and private:
            cat = "SSC"
            reasons.append("A private company within every small-company threshold found.")
        else:
            cat = "MSC"
            reasons.append("Not listed, below the large-company thresholds and not within all small thresholds.")
    facts = {
        "listed": listed, "private": private,
        "turnover": float(t) if t is not None else None,
        "paid_up_capital": float(c) if c is not None else None,
        "employees": emp,
        "unit_multiplier": int(mult),
        "thresholds": th,
    }
    missing = [n for n, v in (("turnover", t), ("paid-up capital", c), ("number of employees", emp)) if v is None]
    if missing:
        reasons.append("Not found in the statements: " + ", ".join(missing) + " — please confirm the category.")
    reasons.append("Classification is based on the preceding year's audited figures (the comparative column).")
    return {"suggested": cat, "reasons": reasons, "facts": facts,
            "categories": {k: dict(v) for k, v in CATEGORIES.items()}}
