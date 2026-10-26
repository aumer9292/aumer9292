"""Reading and writing figures the way financial statements show them."""

import re
from decimal import Decimal, InvalidOperation

DASHES = {"-", "–", "—", "‒", "nil", "Nil", "NIL", "--", "−"}
_NUM = re.compile(r"""^\s*
    (?P<sign>[-−–]?)\s*
    (?P<open>\()?\s*
    (?P<sign2>[-−–]?)\s*
    (?P<cur>[$€£¥₹]|Rs\.?|PKR|USD|AED|GBP|EUR)?\s*
    (?P<num>\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?|\.\d+)
    \s*(?P<close>\))?\s*$""", re.X)
YEAR = re.compile(r"^(19|20)\d{2}$")


def parse_amount(text):
    """Return a Decimal for a figure, Decimal(0) for a dash, or None.

    Accepts 1,234 / (1,234) / -1,234 / 1234.50 / - / nil. Percentages and
    text are not figures.
    """
    if text is None:
        return None
    t = str(text).strip().replace(" ", " ")
    if not t:
        return None
    if t in DASHES:
        return Decimal(0)
    m = _NUM.match(t)
    if not m:
        return None
    if bool(m.group("open")) != bool(m.group("close")):
        return None
    try:
        v = Decimal(m.group("num").replace(",", ""))
    except InvalidOperation:
        return None
    if m.group("open") or m.group("sign") or m.group("sign2"):
        v = -v
    return v


def is_amount(text):
    return parse_amount(text) is not None


def decimals_of(text):
    t = str(text).strip()
    m = _NUM.match(t)
    if not m or "." not in m.group("num"):
        return 0
    return len(m.group("num").split(".")[1])


def format_amount(value, decimals=0, negative="brackets", zero="-"):
    """Format a Decimal the standard way: 1,234 / (1,234) / -."""
    if value == 0 and zero == "-":
        return "-"
    q = abs(value)
    s = f"{q:,.{decimals}f}"
    if value < 0:
        return f"({s})" if negative == "brackets" else f"-{s}"
    return s


def looks_like_note_ref(text):
    t = str(text).strip()
    return bool(re.fullmatch(r"\d{1,2}(\.\d{1,2})?(\s*[,&]\s*\d{1,2}(\.\d{1,2})?)*[a-z]?", t))


def find_amounts(text):
    """All figures written inside running text (used by the integrity check)."""
    out = []
    for m in re.finditer(r"\(?-?\d{1,3}(?:,\d{3})+(?:\.\d+)?\)?|\(?-?\d+\.\d+\)?|\b\d+\b", text):
        v = parse_amount(m.group(0))
        if v is not None:
            out.append(abs(v))
    return out
