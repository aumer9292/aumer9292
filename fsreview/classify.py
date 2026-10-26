"""Work out what each part of the document is.

Sections: the primary statements, notes, accounting policies, estimates,
directors' and auditor's reports and supplementary schedules. Paragraph roles:
titles, headings, note headings, period lines, unit lines and body text.
Table structure: label, note and amount columns, the year above each amount
column, and each row's kind (heading, item, subtotal, total, grand total).
"""

import re

from .numbers import YEAR, is_amount, looks_like_note_ref, parse_amount

SECTIONS = {
    "SOFP": "Statement of Financial Position",
    "SOCI": "Statement of Profit or Loss and Other Comprehensive Income",
    "SOCE": "Statement of Changes in Equity",
    "SOCF": "Statement of Cash Flows",
    "NOTES": "Notes to the Financial Statements",
    "POLICIES": "Significant Accounting Policies",
    "ESTIMATES": "Accounting Estimates and Judgements",
    "DIRECTORS": "Directors' / Management Report",
    "AUDITOR": "Auditor's Report",
    "SUPPLEMENTARY": "Supplementary Information",
    "CONTENTS": "Contents",
    "COVER": "Cover and General Information",
}
PRIMARY = ("SOFP", "SOCI", "SOCE", "SOCF")

_TITLE_PATTERNS = [
    ("SOFP", r"statement of financial position|balance sheet|statement of financial condition|^sofp$|^bs$"),
    ("SOCI", r"statement of (?:profit or loss|comprehensive income|income|profit and loss|operations)"
             r"|income statement|profit and loss account|statement of financial performance|^soci$|^p ?& ?l$|^pl$"),
    ("SOCE", r"statement of changes in (?:equity|net assets)|statement of changes in shareholders|^soce$"),
    ("SOCF", r"statement of cash ?flows?|cash ?flow statement|^socf$|^cfs?$"),
    ("POLICIES", r"(?:significant|material|summary of|principal)\s+(?:\w+\s+)?accounting policies"
                 r"|^accounting policies$|basis of preparation"),
    ("ESTIMATES", r"(?:critical|significant)?\s*accounting estimates|judg(?:e)?ments? and estimates"
                  r"|estimates and judg(?:e)?ments|sources of estimation uncertainty"),
    ("NOTES", r"notes to (?:and forming part of )?the (?:financial statements|accounts)|^notes$"),
    ("DIRECTORS", r"directors'? report|report of the directors|directors’ report|management report"
                  r"|management'?s? discussion|chairman'?s? (?:review|statement)|board of directors' report"),
    ("AUDITOR", r"independent auditor'?s'? report|auditors?'? report|auditor’s report|report of the auditors?"
                r"|review report"),
    ("SUPPLEMENTARY", r"supplementary (?:information|schedules?)|detailed (?:income|trading|profit)"
                      r"|schedule of|additional information"),
    ("CONTENTS", r"^(?:table of )?contents$|^index$"),
]
_TITLE_RES = [(k, re.compile(p, re.I)) for k, p in _TITLE_PATTERNS]

NOTE_HEADING = re.compile(r"^\s*(?:note\s+)?(\d{1,2}(?:\.\d{1,2})?)\.?\s+([A-Z(][^\n]{1,110})$")
PERIOD_LINE = re.compile(
    r"\b(for the (?:year|period|six months|three months|quarter)|as at|as of|at)\b.{0,40}"
    r"\b(\d{1,2}(?:st|nd|rd|th)?\s+[A-Z][a-z]+,?\s+\d{4}|[A-Z][a-z]+\s+\d{1,2},?\s+\d{4})", re.I)
UNIT_LINE = re.compile(r"^\(?\s*(?:all\s+)?(?:amounts?|figures)?\s*(?:are\s+)?(?:in|expressed in|presented in)?\s*"
                       r"(?:[A-Z]{3}|Rs\.?|Rupees|\$|€|£|[A-Z][a-z]+ (?:dollars|rupees|pounds|euros))"
                       r"\s*'?(?:000|thousands?|millions?|m|mn|bn)?'?\s*\)?$", re.I)

GRAND_TOTAL = re.compile(
    r"^total (?:assets|equity and liabilities|liabilities and equity|equity|liabilities)$"
    r"|^net assets$|^(?:profit|loss|\(loss\)|profit/\(loss\)|profit / \(loss\)|net profit|net loss|net income)"
    r"(?: after tax(?:ation)?)? for the (?:year|period)"
    r"|^total comprehensive (?:income|loss|income/\(loss\))"
    r"|^cash and cash equivalents at (?:the )?end|^closing cash|^balance (?:as )?at.*\d{4}$", re.I)
TOTAL = re.compile(r"^(?:total|net |sub-?total|gross (?:profit|loss)|operating (?:profit|loss)"
                   r"|(?:profit|loss|\(loss\)|profit/\(loss\)).*before|profit from operations|"
                   r"net (?:cash|increase|decrease|\(decrease\)|change)|cash generated)", re.I)


def classify_title(text):
    t = text.strip()
    if not t or len(t) > 140:
        return None
    if re.search(r"(\.{3,}|…|\t|\s{3,})\s*\d{1,3}(\s*-\s*\d{1,3})?$", t):
        return "CONTENTS_ENTRY"
    m = re.match(r"^(.*?[A-Za-z)])\s*\d{1,3}(\s*-\s*\d{1,3})?$", t)
    if m and not re.search(r"\b(19|20)\d\d$", t) and not NOTE_HEADING.match(t):
        inner = m.group(1)
        if any(rx.search(inner) for _, rx in _TITLE_RES) and len(inner.split()) <= 10:
            return "CONTENTS_ENTRY"
    for key, rx in _TITLE_RES:
        if rx.search(t):
            if key in PRIMARY or key in ("CONTENTS",):
                # body sentences mention statements too; a title is short and has no full stop
                if len(t) > 110 or t.endswith(".") and len(t.split()) > 12:
                    return None
            elif len(t.split()) > 14:
                return None
            return key
    return None


def classify_document(doc, settings=None):
    section = "COVER"
    in_contents = False
    current_note = None
    tables_since_title = 0
    for b in doc.blocks:
        if b.kind == "paragraph":
            t = b.text.strip()
            if not t:
                b.role = "blank"
                b.section = section
                continue
            key = classify_title(t)
            if key == "CONTENTS_ENTRY":
                b.role = "contents"
                b.section = "CONTENTS"
                continue
            if key:
                if key == "CONTENTS":
                    in_contents = True
                else:
                    in_contents = False
                if key in PRIMARY or key in ("NOTES", "DIRECTORS", "AUDITOR", "SUPPLEMENTARY", "CONTENTS"):
                    section = key
                    b.role = "title"
                    current_note = None
                elif key in ("POLICIES", "ESTIMATES"):
                    section = key
                    b.role = "heading"
                    m = NOTE_HEADING.match(t)
                    if m:
                        b.role = "note_heading"
                        b.note_number = m.group(1).split(".")[0]
                        current_note = b.note_number
                b.section = section
                tables_since_title = 0
                continue
            if in_contents:
                b.role = "contents"
                b.section = "CONTENTS"
                if len(t) > 160:
                    in_contents = False
                else:
                    continue
            if PERIOD_LINE.search(t) and len(t) < 120:
                b.role = "period"
            elif UNIT_LINE.match(t) and len(t) < 60:
                b.role = "unit"
            else:
                m = NOTE_HEADING.match(t)
                if m and section in ("NOTES", "POLICIES", "ESTIMATES", "SUPPLEMENTARY") or \
                        (m and section in PRIMARY and tables_since_title > 0):
                    if len(t) < 110 and not t.endswith(".") and not re.search(r"\d,\d{3}", t):
                        b.role = "note_heading"
                        num = m.group(1)
                        b.note_number = num
                        if "." not in num:
                            if section in PRIMARY or (section in ("POLICIES", "ESTIMATES")
                                                      and current_note and num != current_note):
                                section = "NOTES"
                            current_note = num
                        b.section = section
                        continue
                if _looks_heading(b):
                    b.role = "heading"
                else:
                    b.role = "body"
            b.section = section
            b.note_number = b.note_number or current_note
        else:
            in_contents = False
            b.section = section
            tables_since_title += 1
            analyse_table(b)
            if current_note and section in ("NOTES", "POLICIES", "ESTIMATES"):
                b.note_number = current_note
    return doc


def _looks_heading(p):
    t = p.text.strip()
    if len(t) > 100 or t.endswith((".", ":", ";", ",")) and len(t.split()) > 6:
        return False
    if len(t.split()) > 14:
        return False
    if p.bold or (t.isupper() and len(t) > 3) or (p.style or "").lower().startswith(("heading", "title")):
        return True
    return False


# ---------------------------------------------------------------- tables

def analyse_table(t):
    w = t.width
    if w == 0:
        return t
    for r in t.rows:
        if len(r.cells) < w:
            r.cells = r.cells + [""] * (w - len(r.cells))
    # header rows: before the first row that has a figure in a non-first column
    stats = []
    for c in range(w):
        vals = [r.cells[c] for r in t.rows if r.cells[c].strip()]
        nums = [v for v in vals if is_amount(v) and not YEAR.match(v.strip())]
        stats.append((len(vals), len(nums)))
    note_col = None
    for c in range(1, w):
        head = " ".join(r.cells[c] for r in t.rows[:4]).strip().lower()
        if re.fullmatch(r"notes?|note\s*no\.?|ref\.?", head):
            note_col = c
            break
    if note_col is None:
        for c in range(1, w - 1):
            vals = [r.cells[c].strip() for r in t.rows if r.cells[c].strip()]
            body = [v for v in vals if not YEAR.match(v) and not v.lower().startswith("note")]
            if body and all(looks_like_note_ref(v) for v in body) and \
                    all(parse_amount(v) is not None and abs(parse_amount(v)) < 100 for v in body
                        if parse_amount(v) is not None) and \
                    sum(1 for c2 in range(c + 1, w) if stats[c2][1] > 0) >= 1 and len(body) < len(t.rows):
                note_col = c
                break
    amount_cols = [c for c in range(1, w) if c != note_col and stats[c][1] >= 1
                   and stats[c][1] >= 0.5 * stats[c][0]]
    label_col = 0
    t.note_col, t.amount_cols, t.label_col = note_col, amount_cols, label_col

    # header rows and years per column
    first_data = len(t.rows)
    for i, r in enumerate(t.rows):
        if any(is_amount(r.cells[c]) and not YEAR.match(r.cells[c].strip()) for c in amount_cols):
            first_data = i
            break
    for i, r in enumerate(t.rows[:first_data]):
        for c in amount_cols:
            m = re.search(r"\b((?:19|20)\d{2})\b", r.cells[c])
            if m:
                t.column_years[c] = int(m.group(1))
            if re.search(r"['’]000|(?i:thousands?|millions?)|[$€£₹]|\bRs\.?|^[A-Z]{3}\b", r.cells[c]) \
                    and not t.unit_label:
                if not YEAR.match(r.cells[c].strip()):
                    t.unit_label = r.cells[c].strip()
    for i, r in enumerate(t.rows):
        label = r.cells[label_col].strip()
        r.label = label
        has_amt = any(is_amount(r.cells[c]) and not (i < first_data and YEAR.match(r.cells[c].strip()))
                      for c in amount_cols)
        if i < first_data and not has_amt:
            r.kind = "header" if any(r.cells[c].strip() for c in range(1, w)) or i < first_data else "heading"
            if not any(r.cells[c].strip() for c in range(1, w)) and label and i > 0:
                r.kind = "heading"
            continue
        if not any(x.strip() for x in r.cells):
            r.kind = "blank"
        elif not has_amt:
            r.kind = "heading"
        elif t.section == "SOCE":
            # movements in equity are line items; balances are the totals
            if re.match(r"^(balance|as at|at )", label, re.I):
                r.kind = "total"
            elif re.match(r"^total", label, re.I):
                r.kind = "subtotal"
            else:
                r.kind = "item"
        elif GRAND_TOTAL.search(label):
            r.kind = "grand_total"
        elif TOTAL.search(label) or (not label and has_amt):
            r.kind = "total"
        elif r.bold and has_amt and label:
            r.kind = "subtotal"
        else:
            r.kind = "item"
    if t.section == "SOCE":
        balances = [r for r in t.rows if r.kind == "total"]
        if balances:
            balances[-1].kind = "grand_total"
    return t


def row_values(t, r, cols=None):
    cols = t.amount_cols if cols is None else cols
    return [parse_amount(r.cells[c]) if c < len(r.cells) else None for c in cols]


def is_financial_table(t):
    data = [r for r in t.rows if r.kind in ("item", "total", "grand_total", "subtotal")]
    return bool(t.amount_cols) and len(data) >= 2
