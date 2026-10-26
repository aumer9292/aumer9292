"""Proofreading: spelling, repeated words, spacing, dates, terminology, references
to superseded standards, and wording that is used inconsistently.

Only safe, mechanical fixes (spacing) are applied automatically. Spelling
corrections are applied only when the user turns that option on, and every
change is listed in the review report.
"""

import calendar
import re
from collections import Counter, defaultdict

from .classify import SECTIONS
from .model import Issue

# Common misspellings in financial statements -> correction.
MISSPELLINGS = {
    "accomodation": "accommodation", "accural": "accrual", "accurals": "accruals", "acrual": "accrual",
    "acruals": "accruals", "accruel": "accrual", "acumulated": "accumulated", "accumalated": "accumulated",
    "accumlated": "accumulated", "aquisition": "acquisition", "aquisitions": "acquisitions",
    "acquistion": "acquisition", "ammortisation": "amortisation", "amortisaton": "amortisation",
    "amortiztion": "amortization", "annum": "annum", "apropriate": "appropriate", "assests": "assets",
    "asets": "assets", "assocaite": "associate", "asociate": "associate", "auditer": "auditor",
    "auditers": "auditors", "beggining": "beginning", "begining": "beginning", "benifit": "benefit",
    "benifits": "benefits", "borowings": "borrowings", "borrowins": "borrowings", "buisness": "business",
    "busines": "business", "calender": "calendar", "captial": "capital", "comparitive": "comparative",
    "comparitives": "comparatives", "compliace": "compliance", "comprehensve": "comprehensive",
    "comprehesive": "comprehensive", "contigent": "contingent", "contingencey": "contingency",
    "comission": "commission", "commision": "commission", "consolidted": "consolidated", "cuurent": "current",
    "currrent": "current", "curent": "current", "debtor's": "debtors'", "decleration": "declaration",
    "defered": "deferred", "deffered": "deferred", "depreciaton": "depreciation", "deprecation": "depreciation",
    "depriciation": "depreciation", "derecognised": "derecognised", "diposal": "disposal", "dispoal": "disposal",
    "disclousre": "disclosure", "disclosre": "disclosure", "dividened": "dividend", "divident": "dividend",
    "dividents": "dividends", "eqiuty": "equity", "equty": "equity", "equtiy": "equity", "equivalants": "equivalents",
    "equivelents": "equivalents", "equivalnets": "equivalents", "excercise": "exercise", "exhange": "exchange",
    "expences": "expenses", "expence": "expense", "expensess": "expenses", "finacial": "financial",
    "financal": "financial", "finanical": "financial", "fiancial": "financial", "goverment": "government",
    "guarentee": "guarantee", "guarante": "guarantee", "impairement": "impairment", "impariment": "impairment",
    "incomee": "income", "independant": "independent", "intangable": "intangible", "intangilbe": "intangible",
    "intrest": "interest", "intreset": "interest", "inventries": "inventories", "invetories": "inventories",
    "inventorys": "inventories", "investmnet": "investment", "investmnets": "investments", "judgement": "judgement",
    "liabilites": "liabilities", "liabilties": "liabilities", "liablities": "liabilities", "liabilitis": "liabilities",
    "liabilty": "liability", "liablity": "liability", "maintainance": "maintenance", "managment": "management",
    "materiel": "material", "measurment": "measurement", "occurence": "occurrence", "occured": "occurred",
    "oppurtunity": "opportunity", "operting": "operating", "opertaing": "operating", "payables": "payables",
    "payble": "payable", "payabels": "payables", "perpetual": "perpetual", "porfit": "profit", "proft": "profit",
    "profti": "profit", "propery": "property", "propety": "property", "provison": "provision",
    "provisons": "provisions", "provisoin": "provision", "receieved": "received", "recieve": "receive",
    "recieved": "received", "recievable": "receivable", "recievables": "receivables", "receiveable": "receivable",
    "receiveables": "receivables", "recevables": "receivables", "reconcilation": "reconciliation",
    "reconcilliation": "reconciliation", "recogised": "recognised", "recognsied": "recognised",
    "recognistion": "recognition", "recongnition": "recognition", "relevent": "relevant", "renumeration": "remuneration",
    "reserves": "reserves", "reveune": "revenue", "revenu": "revenue", "revenues": "revenues", "revneue": "revenue",
    "seperate": "separate", "seperately": "separately", "shareholers": "shareholders", "sharholders": "shareholders",
    "signficant": "significant", "significent": "significant", "statment": "statement", "statments": "statements",
    "statemnet": "statement", "subsidary": "subsidiary", "subsidaries": "subsidiaries", "subsidiries": "subsidiaries",
    "taxtion": "taxation", "taxaion": "taxation", "therefor": "therefore", "transcation": "transaction",
    "transcations": "transactions", "transation": "transaction", "untill": "until", "valuaton": "valuation",
    "wich": "which", "withold": "withhold", "witholding": "withholding", "yeild": "yield", "retaind": "retained",
    "retianed": "retained", "earnigs": "earnings", "earings": "earnings", "ernings": "earnings",
    "begininng": "beginning", "acount": "account", "acounts": "accounts", "accouting": "accounting",
    "acounting": "accounting", "accountting": "accounting", "polices": "policies", "policys": "policies",
    "estimatess": "estimates", "comparitively": "comparatively", "excluding": "excluding",
    "goodwil": "goodwill", "leaes": "leases", "lese": "lease", "loanes": "loans", "borowing": "borrowing",
    "debenturs": "debentures", "ordnary": "ordinary", "auhtorised": "authorised", "authorized": "authorized",
    "issueed": "issued", "paidup": "paid-up", "perid": "period", "peroid": "period", "yaer": "year",
    "yera": "year", "endded": "ended", "eneded": "ended", "balanace": "balance", "balnce": "balance",
}
# words above that are actually correct spellings (kept to avoid flagging)
for ok in ("annum", "derecognised", "judgement", "payables", "perpetual", "reserves", "revenues", "excluding",
           "authorized"):
    MISSPELLINGS.pop(ok, None)

# Terms that should be used consistently through one set of statements.
TERM_GROUPS = [
    ("statement of financial position / balance sheet",
     [r"statement of financial position", r"balance sheet"]),
    ("statement of profit or loss / income statement / profit and loss account",
     [r"statement of (?:profit or loss|comprehensive income)", r"income statement", r"profit and loss account"]),
    ("statement of cash flows / cash flow statement", [r"statement of cash flows", r"cash flow statement"]),
    ("notes to the financial statements / notes to the accounts",
     [r"notes to the financial statements", r"notes to the accounts"]),
    ("trade receivables / trade debtors", [r"trade receivables", r"trade debtors"]),
    ("trade payables / trade creditors", [r"trade payables", r"trade creditors"]),
    ("inventories / stock-in-trade / stocks", [r"\binventor(?:y|ies)\b", r"stock[- ]in[- ]trade", r"\bstocks\b"]),
    ("property, plant and equipment / fixed assets", [r"property, plant and equipment", r"\bfixed assets\b"]),
    ("-ise / -ize spelling", [r"\b(?:recognis|amortis|capitalis|organis|realis|authoris|categoris)(?:e|ed|es|ing|ation)\b",
                              r"\b(?:recogniz|amortiz|capitaliz|organiz|realiz|authoriz|categoriz)(?:e|ed|es|ing|ation)\b"]),
    ("judgement / judgment", [r"\bjudgements?\b", r"\bjudgments?\b"]),
]

SUPERSEDED = {
    "IAS 11": "IFRS 15 (Revenue from Contracts with Customers) since 2018",
    "IAS 18": "IFRS 15 (Revenue from Contracts with Customers) since 2018",
    "IAS 17": "IFRS 16 (Leases) since 2019",
    "IAS 39": "IFRS 9 (Financial Instruments) for most entities since 2018",
    "IFRS 4": "IFRS 17 (Insurance Contracts) since 2023",
    "IAS 35": "IFRS 5",
    "SIC-15": "IFRS 16",
    "IFRIC 4": "IFRS 16",
    "IFRIC 13": "IFRS 15",
    "IFRIC 15": "IFRS 15",
    "IFRIC 18": "IFRS 15",
}

MONTHS = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}
DATE_RX = re.compile(r"\b(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|"
                     r"September|October|November|December),?\s+(\d{4})\b|\b(January|February|March|April|May|"
                     r"June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})\b", re.I)
WORD_RX = re.compile(r"[A-Za-z][A-Za-z'’-]*")


def iter_texts(doc):
    """(text, location, section, setter) for every paragraph and table cell."""
    for b in doc.blocks:
        if b.kind == "paragraph":
            yield b.text, b.source, b.section, b
        else:
            for ri, r in enumerate(b.rows):
                for ci, c in enumerate(r.cells):
                    if c:
                        yield c, f"{b.source}, row {ri + 1}", b.section, None


def _sec(s):
    return SECTIONS.get(s, s)


def proofread(doc, settings):
    issues = []
    spelled = Counter()
    where = defaultdict(list)
    extra_words = {w.lower() for w in settings.get("custom_words", [])}
    spell = _optional_spellchecker()
    unknown = Counter()
    unknown_where = {}

    for text, loc, sec, _ in iter_texts(doc):
        if not text.strip():
            continue
        for m in WORD_RX.finditer(text):
            w = m.group(0).strip("'’-")
            lw = w.lower()
            if lw in MISSPELLINGS and lw not in extra_words:
                spelled[(w, MISSPELLINGS[lw])] += 1
                where[(w, MISSPELLINGS[lw])].append(loc)
            elif spell is not None and lw not in extra_words and len(lw) > 3 and w.islower() \
                    and lw not in spell and lw not in FINANCE_WORDS:
                unknown[lw] += 1
                unknown_where.setdefault(lw, loc)
        for m in re.finditer(r"\b(\w+)\s+\1\b", text, re.I):
            if m.group(1).lower() not in ("had", "that", "is") and not m.group(1).isdigit():
                issues.append(Issue("proofreading", "warning", f"Repeated word: '{m.group(0)}'", loc, _sec(sec),
                                    suggestion=f"Delete the second '{m.group(1)}'."))
        if re.search(r"\s[,.;:](?!\d)", text) and not re.search(r"\s\.\d", text):
            issues.append(Issue("proofreading", "info", "Space before punctuation: “" + _snip(text, r"\s[,.;:]") + "”",
                                loc, _sec(sec)))
        if re.search(r"[a-z][,;][A-Za-z]", text):
            issues.append(Issue("proofreading", "info", "Missing space after punctuation: “"
                                + _snip(text, r"[a-z][,;][A-Za-z]") + "”", loc, _sec(sec)))
        if text.count("(") != text.count(")"):
            issues.append(Issue("proofreading", "warning", "Unmatched brackets: “" + text[:80] + "”", loc, _sec(sec)))
        for m in DATE_RX.finditer(text):
            if m.group(1):
                d, mon, y = int(m.group(1)), m.group(2), int(m.group(3))
            else:
                mon, d, y = m.group(4), int(m.group(5)), int(m.group(6))
            last = calendar.monthrange(y, MONTHS[mon.lower()])[1]
            if d > last:
                issues.append(Issue("proofreading", "error", f"Invalid date: '{m.group(0)}' ({mon} has {last} days "
                                    f"in {y}).", loc, _sec(sec), suggestion=f"Did you mean {last} {mon.title()} {y}?"))
        for std, repl in SUPERSEDED.items():
            if re.search(r"\b" + re.escape(std).replace(r"\ ", r"\s*") + r"\b(?!\d)", text):
                if settings.get("framework", "IFRS").upper().startswith(("IFRS", "IAS")):
                    issues.append(Issue("proofreading", "warning", f"Reference to {std}, which has been "
                                        f"superseded by {repl}.", loc, _sec(sec),
                                        suggestion="Update the reference unless it is deliberate (e.g. transition "
                                                   "disclosures)."))

    for (w, fix), n in sorted(spelled.items()):
        locs = where[(w, fix)]
        issues.append(Issue("proofreading", "error", f"Spelling: '{w}' → '{_match_case(w, fix)}' "
                            f"({n} occurrence{'s' if n > 1 else ''})", ", ".join(dict.fromkeys(locs[:3])),
                            suggestion=f"Replace with '{_match_case(w, fix)}'."))
    for w, n in unknown.most_common(40):
        issues.append(Issue("proofreading", "info", f"Possible spelling error: '{w}' ({n}×)", unknown_where[w],
                            suggestion="Check the spelling, or add the word to the custom dictionary."))
    issues += _consistency(doc)
    issues += _years(doc, settings)
    issues += _period_dates(doc)
    return issues


FINANCE_WORDS = set("""
ifrs ias ifric sic gaap smes sme frs aspe ipsas ppe fvoci fvtpl fvtoci ecl oci eps ebitda ebit cgu cgus
receivables payables accruals prepayments impairment derecognition derecognised derecognized amortised amortized
amortisation amortization capitalised capitalized remeasurement remeasurements onerous lessee lessor lessees lessors
subleases sublease debentures debenture sukuk musharaka murabaha ijarah zakat goodwill intangibles
reclassified reclassification reclassifications hyperinflationary undiscounted discounted unrecognised
unrecognized unutilised unutilized unappropriated retained transactional fiduciary interco intercompany
investee investees associates joint ventures overdrafts overdraft repo repos nominee nominees unquoted quoted
counterparties counterparty creditworthiness rateably straightline appropriations dividends dividend
subsidiaries subsidiary consolidated deconsolidated deconsolidation judgements judgement judgments judgment
""".split())


def _optional_spellchecker():
    try:
        from spellchecker import SpellChecker
        return SpellChecker()
    except Exception:
        return None


def _match_case(src, fix):
    if src.isupper():
        return fix.upper()
    if src[:1].isupper():
        return fix[:1].upper() + fix[1:]
    return fix


def _snip(text, rx):
    m = re.search(rx, text)
    if not m:
        return text[:60]
    a, b = max(0, m.start() - 25), min(len(text), m.end() + 25)
    return ("…" if a else "") + text[a:b] + ("…" if b < len(text) else "")


def _consistency(doc):
    issues = []
    full = doc.text()
    for name, pats in TERM_GROUPS:
        used = []
        for p in pats:
            n = len(re.findall(p, full, re.I))
            if n:
                used.append((p, n))
        if len(used) > 1:
            parts = []
            for p, n in used:
                m = re.search(p, full, re.I)
                parts.append(f"'{m.group(0)}' ×{n}")
            issues.append(Issue("consistency", "warning", f"Mixed terminology ({name}): " + ", ".join(parts),
                                suggestion="Use one term consistently throughout the financial statements."))
    # The same title written with different capitalisation
    variants = defaultdict(Counter)
    for p in doc.paragraphs():
        if p.role in ("title", "heading", "note_heading") and p.text.strip():
            key = re.sub(r"^\s*\d+(\.\d+)*\.?\s*", "", p.text.strip())
            variants[key.lower()][key] += 1
    for low, c in variants.items():
        if len(c) > 1:
            issues.append(Issue("consistency", "info", "Same heading capitalised differently: "
                                + " / ".join(f"'{k}'" for k in c),
                                suggestion="The formatter applies one capitalisation style to headings."))
    return issues


def _years(doc, settings):
    issues = []
    try:
        fy = int(settings.get("financial_year") or 0)
    except ValueError:
        fy = 0
    if not fy:
        return issues
    for p in doc.paragraphs():
        if p.role in ("title", "period") or (p.section in ("SOFP", "SOCI", "SOCE", "SOCF") and p.role != "body"):
            for y in re.findall(r"\b(19\d{2}|20\d{2})\b", p.text):
                y = int(y)
                if y > fy + 1:
                    issues.append(Issue("consistency", "error", f"Year {y} in “{p.text.strip()[:90]}” is after the "
                                        f"financial year {fy}.", p.source, _sec(p.section),
                                        suggestion="Check the reporting period."))
                elif p.role == "period" and y not in (fy, fy - 1):
                    issues.append(Issue("consistency", "warning", f"Period line mentions {y}; the financial year "
                                        f"selected is {fy}.", p.source, _sec(p.section)))
    return issues


def _period_dates(doc):
    """Statements in one set should share one period-end date."""
    issues = []
    dates = Counter()
    where = {}
    for p in doc.paragraphs():
        if p.role == "period":
            m = DATE_RX.search(p.text)
            if m:
                d = re.sub(r"\s+", " ", m.group(0).replace(",", ""))
                d = " ".join(w.capitalize() if w.isalpha() else w for w in d.split())
                dates[d] += 1
                where.setdefault(d, p.source)
    if len(dates) > 1:
        years = defaultdict(list)
        for d in dates:
            years[d[-4:]].append(d)
        top_year = max(years, key=lambda y: sum(dates[d] for d in years[y]))
        variants = years[top_year]
        if len(variants) > 1:
            issues.append(Issue("consistency", "error", "Different period-end dates are used in statement headings: "
                                + ", ".join(f"'{d}' ({where[d]})" for d in variants),
                                suggestion="Use the same reporting date throughout."))
    return issues


def safe_text_fixes(text):
    """Mechanical fixes that never change meaning: collapse runs of spaces,
    trim spaces before punctuation, and normalise space around brackets."""
    new = re.sub(r"(?<=\S) {2,}(?=\S)", " ", text)
    new = re.sub(r"(?<=[A-Za-z)]) +([,;:])(?=\s|$)", r"\1", new)
    new = re.sub(r"(?<=[A-Za-z)]) +\.(?=\s|$)", ".", new)
    new = re.sub(r"\( +", "(", new)
    new = re.sub(r" +\)", ")", new)
    return new


def spelling_fix(text):
    def rep(m):
        w = m.group(0)
        fix = MISSPELLINGS.get(w.lower())
        return _match_case(w, fix) if fix else w
    return WORD_RX.sub(rep, text)
