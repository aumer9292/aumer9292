"""Are the relevant notes and disclosures included?

A built-in checklist per entity category (PIC, LSC, MSC, SSC). Each item says
which categories need it, when it applies (always, or only when a related line
appears in the statements), and the wording that shows it is there. Results
are Present, Missing or Not applicable, with the reference.

The checklist is a reviewer's aid built from IFRS, IFRS for SMEs, the Revised
AFRS for SSEs and the Companies Act 2017 schedules. It looks for the presence of
a disclosure, not its adequacy: a 'Present' item still needs reading.
"""

import re
from decimal import Decimal

from .classify import PRIMARY, SECTIONS
from .model import Issue
from .numbers import parse_amount

ALL = ("PIC", "LSC", "MSC", "SSC")
FULL = ("PIC", "LSC")
NOT_SMALL = ("PIC", "LSC", "MSC")


def _sofp(pattern):
    return lambda c: bool(re.search(pattern, c["sofp_labels"], re.I))


def _soci(pattern):
    return lambda c: bool(re.search(pattern, c["soci_labels"], re.I))


def _any(pattern):
    return lambda c: bool(re.search(pattern, c["statement_labels"] + "\n" + c["text"], re.I))


ALWAYS = None

# id, title, reference, categories, applies-when, patterns that show it is there
CHECKLIST = [
    # ---- general information and basis of preparation
    ("legal_status", "Legal status and nature of business", "IAS 1.138 / Companies Act 2017", ALL, ALWAYS,
     [r"legal status", r"(?:the )?company and its operations", r"status and nature of business", r"nature of (?:the )?business",
      r"reporting entity", r"general information", r"incorporated (?:in|under|on)"]),
    ("registered_office", "Registered office and places of business", "Companies Act 2017 schedules", NOT_SMALL, ALWAYS,
     [r"registered office", r"geographical location", r"place[s]? of business"]),
    ("compliance", "Statement of compliance with the applicable reporting framework", "IAS 1.16 / Companies Act 2017 s.225",
     ALL, ALWAYS, [r"statement of compliance", r"prepared in accordance with", r"basis of preparation"]),
    ("measurement_basis", "Basis of measurement", "IAS 1.117", NOT_SMALL, ALWAYS,
     [r"basis of measurement", r"historical cost convention", r"historical cost basis"]),
    ("currency", "Functional and presentation currency", "IAS 21.53 / IAS 1.51", NOT_SMALL, ALWAYS,
     [r"functional (?:and|&) presentation currency", r"presentation currency", r"functional currency"]),
    ("estimates", "Critical accounting estimates and judgements", "IAS 1.122, 1.125", NOT_SMALL, ALWAYS,
     [r"estimates and judg", r"judg(?:e)?ments and estimates", r"accounting estimates", r"use of estimates",
      r"sources of estimation uncertainty"]),
    ("new_standards", "New standards, amendments and interpretations (effective and not yet effective)", "IAS 8.30",
     FULL, ALWAYS, [r"standards?,? (?:interpretations|amendments)", r"not yet effective", r"new (?:and amended )?standards",
                    r"amendments to (?:published|approved) accounting standards", r"initial application of"]),
    ("policies", "Material / significant accounting policies", "IAS 1.117", ALL, ALWAYS,
     [r"accounting policies", r"summary of significant"]),
    # ---- statement of financial position lines
    ("ppe", "Property, plant and equipment: reconciliation, depreciation method and rates", "IAS 16.73 / Section 17 SMEs",
     ALL, _sofp(r"property, plant|operating fixed assets|fixed assets"),
     [r"property, plant and equipment", r"operating fixed assets", r"fixed assets"]),
    ("cwip", "Capital work-in-progress", "IAS 16.74(b)", NOT_SMALL, _any(r"capital work[- ]in[- ]progress"),
     [r"capital work[- ]in[- ]progress"]),
    ("revaluation", "Revaluation: date, valuer, carrying amount under the cost model", "IAS 16.77",
     ALL, _any(r"surplus on revaluation|revaluation surplus|revalued"),
     [r"revaluation", r"revalued"]),
    ("rou", "Right-of-use assets and lease liabilities", "IFRS 16.53, 16.58", FULL, _any(r"right[- ]of[- ]use|lease liabilit"),
     [r"right[- ]of[- ]use", r"lease liabilit"]),
    ("leases_sme", "Leases (finance and operating lease commitments)", "Section 20 SMEs / AFRS SSE", ("MSC", "SSC"),
     _any(r"\blease"), [r"\blease"]),
    ("intangibles", "Intangible assets: reconciliation and amortisation", "IAS 38.118", ALL, _sofp(r"intangible"),
     [r"intangible"]),
    ("investments", "Investments: classification and measurement", "IFRS 7 / IFRS 9", ALL,
     _sofp(r"investment|associate|subsidiar"), [r"investments?", r"associate", r"subsidiar"]),
    ("loans_adv", "Loans and advances", "Companies Act 2017 schedules", ALL, _sofp(r"loans and advances|advances"),
     [r"loans and advances", r"advances"]),
    ("deposits", "Deposits, prepayments and other receivables", "IAS 1.77", ALL,
     _sofp(r"deposits|prepayments|other receivables"), [r"deposits", r"prepayments", r"other receivables"]),
    ("stores", "Stores, spares and loose tools", "IAS 2.36", ALL, _sofp(r"stores"), [r"stores"]),
    ("inventory", "Stock-in-trade / inventories: categories and cost formula", "IAS 2.36", ALL,
     _sofp(r"inventor|stock[- ]in[- ]trade|stocks"), [r"inventor", r"stock[- ]in[- ]trade", r"stocks"]),
    ("receivables", "Trade debts / receivables, ageing and loss allowance", "IFRS 7.35H / IFRS 9", ALL,
     _sofp(r"trade debts|receivable|debtors"), [r"trade debts", r"receivable", r"debtors"]),
    ("ecl", "Expected credit loss / impairment allowance on receivables", "IFRS 7.35H", FULL,
     _sofp(r"trade debts|receivable|debtors"),
     [r"expected credit loss", r"\becl\b", r"loss allowance", r"impairment allowance", r"provision for (?:doubtful|impairment)"]),
    ("cash", "Cash and bank balances", "IAS 7.45", ALL, _sofp(r"cash"), [r"cash and (?:bank|cash equivalents)", r"bank balances",
                                                                       r"cash at bank"]),
    ("share_capital", "Share capital: authorised, issued, subscribed and paid-up; shares held by associates",
     "IAS 1.79 / Companies Act 2017", ALL, _sofp(r"share capital|paid[- ]up"),
     [r"share capital", r"paid[- ]up", r"authori[sz]ed capital"]),
    ("reserves", "Reserves: nature and purpose", "IAS 1.79(b)", NOT_SMALL, _sofp(r"reserve"), [r"reserve"]),
    ("lt_borrowings", "Long term financing: lender, security, mark-up rate and repayment terms", "IFRS 7.7 / Companies Act 2017",
     ALL, _sofp(r"long[- ]term (?:financ|loan|borrow)|borrowings|loans? from"),
     [r"long[- ]term (?:financ|loan|borrow)", r"borrowings", r"loans? from"]),
    ("st_borrowings", "Short term borrowings: limits, security and mark-up", "IFRS 7.7", ALL,
     _sofp(r"short[- ]term (?:borrow|financ)|running finance|overdraft"),
     [r"short[- ]term (?:borrow|financ)", r"running finance", r"overdraft"]),
    ("deferred_tax", "Deferred taxation: components and movement", "IAS 12.81(g)", NOT_SMALL, _sofp(r"deferred tax"),
     [r"deferred tax"]),
    ("employee_benefits", "Staff retirement benefits (defined benefit: actuarial assumptions, movement)", "IAS 19.135",
     ALL, _any(r"gratuity|defined benefit|staff retirement|pension"),
     [r"gratuity", r"defined benefit", r"staff retirement", r"employee benefits", r"pension"]),
    ("payables", "Trade and other payables", "IAS 1.77", ALL, _sofp(r"payables|creditors|accrued"),
     [r"payables", r"creditors"]),
    ("unclaimed_div", "Unclaimed dividend", "Companies Act 2017 s.244", ("PIC",), _sofp(r"unclaimed dividend"),
     [r"unclaimed dividend"]),
    ("contingencies", "Contingencies and commitments", "IAS 37.86 / IAS 16.74(c)", ALL, ALWAYS,
     [r"contingencies", r"commitments", r"contingent liabilit"]),
    # ---- profit or loss lines
    ("revenue", "Revenue: disaggregation and recognition", "IFRS 15.114 / Section 23 SMEs", ALL, _soci(r"revenue|sales|turnover"),
     [r"revenue", r"\bsales\b", r"turnover"]),
    ("cost_of_sales", "Cost of sales", "IAS 1.103", ALL, _soci(r"cost of (?:sales|revenue|goods|services)"),
     [r"cost of (?:sales|revenue|goods|services)"]),
    ("distribution", "Distribution / selling costs", "IAS 1.104", NOT_SMALL, _soci(r"distribution|selling|marketing"),
     [r"distribution", r"selling", r"marketing"]),
    ("admin", "Administrative expenses", "IAS 1.104", NOT_SMALL, _soci(r"administrative"), [r"administrative"]),
    ("other_income", "Other income", "IAS 1.97", ALL, _soci(r"other (?:operating )?income"), [r"other (?:operating )?income"]),
    ("other_expenses", "Other expenses (including WPPF and WWF)", "Companies Act 2017 schedules", NOT_SMALL,
     _soci(r"other (?:operating )?(?:expenses|charges)"), [r"other (?:operating )?(?:expenses|charges)", r"workers'? (?:profit|welfare)"]),
    ("finance_cost", "Finance cost", "IFRS 7.20(b)", ALL, _soci(r"finance (?:cost|charges)|financial charges|interest"),
     [r"finance (?:cost|charges)", r"financial charges", r"mark[- ]up"]),
    ("tax", "Taxation: current, prior and deferred", "IAS 12.79", ALL, _soci(r"tax"), [r"taxation", r"income tax", r"tax expense"]),
    ("tax_recon", "Relationship between tax expense and accounting profit", "IAS 12.81(c)", FULL, _soci(r"tax"),
     [r"relationship between", r"reconciliation .{0,40}tax", r"tax .{0,30}reconciliation", r"applicable tax rate",
      r"effective (?:tax )?rate"]),
    ("auditors_remuneration", "Auditors' remuneration", "Companies Act 2017 schedules", NOT_SMALL, ALWAYS,
     [r"auditors'? remuneration", r"auditor'?s'? remuneration", r"audit fee"]),
    ("donations", "Donations (names of donees in which a director is interested)", "Companies Act 2017 schedules",
     NOT_SMALL, _any(r"donation"), [r"donation"]),
    ("eps", "Earnings per share (basic and diluted)", "IAS 33.66", ("PIC",), lambda c: c["listed"],
     [r"earnings per share", r"\beps\b"]),
    ("segments", "Operating segments", "IFRS 8", ("PIC",), lambda c: c["listed"], [r"segment"]),
    # ---- cash flows
    ("cash_generated", "Cash generated from operations (reconciliation)", "IAS 7.18", NOT_SMALL,
     lambda c: bool(re.search(r"cash generated from operations", c["referenced_labels"], re.I)), [r"cash generated from operations", r"adjustments for"]),
    ("financing_changes", "Changes in liabilities arising from financing activities", "IAS 7.44A", FULL,
     _sofp(r"borrow|financing|loan|lease liabilit"),
     [r"liabilities arising from financing", r"changes in liabilities", r"reconciliation of (?:movement|liabilities)"]),
    # ---- other notes
    ("related_parties", "Transactions and balances with related parties", "IAS 24.18 / Section 33 SMEs", ALL, ALWAYS,
     [r"related part"]),
    ("remuneration", "Remuneration of chief executive, directors and executives", "Companies Act 2017 schedules / IAS 24.17",
     NOT_SMALL, ALWAYS, [r"remuneration of (?:the )?chief executive", r"remuneration of directors", r"chief executive.{0,40}directors.{0,40}executives",
                         r"key management personnel"]),
    ("employees", "Number of employees (at year end and average)", "Companies Act 2017 schedules", NOT_SMALL, ALWAYS,
     [r"number of employees", r"number of persons employed"]),
    ("provident_fund", "Provident fund (investments made in accordance with section 218)", "Companies Act 2017 s.218",
     NOT_SMALL, _any(r"provident fund"), [r"provident fund"]),
    ("capacity", "Plant capacity and actual production (reasons for shortfall)", "Companies Act 2017 schedules", FULL,
     _any(r"manufactur"), [r"capacity", r"actual production"]),
    ("fin_instruments", "Financial instruments by category", "IFRS 7.8", FULL, ALWAYS,
     [r"financial instruments by category", r"financial assets (?:and liabilities )?by category",
      r"financial (?:assets|instruments) at amortised cost"]),
    ("fin_risk", "Financial risk management: credit, liquidity and market risk", "IFRS 7.31-42", FULL, ALWAYS,
     [r"financial risk management", r"credit risk", r"liquidity risk", r"market risk"]),
    ("fair_value", "Fair value of financial instruments / fair value hierarchy", "IFRS 13.93", FULL, ALWAYS,
     [r"fair value (?:hierarchy|measurement|of financial)", r"fair value estimation"]),
    ("capital_risk", "Capital risk management", "IAS 1.134", FULL, ALWAYS, [r"capital (?:risk )?management"]),
    ("shariah", "Shariah-compliant disclosures (listed companies)", "Companies Act 2017 Fourth Schedule", ("PIC",),
     lambda c: c["listed"], [r"shariah", r"islamic mode", r"conventional and shariah"]),
    ("events_after", "Events after the reporting period", "IAS 10.21", NOT_SMALL, ALWAYS,
     [r"events? after the (?:reporting|balance sheet|statement of financial position)", r"non[- ]adjusting event",
      r"subsequent events?", r"proposed (?:final )?dividend"]),
    ("dividends", "Dividends declared / proposed", "IAS 1.137 / IAS 10.13", ALL,
     _any(r"dividends? (?:paid|declared|proposed)"), [r"dividend"]),
    ("corresponding", "Corresponding figures (re-arrangement / reclassification)", "IAS 1.41", NOT_SMALL, ALWAYS,
     [r"corresponding figures", r"comparative (?:figures|information)", r"re-?arranged", r"reclassified"]),
    ("authorisation", "Date of authorisation for issue", "IAS 10.17", ALL, ALWAYS,
     [r"authori[sz]ed for issue", r"date of authori[sz]ation", r"approved .{0,40}board of directors on"]),
    ("rounding", "General: rounding off of figures", "IAS 1.51(e)", ALL, ALWAYS,
     [r"rounded (?:off )?to the nearest", r"rounding"]),
]


def _context(doc, category, settings):
    labels = {s: [] for s in PRIMARY}
    for t in doc.tables():
        if t.section in labels:
            labels[t.section] += [r.label for r in t.rows if r.label]
    referenced = [r.label for t in doc.tables() if t.section in PRIMARY and t.note_col is not None
                  for r in t.rows if r.label and r.cells[t.note_col].strip()]
    notes_text, headings = [], []
    for b in doc.blocks:
        if b.section in ("NOTES", "POLICIES", "ESTIMATES", "SUPPLEMENTARY"):
            if b.kind == "paragraph":
                notes_text.append(b.text)
                if b.role in ("note_heading", "heading", "title"):
                    headings.append(b.text)
            else:
                for r in b.rows:
                    notes_text.append(" ".join(r.cells))
    text = doc.text()
    return {
        "category": category,
        "sofp_labels": "\n".join(labels["SOFP"]),
        "soci_labels": "\n".join(labels["SOCI"]),
        "statement_labels": "\n".join(sum(labels.values(), [])),
        "referenced_labels": "\n".join(referenced),
        "text": text,
        "notes_text": "\n".join(notes_text),
        "headings": "\n".join(headings),
        "sections": {b.section for b in doc.blocks},
        "listed": bool(re.search(r"listed on (?:the )?(?:pakistan|karachi|lahore|islamabad) stock exchange|"
                                 r"shares (?:are|of the company are) (?:listed|quoted)|\blisted company\b", text, re.I))
        or bool((settings or {}).get("listed")),
    }


def run_checklist(doc, category, settings=None):
    """Return (rows, issues). rows: dicts with id, title, ref, status, evidence."""
    ctx = _context(doc, category, settings)
    rows, issues = [], []
    cat_name = category
    for item_id, title, ref, cats, when, patterns in CHECKLIST:
        if category not in cats:
            continue
        applies = when is None or when(ctx)
        if not applies:
            rows.append({"id": item_id, "title": title, "ref": ref, "status": "n/a",
                         "evidence": "No related line or wording in the statements."})
            continue
        status, evidence = "missing", ""
        for p in patterns:
            m = re.search(p, ctx["headings"], re.I)
            if m:
                status, evidence = "present", _line_with(ctx["headings"], m)
                break
        if status == "missing":
            for p in patterns:
                m = re.search(p, ctx["notes_text"], re.I)
                if m:
                    status, evidence = "present", "in note text: " + _line_with(ctx["notes_text"], m)
                    break
        if status == "missing" and item_id in ("authorisation", "rounding", "compliance", "legal_status"):
            for p in patterns:
                m = re.search(p, ctx["text"], re.I)
                if m:
                    status, evidence = "present", "elsewhere: " + _line_with(ctx["text"], m)
                    break
        rows.append({"id": item_id, "title": title, "ref": ref, "status": status, "evidence": evidence[:160]})
        if status == "missing":
            sev = "error" if when is None and item_id not in ("rounding", "corresponding", "new_standards") else "warning"
            issues.append(Issue("disclosure", sev, f"Disclosure not found: {title}", section=SECTIONS["NOTES"],
                                location=ref, suggestion=f"Add the disclosure required for a {cat_name} or confirm "
                                                         "why it is not needed."))
    present = sum(1 for r in rows if r["status"] == "present")
    issues.append(Issue("disclosure", "info", f"Disclosure checklist ({category}): {present} present, "
                        f"{sum(1 for r in rows if r['status'] == 'missing')} not found, "
                        f"{sum(1 for r in rows if r['status'] == 'n/a')} not applicable.",
                        status="passed" if present == len([r for r in rows if r["status"] != "n/a"]) else "open"))
    issues += _unreferenced_lines(doc)
    return rows, issues


def _line_with(text, m):
    a = text.rfind("\n", 0, m.start()) + 1
    b = text.find("\n", m.end())
    return text[a:b if b != -1 else len(text)].strip()


def _unreferenced_lines(doc, share=0.05):
    """Significant statement lines with no note reference."""
    issues = []
    for t in doc.tables():
        if t.section not in ("SOFP", "SOCI") or t.note_col is None or not t.amount_cols:
            continue
        base = max([abs(parse_amount(r.cells[t.amount_cols[0]]) or 0) for r in t.rows] + [0])
        if not base:
            continue
        missing = []
        for r in t.rows:
            if r.kind != "item" or r.cells[t.note_col].strip():
                continue
            v = parse_amount(r.cells[t.amount_cols[0]])
            if v is not None and abs(v) >= base * Decimal(str(share)) and not re.search(
                    r"retained|unappropriated|accumulated (?:profit|loss)|current portion|current tax", r.label, re.I):
                missing.append(r.label)
        if missing:
            issues.append(Issue("disclosure", "warning", "Significant lines with no note reference: "
                                + ", ".join(f"'{m}'" for m in missing[:8]), location=t.source,
                                section=SECTIONS.get(t.section),
                                suggestion="Add a note reference, or confirm the line needs no further disclosure."))
    return issues
