"""Key figures read from the primary statements, for size classification and analysis.

Each figure is a (current year, comparative year) pair of Decimals (either may
be None). Labels are matched with the wording used in IFRS statements and in
Pakistani practice (sales - net, trade debts, stock-in-trade, unappropriated
profit, long term financing ...).
"""

import re

from .classify import is_financial_table, row_values

# key -> (section, label pattern). First match wins, so specific patterns come first.
LINES = [
    ("revenue", "SOCI", r"^(?:net )?(?:revenue|sales|turnover)(?:\s*[-–]\s*net)?(?: from contracts with customers)?(?:, net)?$"
                        r"|^revenue from|^net (?:sales|revenue)|^sales\b|^turnover\b"),
    ("cost_of_sales", "SOCI", r"^cost of (?:sales|revenue|goods sold|services)"),
    ("gross_profit", "SOCI", r"^gross (?:profit|loss|profit\s*/\s*\(loss\))"),
    ("distribution", "SOCI", r"^(?:distribution|selling|marketing)"),
    ("admin", "SOCI", r"^administrative|^general and administrative"),
    ("other_income", "SOCI", r"^other (?:income|operating income)"),
    ("operating_profit", "SOCI", r"^(?:operating profit|profit from operations|operating (?:loss|profit\s*/\s*\(loss\)))"),
    ("finance_cost", "SOCI", r"^finance (?:cost|costs|charges)|^interest expense|^financial charges"),
    ("pbt", "SOCI", r"(?:profit|loss|\(loss\)).*before (?:income )?tax"),
    ("tax", "SOCI", r"^(?:income )?tax(?:ation)?(?: expense| charge)?$|^taxation|^income tax"),
    ("profit", "SOCI", r"^(?:net )?(?:profit|loss|\(loss\)|profit\s*/\s*\(loss\))(?: after (?:income )?tax(?:ation)?)?(?: for the (?:year|period))?$"),
    ("eps", "SOCI", r"earnings per share|^eps"),
    ("depreciation", "SOCF", r"^depreciation"),
    ("cfo", "SOCF", r"^net cash (?:generated |used |inflow |outflow |\(used in\)\s*/?\s*)?.*operating"),
    ("cfi", "SOCF", r"^net cash .*investing"),
    ("cff", "SOCF", r"^net cash .*financing"),
    ("capex", "SOCF", r"(?:purchase|acquisition|additions?) (?:of|to) (?:property|fixed assets|operating fixed)|capital expenditure"),
    ("dividends_paid", "SOCF", r"^dividends? paid"),
    ("ppe", "SOFP", r"^property, plant and equipment|^operating fixed assets|^fixed assets"),
    ("inventory", "SOFP", r"^inventor(?:y|ies)|^stock[- ]in[- ]trade|^stocks?$"),
    ("stores", "SOFP", r"^stores(?:,| and) spares"),
    ("receivables", "SOFP", r"^trade (?:debts|receivables|debtors)|^trade and other receivables|^accounts receivable"),
    ("cash", "SOFP", r"^cash and (?:cash equivalents|bank balances?)|^cash at bank|^bank balances"),
    ("share_capital", "SOFP", r"(?:issued, subscribed and )?paid[- ]up (?:share )?capital|^share capital$|^issued capital"),
    ("payables", "SOFP", r"^trade and other payables|^trade (?:payables|creditors)|^creditors, accrued"),
    ("lt_borrowings", "SOFP", r"^long[- ]term (?:financing|finances|loans?|borrowings)|^borrowings$|^loans? from banks"),
    ("st_borrowings", "SOFP", r"^short[- ]term (?:borrowings|finances?|running finance)|^running finance|^bank overdrafts?"),
    ("current_ltb", "SOFP", r"^current (?:portion|maturity) of"),
    ("lease_liabilities", "SOFP", r"lease liabilit"),
    ("total_assets", "SOFP", r"^total assets$"),
    ("total_equity", "SOFP", r"^total equity$|^total shareholders'? equity|^equity attributable|^total share capital and reserves"),
    ("total_liabilities", "SOFP", r"^total liabilities$"),
    ("total_non_current_assets", "SOFP", r"^total non[- ]current assets"),
    ("total_current_assets", "SOFP", r"^total current assets"),
    ("total_non_current_liabilities", "SOFP", r"^total non[- ]current liabilities"),
    ("total_current_liabilities", "SOFP", r"^total current liabilities"),
]
_RX = [(k, s, re.compile(p, re.I)) for k, s, p in LINES]

GROUPS = [
    ("total_non_current_assets", r"^non[- ]current assets"),
    ("total_current_assets", r"^current assets"),
    ("total_non_current_liabilities", r"^non[- ]current liabilities"),
    ("total_current_liabilities", r"^current liabilities"),
    ("total_equity", r"^(?:equity|share capital and reserves|capital and reserves|shareholders'? equity)$"),
]


def extract_figures(doc):
    found = {}
    for t in doc.tables():
        if t.section not in ("SOFP", "SOCI", "SOCF") or not is_financial_table(t):
            continue
        group = None
        for r in t.rows:
            label = (r.label or "").strip().rstrip(":")
            if r.kind == "heading":
                for key, rx in GROUPS:
                    if re.match(rx, label, re.I):
                        group = key
                        break
                continue
            vals = row_values(t, r)
            if not vals or all(v is None for v in vals):
                continue
            pair = (vals[0], vals[1] if len(vals) > 1 else None)
            if r.kind in ("total", "subtotal") and not label and group:
                found.setdefault(group, pair)
                group = None
                continue
            for key, sec, rx in _RX:
                if sec == t.section and rx.search(label) and key not in found:
                    found[key] = pair
                    break
            if r.kind in ("grand_total",):
                group = None
    _derive(found)
    return found


def _derive(f):
    def get(k, i):
        v = f.get(k)
        return v[i] if v else None

    for i in (0, 1):
        def setd(key, value):
            if value is None:
                return
            cur = list(f.get(key, (None, None)))
            if cur[i] is None:
                cur[i] = value
                f[key] = tuple(cur)

        ta, te, tl = get("total_assets", i), get("total_equity", i), get("total_liabilities", i)
        tca, tncl, tcl = get("total_current_assets", i), get("total_non_current_liabilities", i), get("total_current_liabilities", i)
        if tl is None and ta is not None and te is not None:
            setd("total_liabilities", ta - te)
        if tl is None and tncl is not None and tcl is not None:
            setd("total_liabilities", tncl + tcl)
        if get("total_non_current_assets", i) is None and ta is not None and tca is not None:
            setd("total_non_current_assets", ta - tca)
        rev, cos = get("revenue", i), get("cost_of_sales", i)
        if get("gross_profit", i) is None and rev is not None and cos is not None:
            setd("gross_profit", rev + cos if cos < 0 else rev - cos)
        debt = [get(k, i) for k in ("lt_borrowings", "st_borrowings", "current_ltb", "lease_liabilities")]
        if any(d is not None for d in debt):
            setd("total_debt", sum(abs(d) for d in debt if d is not None))
