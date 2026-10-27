"""Accounting cross-checks: casting, balancing and agreement between statements."""

import re
from decimal import Decimal

from .classify import PRIMARY, SECTIONS, is_financial_table, row_values
from .model import Issue
from .numbers import format_amount, parse_amount

TOL = Decimal("1")      # rounding tolerance in the presentation unit


def _fmt(v):
    return format_amount(v, 2 if v != v.to_integral_value() else 0) if v is not None else "n/a"


def _col_name(t, c):
    y = t.column_years.get(c)
    if y:
        return str(y)
    if c in t.amount_cols and len(t.amount_cols) <= 2:
        return ("current year", "comparative year")[t.amount_cols.index(c)]
    return f"column {c + 1}"


def _find_rows(doc, section, pattern):
    rx = re.compile(pattern, re.I)
    out = []
    for t in doc.tables():
        if t.section != section:
            continue
        for r in t.rows:
            if r.label and rx.search(r.label) and any(v is not None for v in row_values(t, r)):
                out.append((t, r))
    return out


def _first_value(doc, section, pattern, col_index=0):
    """Value in the n-th amount column (0 = current year) of the first matching row."""
    rows = _find_rows(doc, section, pattern)
    for t, r in rows:
        vals = row_values(t, r)
        if len(vals) > col_index and vals[col_index] is not None:
            return vals[col_index], t, r
    return None, None, None


def run_checks(doc, settings):
    issues = []
    issues += cast_tables(doc)
    issues += check_sofp(doc)
    issues += check_soci(doc)
    issues += check_socf(doc)
    issues += check_soce(doc)
    issues += check_comparatives(doc, settings)
    issues += check_note_references(doc)
    issues += check_notes_agree(doc)
    issues += check_presence(doc)
    issues += check_restatement(doc)
    issues += check_annexed_range(doc)
    return issues


def check_restatement(doc):
    """Restated comparatives need a third statement of financial position and a restatement note."""
    issues = []
    restated = any(re.search(r"\brestated\b", " ".join(r.cells), re.I)
                   for t in doc.tables() if t.section in PRIMARY for r in t.rows[:4])
    restated = restated or any(re.search(r"\(restated\)|\brestated\b", p.text, re.I)
                               for p in doc.paragraphs() if p.section in PRIMARY)
    if not restated:
        return issues
    sofp = [t for t in doc.tables() if t.section == "SOFP" and is_financial_table(t)]
    if sofp and len(sofp[0].amount_cols) < 3:
        issues.append(Issue("accounting", "error", "Comparatives are marked 'Restated' but the statement of "
                            "financial position has no third column as at the beginning of the preceding period.",
                            location=sofp[0].source, section=SECTIONS["SOFP"],
                            suggestion="Present a third statement of financial position as at the start of the "
                                       "comparative year (IAS 1.40A–40D; Section 10 of IFRS for SMEs for MSCs)."))
    notes = "\n".join(p.text for p in doc.paragraphs() if p.section in ("NOTES", "POLICIES", "ESTIMATES"))
    if not re.search(r"restat|prior period (?:error|adjustment)|correction of (?:an )?error", notes, re.I):
        issues.append(Issue("disclosure", "error", "Comparatives are restated but no note explains the restatement.",
                            section=SECTIONS["NOTES"],
                            suggestion="Disclose the nature of the prior period error and the correction to each "
                                       "line item affected (IAS 8.49)."))
    else:
        issues.append(Issue("disclosure", "info", "Restatement is explained in the notes.", status="passed"))
    return issues


def check_annexed_range(doc):
    """'The annexed notes 1 to 33 form an integral part…' must name the last note that exists."""
    issues = []
    last = max([int(p.note_number) for p in doc.paragraphs() if p.role == "note_heading" and p.note_number
                and p.note_number.isdigit()] + [0])
    if not last:
        return issues
    seen = set()
    for p in doc.paragraphs():
        m = re.search(r"annexed notes?\s+(?:from\s+)?(\d+)\s*(?:to|-|–)\s*(\d+)", p.text, re.I)
        if not m or m.group(0) in seen:
            continue
        seen.add(m.group(0))
        if int(m.group(2)) != last:
            issues.append(Issue("consistency", "error", f"“{m.group(0)}” — the last note in the financial "
                                f"statements is note {last}.", location=p.source, section=SECTIONS.get(p.section),
                                suggestion=f"Change to 'The annexed notes 1 to {last} form an integral part of these "
                                           "financial statements.'"))
        if re.search(r"notes?\s+from\s+\d+\s*-", m.group(0), re.I):
            issues.append(Issue("proofreading", "info", f"Usual wording is 'The annexed notes 1 to {m.group(2)} form "
                                f"an integral part…' (found “{m.group(0)}”).", location=p.source))
    return issues


# ------------------------------------------------------------- casting

BALANCE = re.compile(r"^(?:balance|opening|closing|as at|at (?:the )?(?:beginning|start|end)|(?:loan|amount)s? "
                     r"(?:payable|receivable|outstanding) at|carrying amount at|written down value as at)", re.I)


def _cast_schedule(t, sec):
    """Movement schedules (fixed assets, loans, provisions): opening + movements = closing, closing
    carried forward as the next opening, and a 'Total' column that adds across."""
    issues = []
    cols = t.amount_cols
    n = len(cols)
    checked = bad = 0
    total_col = None
    if n >= 3 and re.search(r"\btotal\b", " ".join(r.cells[cols[-1]] for r in t.rows[:4]), re.I):
        total_col = cols[-1]
    opening, moves, last_closing = None, [], None
    for r in t.rows:
        if r.kind in ("header", "blank", "info"):
            continue
        vals = row_values(t, r)
        if r.kind == "heading" and not any(v is not None for v in vals):
            opening, moves, last_closing = None, [], None
            continue
        if total_col is not None and r.kind != "heading":
            parts = [parse_amount(r.cells[c]) for c in cols[:-1]]
            tot = parse_amount(r.cells[total_col])
            if tot is not None and any(v is not None for v in parts):
                ssum = sum((v for v in parts if v is not None), Decimal(0))
                checked += 1
                if abs(ssum - tot) > TOL:
                    bad += 1
                    issues.append(Issue("casting", "error", f"'{r.label}': the columns add to {_fmt(ssum)} but the "
                                        f"total column shows {_fmt(tot)}.", location=t.source, section=sec))
        label = r.label or ""
        texts = [r.cells[c] for c in range(len(r.cells)) if c not in cols and r.cells[c].strip()]
        if re.match(r"^(?:wdv|written down value|net book value|carrying amount)", label, re.I):
            continue
        bal = next((x for x in texts if BALANCE.search(x)), None)
        if bal:
            label = bal
            is_open = bool(re.search(r"\b(?:opening|beginning|start)\b|\b0?1(?:st)?,?\s+(?:july|january|april|october)\b|"
                                     r"\b(?:july|january|april|october)\s+0?1\b", label, re.I))
            is_close = bool(re.search(r"\b(?:closing|end)\b|\b(?:30|31)(?:st|th)?,?\s+[a-z]+|[a-z]+\s+(?:30|31)\b",
                                      label, re.I)) and not is_open
            if is_close and opening is None:
                last_closing = vals          # a balance shown on its own (no movements to reconcile)
                continue
            if opening is None:
                # carried forward from the closing just above: only for dated balances of the same schedule
                if last_closing is not None and re.search(r"\b(19|20)\d{2}\b", label):
                    checked += 1
                    diff = [k for k in range(n) if vals[k] is not None and last_closing[k] is not None
                            and abs(vals[k] - last_closing[k]) > TOL]
                    if diff:
                        bad += 1
                        issues.append(Issue("casting", "error", f"'{label}' does not equal the closing balance "
                                            "just above it.", location=t.source, section=sec))
                opening, moves = vals, []
                continue
            calc = _sum([opening] + moves, n)
            checked += 1
            diffs = [f"{_col_name(t, cols[k])}: shown {_fmt(vals[k])}, opening + movements {_fmt(calc[k])}"
                     for k in range(n) if vals[k] is not None and calc[k] is not None and abs(vals[k] - calc[k]) > TOL]
            if diffs:
                bad += 1
                issues.append(Issue("casting", "error", f"'{label}' does not reconcile — " + "; ".join(diffs),
                                    location=t.source, section=sec,
                                    suggestion="Check the movements between the opening and closing balances."))
            last_closing, opening, moves = vals, None, []
            continue
        if opening is not None:
            moves.append(vals)
    if checked and not bad:
        issues.append(Issue("casting", "info", f"Movement schedule reconciles ({checked} check(s)).",
                            location=t.source, section=sec, status="passed"))
    return issues


def _is_schedule(t):
    texts = [x for r in t.rows if r.kind not in ("header", "blank")
             for c, x in enumerate(r.cells) if c not in t.amount_cols and x.strip()]
    return t.section != "SOCE" and sum(1 for x in texts if BALANCE.search(x)) >= 2


def _is_effect_table(doc, t):
    """Tables showing the effect of a restatement or change in policy: increase/(decrease) per line."""
    head = " ".join(" ".join(r.cells) for r in t.rows if r.kind in ("header", "heading"))
    before, heading = [], ""
    for b in doc.blocks:
        if b is t:
            break
        if b.kind == "paragraph" and b.text.strip():
            before = (before + [b.text])[-3:]
            if b.role == "note_heading" and b.note_number and "." not in b.note_number:
                heading = b.text
    return bool(re.search(r"increase\s*/\s*\(?decrease|effect of (?:the )?restatement|impact of",
                          head + " " + " ".join(before), re.I)) or \
        bool(re.search(r"restatement|change in accounting polic|prior period", heading, re.I))


def cast_tables(doc):
    issues = []
    carry = None          # lines left open at the end of a table that continues on the next page
    for ti, t in enumerate(doc.tables()):
        if not is_financial_table(t):
            continue
        sec = SECTIONS.get(t.section, t.section)
        if t.section == "SOCE":
            issues += _cast_soce(t)
            continue
        if _is_effect_table(doc, t):
            continue          # an increase/(decrease) table in a restatement note does not add down
        if _is_schedule(t):
            issues += _cast_schedule(t, sec)
            continue
        cols = t.amount_cols
        n = len(cols)
        items, stack, prev_total = [], [], None
        if carry and carry[0] == (t.note_number, n) and t.note_number:
            items = list(carry[1])
        checked = failed = 0
        for r in t.rows:
            if r.kind in ("header", "blank", "info"):
                continue
            if r.kind == "heading":
                # lines left without their own total (e.g. one borrowing line under
                # 'Non-current liabilities') count as a group of their own
                if items:
                    stack.append(_sum(items, n))
                    items = []
                continue
            vals = row_values(t, r)
            if r.kind == "item":
                items.append(vals)
                continue
            present = [abs(v) for v in vals if v is not None]
            if present and all(v < 1000 and v == v.to_integral_value() for v in present) and \
                    any(abs(v) >= 100000 for it in items + stack for v in it if v is not None):
                continue          # a count (persons, employees) among rupee amounts
            # A total is built in one of the usual ways. 'pop' says how many earlier
            # totals it replaces on the stack when it matches.
            candidates = []
            if items:
                candidates.append(("the lines above it", _sum(items, n), 0))
                if prev_total is not None:
                    candidates.append(("the previous total and the lines since", _sum([prev_total] + items, n), 1))
            for k in range(2, len(stack) + 1):
                candidates.append(("the totals above it", _sum(stack[-k:], n), k))
            if len(stack) >= 2:
                candidates.append(("the difference of the two totals above it",
                                   [a - b if a is not None and b is not None else None
                                    for a, b in zip(stack[-2], stack[-1])], 2))
            if stack and items:
                for k in range(1, len(stack) + 1):
                    candidates.append(("the totals and lines above it", _sum(stack[-k:] + items, n), k))
            hit = next((c for c in candidates if _close(vals, c[1])), None)
            if candidates:
                checked += 1
                if hit is None:
                    best = min(candidates, key=lambda c: _gap(vals, c[1]))
                    diffs = [f"{_col_name(t, cols[k])}: shown {_fmt(vals[k])}, computed {_fmt(best[1][k])}"
                             for k in range(n) if vals[k] is not None and best[1][k] is not None
                             and abs(vals[k] - best[1][k]) > TOL]
                    worst = max((abs(vals[k] - best[1][k]) for k in range(n)
                                 if vals[k] is not None and best[1][k] is not None), default=Decimal(0))
                    rounding = worst <= Decimal("0.5") * (len(items) + len(stack) + 1) + 1
                    if diffs and rounding:
                        issues.append(Issue("casting", "warning",
                                            f"'{r.label or '(unlabelled total)'}' is out by a rounding difference "
                                            f"against {best[0]} — " + "; ".join(diffs),
                                            location=t.source, section=sec,
                                            suggestion="Round each figure (e.g. =ROUND(x,0)) so the printed lines "
                                                       "add up to the printed total."))
                    elif diffs:
                        failed += 1
                        issues.append(Issue("casting", "error",
                                            f"'{r.label or '(unlabelled total)'}' does not add up to "
                                            f"{best[0]} — " + "; ".join(diffs),
                                            location=t.source, section=sec,
                                            suggestion="Re-cast the column and correct the total or the "
                                                       "line items."))
            if hit is not None and hit[2]:
                del stack[-hit[2]:]
            stack.append(vals)
            prev_total = vals
            items = []
        carry = ((t.note_number, n), items) if items and t.section in ("NOTES", "POLICIES", "ESTIMATES") else None
        if checked and not failed:
            issues.append(Issue("casting", "info", f"{checked} total(s) cast correctly.",
                                location=t.source, section=sec, status="passed"))
    return issues


def _cast_soce(t):
    """In the statement of changes in equity, the last column totals the components across."""
    issues = []
    cols = t.amount_cols
    if len(cols) < 3:
        return issues
    *parts, total_col = cols
    bad = 0
    for r in t.rows:
        if r.kind not in ("item", "total", "grand_total", "subtotal"):
            continue
        vals = [parse_amount(r.cells[c]) for c in parts]
        tot = parse_amount(r.cells[total_col])
        if tot is None or all(v is None for v in vals):
            continue
        s = sum((v for v in vals if v is not None), Decimal(0))
        if abs(s - tot) > TOL:
            bad += 1
            issues.append(Issue("casting", "error",
                                f"'{r.label}': the equity components add to {_fmt(s)} but the total column "
                                f"shows {_fmt(tot)}.", location=t.source, section=SECTIONS["SOCE"]))
    # closing balance = opening + movements, per column
    bal_rows = [i for i, r in enumerate(t.rows) if re.match(r"^(balance|as at|at )", r.label or "", re.I)
                and any(v is not None for v in row_values(t, r))]
    for a, b in zip(bal_rows, bal_rows[1:]):
        opening = row_values(t, t.rows[a])
        moves = [row_values(t, t.rows[i]) for i in range(a + 1, b)
                 if t.rows[i].kind not in ("header", "heading", "blank")
                 and not re.match(r"^total", t.rows[i].label, re.I)]
        closing = row_values(t, t.rows[b])
        calc = _sum([opening] + moves, len(cols))
        for k, c in enumerate(cols):
            if closing[k] is not None and calc[k] is not None and abs(closing[k] - calc[k]) > TOL:
                # transactions may be presented with a 'total transactions' row; try without subtotals
                bad += 1
                issues.append(Issue("casting", "error",
                                    f"'{t.rows[b].label}' in column {c + 1}: opening balance plus movements is "
                                    f"{_fmt(calc[k])} but {_fmt(closing[k])} is shown.",
                                    location=t.source, section=SECTIONS["SOCE"]))
    if not bad:
        issues.append(Issue("casting", "info", "Statement of changes in equity casts across and down.",
                            location=t.source, section=SECTIONS["SOCE"], status="passed"))
    return issues


def _sum(rows, n):
    out = []
    for k in range(n):
        vals = [r[k] for r in rows if k < len(r) and r[k] is not None]
        out.append(sum(vals, Decimal(0)) if vals else None)
    return out


def _close(vals, calc):
    pairs = [(v, c) for v, c in zip(vals, calc) if v is not None and c is not None]
    return bool(pairs) and all(abs(v - c) <= TOL for v, c in pairs)


def _gap(vals, calc):
    return sum((abs(v - c) for v, c in zip(vals, calc) if v is not None and c is not None), Decimal(0))


# ------------------------------------------------------ statement checks

def _agree(issues, a, b, what, section, location="", severity="error", suggestion=""):
    if a is None or b is None:
        return
    if abs(a - b) <= TOL:
        issues.append(Issue("accounting", "info", f"{what}: agrees at {_fmt(a)}.", location=location,
                            section=section, status="passed"))
    else:
        issues.append(Issue("accounting", severity, f"{what}: {_fmt(a)} vs {_fmt(b)} "
                                                    f"(difference {_fmt(a - b)}).",
                            location=location, section=section, suggestion=suggestion))


def check_sofp(doc):
    issues = []
    sec = SECTIONS["SOFP"]
    for k, year in ((0, "current year"), (1, "comparative year")):
        ta, t, _ = _first_value(doc, "SOFP", r"^total assets$", k)
        tel, _, _ = _first_value(doc, "SOFP", r"^total (equity and liabilities|liabilities and equity)$", k)
        if ta is not None and tel is not None:
            _agree(issues, ta, tel, f"Statement of financial position balances ({year}): total assets vs "
                                    "total equity and liabilities", sec, t.source,
                   suggestion="Find the difference before formatting is finalised.")
        elif ta is not None:
            te, _, _ = _first_value(doc, "SOFP", r"^total equity$", k)
            tl, _, _ = _first_value(doc, "SOFP", r"^total liabilities$", k)
            na, _, _ = _first_value(doc, "SOFP", r"^net assets$", k)
            if te is not None and tl is not None:
                _agree(issues, ta, te + tl, f"Statement of financial position balances ({year}): total assets "
                                            "vs total equity plus total liabilities", sec, t.source)
            elif na is not None and te is not None:
                _agree(issues, na, te, f"Net assets vs total equity ({year})", sec, t.source)
    if not any(i.section == sec and i.category == "accounting" for i in issues) and \
            any(t.section == "SOFP" for t in doc.tables()):
        issues.append(Issue("accounting", "warning", "Could not find 'Total assets' and 'Total equity and "
                            "liabilities' lines to prove the statement of financial position balances.",
                            section=sec, suggestion="Check the totals are labelled in the usual way."))
    return issues


PROFIT = r"^(?:net )?(?:profit|loss|\(loss\)|profit\s*/\s*\(loss\))(?: after tax(?:ation)?)? for the (?:year|period)"


def check_soci(doc):
    issues = []
    sec = SECTIONS["SOCI"]
    for k, year in ((0, "current year"), (1, "comparative year")):
        p, t, _ = _first_value(doc, "SOCI", PROFIT, k)
        oci, _, _ = _first_value(doc, "SOCI", r"^(total )?other comprehensive (income|loss)(?! that| which| to| items)", k)
        tci, _, _ = _first_value(doc, "SOCI", r"^total comprehensive", k)
        if p is not None and tci is not None:
            if oci is None:
                oci_rows = _find_rows(doc, "SOCI", r"comprehensive")
                if not oci_rows:
                    oci = Decimal(0)
            if oci is not None:
                _agree(issues, p + oci, tci, f"Total comprehensive income = profit + other comprehensive income "
                                             f"({year})", sec, t.source)
    return issues


def check_socf(doc):
    issues = []
    sec = SECTIONS["SOCF"]
    for k, year in ((0, "current year"), (1, "comparative year")):
        op, t, _ = _first_value(doc, "SOCF", r"^net cash (?:generated |used |inflow |outflow )?.*operating", k)
        inv, _, _ = _first_value(doc, "SOCF", r"^net cash .*investing", k)
        fin, _, _ = _first_value(doc, "SOCF", r"^net cash .*financing", k)
        net, _, _ = _first_value(doc, "SOCF", r"^net (?:increase|decrease|\(decrease\)|change|increase\s*/\s*\(decrease\))", k)
        opening, _, _ = _first_value(doc, "SOCF", r"(?:at (?:the )?(?:beginning|start)|opening)", k)
        closing, _, _ = _first_value(doc, "SOCF", r"(?:at (?:the )?end|closing)", k)
        if t is None:
            continue
        if None not in (op, inv, fin, net):
            _agree(issues, op + inv + fin, net, f"Operating + investing + financing cash flows = net change in cash "
                                                f"({year})", sec, t.source)
        if None not in (opening, net, closing):
            fx, _, _ = _first_value(doc, "SOCF", r"exchange (?:rate )?(?:differences|movements|gains|losses)|effect of", k)
            _agree(issues, opening + net + (fx or 0), closing,
                   f"Opening cash + net change{' + exchange differences' if fx else ''} = closing cash ({year})",
                   sec, t.source)
        if closing is not None:
            cash, ts, _ = _first_value(doc, "SOFP", r"^cash and (?:cash equivalents|bank balances?)|^cash at bank|^bank balances", k)
            note_ok = _closing_cash_note(doc, closing, k)
            if note_ok:
                issues.append(Issue("accounting", "info", f"Closing cash and cash equivalents ({year}) agree to "
                                    f"note {note_ok}.", location=t.source, section=sec, status="passed"))
            elif cash is not None:
                _agree(issues, closing, cash, f"Closing cash in the cash flow statement vs cash in the statement "
                                              f"of financial position ({year})", sec, t.source, severity="warning",
                       suggestion="A difference is acceptable only if explained, e.g. bank overdrafts included "
                                  "in cash and cash equivalents; disclose the reconciliation.")
    return issues


def _closing_cash_note(doc, closing, k):
    """The note the closing cash line cites, if that note shows the same figure."""
    for t in doc.tables():
        if t.section != "SOCF" or t.note_col is None:
            continue
        for r in t.rows:
            if re.search(r"(?:at (?:the )?end|closing)", r.label or "", re.I) and r.cells[t.note_col].strip():
                ref = re.match(r"\d{1,2}", r.cells[t.note_col].strip())
                if not ref:
                    continue
                for nt in doc.tables():
                    if nt.note_number and nt.note_number.split(".")[0] == ref.group(0) and nt.section != "SOCF":
                        for nr in nt.rows:
                            vals = row_values(nt, nr)
                            if len(vals) > k and vals[k] is not None and abs(vals[k] - closing) <= TOL:
                                return ref.group(0)
    return None


def check_soce(doc):
    issues = []
    sec = SECTIONS["SOCE"]
    soce = [t for t in doc.tables() if t.section == "SOCE" and is_financial_table(t)]
    if not soce:
        return issues
    t = soce[0]
    total_col = t.amount_cols[-1]
    bal = [r for r in t.rows if re.match(r"^(balance|as at|at )", r.label or "", re.I)
           and parse_amount(r.cells[total_col]) is not None]
    profit_rows = [r for r in t.rows if re.search(r"profit|loss", r.label or "", re.I)
                   and not re.search(r"comprehensive", r.label or "", re.I)
                   and parse_amount(r.cells[total_col]) is not None]
    if bal:
        closing = parse_amount(bal[-1].cells[total_col])
        te, _, _ = _first_value(doc, "SOFP", r"^total equity$|^equity attributable|^total shareholders'? equity", 0)
        _agree(issues, closing, te, "Closing total equity in the statement of changes in equity vs the statement "
                                    "of financial position (current year)", sec, t.source)
        if len(bal) >= 3:
            mid = parse_amount(bal[-2].cells[total_col])
            te1, _, _ = _first_value(doc, "SOFP", r"^total equity$|^equity attributable|^total shareholders'? equity", 1)
            _agree(issues, mid, te1, "Comparative closing equity in the statement of changes in equity vs the "
                                     "statement of financial position", sec, t.source)
    rows = t.rows
    for i, r in enumerate(rows):
        if not re.match(r"^total comprehensive (?:income|loss|\(loss\))", r.label or "", re.I):
            continue
        tci = parse_amount(r.cells[total_col])
        block = []
        for j in range(i - 1, -1, -1):
            if re.match(r"^(balance|as at|at )", rows[j].label or "", re.I):
                break
            block.append(rows[j])
        parts = [parse_amount(x.cells[total_col]) for x in block
                 if re.search(r"profit|loss|other comprehensive", x.label or "", re.I)
                 and parse_amount(x.cells[total_col]) is not None]
        expected = sum(parts, Decimal(0)) if parts else None
        if tci is not None and expected is not None and abs(tci - expected) > TOL:
            nxt = next((parse_amount(x.cells[total_col]) for x in rows[i + 1:]
                        if re.match(r"^(balance|as at|at )", x.label or "", re.I)), None)
            why = " — it shows the closing balance instead" if nxt is not None and abs(nxt - tci) <= TOL else ""
            issues.append(Issue("accounting", "error", f"'{r.label}' in the statement of changes in equity shows "
                                f"{_fmt(tci)}, but profit plus other comprehensive income is {_fmt(expected)}{why}.",
                                location=t.source, section=sec,
                                suggestion="Show total comprehensive income (profit + OCI) on this line; the "
                                           "balance belongs on the 'Balance as at' line."))
    if profit_rows:
        p_soce = parse_amount(profit_rows[-1].cells[total_col])
        p_soci, _, _ = _first_value(doc, "SOCI", PROFIT, 0)
        _agree(issues, p_soce, p_soci, "Profit for the year in the statement of changes in equity vs the "
                                       "statement of profit or loss", sec, t.source)
    return issues


# ------------------------------------------------------ comparatives

def check_comparatives(doc, settings):
    issues = []
    fy = settings.get("financial_year")
    try:
        fy = int(fy) if fy else None
    except ValueError:
        fy = None
    for t in doc.tables():
        if t.section not in PRIMARY or not is_financial_table(t):
            continue
        sec = SECTIONS[t.section]
        years = [t.column_years.get(c) for c in t.amount_cols]
        if t.section != "SOCE":
            if len(t.amount_cols) < 2:
                issues.append(Issue("consistency", "warning", "Only one column of figures — comparative figures "
                                    "appear to be missing.", location=t.source, section=sec,
                                    suggestion="Present comparative figures for the preceding period."))
            if fy and years and years[0] and years[0] != fy:
                issues.append(Issue("consistency", "error", f"First column is headed {years[0]} but the financial "
                                    f"year selected is {fy}.", location=t.source, section=sec))
            ys = [y for y in years if y]
            if len(ys) >= 2 and ys[0] - ys[1] != 1:
                issues.append(Issue("consistency", "warning", f"Column years {ys[0]} and {ys[1]} are not "
                                    "consecutive.", location=t.source, section=sec))
            if len(ys) >= 2 and ys[0] < ys[1]:
                issues.append(Issue("consistency", "warning", "The comparative year is shown before the current "
                                    "year; the current year is normally presented first.",
                                    location=t.source, section=sec))
            if len(t.amount_cols) >= 2:
                missing = [r.label for r in t.rows if r.kind == "item" and r.label and
                           parse_amount(r.cells[t.amount_cols[0]]) is not None and
                           not r.cells[t.amount_cols[1]].strip()]
                if missing:
                    issues.append(Issue("consistency", "warning",
                                        f"{len(missing)} line(s) have a current-year figure but a blank comparative: "
                                        + ", ".join(f"'{m}'" for m in missing[:6]) + ("…" if len(missing) > 6 else ""),
                                        location=t.source, section=sec,
                                        suggestion="Show '-' for nil comparatives so the blank is not mistaken "
                                                   "for an omission."))
    units = sorted({t.unit_label for t in doc.tables() if t.unit_label and t.section in PRIMARY})
    norm = {re.sub(r"[\s'’]", "", u).lower() for u in units}
    if len(norm) > 1:
        issues.append(Issue("consistency", "warning", "Different currency/unit labels are used across the primary "
                            "statements: " + ", ".join(units),
                            suggestion="Use one presentation currency and unit label throughout."))
    return issues


# ---------------------------------------------------------- notes

def _note_numbers(doc):
    nums = []
    for p in doc.paragraphs():
        if p.role == "note_heading" and p.note_number and "." not in p.note_number \
                and not re.search(r"continued|cont'?d\.?", p.text, re.I):
            nums.append((int(p.note_number), p))
    return nums


def check_note_references(doc):
    issues = []
    notes = _note_numbers(doc)
    existing = {n for n, _ in notes}
    if notes:
        seq = [n for n, _ in notes]
        dups = sorted({n for n in seq if seq.count(n) > 1})
        if dups:
            issues.append(Issue("consistency", "error", "Duplicate note numbers: " + ", ".join(map(str, dups)),
                                section=SECTIONS["NOTES"], suggestion="Renumber the notes and update references."))
        gaps = [n for n in range(min(seq), max(seq) + 1) if n not in existing]
        if gaps:
            issues.append(Issue("consistency", "warning", "Note numbering skips: " + ", ".join(map(str, gaps)),
                                section=SECTIONS["NOTES"]))
        out_of_order = [b for a, b in zip(seq, seq[1:]) if b < a]
        if out_of_order:
            issues.append(Issue("consistency", "warning", "Notes are out of numerical order near note "
                                + ", ".join(map(str, out_of_order[:5])), section=SECTIONS["NOTES"]))
    refs = set()
    for t in doc.tables():
        if t.note_col is None or t.section not in PRIMARY:
            continue
        for r in t.rows:
            if r.kind in ("header",):
                continue
            for ref in re.findall(r"\d{1,2}", r.cells[t.note_col]):
                refs.add(int(ref))
                if notes and int(ref) not in existing:
                    issues.append(Issue("consistency", "error", f"'{r.label}' refers to note {ref}, but there is "
                                        f"no note {ref}.", location=t.source, section=SECTIONS.get(t.section),
                                        suggestion="Correct the note reference."))
    if notes and refs:
        unreferenced = sorted(existing - refs - {n for n in existing if n <= 4})
        if unreferenced and len(unreferenced) < len(existing):
            issues.append(Issue("consistency", "info", "Notes not referenced from the primary statements: "
                                + ", ".join(map(str, unreferenced)) + " (this can be correct for narrative notes).",
                                section=SECTIONS["NOTES"]))
    if refs and not notes:
        issues.append(Issue("consistency", "warning", "Note references are used in the statements but note "
                            "headings could not be identified, so the references were not verified.",
                            section=SECTIONS["NOTES"]))
    return issues


def check_notes_agree(doc):
    """A statement line that cites note N should show a figure that appears in note N."""
    issues = []
    note_values = {}
    for t in doc.tables():
        if t.note_number and t.section in ("NOTES", "POLICIES", "ESTIMATES"):
            vals = note_values.setdefault(int(t.note_number), set())
            for r in t.rows:
                for c in t.amount_cols:
                    v = parse_amount(r.cells[c])
                    if v is not None:
                        vals.add(abs(v))
    if not note_values:
        return issues
    agreed = 0
    for t in doc.tables():
        if t.note_col is None or t.section not in ("SOFP", "SOCI") or not t.amount_cols:
            continue
        for r in t.rows:
            refs = re.findall(r"\d{1,2}", r.cells[t.note_col])
            if len(refs) != 1:
                continue
            n = int(refs[0])
            if n not in note_values:
                continue
            v = parse_amount(r.cells[t.amount_cols[0]])
            if v is None or v == 0:
                continue
            if abs(v) in note_values[n]:
                agreed += 1
            else:
                issues.append(Issue("accounting", "warning",
                                    f"'{r.label}' ({_fmt(v)}) cites note {n}, but that figure does not appear in "
                                    f"note {n}.", location=t.source, section=SECTIONS.get(t.section),
                                    suggestion="Agree the statement figure to the note total."))
    if agreed:
        issues.append(Issue("accounting", "info", f"{agreed} statement figure(s) agree to their notes.",
                            section=SECTIONS["NOTES"], status="passed"))
    return issues


def check_presence(doc):
    issues = []
    found = {b.section for b in doc.blocks}
    for key in PRIMARY:
        if key not in found:
            issues.append(Issue("extraction", "warning", f"{SECTIONS[key]} was not found.",
                                suggestion="If it is present, check its title so the tool can recognise it; "
                                           "otherwise confirm it is meant to be omitted."))
    if "NOTES" not in found and "POLICIES" not in found:
        issues.append(Issue("extraction", "warning", "No notes to the financial statements were found."))
    return issues
