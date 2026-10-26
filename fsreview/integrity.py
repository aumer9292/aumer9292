"""Prove that formatting did not change the content.

The formatted Word file is read back and compared with the source: every
figure in every table cell, in order, must have the same value, all figures in
running text must still be there, and the wording may differ only by the
changes the tool recorded (spacing, capitalisation, and spelling corrections
when that option is on).
"""

import re
from collections import Counter

from .model import Issue
from .numbers import find_amounts, parse_amount

_WORD = re.compile(r"[a-z0-9]+")


def _figures(model):
    table_vals, text_vals = [], Counter()
    for b in model.blocks:
        if b.kind == "table":
            for r in b.rows:
                for c in r.cells:
                    v = parse_amount(c)
                    if v is not None:
                        table_vals.append(v)
        else:
            text_vals.update(find_amounts(b.text))
    return table_vals, text_vals


def _words(model):
    out = Counter()
    for b in model.blocks:
        if b.kind == "paragraph":
            out.update(_WORD.findall(b.text.lower()))
        else:
            for r in b.rows:
                for c in r.cells:
                    if parse_amount(c) is None:
                        out.update(_WORD.findall(c.lower()))
    return out


def compare(before, after, changes):
    issues = []
    tb, xb = _figures(before)
    ta, xa = _figures(after)
    if tb != ta:
        n = sum(1 for a, b in zip(tb, ta) if a != b) + abs(len(tb) - len(ta))
        issues.append(Issue("integrity", "error", f"{n} table figure(s) differ between the source and the formatted "
                            "document. Do not rely on the formatted file until this is resolved."))
    missing = xb - xa
    if missing:
        issues.append(Issue("integrity", "error", "Figures in the text are missing after formatting: "
                            + ", ".join(str(v) for v in list(missing)[:10])))
    wb, wa = _words(before), _words(after)
    allowed_out, allowed_in = Counter(), Counter()
    for _, old, new, reason in changes.text:
        allowed_out.update(_WORD.findall(old.lower()))
        allowed_in.update(_WORD.findall(new.lower()))
    lost = (wb - wa) - allowed_out
    added = (wa - wb) - allowed_in
    if lost:
        issues.append(Issue("integrity", "error", f"{sum(lost.values())} word(s) from the source were not found in "
                            "the formatted document: " + ", ".join(list(lost)[:12])))
    if added:
        issues.append(Issue("integrity", "warning", f"{sum(added.values())} word(s) appear in the formatted "
                            "document that were not in the source: " + ", ".join(list(added)[:12])))
    if not issues:
        issues.append(Issue("integrity", "info", f"Content preserved: all {len(tb)} table figures, "
                            f"{sum(xb.values())} figures in text and {sum(wb.values())} words match the source "
                            f"(after {len(changes.text)} recorded edits).", status="passed"))
    return issues


def compare_figures(source, converted):
    """Converting PDF, text or Excel into Word must carry every table figure across."""
    if _figures(source)[0] != _figures(converted)[0]:
        return [Issue("integrity", "error", "Figures changed while converting the source into Word.")]
    return []
