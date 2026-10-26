"""Reference library: approved and illustrative financial statements kept for comparison.

Upload the ICAP illustrative financial statements for each category (PIC, LSC,
MSC, SSC) and any model set your firm follows. Each reference is read once; its
structure (sections and note headings) is kept, and every later review of the
same category is compared with it to point out notes the reference has that the
draft does not.
"""

import datetime as _dt
import json
import os
import re
import shutil
import uuid

from .classify import classify_document
from .entity import HOME
from .extract import read_document
from .model import Issue

LIB = os.path.join(HOME, "library")
STOP = set("the of and to for in on a an its from with by at as other its notes note".split())


def _norm(text):
    t = re.sub(r"^\s*(?:note\s+)?\d+(?:\.\d+)*\.?\s*", "", text.lower())
    t = re.sub(r"[^a-z ]+", " ", t)
    return [w for w in t.split() if w not in STOP and len(w) > 1]


def note_headings(doc):
    out = []
    for b in doc.blocks:
        if b.kind == "paragraph" and b.role == "note_heading" and b.note_number and "." not in b.note_number:
            title = re.sub(r"^\s*(?:note\s+)?\d+\.?\s*", "", b.text.strip())
            if title and title.lower() not in (h.lower() for h in out):
                out.append(title)
    return out


def add(path, name, kind, category):
    doc = classify_document(read_document(path))
    rid = uuid.uuid4().hex[:12]
    folder = os.path.join(LIB, rid)
    os.makedirs(folder, exist_ok=True)
    stored = os.path.join(folder, os.path.basename(path))
    shutil.copyfile(path, stored)
    meta = {
        "id": rid, "name": name or os.path.basename(path), "kind": kind, "category": category,
        "file": os.path.basename(path), "added": _dt.datetime.now().isoformat(timespec="seconds"),
        "sections": sorted({b.section for b in doc.blocks if b.section}),
        "note_headings": note_headings(doc),
        "statement_lines": {s: [r.label for t in doc.tables() if t.section == s for r in t.rows if r.label]
                            for s in ("SOFP", "SOCI", "SOCE", "SOCF")},
    }
    with open(os.path.join(folder, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2, ensure_ascii=False)
    return meta


def items():
    out = []
    if not os.path.isdir(LIB):
        return out
    for rid in sorted(os.listdir(LIB)):
        try:
            with open(os.path.join(LIB, rid, "meta.json"), encoding="utf-8") as f:
                out.append(json.load(f))
        except (OSError, ValueError):
            continue
    return out


def remove(rid):
    if re.fullmatch(r"[0-9a-f]{12}", rid or "") and os.path.isdir(os.path.join(LIB, rid)):
        shutil.rmtree(os.path.join(LIB, rid))
        return True
    return False


def file_of(meta):
    return os.path.join(LIB, meta["id"], meta["file"])


def for_category(category):
    return [m for m in items() if m.get("category") in (category, "ANY")]


def _match(a, b):
    sa, sb = set(a), set(b)
    if not sa or not sb:
        return False
    return len(sa & sb) / len(sa | sb) >= 0.5 or sa <= sb or sb <= sa


def concepts(heading):
    """The checklist items a heading is about, so 'Trade debts' matches 'Trade and other receivables'."""
    from .disclosures import CHECKLIST
    out = set()
    for item_id, _, _, _, _, patterns in CHECKLIST:
        if any(re.search(p, heading, re.I) for p in patterns):
            out.add(item_id)
    return out


def compare(doc, category):
    """Notes in the category's references that the draft does not have."""
    issues, rows = [], []
    mine_headings = note_headings(doc)
    mine = [_norm(h) for h in mine_headings]
    mine_concepts = set().union(*[concepts(h) for h in mine_headings]) if mine_headings else set()
    mine_text = doc.text().lower()
    for ref in for_category(category):
        missing = []
        for h in ref["note_headings"]:
            words = _norm(h)
            if not words:
                continue
            if any(_match(words, m) for m in mine) or (concepts(h) & mine_concepts):
                rows.append({"reference": ref["name"], "note": h, "status": "present"})
                continue
            phrase = " ".join(words)
            if phrase and phrase in re.sub(r"[^a-z ]+", " ", mine_text):
                rows.append({"reference": ref["name"], "note": h, "status": "present"})
                continue
            missing.append(h)
            rows.append({"reference": ref["name"], "note": h, "status": "missing"})
        if missing:
            issues.append(Issue("disclosure", "warning",
                                f"Compared with '{ref['name']}', these notes were not found: " + "; ".join(missing[:25])
                                + ("…" if len(missing) > 25 else ""),
                                suggestion="Check whether each applies to this entity; add those that do."))
        else:
            issues.append(Issue("disclosure", "info", f"Every note in '{ref['name']}' has a counterpart in the draft.",
                                status="passed"))
    return rows, issues
