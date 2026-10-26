"""Document model shared by every reader, check and writer.

A document is an ordered list of blocks. A block is a paragraph or a table.
Readers fill in text and table cells; the classifier adds the section,
paragraph role and row kinds. Nothing in the source is dropped: content that
could not be read reliably is kept as a paragraph and marked ``needs_review``.
"""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Paragraph:
    text: str
    bold: bool = False
    style: str = ""
    source: str = ""            # where it came from, e.g. "page 3" or "paragraph 41"
    needs_review: str = ""      # reason, when extraction was unreliable
    section: str = ""           # set by the classifier
    role: str = ""              # title, heading, note_heading, body, period, unit, contents
    note_number: Optional[str] = None
    index: int = -1             # position in the source container (docx paragraph index)

    kind = "paragraph"


@dataclass
class Row:
    cells: list
    bold: bool = False
    kind: str = ""              # header, heading, item, subtotal, total, grand_total, blank
    label: str = ""


@dataclass
class Table:
    rows: list
    source: str = ""
    needs_review: str = ""
    section: str = ""
    label_col: int = 0
    note_col: Optional[int] = None
    amount_cols: list = field(default_factory=list)
    column_years: dict = field(default_factory=dict)   # col -> year shown in header
    unit_label: str = ""
    note_number: Optional[str] = None
    index: int = -1

    kind = "table"

    @property
    def width(self):
        return max((len(r.cells) for r in self.rows), default=0)

    def cell(self, r, c):
        row = self.rows[r].cells
        return row[c] if c < len(row) else ""


@dataclass
class Document:
    blocks: list
    source_name: str = ""
    source_type: str = ""
    warnings: list = field(default_factory=list)

    def paragraphs(self):
        return [b for b in self.blocks if b.kind == "paragraph"]

    def tables(self):
        return [b for b in self.blocks if b.kind == "table"]

    def text(self):
        out = []
        for b in self.blocks:
            if b.kind == "paragraph":
                out.append(b.text)
            else:
                for r in b.rows:
                    out.append("\t".join(r.cells))
        return "\n".join(out)


@dataclass
class Issue:
    category: str       # accounting, casting, proofreading, formatting, consistency, extraction, integrity
    severity: str       # error, warning, info
    message: str
    location: str = ""
    section: str = ""
    suggestion: str = ""
    status: str = "open"    # open, fixed, passed

    def as_dict(self):
        return {k: getattr(self, k) for k in
                ("category", "severity", "message", "location", "section", "suggestion", "status")}
