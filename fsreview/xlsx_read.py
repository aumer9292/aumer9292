"""Reading financial statements prepared in Excel.

Audit files usually keep the statements, the notes and the working papers
(trial balances, ledgers, adjusting entries, tax workings) in one workbook. Only
the statement and note sheets are read as the financial statements: a sheet
qualifies when it is visible, is not named like a working paper, and has a print
area or a statement title near the top. Within it only the print area is read,
hidden rows and columns are skipped, and every figure is shown the way Excel
displays it. The working papers are listed in the report, not read.

Rows become paragraphs (titles, note headings such as "4.1 Allocation of
depreciation", narrative text) or table rows (anything with figures, plus the
column headings and sub-headings around them).
"""

import datetime as _dt
import re

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

from .classify import PERIOD_LINE, classify_title
from .model import Paragraph, Row, Table
from .numbers import looks_like_note_ref

WORKING = re.compile(
    r"\b(?:tb|t\.b|trial ?balance|ledger|gl|adjust\w*|working|workings|provision|tax|deff?err?ed|"
    r"issues?|sheet\d*|journal|mapping|recon\w*|pivot|data|difference|lead|calc\w*|fy\d{2,4}|"
    r"receivables? fy|payables? fy|payable|receivable|investment in jv|ageing|aging|schedule of entries)\b", re.I)


def _print_bounds(ws):
    """(min_col, min_row, max_col, max_row) of the first print range, or the used range."""
    area = ws.print_area
    if area:
        first = str(area).split(",")[0]
        ref = first.split("!")[-1].replace("$", "")
        try:
            return range_boundaries(ref), True
        except ValueError:
            pass
    return (1, 1, ws.max_column or 1, ws.max_row or 1), False


def _title_near_top(ws, bounds):
    c0, r0, c1, r1 = bounds
    for r in range(r0, min(r1, r0 + 8) + 1):
        for c in range(c0, min(c1, c0 + 4) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and classify_title(v.strip()) not in (None, "CONTENTS_ENTRY"):
                return True
    return False


def select_sheets(wb):
    """Return ([(worksheet, bounds)], [names of working papers], [names of hidden sheets])."""
    chosen, working, hidden = [], [], []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            hidden.append(ws.title)
            continue
        bounds, has_area = _print_bounds(ws)
        named_working = bool(WORKING.search(ws.title.strip()))
        statementish = classify_title(ws.title.strip()) not in (None, "CONTENTS_ENTRY") or \
            re.match(r"^\s*(?:notes?|policies|accounting policies)\b", ws.title, re.I)
        if statementish or (not named_working and (has_area or _title_near_top(ws, bounds))):
            chosen.append((ws, bounds))
        else:
            working.append(ws.title)
    return chosen, working, hidden


# ----------------------------------------------------------------- values

def _decimals(fmt):
    sect = (fmt or "General").split(";")[0]
    m = re.search(r"0\.(0+)", sect)
    return len(m.group(1)) if m else 0


def render(value, fmt="General", is_date=False):
    """Text of a cell as Excel shows it (close enough for review purposes)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value).upper()
    if isinstance(value, (_dt.datetime, _dt.date)) or is_date:
        try:
            return value.strftime("%d %B %Y").lstrip("0")
        except AttributeError:
            return str(value)
    if isinstance(value, (int, float)):
        fmt = fmt or "General"
        if "%" in fmt:
            return f"{value * 100:.{_decimals(fmt)}f}%"
        if fmt in ("General", "@", "0", "0.0", "0.00") or not re.search(r"#,##", fmt):
            if isinstance(value, int) or float(value).is_integer():
                iv = int(value)
                if 1900 <= iv <= 2100 or abs(iv) < 1000 or "#" not in fmt:
                    if abs(iv) >= 10000 and fmt == "General":
                        return f"({abs(iv):,})" if iv < 0 else f"{iv:,}"
                    return str(iv) if fmt in ("General", "@", "0") else f"{iv:.{_decimals(fmt)}f}"
            dec = _decimals(fmt) if fmt != "General" else (0 if abs(value) >= 1000 else 2)
            if fmt == "General" and abs(value) < 100:
                s = f"{value:.4f}".rstrip("0").rstrip(".")      # note numbers such as 4.1
                return s
            s = f"{abs(value):,.{dec}f}"
            return f"({s})" if value < 0 else s
        dec = _decimals(fmt)
        s = f"{abs(value):,.{dec}f}"
        if round(abs(value), dec) == 0 and '"-"' in fmt:
            return "-"
        return f"({s})" if value < 0 and round(value, dec) != 0 else s
    return re.sub(r"\s+", " ", str(value)).strip()


# ----------------------------------------------------------------- rows

class _Cell:
    __slots__ = ("col", "text", "num", "bold", "raw")

    def __init__(self, col, text, num, bold, raw):
        self.col, self.text, self.num, self.bold, self.raw = col, text, num, bold, raw


def _sheet_rows(ws, bounds, hidden_decimals=None):
    c0, r0, c1, r1 = bounds
    cols = [c for c in range(c0, c1 + 1)
            if not ws.column_dimensions[get_column_letter(c)].hidden]
    rows = []
    for r in range(r0, r1 + 1):
        if ws.row_dimensions[r].hidden:
            continue
        cells = []
        for c in cols:
            cell = ws.cell(r, c)
            v = cell.value
            if v is None or (isinstance(v, str) and not v.strip()):
                continue
            text = render(v, cell.number_format, getattr(cell, "is_date", False))
            text = re.sub(r"([-–—_=])\1{12,}", lambda m: m.group(1) * 12, text)     # banners padded to fit a column
            if not text or re.fullmatch(r"_+", text):
                if text:
                    cells.append(_Cell(c, text, False, False, v))
                continue
            num = isinstance(v, (int, float)) and not isinstance(v, bool)
            if hidden_decimals is not None and isinstance(v, float) and abs(v) >= 1 and \
                    abs(v - round(v)) > 0.0049 and _decimals(cell.number_format) == 0 and "%" not in cell.number_format:
                hidden_decimals.append(f"{ws.title}!{get_column_letter(c)}{r}")
            cells.append(_Cell(c, text, num, bool(cell.font and cell.font.bold), v))
        rows.append((r, cells))
    return rows


def _is_note_no(cell):
    return cell.num and looks_like_note_ref(cell.text) and abs(float(cell.raw)) < 100


def _row_kind(cells, label_col):
    if not cells:
        return "blank"
    nums = [c for c in cells if c.num]
    texts = [c for c in cells if not c.num]
    if len(cells) == 2 and _is_note_no(cells[0]) and not cells[1].num:
        return "note_heading"
    figures = [c for c in nums if not (_is_note_no(c) and c is cells[0])]
    if figures:
        return "table"
    if any(re.search(r"\b(19|20)\d{2}\b|\(?rupees\)?|^note$|['’]000", c.text, re.I) for c in texts) and len(cells) >= 2:
        return "table"
    if len(texts) == 1:
        return "text"
    return "multi_text"


def sheet_blocks(ws, bounds, source, hidden_decimals=None):
    rows = _sheet_rows(ws, bounds, hidden_decimals)
    kinds = [(_row_kind(cells, bounds[0]), r, cells) for r, cells in rows]
    blocks = []
    table = []

    def close():
        nonlocal table
        if table:
            blocks.append(_make_table(table, source))
            table = []

    def next_table_within(i, n=4):
        """Is a table row coming, past blank rows and at most n short text rows?"""
        seen = 0
        for k, _, cells in kinds[i + 1:]:
            if k == "table":
                return True
            if k == "blank":
                continue
            text = " ".join(c.text for c in cells)
            if re.fullmatch(r"[-–—=\s]*\(?[A-Za-z' ]{0,20}\)?[-–—=\s]*", text) and "-" in text:
                continue      # a '------ Rupees ------' banner belongs with the table below
            if k not in ("text", "multi_text") or len(text) >= 100 or classify_title(cells[0].text) \
                    or PERIOD_LINE.search(text):
                return False
            seen += 1
            if seen >= n:
                return False
        return False

    for i, (kind, r, cells) in enumerate(kinds):
        loc = f"{source} row {r}"
        # "AS AT JUNE 30, 2026 ........ Restated": the period line is a paragraph; the rest heads the columns
        if cells and not cells[0].num and PERIOD_LINE.search(cells[0].text) and len(cells[0].text) < 120 \
                and kind in ("table", "multi_text") and not any(c.num for c in cells[1:]) \
                and not re.match(r"^\s*(balance|opening|closing)", cells[0].text, re.I):
            close()
            blocks.append(Paragraph(text=cells[0].text, bold=cells[0].bold, source=loc))
            cells = cells[1:]
            if not cells:
                continue
            kind = "table" if next_table_within(i) else "multi_text"
        if kind == "blank":
            if table and next_table_within(i):
                continue
            close()
            continue
        # "4 | OPERATING FIXED ASSETS | Computer equipment | Vehicles ...": the note heading, then column headings
        if len(cells) > 2 and _is_note_no(cells[0]) and not cells[1].num and cells[1].text.isupper() \
                and kind in ("table", "multi_text"):
            close()
            blocks.append(Paragraph(text=f"{cells[0].text} {cells[1].text}", bold=cells[1].bold, source=loc))
            cells = cells[2:]
            kind = "table" if next_table_within(i) or any(c.num for c in cells) else "multi_text"
        if kind == "table":
            table.append((r, cells))
            continue
        if kind == "multi_text" and _is_signature(cells):
            close()
            sig = [(r, cells)]
            blocks.append(_make_table(sig, source))
            continue
        text = "   ".join(c.text for c in cells)
        short = (len(text) < 100 or len(cells) >= 3) and not classify_title(cells[0].text)
        if kind in ("text", "multi_text") and short and (table or kind == "multi_text") and next_table_within(i):
            table.append((r, cells))
            continue
        close()
        if kind == "note_heading":
            text = f"{cells[0].text} {cells[1].text}"
        blocks.append(Paragraph(text=text, bold=all(c.bold for c in cells if not c.num) and bool(cells),
                                source=loc))
    close()
    return blocks


SIGNATORY = re.compile(r"^(?:_{4,}|chief executive(?: officer)?|(?:managing )?director|chief financial officer|ceo|cfo|"
                       r"company secretary|chairman|partner)$", re.I)


def _is_signature(cells):
    """The signature line under a statement: underscores, or 'Chief Executive Officer   Director'."""
    return len(cells) >= 2 and all(SIGNATORY.match(c.text.strip()) for c in cells)


def _make_table(rows, source):
    cols = sorted({c.col for _, cells in rows for c in cells})
    index = {c: i for i, c in enumerate(cols)}
    out = []
    for _, cells in rows:
        vals = [""] * len(cols)
        for c in cells:
            vals[index[c.col]] = (vals[index[c.col]] + " " + c.text).strip()
        bold = bool(cells) and all(c.bold for c in cells if not c.num)
        out.append(Row(cells=vals, bold=bold))
    # Excel spacer columns (narrow gaps between figure columns holding only dashes) are not columns of the statement
    keep = [i for i in range(len(cols))
            if any(r.cells[i] not in ("", "-") for r in out) or i == 0]
    if len(keep) < len(cols):
        for r in out:
            r.cells = [r.cells[i] for i in keep]
    # the label column is the first column holding text; drop empty leading columns
    first = 0
    while first < len(cols) - 1 and not any(r.cells[first] for r in out):
        first += 1
    for r in out:
        r.cells = r.cells[first:]
    # a leading column that only holds note numbers next to a text column becomes part of the label
    if len(out[0].cells) > 2 and all((not r.cells[0]) or looks_like_note_ref(r.cells[0]) for r in out) and \
            any(r.cells[0] for r in out) and all(r.cells[1] == "" or not re.fullmatch(r"[\d,().-]+", r.cells[1])
                                                 for r in out):
        for r in out:
            r.cells = [(r.cells[0] + " " + r.cells[1]).strip()] + r.cells[2:]
    return Table(rows=out, source=f"{source} rows {rows[0][0]}-{rows[-1][0]}")


def read_workbook(wb, values_wb=None):
    """Blocks for the statement and note sheets of an openpyxl workbook (values workbook preferred)."""
    vwb = values_wb or wb
    chosen, working, hidden = select_sheets(wb)
    blocks, hidden_decimals = [], []
    for ws, bounds in chosen:
        vws = vwb[ws.title]
        blocks.extend(sheet_blocks(vws, bounds, f"sheet '{ws.title}'", hidden_decimals))
    return blocks, [ws.title for ws, _ in chosen], working, hidden, hidden_decimals
