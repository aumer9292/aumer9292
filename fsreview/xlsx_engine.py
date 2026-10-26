"""Excel output.

An Excel input keeps every worksheet, formula, defined name and value: only
cell styles, number formats, column widths and print settings change, and the
result is compared cell by cell with the original before it is released.
Other inputs get a workbook with one worksheet per financial table, figures
stored as numbers.
"""

import re

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from .classify import GRAND_TOTAL, PRIMARY, SECTIONS, TOTAL, is_financial_table
from .model import Issue
from .numbers import YEAR, parse_amount

THIN = Side(style="thin", color="000000")
DOUBLE = Side(style="double", color="000000")


def _num_format(decimals, settings):
    d = "." + "0" * decimals if decimals else ""
    zero = '"-"' if settings.get("zero_style", "-") == "-" else "0" + d
    if settings.get("negative_style", "brackets") == "brackets":
        return f"#,##0{d}_);(#,##0{d});{zero}_)"
    return f"#,##0{d};-#,##0{d};{zero}"


def _is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def format_workbook(src, out, settings, changes):
    keep_vba = src.lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(src, keep_vba=keep_vba)
    vals = openpyxl.load_workbook(src, data_only=True)
    st = settings["style"]
    for ws in wb.worksheets:
        vws = vals[ws.title]
        _format_sheet(ws, vws, settings, st, changes)
    wb.save(out)
    return verify_workbook(src, out)


def _format_sheet(ws, vws, settings, st, changes):
    if ws.max_row < 1 or ws.max_column < 1:
        return
    maxc = min(ws.max_column, 60)
    maxr = min(ws.max_row, 5000)
    # label column: the first column holding mostly text
    label_col, best = 1, -1
    for c in range(1, min(maxc, 6) + 1):
        n = sum(1 for r in range(1, maxr + 1) if isinstance(vws.cell(r, c).value, str) and
                vws.cell(r, c).value.strip())
        if n > best:
            label_col, best = c, n
    amount_cols = []
    for c in range(1, maxc + 1):
        if c == label_col:
            continue
        cells = [ws.cell(r, c) for r in range(1, maxr + 1) if ws.cell(r, c).value is not None]
        nums = [x for x in cells if _is_number(x.value) or (isinstance(x.value, str) and x.value.startswith("="))]
        nums = [x for x in nums if not (isinstance(x.value, int) and 1900 <= x.value <= 2100)]
        if nums and len(nums) >= 0.5 * len(cells) and not all(isinstance(x.value, int) and abs(x.value) < 60
                                                               for x in nums):
            amount_cols.append(c)
    first_data = maxr + 1
    for r in range(1, maxr + 1):
        if any(_is_number(vws.cell(r, c).value) and not (isinstance(vws.cell(r, c).value, int)
               and 1900 <= vws.cell(r, c).value <= 2100) for c in amount_cols) or \
                any(isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.startswith("=") for c in amount_cols):
            first_data = r
            break
    decimals = {}
    for c in amount_cols:
        dec = 0
        for r in range(first_data, maxr + 1):
            v = vws.cell(r, c).value
            fmt = ws.cell(r, c).number_format or ""
            if _is_number(v) and isinstance(v, float) and not float(v).is_integer():
                m = re.search(r"\.(0+)", fmt)
                dec = max(dec, len(m.group(1)) if m else 2)
        decimals[c] = min(dec, 4)

    font = st["font"]
    size = st["table_size"]
    prev_kind = None
    for r in range(1, maxr + 1):
        label = str(vws.cell(r, label_col).value or "").strip()
        has_amt = any(vws.cell(r, c).value is not None or ws.cell(r, c).value is not None for c in amount_cols)
        formulas = [str(ws.cell(r, c).value) for c in amount_cols if isinstance(ws.cell(r, c).value, str)
                    and ws.cell(r, c).value.startswith("=")]
        if r < first_data:
            kind = "header" if has_amt or r < first_data else "heading"
        elif not label and not has_amt:
            kind = "blank"
        elif not has_amt:
            kind = "heading"
        elif GRAND_TOTAL.search(label):
            kind = "grand_total"
        elif TOTAL.search(label) or any(re.search(r"SUM\(", f, re.I) for f in formulas) or (not label and has_amt):
            kind = "total"
        else:
            kind = "item"
        for c in range(1, maxc + 1):
            cell = ws.cell(r, c)
            if cell.value is None and c not in amount_cols:
                continue
            f = cell.font
            bold = f.bold
            if kind == "header" or (kind in ("grand_total", "heading") and (c == label_col or c in amount_cols)):
                bold = True
            elif kind == "item" and c in amount_cols:
                bold = False
            cell.font = Font(name=font, size=size, bold=bold, italic=f.italic, underline=f.underline,
                             color=f.color if f.color is not None and f.color.type == "rgb" and f.color.rgb not in
                             (None, "FF000000") else st["color"])
            if c in amount_cols:
                fmt = cell.number_format or "General"
                is_pct_or_date = "%" in fmt or re.search(r"[dmy]{2,}", fmt, re.I)
                if kind == "header":
                    cell.alignment = Alignment(horizontal="right", vertical="bottom", wrap_text=True)
                elif not is_pct_or_date and (_is_number(cell.value) or (isinstance(cell.value, str)
                                                                         and cell.value.startswith("="))):
                    new = _num_format(decimals.get(c, 0), settings)
                    if cell.number_format != new:
                        cell.number_format = new
                        changes.add("Excel number formats standardised")
                    cell.alignment = Alignment(horizontal="right", vertical="bottom")
                if kind in ("total", "grand_total") and prev_kind in ("item", "total") and cell.value is not None:
                    cell.border = Border(top=THIN, bottom=DOUBLE if kind == "grand_total" else None)
                elif kind == "grand_total" and cell.value is not None:
                    cell.border = Border(bottom=DOUBLE)
        if kind != "blank":
            prev_kind = kind
    ws.column_dimensions[get_column_letter(label_col)].width = min(60, max(
        [len(str(vws.cell(r, label_col).value or "")) for r in range(1, maxr + 1)] + [20]) + 2)
    for c in amount_cols:
        ws.column_dimensions[get_column_letter(c)].width = 15 if decimals.get(c) else 14
    ws.sheet_view.showGridLines = False
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER if settings.get("page_size") == "Letter" else ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    changes.add("Excel worksheets formatted")


def verify_workbook(src, out):
    """Every cell's value or formula must be unchanged."""
    a = openpyxl.load_workbook(src)
    b = openpyxl.load_workbook(out)
    issues = []
    if a.sheetnames != b.sheetnames:
        issues.append(Issue("integrity", "error", "Worksheet list changed during formatting."))
        return issues
    diffs = 0
    formulas = 0
    for name in a.sheetnames:
        wa, wb_ = a[name], b[name]
        for row in wa.iter_rows():
            for c in row:
                v2 = wb_.cell(c.row, c.column).value
                if isinstance(c.value, str) and c.value.startswith("="):
                    formulas += 1
                if c.value != v2 and not (hasattr(c.value, "text") and hasattr(v2, "text")):
                    diffs += 1
    if diffs:
        issues.append(Issue("integrity", "error", f"{diffs} cell value(s) differ between the original and formatted "
                            "workbook. Do not use the formatted workbook.", status="open"))
    else:
        issues.append(Issue("integrity", "info", f"Excel workbook: every value and all {formulas} formula(s) "
                            "retained unchanged.", status="passed"))
    return issues


def model_to_workbook(model, out, settings, changes):
    """Workbook with one worksheet per financial table (from Word, PDF or text input)."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    st = settings["style"]
    used = set()
    count = 0
    last_title = None
    for b in model.blocks:
        if b.kind == "paragraph" and b.role == "title":
            last_title = b.text.strip()
        if b.kind != "table" or not is_financial_table(b):
            continue
        count += 1
        base = {"SOFP": "SOFP", "SOCI": "SOCI", "SOCE": "SOCE", "SOCF": "SOCF"}.get(b.section)
        if not base:
            base = f"Note {b.note_number}" if b.note_number else f"Table {count}"
        name = base[:28]
        k = 2
        while name in used:
            name = f"{base[:25]} ({k})"
            k += 1
        used.add(name)
        ws = wb.create_sheet(name)
        r0 = 1
        heading = SECTIONS.get(b.section, "") if b.section in PRIMARY else (last_title or "")
        if heading:
            ws.cell(1, 1, heading).font = Font(name=st["font"], size=st["table_size"] + 2, bold=True)
            r0 = 3
        for ri, row in enumerate(b.rows):
            for ci, txt in enumerate(row.cells):
                v = parse_amount(txt) if ci in b.amount_cols and not YEAR.match(txt.strip() or "x") else None
                if row.kind == "header":
                    v = None
                cell = ws.cell(r0 + ri, ci + 1)
                if v is not None and txt.strip():
                    cell.value = float(v) if v != v.to_integral_value() else int(v)
                elif txt:
                    cell.value = txt
        _format_sheet(ws, ws, settings, st, changes)
    if not wb.worksheets:
        ws = wb.create_sheet("No tables")
        ws["A1"] = "No financial tables were found in the source document."
    wb.save(out)
