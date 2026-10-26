"""Read Word, Excel, PDF and text files into the document model.

Every reader keeps all content. Where text cannot be read reliably (a scanned
page with no OCR available, a text box, a content control) it is kept or
listed as a warning so the reviewer can check it by hand.
"""

import datetime as _dt
import os
import re

from .classify import classify_title
from .model import Document, Paragraph, Row, Table
from .numbers import YEAR, is_amount

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def read_document(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".docx", ".docm", ".dotx"):
        return read_docx(path)
    if ext in (".xlsx", ".xlsm"):
        return read_xlsx(path)
    if ext == ".xls":
        return read_xls(path)
    if ext == ".pdf":
        return read_pdf(path)
    if ext in (".txt", ".text", ".csv", ".md"):
        return read_txt(path)
    raise ValueError(f"Unsupported file type '{ext}'. Use .docx, .xlsx, .xls, .pdf or .txt.")


# --------------------------------------------------------------------- Word

def docx_body_items(doc):
    """Body paragraphs and tables in document order, as python-docx objects."""
    from docx.table import Table as DTable
    from docx.text.paragraph import Paragraph as DPara
    items = []
    for child in doc.element.body.iterchildren():
        if child.tag == W + "p":
            items.append(DPara(child, doc))
        elif child.tag == W + "tbl":
            items.append(DTable(child, doc))
    return items


def unique_cells(row):
    """Cells of a docx row with horizontally merged cells counted once."""
    seen, out = set(), []
    for c in row.cells:
        if id(c._tc) in seen:
            continue
        seen.add(id(c._tc))
        out.append(c)
    return out


def _para_bold(p):
    runs = [r for r in p.runs if r.text.strip()]
    if not runs:
        return False
    style_bold = bool(p.style is not None and p.style.font is not None and p.style.font.bold)
    return all((r.bold if r.bold is not None else style_bold) for r in runs)


def read_docx(path):
    import docx
    d = docx.Document(path)
    blocks, warnings = [], []
    for i, item in enumerate(docx_body_items(d)):
        if item._element.tag == W + "p":
            blocks.append(Paragraph(text=item.text, bold=_para_bold(item),
                                    style=item.style.name if item.style is not None else "",
                                    source=f"paragraph {i + 1}", index=i))
        else:
            rows = []
            for r in item.rows:
                cells = unique_cells(r)
                texts = [c.text.strip() for c in cells]
                bold = all(_para_bold(p) for c in cells for p in c.paragraphs if p.text.strip()) \
                    and any(texts)
                rows.append(Row(cells=texts, bold=bold))
            blocks.append(Table(rows=rows, source=f"table at item {i + 1}", index=i))
    body = d.element.body
    if body.findall(f".//{W}txbxContent"):
        warnings.append("The document contains text boxes. Their text was kept but not reformatted or "
                        "checked; review them by hand.")
    if body.findall(f"{W}sdt"):
        warnings.append("The document contains content controls (for example a table of contents). "
                        "They were kept unchanged; update fields in Word after formatting.")
    if body.findall(f".//{W}ins") or body.findall(f".//{W}del"):
        warnings.append("The document contains tracked changes. Accept or reject them before final "
                        "review; deleted text is not included in the checks.")
    return Document(blocks=blocks, source_name=os.path.basename(path), source_type="docx",
                    warnings=warnings)


# -------------------------------------------------------------------- Excel

def excel_value_text(value, number_format=""):
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value).upper()
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.strftime("%d %B %Y").lstrip("0")
    if isinstance(value, (int, float)):
        if "%" in (number_format or ""):
            return f"{value * 100:.1f}%"
        if isinstance(value, float) and not value.is_integer():
            dec = 2 if abs(value) < 1e9 else 0
            s = f"{abs(value):,.{dec}f}"
        else:
            s = f"{abs(int(round(value))):,}"
        if YEAR.match(str(value)) or (isinstance(value, int) and 1900 <= value <= 2100
                                       and "," not in (number_format or "")):
            return str(value)
        return f"({s})" if value < 0 else s
    return str(value).strip()


def _sheet_blocks(name, rows_values, source):
    blocks = [Paragraph(text=name, bold=True, source=source, style="Sheet name")]
    rows = [r for r in rows_values]
    # trim trailing empty columns and rows
    width = 0
    for r in rows:
        for j in range(len(r) - 1, -1, -1):
            if r[j] != "":
                width = max(width, j + 1)
                break
    while rows and not any(rows[-1]):
        rows.pop()
    rows = [list(r[:width]) + [""] * (width - len(r[:width])) for r in rows]
    if not rows or width == 0:
        blocks.append(Paragraph(text="(empty sheet)", source=source))
        return blocks
    # Leading rows with a single filled cell in column A are titles, not table rows.
    k = 0
    while k < len(rows) and sum(1 for c in rows[k] if c) <= 1 and not any(is_amount(c) for c in rows[k][1:]):
        if rows[k][0]:
            blocks.append(Paragraph(text=rows[k][0], bold=True, source=f"{source} row {k + 1}"))
        k += 1
        if k >= 6:
            break
    if k < len(rows):
        blocks.append(Table(rows=[Row(cells=r) for r in rows[k:]], source=source))
    return blocks


def _recalculated_copy(path):
    """Workbooks saved by tools other than Excel often hold formulas without results.
    LibreOffice, when installed, recalculates a temporary copy so the values can be checked."""
    import shutil
    import subprocess
    import tempfile
    from .pdf_engine import find_office
    office = find_office()
    if not office:
        return None
    out_dir = tempfile.mkdtemp(prefix="fsr_calc_")
    profile = tempfile.mkdtemp(prefix="fsr_lo_")
    try:
        subprocess.run([office, f"-env:UserInstallation=file://{profile}", "--headless", "--calc",
                        "--convert-to", "xlsx", "--outdir", out_dir, path], check=True, timeout=180,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        out = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + ".xlsx")
        return out if os.path.exists(out) else None
    except (subprocess.SubprocessError, OSError):
        return None
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def read_xlsx(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    blocks, warnings = [], []
    formula_wb = openpyxl.load_workbook(path, data_only=False)
    missing_results = any(
        isinstance(c.value, str) and c.value.startswith("=") and wb[ws.title].cell(c.row, c.column).value is None
        for ws in formula_wb.worksheets for row in ws.iter_rows() for c in row)
    if missing_results:
        calc = _recalculated_copy(path)
        if calc:
            wb = openpyxl.load_workbook(calc, data_only=True)
            warnings.append("Some formulas had no saved results; the workbook was recalculated with LibreOffice "
                            "for checking. The formatted workbook keeps the original formulas.")
    for ws in wb.worksheets:
        fws = formula_wb[ws.title]
        rows, missing = [], 0
        for r_idx, row in enumerate(ws.iter_rows(), start=1):
            vals = []
            for c in row:
                v = c.value
                if v is None:
                    f = fws.cell(row=c.row, column=c.column).value
                    if isinstance(f, str) and f.startswith("="):
                        missing += 1
                vals.append(excel_value_text(v, c.number_format))
            rows.append(vals)
        if missing:
            warnings.append(f"Sheet '{ws.title}': {missing} formula cell(s) have no saved result, so their "
                            "values could not be checked. Open and save the workbook in Excel, then run again.")
        blocks.extend(_sheet_blocks(ws.title, rows, f"sheet '{ws.title}'"))
    return Document(blocks=blocks, source_name=os.path.basename(path), source_type="xlsx",
                    warnings=warnings)


def read_xls(path):
    try:
        import xlrd
    except ImportError as e:
        raise RuntimeError("Reading .xls needs the 'xlrd' package (pip install xlrd).") from e
    book = xlrd.open_workbook(path)
    blocks = []
    for sh in book.sheets():
        rows = []
        for r in range(sh.nrows):
            vals = []
            for c in range(sh.ncols):
                cell = sh.cell(r, c)
                v = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    v = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                elif cell.ctype == xlrd.XL_CELL_EMPTY:
                    v = None
                elif cell.ctype == xlrd.XL_CELL_NUMBER and float(v).is_integer():
                    v = int(v)
                vals.append(excel_value_text(v))
            rows.append(vals)
        blocks.extend(_sheet_blocks(sh.name, rows, f"sheet '{sh.name}'"))
    return Document(blocks=blocks, source_name=os.path.basename(path), source_type="xls",
                    warnings=["The .xls format does not keep formulas when converted. The formatted "
                              "workbook is saved as .xlsx with values only; check any formulas by hand."])


# ------------------------------------------------- line-based (PDF and text)

class Seg:
    __slots__ = ("text", "x0", "x1", "bold")

    def __init__(self, text, x0, x1, bold=False):
        self.text, self.x0, self.x1, self.bold = text, x0, x1, bold


class Line:
    def __init__(self, segs, y, size, page, width):
        self.segs, self.y, self.size, self.page, self.width = segs, y, size, page, width

    @property
    def text(self):
        return " ".join(s.text for s in self.segs)

    @property
    def bold(self):
        return bool(self.segs) and all(s.bold for s in self.segs)


CONNECTOR = re.compile(r"(?:\b(?:and|of|in|from|to|for|on|at|the|by|with|or|less|net)|[,&/-])$", re.I)


UNIT_TOKEN = re.compile(r"^(?:[$€£₹]|Rs\.?|[A-Z]{3})?\s?['’]?000$|^(?:[$€£₹]|Rs\.?|[A-Z]{3})$")


def _is_fig(text):
    t = text.strip()
    return is_amount(t) or bool(YEAR.match(t)) or bool(UNIT_TOKEN.match(t))


def _is_column_heading(line):
    """A short line of column captions to the right ('Share capital  Retained earnings  Total')."""
    segs = line.segs
    return 2 <= len(segs) <= 8 and len(line.text) < 90 and \
        sum(1 for s in segs if s.x0 > line.width * 0.3) >= 2 and segs[0].x0 > line.width * 0.22


def _split_numeric_tail(line):
    """Split a line into label segments and the figure segments on its right."""
    segs = line.segs
    nums = []
    i = len(segs) - 1
    while i >= 0:
        if _is_fig(segs[i].text) and segs[i].x0 > line.width * 0.22 and (i > 0 or segs[i].x0 > line.width * 0.3):
            nums.insert(0, segs[i])
            i -= 1
            continue
        break
    return segs[:i + 1], nums


def _is_table_line(line):
    label, nums = _split_numeric_tail(line)
    if not nums:
        return False
    real = [s for s in nums if is_amount(s.text) and not YEAR.match(s.text.strip())]
    heads = [s for s in nums if YEAR.match(s.text.strip()) or UNIT_TOKEN.match(s.text.strip())]
    if not real and len(heads) < 2:
        return False
    # a single small number after a long label is a contents page number, not a figure
    if len(nums) == 1 and label and re.fullmatch(r"\d{1,3}", nums[0].text.strip()) and \
            classify_hint(" ".join(s.text for s in label)):
        return False
    return not label or label[-1].x1 < nums[0].x0


def classify_hint(text):
    return bool(re.search(r"statement|report|notes|contents|policies", text, re.I))


def _continues(prev_label, line):
    """Is this label-only line the wrapped end of the previous row's label?"""
    t = line.text.strip()
    return bool(t) and (t[0].islower() or bool(CONNECTOR.search(prev_label.strip()))
                        or bool(re.fullmatch(r"(?:\d{1,2}\s+)?(?:[A-Z][a-z]+\s+)?(?:19|20)\d{2}\)?", t)))


def _cluster(items, tol):
    """items: (x1, x0) pairs -> clusters as dicts with centre, x0, x1."""
    out = []
    for x1, x0 in sorted(items):
        if out and abs(x1 - out[-1]["c"]) <= tol:
            c = out[-1]
            c["xs"].append(x1)
            c["c"] = sum(c["xs"]) / len(c["xs"])
            c["x0"] = min(c["x0"], x0)
            c["x1"] = max(c["x1"], x1)
        else:
            out.append({"xs": [x1], "c": x1, "x0": x0, "x1": x1})
    return out


def lines_to_blocks(lines, source_prefix):
    blocks = []
    i, n = 0, len(lines)
    para = None
    while i < n:
        ln = lines[i]
        if _is_table_line(ln) or (_is_column_heading(ln) and any(_is_table_line(x) for x in lines[i + 1:i + 3])):
            j, region = i, []
            while j < n:
                cur = lines[j]
                if _is_table_line(cur) or (j == i):
                    region.append(cur)
                    j += 1
                    continue
                look = []
                for x in lines[j:j + 4]:
                    if x.page != cur.page or classify_title(x.text):
                        break
                    look.append(x)
                look = look[1:] if look and look[0] is cur else []
                same_page = not region or cur.page == region[-1].page
                prev_label = " ".join(s.text for s in _split_numeric_tail(region[-1])[0]) if region else ""
                if len(cur.text) < 90 and ((same_page and any(_is_table_line(x) for x in look)) or
                                           (region and _continues(prev_label, cur) and
                                            ((cur.page == region[-1].page and cur.y - region[-1].y < cur.size * 2.2)
                                             or (cur.page == region[-1].page + 1 and j > 0
                                                 and lines[j - 1] is region[-1])))):
                    region.append(cur)
                    j += 1
                    continue
                break
            if para:
                blocks.append(para)
                para = None
            blocks.append(_region_to_table(region, source_prefix))
            i = j
            continue
        text = ln.text.strip()
        if not text:
            i += 1
            continue
        prev = lines[i - 1] if i else None
        joins = (para is not None and prev is not None and prev.page == ln.page
                 and not prev.bold and not ln.bold
                 and abs(ln.size - prev.size) < 0.6
                 and 0 < ln.y - prev.y < prev.size * 1.9
                 and prev.segs and prev.segs[-1].x1 > prev.width * 0.62
                 and not re.match(r"^(\d+(\.\d+)*\.?|[a-z]\)|\([a-z0-9]+\)|[•\-–])\s", text))
        if joins:
            para.text += " " + text
        else:
            if para:
                blocks.append(para)
            para = Paragraph(text=text, bold=ln.bold, source=f"{source_prefix} {ln.page}")
        i += 1
    if para:
        blocks.append(para)
    return blocks


def _region_to_table(region, source_prefix):
    figs = []
    for ln in region:
        _, nums = _split_numeric_tail(ln)
        figs.extend((s.x1, s.x0) for s in nums)
    span = max((ln.width for ln in region), default=600)
    cols = _cluster(figs, tol=span * 0.035)
    label_edge = min((c["x0"] for c in cols), default=span) - 2

    def col_for(seg):
        best, dist = None, None
        for k, c in enumerate(cols):
            if seg.x1 >= c["x0"] - 4 and seg.x0 <= c["x1"] + 4:
                d = abs((seg.x0 + seg.x1) / 2 - (c["x0"] + c["x1"]) / 2)
            else:
                d = min(abs(seg.x1 - c["c"]), abs(seg.x0 - c["x0"])) + 1000
            if dist is None or d < dist:
                best, dist = k, d
        return best if dist is not None and dist < 1000 + span * 0.05 else None

    rows = []
    for ln in region:
        label, nums = _split_numeric_tail(ln)
        cells = [""] * (len(cols) + 1)
        lab = []
        for s in label:
            k = col_for(s) if s.x0 >= label_edge - 4 else None
            if k is None:
                lab.append(s.text)
            else:
                cells[k + 1] = (cells[k + 1] + " " + s.text).strip()
        for s in nums:
            k = min(range(len(cols)), key=lambda c: abs(cols[c]["c"] - s.x1))
            cells[k + 1] = (cells[k + 1] + " " + s.text).strip()
        cells[0] = " ".join(lab)
        has_fig = bool(nums)
        if rows and not has_fig and not any(cells[1:]) and _continues(rows[-1][0].cells[0], ln):
            rows[-1][0].cells[0] += " " + cells[0]
            continue
        if rows and has_fig and rows[-1][1] is False and not any(rows[-1][0].cells[1:]) and \
                rows[-1][0].cells[0] and CONNECTOR.search(rows[-1][0].cells[0]):
            prev = rows.pop()[0]
            cells[0] = (prev.cells[0] + " " + cells[0]).strip()
        rows.append((Row(cells=cells, bold=ln.bold), has_fig))
    pages = sorted({ln.page for ln in region})
    return Table(rows=[r for r, _ in rows],
                 source=f"{source_prefix} {pages[0]}" + (f"-{pages[-1]}" if len(pages) > 1 else ""))


def _words_to_lines(words, page_no, width):
    """words: dicts with text, x0, x1, top, size, bold."""
    if not words:
        return []
    words = sorted(words, key=lambda w: (round(w["top"]), w["x0"]))
    lines, cur = [], []
    for w in words:
        if cur and abs(w["top"] - cur[0]["top"]) > max(2.5, cur[0]["size"] * 0.45):
            lines.append(cur)
            cur = []
        cur.append(w)
    if cur:
        lines.append(cur)
    out = []
    for lw in lines:
        lw.sort(key=lambda w: w["x0"])
        size = sorted(w["size"] for w in lw)[len(lw) // 2]
        segs = []
        for w in lw:
            gap = w["x0"] - segs[-1].x1 if segs else None
            if segs and gap < size * 0.9 and not (_is_fig(w["text"]) and gap > size * 0.5):
                segs[-1].text += " " + w["text"]
                segs[-1].x1 = w["x1"]
                segs[-1].bold = segs[-1].bold and w["bold"]
            else:
                segs.append(Seg(w["text"], w["x0"], w["x1"], w["bold"]))
        out.append(Line(segs, lw[0]["top"], size, page_no, width))
    return out


def _running_lines(pages):
    """Text repeated at the top or bottom of most pages: running headers and footers."""
    if len(pages) < 3:
        return set()
    counts = {}
    for lines, height in pages:
        seen = set()
        for ln in lines:
            if ln.y < height * 0.09 or ln.y > height * 0.9:
                key = re.sub(r"\d+", "#", ln.text.strip().lower())
                if key and key not in seen:
                    seen.add(key)
                    counts[key] = counts.get(key, 0) + 1
    return {k for k, v in counts.items() if v >= max(3, len(pages) * 0.5)}


def read_pdf(path):
    import pdfplumber
    pages, warnings = [], []
    with pdfplumber.open(path) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(extra_attrs=["fontname", "size"], keep_blank_chars=False,
                                       use_text_flow=False)
            ws = [{"text": w["text"], "x0": w["x0"], "x1": w["x1"], "top": w["top"],
                   "size": w.get("size", 10), "bold": bool(re.search(r"bold|black|heavy|semibold",
                                                                      w.get("fontname", ""), re.I))}
                  for w in words]
            if len(ws) < 5:
                ocr = _ocr_page(path, pno - 1)
                if ocr is None:
                    msg = (f"Page {pno} has no readable text (it looks scanned) and OCR is not installed. "
                           "Its content is NOT included; install Tesseract and pytesseract or supply "
                           "the Word/Excel version.")
                    warnings.append(msg)
                    pages.append((None, pno, msg))
                    continue
                ws = ocr
                warnings.append(f"Page {pno} was read with OCR. Check its figures against the original.")
            pages.append((_words_to_lines(ws, pno, float(page.width)), float(page.height), None))
    running = _running_lines([(ls, h) for ls, h, _ in pages if ls is not None])
    dropped = set()
    blocks, chunk = [], []
    for lines, h, msg in pages:
        if lines is None:
            blocks.extend(lines_to_blocks(chunk, "page"))
            chunk = []
            blocks.append(Paragraph(text=f"[Page {h}: content could not be extracted — manual review required]",
                                    source=f"page {h}", needs_review=msg))
            continue
        for ln in lines:
            key = re.sub(r"\d+", "#", ln.text.strip().lower())
            if key in running and (ln.y < h * 0.09 or ln.y > h * 0.9):
                dropped.add(ln.text.strip())
                continue
            chunk.append(ln)
    blocks.extend(lines_to_blocks(chunk, "page"))
    if dropped:
        warnings.append("Running headers/footers repeated on every page were not copied into the body (the "
                        "formatted document has its own): " + "; ".join(sorted(dropped)[:4])
                        + ("…" if len(dropped) > 4 else ""))
    return Document(blocks=blocks, source_name=os.path.basename(path), source_type="pdf",
                    warnings=warnings)


def _ocr_page(path, index):
    try:
        import pypdfium2 as pdfium
        import pytesseract
    except ImportError:
        return None
    try:
        pdf = pdfium.PdfDocument(path)
        scale = 300 / 72
        img = pdf[index].render(scale=scale).to_pil()
        data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
    except Exception:
        return None
    out = []
    for k, text in enumerate(data["text"]):
        if not text.strip() or float(data["conf"][k]) < 0:
            continue
        out.append({"text": text, "x0": data["left"][k] / scale,
                    "x1": (data["left"][k] + data["width"][k]) / scale,
                    "top": data["top"][k] / scale, "size": data["height"][k] / scale, "bold": False})
    return out


def read_txt(path):
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    text = text.replace("\t", "    ")
    raw_lines = text.splitlines()
    width = max((len(l.rstrip()) for l in raw_lines), default=80)
    lines, y = [], 0
    for ln in raw_lines:
        y += 12
        segs = []
        for m in re.finditer(r"\S+(?: \S+)*", ln):
            segs.append(Seg(m.group(0), m.start(), m.end()))
        fixed = []
        for s in segs:   # "Revenue 5 1,200" -> separate figure tokens
            parts = s.text.split(" ")
            k = len(parts)
            while k > 1 and is_amount(parts[k - 1]):
                k -= 1
            if k < len(parts) and k > 0:
                head = " ".join(parts[:k])
                fixed.append(Seg(head, s.x0, s.x0 + len(head)))
                x = s.x0 + len(head) + 1
                for p in parts[k:]:
                    fixed.append(Seg(p, x, x + len(p)))
                    x += len(p) + 1
            else:
                fixed.append(s)
        is_caps = ln.strip().isupper() and len(ln.strip()) > 3
        for s in fixed:
            s.bold = is_caps
        lines.append(Line(fixed, y, 10, 1, max(width, 60)))
    # blank lines end paragraphs in text files
    blocks, chunk = [], []
    for ln in lines:
        if not ln.segs:
            if chunk:
                blocks.extend(lines_to_blocks(chunk, "line"))
                chunk = []
            continue
        chunk.append(ln)
    blocks.extend(lines_to_blocks(chunk, "line"))
    for b in blocks:
        b.source = b.source.replace("line 1", "text")
    return Document(blocks=blocks, source_name=os.path.basename(path), source_type="txt")
