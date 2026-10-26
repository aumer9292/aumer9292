"""PDF output: LibreOffice renders the formatted Word file exactly; without it,
a simpler PDF is drawn from the formatted content with ReportLab."""

import os
import shutil
import subprocess
import tempfile


def find_office():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(p):
            return p
    return None


def docx_to_pdf(docx_path, pdf_path, model=None, settings=None):
    office = find_office()
    if office:
        out_dir = tempfile.mkdtemp(prefix="fsr_pdf_")
        profile = tempfile.mkdtemp(prefix="fsr_lo_")
        try:
            subprocess.run([office, f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf",
                            "--outdir", out_dir, docx_path], check=True, timeout=240,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            produced = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
            if os.path.exists(produced):
                shutil.move(produced, pdf_path)
                return "LibreOffice"
        except (subprocess.SubprocessError, OSError):
            pass
        finally:
            shutil.rmtree(out_dir, ignore_errors=True)
            shutil.rmtree(profile, ignore_errors=True)
    _reportlab_pdf(docx_path, pdf_path, settings or {})
    return "ReportLab (simplified layout — install LibreOffice for an exact copy of the Word layout)"


def _reportlab_pdf(docx_path, pdf_path, settings):
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
    from reportlab.lib.pagesizes import A4, LETTER
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    from .classify import PRIMARY, classify_document, is_financial_table
    from .extract import read_docx

    model = classify_document(read_docx(docx_path))
    size = 10
    body = ParagraphStyle("b", fontName="Times-Roman", fontSize=size, leading=size * 1.3, alignment=TA_JUSTIFY,
                          spaceAfter=5)
    head = ParagraphStyle("h", parent=body, fontName="Times-Bold", alignment=TA_LEFT, spaceBefore=8, spaceAfter=3)
    title = ParagraphStyle("t", parent=head, fontSize=13, leading=16, alignment=TA_CENTER)
    cellst = ParagraphStyle("c", parent=body, fontSize=size - 1, leading=(size - 1) * 1.2, alignment=TA_LEFT,
                            spaceAfter=0)
    page = LETTER if settings.get("page_size") == "Letter" else A4
    story = []
    prev = None
    for b in model.blocks:
        if b.kind == "paragraph":
            if not b.text.strip():
                continue
            if b.role == "title" and b.section in PRIMARY and b.section != prev and story:
                story.append(PageBreak())
            if b.role == "title":
                prev = b.section
            st = title if b.role == "title" or (b.role in ("period", "unit") and b.section in PRIMARY) else \
                head if b.role in ("heading", "note_heading") else body
            story.append(Paragraph(escape(b.text), st))
        else:
            fin = is_financial_table(b)
            data = []
            for r in b.rows:
                row = []
                for c, txt in enumerate(r.cells):
                    bold = r.kind in ("header", "grand_total", "heading")
                    t = f"<b>{escape(txt)}</b>" if bold and txt else escape(txt)
                    al = ParagraphStyle("x", parent=cellst, alignment=2 if c in b.amount_cols else 1
                                        if c == b.note_col else 0)
                    row.append(Paragraph(t, al))
                data.append(row)
            if not data:
                continue
            avail = page[0] - 4.5 * cm
            w = b.width
            if fin:
                amt = min(2.6 * cm, avail * 0.6 / max(1, len(b.amount_cols)))
                note = 1.2 * cm if b.note_col is not None else 0
                lab = avail - amt * len(b.amount_cols) - note
                widths = [lab if c == 0 else note if c == b.note_col else amt for c in range(w)]
            else:
                widths = [avail / w] * w
            t = Table(data, colWidths=widths, repeatRows=sum(1 for r in b.rows if r.kind == "header"))
            style = [("VALIGN", (0, 0), (-1, -1), "BOTTOM"), ("TOPPADDING", (0, 0), (-1, -1), 1),
                     ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]
            if fin:
                prevk = None
                for i, r in enumerate(b.rows):
                    if r.kind in ("total", "grand_total") and prevk in ("item", "total", "subtotal"):
                        for c in b.amount_cols:
                            style.append(("LINEABOVE", (c, i), (c, i), 0.5, colors.black))
                    if r.kind == "grand_total":
                        for c in b.amount_cols:
                            style.append(("LINEBELOW", (c, i), (c, i), 1.2, colors.black))
                    if r.kind != "blank":
                        prevk = r.kind
            else:
                style.append(("GRID", (0, 0), (-1, -1), 0.4, colors.grey))
            t.setStyle(TableStyle(style))
            story.append(t)
            story.append(Spacer(1, 6))

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Times-Roman", 8)
        canvas.drawCentredString(page[0] / 2, 1.2 * cm, f"Page {doc_.page}")
        if settings.get("entity_name"):
            canvas.drawRightString(page[0] - 2 * cm, page[1] - 1.2 * cm, settings["entity_name"])
        canvas.restoreState()

    doc = SimpleDocTemplate(pdf_path, pagesize=page, leftMargin=2.5 * cm, rightMargin=2 * cm, topMargin=2.5 * cm,
                            bottomMargin=2.5 * cm, title=model.source_name)
    doc.build(story or [Paragraph("(empty)", body)], onFirstPage=footer, onLaterPages=footer)
