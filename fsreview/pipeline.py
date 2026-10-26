"""Run the whole review: read, classify, check, proofread, format, verify, report."""

import os

from . import ai_review, checks, docx_engine, integrity, pdf_engine, proofread, report, xlsx_engine
from .classify import classify_document
from .extract import read_document, read_docx
from .model import Issue
from .templates import build_settings, style_from_reference


def run(input_path, out_dir, user_settings=None, reference_path=None, progress=print):
    os.makedirs(out_dir, exist_ok=True)
    ref_style = None
    notes = []
    if reference_path:
        try:
            ref_style = style_from_reference(reference_path)
            notes.append(f"House style taken from the reference document {os.path.basename(reference_path)} "
                         f"({ref_style.get('font', '?')} {ref_style.get('size', '?')}pt).")
        except Exception as e:
            notes.append(f"The reference template could not be read ({e}); the selected template was used.")
    settings = build_settings(user_settings, ref_style)
    wanted = set(settings.get("outputs") or ["docx", "xlsx", "pdf", "report"])
    base = os.path.splitext(os.path.basename(input_path))[0]
    ext = os.path.splitext(input_path)[1].lower()

    progress("Reading document…")
    model = read_document(input_path)
    classify_document(model, settings)
    issues = [Issue("extraction", "warning", w, suggestion="Review by hand.") for w in model.warnings]
    for b in model.blocks:
        if getattr(b, "needs_review", ""):
            issues.append(Issue("extraction", "error", b.needs_review, location=b.source))

    progress("Running accounting cross-checks…")
    issues += checks.run_checks(model, settings)
    progress("Proofreading…")
    issues += proofread.proofread(model, settings)
    if settings.get("ai_review"):
        progress("AI proofreading with Claude…")
        issues += ai_review.ai_proofread(model, settings)

    outputs = {}
    changes = docx_engine.Changes()
    docx_out = os.path.join(out_dir, f"{base} - formatted.docx")
    if wanted & {"docx", "pdf"}:
        progress("Formatting Word document…")
        changes = docx_engine.format_to_file(model, settings, docx_out,
                                             source_docx=input_path if ext in (".docx", ".docm", ".dotx") else None)
        outputs["Word"] = os.path.basename(docx_out)
        progress("Verifying content integrity…")
        after = classify_document(read_docx(docx_out))
        if ext in (".docx", ".docm", ".dotx"):
            issues += integrity.compare(model, after, changes)
        else:
            raw_path = os.path.join(out_dir, ".source-content.docx")
            docx_engine.model_to_docx(model).save(raw_path)
            raw = classify_document(read_docx(raw_path))
            os.remove(raw_path)
            issues += integrity.compare(raw, after, changes)
            issues += integrity.compare_figures(model, raw)

    if "xlsx" in wanted:
        progress("Preparing Excel workbook…")
        xlsx_out = os.path.join(out_dir, f"{base} - formatted.xlsx")
        if ext in (".xlsx", ".xlsm"):
            if ext == ".xlsm":
                xlsx_out = xlsx_out[:-5] + ".xlsm"
            issues += xlsx_engine.format_workbook(input_path, xlsx_out, settings, changes)
        else:
            xlsx_engine.model_to_workbook(model, xlsx_out, settings, changes)
            if ext != ".xls":
                notes.append("The Excel workbook holds the financial tables extracted from the document, with "
                             "figures stored as numbers.")
        outputs["Excel"] = os.path.basename(xlsx_out)

    if "pdf" in wanted:
        progress("Creating PDF…")
        pdf_out = os.path.join(out_dir, f"{base} - formatted.pdf")
        how = pdf_engine.docx_to_pdf(docx_out, pdf_out, settings=settings)
        outputs["PDF"] = os.path.basename(pdf_out)
        notes.append(f"PDF produced with {how}.")
        if "docx" not in wanted:
            os.remove(docx_out)
            outputs.pop("Word", None)

    progress("Writing review report…")
    rep_docx = os.path.join(out_dir, f"{base} - review report.docx")
    rep_json = os.path.join(out_dir, f"{base} - review report.json")
    outputs["Review report"] = os.path.basename(rep_docx)
    data = report.to_json(rep_json, model, settings, issues, changes, outputs, notes)
    report.to_docx(rep_docx, model, settings, issues, changes, outputs, notes)
    progress("Done.")
    return data

