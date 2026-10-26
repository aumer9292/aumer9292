"""Run the whole review: read, classify, check, proofread, format, verify, report."""

import os
import re

from . import (ai_review, analysis, checks, disclosures, docx_engine, entity, integrity, library, pdf_engine,
               proofread, report, xlsx_engine)
from .classify import PRIMARY, classify_document
from .figures import extract_figures
from .extract import read_document, read_docx
from .model import Issue
from .templates import build_settings, style_from_reference


def analyse(input_path, user_settings=None):
    """First look at a draft: what it contains and which category the entity seems to be.
    The user confirms the category before the full review runs."""
    settings = build_settings(user_settings)
    model = classify_document(read_document(input_path), settings)
    figs = extract_figures(model)
    out = entity.suggest(model, figs, settings)
    out["sections"] = report.sections_found(model)
    out["entity_name"] = _entity_name(model)
    out["years"] = list(_years(model, settings))
    return out


def _entity_name(model):
    for b in model.blocks[:15]:
        if b.kind == "paragraph" and re.search(r"\b(limited|ltd\.?|company|corporation|inc\.?|llc|plc)\b", b.text, re.I) \
                and len(b.text) < 90:
            return b.text.strip().title() if b.text.isupper() else b.text.strip()
    return ""


def _years(model, settings):
    for t in model.tables():
        if t.section in PRIMARY and t.column_years:
            ys = [t.column_years.get(c) for c in t.amount_cols if t.column_years.get(c)]
            if len(ys) >= 2:
                return str(ys[0]), str(ys[1])
    fy = settings.get("financial_year")
    if fy and str(fy).isdigit():
        return str(fy), str(int(fy) - 1)
    return "Current year", "Prior year"


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
    category = (settings.get("category") or "").upper()
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

    if category not in entity.CATEGORIES:
        sug = entity.suggest(model, extract_figures(model), settings)
        category = sug["suggested"]
        issues.append(Issue("disclosure", "warning", f"Entity category was not confirmed; the review assumed "
                            f"{entity.CATEGORIES[category]['name']} ({category}). " + " ".join(sug["reasons"]),
                            suggestion="Confirm the category so the right disclosure checklist is applied."))
    settings["category"] = category
    cat = entity.CATEGORIES[category]
    if not (user_settings or {}).get("framework_override"):
        settings["framework"] = cat["framework"]
    notes.insert(0, f"Entity category: {cat['name']} ({category}) — framework: {cat['framework']}; "
                    f"disclosures: {cat['schedule']}.")

    progress("Checking required disclosures…")
    checklist, d_issues = disclosures.run_checklist(model, category, settings)
    issues += d_issues
    ref_rows, r_issues = library.compare(model, category)
    issues += r_issues

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

    progress("Analysing the figures…")
    years = _years(model, settings)
    result = analysis.compute(model)
    chart_dir = os.path.join(out_dir, "charts")
    os.makedirs(chart_dir, exist_ok=True)
    chart_files = analysis.charts(result, chart_dir, years)
    an_docx = os.path.join(out_dir, f"{base} - financial analysis.docx")
    if result["ratios"] or result["horizontal"]:
        analysis.write_report(an_docx, result, chart_files, settings, years)
        outputs["Financial analysis"] = os.path.basename(an_docx)
        if "Excel" in outputs:
            analysis.add_to_workbook(os.path.join(out_dir, outputs["Excel"]), result, years)
    else:
        notes.append("Financial analysis skipped: the primary statements' figures could not be identified.")

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
    extra = {"category": {"code": category, **cat}, "checklist": checklist, "reference_comparison": ref_rows,
             "analysis": result, "years": list(years),
             "charts": [{"title": t, "file": "charts/" + os.path.basename(p)} for t, p in chart_files]}
    data = report.to_json(rep_json, model, settings, issues, changes, outputs, notes, extra)
    report.to_docx(rep_docx, model, settings, issues, changes, outputs, notes, extra)
    progress("Done.")
    return data

