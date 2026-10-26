import json
import os
import sys
from decimal import Decimal

import docx
import openpyxl
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "samples"))

import make_samples  # noqa: E402
from fsreview import checks, pipeline  # noqa: E402
from fsreview.classify import analyse_table, classify_title  # noqa: E402
from fsreview.extract import read_docx  # noqa: E402
from fsreview.model import Document, Row, Table  # noqa: E402
from fsreview.numbers import format_amount, parse_amount  # noqa: E402
from fsreview.pdf_engine import find_office  # noqa: E402

FAST = {"financial_year": "2025", "entity_name": "Northwind Trading Limited", "outputs": ["docx", "xlsx", "report"]}


@pytest.fixture(scope="module")
def samples(tmp_path_factory):
    d = tmp_path_factory.mktemp("samples")
    make_samples.make_docx(str(d / "fs.docx"))
    make_samples.make_xlsx(str(d / "fs.xlsx"))
    make_samples.make_txt(str(d / "fs.txt"))
    return d


@pytest.fixture(scope="module")
def docx_run(samples, tmp_path_factory):
    out = tmp_path_factory.mktemp("out")
    data = pipeline.run(str(samples / "fs.docx"), str(out), dict(FAST), progress=lambda m: None)
    return data, out


def messages(data, category=None, status=None):
    return [i["message"] for i in data["issues"]
            if (category is None or i["category"] == category) and (status is None or i["status"] == status)]


# ------------------------------------------------------------------ numbers

@pytest.mark.parametrize("text,value", [
    ("1,234", Decimal(1234)), ("(1,234)", Decimal(-1234)), ("-1,234", Decimal(-1234)), ("1234.50", Decimal("1234.50")),
    ("-", Decimal(0)), ("nil", Decimal(0)), ("–", Decimal(0)), ("$ 500", Decimal(500)), ("(500", None),
    ("12.5%", None), ("Revenue", None), ("", None),
])
def test_parse_amount(text, value):
    assert parse_amount(text) == value


def test_format_amount():
    assert format_amount(Decimal(-1234)) == "(1,234)"
    assert format_amount(Decimal(-1234), negative="minus") == "-1,234"
    assert format_amount(Decimal(0)) == "-"
    assert format_amount(Decimal("1234.5"), 2) == "1,234.50"


def test_titles_and_contents_entries():
    assert classify_title("STATEMENT OF FINANCIAL POSITION") == "SOFP"
    assert classify_title("Statement of Profit or Loss and Other Comprehensive Income") == "SOCI"
    assert classify_title("Statement of cash flows\t6") == "CONTENTS_ENTRY"
    assert classify_title("Statement of financial position .......... 3") == "CONTENTS_ENTRY"
    assert classify_title("Revenue increased in note 12") is None


# ------------------------------------------------------------------ casting

def _table(rows, section="SOFP"):
    t = Table(rows=[Row(cells=list(r)) for r in rows], section=section)
    return analyse_table(t)


def test_casting_finds_wrong_total():
    t = _table([("", "2025", "2024"), ("Inventories", "100", "90"), ("Cash", "50", "40"), ("Total current assets", "160", "130")])
    doc = Document(blocks=[t])
    found = checks.cast_tables(doc)
    errs = [i for i in found if i.severity == "error"]
    assert len(errs) == 1 and "150" in errs[0].message and "2025" in errs[0].message


def test_casting_accepts_layered_totals():
    t = _table([("", "2025"), ("PPE", "100"), ("", "100"), ("Inventory", "30"), ("Cash", "20"), ("", "50"),
                ("Total assets", "150"), ("Share capital", "100"), ("Total equity", "100"), ("Borrowings", "20"),
                ("Current liabilities", ""), ("Payables", "30"), ("", "30"), ("Total liabilities", "50"),
                ("Total equity and liabilities", "150")])
    found = checks.cast_tables(Document(blocks=[t]))
    assert not [i for i in found if i.severity == "error"], [i.message for i in found]


def test_running_totals_in_profit_or_loss():
    t = _table([("", "2025"), ("Revenue", "1,000"), ("Cost of sales", "(600)"), ("Gross profit", "400"),
                ("Expenses", "(150)"), ("Operating profit", "250"), ("Tax", "(50)"), ("Profit for the year", "200")],
               section="SOCI")
    found = checks.cast_tables(Document(blocks=[t]))
    assert not [i for i in found if i.severity == "error"]


# ------------------------------------------------------------- full review

def test_word_review_finds_planted_problems(docx_run):
    data, _ = docx_run
    msgs = " | ".join(messages(data))
    assert "refers to note 11, but there is no note 11" in msgs
    assert "Invalid date: '31 June 2025'" in msgs
    assert "IAS 18" in msgs
    assert "'recievables' → 'receivables'" in msgs
    assert "Repeated word: 'the the'" in msgs
    assert "Different period-end dates" in msgs
    assert "balance sheet" in msgs.lower()
    assert not messages(data, "casting") or all("cast" in m for m in messages(data, "casting"))


def test_word_review_cross_checks_pass(docx_run):
    data, _ = docx_run
    passed = " | ".join(messages(data, status="passed"))
    assert "total assets vs total equity and liabilities: agrees at 22,000" in passed
    assert "Closing total equity in the statement of changes in equity vs the statement of financial position" in passed
    assert "Profit for the year in the statement of changes in equity vs the statement of profit or loss" in passed
    assert "Opening cash + net change = closing cash (current year): agrees" in passed


def test_sections_identified(docx_run):
    data, _ = docx_run
    for s in ("Statement of Financial Position", "Statement of Profit or Loss and Other Comprehensive Income",
              "Statement of Changes in Equity", "Statement of Cash Flows", "Notes to the Financial Statements",
              "Significant Accounting Policies", "Directors' / Management Report"):
        assert s in data["sections"]


def test_formatting_preserves_content(docx_run, samples):
    data, out = docx_run
    integrity = [i for i in data["issues"] if i["category"] == "integrity"]
    assert integrity and all(i["status"] == "passed" for i in integrity), integrity
    before = read_docx(str(samples / "fs.docx"))
    after = read_docx(str(out / "fs - formatted.docx"))
    vals = lambda m: [parse_amount(c) for t in m.tables() for r in t.rows for c in r.cells if parse_amount(c) is not None]
    assert vals(before) == vals(after)


def test_figures_presented_consistently(docx_run):
    data, out = docx_run
    edits = {(e["before"], e["after"]) for e in data["changes"]["edits"]}
    assert ("12500", "12,500") in edits
    assert ("-5,200", "(5,200)") in edits
    d = docx.Document(str(out / "fs - formatted.docx"))
    sec = d.sections[0]
    assert round(sec.page_width.cm, 1) == 21.0
    assert any("PAGE" in (el.text or "") for el in sec.footer._element.iter(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText"))


def test_spelling_fix_option(samples, tmp_path):
    s = dict(FAST, apply_spelling=True, outputs=["docx", "report"])
    data = pipeline.run(str(samples / "fs.docx"), str(tmp_path), s, progress=lambda m: None)
    text = "\n".join(p.text for p in docx.Document(str(tmp_path / "fs - formatted.docx")).paragraphs)
    assert "depreciaton" not in text and "accumulated depreciation" in text
    assert all(i["status"] == "passed" for i in data["issues"] if i["category"] == "integrity")
    assert any(e["reason"] == "Spelling corrected" for e in data["changes"]["edits"])


def test_casting_error_in_word(samples, tmp_path):
    src = tmp_path / "bad.docx"
    d = docx.Document(str(samples / "fs.docx"))
    for t in d.tables:
        for row in t.rows:
            if row.cells[0].text.strip() == "Total assets":
                row.cells[2].paragraphs[0].runs[0].text = "22,100"
    d.save(str(src))
    data = pipeline.run(str(src), str(tmp_path / "o"), dict(FAST, outputs=["report"]), progress=lambda m: None)
    errs = [i["message"] for i in data["issues"] if i["severity"] == "error"]
    assert any("Total assets" in m and "22,100" in m for m in errs)
    assert any("total assets vs total equity and liabilities" in m for m in errs)


def test_excel_keeps_formulas(samples, tmp_path):
    data = pipeline.run(str(samples / "fs.xlsx"), str(tmp_path), dict(FAST), progress=lambda m: None)
    wb = openpyxl.load_workbook(str(tmp_path / "fs - formatted.xlsx"))
    assert wb["SOFP"]["C10"].value == "=SUM(C8:C9)"
    assert wb["SOFP"]["C16"].value == "=C10+C15"
    assert "(#,##0)" in wb["SOFP"]["C8"].number_format
    assert any("formula(s) retained unchanged" in m for m in messages(data, "integrity", "passed"))


def test_text_input(samples, tmp_path):
    data = pipeline.run(str(samples / "fs.txt"), str(tmp_path), dict(FAST), progress=lambda m: None)
    assert "Statement of Financial Position" in data["sections"]
    assert any("agrees at 22,000" in m for m in messages(data, "accounting", "passed"))
    assert all(i["status"] == "passed" for i in data["issues"] if i["category"] == "integrity")


@pytest.mark.skipif(not find_office(), reason="LibreOffice is needed to make a PDF test input")
def test_pdf_input(samples, tmp_path):
    import subprocess
    subprocess.run([find_office(), "--headless", "--convert-to", "pdf", "--outdir", str(tmp_path),
                    str(samples / "fs.docx")], check=True, capture_output=True, timeout=180)
    data = pipeline.run(str(tmp_path / "fs.pdf"), str(tmp_path / "o"), dict(FAST), progress=lambda m: None)
    msgs = " | ".join(messages(data))
    assert "no note 11" in msgs and "31 June 2025" in msgs
    assert not [i for i in data["issues"] if i["category"] == "casting" and i["severity"] == "error"]
    assert any("agrees at 14,600" in m for m in messages(data, "accounting", "passed"))


def test_report_files(docx_run):
    data, out = docx_run
    assert (out / "fs - review report.docx").exists()
    j = json.loads((out / "fs - review report.json").read_text(encoding="utf-8"))
    assert j["summary"]["error"] == data["summary"]["error"] > 0
