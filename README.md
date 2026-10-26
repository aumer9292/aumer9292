# FS Review — financial statements formatting, proofreading and review

FS Review takes a draft set of financial statements and returns:

| Output | What it is |
| --- | --- |
| **Formatted Word document** | The original document reformatted in place. Fonts, headings, spacing, margins, page size, tables, figure presentation, totals lines, page breaks, headers and footers are standardised. Images, fields, footnotes and anything else the tool does not understand stay where they were. |
| **Excel workbook** | For Excel input, the original workbook with every worksheet, formula and value unchanged. Only number formats, fonts, borders, column widths and print settings change. For other inputs, a workbook with one sheet per financial table, figures stored as numbers. |
| **PDF** | The formatted Word document as a PDF. LibreOffice gives an exact copy; without it, a simpler layout is drawn. |
| **Review report** (Word + JSON) | Every finding: accounting cross-checks, casting, comparatives, note references, proofreading, items for manual review, and an audit trail of every text change the tool made. |

It runs on your own computer. Nothing is uploaded anywhere unless you switch on the optional AI review.

## Getting started

1. Install Python 3.10 or later from python.org (tick "Add Python to PATH" on Windows).
2. Optional but recommended: install [LibreOffice](https://www.libreoffice.org/) for exact PDFs and for checking Excel files whose formulas have no saved results.
3. Double-click **Start FS Review.bat** (Windows) or run `./start-fs-review.sh` (macOS/Linux). Your browser opens the tool.
4. Drop in the draft, fill in the entity name, financial year and framework, choose a template, and press **Review and format**.

Command line, for batches or scripts:

```
python -m fsreview "Draft FS 2025.docx" --entity "Northwind Trading Limited" --year 2025 --framework IFRS
python -m fsreview draft.pdf --template modern --page-size Letter --fix-spelling
python -m fsreview draft.docx --reference "Approved FS 2024.docx"     # copy the house style
python -m fsreview --help
```

The command exits with status 1 when the review found errors, so it can be used in a checklist script.

## Inputs

* **Word** (.docx) — formatted in place; the best route.
* **Excel** (.xlsx, .xlsm, .xls) — each worksheet is read; .xls is read through `xlrd` and saved as .xlsx (formulas cannot be carried over from .xls; the report says so).
* **PDF** — text PDFs are rebuilt into Word: paragraphs are rejoined, figure columns are found from their positions, wrapped labels are rejoined, and running headers/footers are dropped from the body (the report lists them). Scanned pages are read with OCR when `pytesseract` and Tesseract are installed; otherwise the page is marked in the document and the report as needing manual work.
* **Text** (.txt) — columns are found from spacing.
* **Reference template** — an approved Word set; its font, sizes and margins become the house style.

Nothing is dropped silently: text boxes, content controls, tracked changes, unreadable pages and formulas without results are all listed as items for manual review.

## What the formatting engine does

It first works out what each part of the document is — statement titles, period lines ("For the year ended…"), unit lines, headings, numbered note headings, body text and contents entries — and, for every table, the label, note and amount columns, the year above each column, and whether each row is a heading, line item, subtotal, total or grand total. Then each part gets its own treatment:

* **Primary statements** start on a new page with a centred title block (entity name, title, period). Amounts are right-aligned in fixed-width columns, the note column is centred, column headings repeat on each page, rows never split across pages, and short statements are kept on one page. Totals get a single rule above; grand totals (total assets, total equity and liabilities, profit for the year, total comprehensive income, closing cash, closing equity) are bold with a double rule below.
* **Figures** are written one way throughout: thousands separators, negatives in brackets (or with a minus sign), nil as a dash, and the same decimal places down each column. Only the presentation changes; each rewritten figure is parsed again and must have the same value, and each change is in the audit trail.
* **Statement of changes in equity** — movements are line items and balances are the totals, with the closing balance double-ruled.
* **Notes** — note headings get one numbering style (`12.` vs `12`), sub-notes keep theirs (`3.1`), headings are kept with the text that follows.
* **Body text** — one font, size, colour, line spacing and paragraph spacing; long paragraphs are justified; double spaces and spaces before punctuation are removed; runs of empty paragraphs are collapsed.
* **Contents** — page numbers on a right tab with dot leaders.
* **Page** — page size, margins, header with the entity name, footer with "Page X of Y".

Templates: Classic (Times New Roman), Modern (Arial), Calibri and Annual report (Garamond), or the house style read from a reference document.

## What the review checks

**Accounting cross-checks**

* Statement of financial position balances (total assets = total equity and liabilities, or equity + liabilities, or net assets = equity), current year and comparative.
* Total comprehensive income = profit + other comprehensive income.
* Cash flows: operating + investing + financing = net change; opening + net change (+ FX) = closing; closing cash agrees to the statement of financial position.
* Statement of changes in equity: components add across to the total column; opening + movements = closing for each column; closing equity agrees to the statement of financial position (current and comparative); profit agrees to the statement of profit or loss.
* Every statement line that cites a note shows a figure that appears in that note.

**Casting** — every subtotal and total is recomputed. The engine understands the usual layouts: a total of the lines above, a total of earlier subtotals (total assets), a running total (gross profit → operating profit → profit before tax → profit for the year), the difference of two totals (net current assets) and lines left without their own subtotal. A failing total shows the figure shown, the figure computed and the year.

**Consistency and comparatives** — comparative column present; first column matches the selected financial year; years consecutive and current year first; blank comparatives; one currency/unit label; one period-end date across the statement headings; duplicate, missing or out-of-order note numbers; references to notes that do not exist; mixed terminology (balance sheet vs statement of financial position, debtors vs receivables, -ise vs -ize, judgement vs judgment, …).

**Proofreading** — a list of misspellings common in financial statements (applied only if you ask, and always listed), repeated words, spacing and punctuation, unmatched brackets, impossible dates (31 June), years after the financial year in headings, and references to superseded standards (IAS 18, IAS 17, IAS 39, IFRS 4, IAS 11 …). Install `pyspellchecker` for a general dictionary check as well.

**AI review (optional)** — with `pip install anthropic` and an `ANTHROPIC_API_KEY`, the narrative paragraphs (policies, notes, reports) are sent to Claude for grammar, clarity and framework-wording suggestions. Suggestions go into the report only; nothing is changed in the document, and Claude is told never to propose changes to figures, dates or note numbers.

**Content integrity** — after formatting, the Word file is read back and compared with the source: every table figure in order, every figure in the text, and every word except the changes recorded in the audit trail. For Excel, every cell value and formula is compared. Any difference is reported as an error.

## Limits worth knowing

* Recognition relies on the usual titles and labels ("Total assets", "Net cash from operating activities", "Balance at …"). Unusual wording may leave a check unperformed; the report then says what it could not verify instead of passing it.
* PDF rebuilding works well for text PDFs with a conventional layout. Always compare the review's figures with the source for scanned or unusual PDFs.
* The tool supports professional judgement; it does not replace review by the preparer and the engagement team.

## For developers

```
pip install -r requirements.txt pytest
python samples/make_samples.py      # deliberately messy sample statements
python -m pytest tests
```

| Module | Role |
| --- | --- |
| `fsreview/extract.py` | Readers for Word, Excel, PDF (with optional OCR) and text into one document model |
| `fsreview/classify.py` | Sections, paragraph roles, table columns and row kinds |
| `fsreview/checks.py` | Casting and accounting cross-checks |
| `fsreview/proofread.py` | Proofreading and consistency |
| `fsreview/docx_engine.py` | Word formatting engine |
| `fsreview/xlsx_engine.py` | Excel formatting and verification |
| `fsreview/pdf_engine.py` | PDF output |
| `fsreview/integrity.py` | Proves the content was preserved |
| `fsreview/report.py` | Review report |
| `fsreview/server.py`, `fsreview/web/` | Local browser interface |
