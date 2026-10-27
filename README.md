# FS Review — financial statements formatting, disclosure check, analysis and review

FS Review takes a draft set of financial statements and, after asking you which category the entity falls in,
returns:

| Output | What it is |
| --- | --- |
| **Formatted Word document** | The draft laid out for printing in Big 4 style (or another template): each statement on its own page with the entity name, title and period; `------ Rupees ------` over the figure columns; totals ruled; notes with hanging numbers; signature blocks; header, footer and page numbers. Images, fields and footnotes stay where they were. |
| **Excel workbook** | For Excel input, the original workbook with every worksheet, formula and value unchanged (only formatting changes). For other inputs, one sheet per financial table. Both get **Ratios** and **Analysis** sheets. |
| **PDF** | The formatted Word document as a PDF (exact with LibreOffice installed). |
| **Financial analysis** (Word) | Charts of results, margins, financial position, cash flows, liquidity/leverage and working-capital days; ratio analysis with formulas and observations; growth; horizontal and vertical analysis with large movements highlighted. |
| **Review report** (Word + JSON) | The entity category and framework, the disclosure checklist, the comparison with your reference statements, accounting cross-checks, casting, comparatives, note references, proofreading, items for manual review, and an audit trail of every text change. |

It runs on your own computer. Nothing is uploaded anywhere unless you switch on the optional AI review.

## How a review runs

1. **Upload** the draft (Word, Excel, PDF or text) and fill in the financial year and layout choices.
2. **Confirm the entity.** The tool reads turnover, paid-up capital, employees and listing status from the
   draft, suggests a category under the Third Schedule to the Companies Act 2017 and **asks you to confirm**:

   | Category | Framework | Disclosures |
   | --- | --- | --- |
   | Public Interest Company (PIC) | IFRS | Fourth Schedule |
   | Large Sized Company (LSC) | IFRS | Fifth Schedule |
   | Medium Sized Company (MSC) | IFRS for SMEs | Fifth Schedule |
   | Small Sized Company (SSC) | Revised AFRS for SSEs (ICAP) | Fifth Schedule |

   Default thresholds: large at or above Rs 200m paid-up capital, Rs 1bn turnover or 750 employees; small within
   Rs 10m capital, Rs 100m turnover and 250 employees (preceding year's figures). **Check these against the
   current SECP notifications**; if they have changed, put the new figures in `~/.fsreview/thresholds.json`
   (same shape as `DEFAULT_THRESHOLDS` in `fsreview/entity.py`).
3. **Review.** Formatting, disclosure checklist, reference comparison, cross-checks, proofreading and analysis run
   together; results appear in tabs (Findings, Disclosure checklist, Financial analysis, Formatting applied) with
   all files to download.

## Are the relevant notes included?

* **Built-in checklist per category** (`fsreview/disclosures.py`), about 60 items: general information, basis of
  preparation, policies and estimates; a note for every significant line in the statements (property, plant and
  equipment, stock-in-trade, trade debts, long term financing, deferred tax, staff retirement benefits …);
  revenue, cost of sales, expenses, finance cost and taxation (with the tax reconciliation for PIC/LSC);
  related parties; remuneration of chief executive, directors and executives; number of employees; provident
  fund; plant capacity; financial instruments, risk management, fair value and capital management (PIC/LSC);
  EPS, segments and Shariah disclosures (listed); events after the reporting period; date of authorisation.
  Line-driven items are marked *Not applicable* when the related line is absent. Each item shows its reference
  and where it was found. The checklist confirms a disclosure is present; its content still needs reading.
* **Significant lines with no note reference** are flagged.
* **Reference library.** Upload the ICAP illustrative financial statements for each category (PIC, LSC, MSC,
  SSC) and any model set your firm follows (for example a published set you consider best practice). They are
  kept in `~/.fsreview/library` and every review of that category lists the reference's notes that the draft
  does not have, matching by meaning ("Trade debts" = "Trade and other receivables"). A Word reference can also
  be chosen as the **house style** for formatting.

  From the browser: *Reference library → Add to library*. From the command line:
  `python -m fsreview --add-reference "ICAP illustrative MSC.pdf" --ref-category MSC`.

## Getting started

1. Install Python 3.10 or later from python.org (tick "Add Python to PATH" on Windows).
2. Optional but recommended: install [LibreOffice](https://www.libreoffice.org/) for exact PDFs and for checking Excel files whose formulas have no saved results.
3. Double-click **Start FS Review.bat** (Windows) or run `./start-fs-review.sh` (macOS/Linux). Your browser opens the tool.
4. Drop in the draft, fill in the financial year, press **Next**, confirm the entity category, and press **Confirm and review**.

Command line, for batches or scripts:

```
python -m fsreview "Draft FS 2025.docx" --year 2025                  # asks which category the entity is
python -m fsreview "Draft FS 2025.docx" --year 2025 --category MSC
python -m fsreview draft.pdf --template modern --page-size Letter --fix-spelling
python -m fsreview --list-references
python -m fsreview draft.docx --reference "Approved FS 2024.docx"     # copy the house style
python -m fsreview --help
```

The command exits with status 1 when the review found errors, so it can be used in a checklist script.

## Inputs

* **Word** (.docx) — formatted in place; the best route.
* **Excel audit files** (.xlsx, .xlsm) — firms keep the statements, notes and working papers in one workbook.
  Only the statement and note sheets are read as the financial statements: a sheet qualifies when it is visible,
  not named like a working paper (TB, ledger, adjusting entries, workings, provision, tax …), and has a print area or
  a statement title at the top. Only the print area is read; hidden rows and columns are skipped; every figure is
  read as Excel displays it; note numbers kept in their own column (4, 4.1 …) become note headings; narrative,
  tables and signature lines are separated. The report lists the working papers and hidden sheets that were not
  read. Figures carrying decimals hidden by the number format are counted (they make printed totals fail to add
  up). **The client's workbook is never re-saved**: only the print settings of the statement sheets are patched
  (paper size, fit to page width, centred, page numbers, gridlines off) and every other part of the file — formulas,
  logos, working papers — is proved byte-for-byte identical. Ratios and analysis go into a separate workbook.
* **.xls** — read through `xlrd`; the formatted workbook is saved as .xlsx (formulas cannot be carried over).
* **PDF** — text PDFs are rebuilt into Word: paragraphs are rejoined, figure columns are found from their positions, wrapped labels are rejoined, and running headers/footers are dropped from the body (the report lists them). Scanned pages are read with OCR when `pytesseract` and Tesseract are installed; otherwise the page is marked in the document and the report as needing manual work.
* **Text** (.txt) — columns are found from spacing.
* **Reference template** — an approved Word set; its font, sizes and margins become the house style.

Nothing is dropped silently: text boxes, content controls, tracked changes, unreadable pages and formulas without results are all listed as items for manual review.

## What the formatting engine does

It first works out what each part of the document is — statement titles, period lines ("For the year ended…"), unit lines, headings, numbered note headings, body text and contents entries — and, for every table, the label, note and amount columns, the year above each column, and whether each row is a heading, line item, subtotal, total or grand total. Then each part gets its own treatment:

* **Big 4 print layout** (default template): Arial, left-aligned upper-case title blocks, period line in capitals,
  a centred `------ Rupees ------` banner over the year columns, note headings in capitals with the number at
  the margin and the text on a 1.27 cm hanging indent (sub-notes 3.1, 3.2 the same), note tables indented to
  match, "The annexed notes 1 to N form an integral part of these financial statements." and signature blocks
  (Chief Executive, Director, and Chief Financial Officer for a PIC) added under each primary statement when
  missing, optional binding gutter. Every added line is in the audit trail.
* **Primary statements** start on a new page with a centred title block (entity name, title, period). Amounts are right-aligned in fixed-width columns, the note column is centred, column headings repeat on each page, rows never split across pages, and short statements are kept on one page. Totals get a single rule above; grand totals (total assets, total equity and liabilities, profit for the year, total comprehensive income, closing cash, closing equity) are bold with a double rule below.
* **Figures** are written one way throughout: thousands separators, negatives in brackets (or with a minus sign), nil as a dash, and the same decimal places down each column. Only the presentation changes; each rewritten figure is parsed again and must have the same value, and each change is in the audit trail.
* **Statement of changes in equity** — movements are line items and balances are the totals, with the closing balance double-ruled.
* **Notes** — note headings get one numbering style (`12.` vs `12`), sub-notes keep theirs (`3.1`), headings are kept with the text that follows.
* **Body text** — one font, size, colour, line spacing and paragraph spacing; long paragraphs are justified; double spaces and spaces before punctuation are removed; runs of empty paragraphs are collapsed.
* **Contents** — page numbers on a right tab with dot leaders.
* **Page** — page size, margins, header with the entity name, footer with "Page X of Y".

Templates: Big 4 print layout, Classic (Times New Roman), Modern (Arial), Calibri and Annual report (Garamond), or the house style read from a reference document.

## Financial analysis

Ratios (current and comparative, with formula and observation): gross, operating and net margins, return on
equity and assets, effective tax rate, current, quick and cash ratios, inventory, receivable and payable days,
asset turnover, debt to equity, gearing, liabilities to assets, interest cover, cash flow to profit and free
cash flow. Growth of revenue, profits, assets, equity and operating cash flow. Horizontal (change and %) and
vertical (% of revenue / % of total assets) analysis of every line. Points for the reviewer list movements over
10% and ratios that deteriorated, and a current ratio below 1 is flagged for the going concern assessment. The
observations are mechanical starting points for your analytical procedures.

## What the review checks

**Accounting cross-checks**

* Statement of financial position balances (total assets = total equity and liabilities, or equity + liabilities, or net assets = equity), current year and comparative.
* Total comprehensive income = profit + other comprehensive income.
* Cash flows: operating + investing + financing = net change; opening + net change (+ FX) = closing; closing cash agrees to the statement of financial position.
* Statement of changes in equity: components add across to the total column; opening + movements = closing for each column; closing equity agrees to the statement of financial position (current and comparative); profit agrees to the statement of profit or loss.
* Every statement line that cites a note shows a figure that appears in that note.

**Casting** — every subtotal and total is recomputed. The engine understands the usual layouts: a total of the lines above, a total of earlier subtotals (total assets), a running total (gross profit → operating profit → profit before tax → profit for the year), the difference of two totals (net current assets) and lines left without their own subtotal. A failing total shows the figure shown, the figure computed and the year.

**Consistency and comparatives** — comparative column present; first column matches the selected financial year; years consecutive and current year first; blank comparatives; one currency/unit label; one period-end date across the statement headings; duplicate, missing or out-of-order note numbers; references to notes that do not exist; mixed terminology (balance sheet vs statement of financial position, debtors vs receivables, -ise vs -ize, judgement vs judgment, …).

**Further checks** — a restated set must show a third statement of financial position at the start of the
comparative year (IAS 1.40A) and a restatement note (IAS 8.49); "The annexed notes 1 to N" must name the last note;
duplicate note numbers; in the statement of changes in equity, a "Total comprehensive income" line that shows a
balance instead of profit + OCI; closing cash agreed to the cash note it cites; the statement of compliance must
name the framework the confirmed category requires (e.g. IFRS for SMEs for an MSC, not the withdrawn AFRS for
SMEs), and only one framework may be named throughout; the same paragraph written twice with different figures;
movement schedules (fixed assets, loans, provisions) reconcile opening + movements = closing, carry forward, and
add across to the total column; note tables continued across pages are cast as one.

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
| `fsreview/xlsx_read.py` | Statement/note sheet selection, print areas and row layout for Excel audit files |
| `fsreview/xlsx_patch.py` | Lossless print set-up of the client's workbook and the byte-for-byte proof |
| `fsreview/classify.py` | Sections, paragraph roles, table columns and row kinds |
| `fsreview/checks.py` | Casting and accounting cross-checks |
| `fsreview/entity.py` | Entity category suggestion and thresholds |
| `fsreview/disclosures.py` | Disclosure checklist per category |
| `fsreview/library.py` | Reference library and note-by-note comparison |
| `fsreview/figures.py`, `fsreview/analysis.py` | Key figures, ratios, charts and the analysis report |
| `fsreview/proofread.py` | Proofreading and consistency |
| `fsreview/docx_engine.py` | Word formatting engine |
| `fsreview/xlsx_engine.py` | Excel formatting and verification |
| `fsreview/pdf_engine.py` | PDF output |
| `fsreview/integrity.py` | Proves the content was preserved |
| `fsreview/report.py` | Review report |
| `fsreview/server.py`, `fsreview/web/` | Local browser interface |
