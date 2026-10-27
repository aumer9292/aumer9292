# FS Review — goal and what remains

## The goal

A reviewer drops a draft set of financial statements into one tool and, in minutes, gets back:

1. **Print-ready statements** in the firm's Big 4 layout (Word and PDF), with no figure or wording changed
   silently.
2. **Confirmation that the right notes and disclosures are there** for the entity's category under the
   Companies Act 2017 (PIC, LSC, MSC, SSC), compared with the ICAP illustrative set and the firm's model set.
3. **A review report** of every error a manager would otherwise find by hand: casting, cross-statement
   agreement, comparatives and restatements, note references and numbering, framework wording, spelling.
4. **Analytical review**: ratios, growth, horizontal/vertical analysis and charts.

The tool is complete when a manager can run a real client file through it and sign off the output with only
the judgement calls left to make.

## Done

| Area | Status |
| --- | --- |
| Inputs: Word, Excel audit files (statements + working papers), PDF (OCR optional), text | Done |
| Entity category suggested from turnover / capital / employees / listing, confirmed by the user | Done |
| Disclosure checklist per category (~60 items, line-driven items marked not applicable) | Done |
| Reference library: ICAP illustratives and model sets, note-by-note comparison by meaning | Done (library empty until sets are uploaded) |
| Casting, schedules, cross-checks, restatement, framework and consistency checks | Done |
| Proofreading (finance spelling list, dates, repeated paragraphs with different figures, superseded standards) | Done |
| Big 4 print layout: title blocks, Rupees banner, totals rules, hanging note numbers, signatures, landscape for wide tables | Done |
| Excel workbook left byte-for-byte intact except print settings | Done |
| Financial analysis report, charts and analysis workbook | Done |
| Browser interface, command line, Windows launcher, 34 automated tests | Done |
| Tried on a real file (Asghar & Co, 57-sheet audit workbook) | Done — 13 genuine findings |

## What remains to complete the goal

| # | Step | Who | Done when |
| --- | --- | --- | --- |
| 1 | Upload the ICAP illustrative sets (PIC, LSC, MSC, SSC) into the Reference library | You | Each category lists its reference notes |
| 2 | Send the Service Industries model PDF (attach it in the chat) so the reader is tested on it and it is added as the model set | You → Claude | The model's notes and layout are read correctly |
| 3 | Confirm the category thresholds against the current SECP notification; correct `thresholds.json` if needed | You | Thresholds signed off |
| 4 | Partner/manager review of the disclosure checklist per category, especially SSC (Revised AFRS for SSEs) and the exact Companies Act schedule paragraphs | You (with Claude making the edits) | Checklist signed off |
| 5 | Run 3–5 more real files (one Word set, one PDF, one small company, one listed/PIC) and fix what they show | You → Claude | No false errors on those files |
| 6 | Compare the printed output with an approved signed set from your firm and tune fonts, spacing and title layout | You → Claude | Output accepted as house style |
| 7 | Next level of disclosure checks: content, not only presence (e.g. PPE reconciliation complete, related-party note has nature/amount/balance, remuneration table has all columns) | Claude | Agreed items checked for content |
| 8 | Optional: a Windows installer (`FSReview.exe`) so Python need not be installed; OCR for scanned PDFs; AI review with an API key | Claude (built on a Windows PC) | Installs with one click |

Steps 1–3 need only you and take an afternoon. Steps 4–6 are where most of the remaining value is: each real file
you run makes the checks sharper.
