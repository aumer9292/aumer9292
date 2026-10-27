"""Build deliberately messy sample financial statements for trying the tool.

    python samples/make_samples.py

Writes sample-fs.docx, sample-fs.xlsx and sample-fs.txt into this folder. The
documents contain planted problems the review should find: a casting error, a
note reference to a missing note, an impossible date, a superseded standard,
misspellings, mixed terminology, and inconsistent fonts and number formats.
"""

import os
import random

import docx
import openpyxl
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__))

SOFP = [
    ("", "Note", "2025", "2024"),
    ("", "", "$'000", "$'000"),
    ("ASSETS", "", "", ""),
    ("Non-current assets", "", "", ""),
    ("Property, plant and equipment", "5", "12500", "11800"),
    ("Intangible assets", "6", "1,250", "1,400"),
    ("", "", "13750", "13,200"),
    ("Current assets", "", "", ""),
    ("Inventories", "7", "3,420", "3,100"),
    ("Trade and other recievables", "8", "2,880", "2,640"),
    ("Cash and cash equivalents", "9", "1,950", "1,460"),
    ("", "", "8,250", "7,200"),
    ("Total assets", "", "22,000", "20,400"),
    ("EQUITY AND LIABILITIES", "", "", ""),
    ("Equity", "", "", ""),
    ("Share capital", "10", "5,000", "5,000"),
    ("Retained earnings", "", "9,600", "8,100"),
    ("Total equity", "", "14,600", "13,100"),
    ("Non-current liabilities", "", "", ""),
    ("Borrowings", "11", "3,500", "4,000"),
    ("Current liabilities", "", "", ""),
    ("Trade and other payables", "12", "3,100", "2,700"),
    ("Current tax liability", "", "800", "600"),
    ("", "", "3,900", "3,300"),
    ("Total liabilities", "", "7,400", "7,300"),
    ("Total equity and liabilities", "", "22,000", "20,400"),
]
# 12,500 + 1,250 = 13,750; current 3,420+2,880+1,950 = 8,250  (13,750+8,250 = 22,000)

SOCI = [
    ("", "Note", "2025", "2024"),
    ("", "", "$'000", "$'000"),
    ("Revenue", "13", "28,400", "25,900"),
    ("Cost of sales", "", "(17,900)", "(16,500)"),
    ("Gross profit", "", "10,500", "9,400"),
    ("Administrative expenses", "", "-5,200", "(4,900)"),
    ("Distribution costs", "", "(2,100)", "(1,950)"),
    ("Operating profit", "", "3,200", "2,550"),
    ("Finance costs", "", "(350)", "(400)"),
    ("Profit before tax", "", "2,850", "2,150"),
    ("Income tax expense", "14", "(850)", "(650)"),
    ("Profit for the year", "", "2,000", "1,500"),
    ("Other comprehensive income", "", "0", "-"),
    ("Total comprehensive income for the year", "", "2,000", "1,500"),
]

SOCE = [
    ("", "Share capital", "Retained earnings", "Total"),
    ("", "$'000", "$'000", "$'000"),
    ("Balance at 1 January 2024", "5,000", "7,100", "12,100"),
    ("Profit for the year", "-", "1,500", "1,500"),
    ("Dividends paid", "-", "(500)", "(500)"),
    ("Balance at 31 December 2024", "5,000", "8,100", "13,100"),
    ("Profit for the year", "-", "2,000", "2,000"),
    ("Dividends paid", "-", "(500)", "(500)"),
    ("Balance at 31 December 2025", "5,000", "9,600", "14,600"),
]

SOCF = [
    ("", "Note", "2025", "2024"),
    ("", "", "$'000", "$'000"),
    ("Cash flows from operating activities", "", "", ""),
    ("Profit before tax", "", "2,850", "2,150"),
    ("Depreciation and amortisation", "", "1,300", "1,250"),
    ("Finance costs", "", "350", "400"),
    ("Increase in inventories", "", "(320)", "(150)"),
    ("Increase in trade receivables", "", "(240)", "(180)"),
    ("Increase in trade payables", "", "400", "120"),
    ("Cash generated from operations", "", "4,340", "3,590"),
    ("Interest paid", "", "(350)", "(400)"),
    ("Income tax paid", "", "(650)", "(600)"),
    ("Net cash from operating activities", "", "3,340", "2,590"),
    ("Cash flows from investing activities", "", "", ""),
    ("Purchase of property, plant and equipment", "", "(1,850)", "(1,600)"),
    ("Net cash used in investing activities", "", "(1,850)", "(1,600)"),
    ("Cash flows from financing activities", "", "", ""),
    ("Repayment of borrowings", "", "(500)", "(500)"),
    ("Dividends paid", "", "(500)", "(500)"),
    ("Net cash used in financing activities", "", "(1,000)", "(1,000)"),
    ("Net increase in cash and cash equivalents", "", "490", "(10)"),
    ("Cash and cash equivalents at beginning of the year", "", "1,460", "1,470"),
    ("Cash and cash equivalents at end of the year", "9", "1,950", "1,460"),
]

PPE_NOTE = [
    ("", "Land and buildings", "Plant and machinery", "Total"),
    ("", "$'000", "$'000", "$'000"),
    ("Cost at 1 January 2025", "8,000", "9,600", "17,600"),
    ("Additions", "400", "1,450", "1,850"),
    ("Cost at 31 December 2025", "8,400", "11,050", "19,450"),
    ("Accumulated depreciation", "(1,600)", "(5,350)", "(6,950)"),
    ("Carrying amount at 31 December 2025", "6,800", "5,700", "12,500"),
]

FONTS = ["Arial", "Calibri", "Times New Roman", "Verdana"]


def _para(d, text, bold=False, size=None, font=None):
    p = d.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size or random.choice([9, 10, 11, 12]))
    r.font.name = font or random.choice(FONTS)
    return p


def _table(d, rows):
    t = d.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            if v:
                r = t.cell(i, j).paragraphs[0].add_run(v)
                r.font.name = random.choice(FONTS)
                r.font.size = Pt(random.choice([9, 10, 11]))
                if i < 2 or row[0].lower().startswith("total"):
                    r.bold = random.random() < 0.6
    return t


def make_docx(path):
    random.seed(7)
    d = docx.Document()
    _para(d, "NORTHWIND TRADING LIMITED", True, 16)
    _para(d, "Financial Statements for the year ended 31 December 2025", False, 13)
    _para(d, "")
    _para(d, "Contents", True)
    for line in ("Directors' report\t1", "Statement of financial position\t3", "Statement of profit or loss\t4",
                 "Statement of changes in equity\t5", "Statement of cash flows\t6",
                 "Notes to the financial statements\t7"):
        _para(d, line)
    _para(d, "")
    _para(d, "Directors' Report", True, 13)
    _para(d, "The directors present their report together with the audited financial statements for the year "
             "ended 31 December 2025.  The company's principal activity is the wholesale distribution of "
             "household goods , and there was no significant change in the nature of the business during the year.")
    _para(d, "Revenue increased by 9.7% to $28.4 million, mainly due to higher volumes in the the northern region. "
             "The balance sheet remains strong with net assets of $14.6 million.")
    _para(d, "")
    _para(d, "")
    _para(d, "")
    _para(d, "NORTHWIND TRADING LIMITED", True, 12)
    _para(d, "Statement of Financial Position", True, 12)
    _para(d, "As at 31 December 2025", False, 10)
    _table(d, SOFP)
    _para(d, "The notes on pages 7 to 12 form an integral part of these financial statements.", size=9)
    _para(d, "NORTHWIND TRADING LIMITED", True, 12)
    _para(d, "STATEMENT OF PROFIT OR LOSS AND OTHER COMPREHENSIVE INCOME", True, 11)
    _para(d, "For the year ended 31 December 2025", False, 10)
    _table(d, SOCI)
    _para(d, "NORTHWIND TRADING LIMITED", True)
    _para(d, "Statement of changes in equity", True)
    _para(d, "For the year ended 31 December 2025")
    _table(d, SOCE)
    _para(d, "NORTHWIND TRADING LIMITED", True)
    _para(d, "Statement of Cash Flows", True)
    _para(d, "For the year ended 31 June 2025")
    _table(d, SOCF)
    _para(d, "Notes to the Financial Statements", True, 13)
    _para(d, "For the year ended 31 December 2025")
    _para(d, "1. General information", True)
    _para(d, "Northwind Trading Limited is a private company incorporated and domiciled in the country. Its "
             "registered office is at 1 Harbour Road.")
    _para(d, "2. Basis of preparation", True)
    _para(d, "These financial statements have been prepared in accordance with International Financial Reporting "
             "Standards (IFRS) on the historical cost basis. They are presented in US dollars, rounded to the "
             "nearest thousand.")
    _para(d, "3. Significant accounting policies", True)
    _para(d, "3.1 Revenue", True)
    _para(d, "Revenue is recognised in accordance with IAS 18 when the significant risks and rewards of ownership "
             "have transferred to the buyer. Revenue is measured at the fair value of the consideration recieved, "
             "net of discounts.")
    _para(d, "3.2 Property, plant and equipment", True)
    _para(d, "Property, plant and equipment is stated at cost less accumulated depreciaton and impairment. "
             "Depreciation is recognized on a straight-line basis over the estimated useful lives.")
    _para(d, "4. Critical accounting estimates and judgements", True)
    _para(d, "The preparation of financial statements requires management to make judgements and estimates "
             "that affect the reported amounts of assets and liabilites.")
    _para(d, "5 Property, plant and equipment", True)
    _table(d, PPE_NOTE)
    _para(d, "6. Intangible assets", True)
    _para(d, "Intangible assets comprise software licences with a carrying amount of $1,250,000 (2024: $1,400,000).")
    _para(d, "7. Inventories", True)
    _para(d, "Inventories comprise finished goods held for resale.")
    _para(d, "8. Trade and other receivables", True)
    _para(d, "Trade receivables are non-interest bearing and are generally on 30-day terms.")
    _para(d, "9. Cash and cash equivalents", True)
    _para(d, "Cash and cash equivalents comprise cash at bank and short-term deposits.")
    _para(d, "10. Share capital", True)
    _para(d, "5,000,000 ordinary shares of $1 each, fully paid.")
    _para(d, "12. Trade and other payables", True)
    _para(d, "Trade creditors are settled on 45-day terms.")
    _para(d, "13. Revenue", True)
    _para(d, "Revenue arises from the sale of goods.")
    _para(d, "14. Income tax", True)
    _para(d, "The tax charge is based on the profit for the year at the standard rate.")
    d.save(path)


def make_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SOFP"
    ws["A1"] = "Northwind Trading Limited"
    ws["A2"] = "Statement of Financial Position as at 31 December 2025"
    r = 4
    ws.cell(r, 1, "")
    ws.cell(r, 2, "Note")
    ws.cell(r, 3, 2025)
    ws.cell(r, 4, 2024)
    r += 1
    for label, note, cy, py in SOFP[2:]:
        r += 1
        ws.cell(r, 1, label)
        ws.cell(r, 2, int(note) if note else None)
        for c, v in ((3, cy), (4, py)):
            if v:
                ws.cell(r, c, int(v.replace(",", "")))
    # subtotals and totals as formulas, so the workbook carries formulas to preserve
    for col in "CD":
        ws[f"{col}10"] = f"=SUM({col}8:{col}9)"
        ws[f"{col}15"] = f"=SUM({col}12:{col}14)"
        ws[f"{col}16"] = f"={col}10+{col}15"
    ws2 = wb.create_sheet("SOCI")
    ws2["A1"] = "Statement of Profit or Loss"
    for i, (label, note, cy, py) in enumerate(SOCI, start=3):
        ws2.cell(i, 1, label)
        for c, v in ((3, cy), (4, py)):
            if v and v not in ("-",) and not v.startswith("$"):
                n = v.replace(",", "").replace("(", "-").replace(")", "")
                try:
                    ws2.cell(i, c, int(n))
                except ValueError:
                    ws2.cell(i, c, v)
            elif v:
                ws2.cell(i, c, v)
    wb.save(path)


def make_txt(path):
    lines = ["NORTHWIND TRADING LIMITED", "STATEMENT OF FINANCIAL POSITION", "AS AT 31 DECEMBER 2025", ""]
    for label, note, cy, py in SOFP:
        lines.append(f"{label:<45}{note:>6}{cy:>12}{py:>12}")
    lines += ["", "STATEMENT OF PROFIT OR LOSS", "FOR THE YEAR ENDED 31 DECEMBER 2025", ""]
    for label, note, cy, py in SOCI:
        lines.append(f"{label:<45}{note:>6}{cy:>12}{py:>12}")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


if __name__ == "__main__":
    make_docx(os.path.join(HERE, "sample-fs.docx"))
    make_xlsx(os.path.join(HERE, "sample-fs.xlsx"))
    make_txt(os.path.join(HERE, "sample-fs.txt"))
    print("Samples written to", HERE)


def make_audit_pack(path):
    """An Excel audit file like the ones firms keep: statements and notes with print areas, note numbers in
    their own column, a working-paper sheet, a hidden old version and a logo image."""
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image as PILImage

    wb = openpyxl.Workbook()
    bs = wb.active
    bs.title = "BS"
    rows = [
        ("ALPHA BUILDERS (PRIVATE) LIMITED",), ("STATEMENT OF FINANCIAL POSITION",),
        ("AS AT JUNE 30, 2026", None, None, None, None, "Restated"), (),
        (None, "Note", None, 2026, None, 2025), (None, None, None, "(Rupees)", None, "(Rupees)"), (),
        ("NON-CURRENT ASSETS",),
        ("Operating fixed assets", 4, None, 1000000.4, None, 900000.0),
        ("CURRENT ASSETS",),
        ("Trade receivables", 5, None, 500000.3, None, 400000.0),
        ("Cash and bank balances", 6, None, 250000, None, 200000),
        (None, None, None, "=D11+D12", None, "=F11+F12"),
        ("TOTAL ASSETS", None, None, "=D9+D13", None, "=F9+F13"), (),
        ("SAHRE CAPITAL AND RESERVES",),
        ("Issued, subscribed and paid-up capital", 7, None, 1000000, None, 1000000),
        ("Accumulated profit", None, None, 250000.7, None, 100000),
        (None, None, None, "=D17+D18", None, "=F17+F18"),
        ("CURRENT LIABILITIES",),
        ("Creditors, accruals and other payables", 8, None, 500000, None, 400000),
        ("TOTAL EQUITY AND LIABILITIES", None, None, "=D19+D21", None, "=F19+F21"), (),
        ("The annexed notes from 1-9 form an integral part of these financial statements.",), (),
        ("CHIEF EXECUTIVE", None, None, None, None, "DIRECTOR"),
    ]
    for r in rows:
        bs.append(list(r))
    for row in bs.iter_rows(min_row=9, max_row=22, min_col=4, max_col=6):
        for c in row:
            c.number_format = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'
    bs.print_area = "A1:F26"
    bs["H9"] = "check"            # outside the print area: a working note, not part of the statement
    logo = PILImage.new("RGB", (60, 30), (31, 78, 156))
    logo_path = os.path.splitext(path)[0] + "-logo.png"
    logo.save(logo_path)
    bs.add_image(XLImage(logo_path), "H1")

    notes = wb.create_sheet("Notes")
    nrows = [
        ("ALPHA BUILDERS (PRIVATE) LIMITED",), ("NOTES TO THE FINANCIAL STATEMENTS",),
        ("FOR THE YEAR ENDED JUNE 30, 2026",), (),
        (1, "LEGAL STATUS AND OPERATIONS"),
        (None, "Alpha Builders (Private) Limited is a private limited company incorporated in Pakistan under the "
               "Companies Act, 2017. Its registered office is in Lahore."), (),
        (4, "OPERATING FIXED ASSETS"),
        (4.1, "Allocation of depreciation"),
        (None, "Cost of services", None, 60000, None, 50000),
        (None, "Administrative expenses", None, 40000, None, 30000),
        (None, None, None, "=D10+D11", None, "=F10+F11"),
    ]
    for r in nrows:
        notes.append(list(r))
    notes["A9"].number_format = "0.0"
    notes.print_area = "A1:F12"

    tb = wb.create_sheet("TB 2026")
    tb.append(["Account", "Debit", "Credit"])
    tb.append(["Cash", 250000, None])
    old = wb.create_sheet("BS (old)")
    old["A1"] = "STATEMENT OF FINANCIAL POSITION"
    old.sheet_state = "hidden"
    wb.save(path)
    os.remove(logo_path)
