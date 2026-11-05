import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as CL

OUT = "/home/user/aumer9292/audit/SSSF_FY2026_PMIC_Financing_WP.xlsx"
D = dt.date
F = "Arial"
NUM = '#,##0;(#,##0);"-"'
PCT = '0.00%;(0.00%);"-"'
DATEF = 'dd-mmm-yyyy'
NAVY = PatternFill("solid", fgColor="1F3864")
AMBER = PatternFill("solid", fgColor="FFC000")
GREY = PatternFill("solid", fgColor="D9E1F2")
LIGHT = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="808080")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
TOP = Border(top=thin)
DBL = Border(top=thin, bottom=Side(style="double", color="000000"))

BLUE = "0000FF"; GREEN = "008000"; BLACK = "000000"

wb = Workbook()
wb.remove(wb.active)


def put(ws, ref, v, bold=False, color=BLACK, fmt=None, fill=None, wrap=False, italic=False,
        border=None, align=None, size=10, comment=None):
    c = ws[ref]
    c.value = v
    c.font = Font(name=F, bold=bold, color=color, italic=italic, size=size)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if border: c.border = border
    if wrap or align:
        c.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal=align)
    if comment: c.comment = Comment(comment, "RSM audit team")
    return c


def header(ws, ref_code, title, widths):
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[CL(i)].width = w
    put(ws, "A1", "Shah Sachal Sami Foundation (SSSF)", bold=True, size=12)
    put(ws, "A2", "Statutory audit – year ended 30 June 2026 (RSM first-year audit)")
    put(ws, "A3", f"{ref_code}  |  {title}", bold=True, size=11, color="1F3864")
    put(ws, "A4", "Area: Financing for microcredit lending – secured (PMIC term finance facility)", italic=True)
    last = CL(len(widths))
    put(ws, f"{CL(len(widths)-1)}1", "Prepared by / date:", italic=True)
    put(ws, f"{last}1", None, fill=AMBER, border=BOX)
    put(ws, f"{CL(len(widths)-1)}2", "Reviewed by / date:", italic=True)
    put(ws, f"{last}2", None, fill=AMBER, border=BOX)
    put(ws, f"{CL(len(widths)-1)}3", "WP ref:", italic=True)
    put(ws, f"{last}3", ref_code, bold=True)


def thead(ws, row, labels, col=1):
    for i, l in enumerate(labels):
        put(ws, f"{CL(col+i)}{row}", l, bold=True, color="FFFFFF", fill=NAVY, wrap=True, border=BOX)
    ws.row_dimensions[row].height = 30


def section(ws, ref, text):
    put(ws, ref, text, bold=True, fill=GREY)


# ===================================================================== INDEX
ix = wb.create_sheet("Index")
header(ix, "A-0", "Index, legend and conclusion", [14, 34, 70, 18, 22])

# ===================================================================== INPUTS
inp = wb.create_sheet("Inputs")
header(inp, "A-IN", "Inputs – facts, figures and contract terms (all formulas link here)",
       [13, 58, 18, 8, 46, 24])
thead(inp, 6, ["Key", "Description", "Value", "Unit", "Source / contract clause", "Status"])
I = {}
rows = [
    ("§", "MATERIALITY", None, None, None, None),
    ("OM", "Overall materiality", 600000, "Rs", "RSM planning memorandum", "Given"),
    ("PM", "Performance materiality", 300000, "Rs", "RSM planning memorandum", "Given"),
    ("YE", "Reporting date", D(2026, 6, 30), "date", "Financial year end", "Given"),
    ("PYE", "Prior reporting date (opening balance date)", D(2025, 6, 30), "date", "FY2025 audited FS (predecessor auditor)", "Given"),
    ("§", "BALANCES PER DRAFT FINANCIAL STATEMENTS", None, None, None, None),
    ("FS_NC", "Financing – non-current (Note 14)", 113750000, "Rs", "Draft FS Note 14", "Given"),
    ("FS_CUR", "Financing – current portion (Note 14)", 35000000, "Rs", "Draft FS Note 14", "Given"),
    ("GL_NC", "Financing – non-current per trial balance / GL", 113750000, "Rs", "TB 30-Jun-26 – agree GL code", "AMBER: agree to TB"),
    ("GL_CUR", "Financing – current portion per trial balance / GL", 35000000, "Rs", "TB 30-Jun-26 – agree GL code", "AMBER: agree to TB"),
    ("OPEN_FS", "Opening balance 1-Jul-25 per draft FS comparative", 148750000, "Rs", "Draft FS Note 14 (comparative)", "Given"),
    ("OPEN_PY", "Opening balance per FY2025 AUDITED financial statements", 148750000, "Rs", "FY2025 audited FS (ISA 510) – obtain signed copy", "AMBER: agree to signed FS"),
    ("DRAW", "Drawdowns received in FY2026", 0, "Rs", "GL / bank statements – PMIC receipts", "AMBER: bank data"),
    ("REPAY", "Principal repaid in FY2026", 0, "Rs", "GL / bank statements – PMIC payments", "AMBER: bank data"),
    ("CONF_BAL", "Principal outstanding per PMIC confirmation (ISA 505)", None, "Rs", "PMIC direct confirmation – to be received", "AMBER: PMIC confirmation"),
    ("FC_BOOK", "Finance cost / markup on PMIC facility booked (Note 23)", 3337705, "Rs", "Draft FS Note 23", "Given"),
    ("MP_BOOK", "Markup payable to PMIC at year end (Note 15)", 1112568, "Rs", "Draft FS Note 15", "Given"),
    ("FUNDS_PY", "Funds (equity) at 30-Jun-25", 50652884, "Rs", "FY2025 audited FS – incl. Rs 10m director sub-loan", "Given"),
    ("FUNDS_CY", "Funds (equity) at 30-Jun-26 – draft", 45406748, "Rs", "Draft FS", "Given"),
    ("SUBLOAN", "Director sub-loan included in funds", 10000000, "Rs", "Clause 6.2.13 sub-loan covenant", "Given"),
    ("FLOOR", "NBFC minimum equity requirement", 50000000, "Rs", "NBFC Regulations 2008 – minimum equity (confirm regulation ref.)", "Given"),
    ("§", "MARKUP RATES & PERIODS", None, None, None, None),
    ("KIBOR_OLD", "6-month KIBOR applicable to the Jul–Sep 2025 quarter (old-facility reset)", 0.11, "%", "Amendment Agreement 27-Aug-24; SBP/PMIC KIBOR reset – CONFIRM EXACT RATE", "AMBER: PMIC / bank data"),
    ("SPR_OLD", "Spread – old facility (6M KIBOR + 350 bps)", 0.035, "%", "Amendment Agreement 27-Aug-24 / TFA II", "Given"),
    ("OLD_FROM", "Old-rate period start", D(2025, 7, 1), "date", "Day after reporting date FY2025", "Given"),
    ("T1_FROM", "New rate tier (i) start – 3% p.a.", D(2025, 10, 1), "date", "Term Facility Agreement 2025 – Schedule B (i)", "Given"),
    ("R1", "Tier (i) markup rate 1-Oct-25 to 30-Jun-26", 0.03, "%", "Schedule B (i)", "Given"),
    ("T2_FROM", "New rate tier (ii) start – 7% p.a.", D(2026, 7, 1), "date", "Schedule B (ii)", "Given"),
    ("R2", "Tier (ii) markup rate 1-Jul-26 to 31-Dec-27", 0.07, "%", "Schedule B (ii)", "Given"),
    ("T3_FROM", "New rate tier (iii) start – 6M KIBOR + 100 bps", D(2028, 1, 1), "date", "Schedule B (iii)", "Given"),
    ("KIBOR_FWD", "Forecast 6-month KIBOR for 2028–2030 (for DCF)", 0.11, "%", "Market forward curve / SBP – ESTIMATE (ISA 540)", "AMBER: market data"),
    ("SPR3", "Tier (iii) spread", 0.01, "%", "Schedule B (iii)", "Given"),
    ("BASIS", "Day-count basis (days per year)", 365, "days", "Agreement – confirm day-count convention", "AMBER: PMIC confirmation"),
    ("§", "IFRS 9 MODIFICATION (Schedule F novation)", None, None, None, None),
    ("VAL_DATE", "Modification / novation effective date (valuation date)", D(2025, 6, 30), "date", "TFA signed ~24-Jun-25; first Disbursement Date – CONFIRM", "AMBER: PMIC confirmation"),
    ("MKT", "Market rate for equivalent unsubsidised financing (discount rate)", 0.15, "%", "ISA 540 estimate – SSSF/peer NBFC borrowing rate June 2025", "AMBER: market data"),
    ("TEST10", "IFRS 9 B3.3.6 quantitative threshold", 0.10, "%", "IFRS 9 B3.3.6", "Given"),
    ("§", "COVENANT DATA (clause 6.2.12)", None, None, None, None),
    ("CR_CLIENT", "Current ratio per client computation", 2.7, "x", "Client covenant computation", "Given"),
    ("CA", "Current assets at 30-Jun-26 (draft FS)", None, "Rs", "Draft FS – statement of financial position", "AMBER: input from FS"),
    ("CLI", "Current liabilities at 30-Jun-26 (draft FS, incl. Rs 35m current portion)", None, "Rs", "Draft FS – statement of financial position", "AMBER: input from FS"),
    ("OSS_FS", "Operational self-sufficiency (OSS) per FS / client", 1.41, "%", "Draft FS / covenant certificate", "Given"),
    ("OSS_CALC", "OSS recomputed = total income ÷ total expenditure", 0.86, "%", "RSM recomputation from draft FS", "Given"),
    ("PORT", "Gross microcredit portfolio at 30-Jun-26", None, "Rs", "Loan-loss provision WP", "AMBER: link to LLP WP"),
    ("PROV_FS", "Loan-loss provision per draft FS", None, "Rs", "Loan-loss provision WP", "AMBER: link to LLP WP"),
    ("PROV_REG", "Provision required under NBFC Regulations 2008", None, "Rs", "Loan-loss provision WP", "AMBER: link to LLP WP"),
    ("SUB_DATE", "Date director sub-loan of Rs 10m received", None, "date", "Bank statement / board minutes – clause 6.2.13", "AMBER: bank data"),
]
r = 7
for key, desc, val, unit, src, status in rows:
    if key == "§":
        section(inp, f"A{r}", desc)
        for c in "BCDEF": inp[f"{c}{r}"].fill = GREY
        r += 1
        continue
    put(inp, f"A{r}", key, italic=True)
    put(inp, f"B{r}", desc, wrap=True)
    fmt = NUM if unit == "Rs" else PCT if unit == "%" else DATEF if unit == "date" else '0.00"x"' if unit == "x" else "0"
    amber = status.startswith("AMBER")
    put(inp, f"C{r}", val, color=BLUE, fmt=fmt, fill=AMBER if amber else None, border=BOX)
    put(inp, f"D{r}", unit)
    put(inp, f"E{r}", src, wrap=True)
    put(inp, f"F{r}", status, bold=amber, color="C00000" if amber else BLACK)
    I[key] = f"Inputs!$C${r}"
    r += 1

# derived
r += 1
section(inp, f"A{r}", "DERIVED (formulas)")
for c in "BCDEF": inp[f"{c}{r}"].fill = GREY
r += 1
derived = [
    ("FS_TOT", "Financing total per draft FS", f"={I['FS_NC']}+{I['FS_CUR']}", NUM),
    ("R_OLD", "Old-facility all-in rate (6M KIBOR + 350 bps)", f"={I['KIBOR_OLD']}+{I['SPR_OLD']}", PCT),
    ("R3", "Tier (iii) all-in rate (forecast 6M KIBOR + 100 bps)", f"={I['KIBOR_FWD']}+{I['SPR3']}", PCT),
    ("ORIG_EIR", "Original effective rate of old facility (no fees – = contractual rate)", f"={I['KIBOR_OLD']}+{I['SPR_OLD']}", PCT),
    ("YE12", "End of 12-month window after reporting date", f"=DATE(YEAR({I['YE']})+1,MONTH({I['YE']}),DAY({I['YE']}))", DATEF),
]
for key, desc, f, fmt in derived:
    put(inp, f"A{r}", key, italic=True)
    put(inp, f"B{r}", desc, wrap=True)
    put(inp, f"C{r}", f, fmt=fmt, border=BOX)
    put(inp, f"E{r}", "Formula")
    I[key] = f"Inputs!$C${r}"
    r += 1

r += 1
section(inp, f"A{r}", "FACILITY & RESTRUCTURING HISTORY (contract references)")
for c in "BCDEF": inp[f"{c}{r}"].fill = GREY
r += 1
thead(inp, r, ["#", "Document / event", "Amount (Rs)", "", "Key terms", "Reference"])
r += 1
hist = [
    ("1", "Term Facility Agreement I (~March 2021)", 171250000, "KIBOR + 3.75% p.a.", "TFA I – obtain copy"),
    ("2", "Term Facility Agreement II (~2022)", 129000000, "6M KIBOR + 350 bps", "TFA II – obtain copy"),
    ("3", "Restructuring request letter 12-Jun-24; Amendment Agreement 27-Aug-24", 148750000,
     "12-month Principal Deferment Period 30-Jun-24 – 29-Jun-25 (principal deferred, markup continues); 6M KIBOR + 350 bps, quarterly; OSS/PAR covenants from Jul-25",
     "Amendment Agreement 27-Aug-24"),
    ("4", "Term Facility Agreement (Murabaha) signed ~24-Jun-25 – NOVATION", 148750000,
     "Sale Price Rs 148.75m settles existing facilities; Purchase Price / Mark-up structure; tiered rate; 15-month grace from first disbursement; 17 graduated quarterly instalments 30-Sep-26 to 30-Sep-30",
     "TFA 2025 – Schedule F (novation), Schedule B (rate & repayment)"),
    ("5", "Security – Letter of Hypothecation 11-Dec-17 (suppl. 1-Feb-19, 4-Mar-21, 15-Sep-22)", 250000000,
     "Hypothecation over current assets up to Rs 250m; SECP charge registration and post-novation modification to be confirmed",
     "Letter of Hypothecation; SECP Form 10 / charge certificate"),
    ("6", "Financial covenants", None,
     "Current ratio ≥ 1:1; D/E < 8x; OSS > 100% from Dec-26; PAR>29 < 5% from Dec-29; minimum provision = higher of 5% of portfolio or NBFC-regulatory provision",
     "TFA 2025 clause 6.2.12"),
    ("7", "Sub-loan covenant", 10000000, "Arrange Rs 10m sub-loan before 30-Jun-25", "TFA 2025 clause 6.2.13"),
]
for h in hist:
    put(inp, f"A{r}", h[0]); put(inp, f"B{r}", h[1], wrap=True)
    put(inp, f"C{r}", h[2], fmt=NUM, color=BLUE)
    put(inp, f"E{r}", h[3], wrap=True); put(inp, f"F{r}", h[4], wrap=True)
    inp.row_dimensions[r].height = 48
    r += 1
inp.freeze_panes = "A7"

# ===================================================================== A2 split (built before Lead because Lead references it)
sp = wb.create_sheet("A2 CNC Split")
lead = wb.create_sheet("A Lead", 2)
mv = wb.create_sheet("A1 Movement", 3)
# order: Index, Inputs, A Lead, A1 Movement, A2 CNC Split ...

header(sp, "A2", "Current vs non-current classification (IAS 1.69–76ZA) – Schedule B repayment profile",
       [6, 16, 18, 26, 14, 18, 18, 22])
put(sp, "A5", "Objective: principal contractually due within 12 months of 30-Jun-26 is current (IAS 1.69(d)); SSSF has the right at the reporting date to defer the remainder (IAS 1.72A).", italic=True)
thead(sp, 7, ["#", "Due date", "Instalment (Rs)", "Contract reference", "Due ≤ 12m? (1/0)", "Current (Rs)", "Non-current (Rs)", "Cumulative repaid"])
sched = ([D(2026, 9, 30), D(2026, 12, 31), D(2027, 3, 31), D(2027, 6, 30)], 1500000), \
        ([D(2027, 9, 30), D(2027, 12, 31), D(2028, 3, 31), D(2028, 6, 30)], 3000000), \
        ([D(2028, 9, 30)], 5000000), ([D(2028, 12, 31), D(2029, 3, 31), D(2029, 6, 30)], 6000000), \
        ([D(2029, 9, 30), D(2029, 12, 31), D(2030, 3, 31), D(2030, 6, 30), D(2030, 9, 30)], 21550000)
r = 8; n = 1
for dates, amt in sched:
    for d in dates:
        put(sp, f"A{r}", n)
        put(sp, f"B{r}", d, color=BLUE, fmt=DATEF)
        put(sp, f"C{r}", amt, color=BLUE, fmt=NUM)
        put(sp, f"D{r}", "TFA 2025 – Schedule B")
        put(sp, f"E{r}", f"=IF(AND(B{r}>{I['YE']},B{r}<={I['YE12']}),1,0)", fmt="0", align="center")
        put(sp, f"F{r}", f"=C{r}*E{r}", fmt=NUM)
        put(sp, f"G{r}", f"=C{r}-F{r}", fmt=NUM)
        put(sp, f"H{r}", f"=SUM($C$8:C{r})", fmt=NUM)
        r += 1; n += 1
SP_FIRST, SP_LAST = 8, r - 1
SPD = f"'A2 CNC Split'!$B${SP_FIRST}:$B${SP_LAST}"
SPA = f"'A2 CNC Split'!$C${SP_FIRST}:$C${SP_LAST}"
put(sp, f"B{r}", "Total", bold=True)
for c in "CFG":
    put(sp, f"{c}{r}", f"=SUM({c}{SP_FIRST}:{c}{SP_LAST})", bold=True, fmt=NUM, border=DBL)
TOT = r
r += 1
put(sp, f"B{r}", "Check: Schedule B total – facility outstanding (should be nil)", italic=True)
put(sp, f"C{r}", f"=C{TOT}-{I['FS_TOT']}", fmt=NUM, bold=True)
put(sp, f"D{r}", f'=IF(ABS(C{r})<1,"✓ Agrees","✗ Investigate")', bold=True)
r += 2
section(sp, f"A{r}", "Comparison to draft financial statements")
for c in "BCDEFGH": sp[f"{c}{r}"].fill = GREY
r += 1
thead(sp, r, ["", "Line item", "Per audit (Rs)", "Per draft FS Note 14 (Rs)", "", "Difference (Rs)", "Exceeds PM?", "Proposed AJE"])
r += 1
SP_CMP = r
put(sp, f"B{r}", "Current portion"); put(sp, f"C{r}", f"=F{TOT}", fmt=NUM)
put(sp, f"D{r}", f"={I['FS_CUR']}", fmt=NUM, color=GREEN)
put(sp, f"F{r}", f"=C{r}-D{r}", fmt=NUM, bold=True)
put(sp, f"G{r}", f'=IF(ABS(F{r})>{I["PM"]},"YES – misstatement","No")', bold=True, color="C00000")
put(sp, f"H{r}", "AJE (a) – see PF")
r += 1
put(sp, f"B{r}", "Non-current portion"); put(sp, f"C{r}", f"=G{TOT}", fmt=NUM)
put(sp, f"D{r}", f"={I['FS_NC']}", fmt=NUM, color=GREEN)
put(sp, f"F{r}", f"=C{r}-D{r}", fmt=NUM, bold=True)
put(sp, f"G{r}", f'=IF(ABS(F{r})>{I["PM"]},"YES – misstatement","No")', bold=True, color="C00000")
put(sp, f"H{r}", "AJE (a) – see PF")
r += 1
put(sp, f"B{r}", "Total", bold=True)
for c in "CDF":
    put(sp, f"{c}{r}", f"=SUM({c}{r-2}:{c}{r-1})", fmt=NUM, bold=True, border=DBL)
r += 2
SP_RECLASS = f"'A2 CNC Split'!$F${SP_CMP}"
section(sp, f"A{r}", "Proposed reclassification AJE (a)")
for c in "BCDEFGH": sp[f"{c}{r}"].fill = GREY
r += 1
put(sp, f"B{r}", "Dr Financing – current portion (Note 14)"); put(sp, f"F{r}", f"=-F{SP_CMP}", fmt=NUM)
r += 1
put(sp, f"B{r}", "    Cr Financing – non-current (Note 14)"); put(sp, f"G{r}", f"=-F{SP_CMP}", fmt=NUM)
r += 2
notes = [
    "Observations:",
    "1. The client classified Rs 35,000,000 as current; Schedule B requires only the four Rs 1,500,000 instalments (30-Sep-26 to 30-Jun-27) to be settled within 12 months → Rs 6,000,000 current.",
    "2. The Rs 35m figure appears to reflect an earlier (pre-2025 restructuring) repayment profile – ask management for its basis.",
    "3. IAS 1.72B (2020/2022 amendments): covenants to be complied with only AFTER the reporting date (OSS > 100% from Dec-26; PAR from Dec-29) do not affect classification at 30-Jun-26, but IAS 1.76ZA requires disclosure of those covenants and of facts indicating difficulty complying (OSS recomputed at 86% – see A5).",
    "4. If any covenant tested ON or BEFORE 30-Jun-26 was breached and no waiver/grace period existed at the reporting date, the whole balance becomes current (IAS 1.74). Obtain PMIC confirmation (ISA 505) – see A5.",
    "5. If AJE (b) is booked (A4), classification applies to the amortised-cost carrying amount; the IFRS 9 discount is presented against the non-current element (see A Lead).",
]
for t in notes:
    put(sp, f"A{r}", t, bold=t.endswith(":"), wrap=False); r += 1
sp.freeze_panes = "A8"

# ===================================================================== A3 markup
mk = wb.create_sheet("A3 Markup")
header(mk, "A3", "Markup / finance cost recomputation – full 12 months (IFRS 9 effective interest; contractual basis)",
       [34, 15, 15, 9, 18, 11, 18, 40])
put(mk, "A5", "Objective: recompute finance cost for 1-Jul-25 to 30-Jun-26 and markup payable at 30-Jun-26 (occurrence, completeness, accuracy, cut-off).", italic=True)
thead(mk, 7, ["Period", "From", "To", "Days", "Principal (Rs)", "Rate p.a.", "Markup (Rs)", "Contract reference / basis"])
put(mk, "A8", "Q1: Jul–Sep 2025 – old-facility rate")
put(mk, "B8", f"={I['OLD_FROM']}", fmt=DATEF, color=GREEN)
put(mk, "C8", f"={I['T1_FROM']}-1", fmt=DATEF)
put(mk, "D8", "=C8-B8+1", fmt="0")
put(mk, "E8", f"={I['OPEN_FS']}", fmt=NUM, color=GREEN)
put(mk, "F8", f"={I['R_OLD']}", fmt=PCT, color=GREEN, fill=AMBER, comment="Old-facility rate: 6M KIBOR + 350 bps (Amendment Agreement 27-Aug-24). Confirm exact KIBOR reset with PMIC / SBP data.")
put(mk, "G8", f"=E8*F8*D8/{I['BASIS']}", fmt=NUM)
put(mk, "H8", "Amendment Agreement 27-Aug-24 (6M KIBOR + 350 bps) – Schedule B new rate only from 1-Oct-25", wrap=True)
put(mk, "A9", "Q2–Q4: Oct 2025 – Jun 2026 – Schedule B (i)")
put(mk, "B9", f"={I['T1_FROM']}", fmt=DATEF, color=GREEN)
put(mk, "C9", f"={I['YE']}", fmt=DATEF, color=GREEN)
put(mk, "D9", "=C9-B9+1", fmt="0")
put(mk, "E9", f"={I['OPEN_FS']}", fmt=NUM, color=GREEN)
put(mk, "F9", f"={I['R1']}", fmt=PCT, color=GREEN)
put(mk, "G9", f"=E9*F9*D9/{I['BASIS']}", fmt=NUM)
put(mk, "H9", "TFA 2025 Schedule B (i): 3% p.a. 1-Oct-25 to 30-Jun-26", wrap=True)
put(mk, "A10", "Expected finance cost FY2026 (contractual)", bold=True)
put(mk, "D10", "=SUM(D8:D9)", fmt="0", bold=True)
put(mk, "G10", "=SUM(G8:G9)", fmt=NUM, bold=True, border=DBL)
put(mk, "A11", "Finance cost booked – Note 23")
put(mk, "G11", f"={I['FC_BOOK']}", fmt=NUM, color=GREEN)
put(mk, "A12", "Finance cost UNDERSTATED – proposed AJE (c)", bold=True, color="C00000")
put(mk, "G12", "=G10-G11", fmt=NUM, bold=True, border=DBL)
put(mk, "H12", f'=IF(ABS(G12)>{I["PM"]},"Exceeds PM – misstatement","Within PM")', bold=True, color="C00000")

section(mk, "A14", "Reconciliation of booked amount – what did the client compute?")
for c in "BCDEFGH": mk[f"{c}14"].fill = GREY
put(mk, "A15", "Oct–Jun at 3% on actual/365 (Q2–Q4 only)")
put(mk, "G15", "=G9", fmt=NUM)
put(mk, "A16", "Booked – Note 23"); put(mk, "G16", "=G11", fmt=NUM)
put(mk, "A17", "Difference", bold=True); put(mk, "G17", "=G16-G15", fmt=NUM, bold=True, border=TOP)
put(mk, "H15", "Booked figure reproduces EXACTLY as Q2–Q4 at 3% on actual/365 → the Jul–Sep 2025 quarter has not been accrued.", wrap=True)
mk.row_dimensions[15].height = 40
put(mk, "A18", "Memo: simple 9/12 basis at 3%")
put(mk, "G18", f"=E9*F9*9/12", fmt=NUM)
put(mk, "H18", "Day-count difference vs actual/365 (immaterial)")
put(mk, "A19", "Memo: day-count difference"); put(mk, "G19", "=G18-G15", fmt=NUM)

section(mk, "A21", "Markup payable at 30-Jun-26 (Note 15) – cut-off")
for c in "BCDEFGH": mk[f"{c}21"].fill = GREY
put(mk, "A22", "Q4: Apr–Jun 2026 accrued, unpaid at year end")
put(mk, "B22", f"=DATE(YEAR({I['YE']}),4,1)", fmt=DATEF)
put(mk, "C22", f"={I['YE']}", fmt=DATEF, color=GREEN)
put(mk, "D22", "=C22-B22+1", fmt="0")
put(mk, "E22", f"={I['OPEN_FS']}", fmt=NUM, color=GREEN)
put(mk, "F22", f"={I['R1']}", fmt=PCT, color=GREEN)
put(mk, "G22", f"=E22*F22*D22/{I['BASIS']}", fmt=NUM)
put(mk, "H22", "Schedule B (i) – quarterly markup")
put(mk, "A23", "Markup payable per Note 15"); put(mk, "G23", f"={I['MP_BOOK']}", fmt=NUM, color=GREEN)
put(mk, "A24", "Difference", bold=True); put(mk, "G24", "=G23-G22", fmt=NUM, bold=True, border=TOP)
put(mk, "H24", '=IF(ABS(G24)<1,"✓ Agrees – Q2 & Q3 markup assumed paid; vouch to bank","✗ Investigate")')

section(mk, "A26", "Quarterly analysis – payment status (vouch to bank statements)")
for c in "BCDEFGH": mk[f"{c}26"].fill = GREY
thead(mk, 27, ["Quarter", "From", "To", "Days", "Principal (Rs)", "Rate p.a.", "Markup (Rs)", "Paid per bank? date / amount (AMBER)"])
qs = [("Q1 Jul–Sep 2025", "=B8", "=C8", "=F8"),
      ("Q2 Oct–Dec 2025", "=B9", f"=DATE(YEAR(B9),12,31)", "=F9"),
      ("Q3 Jan–Mar 2026", "=C29+1", "=DATE(YEAR(B30),3,31)", "=F9"),
      ("Q4 Apr–Jun 2026", "=C30+1", "=C9", "=F9")]
for i, (q, b, c, f) in enumerate(qs):
    rr = 28 + i
    put(mk, f"A{rr}", q); put(mk, f"B{rr}", b, fmt=DATEF); put(mk, f"C{rr}", c, fmt=DATEF)
    put(mk, f"D{rr}", f"=C{rr}-B{rr}+1", fmt="0"); put(mk, f"E{rr}", "=$E$8", fmt=NUM)
    put(mk, f"F{rr}", f, fmt=PCT); put(mk, f"G{rr}", f"=E{rr}*F{rr}*D{rr}/{I['BASIS']}", fmt=NUM)
    put(mk, f"H{rr}", None, fill=AMBER, border=BOX)
put(mk, "A32", "Total", bold=True); put(mk, "D32", "=SUM(D28:D31)", fmt="0", bold=True)
put(mk, "G32", "=SUM(G28:G31)", fmt=NUM, bold=True, border=DBL)
put(mk, "H32", '=IF(ABS(G32-G10)<1,"✓ Ties to expected finance cost","✗ Check")')
r = 34
for t in ["Observations:",
          "1. Schedule B prices the new facility only from 1-Oct-25. The Jul–Sep 2025 quarter therefore runs at the old rate (6M KIBOR + 350 bps) – confirm with PMIC (ISA 505) whether markup for that quarter was charged, paid, capitalised into the Sale Price or waived.",
          "2. Trace any PMIC payment around 30-Sep-25 / 1-Oct-25 in the bank statement: if paid and debited to another account (e.g. advances / opening payable), AJE (c) credits that account instead of markup payable.",
          "3. If the novation is treated as a derecognition (A4, AJE (b)), finance cost under IFRS 9 is the EIR unwind on the new loan, not the coupon – see A4 section F and AJE (d) in PF.",
          "4. Mark-up under the Murabaha structure is a financing cost under IFRS 9 (substance over form) – the Shariah legal form does not change amortised-cost accounting."]:
    put(mk, f"A{r}", t, bold=t.endswith(":")); r += 1

# ===================================================================== A4 IFRS 9
md = wb.create_sheet("A4 IFRS9 Modification")
header(md, "A4", "IFRS 9 modification of financial liability – Schedule F novation (IFRS 9 3.3.2, B3.3.6) & ISA 540 estimate",
       [6, 13, 13, 7, 16, 9, 14, 14, 15, 9, 11, 15, 11, 15, 3, 16, 15, 15, 16, 8])

section(md, "A5", "A. Qualitative assessment – IFRS 9 3.3.2 (exchange with substantially different terms)")
for c in range(2, 15): md[f"{CL(c)}5"].fill = GREY
thead(md, 6, ["#", "Factor", "", "", "Old facility (Amendment 27-Aug-24)", "", "", "New facility (TFA 2025)", "", "", "", "Substantial?", "", "Reference"])
qual = [
    ("Legal form", "Conventional KIBOR term finance", "NOVATION – Sale Price Rs 148.75m settles existing facilities; Murabaha", "Yes", "Schedule F"),
    ("Pricing", "6M KIBOR + 350 bps (~14.5%)", "3% → 7% → 6M KIBOR + 100 bps", "Yes", "Schedule B"),
    ("Tenor / profile", "Restructured 2024 profile", "15-month grace; 17 graduated instalments to 30-Sep-30", "Yes", "Schedule B"),
    ("Covenants", "OSS / PAR tested from Jul-25", "OSS from Dec-26; PAR from Dec-29", "Indicator", "Clause 6.2.12"),
    ("Security", "Hypothecation Rs 250m", "Same Letter of Hypothecation", "No", "LoH 11-Dec-17"),
]
r = 7
for i, (fac, old, new, s, ref) in enumerate(qual, 1):
    put(md, f"A{r}", i); put(md, f"B{r}", fac, bold=True)
    put(md, f"E{r}", old, wrap=True); md.merge_cells(f"E{r}:G{r}")
    put(md, f"H{r}", new, wrap=True); md.merge_cells(f"H{r}:K{r}")
    put(md, f"L{r}", s, bold=True); put(md, f"N{r}", ref)
    md.row_dimensions[r].height = 30
    r += 1

r += 1
section(md, f"A{r}", "B. Quantitative 10% test – IFRS 9 B3.3.6 (PV of new cash flows at ORIGINAL effective rate vs carrying amount)")
for c in range(2, 15): md[f"{CL(c)}{r}"].fill = GREY
QB = r + 1
# placeholders, filled after schedule rows known
r += 7

section(md, f"A{r}", "C. Cash-flow schedule of new facility (Schedule B) – DCF and amortised-cost schedule")
for c in range(2, 20): md[f"{CL(c)}{r}"].fill = GREY
r += 1
thead(md, r, ["#", "Period from", "Period to (payment)", "Days", "Opening principal (Rs)", "Rate p.a.",
              "Markup (Rs)", "Principal (Rs)", "Total cash flow (Rs)", "t (yrs)", "DF @ market",
              "PV @ market (Rs)", "DF @ orig. EIR", "PV @ orig. EIR (Rs)", "",
              "Amortised cost – opening (Rs)", "EIR interest (Rs)", "Cash flow (Rs)", "Amortised cost – closing (Rs)", "In FY26? (1/0)"])
hdr = r
pay_dates = [D(2025, 9, 30), D(2025, 12, 31), D(2026, 3, 31), D(2026, 6, 30)] + [d for ds, _ in sched for d in ds]
first = hdr + 1
last = hdr + len(pay_dates)
FV_CELL = f"$L${last+1}"
for i, d in enumerate(pay_dates):
    rr = first + i
    put(md, f"A{rr}", i + 1)
    put(md, f"B{rr}", f"={I['VAL_DATE']}+1" if i == 0 else f"=C{rr-1}+1", fmt=DATEF)
    put(md, f"C{rr}", d, fmt=DATEF, color=BLUE)
    put(md, f"D{rr}", f"=C{rr}-B{rr}+1", fmt="0")
    put(md, f"E{rr}", f"={I['OPEN_FS']}" if i == 0 else f"=E{rr-1}-H{rr-1}", fmt=NUM)
    put(md, f"F{rr}", f"=IF(B{rr}<{I['T1_FROM']},{I['R_OLD']},IF(B{rr}<{I['T2_FROM']},{I['R1']},IF(B{rr}<{I['T3_FROM']},{I['R2']},{I['R3']})))", fmt=PCT)
    put(md, f"G{rr}", f"=E{rr}*F{rr}*D{rr}/{I['BASIS']}", fmt=NUM)
    put(md, f"H{rr}", f"=SUMIF({SPD},C{rr},{SPA})", fmt=NUM, color=GREEN)
    put(md, f"I{rr}", f"=G{rr}+H{rr}", fmt=NUM)
    put(md, f"J{rr}", f"=(C{rr}-{I['VAL_DATE']})/365", fmt="0.000")
    put(md, f"K{rr}", f"=1/(1+{I['MKT']})^J{rr}", fmt="0.0000")
    put(md, f"L{rr}", f"=I{rr}*K{rr}", fmt=NUM)
    put(md, f"M{rr}", f"=1/(1+{I['ORIG_EIR']})^J{rr}", fmt="0.0000")
    put(md, f"N{rr}", f"=I{rr}*M{rr}", fmt=NUM)
    put(md, f"P{rr}", f"={FV_CELL}" if i == 0 else f"=S{rr-1}", fmt=NUM)
    put(md, f"Q{rr}", f"=P{rr}*((1+{I['MKT']})^((C{rr}-B{rr}+1)/365)-1)", fmt=NUM)
    put(md, f"R{rr}", f"=I{rr}", fmt=NUM)
    put(md, f"S{rr}", f"=P{rr}+Q{rr}-R{rr}", fmt=NUM)
    put(md, f"T{rr}", f"=IF(C{rr}<={I['YE']},1,0)", fmt="0", align="center")
tr = last + 1
put(md, f"B{tr}", "Total", bold=True)
for c in "DGHILNQR":
    put(md, f"{c}{tr}", f"=SUM({c}{first}:{c}{last})", fmt=NUM if c != "D" else "0", bold=True, border=DBL)
put(md, f"B{tr+1}", "Checks:", bold=True)
put(md, f"E{tr+1}", "Principal = Schedule B total", italic=True)
put(md, f"I{tr+1}", f'=IF(ABS(H{tr}-{I["FS_TOT"]})<1,"✓","✗")', bold=True)
put(md, f"P{tr+1}", "Closing AC at final date = nil", italic=True)
put(md, f"S{tr+1}", f'=IF(ABS(S{last})<1,"✓","✗")', bold=True)
put(md, f"E{tr+2}", "FY26 contractual markup = A3", italic=True)
put(md, f"I{tr+2}", f"=SUMPRODUCT(T{first}:T{last},G{first}:G{last})", fmt=NUM)
put(md, f"J{tr+2}", f"=IF(ABS(I{tr+2}-'A3 Markup'!G10)<1,\"✓\",\"✗\")", bold=True)
RNG = (first, last)

# --- section B fill
put(md, f"B{QB}", "Carrying amount of old facility at modification date")
put(md, f"L{QB}", f"={I['OPEN_FS']}", fmt=NUM, color=GREEN)
put(md, f"B{QB+1}", "PV of new cash flows at original effective rate (col N)")
put(md, f"L{QB+1}", f"=N{tr}", fmt=NUM)
put(md, f"B{QB+2}", "Difference")
put(md, f"L{QB+2}", f"=L{QB}-L{QB+1}", fmt=NUM, border=TOP)
put(md, f"B{QB+3}", "Difference as % of carrying amount", bold=True)
put(md, f"L{QB+3}", f"=L{QB+2}/L{QB}", fmt=PCT, bold=True)
put(md, f"B{QB+4}", "Conclusion (≥ 10% → substantial modification → extinguish old / recognise new at fair value)", bold=True)
put(md, f"L{QB+4}", f'=IF(L{QB+3}>={I["TEST10"]},"SUBSTANTIAL – derecognise","Not substantial (qualitative still applies)")', bold=True, color="C00000")
put(md, f"N{QB}", "Fees paid to PMIC on restructuring (B3.3.6A) – nil assumed; confirm", italic=True, fill=AMBER)

r = tr + 4
section(md, f"A{r}", "D. Fair value of new loan and gain on extinguishment (IFRS 9 3.3.3, 5.1.1; IFRS 13 Level 3)")
for c in range(2, 15): md[f"{CL(c)}{r}"].fill = GREY
r += 1
D0 = r
lines = [
    ("Carrying amount of old facility derecognised", f"={I['OPEN_FS']}"),
    ("Fair value of new loan = PV of Schedule B cash flows at market rate", f"={FV_CELL}"),
    ("Gain on extinguishment – recognised in FY2025 (prior-period, IAS 8)", f"=L{D0}-L{D0+1}"),
    ("Market (discount) rate used", f"={I['MKT']}"),
    ("Funds at 30-Jun-25 as reported", f"={I['FUNDS_PY']}"),
    ("Funds at 30-Jun-25 restated", f"=L{D0+4}+L{D0+2}"),
    ("NBFC minimum equity floor", f"={I['FLOOR']}"),
    ("Headroom / (shortfall) after restatement", f"=L{D0+5}-L{D0+6}"),
    ("Headroom / (shortfall) excluding Rs 10m director sub-loan (IAS 32 risk)", f"=L{D0+7}-{I['SUBLOAN']}"),
]
for i, (lbl, f) in enumerate(lines):
    rr = D0 + i
    put(md, f"B{rr}", lbl, bold=i in (2, 5))
    put(md, f"L{rr}", f, fmt=PCT if i == 3 else NUM, bold=i in (2, 5),
        border=DBL if i in (2, 5) else None)
GAIN = f"'A4 IFRS9 Modification'!$L${D0+2}"
FUNDS_RESTATED = f"'A4 IFRS9 Modification'!$L${D0+5}"
put(md, f"N{D0+2}", "Planning estimate ≈ Rs 24m; recomputed at inputs above – see sensitivity (E)", italic=True, wrap=True)
md.row_dimensions[D0+2].height = 28
put(md, f"N{D0+7}", f'=IF(L{D0+7}>=0,"Floor met after restatement","Floor NOT met")', bold=True)
put(md, f"N{D0+8}", f'=IF(L{D0+8}>=0,"Met","Not met if sub-loan is a liability")', bold=True)

r = D0 + len(lines) + 1
section(md, f"A{r}", "E. Sensitivity of fair value / gain to the market rate (ISA 540.18 – range of outcomes)")
for c in range(2, 15): md[f"{CL(c)}{r}"].fill = GREY
r += 1
thead(md, r, ["", "Market rate", "", "", "Fair value (Rs)", "", "", "Gain (Rs)", "", "", "", "Funds 30-Jun-25 restated (Rs)", "", "vs floor"])
r += 1
for rate in [0.13, 0.14, 0.15, 0.16, 0.17]:
    put(md, f"B{r}", rate, fmt=PCT, color=BLUE)
    put(md, f"E{r}", f"=SUMPRODUCT(I${first}:I${last}/(1+B{r})^J${first}:J${last})", fmt=NUM)
    put(md, f"H{r}", f"={I['OPEN_FS']}-E{r}", fmt=NUM)
    put(md, f"L{r}", f"={I['FUNDS_PY']}+H{r}", fmt=NUM)
    put(md, f"N{r}", f'=IF(L{r}>={I["FLOOR"]},"Above","Below")')
    r += 1
put(md, f"B{r}", "Alternative – if judged NOT substantial: carrying amount reset to PV at original EIR (IFRS 9 B5.4.6 analogy, IASB BC4.252–253); gain =", italic=True)
put(md, f"L{r}", f"={I['OPEN_FS']}-N{tr}", fmt=NUM, italic=True)
r += 2

section(md, f"A{r}", "F. Consequential FY2026 effect – amortised cost of the new loan (EIR = market rate)")
for c in range(2, 15): md[f"{CL(c)}{r}"].fill = GREY
r += 1
F0 = r
flines = [
    ("Amortised cost at 1-Jul-25 (fair value)", f"={FV_CELL}"),
    ("EIR interest FY2026 (col Q, FY26 rows)", f"=SUMPRODUCT(T{first}:T{last},Q{first}:Q{last})"),
    ("Contractual markup FY2026 (col G, FY26 rows) – = A3 expected", f"=SUMPRODUCT(T{first}:T{last},G{first}:G{last})"),
    ("Principal repaid FY2026", f"=SUMPRODUCT(T{first}:T{last},H{first}:H{last})"),
    ("Amortised cost at 30-Jun-26 (excl. markup payable)", f"=SUMIF(C{first}:C{last},{I['YE']},S{first}:S{last})"),
    ("Nominal principal at 30-Jun-26", f"={I['FS_TOT']}"),
    ("Unamortised discount at 30-Jun-26", f"=L{F0+5}-L{F0+4}"),
    ("Finance cost booked – Note 23", f"={I['FC_BOOK']}"),
    ("Total finance cost adjustment under IFRS 9 (= AJE (c) + AJE (d))", f"=L{F0+1}-L{F0+7}"),
    ("  of which AJE (c): missing Jul–Sep 2025 contractual markup", f"='A3 Markup'!G12"),
    ("  of which AJE (d): EIR unwind in excess of contractual markup", f"=L{F0+1}-L{F0+2}"),
]
for i, (lbl, f) in enumerate(flines):
    rr = F0 + i
    put(md, f"B{rr}", lbl, bold=i in (4, 8))
    put(md, f"L{rr}", f, fmt=NUM, bold=i in (4, 8), border=DBL if i in (4, 8) else None,
        color=GREEN if "A3" in f else BLACK)
put(md, f"N{F0+4}", f'=IF(ABS(L{F0}+L{F0+1}-L{F0+2}-L{F0+3}-L{F0+4})<1,"✓ roll-forward ties","✗")', bold=True)
EIR_INT = f"'A4 IFRS9 Modification'!$L${F0+1}"
AC_CLOSE = f"'A4 IFRS9 Modification'!$L${F0+4}"
AJE_D = f"'A4 IFRS9 Modification'!$L${F0+10}"
CONTR = f"'A4 IFRS9 Modification'!$L${F0+2}"
r = F0 + len(flines) + 1
for t in ["Conclusion & ISA 540 matters:",
          "1. Both the qualitative factors (novation under Schedule F, change from conventional to Murabaha structure, rate cut ~14.5% → 3%/7%) and the 10% test indicate a SUBSTANTIAL modification: derecognise the old liability and recognise the new one at fair value (IFRS 9 3.3.2, 3.3.3, 5.1.1).",
          "2. The difference between carrying amount and fair value is a gain arising on ~24-30 Jun 2025 → belongs to FY2025. Not recognised by the client → prior-period error (IAS 8.41-49); restate comparatives and opening funds (ISA 510, ISA 710).",
          "3. Key estimate inputs (AMBER): market rate, forward KIBOR, modification date, any fees. Evaluate management's point estimate vs RSM range (section E) – ISA 540.18/.29; consider management bias.",
          "4. Below-market element: PMIC is a government-sponsored apex lender – consider whether the benefit is a government grant (IAS 20.10A) rather than a gain; the measurement of the liability is the same, presentation of the credit differs. Document judgement.",
          "5. Consequential: FY2026 finance cost on EIR basis is materially higher than the coupon; the net equity effect at 30-Jun-26 is shown in PF."]:
    put(md, f"A{r}", t, bold=t.endswith(":")); r += 1
r += 1
section(md, f"A{r}", "G. How amortised cost works – explanation of the mechanics (IFRS 9 4.2.1, 5.1.1, App. A 'effective interest method')")
for c in range(2, 15): md[f"{CL(c)}{r}"].fill = GREY
r += 1
expl = [
    "Step 1 – Day one: ALL future contractual cash flows of the new loan (Schedule B principal instalments AND markup) are discounted to present value. The discount rate is the MARKET rate, not the contract rate, because the novation creates a NEW loan that must be recognised at fair value (IFRS 9 5.1.1).",
    "Step 2 – That present value is the opening carrying amount. The difference from the old carrying amount (Rs 148.75m) is the gain on extinguishment (section D / AJE b).",
    "Step 3 – Each period: EIR interest = opening carrying amount × EIR (= market rate, as no fees were paid; any fees would change the EIR – B3.3.6A / B5.4.1-B5.4.3). This is the finance cost charged to P&L.",
    "Step 4 – Closing carrying amount = opening + EIR interest – cash paid (markup + principal). When the coupon is below the EIR the unpaid part of the interest is added to the loan, so the carrying amount INCREASES; principal repayments reduce it.",
    "Step 5 – By the final instalment (30-Sep-2030) the carrying amount unwinds to exactly nil (check in col S). The total EIR interest over the life = total cash paid – initial fair value.",
]
for t in expl:
    put(md, f"B{r}", t, wrap=True); md.merge_cells(f"B{r}:N{r}")
    md.row_dimensions[r].height = 30
    r += 1
r += 1
thead(md, r, ["", "FY2026 quarter", "", "", "Opening carrying (Rs)", "", "EIR interest – P&L (Rs)", "", "Cash markup (Rs)", "", "", "Closing carrying (Rs)", "", "Comment"])
r += 1
G0 = r
qnames = ["Jul–Sep 2025", "Oct–Dec 2025", "Jan–Mar 2026", "Apr–Jun 2026"]
qcom = ["Carrying FALLS: old-rate coupon (14.5% on Rs 148.75m) exceeds EIR (15% on fair value)",
        "Carrying RISES: 3% coupon < 15% EIR – shortfall added to the loan",
        "Carrying RISES – as above", "Carrying RISES – closing = amortised cost at 30-Jun-26"]
for i in range(4):
    sr = first + i
    put(md, f"B{r}", qnames[i])
    put(md, f"E{r}", f"=P{sr}", fmt=NUM); put(md, f"G{r}", f"=Q{sr}", fmt=NUM)
    put(md, f"I{r}", f"=R{sr}", fmt=NUM); put(md, f"L{r}", f"=S{sr}", fmt=NUM)
    put(md, f"N{r}", qcom[i], wrap=True); md.row_dimensions[r].height = 28
    r += 1
put(md, f"B{r}", "FY2026 total", bold=True)
put(md, f"E{r}", f"=E{G0}", fmt=NUM, bold=True, border=DBL)
put(md, f"G{r}", f"=SUM(G{G0}:G{r-1})", fmt=NUM, bold=True, border=DBL)
put(md, f"I{r}", f"=SUM(I{G0}:I{r-1})", fmt=NUM, bold=True, border=DBL)
put(md, f"L{r}", f"=L{r-1}", fmt=NUM, bold=True, border=DBL)
put(md, f"N{r}", f'=IF(ABS(E{r}+G{r}-I{r}-L{r})<1,"✓ opening + interest – cash = closing","✗")', bold=True)
r += 1
put(md, f"B{r}", "Interest not paid in cash, added to the loan = AJE (d): Dr Finance cost / Cr Financing")
put(md, f"L{r}", f"=G{r-1}-I{r-1}", fmt=NUM, bold=True)
put(md, f"N{r}", f'=IF(ABS(L{r}-{AJE_D})<1,"✓ = AJE (d) in PF","✗")', bold=True)
r += 1
put(md, f"B{r}", "Unamortised discount at 30-Jun-26 = nominal Rs 148.75m – closing carrying (shown in Lead, column 'AJE (b)+(d)')")
put(md, f"L{r}", f"={I['FS_TOT']}-L{r-2}", fmt=NUM, bold=True)
put(md, f"N{r}", f'=IF(ABS(L{r}-L{F0+6})<1,"✓ = section F","✗")', bold=True)
r += 2
put(md, f"B{r}", "Conclusion: under IFRS 9 the PMIC loan is carried at amortised cost of the amount in section F, not at the nominal Rs 148.75m; FY2026 finance cost is the EIR interest above, not the 3% coupon. AJE (d) applies only if AJE (b) (derecognition at fair value) is accepted.", bold=True, wrap=True)
md.merge_cells(f"B{r}:N{r}"); md.row_dimensions[r].height = 30
md.freeze_panes = f"C{hdr+1}"

# ===================================================================== A5 covenants
cv = wb.create_sheet("A5 Covenants")
header(cv, "A5", "Covenant compliance – TFA 2025 clauses 6.2.12 & 6.2.13; NBFC minimum equity (ISA 250 / ISA 570)",
       [30, 14, 22, 16, 16, 16, 24, 50])
thead(cv, 6, ["Covenant", "Clause", "Requirement", "Per client", "Per audit – draft FS", "Per audit – after AJEs", "Result", "Comment / follow-up"])
# helper calcs
put(cv, "A20", "Supporting calculations", bold=True, fill=GREY)
for c in "BCDEFGH": cv[f"{c}20"].fill = GREY
calc = [
    ("Debt – PMIC facility, nominal", f"={I['FS_TOT']}"),
    ("Debt – PMIC facility, amortised cost after AJE (b)/(d)", f"={AC_CLOSE}"),
    ("Funds 30-Jun-26 – draft", f"={I['FUNDS_CY']}"),
    ("Funds 30-Jun-26 – after AJE (b), (c), (d)", "='PF Points Forward'!F31"),
    ("Current liabilities after AJE (a) (and (c) if unpaid)", f"=IF(ISBLANK({I['CLI']}),0,{I['CLI']}+'A2 CNC Split'!F{SP_CMP}+'A3 Markup'!G12)"),
    ("Minimum provision required = MAX(5% × portfolio, regulatory)", f"=IF(ISBLANK({I['PORT']}),0,MAX(0.05*{I['PORT']},N({I['PROV_REG']})))"),
]
for i, (lbl, f) in enumerate(calc):
    put(cv, f"A{21+i}", lbl); put(cv, f"E{21+i}", f, fmt=NUM, color=GREEN)
covs = [
    ("Current ratio", "6.2.12", "≥ 1.00 : 1",
     f"={I['CR_CLIENT']}",
     f'=IF(OR(ISBLANK({I["CA"]}),ISBLANK({I["CLI"]})),"Input CA/CL",{I["CA"]}/{I["CLI"]})',
     f'=IF(OR(ISBLANK({I["CA"]}),E25=0),"Input CA/CL",{I["CA"]}/E25)',
     '=IF(ISNUMBER(F7),IF(F7>=1,"Pass","BREACH"),IF(D7>=1,"Pass (per client)","BREACH"))',
     "Reclassification AJE (a) reduces current liabilities by Rs 29m → ratio improves. Enter CA/CL (amber, Inputs) to recompute."),
    ("Debt-to-equity", "6.2.12", "< 8.0x", None,
     "=E21/E23", "=E22/E24",
     '=IF(MAX(E8,F8)<8,"Pass","BREACH")',
     "Confirm definition of 'debt' and 'equity' in the agreement (does equity include the Rs 10m sub-loan?)."),
    ("Operational self-sufficiency (OSS)", "6.2.12", "> 100% from Dec-2026",
     f"={I['OSS_FS']}", f"={I['OSS_CALC']}", None,
     f'=IF({I["YE"]}<DATE(2026,12,31),IF(E9>1,"Not yet tested","Not yet tested – AT RISK"),IF(E9>1,"Pass","BREACH"))',
     "FS shows 141% but income ÷ expenditure = 86% → obtain client OSS computation and reconcile definition. First test Dec-26 (within 12m): disclose under IAS 1.76ZA; ISA 570 indicator."),
    ("PAR > 29 days", "6.2.12", "< 5% from Dec-2029", None, None, None, '="Not applicable at 30-Jun-26"',
     "Monitor; 2024 Amendment required compliance from Jul-25 – confirm the 2025 novation superseded it (Schedule F)."),
    ("Minimum loan-loss provision", "6.2.12", "≥ higher of 5% of portfolio or NBFC-regulatory provision",
     f'=IF(ISBLANK({I["PROV_FS"]}),"Pending",{I["PROV_FS"]})', '=IF(E26=0,"Pending",E26)', None,
     f'=IF(ISBLANK({I["PORT"]}),"Pending LLP WP",IF(N({I["PROV_FS"]})>=E26,"Pass","BREACH"))',
     "Ties to loan-loss provision WP – any shortfall is both a misstatement (NBFC Regs) and a covenant breach."),
    ("Sub-loan Rs 10m", "6.2.13", "Arranged before 30-Jun-2025",
     f"={I['SUBLOAN']}", f'=IF(ISBLANK({I["SUB_DATE"]}),"Pending",{I["SUB_DATE"]})', None,
     f'=IF(ISBLANK({I["SUB_DATE"]}),"Pending – date of receipt",IF({I["SUB_DATE"]}<=DATE(2025,6,30),"Pass","BREACH"))',
     "Explains the Rs 10m director loan in funds. Assess IAS 32 classification (equity only if no contractual obligation to repay)."),
    ("NBFC minimum equity", "NBFC Regs 2008", "≥ Rs 50,000,000",
     f"={I['FUNDS_CY']}", "=E23", "=E24",
     f'=IF(F13>={I["FLOOR"]},IF(E13>={I["FLOOR"]},"Pass","Breach resolved by AJEs"),"BREACH")',
     "Draft funds Rs 45.4m below floor. After AJEs (b)-(d) – see PF. ISA 250: non-compliance with NBFC Regs; consider SECP reporting and ISA 570."),
]
for i, row in enumerate(covs):
    rr = 7 + i
    for j, v in enumerate(row):
        col = CL(j + 1)
        if v is None: continue
        fmt = None
        if j in (3, 4, 5):
            fmt = '0.00"x"' if i in (0, 1) else PCT if i == 2 else DATEF if (i == 5 and j == 4) else NUM
        put(cv, f"{col}{rr}", v, fmt=fmt, wrap=(j in (2, 7)), bold=(j == 6),
            color=GREEN if (isinstance(v, str) and "Inputs!" in v) else BLACK)
    cv.row_dimensions[rr].height = 52
put(cv, "A15", "Overall: obtain PMIC confirmation (ISA 505) stating whether any covenant was breached at/before 30-Jun-26 and whether a waiver was granted BEFORE the reporting date (IAS 1.74–75).", bold=True, color="C00000")

# ===================================================================== Lead
header(lead, "A", "Lead schedule – Financing for microcredit lending – secured (PMIC)",
       [36, 9, 17, 17, 15, 16, 17, 16, 17, 17, 12])
thead(lead, 6, ["Line item", "FS note", "Per draft FS 30-Jun-26", "Per TB / GL 30-Jun-26", "Difference FS – GL",
                "AJE (a) reclass", "AJE (b)+(d) IFRS 9", "AJE (c) markup", "Audited 30-Jun-26 (proposed)", "Audited 30-Jun-25 (opening)", "WP ref"])
put(lead, "A7", "Financing – non-current"); put(lead, "B7", "14")
put(lead, "C7", f"={I['FS_NC']}", fmt=NUM, color=GREEN); put(lead, "D7", f"={I['GL_NC']}", fmt=NUM, color=GREEN)
put(lead, "F7", f"=-'A2 CNC Split'!F{SP_CMP}", fmt=NUM, color=GREEN)
put(lead, "G7", f"={AC_CLOSE}-{I['FS_TOT']}", fmt=NUM, color=GREEN)
put(lead, "K7", "A2 / A4")
put(lead, "A8", "Financing – current portion"); put(lead, "B8", "14")
put(lead, "C8", f"={I['FS_CUR']}", fmt=NUM, color=GREEN); put(lead, "D8", f"={I['GL_CUR']}", fmt=NUM, color=GREEN)
put(lead, "F8", f"={SP_RECLASS}", fmt=NUM, color=GREEN); put(lead, "K8", "A2")
put(lead, "A9", "Total financing – PMIC", bold=True)
for rr in (7, 8):
    put(lead, f"E{rr}", f"=C{rr}-D{rr}", fmt=NUM)
    put(lead, f"I{rr}", f"=C{rr}+SUM(F{rr}:H{rr})", fmt=NUM)
for c in "CDEFGHI":
    put(lead, f"{c}9", f"=SUM({c}7:{c}8)", fmt=NUM, bold=True, border=DBL)
put(lead, "J9", f"={I['OPEN_PY']}", fmt=NUM, bold=True, color=GREEN, border=DBL)
put(lead, "K9", "A1")
put(lead, "A11", "Related balances", bold=True, fill=GREY)
for c in "BCDEFGHIJK": lead[f"{c}11"].fill = GREY
put(lead, "A12", "Markup payable – PMIC"); put(lead, "B12", "15")
put(lead, "C12", f"={I['MP_BOOK']}", fmt=NUM, color=GREEN); put(lead, "D12", f"={I['MP_BOOK']}", fmt=NUM, color=GREEN)
put(lead, "E12", "=C12-D12", fmt=NUM)
put(lead, "H12", "='A3 Markup'!G12", fmt=NUM, color=GREEN,
    comment="Credit to markup payable assumes the Jul–Sep 2025 markup is unpaid. If paid, credit the account debited at payment.")
put(lead, "I12", "=C12+SUM(F12:H12)", fmt=NUM); put(lead, "K12", "A3")
put(lead, "A13", "Finance cost – PMIC (P&L)"); put(lead, "B13", "23")
put(lead, "C13", f"={I['FC_BOOK']}", fmt=NUM, color=GREEN); put(lead, "D13", f"={I['FC_BOOK']}", fmt=NUM, color=GREEN)
put(lead, "E13", "=C13-D13", fmt=NUM)
put(lead, "G13", f"={AJE_D}", fmt=NUM, color=GREEN)
put(lead, "H13", "='A3 Markup'!G12", fmt=NUM, color=GREEN)
put(lead, "I13", "=C13+SUM(F13:H13)", fmt=NUM); put(lead, "K13", "A3 / A4")
put(lead, "A15", "Tie-outs", bold=True, fill=GREY)
for c in "BCDEFGHIJK": lead[f"{c}15"].fill = GREY
ties = [
    ("FS vs GL difference (must be nil)", "=E9", '=IF(ABS(B16)<1,"✓ Nil","✗ Investigate")'),
    ("Closing nominal per A1 roll-forward – FS", f"='A1 Movement'!C13-C9", '=IF(ABS(B17)<1,"✓ Nil","✗")'),
    ("Schedule B total – FS total", f"='A2 CNC Split'!C{TOT}-C9", '=IF(ABS(B18)<1,"✓ Nil","✗")'),
    ("Reclass AJE (a) nets to nil", "=F9", '=IF(ABS(B19)<1,"✓ Nil","✗")'),
    ("Audited finance cost – EIR interest (A4)", f"=I13-{EIR_INT}", '=IF(ABS(B20)<1,"✓ Nil","✗")'),
    ("PMIC confirmation – audited principal (ISA 505)", f'=IF(ISBLANK({I["CONF_BAL"]}),"Pending",{I["CONF_BAL"]}-C9)', '=IF(ISNUMBER(B21),IF(ABS(B21)<1,"✓ Agreed","✗ Reconcile"),"AMBER – awaiting")'),
]
for i, (lbl, f, chk) in enumerate(ties):
    rr = 16 + i
    put(lead, f"A{rr}", lbl); put(lead, f"B{rr}", f, fmt=NUM); put(lead, f"C{rr}", chk, bold=True)
put(lead, "A23", "Notes:", bold=True)
put(lead, "A24", "• AJE (b)+(d) column shows the IFRS 9 discount on the new loan (fair value at 30-Jun-25 unwound at EIR to 30-Jun-26). Contractual principal remains Rs 148.75m; classification of the nominal amount per A2.")
put(lead, "A25", "• Opening column: agree to signed FY2025 audited FS (predecessor) – ISA 510. If AJE (b) is booked, the restated opening carrying amount is the fair value in A4.")
put(lead, "A26", "• Colour code: blue = hard input; green = link to another sheet; black = formula; amber fill = awaiting PMIC confirmation / bank / market data.")
lead.freeze_panes = "B7"

# ===================================================================== A1 movement
header(mv, "A1", "Movement / roll-forward and opening balance (ISA 510 – first-year audit)",
       [52, 14, 18, 18, 46])
thead(mv, 6, ["Item", "Ref", "Amount (Rs)", "Check", "Explanation / contract reference"])
put(mv, "A7", "Opening balance 1-Jul-25 – per FY2025 AUDITED FS (predecessor)")
put(mv, "C7", f"={I['OPEN_PY']}", fmt=NUM, color=GREEN, fill=AMBER)
put(mv, "E7", "Agree to signed FY2025 audited FS & predecessor's file (ISA 510.6)", wrap=True)
put(mv, "A8", "Opening balance 1-Jul-25 – per draft FY2026 FS comparative")
put(mv, "C8", f"={I['OPEN_FS']}", fmt=NUM, color=GREEN)
put(mv, "A9", "Difference", italic=True); put(mv, "C9", "=C8-C7", fmt=NUM, border=TOP)
put(mv, "D9", '=IF(ABS(C9)<1,"✓ Nil","✗")', bold=True)
put(mv, "A10", "Opening balance used for roll-forward", bold=True); put(mv, "C10", "=C8", fmt=NUM, bold=True)
put(mv, "A11", "Add: drawdowns received"); put(mv, "C11", f"={I['DRAW']}", fmt=NUM, color=GREEN)
put(mv, "E11", "No disbursement – Sale Price of Rs 148.75m under the 2025 TFA settled the existing facilities on a non-cash basis (Schedule F novation)", wrap=True)
put(mv, "A12", "Less: principal repaid"); put(mv, "C12", f"=-{I['REPAY']}", fmt=NUM, color=GREEN)
put(mv, "E12", "No instalment due: Principal Deferment Period 30-Jun-24 – 29-Jun-25 (Amendment 27-Aug-24) followed by 15-month grace; first instalment 30-Sep-26 (Schedule B)", wrap=True)
put(mv, "A13", "Closing balance 30-Jun-26 – computed", bold=True); put(mv, "C13", "=SUM(C10:C12)", fmt=NUM, bold=True, border=DBL)
put(mv, "A14", "Closing balance per draft FS (Note 14)"); put(mv, "C14", f"={I['FS_TOT']}", fmt=NUM, color=GREEN)
put(mv, "A15", "Difference", italic=True); put(mv, "C15", "=C13-C14", fmt=NUM, border=TOP)
put(mv, "D15", '=IF(ABS(C15)<1,"✓ Nil","✗")', bold=True)
put(mv, "A16", "Principal due in FY2026 under Schedule B (should be nil)")
put(mv, "C16", f"=SUMIFS({SPA},{SPD},\">\"&{I['PYE']},{SPD},\"<=\"&{I['YE']})", fmt=NUM)
put(mv, "D16", '=IF(C16=-C12,"✓ Consistent","✗")', bold=True)
for rr in (11, 12): mv.row_dimensions[rr].height = 42
mv.row_dimensions[7].height = 28
put(mv, "A18", "Reconciliation of facility history", bold=True, fill=GREY)
for c in "BCDE": mv[f"{c}18"].fill = GREY
put(mv, "A19", "TFA I (~Mar-21)"); put(mv, "C19", 171250000, fmt=NUM, color=BLUE)
put(mv, "A20", "TFA II (~2022)"); put(mv, "C20", 129000000, fmt=NUM, color=BLUE)
put(mv, "A21", "Total facilities sanctioned"); put(mv, "C21", "=C19+C20", fmt=NUM, border=TOP)
put(mv, "A22", "Restructured outstanding (Amendment 27-Aug-24)"); put(mv, "C22", 148750000, fmt=NUM, color=BLUE)
put(mv, "A23", "Implied repaid / undrawn before restructuring"); put(mv, "C23", "=C21-C22", fmt=NUM)
put(mv, "E23", "Obtain TFA I & II, drawdown and repayment history to confirm (completeness of liability).", wrap=True)
put(mv, "A24", "Sale Price under 2025 TFA (Schedule F)"); put(mv, "C24", f"={I['OPEN_FS']}", fmt=NUM, color=GREEN)
put(mv, "D24", '=IF(C24=C22,"✓ = restructured amount","✗")', bold=True)
r = 26
for t in ["ISA 510 procedures (first-year audit):",
          "1. Read FY2025 audited FS & predecessor's report – any modification / emphasis of matter on going concern or NBFC equity?",
          "2. Review predecessor's working papers on borrowings (ISA 510.6(c)) – did they consider the June 2025 novation and IFRS 9 modification?",
          "3. Opening balance contains a misstatement if the 2025 modification gain was not recognised (A4) → IAS 8 restatement; communicate with management/TCWG (ISA 510.7, ISA 710).",
          "4. Confirm opening principal and terms with PMIC (ISA 505) as part of the year-end confirmation.",
          "Conclusion: Nil movement in principal is consistent with the contractual deferment/grace; opening balance agrees to prior-year audited FS subject to amber items."]:
    put(mv, f"A{r}", t, bold=t.endswith(":") or t.startswith("Conclusion")); r += 1

# ===================================================================== Standards
st = wb.create_sheet("Standards & ISA")
header(st, "STD", "Applicable financial reporting standards, regulations and ISAs",
       [24, 16, 58, 62, 12])
thead(st, 6, ["Standard / ISA", "Paragraph(s)", "Requirement", "Application to PMIC facility", "WP ref"])
std = [
    ("IFRS 9", "4.2.1, 5.1.1, App. A", "Financial liabilities at amortised cost; initial recognition at fair value; effective interest method.", "New Murabaha facility is a financial liability at amortised cost; recognised at FV on novation; finance cost = EIR unwind.", "A3 / A4"),
    ("IFRS 9", "3.3.1–3.3.3", "Derecognise a liability when extinguished; an exchange with substantially different terms is an extinguishment; difference to P&L.", "Schedule F novation – old KIBOR facilities settled by Sale Price → extinguish and recognise new loan; gain in FY2025.", "A4"),
    ("IFRS 9", "B3.3.6, B3.3.6A", "10% test: PV of new cash flows (net of fees) at original EIR vs carrying amount; fees included only between borrower and lender.", "PV at original ~14.5% vs Rs 148.75m – see A4 section B.", "A4"),
    ("IFRS 9", "B5.4.6 / BC4.252–253", "Non-substantial modification of a liability: adjust carrying amount to PV of revised cash flows at original EIR; gain/loss in P&L.", "Alternative outcome if derecognition rejected – gain still arises (A4 section E memo).", "A4"),
    ("IFRS 13", "9, 72–90", "Fair value = exit price; Level 3 inputs where unobservable.", "Market rate for SSSF unsubsidised borrowing and forward KIBOR are Level 3 inputs.", "A4"),
    ("IAS 20", "10A", "Benefit of a government loan at below-market rate is a government grant.", "Consider whether PMIC concession is a government grant – presentation of the credit.", "A4"),
    ("IAS 1", "69, 72A, 72B, 74–76, 76ZA", "Current if due within 12 months; right to defer must exist at reporting date; covenants after reporting date do not affect classification but must be disclosed; breach without pre-YE waiver → current.", "Current = Rs 6.0m not Rs 35m (AJE (a)); OSS covenant (Dec-26) → 76ZA disclosure.", "A2 / A5"),
    ("IAS 8", "41–49", "Correction of material prior-period errors by retrospective restatement.", "Unrecognised FY2025 modification gain → restate opening funds.", "A4 / PF"),
    ("IAS 32", "11, 16", "Equity vs liability classification.", "Rs 10m director sub-loan in funds – equity only if no contractual obligation to deliver cash.", "A5"),
    ("IFRS 7", "7, 18–19, 39(a)", "Significance of financial instruments; defaults & breaches; maturity analysis.", "Disclose terms, security, covenant status and contractual maturity (Schedule B).", "PF"),
    ("NBFC Regulations 2008", "Minimum equity & prudential regulations (confirm refs)", "Minimum equity; provisioning for microfinance portfolio.", "Funds Rs 45.4m < Rs 50m floor; minimum provision covenant links to NBFC provisioning.", "A5"),
    ("Companies Act 2017", "s.42; s.100 et seq.; Fourth Schedule", "Section 42 licence conditions; registration/modification of charges with SECP; disclosure of secured borrowings.", "Confirm SECP registration of the Rs 250m hypothecation charge and modification after the 2025 novation.", "Inputs / PF"),
    ("ISA 505", "7–16", "External confirmation – control over requests; responses; exceptions.", "Direct confirmation to PMIC: balance, markup, terms, security, covenant breaches/waivers.", "PF"),
    ("ISA 540 (Revised)", "13, 17–18, 21, 29, 32", "Risk assessment of estimates; evaluate method, assumptions and data; develop auditor range; management bias.", "Fair value of new loan (market rate, KIBOR curve, modification date).", "A4"),
    ("ISA 510", "6–8", "Opening balances – sufficient appropriate evidence; misstatements affecting current period.", "Agree opening to FY2025 audited FS; unrecorded FY2025 gain is an opening-balance misstatement.", "A1"),
    ("ISA 710", "7–8", "Comparative information – prior-period misstatement.", "Restated comparatives if AJE (b) booked.", "PF"),
    ("ISA 250 (Revised)", "14, 19–22, 29", "Compliance with laws and regulations; reporting to regulators where required.", "NBFC minimum equity breach; consider SECP reporting.", "A5"),
    ("ISA 570 (Revised)", "10, 16–20, 22–24", "Going concern – events/conditions; management's assessment; disclosures.", "Equity below floor; OSS 86% vs >100% from Dec-26; dependence on PMIC concession.", "A5 / PF"),
    ("ISA 450", "5–12", "Accumulate and communicate misstatements.", "AJEs (a)–(d) to summary of misstatements.", "PF"),
    ("ISA 330 / 500", "—", "Responses to assessed risks; sufficient appropriate evidence.", "Audit programme on 'Audit Programme' tab.", "AP"),
]
for i, row in enumerate(std):
    rr = 7 + i
    for j, v in enumerate(row):
        put(st, f"{CL(j+1)}{rr}", v, wrap=True, bold=(j == 0))
    st.row_dimensions[rr].height = 45
st.freeze_panes = "A7"

# ===================================================================== Audit programme
ap = wb.create_sheet("Audit Programme")
header(ap, "AP", "Audit programme – PMIC financing", [6, 20, 70, 20, 10, 14, 14, 30])
thead(ap, 6, ["#", "Assertion", "Procedure", "ISA / standard", "WP ref", "Done by / date", "Status", "Conclusion / exceptions"])
prog = [
    ("Existence / rights & obligations", "Send PMIC confirmation (ISA 505) requesting: principal 30-Jun-26, markup payable, rates applied (incl. Jul–Sep 2025), repayment schedule, security, covenant breaches and waivers, first Disbursement Date.", "ISA 505", "PF"),
    ("Completeness / accuracy", "Agree FS Note 14 to TB/GL; tie lead schedule (difference nil).", "ISA 500", "A"),
    ("Opening balances", "Agree opening Rs 148.75m to signed FY2025 audited FS; review predecessor's WPs on borrowings and restructuring.", "ISA 510", "A1"),
    ("Completeness", "Roll forward: confirm nil drawdowns/repayments to bank statements; explain by deferment (Amendment 27-Aug-24) and grace (Schedule B).", "ISA 500", "A1"),
    ("Rights & obligations", "Obtain and read TFA I (2021), TFA II (2022), Amendment Agreement 27-Aug-24 (and request letter 12-Jun-24) and TFA 2025 incl. Schedules B and F; document key terms.", "ISA 500", "Inputs"),
    ("Classification", "Recompute current portion from Schedule B (12 months after 30-Jun-26) – propose reclassification AJE (a).", "IAS 1.69–76ZA", "A2"),
    ("Accuracy / cut-off", "Recompute finance cost for full 12 months (Jul–Sep 2025 at old rate; Oct–Jun at 3%) and markup payable; vouch markup payments to bank.", "IFRS 9 EIR", "A3"),
    ("Accuracy / valuation", "Confirm 6M KIBOR reset for Jul–Sep 2025 quarter from PMIC / SBP published rates.", "ISA 500", "A3"),
    ("Valuation / presentation", "Assess Schedule F novation under IFRS 9 3.3.2 (qualitative + 10% test); conclude on derecognition.", "IFRS 9 / ISA 540", "A4"),
    ("Valuation (estimate)", "Evaluate market rate, KIBOR forecast and modification date; build auditor range; consider management bias.", "ISA 540", "A4"),
    ("Opening balances / comparatives", "Quantify FY2025 modification gain and EIR unwind; propose IAS 8 restatement; communicate with TCWG.", "ISA 510 / 710 / IAS 8", "A4 / PF"),
    ("Presentation", "Consider IAS 20 government-grant view of the PMIC concession.", "IAS 20.10A", "A4"),
    ("Compliance", "Recompute covenants (6.2.12) at 30-Jun-26 pre- and post-AJE; obtain client covenant certificate; reconcile OSS 141% vs 86%.", "ISA 250", "A5"),
    ("Compliance", "Verify sub-loan Rs 10m arranged before 30-Jun-25 (6.2.13); assess IAS 32 classification.", "IAS 32", "A5"),
    ("Compliance", "Assess NBFC minimum equity before/after AJEs; consider reporting to SECP.", "ISA 250 / NBFC Regs", "A5 / PF"),
    ("Going concern", "Evaluate effect of equity shortfall and forward OSS covenant on going concern; review cash-flow forecast to Dec-26 covering instalments from 30-Sep-26.", "ISA 570", "PF"),
    ("Rights & obligations", "Inspect Letter of Hypothecation (11-Dec-17 + supplements) and SECP charge certificate; confirm charge modified post-novation.", "Companies Act 2017", "Inputs"),
    ("Presentation & disclosure", "Review Note 14/15/23 disclosures: terms, rates, security, maturity analysis, covenants (IAS 1.76ZA, IFRS 7.18–19, 39).", "IAS 1 / IFRS 7", "PF"),
    ("Subsequent events", "Verify payment of 30-Sep-26 instalment (Rs 1.5m) and markup after year end.", "ISA 560", "PF"),
    ("Misstatements", "Post AJEs (a)–(d) to summary of uncorrected misstatements; evaluate vs materiality.", "ISA 450", "PF"),
]
for i, (a, p, isa, ref) in enumerate(prog, 1):
    rr = 6 + i
    put(ap, f"A{rr}", i); put(ap, f"B{rr}", a, wrap=True); put(ap, f"C{rr}", p, wrap=True)
    put(ap, f"D{rr}", isa, wrap=True); put(ap, f"E{rr}", ref)
    put(ap, f"F{rr}", None, fill=AMBER, border=BOX)
    put(ap, f"G{rr}", "Open"); put(ap, f"H{rr}", None, border=BOX)
    ap.row_dimensions[rr].height = 42
ap.freeze_panes = "A7"

# ===================================================================== Points forward
pf = wb.create_sheet("PF Points Forward")
header(pf, "PF", "Points forward, proposed adjustments, equity impact and evidence outstanding",
       [6, 46, 30, 30, 16, 16, 16, 34])
thead(pf, 6, ["#", "Description", "Debit", "Credit", "Amount (Rs)", "P&L FY2026 (Rs)", "Funds effect (Rs)", "Status / basis"])
ajes = [
    ("a", "Reclassify non-current to current per Schedule B (IAS 1.69)", "Financing – current portion", "Financing – non-current",
     f"=-{SP_RECLASS}", "=0", "=0", "Factual misstatement"),
    ("b", "FY2025 gain on extinguishment – Schedule F novation (IFRS 9 3.3.3; IAS 8 restatement)", "Financing (discount on new loan)", "Opening funds (FY2025 gain)",
     f"={GAIN}", "=0", f"={GAIN}", "Judgemental – ISA 540; subject to PMIC terms & market rate"),
    ("c", "Jul–Sep 2025 markup not accrued (old-facility rate)", "Finance cost (Note 23)", "Markup payable (Note 15) / cash",
     "='A3 Markup'!G12", "=-E9", "=-E9", "Factual – confirm KIBOR with PMIC"),
    ("d", "EIR unwind on new loan in excess of contractual markup (consequential on b)", "Finance cost (Note 23)", "Financing (discount unwind)",
     f"={AJE_D}", "=-E10", "=-E10", "Projected – only if (b) is booked"),
]
for i, row in enumerate(ajes):
    rr = 7 + i
    for j, v in enumerate(row):
        col = CL(j + 1)
        put(pf, f"{col}{rr}", v, fmt=NUM if j in (4, 5, 6) else None, wrap=j in (1, 2, 3, 7),
            color=GREEN if (isinstance(v, str) and "!" in v) else BLACK)
    pf.row_dimensions[rr].height = 40
put(pf, "B11", "Total effect", bold=True)
for c in "FG":
    put(pf, f"{c}11", f"=SUM({c}7:{c}10)", fmt=NUM, bold=True, border=DBL)
put(pf, "B12", "Each AJE vs performance materiality", italic=True)
put(pf, "H12", f'=IF(MIN(ABS(E7),ABS(E8),ABS(E9),ABS(E10))>{I["PM"]},"All exceed PM – material individually","Check individual items vs PM")', bold=True, color="C00000")

section(pf, "A14", "Equity / NBFC minimum-funds bridge")
for c in "BCDEFGH": pf[f"{c}14"].fill = GREY
thead(pf, 15, ["", "Scenario", "", "", "30-Jun-25 (Rs)", "", "30-Jun-26 (Rs)", "vs Rs 50m floor (30-Jun-26)"])
put(pf, "B16", "Funds as reported / draft"); put(pf, "E16", f"={I['FUNDS_PY']}", fmt=NUM, color=GREEN); put(pf, "G16", f"={I['FUNDS_CY']}", fmt=NUM, color=GREEN)
put(pf, "B17", "AJE (b) – FY2025 modification gain"); put(pf, "E17", "=G8", fmt=NUM); put(pf, "G17", "=G8", fmt=NUM)
put(pf, "B18", "AJE (c) – missing Jul–Sep markup (FY2026)"); put(pf, "E18", "=0", fmt=NUM); put(pf, "G18", "=G9", fmt=NUM)
put(pf, "B19", "AJE (d) – EIR unwind (FY2026)"); put(pf, "E19", "=0", fmt=NUM); put(pf, "G19", "=G10", fmt=NUM)
put(pf, "B20", "Funds after all AJEs", bold=True)
put(pf, "E20", "=SUM(E16:E19)", fmt=NUM, bold=True, border=DBL); put(pf, "G20", "=SUM(G16:G19)", fmt=NUM, bold=True, border=DBL)
for rr in (16, 20):
    put(pf, f"H{rr}", f'=IF(G{rr}>={I["FLOOR"]},"Above floor","BELOW floor")', bold=True, color="C00000")
put(pf, "B21", "Funds if only AJE (c) booked (no IFRS 9 modification)")
put(pf, "E21", "=E16", fmt=NUM); put(pf, "G21", "=G16+G18", fmt=NUM)
put(pf, "H21", f'=IF(G21>={I["FLOOR"]},"Above floor","BELOW floor")', bold=True, color="C00000")
put(pf, "B22", "Funds after all AJEs, excluding Rs 10m director sub-loan (IAS 32 risk)")
put(pf, "E22", f"=E20-{I['SUBLOAN']}", fmt=NUM); put(pf, "G22", f"=G20-{I['SUBLOAN']}", fmt=NUM)
put(pf, "H22", f'=IF(G22>={I["FLOOR"]},"Above floor","BELOW floor")', bold=True, color="C00000")
put(pf, "B23", "Check: 30-Jun-26 after AJEs = draft + gain – (EIR interest – booked)", italic=True)
put(pf, "G23", f"=G20-({I['FUNDS_CY']}+{GAIN}-({EIR_INT}-{I['FC_BOOK']}))", fmt=NUM)
put(pf, "H23", '=IF(ABS(G23)<1,"✓ Nil","✗")', bold=True)
# cell used by covenants sheet
put(pf, "B31", "Link cell – funds 30-Jun-26 after AJEs (used in A5)", italic=True, color="808080")
put(pf, "F31", "=G20", fmt=NUM, color="808080")

section(pf, "A25", "Matters for partner / TCWG attention")
for c in "BCDEFGH": pf[f"{c}25"].fill = GREY
mtrs = [
    "1. Classification error of Rs 29m (AJE a) – factual; discuss basis of client's Rs 35m.",
    "2. Unrecognised IFRS 9 modification gain (AJE b) is the decisive item for the NBFC Rs 50m equity floor at both year ends – ISA 540 sensitivity in A4 section E.",
    "3. Booking (b) without (d) overstates 30-Jun-26 equity: the discount unwinds at ~15% through FY2026 finance cost.",
    "4. OSS: 141% (client) vs 86% (income ÷ expenditure) – covenant first tested Dec-26 → IAS 1.76ZA disclosure; ISA 570 indicator together with equity position.",
]
for i, t in enumerate(mtrs):
    put(pf, f"B{26+i}", t)

section(pf, "A33", "Evidence outstanding (amber items)")
for c in "BCDEFGH": pf[f"{c}33"].fill = GREY
thead(pf, 34, ["#", "Evidence", "Purpose", "Source", "Requested", "Received", "", "WP ref"])
ev = [
    ("PMIC confirmation – balance, markup, terms, covenant breach/waiver", "Existence, completeness, classification, covenants", "PMIC (ISA 505)", "A / A5"),
    ("6M KIBOR applied for Jul–Sep 2025 quarter", "Accuracy of AJE (c)", "PMIC / SBP", "A3"),
    ("Original TFA I (2021) and TFA II (2022)", "Rights & obligations; history", "Client / PMIC", "Inputs / A1"),
    ("Amendment Agreement 27-Aug-24 and request letter 12-Jun-24", "Deferment period; old rate", "Client", "A1 / A3"),
    ("TFA 2025 incl. Schedules B and F; first Disbursement Date", "Modification date; grace period; novation", "Client / PMIC", "A4"),
    ("Market rate evidence (SSSF/peer unsubsidised borrowing June 2025) and KIBOR curve", "ISA 540 – fair value", "Bank quotes / SBP / Bloomberg", "A4"),
    ("FY2025 signed audited FS and predecessor's WPs", "Opening balance", "Predecessor auditor", "A1"),
    ("SECP charge certificate and modification filing", "Security / disclosure", "SECP / client", "Inputs"),
    ("Bank statements – PMIC markup payments FY2026 and after YE", "Cut-off; subsequent events", "Client bank", "A3"),
    ("Client covenant computation & OSS workings; CA/CL; LLP data", "Covenants", "Client / LLP WP", "A5"),
    ("Sub-loan agreement and receipt date", "Clause 6.2.13; IAS 32", "Client / bank", "A5"),
]
for i, (e, p, s, ref) in enumerate(ev, 1):
    rr = 34 + i
    put(pf, f"A{rr}", i); put(pf, f"B{rr}", e, wrap=True); put(pf, f"C{rr}", p, wrap=True); put(pf, f"D{rr}", s, wrap=True)
    put(pf, f"E{rr}", None, fill=AMBER, border=BOX); put(pf, f"F{rr}", None, fill=AMBER, border=BOX)
    put(pf, f"H{rr}", ref)
    pf.row_dimensions[rr].height = 30
# move link cell well below
pf["B31"].value = None; pf["F31"].value = None
put(pf, "B48", "Link cell – funds 30-Jun-26 after AJEs (used in A5)", italic=True, color="808080")
put(pf, "F48", "=G20", fmt=NUM, color="808080")
cv["E24"].value = "='PF Points Forward'!F48"

# ===================================================================== Index content
put(ix, "A6", "Tab", bold=True, color="FFFFFF", fill=NAVY)
put(ix, "B6", "Title", bold=True, color="FFFFFF", fill=NAVY)
put(ix, "C6", "Purpose", bold=True, color="FFFFFF", fill=NAVY)
put(ix, "D6", "Key result", bold=True, color="FFFFFF", fill=NAVY)
put(ix, "E6", "Check", bold=True, color="FFFFFF", fill=NAVY)
idx = [
    ("Inputs", "Inputs", "All facts, rates, dates and contract references; amber = confirm", None, None),
    ("A", "A Lead", "Lead schedule – FS vs GL, AJEs, audited balance", "='A Lead'!I9", "='A Lead'!C16"),
    ("A1", "A1 Movement", "Roll-forward; opening balance (ISA 510)", "='A1 Movement'!C13", "='A1 Movement'!D15"),
    ("A2", "A2 CNC Split", "Current vs non-current (IAS 1) – current portion per audit", f"='A2 CNC Split'!C{SP_CMP}", f"='A2 CNC Split'!D{TOT+1}"),
    ("A3", "A3 Markup", "Markup recompute – understatement (AJE c)", "='A3 Markup'!G12", "='A3 Markup'!H24"),
    ("A4", "A4 IFRS9 Modification", "Novation – DCF, 10% test, gain (AJE b), EIR unwind (AJE d)", f"={GAIN}", f"='A4 IFRS9 Modification'!L{QB+4}"),
    ("A5", "A5 Covenants", "Clause 6.2.12 / 6.2.13; NBFC minimum equity", "='A5 Covenants'!F13", "='A5 Covenants'!G13"),
    ("STD", "Standards & ISA", "IFRS / IAS / ISA / regulatory references", None, None),
    ("AP", "Audit Programme", "Procedures, assertions, sign-off", None, None),
    ("PF", "PF Points Forward", "AJEs, equity bridge, evidence outstanding", "='PF Points Forward'!G20", "='PF Points Forward'!H20"),
]
for i, (code, tab, purp, key, chk) in enumerate(idx):
    rr = 7 + i
    c = put(ix, f"A{rr}", code, bold=True)
    c.hyperlink = f"#'{tab}'!A1"
    c.font = Font(name=F, bold=True, color="0563C1", underline="single")
    put(ix, f"B{rr}", tab); put(ix, f"C{rr}", purp, wrap=True)
    if key: put(ix, f"D{rr}", key, fmt=NUM, color=GREEN)
    if chk: put(ix, f"E{rr}", chk, color=GREEN)
put(ix, "A18", "Legend", bold=True, fill=GREY)
for c in "BCDE": ix[f"{c}18"].fill = GREY
leg = [("123", "Blue text – hard-coded input (from client FS / agreements / user brief)", None, BLUE),
       ("123", "Green text – link to another sheet", None, GREEN),
       ("123", "Black text – formula on this sheet", None, BLACK),
       ("", "Amber fill – input awaiting PMIC confirmation, bank data or market data; replace and the workbook recalculates", AMBER, BLACK),
       ("✓ / ✗", "Automated tie-out check", None, BLACK)]
for i, (s, t, fill, col) in enumerate(leg):
    rr = 19 + i
    put(ix, f"A{rr}", s, color=col, fill=fill, border=BOX); put(ix, f"B{rr}", t)
put(ix, "A25", "Entity & framework", bold=True, fill=GREY)
for c in "BCDE": ix[f"{c}25"].fill = GREY
for i, t in enumerate([
    "SSSF – Section 42 (Companies Act 2017) not-for-profit microfinance NBFC, Sindh, Pakistan.",
    "Framework: IFRS as applicable in Pakistan; NBFC (Establishment & Regulation) Rules 2003 / NBFC Regulations 2008; Companies Act 2017.",
    "Materiality: overall Rs 600,000; performance Rs 300,000 (see Inputs).",
]):
    put(ix, f"A{26+i}", t)
put(ix, "A30", "Overall conclusion (preparer)", bold=True, fill=GREY)
for c in "BCDE": ix[f"{c}30"].fill = GREY
put(ix, "A31", "Subject to the amber items (PMIC confirmation, KIBOR, market rate, first Disbursement Date), the PMIC facility is NOT fairly stated in the draft FS: (a) current portion overstated by Rs 29m; (b) FY2025 modification gain unrecognised; (c) Jul–Sep 2025 markup not accrued; (d) EIR unwind not recognised. See PF.", wrap=True)
ix.merge_cells("A31:E33")
ix.row_dimensions[31].height = 30

# sheet order
order = ["Index", "Inputs", "A Lead", "A1 Movement", "A2 CNC Split", "A3 Markup", "A4 IFRS9 Modification",
         "A5 Covenants", "Standards & ISA", "Audit Programme", "PF Points Forward"]
wb._sheets = [wb[n] for n in order]
for ws in wb.worksheets:
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 0
tabc = {"Index": "1F3864", "Inputs": "FFC000", "Standards & ISA": "808080", "Audit Programme": "808080", "PF Points Forward": "C00000"}
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = tabc.get(ws.title, "2E75B6")
import os
os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)
print("saved")
