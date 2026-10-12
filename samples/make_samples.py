#!/usr/bin/env python3
"""Ardent Labs, Inc.: FP&A / corporate finance sample workbooks for the Finance Brain demo.

Deterministic (seed 26). Three workbooks. Cross-workbook dependencies are pasted values
with provenance comments (no live external links, so Excel never repairs the files).
Run:  python make_samples.py [out_dir]
"""
from __future__ import annotations
import datetime as dt, os, random, sys
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

OUT = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
R = random.Random(26)
FX = 1.08
TITLE = Font(bold=True, size=14); BOLD = Font(bold=True); GREY = Font(color="808080", italic=True)
HDR = PatternFill("solid", fgColor="DDEBF7"); INPUT = PatternFill("solid", fgColor="FFF2CC"); TOTAL = PatternFill("solid", fgColor="F2F2F2")
MONEY = '#,##0;(#,##0)'; PCT = '0.0%'; DATE = 'yyyy-mm-dd'; MON = 'mmm-yy'

def m(i): return dt.date(2026, i, 1)
def note(cell, text, author="FP&A"): cell.comment = Comment(text, author)
def header(ws, row, labels, col=1):
    for j, lab in enumerate(labels):
        c = ws.cell(row=row, column=col + j, value=lab); c.font = BOLD; c.fill = HDR
def widths(ws, w):
    for j, x in enumerate(w, 1): ws.column_dimensions[get_column_letter(j)].width = x
def add_name(wb, name, ref):
    dn = DefinedName(name, attr_text=ref)
    try: wb.defined_names.add(dn)
    except Exception: wb.defined_names[name] = dn
def put(ws, r, c, v, fmt=None, font=None, fill=None):
    cell = ws.cell(row=r, column=c, value=v)
    if fmt: cell.number_format = fmt
    if font: cell.font = font
    if fill: cell.fill = fill
    return cell

# ----------------------------------------------------------------------------- master data
DEPTS = [("SAL", "Sales", "CC-100", "Priya Natarajan", "Dan Okafor (CRO)"),
         ("MKT", "Marketing", "CC-200", "Luis Ortega", "Dan Okafor (CRO)"),
         ("CS", "Customer Success", "CC-300", "Hannah Weiss", "Dan Okafor (CRO)"),
         ("RND", "R&D", "CC-400", "Wei Zhang", "Sofia Almeida (CTO)"),
         ("PRD", "Product", "CC-450", "Aisha Bello", "Sofia Almeida (CTO)"),
         ("GA", "G&A", "CC-500", "Tom Hadley", "Grace Park (CFO)")]
COA = [("4000", "Subscription revenue", "Revenue", "Recurring"),
       ("4010", "Professional services revenue", "Revenue", "Non-recurring"),
       ("5000", "Hosting & infrastructure", "COGS", "Cost of revenue"),
       ("5020", "Services delivery", "COGS", "Cost of revenue"),
       ("6000", "Salaries & wages", "Opex", "People"), ("6010", "Benefits", "Opex", "People"),
       ("6020", "Commissions", "Opex", "People"), ("6030", "Bonus accrual", "Opex", "People"),
       ("6100", "Software & tools", "Opex", "Non-people"), ("6200", "Travel & entertainment", "Opex", "Non-people"),
       ("6300", "Marketing programs", "Opex", "Non-people"), ("6400", "Rent & facilities", "Opex", "Non-people"),
       ("6500", "Professional fees", "Opex", "Non-people"), ("6600", "Recruiting", "Opex", "Non-people"),
       ("6700", "Contractors", "Opex", "Non-people"), ("6900", "Other opex", "Opex", "Non-people")]
SAL_BASE = {"SAL": 330_000, "MKT": 150_000, "CS": 170_000, "RND": 520_000, "PRD": 140_000, "GA": 220_000}
COMMON = {"6000": 1.0, "6010": 0.22, "6030": 0.10, "6100": 0.06, "6200": 0.04, "6600": 0.03, "6900": 0.02}
EXTRA = {"SAL": {"6020": 0.35, "6200": 0.08}, "MKT": {"6300": 1.35, "6200": 0.08}, "GA": {"6400": 0.58, "6500": 0.32},
         "RND": {"6700": 0.17, "5000": 1.00}, "CS": {"5020": 0.40}, "PRD": {}}
EU_SHARE = 0.35

def budget_rows(periods):
    rows = []
    for p in periods:
        g = 1.02 ** (p.month - 1)
        for ent in ("US", "EU"):
            ef = 1.0 if ent == "US" else EU_SHARE
            sub = 2_600_000 * g * ef
            rows.append((p, ent, "REV", "4000", "Core", round(sub * 0.65)))
            rows.append((p, ent, "REV", "4000", "Analytics", round(sub * 0.35)))
            rows.append((p, ent, "REV", "4010", "Services", round(sub * 0.10)))
            for code, *_ in DEPTS:
                fr = dict(COMMON); fr.update(EXTRA[code])
                for acct in sorted(fr):
                    rows.append((p, ent, code, acct, "", round(SAL_BASE[code] * fr[acct] * (1 + 0.01 * (p.month - 1)) * ef)))
    return rows

def actual_rows(periods):
    rows = []
    for (p, ent, dept, acct, prod, usd) in budget_rows(periods):
        f = R.gauss(1.0, 0.04)
        if acct.startswith("4"): f = R.gauss(1.03, 0.02)
        if dept == "MKT" and acct == "6300" and 4 <= p.month <= 6: f = R.gauss(1.60, 0.05)
        if dept == "RND" and acct in ("6000", "6010", "6030") and 7 <= p.month <= 9: f = R.gauss(0.92, 0.01)
        amt = usd * f
        if ent == "EU": rows.append((p, ent, dept, acct, prod, round(amt / FX), "EUR"))
        else: rows.append((p, ent, dept, acct, prod, round(amt), "USD"))
    return rows

ACT = actual_rows([m(i) for i in range(1, 10)])
BUD = budget_rows([m(i) for i in range(1, 13)])
# stale hard-coded range: cuts into the September EU block (rows were appended after the ranges were written)
first_sep_eu = next(i for i, r in enumerate(ACT) if r[0] == m(9) and r[1] == "EU") + 2   # +2: header row and 1-based
ACT_END = first_sep_eu + 5
BUD_END = 1300
excluded = ACT[ACT_END - 1:]
cat = {a: c for a, _, c, _ in COA}
excl_by_cat = {}
for r in excluded:
    excl_by_cat[cat[r[3]]] = excl_by_cat.get(cat[r[3]], 0) + r[5] * FX

FIRST = ["Ava", "Noah", "Mia", "Liam", "Zoe", "Ethan", "Isla", "Lucas", "Nora", "Mateo", "Lena", "Kai", "Ines", "Omar", "Sara", "Felix", "Yara", "Jonas", "Leila", "Arjun", "Maya", "Diego", "Elif", "Ravi", "Chloe", "Tariq", "Hana", "Ivan", "Rosa", "Theo"]
LAST = ["Okafor", "Lindqvist", "Moreau", "Tanaka", "Haddad", "Novak", "Mensah", "Petrova", "Silva", "Brennan", "Kowalski", "Duarte", "Nakamura", "Achebe", "Fischer", "Rahman", "Castillo", "Ibrahim", "Jensen", "Varga"]
BANDS = {"L1": (90_000, 110_000, 130_000, 0.05), "L2": (110_000, 135_000, 160_000, 0.08), "L3": (140_000, 170_000, 200_000, 0.10),
         "L4": (175_000, 210_000, 250_000, 0.12), "L5": (220_000, 265_000, 320_000, 0.15), "L6": (280_000, 340_000, 420_000, 0.20)}
TITLES = {"SAL": ["Account Executive", "Sales Development Rep", "Sales Engineer", "Regional Sales Manager"],
          "MKT": ["Marketing Manager", "Content Lead", "Demand Gen Manager", "Product Marketing Manager"],
          "CS": ["Customer Success Manager", "Support Engineer", "Onboarding Specialist", "CS Team Lead"],
          "RND": ["Software Engineer", "Senior Software Engineer", "Staff Engineer", "Engineering Manager", "Data Engineer", "SRE"],
          "PRD": ["Product Manager", "Senior Product Manager", "Product Designer"],
          "GA": ["Accountant", "FP&A Analyst", "Recruiter", "People Ops Partner", "Legal Counsel", "IT Administrator"]}
HC_MIX = [("RND", 45), ("SAL", 25), ("CS", 18), ("MKT", 12), ("PRD", 8), ("GA", 12)]
roster = []; eid = 1001; used = set()
for dept, n in HC_MIX:
    for _ in range(n):
        while True:
            name = f"{R.choice(FIRST)} {R.choice(LAST)}"
            if name not in used: used.add(name); break
        level = R.choices(["L1", "L2", "L3", "L4", "L5", "L6"], weights=[10, 25, 30, 20, 10, 5])[0]
        lo, mid, hi, _ = BANDS[level]
        salary = int(round(R.uniform(lo, hi) / 1000) * 1000)
        start = dt.date(2019, 1, 1) + dt.timedelta(days=R.randint(0, 365 * 7 + 200))
        loc = R.choices(["Austin, TX", "Denver, CO", "Remote US", "Berlin, DE", "Lisbon, PT"], weights=[35, 15, 20, 20, 10])[0]
        status = "Active"; end = None
        if R.random() < 0.08: status = "Terminated"; end = dt.date(2026, R.randint(1, 9), R.randint(1, 28))
        roster.append((f"E{eid}", name, dept, R.choice(TITLES[dept]), level, loc, start, salary, 1.0 if R.random() > 0.05 else 0.8, status, end)); eid += 1
HC_SEP = {code: sum(1 for r_ in roster if r_[2] == code and r_[6] <= dt.date(2026, 9, 30) and (r_[9] == "Active" or (r_[10] and r_[10] > dt.date(2026, 9, 30)))) for code, *_ in DEPTS}

# ----------------------------------------------------------------------------- workbook 1: Budget vs Actuals
wb = Workbook()
ws = wb.active; ws.title = "Cover"
put(ws, 1, 1, "Ardent Labs, Inc.", font=TITLE); ws.merge_cells("A1:F1")
put(ws, 2, 1, "FY2026 Budget vs Actuals, September 2026 close", font=BOLD)
for i, (k, v) in enumerate([("Prepared by", "FP&A (Maya Chen)"), ("Version", "v9, 2026-10-06"), ("Source", "GL export 2026-10-03 (Jan to Sep)"),
                            ("Currency", "USD; EU entity converted at Assumptions!FX_EUR_USD"), ("Status", "Draft for CFO review"),
                            ("Reviewer", "Grace Park (CFO)")], start=4):
    put(ws, i, 1, k, font=BOLD); put(ws, i, 2, v)
put(ws, 11, 1, "Change log", font=BOLD)
for i, line in enumerate(["v9 2026-10-06: September actuals appended to Actuals tab (rows added at the bottom).",
                          "v8 2026-09-08: Q4 budget reforecast loaded to Budget tab as version 'FY26 RF2'.",
                          "v7 2026-08-05: Dept Variance tab added at CFO request; commentary moved to its own tab.",
                          "v6 2026-07-14: Benefits load changed from 20% to 22% (Treasury memo). Headcount model copied the value."], start=12):
    put(ws, i, 1, line)
widths(ws, [16, 70])

ws = wb.create_sheet("Assumptions")
header(ws, 1, ["Assumption", "Value", "Unit", "Owner", "Note"])
A = [("FX EUR to USD (budget rate)", FX, "rate", "Treasury", "Manual input; updated quarterly. Actual monthly rates live in the Treasury FX file."),
     ("Benefits load", 0.22, "% of salaries", "People Ops", "Changed from 20% on 2026-07-14 (Treasury memo). Headcount Plan copies this value."),
     ("Bonus accrual", 0.10, "% of salaries", "FP&A", "Company plan at target."),
     ("Commission rate", 0.08, "% of new ARR", "Sales Ops", "Blended; plan document v3."),
     ("Statutory tax rate", 0.21, "rate", "Tax", "US federal only; state blended rate in the provision model."),
     ("Planned monthly attrition", 0.012, "% of HC", "People Ops", "Used by Headcount Plan."),
     ("Hosting cost per 1k API calls", 0.42, "USD", "R&D", "Vendor contract FY26.")]
for i, row in enumerate(A, start=2):
    for j, v in enumerate(row, 1): put(ws, i, j, v)
    ws.cell(row=i, column=2).fill = INPUT
    ws.cell(row=i, column=2).number_format = PCT if row[2].startswith("%") else "0.0000"
note(ws["B2"], "Manual input. Updated quarterly by Treasury. Do not overwrite with actual rates.", "Treasury")
note(ws["B3"], "Changed from 20% to 22% on 2026-07-14. The Headcount Plan still has 22% typed in; check it when this changes.", "Maya Chen")
add_name(wb, "FX_EUR_USD", "Assumptions!$B$2"); add_name(wb, "BenefitsLoad", "Assumptions!$B$3")
add_name(wb, "BonusPct", "Assumptions!$B$4"); add_name(wb, "CommissionRate", "Assumptions!$B$5")
widths(ws, [32, 10, 14, 12, 80])

ws = wb.create_sheet("COA")
header(ws, 1, ["Account", "Account name", "Category", "Subcategory", "P&L order"])
for i, (a, n, c, s) in enumerate(COA, start=2):
    put(ws, i, 1, a); put(ws, i, 2, n); put(ws, i, 3, c); put(ws, i, 4, s); put(ws, i, 5, i - 1)
widths(ws, [10, 32, 10, 16, 10])

ws = wb.create_sheet("Departments")
header(ws, 1, ["Dept code", "Department", "Cost center", "Owner", "VP"])
for i, d in enumerate(DEPTS, start=2):
    for j, v in enumerate(d, 1): put(ws, i, j, v)
put(ws, 8, 1, "REV"); put(ws, 8, 2, "Revenue (no department)"); put(ws, 8, 3, "n/a"); put(ws, 8, 4, "Revenue accounting"); put(ws, 8, 5, "Grace Park (CFO)")
widths(ws, [10, 26, 12, 18, 22])

ws = wb.create_sheet("Actuals")
header(ws, 1, ["Period", "Entity", "Department", "Account", "Product", "Amount", "Currency", "Source"])
for i, (p, ent, dept, acct, prod, amt, cur) in enumerate(ACT, start=2):
    put(ws, i, 1, p, DATE); put(ws, i, 2, ent); put(ws, i, 3, dept); put(ws, i, 4, acct); put(ws, i, 5, prod)
    put(ws, i, 6, amt, MONEY); put(ws, i, 7, cur); put(ws, i, 8, "GL export 2026-10-03")
note(ws["A1"], "Long format GL extract. EU rows are in EUR; the BvA Summary converts them with FX_EUR_USD. September rows were appended on 2026-10-06.", "Maya Chen")
ws.freeze_panes = "A2"; widths(ws, [12, 8, 12, 10, 11, 14, 10, 22])

ws = wb.create_sheet("Budget")
header(ws, 1, ["Period", "Entity", "Department", "Account", "Product", "Amount (USD)", "Version"])
for i, (p, ent, dept, acct, prod, amt) in enumerate(BUD, start=2):
    put(ws, i, 1, p, DATE); put(ws, i, 2, ent); put(ws, i, 3, dept); put(ws, i, 4, acct); put(ws, i, 5, prod)
    put(ws, i, 6, amt, MONEY); put(ws, i, 7, "FY26 Board" if p.month <= 9 else "FY26 RF2")
ws.freeze_panes = "A2"; widths(ws, [12, 8, 12, 10, 11, 14, 12])

ws = wb.create_sheet("BvA Summary")
put(ws, 1, 1, "Ardent Labs, Inc.  FY2026 Budget vs Actuals (USD)", font=TITLE); ws.merge_cells("A1:M1")
put(ws, 2, 1, "September 2026 close. EU entity converted at FX_EUR_USD. Prepared by FP&A.", font=GREY)
put(ws, 3, 1, "period key (hidden)", font=GREY)
for col, d in zip("DEFGHI", [m(7), m(7), m(8), m(8), m(9), m(9)]): put(ws, 3, ord(col) - 64, d, DATE, font=GREY)
put(ws, 3, 10, m(9), DATE, font=GREY); put(ws, 3, 11, m(9), DATE, font=GREY)
ws.row_dimensions[3].hidden = True
header(ws, 4, ["Account", "Account name", "Category", "Jul-26 Actual", "Jul-26 Budget", "Aug-26 Actual", "Aug-26 Budget",
               "Sep-26 Actual", "Sep-26 Budget", "YTD Actual", "YTD Budget", "YTD Variance", "YTD Var %"])
AR = f"$2:${ACT_END}"; BR = f"$2:${BUD_END}"
def act_f(acct_cell, period_cell=None):
    crit = f",Actuals!$A{AR.replace('$2:$', '$2:$A$') if False else ''}" if False else ""
    base = f"Actuals!$F$2:$F${ACT_END},Actuals!$D$2:$D${ACT_END},{acct_cell}"
    per = f",Actuals!$A$2:$A${ACT_END},{period_cell}" if period_cell else ""
    us = f"SUMIFS({base}{per},Actuals!$B$2:$B${ACT_END},\"US\")"
    eu = f"SUMIFS({base}{per},Actuals!$B$2:$B${ACT_END},\"EU\")*FX_EUR_USD"
    return f"={us}+{eu}"
def bud_f(acct_cell, period_cell, ytd=False):
    per = f"\"<=\"&{period_cell}" if ytd else period_cell
    return f"=SUMIFS(Budget!$F$2:$F${BUD_END},Budget!$A$2:$A${BUD_END},{per},Budget!$D$2:$D${BUD_END},{acct_cell})"
groups = [("Revenue", ["4000", "4010"]), ("COGS", ["5000", "5020"]),
          ("Opex", ["6000", "6010", "6020", "6030", "6100", "6200", "6300", "6400", "6500", "6600", "6700", "6900"])]
r = 5; totals = {}
for gname, accts in groups:
    first = r
    for a in accts:
        put(ws, r, 1, a); put(ws, r, 2, f'=VLOOKUP($A{r},COA!$A$2:$E$30,2,FALSE)'); put(ws, r, 3, f'=VLOOKUP($A{r},COA!$A$2:$E$30,3,FALSE)')
        for col, pc in (("D", "D$3"), ("F", "F$3"), ("H", "H$3")): put(ws, r, ord(col) - 64, act_f(f"$A{r}", pc), MONEY)
        for col, pc in (("E", "E$3"), ("G", "G$3"), ("I", "I$3")): put(ws, r, ord(col) - 64, bud_f(f"$A{r}", pc), MONEY)
        put(ws, r, 10, act_f(f"$A{r}"), MONEY); put(ws, r, 11, bud_f(f"$A{r}", "K$3", ytd=True), MONEY)
        put(ws, r, 12, f"=J{r}-K{r}", MONEY); put(ws, r, 13, f"=IFERROR(L{r}/K{r},0)", PCT)
        r += 1
    put(ws, r, 2, f"Total {gname}", font=BOLD)
    for c in range(4, 13): put(ws, r, c, f"=SUM({get_column_letter(c)}{first}:{get_column_letter(c)}{r-1})", MONEY, font=BOLD, fill=TOTAL)
    put(ws, r, 13, f"=IFERROR(L{r}/K{r},0)", PCT, font=BOLD, fill=TOTAL)
    totals[gname] = r; r += 1
    if gname == "COGS":
        put(ws, r, 2, "Gross profit", font=BOLD)
        for c in range(4, 13): L = get_column_letter(c); put(ws, r, c, f"={L}{totals['Revenue']}-{L}{totals['COGS']}", MONEY, font=BOLD, fill=TOTAL)
        put(ws, r, 13, f"=IFERROR(L{r}/K{r},0)", PCT, font=BOLD, fill=TOTAL); totals["GP"] = r; r += 1
        put(ws, r, 2, "Gross margin %", font=BOLD)
        for c in range(4, 12): L = get_column_letter(c); put(ws, r, c, f"=IFERROR({L}{totals['GP']}/{L}{totals['Revenue']},0)", PCT, font=BOLD)
        put(ws, r, 12, f"=J{r}-K{r}", PCT, font=BOLD); totals["GM"] = r
        note(ws.cell(row=r, column=2), "FP&A definition: (Revenue - COGS) / Revenue. COGS = hosting + services delivery. Customer Success is Opex here. The Board KPI pack uses the investor definition that includes CS cost.", "Maya Chen")
        r += 1
put(ws, r, 2, "EBITDA", font=BOLD)
for c in range(4, 13): L = get_column_letter(c); put(ws, r, c, f"={L}{totals['GP']}-{L}{totals['Opex']}", MONEY, font=BOLD, fill=TOTAL)
put(ws, r, 13, f"=IFERROR(L{r}/K{r},0)", PCT, font=BOLD, fill=TOTAL); totals["EBITDA"] = r; r += 1
put(ws, r, 2, "EBITDA margin %", font=BOLD)
for c in range(4, 12): L = get_column_letter(c); put(ws, r, c, f"=IFERROR({L}{totals['EBITDA']}/{L}{totals['Revenue']},0)", PCT, font=BOLD)
put(ws, r, 12, f"=J{r}-K{r}", PCT, font=BOLD); totals["EBITDAm"] = r; r += 3
hc_title = r
put(ws, r, 1, "Headcount snapshot (pasted from Headcount Plan, HC Summary, 2026-10-02)", font=BOLD); r += 1
header(ws, r, ["Department", "Sep-26 HC", "Budget HC", "Variance"]); r += 1
HC_SNAP = [(name, HC_SEP[code] + (2 if code == "RND" else 0), HC_SEP[code] + {"SAL": 1, "MKT": 1, "CS": 0, "RND": 8, "PRD": 1, "GA": 0}[code]) for code, name, *_ in DEPTS]
hc_first = r
for d, a, b in HC_SNAP:
    put(ws, r, 1, d); put(ws, r, 2, a); put(ws, r, 3, b); put(ws, r, 4, f"=B{r}-C{r}"); r += 1
put(ws, r, 1, "Total", font=BOLD)
for c in (2, 3, 4): put(ws, r, c, f"=SUM({get_column_letter(c)}{hc_first}:{get_column_letter(c)}{r-1})", font=BOLD, fill=TOTAL)
note(ws.cell(row=hc_first, column=2), "Pasted from Ardent_FY26_Headcount_Plan.xlsx, Roster, active at Sep 30, on 2026-10-02 before two R&D September terminations were processed. Not linked.", "Maya Chen")
r += 3
put(ws, r, 1, "Notes", font=BOLD); r += 1
for line in ["1. Variance = Actual - Budget. For expense lines a negative variance is favourable.",
             "2. EU entity amounts are converted at the budget rate (Assumptions). Treasury reports actual-rate figures separately.",
             "3. Q4 budget columns use reforecast RF2 (Budget tab, Version column).",
             "4. SUMIFS ranges on this tab were last extended in v5 (June). Check the range end when actuals rows are appended."]:
    put(ws, r, 1, line); r += 1
ws.freeze_panes = "D5"; widths(ws, [10, 30, 10] + [13] * 10)

ws = wb.create_sheet("Dept Variance")
header(ws, 1, ["Dept code", "Department", "Owner", "YTD Actual", "YTD Budget", "Variance", "Var %", "Flag", "Commentary"])
COMMENT = {"SAL": "Commissions above plan on 6% ARR over-attainment; T&E over by $62K for the EMEA field kickoff.",
           "MKT": "Q2 programs over budget by about $410K: Series B launch events pulled forward from Q3, approved by CMO 2026-05-12. Expect Q3 and Q4 under budget.",
           "CS": "On plan. Services delivery cost (5020) tracked in R&D hosting review, not here.",
           "RND": "Under budget on salaries, benefits and bonus from July: six requisitions slipped to Q4 (see Headcount Plan, Hiring Plan tab, status 'Slipped to Q4'). Hosting (5000) tracks usage, on plan.",
           "PRD": "On plan.",
           "GA": "Professional fees over by $48K: audit readiness project (Deloitte SOW 2026-03)."}
for i, (code, *_rest) in enumerate(DEPTS, start=2):
    put(ws, i, 1, code); put(ws, i, 2, f"=VLOOKUP($A{i},Departments!$A$2:$E$8,2,FALSE)"); put(ws, i, 3, f"=VLOOKUP($A{i},Departments!$A$2:$E$8,4,FALSE)")
    put(ws, i, 4, f"=SUMIFS(Actuals!$F$2:$F${ACT_END},Actuals!$C$2:$C${ACT_END},$A{i},Actuals!$B$2:$B${ACT_END},\"US\")+SUMIFS(Actuals!$F$2:$F${ACT_END},Actuals!$C$2:$C${ACT_END},$A{i},Actuals!$B$2:$B${ACT_END},\"EU\")*FX_EUR_USD", MONEY)
    put(ws, i, 5, f"=SUMIFS(Budget!$F$2:$F${BUD_END},Budget!$C$2:$C${BUD_END},$A{i},Budget!$A$2:$A${BUD_END},\"<=\"&DATE(2026,9,1))", MONEY)
    put(ws, i, 6, f"=D{i}-E{i}", MONEY); put(ws, i, 7, f"=IFERROR(F{i}/E{i},0)", PCT); put(ws, i, 8, f'=IF(ABS(G{i})>0.05,"Review","")'); put(ws, i, 9, COMMENT[code])
put(ws, 8, 1, "Total", font=BOLD)
for c in (4, 5, 6): put(ws, 8, c, f"=SUM({get_column_letter(c)}2:{get_column_letter(c)}7)", MONEY, font=BOLD, fill=TOTAL)
put(ws, 8, 7, "=IFERROR(F8/E8,0)", PCT, font=BOLD, fill=TOTAL)
widths(ws, [10, 20, 18, 14, 14, 14, 9, 9, 110])

ws = wb.create_sheet("Commentary")
header(ws, 1, ["Period", "Department", "Author", "Commentary"])
C = [(m(9), "FP&A", "Maya Chen", "September close: revenue 3% over plan on Analytics upsell; EBITDA margin 2 points above plan mostly from R&D hiring slip, which is timing, not savings."),
     (m(9), "RND", "Wei Zhang", "Six reqs (REQ-0412 to REQ-0417) slipped to Q4; offers out for two. Salary underspend will reverse in Q1 FY27 when the cohort starts."),
     (m(9), "MKT", "Luis Ortega", "Q3 programs 18% under budget to offset the Q2 launch pull-forward. Full-year marketing programs expected on plan."),
     (m(8), "FP&A", "Maya Chen", "Benefits load changed to 22% in July. Headcount Plan was updated by copying the value; it is not linked. Q4 rows there use 25% per People Ops guidance, which has not been approved by Finance."),
     (m(8), "GA", "Tom Hadley", "Audit readiness SOW runs through November. $48K over plan YTD, expected $90K over by year end."),
     (m(7), "SAL", "Priya Natarajan", "EMEA field kickoff in Lisbon drove T&E over plan. Commissions follow bookings; over-attainment is good news."),
     (m(7), "FP&A", "Maya Chen", "Board KPI pack reports gross margin on the investor definition (includes Customer Success cost). BvA gross margin excludes it. Reconcile both before the October board meeting."),
     (m(6), "FP&A", "Maya Chen", "Q4 reforecast RF2 loaded to the Budget tab. Board version kept for Jan to Sep. Versions are distinguished only by the Version column."),
     (m(6), "CS", "Hannah Weiss", "Services delivery cost moved under COGS 5020 in June (was Opex 6700 for the CS team). Prior months restated in the GL, not in older BvA versions."),
     (m(5), "MKT", "Luis Ortega", "Series B launch events pulled into May and June. CMO approval 2026-05-12 attached to the spend request."),
     (m(4), "FP&A", "Maya Chen", "Revenue in this file is GAAP recognised revenue from the GL. The Board KPI pack shows ARR divided by 12 as run-rate revenue. They will never tie; do not try to reconcile them to the dollar."),
     (m(3), "GA", "Tom Hadley", "FX: Treasury asked FP&A to keep the budget rate (1.08) for variance reporting. Actual-rate P&L is in the Treasury monthly file."),
     (m(2), "RND", "Wei Zhang", "Hosting cost per 1k API calls renegotiated to $0.42 (from $0.47) effective March."),
     (m(1), "FP&A", "Maya Chen", "FY26 budget approved by the board on 2025-12-16. This file is the single source for monthly budget vs actuals.")]
for i, row in enumerate(C, start=2):
    put(ws, i, 1, row[0], MON); put(ws, i, 2, row[1]); put(ws, i, 3, row[2]); put(ws, i, 4, row[3])
widths(ws, [10, 12, 18, 140])
wb.save(os.path.join(OUT, "Ardent_FY26_Budget_vs_Actuals.xlsx"))

# ----------------------------------------------------------------------------- workbook 2: Headcount & Compensation Plan
begin_hc = sum(1 for r_ in roster if r_[6] < m(1) and (r_[9] == "Active" or r_[10] >= m(1)))

wb = Workbook(); ws = wb.active; ws.title = "Roster"
header(ws, 1, ["Employee ID", "Name", "Department", "Title", "Level", "Location", "Start date", "Annual base (USD)", "FTE", "Status", "End date"])
for i, row in enumerate(roster, start=2):
    for j, v in enumerate(row, 1):
        c = put(ws, i, j, v)
        if j in (7, 11) and v: c.number_format = DATE
        if j == 8: c.number_format = MONEY
note(ws["H1"], "EU employees are shown in USD at the budget FX rate for planning. Payroll pays EUR.", "People Ops")
ws.freeze_panes = "A2"; widths(ws, [12, 22, 12, 28, 7, 14, 12, 16, 6, 11, 12])

ws = wb.create_sheet("Comp Bands")
header(ws, 1, ["Level", "Min", "Mid", "Max", "Bonus target %"])
for i, (lvl, (lo, mid, hi, b)) in enumerate(BANDS.items(), start=2):
    put(ws, i, 1, lvl); put(ws, i, 2, lo, MONEY); put(ws, i, 3, mid, MONEY); put(ws, i, 4, hi, MONEY); put(ws, i, 5, b, PCT)
put(ws, 9, 1, "Bands effective 2026-01-01 (Radford Q4 FY25 refresh).", font=GREY)
widths(ws, [8, 12, 12, 12, 14])

ws = wb.create_sheet("Hiring Plan")
header(ws, 1, ["Req ID", "Department", "Title", "Level", "Planned start", "Status", "Offer salary (USD)", "Hiring manager"])
reqs = []
plan = [("SAL", 5), ("MKT", 2), ("CS", 3), ("RND", 14), ("PRD", 2), ("GA", 4)]
n = 401
for dept, k in plan:
    for _ in range(k):
        mon = R.randint(2, 11); level = R.choice(["L2", "L3", "L4"]); lo, mid, hi, _ = BANDS[level]
        status = "Hired" if mon <= 8 and R.random() < 0.7 else ("Open" if mon >= 9 else R.choice(["Open", "Cancelled", "Offer out"]))
        reqs.append([f"REQ-0{n}", dept, R.choice(TITLES[dept]), level, m(mon), status, int(round(R.uniform(lo, hi) / 1000) * 1000), dict((d[0], d[3]) for d in DEPTS)[dept]]); n += 1
slipped = [r_ for r_ in reqs if r_[1] == "RND"][:6]
for k, r_ in enumerate(slipped):
    r_[0] = f"REQ-04{12 + k}"; r_[4] = m(7 + k % 3); r_[5] = "Slipped to Q4"
for i, row in enumerate(reqs, start=2):
    for j, v in enumerate(row, 1):
        c = put(ws, i, j, v)
        if j == 5: c.number_format = DATE
        if j == 7: c.number_format = MONEY
note(ws["F1"], "Status values: Open, Offer out, Hired, Slipped to Q4, Cancelled. HC Summary counts only 'Hired' by planned start month.", "People Ops")
widths(ws, [10, 12, 28, 7, 13, 14, 16, 18])

ws = wb.create_sheet("HC Summary")
put(ws, 1, 1, "Ardent Labs, Inc.  FY2026 Headcount and loaded compensation", font=TITLE); ws.merge_cells("A1:I1")
put(ws, 2, 1, "Benefits load (Jan to Sep)", font=BOLD); c = put(ws, 2, 2, 0.22, PCT, fill=INPUT)
note(c, "Copied from Ardent_FY26_Budget_vs_Actuals.xlsx, Assumptions, Benefits load (B3) on 2026-07-14. Not linked; re-copy when Finance changes it.", "People Ops")
put(ws, 3, 1, "Benefits load (Q4 forecast)", font=BOLD); c = put(ws, 3, 2, 0.25, PCT, fill=INPUT)
note(c, "People Ops guidance for the FY27 plan year, applied to Q4 early. Not approved by Finance (see BvA Commentary, Aug-26).", "People Ops")
header(ws, 5, ["Month", "Beginning HC", "Hires", "Attrition", "Ending HC", "Monthly base comp", "Loaded comp", "Budget loaded comp", "Variance"])
attr = [1, 2, 1, 0, 2, 1, 3, 1, 2, 1, 1, 1]
for k in range(12):
    rr = 6 + k; put(ws, rr, 1, m(k + 1), MON)
    if k == 0:
        c = put(ws, rr, 2, begin_hc, fill=INPUT); note(c, "Opening headcount per Dec-25 payroll register.", "People Ops")
    else: put(ws, rr, 2, f"=E{rr-1}")
    put(ws, rr, 3, f'=COUNTIFS(\'Hiring Plan\'!$E:$E,">="&A{rr},\'Hiring Plan\'!$E:$E,"<="&EOMONTH(A{rr},0),\'Hiring Plan\'!$F:$F,"Hired")')
    put(ws, rr, 4, attr[k], fill=INPUT); put(ws, rr, 5, f"=B{rr}+C{rr}-D{rr}")
    put(ws, rr, 6, f'=SUMIFS(Roster!$H:$H,Roster!$G:$G,"<="&EOMONTH(A{rr},0),Roster!$J:$J,"Active")/12', MONEY)
    put(ws, rr, 7, f"=F{rr}*(1+$B$2)" if k < 9 else f"=F{rr}*(1+$B$3)", MONEY)
    put(ws, rr, 8, round(2_200_000 * (1 + 0.012 * k) * R.gauss(1.0, 0.01)), MONEY, fill=INPUT); put(ws, rr, 9, f"=G{rr}-H{rr}", MONEY)
put(ws, 18, 1, "Total", font=BOLD)
for cc in (3, 4, 6, 7, 8, 9): put(ws, 18, cc, f"=SUM({get_column_letter(cc)}6:{get_column_letter(cc)}17)", MONEY if cc > 4 else None, font=BOLD, fill=TOTAL)
note(ws["H5"], "Budget loaded comp pasted from the FY26 budget model (People tab) on 2026-01-12.", "People Ops")
put(ws, 20, 1, "Notes", font=BOLD)
for i, line in enumerate(["Beginning HC chains from the prior month's Ending HC.", "Hires counts Hiring Plan rows with status Hired by planned start month; actual start dates are in Workday, not here.",
                          "Monthly base comp sums active roster base salaries with a start date on or before month end, divided by 12.",
                          "Loaded comp applies the benefits load in B2 for Jan to Sep and B3 for Q4."], start=21): put(ws, i, 1, line)
widths(ws, [12, 13, 8, 9, 10, 17, 15, 18, 13])
wb.save(os.path.join(OUT, "Ardent_FY26_Headcount_Plan.xlsx"))

# ----------------------------------------------------------------------------- workbook 3: Board KPI pack
COMPANIES = ["Harbor", "Northwind", "Pinecrest", "Bluefin", "Quarry", "Lantern", "Saffron", "Granite", "Juniper", "Kestrel", "Marlin", "Orchard", "Tidal", "Vantage", "Willow", "Zenith", "Beacon", "Cinder", "Delta", "Ember"]
SUFFIX = ["Health", "Logistics", "Financial", "Retail", "Media", "Energy", "Labs", "Systems", "Foods", "Mobility"]
custs = []; cid = 1
for seg, n, lo, hi in (("SMB", 90, 8_000, 32_000), ("Mid-Market", 80, 45_000, 160_000), ("Enterprise", 30, 220_000, 950_000)):
    for _ in range(n):
        start = dt.date(2022, 1, 1) + dt.timedelta(days=R.randint(0, 365 * 4 + 240))
        churn = None
        if start < m(1) and R.random() < 0.12: churn = dt.date(2026, R.randint(1, 9), R.randint(1, 28))
        custs.append((f"C{cid:04d}", f"{R.choice(COMPANIES)} {R.choice(SUFFIX)} {cid}", seg, R.choices(["NA", "EMEA"], weights=[70, 30])[0],
                      R.choices(["Core", "Core + Analytics"], weights=[60, 40])[0], start, churn, int(round(R.uniform(lo, hi) / 100) * 100))); cid += 1
begin_arr = sum(c_[7] for c_ in custs if c_[5] < m(1))

wb = Workbook(); ws = wb.active; ws.title = "Customers"
header(ws, 1, ["Customer ID", "Customer", "Segment", "Region", "Product", "Start date", "Churn date", "ARR (USD)"])
for i, row in enumerate(custs, start=2):
    for j, v in enumerate(row, 1):
        c = put(ws, i, j, v)
        if j in (6, 7) and v: c.number_format = DATE
        if j == 8: c.number_format = MONEY
note(ws["H1"], "Contracted ARR at the latest renewal, from the billing system export 2026-09-30. Expansion since renewal is in the ARR Bridge, not here.", "RevOps")
ws.freeze_panes = "A2"; widths(ws, [11, 26, 12, 8, 16, 12, 12, 12])

ws = wb.create_sheet("ARR Bridge")
put(ws, 1, 1, "Ardent Labs, Inc.  ARR bridge FY2026 (USD)", font=TITLE); ws.merge_cells("A1:H1")
header(ws, 3, ["Month", "Beginning ARR", "New", "Expansion", "Contraction", "Churn", "Ending ARR", "Net new ARR"])
for k in range(9):
    rr = 4 + k; put(ws, rr, 1, m(k + 1), MON)
    if k == 0:
        c = put(ws, rr, 2, begin_arr, MONEY, fill=INPUT); note(c, "Dec-25 ending ARR per the FY25 Q4 board pack. Typed, not linked.", "RevOps")
    else: put(ws, rr, 2, f"=G{rr-1}", MONEY)
    put(ws, rr, 3, f'=SUMIFS(Customers!$H:$H,Customers!$F:$F,">="&A{rr},Customers!$F:$F,"<="&EOMONTH(A{rr},0))', MONEY)
    put(ws, rr, 4, round(R.uniform(180_000, 420_000), -3), MONEY, fill=INPUT); put(ws, rr, 5, -round(R.uniform(40_000, 120_000), -3), MONEY, fill=INPUT)
    put(ws, rr, 6, f'=-SUMIFS(Customers!$H:$H,Customers!$G:$G,">="&A{rr},Customers!$G:$G,"<="&EOMONTH(A{rr},0))', MONEY)
    put(ws, rr, 7, f"=B{rr}+C{rr}+D{rr}+E{rr}+F{rr}", MONEY); put(ws, rr, 8, f"=G{rr}-B{rr}", MONEY)
note(ws["D3"], "Expansion and contraction come from the billing system upsell report, pasted monthly by RevOps. Not derivable from the Customers tab.", "RevOps")
put(ws, 14, 1, "NRR (YTD, simplified)", font=BOLD); put(ws, 14, 2, "=(B4+SUM(D4:D12)+SUM(E4:E12)+SUM(F4:F12))/B4", PCT)
note(ws["B14"], "Simplified: (beginning ARR + expansion + contraction + churn) / beginning ARR, YTD not trailing twelve months. The Definitions tab says TTM; this cell does not do that yet.", "RevOps")
widths(ws, [22, 14, 12, 12, 12, 12, 14, 13])

ws = wb.create_sheet("KPI Summary")
put(ws, 1, 1, "Ardent Labs, Inc.  Q3 FY2026 board KPIs", font=TITLE); ws.merge_cells("A1:D1")
header(ws, 3, ["Metric", "Q3 FY26", "Definition", "Source"])
q3_rev = sum((r_[5] * (FX if r_[6] == "EUR" else 1)) for r_ in ACT if r_[0].month >= 7 and r_[3].startswith("4"))
q3_cogs = sum((r_[5] * (FX if r_[6] == "EUR" else 1)) for r_ in ACT if r_[0].month >= 7 and r_[3].startswith("5"))
q3_cs = sum((r_[5] * (FX if r_[6] == "EUR" else 1)) for r_ in ACT if r_[0].month >= 7 and r_[2] == "CS" and r_[3].startswith("6"))
q3_ebitda = q3_rev - q3_cogs - sum((r_[5] * (FX if r_[6] == "EUR" else 1)) for r_ in ACT if r_[0].month >= 7 and r_[3].startswith("6"))
K = [("Ending ARR", "='ARR Bridge'!G12", "Contracted annual recurring revenue at Sep 30", "ARR Bridge"),
     ("Run-rate revenue (monthly)", "=B4/12", "Ending ARR / 12", "Derived"),
     ("GAAP revenue, Q3", round(q3_rev), "Recognised revenue Jul to Sep, USD", "Pasted from BvA Summary (Jul to Sep actual) 2026-10-03"),
     ("COGS, Q3", round(q3_cogs), "Hosting + services delivery", "Pasted from BvA Summary 2026-10-03"),
     ("Customer Success cost, Q3", round(q3_cs), "CS department opex, all accounts", "Pasted from GL export 2026-10-03"),
     ("Gross margin % (investor)", "=(B6-B7-B8)/B6", "(Revenue - COGS - Customer Success cost) / Revenue", "Derived; differs from FP&A BvA definition"),
     ("Net new ARR, Q3", "=SUM('ARR Bridge'!H10:H12)", "Ending minus beginning ARR, Jul to Sep", "ARR Bridge"),
     ("NRR", "='ARR Bridge'!B14", "Net revenue retention (see Definitions)", "ARR Bridge"),
     ("EBITDA, Q3", round(q3_ebitda), "Revenue - COGS - Opex", "Pasted from BvA Summary 2026-10-03"),
     ("EBITDA margin, Q3", "=B12/B6", "EBITDA / GAAP revenue", "Derived"),
     ("ARR growth YoY", "=B4/B15-1", "Ending ARR vs Sep-25 ARR", "Derived"),
     ("Sep-25 ARR (prior year)", round(begin_arr * 0.78, -3), "Per FY25 board pack", "Typed"),
     ("Rule of 40", "=B14+B13", "ARR growth + EBITDA margin", "Derived"),
     ("Net burn, Q3", -1_850_000, "Operating cash outflow", "Cash model (not in this pack)"),
     ("Burn multiple", "=-B17/B10", "Net burn / net new ARR", "Derived")]
for i, (k, v, d, s) in enumerate(K, start=4):
    put(ws, i, 1, k, font=BOLD); c = put(ws, i, 2, v, PCT if ("%" in k or "growth" in k or "NRR" in k or "margin" in k) else MONEY)
    if isinstance(v, (int, float)): c.fill = INPUT
    put(ws, i, 3, d); put(ws, i, 4, s)
note(ws["B6"], "Pasted from the BvA Summary tab Jul to Sep actual revenue before the September EU rows were appended there. Will not tie to the current BvA file.", "Maya Chen")
note(ws["B9"], "Investor convention: includes Customer Success cost. FP&A BvA gross margin excludes it. Both are correct; label them.", "Grace Park")
put(ws, 21, 1, "ARR by segment (active customers)", font=BOLD)
header(ws, 22, ["Segment", "ARR", "Customers", "Average ARR"])
for i, seg in enumerate(["SMB", "Mid-Market", "Enterprise"], start=23):
    put(ws, i, 1, seg); put(ws, i, 2, f'=SUMIFS(Customers!$H:$H,Customers!$C:$C,A{i},Customers!$G:$G,"")', MONEY)
    put(ws, i, 3, f'=COUNTIFS(Customers!$C:$C,A{i},Customers!$G:$G,"")'); put(ws, i, 4, f"=IFERROR(B{i}/C{i},0)", MONEY)
put(ws, 26, 1, "Total", font=BOLD); put(ws, 26, 2, "=SUM(B23:B25)", MONEY, font=BOLD, fill=TOTAL); put(ws, 26, 3, "=SUM(C23:C25)", font=BOLD, fill=TOTAL); put(ws, 26, 4, "=IFERROR(B26/C26,0)", MONEY, font=BOLD, fill=TOTAL)
put(ws, 29, 1, "Notes", font=BOLD)
for i, line in enumerate(["Gross margin here follows the investor convention (includes Customer Success cost). The FP&A BvA file excludes it. Reconcile before the board meeting.",
                          "Run-rate revenue is ARR / 12 and will not tie to GAAP revenue in the BvA file.",
                          "Segment ARR excludes churned customers (blank churn date) but ignores expansion since renewal; the bridge total is the governed ARR figure."], start=30): put(ws, i, 1, line)
widths(ws, [30, 14, 52, 56])

ws = wb.create_sheet("Cohorts")
put(ws, 1, 1, "Logo retention by start-quarter cohort (%)", font=BOLD)
put(ws, 2, 1, "Pasted from Looker dashboard 'Retention v3' on 2026-09-30. Not calculated in this workbook.", font=GREY)
header(ws, 4, ["Cohort"] + [f"M{k}" for k in (0, 3, 6, 9, 12, 18, 24)])
for i, q in enumerate(["2024-Q1", "2024-Q2", "2024-Q3", "2024-Q4", "2025-Q1", "2025-Q2", "2025-Q3", "2025-Q4", "2026-Q1", "2026-Q2"], start=5):
    put(ws, i, 1, q); v = 1.0
    for j in range(2, 9):
        if (i - 5) >= 10 - (j - 1): break
        put(ws, i, j, round(v, 3), PCT); v *= R.uniform(0.93, 0.99)
widths(ws, [10] + [8] * 7)

ws = wb.create_sheet("Definitions")
header(ws, 1, ["Metric", "Definition", "Owner", "Last reviewed", "Known differences"])
D = [("ARR", "Sum of contracted annual recurring revenue for active customers, plus expansion since renewal per the billing system.", "RevOps", dt.date(2026, 9, 30), "Customers tab excludes post-renewal expansion; ARR Bridge includes it."),
     ("Revenue (board pack)", "ARR / 12, run-rate.", "FP&A", dt.date(2026, 7, 1), "BvA file uses GAAP recognised revenue from the GL."),
     ("Gross margin (investor)", "(Revenue - COGS - Customer Success cost) / Revenue.", "CFO", dt.date(2026, 7, 1), "FP&A BvA gross margin = (Revenue - COGS) / Revenue; CS is Opex there."),
     ("NRR", "Trailing twelve months: (beginning ARR + expansion + contraction + churn) / beginning ARR for the cohort active 12 months ago.", "RevOps", dt.date(2026, 3, 1), "ARR Bridge B14 computes a YTD simplification, not TTM."),
     ("Rule of 40", "ARR growth YoY + EBITDA margin.", "CFO", dt.date(2026, 7, 1), "Some investor decks use revenue growth + FCF margin."),
     ("Burn multiple", "Net burn / net new ARR for the period.", "CFO", dt.date(2026, 7, 1), ""),
     ("Headcount", "Active employees at month end, FTE not weighted.", "People Ops", dt.date(2026, 1, 12), "Headcount Plan counts contractors separately; BvA snapshot pastes ending HC.")]
for i, row in enumerate(D, start=2):
    for j, v in enumerate(row, 1):
        c = put(ws, i, j, v)
        if j == 4: c.number_format = DATE
widths(ws, [24, 90, 12, 13, 80])
wb.save(os.path.join(OUT, "Ardent_Q3_FY26_Board_KPIs.xlsx"))

# ----------------------------------------------------------------------------- README with the planted findings
sep_eu_rows = sum(1 for r_ in ACT if r_[0] == m(9) and r_[1] == "EU")
readme = f"""# Ardent Labs sample workbooks (generated by make_samples.py, seed 26)

Fictional B2B SaaS company, about $30M ARR, two entities (US, EU). Three workbooks an FP&A team would actually keep.
No live external links: cross-workbook dependencies are pasted values with cell comments naming the source, which is how most
finance teams link files and the most common way lineage is lost. The shipped files were recalculated in Excel on 2026-10-12, so formula cells carry cached values.
If you regenerate with make_samples.py, repeat the recalc (see recalc note below).

## Files

| File | Sheets | What it exercises |
|---|---|---|
| Ardent_FY26_Budget_vs_Actuals.xlsx | Cover, Assumptions, COA, Departments, Actuals ({len(ACT)} rows, EU in EUR), Budget ({len(BUD)} rows), BvA Summary, Dept Variance, Commentary | key-value blocks, named ranges, long fact tables, dimension tables, wide report with SUMIFS + VLOOKUP + FX conversion, subtotal rows, a second table on the same sheet, hidden helper row, narrative commentary |
| Ardent_FY26_Headcount_Plan.xlsx | Roster ({len(roster)} rows), Comp Bands, Hiring Plan ({len(reqs)} rows), HC Summary | COUNTIFS/SUMIFS with date criteria and whole-column ranges, month-to-month chain (Beginning = prior Ending), typed constants copied from another workbook |
| Ardent_Q3_FY26_Board_KPIs.xlsx | Customers ({len(custs)} rows), ARR Bridge, KPI Summary, Cohorts, Definitions | cross-sheet formula chains, pasted values from the BvA file, a conflicting gross margin definition, a metric whose formula does not match its definition, a pasted Looker grid |

## Planted findings (what Finance Brain should surface)

1. **Stale hard-coded ranges.** BvA Summary and Dept Variance use `Actuals!$F$2:$F${ACT_END}` while Actuals has {len(ACT)} data rows
   (through row {len(ACT) + 1}). {len(excluded)} September EU rows are silently excluded ({sep_eu_rows} Sep EU rows exist).
   Excluded amounts at budget FX: {", ".join(f"{k} ${v:,.0f}" for k, v in sorted(excl_by_cat.items()))}. Cover change log v9 and Notes line 4 hint at it.
2. **Copied constant drift.** Benefits load is 22% in BvA Assumptions!B3 (named BenefitsLoad). Headcount Plan HC Summary!B2 has 22% typed
   (comment says copied 2026-07-14) and B3 has 25% for Q4, unapproved. Commentary Aug-26 describes it.
3. **Conflicting metric definitions.** Gross margin: BvA = (Revenue - COGS) / Revenue; Board KPI pack = (Revenue - COGS - CS cost) / Revenue.
   Revenue: BvA = GAAP from GL; board pack = ARR / 12. Both documented in Definitions and Commentary; the brain should answer "do these agree" with no.
4. **Formula does not match its definition.** Definitions says NRR is trailing twelve months; ARR Bridge!B14 computes a YTD simplification (comment admits it).
5. **Pasted values that will not tie.** KPI Summary!B6 (Q3 GAAP revenue) was pasted from the full GL, so it differs from the current BvA Summary
   Q3 revenue (which drops the excluded rows). Headcount snapshot on BvA Summary is pasted from HC Summary. Cohorts is pasted from Looker.
6. **Multiple tables per sheet and hidden structure.** BvA Summary has the P&L, a headcount table, and a notes block; row 3 is a hidden period-key row
   the SUMIFS depend on. KPI Summary has two tables plus notes.
7. **Joins to discover.** Actuals.Account -> COA.Account, Actuals.Department -> Departments.Dept code, Roster.Level -> Comp Bands.Level,
   Hiring Plan.Department -> Departments, Customers.Segment used by KPI Summary SUMIFS. VLOOKUPs make two of them explicit.
8. **Institutional knowledge in text.** Commentary tab (14 entries), cell comments (owners, dates, provenance), Definitions, Cover change log.

## Recalc note

openpyxl writes formulas without cached values. To store values: open each file in Excel and save, or run
`soffice --headless --convert-to xlsx --outdir recalc/ *.xlsx` with LibreOffice. The capture pipeline reads formulas with
`data_only=False`, so lineage works either way; cached values only matter for showing numbers in the UI.
"""
open(os.path.join(OUT, "README.md"), "w").write(readme)
print(f"wrote 3 workbooks to {OUT}; actuals={len(ACT)} budget={len(BUD)} roster={len(roster)} reqs={len(reqs)} customers={len(custs)} ACT_END={ACT_END} excluded={len(excluded)}")
