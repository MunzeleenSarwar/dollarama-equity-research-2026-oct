#!/usr/bin/env python3
"""
Build the Dollarama (TSX: DOL) Excel financial model.

Sheets: Cover | Historicals & Ratios | Forecast & 3-Statement | DCF & Sensitivity | Comps
All forecast, valuation and sensitivity cells are live Excel formulas.
Historical figures come from data/statements.json (StockAnalysis.com, tracing
to Dollarama's filed statements). CAD millions unless noted.

Method notes (also on the Cover sheet):
- IFRS 16 leases are treated as financing: FCFF is pre-lease-payment and total
  debt in the EV-to-equity bridge INCLUDES lease liabilities.
- WACC uses stated capital-market assumptions (see DCF sheet), not fetched data.
"""
import json, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
DATA = os.path.join(ROOT, "data")
OUT = os.path.join(ROOT, "Dollarama_Financial_Model.xlsx")

S = json.load(open(os.path.join(DATA, "statements.json")))
inc, bal, cf = S["income"], S["balance"], S["cashflow"]
FYS = ["FY 2022", "FY 2023", "FY 2024", "FY 2025", "FY 2026"]
FC = ["FY 2027E", "FY 2028E", "FY 2029E", "FY 2030E", "FY 2031E"]


def g(d, k, p):
    v = d.get(k, {}).get(p)
    return None if v is None else float(v)


NAVY = "1F3864"
HDR = PatternFill("solid", fgColor=NAVY)
SUB = PatternFill("solid", fgColor="D9E2EC")
INP = PatternFill("solid", fgColor="FFF4CC")
WHT = Font(color="FFFFFF", bold=True, size=11)
BOLD = Font(color="111111", bold=True, size=11)
NORM = Font(color="111111", size=11)
TITLE = Font(color=NAVY, bold=True, size=15)
THIN = Border(bottom=Side(style="thin", color="B0B0B0"))
MONEY = '#,##0'
PCT = '0.0%'
PRICE = '0.00'


def put(ws, row, col, val, font=NORM, fill=None, fmt=None, align=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font = font
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    if align: c.alignment = align
    return c


def header_row(ws, row, text, ncols=8):
    put(ws, row, 1, text, font=WHT, fill=HDR)
    for i in range(2, ncols + 1):
        put(ws, row, i, None, fill=HDR)


wb = Workbook()

# ============================ COVER ============================
ws = wb.active
ws.title = "Cover"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.column_dimensions["A"].width = 26
for col in "BCDEFG":
    ws.column_dimensions[col].width = 15
put(ws, 1, 1, "Dollarama Inc. (TSX: DOL) — Financial Model & Valuation", font=TITLE)
put(ws, 2, 1, "Prepared by Munzeleen Sarwar  •  October 2026  •  Educational portfolio project — not investment advice", font=Font(italic=True, color="555555"))
rows = [
    ("", ""),
    ("CONTENTS", ""),
    ("Historicals & Ratios", "FY2022–FY2026 financials from Dollarama's filed statements (via StockAnalysis.com) with computed ratios"),
    ("Forecast & 3-Statement", "Integrated income statement, condensed balance sheet and cash-flow summary, FY2027E–FY2031E, driven by explicit assumptions"),
    ("DCF & Sensitivity", "Free-cash-flow-to-firm DCF, WACC build, sensitivity grid and valuation football field"),
    ("Comps", "Trading comparables: North American value retail and Canadian food/convenience retail, October 2026"),
    ("", ""),
    ("KEY OUTPUTS (live from DCF sheet)", ""),
    ("DCF implied share price (CAD)", "='DCF & Sensitivity'!B31"),
    ("Share price, 21 Oct 2026 close (CAD)", 204.59),
    ("Implied upside / (downside)", "='DCF & Sensitivity'!B32"),
    ("Recommendation implied by base case", "HOLD — a high-quality compounder priced at a premium to base-case value"),
    ("", ""),
    ("METHOD NOTES", ""),
    ("Leases (IFRS 16)", "FCFF is calculated before lease payments; total debt in the bridge includes lease liabilities (FY2026: $5,396M incl. $2,771M leases)"),
    ("WACC inputs", "Risk-free rate, beta and equity risk premium are stated assumptions on the DCF sheet (beta is an analyst assumption for a defensive retailer, not a fetched figure)"),
    ("Equity base", "Dollarama's book equity is small (and was negative in FY2022–23) because of large buybacks — ROE is not meaningful and is excluded; ROA and margins are used instead"),
    ("Data check", "All historical totals are as reported by the data source; check Dollarama's SEDAR+ filings before any investment use"),
]
r = 4
for a, b in rows:
    if a and not b:
        header_row(ws, r, a, ncols=4)
    else:
        put(ws, r, 1, a, font=BOLD if a else NORM)
        if isinstance(b, str) and b.startswith("="):
            put(ws, r, 2, b, font=BOLD, fmt=PRICE if "price" in a.lower() else PCT)
        elif isinstance(b, (int, float)):
            put(ws, r, 2, b, font=BOLD, fmt=PRICE)
        else:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            put(ws, r, 2, b)
    r += 1

# ===================== HISTORICALS & RATIOS =====================
ws = wb.create_sheet("Historicals & Ratios")
ws.column_dimensions["A"].width = 42
for i in range(2, 8):
    ws.column_dimensions[get_column_letter(i)].width = 14
put(ws, 1, 1, "Dollarama — Historical Financials & Ratios (CAD millions, fiscal years ending late Jan / early Feb)", font=TITLE)
header_row(ws, 3, "", ncols=7)
put(ws, 3, 1, "", fill=HDR)
for j, fy in enumerate(FYS):
    put(ws, 3, 2 + j, fy, font=WHT, fill=HDR, align=Alignment(horizontal="right"))

IS_LINES = [("Revenue", "Revenue"), ("Cost of Revenue", "Cost of Revenue"), ("Gross Profit", "Gross Profit"),
            ("Selling, General & Admin", "Selling, General & Admin"), ("Operating Expenses (total)", "Operating Expenses"),
            ("Operating Income (EBIT)", "Operating Income"), ("Interest Expense", "Interest Expense"),
            ("Pretax Income", "Pretax Income"), ("Income Tax Expense", "Income Tax Expense"),
            ("Net Income", "Net Income"), ("EBITDA (as reported by source)", "EBITDA"),
            ("EPS Diluted (CAD)", "EPS (Diluted)"), ("Shares Outstanding Diluted (M)", "Shares Outstanding (Diluted)")]
BS_LINES = [("Cash & Short-Term Investments", "Cash & Short-Term Investments"), ("Inventory", "Inventory"),
            ("Total Current Assets", "Total Current Assets"), ("Property, Plant & Equipment (incl. ROU as aggregated)", "Property, Plant & Equipment"),
            ("Total Assets", "Total Assets"), ("Total Current Liabilities", "Total Current Liabilities"),
            ("Long-Term Debt", "Long-Term Debt"), ("Long-Term Leases", "Long-Term Leases"),
            ("Total Liabilities", "Total Liabilities"), ("Shareholders' Equity", "Shareholders' Equity"),
            ("Shares Outstanding (M, period end)", "Total Common Shares Outstanding")]
CF_LINES = [("Depreciation & Amortization", "Depreciation & Amortization"), ("Operating Cash Flow", "Operating Cash Flow"),
            ("Capital Expenditures", "Capital Expenditures"), ("Free Cash Flow", "Free Cash Flow"),
            ("Lease Payments", "Lease Payments"), ("Repurchase of Common Stock (buybacks)", "Repurchase of Common Stock"),
            ("Common Dividends Paid", "Common Dividends Paid")]

r = 4
sheet_rows = {}
for title, lines, src in [("INCOME STATEMENT", IS_LINES, inc), ("BALANCE SHEET", BS_LINES, bal), ("CASH FLOW", CF_LINES, cf)]:
    header_row(ws, r, title, ncols=7)
    r += 1
    for label, key in lines:
        put(ws, r, 1, label)
        for j, fy in enumerate(FYS):
            v = g(src, key, fy)
            fmt = PRICE if ("EPS" in label or "Shares" in label) else MONEY
            put(ws, r, 2 + j, v, fmt=fmt, align=Alignment(horizontal="right"))
        sheet_rows[label] = r
        r += 1
    r += 1

header_row(ws, r, "COMPUTED RATIOS", ncols=7)
r += 1
REV = sheet_rows["Revenue"]; GP = sheet_rows["Gross Profit"]; EBIT = sheet_rows["Operating Income (EBIT)"]
NI = sheet_rows["Net Income"]; TA = sheet_rows["Total Assets"]; TCA = sheet_rows["Total Current Assets"]
TCL = sheet_rows["Total Current Liabilities"]; INV = sheet_rows["Inventory"]; COR = sheet_rows["Cost of Revenue"]
FCF = sheet_rows["Free Cash Flow"]; EQ = sheet_rows["Shareholders' Equity"]
ratio_defs = [
    ("Revenue growth", f"={get_column_letter(3)}{REV}/{get_column_letter(2)}{REV}-1", PCT, 3),  # computed per-col below instead
]
ratio_rows = [
    ("Gross margin", lambda c: f"={c}{GP}/{c}{REV}", PCT),
    ("Operating (EBIT) margin", lambda c: f"={c}{EBIT}/{c}{REV}", PCT),
    ("Net margin", lambda c: f"={c}{NI}/{c}{REV}", PCT),
    ("Return on assets (NI / total assets)", lambda c: f"={c}{NI}/{c}{TA}", PCT),
    ("Current ratio", lambda c: f"={c}{TCA}/{c}{TCL}", '0.00'),
    ("Inventory days (on cost of revenue)", lambda c: f"={c}{INV}/{c}{COR}*365", '0'),
    ("FCF margin", lambda c: f"={c}{FCF}/{c}{REV}", PCT),
]
for label, fn, fmt in ratio_rows:
    put(ws, r, 1, label)
    for j in range(5):
        c = get_column_letter(2 + j)
        put(ws, r, 2 + j, fn(c), fmt=fmt, align=Alignment(horizontal="right"))
    r += 1
# revenue growth row (needs prior column; first col blank)
put(ws, r, 1, "Revenue growth (YoY)")
for j in range(1, 5):
    c, p = get_column_letter(2 + j), get_column_letter(1 + j)
    put(ws, r, 2 + j, f"={c}{REV}/{p}{REV}-1", fmt=PCT, align=Alignment(horizontal="right"))
r += 2
put(ws, r, 1, "Note: ROE is deliberately excluded — buybacks shrank the equity base (negative in FY2022–FY2023), so ROE is not meaningful for Dollarama.", font=Font(italic=True, color="555555"))

# ================= FORECAST & 3-STATEMENT =================
ws = wb.create_sheet("Forecast & 3-Statement")
ws.column_dimensions["A"].width = 46
for i in range(2, 8):
    ws.column_dimensions[get_column_letter(i)].width = 14
put(ws, 1, 1, "Dollarama — Integrated Forecast (CAD millions unless noted). Yellow cells are assumptions; everything else is a formula.", font=TITLE)
header_row(ws, 3, "", ncols=7)
put(ws, 3, 2, "FY 2026A", font=WHT, fill=HDR, align=Alignment(horizontal="right"))
for j, fy in enumerate(FC):
    put(ws, 3, 3 + j, fy, font=WHT, fill=HDR, align=Alignment(horizontal="right"))

fy26 = {
    "rev": g(inc, "Revenue", "FY 2026"), "cor": g(inc, "Cost of Revenue", "FY 2026"),
    "gp": g(inc, "Gross Profit", "FY 2026"), "sga": g(inc, "Selling, General & Admin", "FY 2026"),
    "opex": g(inc, "Operating Expenses", "FY 2026"), "ebit": g(inc, "Operating Income", "FY 2026"),
    "tax_rate": g(inc, "Effective Tax Rate", "FY 2026"),
}
fy26["da"] = fy26["opex"] - fy26["sga"]
fy26["interest"] = -g(inc, "Interest Expense", "FY 2026")  # source stores expense as negative
fy26["other_inc"] = g(inc, "Pretax Income", "FY 2026") - fy26["ebit"] + fy26["interest"]
fy26["pretax"] = g(inc, "Pretax Income", "FY 2026"); fy26["tax"] = g(inc, "Income Tax Expense", "FY 2026")
fy26["ni"] = g(inc, "Net Income", "FY 2026")
cash26 = g(bal, "Cash & Short-Term Investments", "FY 2026")
opca26 = g(bal, "Total Current Assets", "FY 2026") - cash26
ppe26 = g(bal, "Property, Plant & Equipment", "FY 2026")
gw26 = g(bal, "Total Assets", "FY 2026") - g(bal, "Total Current Assets", "FY 2026") - ppe26
curr_lease26 = g(bal, "Current Portion of Leases", "FY 2026") or 0
curr_ltd26 = g(bal, "Current Portion of Long-Term Debt", "FY 2026") or 0
st_debt26 = g(bal, "Short-Term Debt", "FY 2026") or 0
opcl26 = g(bal, "Total Current Liabilities", "FY 2026") - curr_lease26 - curr_ltd26 - st_debt26
lt_debt26 = g(bal, "Long-Term Debt", "FY 2026")
debt26 = lt_debt26 + curr_ltd26 + st_debt26
leases26 = g(bal, "Long-Term Leases", "FY 2026") + curr_lease26
otherlt26 = g(bal, "Total Liabilities", "FY 2026") - g(bal, "Total Current Liabilities", "FY 2026") - lt_debt26 - g(bal, "Long-Term Leases", "FY 2026")
eq26 = g(bal, "Shareholders' Equity", "FY 2026")
shares26 = g(bal, "Total Common Shares Outstanding", "FY 2026")
capex26 = -g(cf, "Capital Expenditures", "FY 2026")
leasepay26 = -g(cf, "Lease Payments", "FY 2026")
div26 = -g(cf, "Common Dividends Paid", "FY 2026")
buy26 = -g(cf, "Repurchase of Common Stock", "FY 2026")

DRV = {}
header_row(ws, 4, "DRIVERS & ASSUMPTIONS", ncols=7)
drivers = [
    ("Revenue growth", [None, .13, .10, .085, .075, .065], PCT),
    ("Gross margin", [fy26["gp"] / fy26["rev"], .450, .450, .450, .450, .450], PCT),
    ("SG&A (% of revenue)", [fy26["sga"] / fy26["rev"], .150, .150, .149, .149, .148], PCT),
    ("D&A (% of revenue)", [fy26["da"] / fy26["rev"], .060, .060, .060, .060, .060], PCT),
    ("Capex (% of revenue)", [capex26 / fy26["rev"], .035, .035, .035, .035, .035], PCT),
    ("Operating current assets (% of revenue)", [opca26 / fy26["rev"]] + [opca26 / fy26["rev"]] * 5, PCT),
    ("Operating current liabilities (% of revenue)", [opcl26 / fy26["rev"]] + [opcl26 / fy26["rev"]] * 5, PCT),
    ("PP&E/ROU additions from leases (% of revenue)", [None, .045, .045, .045, .045, .045], PCT),
    ("Tax rate", [fy26["tax_rate"], .258, .258, .258, .258, .258], PCT),
    ("Interest expense ($M)", [fy26["interest"], 215, 220, 225, 230, 235], MONEY),
    ("Equity & other income ($M)", [fy26["other_inc"], 200, 205, 210, 215, 220], MONEY),
    ("Dividend payout (% of net income)", [div26 / fy26["ni"], .09, .09, .09, .09, .09], PCT),
    ("Share buybacks ($M per year)", [buy26, 900, 900, 900, 900, 900], MONEY),
    ("Lease liabilities growth", [None, .05, .05, .05, .05, .05], PCT),
    ("Share count change (buyback effect)", [None, -.012, -.012, -.012, -.012, -.012], PCT),
]
r = 5
for label, vals, fmt in drivers:
    put(ws, r, 1, label)
    for j, v in enumerate(vals):
        put(ws, r, 2 + j, v, fmt=fmt, fill=INP if j > 0 else None, align=Alignment(horizontal="right"))
    DRV[label] = r
    r += 1
r += 1
IS0 = r
header_row(ws, r, "INCOME STATEMENT", ncols=7)
r += 1
IS = {}
is_lines = ["Revenue", "Cost of Revenue", "Gross Profit", "SG&A", "Depreciation & Amortization", "Operating Income (EBIT)",
            "Interest Expense", "Equity & Other Income", "Pretax Income", "Income Tax", "Net Income", "EPS Diluted (CAD)",
            "EBIT Margin", "Net Margin"]
for lab in is_lines:
    IS[lab] = r
    put(ws, r, 1, lab, font=BOLD if lab in ("Gross Profit", "Operating Income (EBIT)", "Net Income") else NORM)
    r += 1
# FY26 actuals in col B
actuals = {"Revenue": fy26["rev"], "Cost of Revenue": fy26["cor"], "Gross Profit": fy26["gp"], "SG&A": fy26["sga"],
           "Depreciation & Amortization": fy26["da"], "Operating Income (EBIT)": fy26["ebit"],
           "Interest Expense": fy26["interest"], "Equity & Other Income": fy26["other_inc"],
           "Pretax Income": fy26["pretax"], "Income Tax": fy26["tax"], "Net Income": fy26["ni"]}
for lab, v in actuals.items():
    put(ws, IS[lab], 2, v, fmt=MONEY, align=Alignment(horizontal="right"))
put(ws, IS["EPS Diluted (CAD)"], 2, g(inc, "EPS (Diluted)", "FY 2026"), fmt=PRICE, align=Alignment(horizontal="right"))
put(ws, IS["EBIT Margin"], 2, f"=B{IS['Operating Income (EBIT)']}/B{IS['Revenue']}", fmt=PCT, align=Alignment(horizontal="right"))
put(ws, IS["Net Margin"], 2, f"=B{IS['Net Income']}/B{IS['Revenue']}", fmt=PCT, align=Alignment(horizontal="right"))

BS0 = r + 1
header_row(ws, BS0, "BALANCE SHEET (condensed, integrated)", ncols=7)
BS = {}
bs_lines = ["Cash & Short-Term Investments (plug)", "Operating Current Assets", "Net PP&E and Right-of-Use Assets",
            "Goodwill & Other Long-Term Assets", "Total Assets", "Operating Current Liabilities", "Debt (excl. leases)",
            "Lease Liabilities", "Other Long-Term Liabilities", "Total Liabilities", "Shareholders' Equity",
            "Total Liabilities & Equity", "Balance Check (must be 0)", "Shares Outstanding (M)"]
rr = BS0 + 1
for lab in bs_lines:
    BS[lab] = rr
    put(ws, rr, 1, lab, font=BOLD if lab in ("Total Assets", "Total Liabilities", "Total Liabilities & Equity", "Balance Check (must be 0)") else NORM)
    rr += 1
bs_actuals = {"Cash & Short-Term Investments (plug)": cash26, "Operating Current Assets": opca26,
              "Net PP&E and Right-of-Use Assets": ppe26, "Goodwill & Other Long-Term Assets": gw26,
              "Total Assets": g(bal, "Total Assets", "FY 2026"), "Operating Current Liabilities": opcl26,
              "Debt (excl. leases)": debt26, "Lease Liabilities": leases26, "Other Long-Term Liabilities": otherlt26,
              "Total Liabilities": g(bal, "Total Liabilities", "FY 2026"), "Shareholders' Equity": eq26,
              "Total Liabilities & Equity": g(bal, "Total Assets", "FY 2026"), "Balance Check (must be 0)": 0,
              "Shares Outstanding (M)": shares26}
for lab, v in bs_actuals.items():
    put(ws, BS[lab], 2, v, fmt=PRICE if lab == "Shares Outstanding (M)" else MONEY, align=Alignment(horizontal="right"))

CF0 = rr + 1
header_row(ws, CF0, "CASH FLOW (summary, derived from IS & BS)", ncols=7)
CF = {}
cf_lines = ["Net Income", "+ Depreciation & Amortization", "- Increase in Net Working Capital", "Operating Cash Flow",
            "Capital Expenditures", "Free Cash Flow (pre-lease payments)", "Lease Payments", "Dividends Paid",
            "Share Buybacks", "Change in Cash (check vs BS)"]
rr = CF0 + 1
for lab in cf_lines:
    CF[lab] = rr
    put(ws, rr, 1, lab, font=BOLD if lab in ("Operating Cash Flow", "Free Cash Flow (pre-lease payments)") else NORM)
    rr += 1
cf_actuals = {"Net Income": fy26["ni"], "+ Depreciation & Amortization": fy26["da"],
              "Operating Cash Flow": g(cf, "Operating Cash Flow", "FY 2026"), "Capital Expenditures": -capex26,
              "Lease Payments": -leasepay26, "Dividends Paid": -div26, "Share Buybacks": -buy26}
for lab, v in cf_actuals.items():
    put(ws, CF[lab], 2, v, fmt=MONEY, align=Alignment(horizontal="right"))

# Forecast formulas, columns C..G (3..7)
for j in range(5):
    col = 3 + j
    L = get_column_letter(col); P = get_column_letter(col - 1)
    put(ws, IS["Revenue"], col, f"={P}{IS['Revenue']}*(1+{L}{DRV['Revenue growth']})", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["Cost of Revenue"], col, f"={L}{IS['Revenue']}*(1-{L}{DRV['Gross margin']})", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["Gross Profit"], col, f"={L}{IS['Revenue']}-{L}{IS['Cost of Revenue']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["SG&A"], col, f"={L}{IS['Revenue']}*{L}{DRV['SG&A (% of revenue)']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["Depreciation & Amortization"], col, f"={L}{IS['Revenue']}*{L}{DRV['D&A (% of revenue)']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["Operating Income (EBIT)"], col, f"={L}{IS['Gross Profit']}-{L}{IS['SG&A']}-{L}{IS['Depreciation & Amortization']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["Interest Expense"], col, f"={L}{DRV['Interest expense ($M)']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["Equity & Other Income"], col, f"={L}{DRV['Equity & other income ($M)']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["Pretax Income"], col, f"={L}{IS['Operating Income (EBIT)']}-{L}{IS['Interest Expense']}+{L}{IS['Equity & Other Income']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["Income Tax"], col, f"={L}{IS['Pretax Income']}*{L}{DRV['Tax rate']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["Net Income"], col, f"={L}{IS['Pretax Income']}-{L}{IS['Income Tax']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, IS["EPS Diluted (CAD)"], col, f"={L}{IS['Net Income']}/{L}{BS['Shares Outstanding (M)']}", fmt=PRICE, align=Alignment(horizontal="right"))
    put(ws, IS["EBIT Margin"], col, f"={L}{IS['Operating Income (EBIT)']}/{L}{IS['Revenue']}", fmt=PCT, align=Alignment(horizontal="right"))
    put(ws, IS["Net Margin"], col, f"={L}{IS['Net Income']}/{L}{IS['Revenue']}", fmt=PCT, align=Alignment(horizontal="right"))
    # BS
    put(ws, BS["Operating Current Assets"], col, f"={L}{IS['Revenue']}*{L}{DRV['Operating current assets (% of revenue)']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Operating Current Liabilities"], col, f"={L}{IS['Revenue']}*{L}{DRV['Operating current liabilities (% of revenue)']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Net PP&E and Right-of-Use Assets"], col, f"={P}{BS['Net PP&E and Right-of-Use Assets']}+{L}{IS['Revenue']}*{L}{DRV['Capex (% of revenue)']}+{L}{IS['Revenue']}*{L}{DRV['PP&E/ROU additions from leases (% of revenue)']}-{L}{IS['Depreciation & Amortization']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Goodwill & Other Long-Term Assets"], col, f"={P}{BS['Goodwill & Other Long-Term Assets']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Debt (excl. leases)"], col, f"={P}{BS['Debt (excl. leases)']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Lease Liabilities"], col, f"={P}{BS['Lease Liabilities']}*(1+{L}{DRV['Lease liabilities growth']})", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Other Long-Term Liabilities"], col, f"={P}{BS['Other Long-Term Liabilities']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Total Liabilities"], col, f"={L}{BS['Operating Current Liabilities']}+{L}{BS['Debt (excl. leases)']}+{L}{BS['Lease Liabilities']}+{L}{BS['Other Long-Term Liabilities']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Shares Outstanding (M)"], col, f"={P}{BS['Shares Outstanding (M)']}*(1+{L}{DRV['Share count change (buyback effect)']})", fmt=PRICE, align=Alignment(horizontal="right"))
    # Equity, totals, cash plug and checks are written in the fix-up loop below
    # (equity needs the cash-flow rows for dividends and buybacks).
    # CF first (equity needs dividends & buybacks)
    put(ws, CF["Net Income"], col, f"={L}{IS['Net Income']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, CF["+ Depreciation & Amortization"], col, f"={L}{IS['Depreciation & Amortization']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, CF["- Increase in Net Working Capital"], col, f"=-(({L}{BS['Operating Current Assets']}-{L}{BS['Operating Current Liabilities']})-({P}{BS['Operating Current Assets']}-{P}{BS['Operating Current Liabilities']}))", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, CF["Operating Cash Flow"], col, f"=SUM({L}{CF['Net Income']}:{L}{CF['- Increase in Net Working Capital']})", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, CF["Capital Expenditures"], col, f"=-{L}{IS['Revenue']}*{L}{DRV['Capex (% of revenue)']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, CF["Free Cash Flow (pre-lease payments)"], col, f"={L}{CF['Operating Cash Flow']}+{L}{CF['Capital Expenditures']}", fmt=MONEY, align=Alignment(horizontal="right"))
    # Lease Payments written in the fix-up loop below (it chains off the prior year).
    put(ws, CF["Dividends Paid"], col, f"=-{L}{IS['Net Income']}*{L}{DRV['Dividend payout (% of net income)']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, CF["Share Buybacks"], col, f"=-{L}{DRV['Share buybacks ($M per year)']}", fmt=MONEY, align=Alignment(horizontal="right"))
# Fix the two cells written awkwardly above, per forecast column
for j in range(5):
    col = 3 + j
    L = get_column_letter(col); P = get_column_letter(col - 1)
    # Lease payments: grow FY26 payment with lease-liability growth driver
    prev = f"B{CF['Lease Payments']}" if j == 0 else f"{P}{CF['Lease Payments']}"
    put(ws, CF["Lease Payments"], col, f"={prev}*(1+{L}{DRV['Lease liabilities growth']})", fmt=MONEY, align=Alignment(horizontal="right"))
    # Equity roll-forward: prior equity + net income - dividends - buybacks
    # (Dividends Paid and Share Buybacks are negative cash-flow rows, so they are added.)
    EQ = BS["Shareholders' Equity"]
    put(ws, EQ, col, f"={P}{EQ}+{L}{IS['Net Income']}+{L}{CF['Dividends Paid']}+{L}{CF['Share Buybacks']}", fmt=MONEY, align=Alignment(horizontal="right"))
    # Totals & plug & checks
    put(ws, BS["Total Liabilities & Equity"], col, f"={L}{BS['Total Liabilities']}+{L}{EQ}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Cash & Short-Term Investments (plug)"], col, f"={L}{BS['Total Liabilities & Equity']}-{L}{BS['Operating Current Assets']}-{L}{BS['Net PP&E and Right-of-Use Assets']}-{L}{BS['Goodwill & Other Long-Term Assets']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Total Assets"], col, f"={L}{BS['Cash & Short-Term Investments (plug)']}+{L}{BS['Operating Current Assets']}+{L}{BS['Net PP&E and Right-of-Use Assets']}+{L}{BS['Goodwill & Other Long-Term Assets']}", fmt=MONEY, align=Alignment(horizontal="right"))
    put(ws, BS["Balance Check (must be 0)"], col, f"={L}{BS['Total Assets']}-{L}{BS['Total Liabilities & Equity']}", fmt=MONEY, align=Alignment(horizontal="right"))
    # Cash check: lease payments are already inside the movement in lease
    # liabilities (additions - payments), so the BS cash movement must equal
    # FCF + ΔLease liabilities - non-cash lease (ROU) additions + dividends + buybacks.
    put(ws, CF["Change in Cash (check vs BS)"], col, f"={L}{CF['Free Cash Flow (pre-lease payments)']}+({L}{BS['Lease Liabilities']}-{P}{BS['Lease Liabilities']})-{L}{IS['Revenue']}*{L}{DRV['PP&E/ROU additions from leases (% of revenue)']}+{L}{CF['Dividends Paid']}+{L}{CF['Share Buybacks']}-({L}{BS['Cash & Short-Term Investments (plug)']}-{P}{BS['Cash & Short-Term Investments (plug)']})", fmt=MONEY, align=Alignment(horizontal="right"))

FORECAST_ROWS = {"IS": IS, "BS": BS, "CF": CF, "DRV": DRV}

# ===================== DCF & SENSITIVITY =====================
ws = wb.create_sheet("DCF & Sensitivity")
ws.column_dimensions["A"].width = 46
for i in range(2, 9):
    ws.column_dimensions[get_column_letter(i)].width = 14
FS = "'Forecast & 3-Statement'"
put(ws, 1, 1, "Dollarama — DCF Valuation (FCFF, leases treated as debt)", font=TITLE)
header_row(ws, 3, "WACC BUILD — inputs are stated assumptions (yellow)", ncols=3)
wacc_rows = [
    ("Risk-free rate (GoC 10Y, assumption)", .034, PCT),
    ("Beta (analyst assumption — defensive retailer)", .60, '0.00'),
    ("Equity risk premium (assumption)", .050, PCT),
    ("Cost of equity (CAPM)", "=B4+B5*B6", PCT),
    ("Pre-tax cost of debt (assumption)", .043, PCT),
    ("Tax rate", .258, PCT),
    ("After-tax cost of debt", "=B8*(1-B9)", PCT),
    ("Market value of equity ($M, Oct 2026)", 54906, MONEY),
    ("Total debt incl. leases ($M, FY2026)", 5396, MONEY),
    ("Weight of equity", "=B11/(B11+B12)", PCT),
    ("Weight of debt", "=B12/(B11+B12)", PCT),
    ("WACC", "=B13*B7+B14*B10", PCT),
]
r = 4
for lab, v, fmt in wacc_rows:
    put(ws, r, 1, lab)
    put(ws, r, 2, v, fmt=fmt, fill=INP if isinstance(v, float) and v < 10 else None, align=Alignment(horizontal="right"))
    r += 1
WACC_CELL = "B18"
# NOTE: rows 4..15 as written; WACC is the last row
WACC_CELL = f"B{r-1}"
r += 1
put(ws, r, 1, "DCF", font=WHT, fill=HDR)
for j, fy in enumerate(FC):
    put(ws, r, 3 + j, fy, font=WHT, fill=HDR, align=Alignment(horizontal="right"))
DCFH = r
r += 1
fcff_r, pv_r = r, r + 3
put(ws, fcff_r, 1, "Free Cash Flow to Firm (EBIT(1-t) + D&A - Capex - ΔWC)")
for j in range(5):
    L = get_column_letter(3 + j)
    put(ws, fcff_r, 3 + j, f"=({FS}!{L}{IS['Operating Income (EBIT)']}*(1-{FS}!{L}{DRV['Tax rate']})+{FS}!{L}{IS['Depreciation & Amortization']}+{FS}!{L}{CF['Capital Expenditures']}+{FS}!{L}{CF['- Increase in Net Working Capital']})", fmt=MONEY, align=Alignment(horizontal="right"))
put(ws, fcff_r + 1, 1, "Discount years")
for j in range(5):
    put(ws, fcff_r + 1, 3 + j, j + 1, align=Alignment(horizontal="right"))
put(ws, fcff_r + 2, 1, "Discount factor")
for j in range(5):
    L = get_column_letter(3 + j)
    put(ws, fcff_r + 2, 3 + j, f"=1/(1+${WACC_CELL})^{L}{fcff_r+1}", fmt='0.000', align=Alignment(horizontal="right"))
put(ws, pv_r, 1, "PV of FCFF")
for j in range(5):
    L = get_column_letter(3 + j)
    put(ws, pv_r, 3 + j, f"={L}{fcff_r}*{L}{fcff_r+2}", fmt=MONEY, align=Alignment(horizontal="right"))
r = pv_r + 2
outs = [
    ("PV of FCFF FY2027E–FY2031E", f"=SUM(C{pv_r}:G{pv_r})", MONEY),
    ("Terminal growth rate (assumption)", .025, PCT),
    ("Terminal value", f"=G{fcff_r}*(1+B{r+1})/({WACC_CELL}-B{r+1})", MONEY),
    ("PV of terminal value", f"=B{r+2}*G{fcff_r+2}", MONEY),
    ("Enterprise value", f"=B{r}+B{r+3}", MONEY),
    ("Less: net debt incl. leases (FY2026 total debt 5,396 - cash 340)", -5056, MONEY),
    ("Equity value", f"=B{r+4}+B{r+5}", MONEY),
    ("Shares outstanding (M)", shares26, PRICE),
    ("Implied share price (CAD)", f"=B{r+6}*1000000/(B{r+7}*1000000)", PRICE),
    ("Implied upside/(downside) vs $204.59", f"=B{r+8}/204.59-1", PCT),
]
price_row = None
for lab, v, fmt in outs:
    put(ws, r, 1, lab, font=BOLD if "Implied share price" in lab else NORM)
    put(ws, r, 2, v, fmt=fmt, fill=INP if lab.startswith("Terminal growth") else None, align=Alignment(horizontal="right"))
    if "Implied share price" in lab: price_row = r
    r += 1
# Cover references B22/B23 assumed implied price at row (price_row) — record actual rows for post-fix
MODEL_ROWS = {"wacc": WACC_CELL, "fcff": fcff_r, "price_row": price_row, "pv_row": pv_r}
r += 1
header_row(ws, r, "SENSITIVITY — Implied share price (CAD): WACC (rows) vs terminal growth (columns)", ncols=7)
r += 1
g_vals = [.015, .02, .025, .03, .035]
w_vals = [.051, .056, .061, .066, .071]
put(ws, r, 1, "WACC \\ g")
for j, gv in enumerate(g_vals):
    put(ws, r, 2 + j, gv, fmt=PCT, font=BOLD, align=Alignment(horizontal="right"))
grid_top = r + 1
FCFF_CELLS = [f"{get_column_letter(3+j)}{fcff_r}" for j in range(5)]
for i, wv in enumerate(w_vals):
    rr_ = grid_top + i
    put(ws, rr_, 1, wv, fmt=PCT, font=BOLD, align=Alignment(horizontal="right"))
    for j, gv in enumerate(g_vals):
        pv_parts = "+".join([f"{FCFF_CELLS[k]}/(1+{wv})^{k+1}" for k in range(5)])
        tv = f"({FCFF_CELLS[4]}*(1+{gv})/({wv}-{gv}))/(1+{wv})^5"
        put(ws, rr_, 2 + j, f"=({pv_parts}+{tv}-5056)/{shares26}", fmt=PRICE, align=Alignment(horizontal="right"))
r = grid_top + len(w_vals) + 1
header_row(ws, r, "FOOTBALL FIELD (CAD per share)", ncols=4)
r += 1
ff = [
    # low = WACC 6.6% / g 2.0% cell; high = WACC 5.6% / g 3.0% cell of the grid above
    ("DCF — sensitivity range (WACC 5.6–6.6%, g 2.0–3.0%)",
     f"=C{grid_top+3}", f"=E{grid_top+1}"),
    ("DCF — base case", f"=B{price_row}", f"=B{price_row}"),
    ("52-week trading range (Oct 2026)", 163.25, 209.96),
    ("Current price (21 Oct 2026)", 204.59, 204.59),
]
for lab, lo, hi in ff:
    put(ws, r, 1, lab)
    put(ws, r, 2, lo, fmt=PRICE, align=Alignment(horizontal="right"))
    put(ws, r, 3, hi, fmt=PRICE, align=Alignment(horizontal="right"))
    r += 1
FF_ROWS = (grid_top + 1, grid_top + 3)  # sensitivity rows used for range: middle-ish rows

# ============================ COMPS ============================
ws = wb.create_sheet("Comps")
ws.column_dimensions["A"].width = 34
for i in range(2, 11):
    ws.column_dimensions[get_column_letter(i)].width = 13.5
put(ws, 1, 1, "Trading Comparables — October 2026 (TTM; US peers in USD — ratios are currency-neutral)", font=TITLE)
headers = ["Company", "P/E", "Forward P/E", "P/S", "P/FCF", "Gross Margin", "Operating Margin", "Net Margin", "Revenue Growth", "Dividend Yield"]
header_row(ws, 3, "", ncols=10)
for j, h in enumerate(headers):
    put(ws, 3, 1 + j, h, font=WHT, fill=HDR, align=Alignment(horizontal="right") if j else None)
peers = json.load(open(os.path.join(DATA, "peers.json")))
want = ["PE Ratio", "Forward PE", "PS Ratio", "P/FCF Ratio", "Gross Margin", "Operating Margin", "Profit Margin", "Revenue Growth", "Dividend Yield"]
r = 4
first_peer = r
for name, d in peers.items():
    put(ws, r, 1, name, font=BOLD if name.startswith("Dollarama") else NORM)
    for j, k in enumerate(want):
        v = d.get(k, {}).get("TTM", d.get(k, {}).get("Current"))
        fmt = PCT if j >= 4 else '0.00'
        put(ws, r, 2 + j, v, fmt=fmt, align=Alignment(horizontal="right"))
    r += 1
last_peer = r - 1
put(ws, r, 1, "Peer median (excl. Dollarama)", font=BOLD)
for j in range(len(want)):
    L = get_column_letter(2 + j)
    put(ws, r, 2 + j, f"=MEDIAN({L}{first_peer+1}:{L}{last_peer})", fmt=PCT if j >= 4 else '0.00', font=BOLD, align=Alignment(horizontal="right"))
r += 2
put(ws, r, 1, "Read: Dollarama has the group's best margins and among its fastest growth — and by far its highest multiples. The investment question is whether the quality justifies a P/E roughly double the peer median.", font=Font(italic=True, color="555555"))
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)

wb.save(OUT)
print("Saved", OUT)
json.dump({"forecast_rows": FORECAST_ROWS, "model_rows": MODEL_ROWS, "fy26": {"debt_incl_leases": 5396, "cash": cash26, "net_debt": 5396 - cash26, "shares": shares26}},
          open(os.path.join(DATA, "model_meta.json"), "w"), indent=1, default=str)
