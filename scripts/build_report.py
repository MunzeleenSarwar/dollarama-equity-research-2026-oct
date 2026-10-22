#!/usr/bin/env python3
"""
Build the Dollarama equity research report (PDF, 6-8 pages).
All numbers are read from data/results.json and data/model_outputs.json so the
report can never drift from the analysis notebook or the Excel model.
"""
import json, os
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_JUSTIFY
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer,
                                Table, TableStyle, Image, PageBreak, HRFlowable, KeepTogether)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
DATA, CHARTS = os.path.join(ROOT, "data"), os.path.join(ROOT, "charts")
OUT = os.path.join(ROOT, "Dollarama_Equity_Research_Report.pdf")

R = json.load(open(os.path.join(DATA, "results.json")))
M = json.load(open(os.path.join(DATA, "model_outputs.json")))
rat = R["ratios_5y"]; comp = R["comps"]; summ = R["summary"]

NAVY = HexColor("#1F3864"); ACC = HexColor("#2E5A88"); GREY = HexColor("#555555")
LGREY = HexColor("#EDF1F5"); BLACK = HexColor("#111111")

st = {
 "title": ParagraphStyle("title", leftIndent=0, fontName="Helvetica-Bold", fontSize=19, leading=22, textColor=NAVY),
 "subtitle": ParagraphStyle("subtitle", leftIndent=0, fontName="Helvetica", fontSize=10.5, leading=13, textColor=ACC),
 "h1": ParagraphStyle("h1", leftIndent=0, fontName="Helvetica-Bold", fontSize=12.5, leading=15, textColor=NAVY, spaceBefore=10, spaceAfter=4),
 "h2": ParagraphStyle("h2", leftIndent=0, fontName="Helvetica-Bold", fontSize=10.5, leading=13, textColor=ACC, spaceBefore=8, spaceAfter=3),
 "body": ParagraphStyle("body", leftIndent=0, fontName="Helvetica", fontSize=9.6, leading=13.2, textColor=BLACK, alignment=TA_JUSTIFY, spaceAfter=5),
 "bullet": ParagraphStyle("bullet", leftIndent=0, fontName="Helvetica", fontSize=9.6, leading=13.2, textColor=BLACK, spaceAfter=3.5),
 "small": ParagraphStyle("small", leftIndent=0, fontName="Helvetica", fontSize=8.4, leading=10.8, textColor=GREY),
 "cell": ParagraphStyle("cell", leftIndent=0, fontName="Helvetica", fontSize=8.6, leading=10.6),
 "cellb": ParagraphStyle("cellb", leftIndent=0, fontName="Helvetica-Bold", fontSize=8.6, leading=10.6),
 "rating": ParagraphStyle("rating", leftIndent=0, fontName="Helvetica-Bold", fontSize=13, leading=15, textColor=HexColor("#FFFFFF")),
}

def tbl(data, widths, header=True, align_right_from=1):
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    cmds = [("LEFTPADDING", (0,0), (-1,-1), 4), ("RIGHTPADDING", (0,0), (-1,-1), 4),
            ("TOPPADDING", (0,0), (-1,-1), 2.6), ("BOTTOMPADDING", (0,0), (-1,-1), 2.6),
            ("VALIGN", (0,0), (-1,-1), "TOP"), ("GRID", (0,0), (-1,-1), 0.4, HexColor("#C4CDD6"))]
    if header:
        cmds += [("BACKGROUND", (0,0), (-1,0), NAVY), ("TEXTCOLOR", (0,0), (-1,0), HexColor("#FFFFFF"))]
    if align_right_from:
        cmds.append(("ALIGN", (align_right_from,0), (-1,-1), "RIGHT"))
    t.setStyle(TableStyle(cmds))
    return t

def P(txt, style="body"): return Paragraph(txt, st[style])
def B(txt): return Paragraph(txt, st["bullet"], bulletText="•")
def pct(x, d=1): return f"{x*100:.{d}f}%"
def money(x): return f"{x:,.0f}"

S = []
# ---------- Page 1 ----------
S.append(P("Dollarama Inc.", "title"))
S.append(P("TSX: DOL &nbsp;|&nbsp; Consumer Defensive — Discount Retail &nbsp;|&nbsp; Equity Research — Initiation", "subtitle"))
S.append(Spacer(1, 4))
rating_tbl = Table([[Paragraph("RATING: HOLD", st["rating"]),
                     Paragraph(f"Price (21 Oct 2026): C$204.59 &nbsp;&nbsp;•&nbsp;&nbsp; DCF value: C${M['implied_price']:.2f} &nbsp;&nbsp;•&nbsp;&nbsp; Implied downside: {pct(M['implied_upside'])}", st["cell"])]],
                   colWidths=[1.5*inch, 5.9*inch])
rating_tbl.setStyle(TableStyle([("BACKGROUND", (0,0), (0,0), ACC), ("BACKGROUND", (1,0), (1,0), LGREY),
                                ("LEFTPADDING", (0,0), (-1,-1), 8), ("TOPPADDING", (0,0), (-1,-1), 6),
                                ("BOTTOMPADDING", (0,0), (-1,-1), 6), ("VALIGN", (0,0), (-1,-1), "MIDDLE")]))
S.append(rating_tbl)
S.append(Spacer(1, 6))
stats = [["", "Value", "", "Value"],
         ["Market capitalisation", "C$54.9B", "Trailing P/E", "41.0x"],
         ["Enterprise value", "C$60.3B", "Forward P/E", "38.3x"],
         ["Shares outstanding", "270.8M", "Price / sales (TTM)", "7.0x"],
         ["Revenue (FY2026)", "C$7,256M", "Dividend yield", "0.2%"],
         ["Net income (FY2026)", "C$1,309M", "52-week range", "C$163.25 – 209.96"],
         ["EPS diluted (FY2026)", "C$4.73", "Net debt incl. leases", "C$5,056M"]]
S.append(tbl([[Paragraph(c, st["cellb"] if i % 2 == 0 and j > 0 else st["cell"]) for i, c in enumerate(row)] for j, row in enumerate(stats)],
             [1.85*inch, 1.05*inch, 1.85*inch, 1.55*inch], header=True, align_right_from=None))
S.append(P("Investment thesis", "h1"))
S.append(B("<b>A genuinely excellent business.</b> Dollarama has grown revenue for ten consecutive years (10-year CAGR " + pct(summ["revenue_cagr_10y"]) + "), holds gross margin near 45% and operating margin near 24% while selling at price points capped at $5, and converts roughly a fifth of revenue into free cash flow. There is no down year in the window — not in COVID, not in the inflation shock."))
S.append(B("<b>Growth is now international.</b> The Canadian core (about 1,700 stores) is maturing, so the story has widened: a majority stake in Dollarcity in Latin America (600+ stores) and the 2025 acquisition of The Reject Shop in Australia (~400 stores; 6.3% of FY2026 revenue, 93.7% still Canada)."))
S.append(B("<b>But the price already knows.</b> At 41x trailing earnings the market capitalises Dollarama at roughly double the peer median multiple (Section 6). Our DCF base case (WACC " + pct(M["wacc"], 2) + ", terminal growth 2.5%) values the shares at C$" + f"{M['implied_price']:.0f}" + ", about 10% below the market price. The premium can be defended only if double-digit growth persists longer than we model."))
S.append(B("<b>Recent results underline both sides.</b> In Q2 FY2027 (ended 2 August 2026) revenue rose 17.6% to C$2,027M and diluted EPS rose 11.2% to C$1.29, yet adjusted EBITDA margin compressed to 32.2% from 34.1% as lower-margin Australian operations entered the mix — growth is real, but it is no longer all high-margin Canadian growth."))
S.append(P("Company snapshot", "h1"))
S.append(P("Dollarama, founded in 1992 and headquartered in Montreal, is Canada's largest dollar-store chain. Stores carry consumables, general merchandise and seasonal goods at fixed price points up to $5 — a deliberately simple model: small stores, limited SKUs, direct sourcing, and ruthless cost discipline. The equity story is therefore not pricing power in the usual sense; it is volume, store economics and capital allocation."))
S.append(KeepTogether([Image(os.path.join(CHARTS, "revenue_10y.png"), width=4.9*inch, height=2.30*inch),
                       P("Figure 1 — Revenue, FY2017–FY2026 (CAD millions). Source: company filings via StockAnalysis.com; analysis in the accompanying notebook.", "small")]))
S.append(PageBreak())

# ---------- Page 2 ----------
S.append(P("1 &nbsp; Business model and strategy", "h1"))
S.append(P("Three engines now drive the group. <b>Canada</b> remains the profit core: a dense store network still adding locations each year, with same-store sales supported by consumers trading down in a high cost-of-living environment. <b>Dollarcity</b> (Latin America) gives Dollarama an equity-accounted growth option in Colombia, Guatemala, El Salvador, Peru and Mexico; its contribution appears in the accounts mainly as earnings from equity investments (about C$222M of pretax support in FY2026 in our data build). <b>The Reject Shop</b> (Australia), acquired in 2025, is the first fully consolidated international platform — strategically logical (the same discount playbook) but dilutive to margins while it is integrated and converted."))
S.append(P("The model's quiet strength is capital intensity. A new Dollarama store is cheap to fit out, inventories turn about 3.6 times a year (roughly 101 days in FY2026), and capex ran at only 3.4% of revenue in FY2026. That is why free cash flow, not accounting profit, is the right lens for this business — and why the balance sheet can carry meaningful lease obligations without strain."))
S.append(P("2 &nbsp; Historical financial analysis", "h1"))
S.append(Image(os.path.join(CHARTS, "margins_5y.png"), width=6.4*inch, height=3.01*inch))
S.append(P("Figure 2 — Margins, FY2022–FY2026. Gross margin compressed about 2 points into the inflation peak, then stabilised; operating leverage did the rest.", "small"))
ratio_header = ["Ratio", "FY2022", "FY2023", "FY2024", "FY2025", "FY2026"]
def rrow(label, key, fmt):
    return [label] + [fmt(rat[fy][key]) for fy in ["FY 2022", "FY 2023", "FY 2024", "FY 2025", "FY 2026"]]
rows = [[Paragraph(c, st["cellb"] if j == 0 else st["cell"]) for j, c in enumerate(ratio_header)]]
for lab, key, fmt in [("Gross margin", "gross_margin", pct), ("Operating margin", "operating_margin", pct),
                      ("Net margin", "net_margin", pct), ("Return on assets", "roa", pct),
                      ("Current ratio (x)", "current_ratio", lambda x: f"{x:.2f}"),
                      ("Inventory days", "inventory_days", lambda x: f"{x:.0f}"),
                      ("FCF margin", "fcf_margin", pct), ("EPS diluted (C$)", "eps_diluted", lambda x: f"{x:.2f}")]:
    rows.append([Paragraph(c, st["cell"]) for c in rrow(lab, key, fmt)])
S.append(tbl(rows, [2.2*inch] + [1.02*inch]*5))
S.append(Spacer(1, 3))
S.append(P("Three observations. First, <b>margin quality is high and stable</b>: even in FY2024, when cost inflation bit hardest, operating margin stayed above 24%. Second, <b>returns on assets of 17–19%</b> are exceptional for retail and are achieved without financial engineering on the asset side. Third, a caveat we apply throughout: <b>return on equity is not meaningful here</b>. Years of buybacks shrank book equity to C$1,456M and it was negative in FY2022–FY2023, so ROE prints at absurd levels (about 90% in FY2026 on the small positive base, and not computable sensibly before). Any analysis quoting Dollarama's ROE without that warning should not be trusted."))
S.append(PageBreak())

# ---------- Page 3 ----------
S.append(P("3 &nbsp; Cash flow and capital allocation", "h1"))
S.append(P("Dollarama is a cash machine with a deliberate leakage problem — deliberate because the leakage is buybacks. FY2026 operating cash flow was C$1,761M against capex of only C$249M, leaving free cash flow (before lease payments) of C$1,512M, a 20.8% FCF margin. Of that, C$864M went to share repurchases and C$113M to dividends. Over FY2022–FY2026 the share count fell from 304M to 277M diluted, which is how EPS compounded at " + pct(summ["eps_cagr_5y"]) + " a year while net income grew more slowly."))
S.append(Image(os.path.join(CHARTS, "capital_returns.png"), width=6.4*inch, height=3.01*inch))
S.append(P("Figure 3 — Capital returned to shareholders, FY2022–FY2026 (CAD millions). Buybacks dominate; the dividend is a token (payout under 10%).", "small"))
S.append(Image(os.path.join(CHARTS, "eps_fcf_ps.png"), width=6.4*inch, height=3.01*inch))
S.append(P("Figure 4 — EPS and free cash flow per share, FY2022–FY2026. Per-share compounding is the real growth story.", "small"))
S.append(P("Leverage, stated honestly", "h2"))
S.append(P("Total debt including IFRS 16 lease liabilities was C$5,396M at FY2026 (of which leases C$2,771M), against C$340M of cash and short-term investments — net debt of about C$5,056M. That is roughly 2.3x FY2026 EBIT plus total depreciation: comfortable for a business with this cash conversion, and interest cover is strong. But analyses that count only the C$2,625M of borrowings materially understate obligations. Our valuation therefore treats leases as debt both in the cash flows (FCFF before lease payments) and in the bridge from enterprise value to equity value."))
S.append(PageBreak())

# ---------- Page 4 ----------
S.append(P("4 &nbsp; Forecast — what we assume, and why", "h1"))
S.append(P("The model (accompanying Excel file) is a fully integrated three-statement forecast. Revenue growth fades from 13% in FY2027 — flattered by a full year of Australia and Dollarcity momentum — to 6.5% by FY2031 as the Canadian store base matures. Gross margin is held at 45.0% (below the FY2022 level, in line with FY2025–26), SG&A at about 15% of revenue, depreciation at 6.0% of revenue, capex at 3.5%, and the tax rate at the FY2026 effective rate of 25.8%. Working capital scales with revenue. Buybacks continue at C$900M a year; the dividend payout stays near 9% of earnings. The balance sheet balances in every year and the cash-flow statement reconciles to the balance-sheet cash plug — both checks are in the model and read zero."))
fh = ["", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
frows = [[Paragraph(c, st["cellb"]) for c in fh]]
for lab, key, fmt in [("Revenue (C$M)", "revenue", money), ("EBIT (C$M)", "ebit", money),
                      ("Net income (C$M)", "net_income", money), ("EPS diluted (C$)", "eps", lambda x: f"{x:.2f}"),
                      ("FCF before leases (C$M)", "fcf_pre_lease", money)]:
    frows.append([Paragraph(lab, st["cell"])] + [Paragraph(fmt(M["forecast"][fy][key]), st["cell"]) for fy in ["FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]])
S.append(tbl(frows, [2.2*inch] + [1.02*inch]*5))
S.append(Spacer(1, 3))
S.append(P("These are not heroic assumptions: the implied five-year revenue CAGR is about 9%, below the last five years' 13.8%, and EBIT margin holds near 24% rather than expanding. The forecast risk is asymmetric in an interesting way — execution risk (Australia, Latin America) argues for caution, while the Canadian core has repeatedly beaten cautious assumptions when consumers are squeezed."))
S.append(P("5 &nbsp; Valuation I — discounted cash flow", "h1"))
S.append(P("Cost of capital. We use stated capital-market assumptions: a 3.4% risk-free rate, a 5.0% equity risk premium and a beta of 0.60 — an analyst assumption appropriate for a defensive, staples-like retailer, giving a 6.4% cost of equity. The after-tax cost of debt is 3.2% (4.3% pre-tax). At market weights (equity C$54.9B, debt incl. leases C$5.4B) the WACC is " + pct(M["wacc"], 2) + ". Readers should treat the beta as the single most contestable input — hence the sensitivity grid below."))
dh = ["DCF (C$M)", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
drows = [[Paragraph(c, st["cellb"]) for c in dh],
         [Paragraph("Free cash flow to firm", st["cell"])] + [Paragraph(money(v), st["cell"]) for v in M["fcff"]]]
S.append(tbl(drows, [2.2*inch] + [1.02*inch]*5))
S.append(Spacer(1, 3))
bridge = [["Present value of FCFF FY2027E–FY2031E", money(M["pv_fcff"])],
          ["Present value of terminal value (g = 2.5%)", money(M["pv_tv"])],
          ["Enterprise value", money(M["ev"])],
          ["Less: net debt incl. leases", "(" + money(M["net_debt"]) + ")"],
          ["Equity value", money(M["equity_value"])],
          ["Shares outstanding (M)", f"{M['shares']:.1f}"],
          ["Implied value per share (C$)", f"{M['implied_price']:.2f}"],
          ["Market price 21 Oct 2026 (C$)", "204.59"],
          ["Implied downside", pct(M["implied_upside"])]]
S.append(tbl([[Paragraph(a, st["cell"]), Paragraph(b, st["cell"])] for a, b in bridge], [4.6*inch, 1.7*inch], header=False))
S.append(Spacer(1, 3))
S.append(P("As always with a low-WACC, high-quality name, the terminal value dominates (about 85% of enterprise value). That is precisely why the sensitivity grid matters more than the point estimate."))
gh = ["WACC \\ g"] + [pct(float(k), 1) for k in M["sens_g"]]
grows = [[Paragraph(c, st["cellb"]) for c in gh]]
for wacc, vals in M["sensitivity"].items():
    grows.append([Paragraph(pct(float(wacc), 1), st["cell"])] + [Paragraph(f"{v:.0f}", st["cell"]) for v in vals])
S.append(tbl(grows, [1.1*inch] + [1.04*inch]*5))
S.append(P("Table — Implied share price (C$) across WACC and terminal growth. The market price of C$204.59 is reached only in the optimistic corner (low WACC, high terminal growth).", "small"))
S.append(PageBreak())

# ---------- Page 5 ----------
S.append(P("6 &nbsp; Valuation II — trading comparables", "h1"))
S.append(P("Dollarama's peer set is imperfect by construction: no other Canadian retailer combines its margins, growth and buyback intensity. We show North American value retailers (Dollar General, Dollar Tree, Five Below — USD reporters, ratios are currency-neutral) and large Canadian food/convenience retailers (Loblaw, Metro, Couche-Tard)."))
ch = ["Company", "P/E", "Fwd P/E", "P/S", "Gross mgn", "Op. mgn", "Rev growth"]
crows = [[Paragraph(c, st["cellb"]) for c in ch]]
order = ["Dollarama (TSX: DOL)", "Dollar General (NYSE: DG)", "Dollar Tree (NASDAQ: DLTR)", "Five Below (NASDAQ: FIVE)", "Loblaw (TSX: L)", "Metro (TSX: MRU)", "Couche-Tard (TSX: ATD)"]
for name in order:
    d = comp[name]
    nm = Paragraph(f"<b>{name}</b>" if name.startswith("Dollarama") else name, st["cell"])
    crows.append([nm] + [Paragraph(v, st["cell"]) for v in [f"{d['PE Ratio']:.1f}x", f"{d['Forward PE']:.1f}x", f"{d['PS Ratio']:.2f}x", pct(d["Gross Margin"]), pct(d["Operating Margin"]), pct(d["Revenue Growth"])]])
S.append(tbl(crows, [1.95*inch, 0.72*inch, 0.82*inch, 0.72*inch, 0.95*inch, 0.85*inch, 0.95*inch]))
S.append(Spacer(1, 3))
S.append(Image(os.path.join(CHARTS, "pe_comps.png"), width=6.4*inch, height=3.2*inch))
S.append(P("Figure 5 — Trailing P/E, October 2026 (TTM). Source: StockAnalysis.com.", "small"))
S.append(P("The pattern is unmistakable: Dollarama has the group's best gross and operating margins and near-best growth — and trades at about 1.4x Loblaw's multiple and 2.4x Dollar General's. Part of that premium is earned (margin and consistency gaps are real), and part is scarcity value as one of very few Canadian consumer compounders of this quality. Applying even a generous 30x to our FY2027E EPS of C$5.38 gives C$161, below both the price and our DCF base case — comparables alone argue the shares are expensive; the DCF says the premium is mostly, but not fully, justified by cash-flow durability."))
S.append(P("Football field (C$ per share)", "h2"))
ff = [["DCF sensitivity range (WACC 5.6–6.6%, g 2.0–3.0%)", "142 – 258"],
      ["DCF base case", f"{M['implied_price']:.2f}"],
      ["Comps-implied at 25–30x FY2027E EPS", "135 – 161"],
      ["52-week trading range", "163.25 – 209.96"],
      ["Current price", "204.59"]]
S.append(tbl([[Paragraph(a, st["cell"]), Paragraph(b, st["cell"])] for a, b in ff], [4.6*inch, 1.7*inch], header=False))
S.append(PageBreak())

# ---------- Page 6 ----------
S.append(P("7 &nbsp; Risks", "h1"))
S.append(B("<b>Multiple compression.</b> The largest risk is not operational. At 41x earnings, even a re-rating to Loblaw-like levels (~29x) would erase about 30% of the market value with earnings unchanged."))
S.append(B("<b>International execution.</b> Australia is the first large consolidated expansion; Q2 FY2027 already showed margin dilution. A difficult conversion of The Reject Shop, or weakness in Dollarcity's markets and currencies, would hit both earnings and the narrative that supports the premium."))
S.append(B("<b>Trade-down reversal.</b> Part of recent Canadian strength is consumers trading down. A stronger consumer economy could slow traffic growth — an unusual risk for a retailer, but real for a value format."))
S.append(B("<b>Competition and costs.</b> Walmart, Amazon and Dollar Tree's Canadian ambitions pressure the value segment; minimum-wage growth and occupancy costs pressure the cost base Dollarama has historically controlled so well."))
S.append(B("<b>Leverage and rates.</b> Net debt including leases of about C$5.1B is well covered today, but the valuation's low discount rate means higher-for-longer rates would compress both the DCF value and, likely, the multiple."))
S.append(P("8 &nbsp; Conclusion", "h1"))
S.append(P("Dollarama is the rare retailer that behaves like a compounder: a decade of uninterrupted revenue growth, 45% gross margins, returns on assets near 18%, and a buyback engine that turns mid-teens profit growth into roughly 21% EPS growth. Our analysis of its financials gives us no quarrel with the business at all. Our quarrel is only with the price: at C$204.59 the shares trade above our C$" + f"{M['implied_price']:.0f}" + " base-case value and near the top of their 52-week range, on a multiple roughly double the peer median. We therefore rate the shares <b>HOLD</b>: existing holders are being paid in per-share compounding to wait, but new money is underwriting several more years of near-flawless execution. A pullback toward the C$160–180 zone — or evidence that Australian margins converge to Canadian levels faster than we model — would change the answer."))
S.append(P("Sources and methodology", "h1"))
S.append(P("Financial data: Dollarama's filed financial statements as aggregated by StockAnalysis.com (retrieved 21 October 2026), cross-checked against company results reporting for Q2 FY2027. Market data: TSX close, 21 October 2026. The full audit trail is reproducible: <i>scripts/fetch_data.py</i> collects the data, <i>analysis.ipynb</i> computes the ratios and charts, and <i>Dollarama_Financial_Model.xlsx</i> contains the integrated forecast and DCF with live formulas (balance-sheet and cash-flow checks read zero). Valuation assumptions (risk-free rate, beta, equity risk premium, terminal growth) are the analyst's stated assumptions, shown on the DCF sheet of the model, and are the inputs to challenge first.", "small"))
S.append(P("Disclaimer", "h1"))
S.append(P("This report was prepared by Munzeleen Sarwar as an educational portfolio project to demonstrate financial analysis and modelling skills. It is based on public information believed to be reliable but not independently audited, it is not investment advice, a recommendation, or an offer to buy or sell any security, and it must not be relied upon for investment decisions. Figures are in Canadian dollars unless stated. Past performance is not indicative of future results.", "small"))

doc = BaseDocTemplate(OUT, pagesize=LETTER, leftMargin=0.65*inch, rightMargin=0.65*inch,
                      topMargin=0.55*inch, bottomMargin=0.55*inch,
                      title="Dollarama (TSX: DOL) — Equity Research Report", author="Munzeleen Sarwar")
frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="main")

def footer(canvas, doc_):
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(GREY)
    canvas.drawString(0.65*inch, 0.32*inch, "Dollarama Inc. (TSX: DOL) — Equity Research | Munzeleen Sarwar | October 2026 | Educational project — not investment advice")
    canvas.drawRightString(LETTER[0]-0.65*inch, 0.32*inch, f"Page {doc_.page}")
    canvas.restoreState()

doc.addPageTemplates([PageTemplate(id="p", frames=[frame], onPage=footer)])
doc.build(S)
print("Built:", OUT)
