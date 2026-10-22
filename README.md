# Dollarama Inc. (TSX: DOL) — Equity Research, Financial Model & Valuation

**Author:** Munzeleen Sarwar · **Date:** October 2026 · **Rating: HOLD** · Price C$204.59 (21 Oct 2026) vs DCF value **C$183.49**

An end-to-end equity research project on Dollarama, Canada's largest dollar-store chain:
ratio analysis on ten years of financials, a fully integrated three-statement Excel
model with a DCF valuation and trading comparables, and a six-page research report.

> Educational portfolio project — not investment advice.

## The story in one paragraph

Dollarama has grown revenue for ten consecutive years (10-year CAGR ≈ 10.5%), holds
gross margin near 45% and operating margin near 24% while selling at price points
capped at $5, and converts about a fifth of revenue into free cash flow. EPS compounded
at ~21% a year (FY2022–FY2026) thanks to buybacks that cut the share count from 304M to
277M. The business is excellent — the price is the problem: at ~41x trailing earnings
(roughly double the peer median), the shares trade above our DCF base case of C$183.49.

## Key findings

| Metric | Value |
|---|---|
| Revenue FY2026 | C$7,256M (+13.1%) |
| Net income FY2026 | C$1,309M |
| EPS diluted FY2026 | C$4.73 |
| Gross / operating margin FY2026 | 45.1% / 24.1% |
| Free cash flow FY2026 (pre-lease) | C$1,512M |
| Net debt incl. leases | C$5,056M |
| DCF implied value | **C$183.49** (WACC 6.11%, terminal growth 2.5%) |
| DCF sensitivity range | C$142 – C$258 |

## What's in this repo

| File | What it is |
|---|---|
| `Dollarama_Equity_Research_Report.pdf` | Six-page equity research report (thesis, financial analysis, forecast, DCF, comps, risks) |
| `Dollarama_Financial_Model.xlsx` | Integrated model: historicals & ratios, FY2027E–FY2031E three-statement forecast with live formulas, DCF with WACC build and sensitivity grid, trading comps. Balance-sheet and cash-flow checks read zero |
| `analysis.ipynb` | Executed Python notebook: 10-year ratio analysis and charts (all outputs embedded) |
| `charts/` | Chart images used in the report |
| `data/` | Source data and computed results (`statements.json`, `revenue_history.csv`, `peers.json`, `results.json`, `model_outputs.json`) |
| `scripts/` | The code that builds everything (see below) |

## Reproduce it

```bash
pip install pandas matplotlib openpyxl reportlab nbformat nbclient ipykernel requests lxml
python scripts/fetch_data.py      # collect financials, history and peer data -> data/
python scripts/run_analysis.py    # ratios + charts -> charts/, data/results.json
python scripts/build_model.py     # Excel model (formulas, not hard-coded values)
python scripts/build_notebook.py  # generate + execute analysis.ipynb
python scripts/build_report.py    # six-page PDF report
```

## Method notes (the honest bits)

- **Leases (IFRS 16)** are treated as financing throughout: free cash flow to firm is
  before lease payments, and total debt in the valuation bridge *includes* lease
  liabilities (C$2,771M of the C$5,396M total at FY2026).
- **ROE is deliberately excluded.** Buybacks shrank book equity so far that it was
  negative in FY2022–FY2023 — ROE is arithmetic noise here. ROA (17–19%) and margins
  carry the analysis.
- **WACC inputs are stated assumptions**, not fetched data: risk-free rate 3.4%,
  beta 0.60 (defensive retailer), equity risk premium 5.0%. They are labelled as
  assumptions in the model — challenge those first.
- **Data source:** Dollarama's filed financial statements as aggregated by
  StockAnalysis.com (retrieved 21 October 2026), cross-checked against Q2 FY2027
  results reporting. Verify against SEDAR+ filings before any investment use.

## Why Dollarama (and not a bank)?

A standard three-statement + DCF framework fits a retailer cleanly. Canadian banks
need specialist dividend-discount / residual-income models, and Canadian Tire
consolidates a bank, a REIT and financial services, which would have buried a first
model in consolidation noise. The same template in `scripts/` can be re-pointed at
another company by changing the data inputs.
