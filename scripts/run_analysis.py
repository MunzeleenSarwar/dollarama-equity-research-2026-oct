#!/usr/bin/env python3
"""
Ratio analysis and charts for Dollarama (TSX: DOL).

Reads data/statements.json and data/revenue_history.csv (from fetch_data.py),
computes a 10-year ratio history where the data supports it and a 5-year
detailed ratio set from the full statements, and writes charts to ../charts/
plus a results.json that the Excel model and PDF report builders consume,
so every document quotes the same numbers.
"""
import json, os, re
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "..")
DATA, CHARTS = os.path.join(ROOT, "data"), os.path.join(ROOT, "charts")
os.makedirs(CHARTS, exist_ok=True)
plt.rcParams.update({"font.size": 9, "axes.edgecolor": "#888888",
                     "axes.labelcolor": "#222222", "figure.dpi": 150})
NAVY, TEAL, GREY = "#1F3864", "#2E86AB", "#9AA5B1"

stmts = json.load(open(os.path.join(DATA, "statements.json")))
inc, bal, cf = stmts["income"], stmts["balance"], stmts["cashflow"]
FYS = ["FY 2022", "FY 2023", "FY 2024", "FY 2025", "FY 2026"]


def g(d, k, p):
    v = d.get(k, {}).get(p)
    return None if v is None else float(v)


# ---------- 10-year revenue history ----------
rh = pd.read_csv(os.path.join(DATA, "revenue_history.csv"))


def to_num(s):
    s = str(s).replace(",", "").strip()
    m = re.match(r"([\d.]+)\s*([BM])", s)
    if m:
        return float(m.group(1)) * (1000 if m.group(2) == "B" else 1)
    try:
        return float(s)
    except Exception:
        return None


rev_hist = []
for _, row in rh.iterrows():
    fy_end = str(row.iloc[0])
    year = int(re.search(r"(\d{4})", fy_end).group(1))
    gv = to_num(str(row.iloc[3]).replace("%", "")) if pd.notna(row.iloc[3]) else None
    rev_hist.append({"fy_label": f"FY{year}", "fy_end": fy_end, "revenue": to_num(row.iloc[1]),
                     "growth": (gv / 100) if gv is not None else None})
rev_hist = sorted(rev_hist, key=lambda x: x["fy_label"])
rev10 = [r for r in rev_hist if r["fy_label"] >= "FY2017"][-10:]
print("Revenue history:", [(r["fy_label"], r["revenue"]) for r in rev10])

# ---------- 5-year detailed ratios ----------
rows = {}
for fy in FYS:
    rev = g(inc, "Revenue", fy)
    rows[fy] = {
        "revenue": rev,
        "gross_margin": g(inc, "Gross Profit", fy) / rev,
        "operating_margin": g(inc, "Operating Income", fy) / rev,
        "net_margin": g(inc, "Net Income", fy) / rev,
        "ebitda_sa_margin": g(inc, "EBITDA", fy) / rev,
        "effective_tax": g(inc, "Effective Tax Rate", fy),
        "roe": g(inc, "Net Income", fy) / g(bal, "Shareholders' Equity", fy),
        "roa": g(inc, "Net Income", fy) / g(bal, "Total Assets", fy),
        "current_ratio": g(bal, "Total Current Assets", fy) / g(bal, "Total Current Liabilities", fy),
        "inventory_days": g(bal, "Inventory", fy) / g(inc, "Cost of Revenue", fy) * 365,
        "net_debt_incl_leases": (g(bal, "Total Debt", fy) or 0) - (g(bal, "Cash & Short-Term Investments", fy) or 0),
        "fcf": g(cf, "Free Cash Flow", fy),
        "fcf_margin": g(cf, "Free Cash Flow", fy) / rev,
        "fcf_after_leases": g(cf, "Free Cash Flow", fy) + g(cf, "Lease Payments", fy),
        "capex": -g(cf, "Capital Expenditures", fy),
        "buybacks": -g(cf, "Repurchase of Common Stock", fy),
        "dividends": -g(cf, "Common Dividends Paid", fy),
        "eps_diluted": g(inc, "EPS (Diluted)", fy),
        "shares_diluted_m": g(inc, "Shares Outstanding (Diluted)", fy),
        "ocf": g(cf, "Operating Cash Flow", fy),
    }
ratios = pd.DataFrame(rows).T
print(ratios.round(3).to_string())

cagr = lambda a, b, n: (b / a) ** (1 / n) - 1
summary = {
    "revenue_cagr_10y": cagr(rev10[0]["revenue"], rev10[-1]["revenue"], 9),
    "revenue_cagr_5y": cagr(g(inc, "Revenue", "FY 2022"), g(inc, "Revenue", "FY 2026"), 4),
    "eps_cagr_5y": cagr(g(inc, "EPS (Diluted)", "FY 2022"), g(inc, "EPS (Diluted)", "FY 2026"), 4),
    "price_oct21_2026": 204.59,
    "market_cap_ttm": 54906.0,
    "ev_ttm": 60329.0,
}
print(summary)

# ---------- charts ----------
yrs = [r["fy_label"] for r in rev10]
revs = [r["revenue"] for r in rev10]
fig, ax = plt.subplots(figsize=(6.6, 3.1))
ax.bar(yrs, revs, color=NAVY, alpha=.88)
for x, v in zip(yrs, revs):
    ax.annotate(f"{v:,.0f}", (x, v), ha="center", va="bottom", fontsize=7.5)
ax.set_ylabel("Revenue (CAD millions)")
ax.set_title("Dollarama revenue — FY2017 to FY2026")
ax.spines[["top", "right"]].set_visible(False)
plt.tight_layout(); plt.savefig(os.path.join(CHARTS, "revenue_10y.png")); plt.close()

fy5 = [f.replace("FY ", "FY") for f in FYS]
fig, ax = plt.subplots(figsize=(6.6, 3.1))
for col, lab, c in [("gross_margin", "Gross margin", NAVY), ("operating_margin", "Operating margin", TEAL), ("net_margin", "Net margin", GREY)]:
    ax.plot(fy5, [rows[f][col] * 100 for f in FYS], marker="o", label=lab, color=c)
ax.set_ylabel("% of revenue"); ax.set_title("Profitability margins — FY2022 to FY2026")
ax.legend(frameon=False, fontsize=8); ax.spines[["top", "right"]].set_visible(False)
plt.tight_layout(); plt.savefig(os.path.join(CHARTS, "margins_5y.png")); plt.close()

fig, ax = plt.subplots(figsize=(6.6, 3.1))
ax.bar(fy5, [rows[f]["eps_diluted"] for f in FYS], color=TEAL, alpha=.9, label="EPS diluted (CAD)")
ax2 = ax.twinx()
ax2.plot(fy5, [rows[f]["fcf"] / rows[f]["shares_diluted_m"] for f in FYS], marker="s", color=NAVY, label="FCF per share (CAD)")
ax.set_ylabel("CAD per share"); ax2.set_ylabel("CAD per share")
ax.set_title("EPS and free cash flow per share — FY2022 to FY2026")
ax.spines[["top"]].set_visible(False); ax2.spines[["top"]].set_visible(False)
plt.tight_layout(); plt.savefig(os.path.join(CHARTS, "eps_fcf_ps.png")); plt.close()

fig, ax = plt.subplots(figsize=(6.6, 3.1))
ax.bar(fy5, [rows[f]["buybacks"] for f in FYS], color=NAVY, alpha=.88, label="Share buybacks")
ax.bar(fy5, [rows[f]["dividends"] for f in FYS], bottom=[rows[f]["buybacks"] for f in FYS], color=TEAL, alpha=.9, label="Dividends paid")
ax.set_ylabel("CAD millions"); ax.set_title("Capital returned to shareholders — FY2022 to FY2026")
ax.legend(frameon=False, fontsize=8); ax.spines[["top", "right"]].set_visible(False)
plt.tight_layout(); plt.savefig(os.path.join(CHARTS, "capital_returns.png")); plt.close()

# ---------- peers ----------
peers = json.load(open(os.path.join(DATA, "peers.json")))
want = ["PE Ratio", "Forward PE", "PS Ratio", "P/FCF Ratio", "Gross Margin", "Operating Margin", "Profit Margin", "Revenue Growth", "Dividend Yield"]
comp = {}
for name, d in peers.items():
    comp[name] = {k: (d.get(k, {}).get("TTM", d.get(k, {}).get("Current"))) for k in want}
comp_df = pd.DataFrame(comp).T
print(comp_df.round(3).to_string())
comp_df.to_csv(os.path.join(DATA, "comps_snapshot.csv"))

fig, ax = plt.subplots(figsize=(6.6, 3.3))
names = [n.split(" (")[0] for n in comp_df.index]
ax.barh(names, comp_df["PE Ratio"], color=[NAVY if n.startswith("Dollarama") else GREY for n in comp_df.index])
ax.set_xlabel("Trailing P/E (x), October 2026")
ax.set_title("Trailing P/E — Dollarama vs comparable retailers")
ax.spines[["top", "right"]].set_visible(False)
plt.tight_layout(); plt.savefig(os.path.join(CHARTS, "pe_comps.png")); plt.close()

results = {"summary": summary, "ratios_5y": {k: {c: (None if pd.isna(v) else float(v)) for c, v in row.items()} for k, row in rows.items()},
           "revenue_10y": rev10, "comps": {k: {kk: (None if vv is None or (isinstance(vv, float) and pd.isna(vv)) else float(vv)) for kk, vv in v.items()} for k, v in comp.items()}}
json.dump(results, open(os.path.join(DATA, "results.json"), "w"), indent=1)
print("charts + results.json written")
