#!/usr/bin/env python3
"""
Fetch Dollarama (TSX: DOL) financial data from StockAnalysis.com.

Source: StockAnalysis.com quote pages (publicly available summary financials,
which trace back to Dollarama's filed financial statements).
All figures in CAD millions unless noted. Fiscal years end late Jan / early Feb
and are labelled by the calendar year they end in (FY2026 ended 1 Feb 2026).

Outputs (into ../data/):
  - statements.json      full income statement / balance sheet / cash flow, FY2022-FY2026 + TTM
  - revenue_history.csv  revenue history back to FY2009
  - peers.json           current valuation & margin snapshot for comparable companies
  - snapshot.json        current price / market snapshot for DOL
"""
import requests, pandas as pd, io, json, os

H = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"}
HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "..", "data")
os.makedirs(DATA, exist_ok=True)
BASE = "https://stockanalysis.com/quote/tsx/DOL/"


def get_tables(url):
    r = requests.get(url, headers=H, timeout=30)
    if r.status_code != 200:
        return None
    try:
        return pd.read_html(io.StringIO(r.text))
    except Exception:
        return None


def period_name(c):
    return str(c[0]) if isinstance(c, tuple) else str(c)


def clean_val(v):
    if v is None:
        return None
    s = str(v).strip()
    if s in ("", "-", "nan"):
        return None
    pct = s.endswith("%")
    s = s.replace(",", "").replace("%", "")
    try:
        x = float(s)
    except Exception:
        return None
    return x / 100 if pct else x


def parse_tables(tabs):
    out = {}
    if not tabs:
        return out
    for t in tabs:
        if t.shape[1] < 3:
            continue
        periods = [period_name(c) for c in t.columns][1:]
        for _, row in t.iterrows():
            label = str(row.iloc[0]).strip()
            if not label or label.lower().startswith("period ending") or label == "nan":
                continue
            if label not in out:
                out[label] = {p: clean_val(v) for p, v in zip(periods, row.iloc[1:])}
    return out


def main():
    stmts = {}
    for name, path in [("income", "financials/income-statement/"),
                       ("balance", "financials/balance-sheet/"),
                       ("cashflow", "financials/cash-flow-statement/")]:
        stmts[name] = parse_tables(get_tables(BASE + path))
        print(name, len(stmts[name]), "labels")
    json.dump(stmts, open(os.path.join(DATA, "statements.json"), "w"), indent=1)

    # Revenue history (long-run)
    tabs = get_tables(BASE + "revenue/")
    hist = None
    for t in tabs or []:
        cols = [str(c).lower() for c in t.columns]
        if any("fiscal year" in c for c in cols) and t.shape[0] > 8:
            hist = t
            break
    if hist is not None:
        hist.to_csv(os.path.join(DATA, "revenue_history.csv"), index=False)
        print("revenue history rows:", hist.shape[0])

    # Peers snapshot
    peers = {
        "Dollarama (TSX: DOL)": BASE + "financials/",
        "Dollar General (NYSE: DG)": "https://stockanalysis.com/stocks/dg/financials/",
        "Dollar Tree (NASDAQ: DLTR)": "https://stockanalysis.com/stocks/dltr/financials/",
        "Five Below (NASDAQ: FIVE)": "https://stockanalysis.com/stocks/five/financials/",
        "Loblaw (TSX: L)": "https://stockanalysis.com/quote/tsx/l/financials/",
        "Metro (TSX: MRU)": "https://stockanalysis.com/quote/tsx/mru/financials/",
        "Couche-Tard (TSX: ATD)": "https://stockanalysis.com/quote/tsx/atd/financials/",
    }
    res = {name: parse_tables(get_tables(url)) for name, url in peers.items()}
    json.dump(res, open(os.path.join(DATA, "peers.json"), "w"), indent=1)
    print("peers fetched:", {k: len(v) for k, v in res.items()})

    # Statistics page (beta, price snapshot)
    stats = parse_tables(get_tables(BASE + "statistics/"))
    json.dump(stats, open(os.path.join(DATA, "statistics.json"), "w"), indent=1)
    print("statistics labels:", len(stats))


if __name__ == "__main__":
    main()
