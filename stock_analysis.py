#!/usr/bin/env python3
"""
Stock Analysis Tool
Usage: python stock_analysis.py TICKER [TICKER2 TICKER3 ...]
       python stock_analysis.py AAPL MSFT GOOG   # compare up to 5 tickers
"""

import sys
import os
from datetime import datetime, date

import yfinance as yf
import pandas as pd


# ── helpers ──────────────────────────────────────────────────────────────────

def _fmt(value, fmt=".2f", fallback="N/A"):
    """Format a numeric value, returning fallback on None/NaN."""
    try:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return fallback
        return format(float(value), fmt)
    except (TypeError, ValueError):
        return fallback


def _fmt_large(value, fallback="N/A"):
    """Human-readable large numbers (billions / millions)."""
    try:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return fallback
        v = float(value)
        if abs(v) >= 1e12:
            return f"${v/1e12:.2f}T"
        if abs(v) >= 1e9:
            return f"${v/1e9:.2f}B"
        if abs(v) >= 1e6:
            return f"${v/1e6:.2f}M"
        return f"${v:,.0f}"
    except (TypeError, ValueError):
        return fallback


def _calc_rsi(close: pd.Series, period: int = 14) -> float:
    """Wilder's RSI over `period` days using the closing price series."""
    delta = close.diff().dropna()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1 / period, min_periods=period, adjust=False).mean()
    avg_loss = loss.ewm(alpha=1 / period, min_periods=period, adjust=False).mean()
    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))
    return float(rsi.iloc[-1]) if not rsi.empty else float("nan")


def _signal(price, ma50, ma200, rsi):
    """Simple BUY / HOLD / SELL signal based on MAs and RSI."""
    if any(v is None or (isinstance(v, float) and pd.isna(v))
           for v in [price, ma50, ma200, rsi]):
        return "INSUFFICIENT DATA"

    bullish_ma  = price > ma50 > ma200
    bearish_ma  = price < ma50 < ma200
    overbought  = rsi > 70
    oversold    = rsi < 30

    if bullish_ma and not overbought:
        return "BUY"
    if bearish_ma and not oversold:
        return "SELL"
    if oversold:
        return "BUY (oversold)"
    if overbought:
        return "SELL (overbought)"
    return "HOLD"


# ── single-ticker analysis ────────────────────────────────────────────────────

def analyze_ticker(symbol: str) -> dict | None:
    """
    Fetch and compute all metrics for one ticker.
    Returns a dict of results, or None if the ticker is invalid.
    """
    symbol = symbol.upper().strip()
    try:
        tk = yf.Ticker(symbol)
        info = tk.info
    except Exception as e:
        print(f"[ERROR] Could not fetch data for {symbol}: {e}")
        return None

    # Validate – yfinance returns a mostly-empty dict for unknown tickers
    if not info or info.get("regularMarketPrice") is None and info.get("currentPrice") is None:
        print(f"[ERROR] No data found for ticker '{symbol}'. Check the symbol and try again.")
        return None

    # ── price data ───────────────────────────────────────────────────────────
    price     = info.get("currentPrice") or info.get("regularMarketPrice")
    week52hi  = info.get("fiftyTwoWeekHigh")
    week52lo  = info.get("fiftyTwoWeekLow")

    # ── fundamentals ─────────────────────────────────────────────────────────
    pe_ratio  = info.get("trailingPE") or info.get("forwardPE")
    eps       = info.get("trailingEps")
    mkt_cap   = info.get("marketCap")
    div_yield = info.get("dividendYield")

    # ── moving averages ───────────────────────────────────────────────────────
    ma50  = info.get("fiftyDayAverage")
    ma200 = info.get("twoHundredDayAverage")

    # ── RSI: need ~3 months of daily history ─────────────────────────────────
    hist = tk.history(period="3mo")
    if hist.empty or len(hist) < 15:
        rsi_val = float("nan")
    else:
        rsi_val = _calc_rsi(hist["Close"])

    # ── signal ───────────────────────────────────────────────────────────────
    sig = _signal(price, ma50, ma200, rsi_val)

    return {
        "symbol":       symbol,
        "name":         info.get("longName") or info.get("shortName", symbol),
        "price":        price,
        "week52hi":     week52hi,
        "week52lo":     week52lo,
        "pe_ratio":     pe_ratio,
        "eps":          eps,
        "market_cap":   mkt_cap,
        "div_yield":    div_yield,
        "ma50":         ma50,
        "ma200":        ma200,
        "rsi":          rsi_val,
        "signal":       sig,
        "as_of":        date.today().isoformat(),
    }


# ── report formatting ─────────────────────────────────────────────────────────

def _build_single_report(d: dict) -> str:
    div = (f"{d['div_yield']*100:.2f}%" if d["div_yield"] and not pd.isna(d["div_yield"])
           else "N/A")
    lines = [
        "=" * 58,
        f"  STOCK ANALYSIS REPORT — {d['symbol']}",
        f"  {d['name']}",
        f"  Generated: {d['as_of']}",
        "=" * 58,
        "",
        "── Price ───────────────────────────────────────────────",
        f"  Current Price   : ${_fmt(d['price'])}",
        f"  52-Week High    : ${_fmt(d['week52hi'])}",
        f"  52-Week Low     : ${_fmt(d['week52lo'])}",
        "",
        "── Fundamentals ────────────────────────────────────────",
        f"  P/E Ratio       : {_fmt(d['pe_ratio'])}",
        f"  EPS             : ${_fmt(d['eps'])}",
        f"  Market Cap      : {_fmt_large(d['market_cap'])}",
        f"  Dividend Yield  : {div}",
        "",
        "── Moving Averages ─────────────────────────────────────",
        f"  50-Day MA       : ${_fmt(d['ma50'])}",
        f"  200-Day MA      : ${_fmt(d['ma200'])}",
        "",
        "── Momentum ────────────────────────────────────────────",
        f"  RSI (14-day)    : {_fmt(d['rsi'])}",
        "",
        "── Signal ──────────────────────────────────────────────",
        f"  Recommendation  : {d['signal']}",
        "",
        "  Signal logic:",
        "    BUY  — price > MA50 > MA200 and RSI ≤ 70",
        "    SELL — price < MA50 < MA200 and RSI ≥ 30",
        "    BUY (oversold)   — RSI < 30",
        "    SELL (overbought)— RSI > 70",
        "    HOLD — none of the above",
        "=" * 58,
    ]
    return "\n".join(lines)


def print_single_report(d: dict):
    print(_build_single_report(d))


def save_single_report(d: dict):
    filename = f"{d['symbol']}_report.txt"
    with open(filename, "w") as f:
        f.write(_build_single_report(d))
    print(f"\nReport saved → {filename}")
    return filename


# ── multi-ticker comparison ───────────────────────────────────────────────────

def compare_tickers(symbols: list[str]):
    """Fetch up to 5 tickers and print a side-by-side comparison table."""
    if len(symbols) > 5:
        print("[WARNING] Comparing only the first 5 tickers.")
        symbols = symbols[:5]

    results = []
    for sym in symbols:
        print(f"Fetching {sym}…", end=" ", flush=True)
        d = analyze_ticker(sym)
        if d:
            results.append(d)
            print("done")
        else:
            print("skipped (invalid)")

    if not results:
        print("[ERROR] No valid tickers to compare.")
        return

    # Column widths
    col = 16
    header_row = f"{'Metric':<22}" + "".join(f"{r['symbol']:>{col}}" for r in results)
    sep = "-" * (22 + col * len(results))

    def row(label, vals):
        return f"{label:<22}" + "".join(f"{v:>{col}}" for v in vals)

    div_vals = []
    for r in results:
        dy = r["div_yield"]
        div_vals.append(f"{dy*100:.2f}%" if dy and not pd.isna(dy) else "N/A")

    table_lines = [
        "",
        "=" * (22 + col * len(results)),
        f"  MULTI-TICKER COMPARISON  ({date.today().isoformat()})",
        "=" * (22 + col * len(results)),
        header_row,
        sep,
        row("Price ($)",          [f"${_fmt(r['price'])}"    for r in results]),
        row("52W High ($)",       [f"${_fmt(r['week52hi'])}" for r in results]),
        row("52W Low ($)",        [f"${_fmt(r['week52lo'])}" for r in results]),
        sep,
        row("P/E Ratio",          [_fmt(r["pe_ratio"])       for r in results]),
        row("EPS ($)",            [f"${_fmt(r['eps'])}"      for r in results]),
        row("Market Cap",         [_fmt_large(r["market_cap"]) for r in results]),
        row("Dividend Yield",     div_vals),
        sep,
        row("MA 50-Day ($)",      [f"${_fmt(r['ma50'])}"     for r in results]),
        row("MA 200-Day ($)",     [f"${_fmt(r['ma200'])}"    for r in results]),
        row("RSI (14-day)",       [_fmt(r["rsi"])            for r in results]),
        sep,
        row("Signal",             [r["signal"]               for r in results]),
        "=" * (22 + col * len(results)),
        "",
    ]

    report_text = "\n".join(table_lines)
    print(report_text)

    # Save combined comparison report
    fname = "comparison_report.txt"
    with open(fname, "w") as f:
        f.write(report_text)
    print(f"Comparison report saved → {fname}")


# ── CLI entrypoint ────────────────────────────────────────────────────────────

def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]

    if not args:
        print(__doc__)
        print("Examples:")
        print("  python stock_analysis.py AAPL")
        print("  python stock_analysis.py AAPL MSFT GOOG AMZN TSLA")
        sys.exit(1)

    if len(args) == 1:
        symbol = args[0]
        print(f"Analyzing {symbol.upper()}…\n")
        data = analyze_ticker(symbol)
        if data is None:
            sys.exit(1)
        print_single_report(data)
        save_single_report(data)
    else:
        compare_tickers(args)


if __name__ == "__main__":
    main()
