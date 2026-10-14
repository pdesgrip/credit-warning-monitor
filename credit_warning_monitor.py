#!/usr/bin/env python3
"""
credit_warning_monitor.py
-------------------------
Early-warning monitor for credit-cycle stress, built around the signals that
historically lead equity bear markets rather than confirm them:

  1. Bank lending standards  -- Fed SLOOS net % tightening (FRED, quarterly)
  2. High-yield primary market -- monthly HY issuance (SIFMA) vs its run-rate,
     with HY OAS (FRED) as the "do spreads still look fine?" context
  3. Credit cross-checks -- CCC & lower OAS (weakest credits crack first) and
     the HYG/LQD ratio (separates credit-driven from rates-driven weakness)

Each leg is scored GREEN / AMBER / RED; the composite is the worst leg.
Outputs a terminal summary, and optionally a self-contained HTML dashboard
and a JSON snapshot.

Usage:
    python credit_warning_monitor.py --sifma-file US-Corporate-Bonds-Statistics-SIFMA.xlsx
    python credit_warning_monitor.py --sifma-file X.xlsx --html report.html --open
    python credit_warning_monitor.py --hy-csv my_hy.csv --json latest.json

No API keys needed: FRED's public CSV endpoint, Stooq (Yahoo fallback) for ETF
prices, and a SIFMA workbook downloaded by hand (their download is form-gated).
"""

import argparse
import io
import json
import sys
from datetime import date, datetime

import pandas as pd
import requests

# ----------------------------------------------------------------------------
# Config: series + thresholds. Tune these; they are starting points, not gospel.
# ----------------------------------------------------------------------------
FRED_CSV = "https://fred.stlouisfed.org/graph/fredgraph.csv?id={sid}"

SLOOS_SERIES = {
    # net % of banks tightening standards (quarterly, positive = tightening)
    "DRTSCILM": "C&I loans, large/mid firms",
    "DRTSCIS":  "C&I loans, small firms",
    "DRTSCLCC": "Consumer credit cards",
}
HY_SPREAD_SERIES = "BAMLH0A0HYM2"   # ICE BofA US HY OAS, daily, in %
CCC_SPREAD_SERIES = "BAMLH0A3HYC"    # ICE BofA CCC & lower OAS, daily, in %

# CCC stress thresholds (1-month change, bp)
CCC_AMBER_BP = 75
CCC_RED_BP = 150
DIFF_AMBER_BP = 50    # CCC minus headline HY, 1m change
DIFF_RED_BP = 100
# HYG/LQD ratio 1m change; capped at AMBER (noisy, needs confirmation)
RATIO_AMBER = -0.015

# SLOOS thresholds (net % tightening). History: >~20 is stress, >~40 is recession-adjacent.
SLOOS_AMBER = 20
SLOOS_RED = 40
SLOOS_RISE_QTRS = 2          # consecutive quarterly increases = momentum flag

# HY issuance thresholds: trailing 3m volume vs trailing 12m avg monthly volume
ISS_AMBER_RATIO = 0.70       # 3m run-rate < 70% of 12m average
ISS_RED_RATIO = 0.50
ISS_MIN_HISTORY_MONTHS = 15

# HY spread context: "spreads still look fine" = below this percentile of last 3y
SPREAD_BENIGN_PCTL = 0.50

SIFMA_PAGE = "https://www.sifma.org/research/statistics/us-corporate-bonds-statistics"


# ----------------------------------------------------------------------------
# Data fetchers
# ----------------------------------------------------------------------------
def fred(sid: str) -> pd.Series:
    r = requests.get(FRED_CSV.format(sid=sid), timeout=30)
    r.raise_for_status()
    df = pd.read_csv(io.StringIO(r.text))
    df.columns = ["date", "value"]
    df["date"] = pd.to_datetime(df["date"])
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    return df.dropna().set_index("date")["value"]


def stooq(sym: str) -> pd.Series:
    """Daily close for an ETF, trying Stooq (.com then .pl mirror) and falling
    back to Yahoo's chart API. Stooq 404s requests with a bare python UA, so a
    browser UA is sent everywhere. sym like 'hyg.us'."""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    for host in ("stooq.com", "stooq.pl"):
        try:
            r = requests.get(f"https://{host}/q/d/l/?s={sym}&i=d", headers=headers, timeout=30)
            r.raise_for_status()
            df = pd.read_csv(io.StringIO(r.text))
            if "Close" in df.columns and len(df) > 50:
                df["Date"] = pd.to_datetime(df["Date"])
                return df.set_index("Date")["Close"].astype(float)
        except Exception:
            continue
    # Yahoo fallback
    ticker = sym.split(".")[0].upper()
    r = requests.get(
        f"https://query1.finance.yahoo.com/v8/finance/chart/{ticker}",
        params={"range": "2y", "interval": "1d"},
        headers=headers, timeout=30)
    r.raise_for_status()
    res = r.json()["chart"]["result"][0]
    idx = pd.to_datetime(res["timestamp"], unit="s").normalize()
    close = res["indicators"]["quote"][0]["close"]
    return pd.Series(close, index=idx, dtype=float).dropna()


def hy_issuance_from_csv(path: str) -> pd.Series:
    """CSV with columns: date, volume (USD bn, monthly). Anything else is ignored."""
    df = pd.read_csv(path)
    df.columns = [c.strip().lower() for c in df.columns]
    df["date"] = pd.to_datetime(df["date"])
    s = df.set_index("date")["volume"].astype(float).sort_index()
    return s.resample("MS").sum()


def _parse_sifma_workbook(xls: pd.ExcelFile) -> dict:
    """
    Parse SIFMA's 'Issuance' sheet. Layout (as of Aug 2026): one table whose
    first column mixes row types -- annual (int year), 'YTD 2026', quarterly
    ('2Q26'), and monthly (datetime, only the last ~13 months). The High Yield
    (nonconvertible) column is the first 'High Yield' header, in $bn.

    Returns {"monthly": Series, "quarterly": Series, "ytd": {year: value}}.
    We stitch monthly + quarterly/3 into a longer monthly history so the
    12-month baseline isn't starved by SIFMA's short monthly window.
    """
    import re
    raw = xls.parse("Issuance", header=None)
    hdr = None
    for i in range(min(40, len(raw))):
        row = raw.iloc[i].astype(str).str.lower()
        if row.str.contains("high yield").any():
            hdr = i
            break
    if hdr is None:
        raise RuntimeError("No 'High Yield' header on Issuance sheet")
    hycol = [j for j, v in enumerate(raw.iloc[hdr].tolist()) if "high yield" in str(v).lower()][0]
    monthly, quarterly, ytd = {}, {}, {}
    for _, r in raw.iloc[hdr + 1:].iterrows():
        lab, val = r.iloc[0], r.iloc[hycol]
        if pd.isna(lab) or pd.isna(val):
            continue
        val = float(val)
        if isinstance(lab, (pd.Timestamp, datetime)):
            monthly[pd.Timestamp(lab).to_period("M").to_timestamp()] = val
            continue
        ls = str(lab).strip()
        m = re.fullmatch(r"([1-4])Q(\d{2})", ls)
        if m:
            q, yy = int(m.group(1)), 2000 + int(m.group(2))
            quarterly[pd.Timestamp(year=yy, month=3 * q - 2, day=1)] = val
            continue
        m = re.fullmatch(r"YTD\s*(\d{4})", ls, flags=re.I)
        if m:
            ytd[int(m.group(1))] = val
            continue
        # annual rows ignored
    mo = pd.Series(monthly).sort_index()
    qt = pd.Series(quarterly).sort_index()
    if mo.empty:
        raise RuntimeError("No monthly rows found on Issuance sheet")
    # stitch: for months not covered by monthly rows, use quarterly/3
    stitched = {}
    for qstart, v in qt.items():
        for k in range(3):
            mth = qstart + pd.DateOffset(months=k)
            stitched[mth] = v / 3.0
    stitched.update(mo.to_dict())          # real monthly values win
    full = pd.Series(stitched).sort_index()
    return {"monthly": full, "quarterly": qt, "ytd": ytd, "monthly_raw": mo}


def hy_issuance_from_sifma_file(path: str) -> dict:
    """Parse a locally downloaded SIFMA 'US Corporate Bonds Statistics' workbook.
    SIFMA now gates the download behind a form, so fetch it by hand from
    https://www.sifma.org/research/statistics/us-corporate-bonds-statistics
    and pass the file with --sifma-file. .xls needs `pip install xlrd`."""
    engine = "xlrd" if path.lower().endswith(".xls") else None
    return _parse_sifma_workbook(pd.ExcelFile(path, engine=engine))


def hy_issuance_from_sifma() -> dict:
    """Best-effort live fetch. As of Aug 2026 the SIFMA page gates the xls behind a
    HubSpot form, so this will usually fail -- it's kept in case they reopen it."""
    import re
    html = requests.get(SIFMA_PAGE, timeout=30).text
    links = re.findall(r'https://[^"\']+\.xlsx?', html)
    links = [l for l in links if "corporate" in l.lower()]
    if not links:
        raise RuntimeError("No direct xls link on SIFMA page (download is form-gated); use --sifma-file")
    return _parse_sifma_workbook(pd.ExcelFile(io.BytesIO(requests.get(links[0], timeout=60).content)))


# ----------------------------------------------------------------------------
# Evaluators
# ----------------------------------------------------------------------------
def status(level: float, amber: float, red: float) -> str:
    return "RED" if level >= red else "AMBER" if level >= amber else "GREEN"


def eval_sloos() -> dict:
    out = {"indicators": {}, "overall": "GREEN", "_series": {}}
    worst = 0
    for sid, name in SLOOS_SERIES.items():
        try:
            s = fred(sid)
        except Exception as e:  # noqa
            out["indicators"][name] = {"error": str(e)}
            continue
        out["_series"][name] = s
        latest, prev = s.iloc[-1], s.iloc[-2]
        diffs = s.diff().tail(SLOOS_RISE_QTRS)
        rising = bool((diffs > 0).all())
        st = status(latest, SLOOS_AMBER, SLOOS_RED)
        # momentum upgrade: below threshold but rising fast from a low base
        if st == "GREEN" and rising and (latest - s.iloc[-1 - SLOOS_RISE_QTRS]) >= 15:
            st = "AMBER"
        worst = max(worst, {"GREEN": 0, "AMBER": 1, "RED": 2}[st])
        out["indicators"][name] = {
            "series": sid,
            "as_of": s.index[-1].date().isoformat(),
            "net_pct_tightening": round(float(latest), 1),
            "prev_quarter": round(float(prev), 1),
            "rising_consecutive_qtrs": rising,
            "status": st,
        }
    out["overall"] = ["GREEN", "AMBER", "RED"][worst]
    return out


def eval_hy(issuance, spread: pd.Series) -> dict:
    out = {}
    ytd = {}
    if isinstance(issuance, dict):
        ytd = issuance.get("ytd", {})
        issuance = issuance["monthly"]
    # --- spreads context ---
    sp3y = spread[spread.index >= spread.index[-1] - pd.DateOffset(years=3)]
    pctl = float((sp3y < spread.iloc[-1]).mean())
    out["spread"] = {
        "as_of": spread.index[-1].date().isoformat(),
        "oas_pct": round(float(spread.iloc[-1]), 2),
        "pctl_3y": round(pctl, 2),
        "benign": pctl < SPREAD_BENIGN_PCTL,
        "chg_1m_bp": round(float(spread.iloc[-1] - spread.iloc[-22]) * 100, 0) if len(spread) > 22 else None,
    }
    # --- issuance ---
    if issuance is None or len(issuance) < ISS_MIN_HISTORY_MONTHS:
        out["issuance"] = {"error": "insufficient issuance history"}
        out["overall"] = "UNKNOWN"
        return out
    # drop the current partial month if we're mid-month
    if issuance.index[-1].month == date.today().month and issuance.index[-1].year == date.today().year:
        issuance = issuance.iloc[:-1]
    run3 = issuance.tail(3).mean()
    avg12 = issuance.tail(15).head(12).mean()   # the 12 months before the last 3
    ratio = float(run3 / avg12) if avg12 else float("nan")
    st = "RED" if ratio < ISS_RED_RATIO else "AMBER" if ratio < ISS_AMBER_RATIO else "GREEN"
    # the specific "quiet shutdown" pattern: issuance collapsing while spreads still benign
    quiet_shutdown = st != "GREEN" and out["spread"]["benign"]
    yoy = None
    if len(ytd) >= 2:
        yrs = sorted(ytd)
        yoy = ytd[yrs[-1]] / ytd[yrs[-2]] - 1 if ytd[yrs[-2]] else None
    last_m = float(issuance.iloc[-1])
    same_m_ly = issuance.get(issuance.index[-1] - pd.DateOffset(years=1))
    out["issuance"] = {
        "as_of": issuance.index[-1].strftime("%Y-%m"),
        "last_month_bn": round(last_m, 1),
        "same_month_ly_bn": round(float(same_m_ly), 1) if same_m_ly is not None and not pd.isna(same_m_ly) else None,
        "ytd_yoy": round(yoy, 3) if yoy is not None else None,
        "last_3m_avg_bn": round(float(run3), 1),
        "prior_12m_avg_bn": round(float(avg12), 1),
        "ratio": round(ratio, 2),
        "status": st,
        "quiet_shutdown_flag": quiet_shutdown,
    }
    out["overall"] = st
    return out



def eval_market(spread: pd.Series) -> dict:
    """Cross-checks: CCC OAS (weakest credits crack first) and HYG/LQD
    (separates credit-driven from rates-driven ETF weakness)."""
    out = {"overall": "GREEN", "_series": {}}
    rank = {"GREEN": 0, "AMBER": 1, "RED": 2}
    worst = 0
    # CCC
    try:
        ccc = fred(CCC_SPREAD_SERIES)
        out["_series"]["ccc"] = ccc
        c3 = ccc[ccc.index >= ccc.index[-1] - pd.DateOffset(years=3)]
        chg = float(ccc.iloc[-1] - ccc.iloc[-22]) * 100 if len(ccc) > 22 else 0.0
        diff = ccc - spread.reindex(ccc.index).ffill()
        dchg = float(diff.iloc[-1] - diff.iloc[-22]) * 100 if len(diff) > 22 else 0.0
        st = "GREEN"
        if chg >= CCC_AMBER_BP or dchg >= DIFF_AMBER_BP:
            st = "AMBER"
        if chg >= CCC_RED_BP or dchg >= DIFF_RED_BP:
            st = "RED"
        out["ccc"] = {
            "as_of": ccc.index[-1].date().isoformat(),
            "oas_pct": round(float(ccc.iloc[-1]), 2),
            "pctl_3y": round(float((c3 < ccc.iloc[-1]).mean()), 2),
            "chg_1m_bp": round(chg, 0),
            "diff_vs_hy_pct": round(float(diff.iloc[-1]), 2),
            "diff_chg_1m_bp": round(dchg, 0),
            "status": st,
        }
        worst = max(worst, rank[st])
    except Exception as e:
        out["ccc"] = {"error": str(e)}
    # HYG/LQD
    try:
        hyg, lqd = stooq("hyg.us"), stooq("lqd.us")
        ratio = (hyg / lqd.reindex(hyg.index).ffill()).dropna()
        ratio = ratio[ratio.index >= ratio.index[-1] - pd.DateOffset(years=1)]
        out["_series"]["ratio"] = ratio
        rchg = float(ratio.iloc[-1] / ratio.iloc[-22] - 1) if len(ratio) > 22 else 0.0
        hchg = float(hyg.iloc[-1] / hyg.iloc[-22] - 1) if len(hyg) > 22 else 0.0
        st = "AMBER" if rchg <= RATIO_AMBER else "GREEN"
        sp_chg_bp = float(spread.iloc[-1] - spread.iloc[-22]) * 100 if len(spread) > 22 else 0.0
        if st == "AMBER" and sp_chg_bp >= 25:
            driver = "credit-driven: HYG lagging LQD while spreads widen"
        elif hchg <= -0.02 and rchg > RATIO_AMBER:
            driver = "rates-driven: HYG down but keeping pace with LQD"
        else:
            driver = ""
        out["ratio"] = {
            "as_of": ratio.index[-1].date().isoformat(),
            "chg_1m": round(rchg, 4),
            "hyg_chg_1m": round(hchg, 4),
            "status": st,
            "driver": driver,
        }
        worst = max(worst, rank[st])
    except Exception as e:
        out["ratio"] = {"error": str(e)}
    out["overall"] = ["GREEN", "AMBER", "RED"][worst]
    return out


# ----------------------------------------------------------------------------
# HTML report
# ----------------------------------------------------------------------------
_COL = {"GREEN": "#2E7D5B", "AMBER": "#C98A1B", "RED": "#B3352C", "UNKNOWN": "#6B7A86"}

def _axis_labels(series: pd.Series, w: int, h: int, pad: int) -> str:
    """Start / middle / end date labels along the bottom of a sparkline.
    Spans over ~2 years label as years ("2014"); shorter spans as "Jul '25"
    so a 12y quarterly window can't be misread as days in one month."""
    span_days = (series.index[-1] - series.index[0]).days
    if span_days > 730:
        fmt = lambda t: pd.Timestamp(t).strftime("%Y")
    else:
        fmt = lambda t: pd.Timestamp(t).strftime("%b '%y")
    mid = series.index[len(series) // 2]
    return (
        f'<text x="{pad}" y="{h-2}" font-size="9" font-family="Consolas,monospace" fill="#8A97A3">{fmt(series.index[0])}</text>'
        f'<text x="{w/2:.0f}" y="{h-2}" font-size="9" font-family="Consolas,monospace" fill="#8A97A3" text-anchor="middle">{fmt(mid)}</text>'
        f'<text x="{w-pad}" y="{h-2}" font-size="9" font-family="Consolas,monospace" fill="#8A97A3" text-anchor="end">{fmt(series.index[-1])}</text>')


def _svg_line(series: pd.Series, w=420, h=102, color="#1F3A5F", band=None, fmt="{:.1f}") -> str:
    """Tiny sparkline with date axis and optional horizontal threshold band."""
    if series is None or len(series) < 2:
        return ""
    lab_h = 12
    ph = h - lab_h
    y = series.values.astype(float)
    x = range(len(y))
    lo, hi = float(min(y.min(), 0) if band else y.min()), float(y.max())
    if band:
        hi = max(hi, max(b[0] for b in band))
    rng = (hi - lo) or 1.0
    pad = 6
    def X(i): return pad + (w - 2 * pad) * i / (len(y) - 1)
    def Y(v): return pad + (ph - 2 * pad) * (1 - (v - lo) / rng)
    pts = " ".join(f"{X(i):.1f},{Y(v):.1f}" for i, v in zip(x, y))
    out = [f'<svg viewBox="0 0 {w} {h}" width="100%" height="{h}" preserveAspectRatio="none">']
    for lvl, c in (band or []):
        out.append(f'<line x1="{pad}" y1="{Y(lvl):.1f}" x2="{w-pad}" y2="{Y(lvl):.1f}" stroke="{c}" stroke-dasharray="3 3" stroke-width="1"/>')
    if lo < 0 < hi:
        out.append(f'<line x1="{pad}" y1="{Y(0):.1f}" x2="{w-pad}" y2="{Y(0):.1f}" stroke="#9AA7B2" stroke-width="1"/>')
    out.append(f'<polyline fill="none" stroke="{color}" stroke-width="1.8" points="{pts}"/>')
    out.append(f'<circle cx="{X(len(y)-1):.1f}" cy="{Y(y[-1]):.1f}" r="3" fill="{color}"/>')
    # min/max value hints on the left
    out.append(f'<text x="{pad}" y="{Y(hi)+10:.1f}" font-size="9" font-family="Consolas,monospace" fill="#8A97A3">{fmt.format(hi)}</text>')
    out.append(f'<text x="{pad}" y="{Y(lo)-3:.1f}" font-size="9" font-family="Consolas,monospace" fill="#8A97A3">{fmt.format(lo)}</text>')
    out.append(_axis_labels(series, w, h, pad))
    out.append("</svg>")
    return "".join(out)


def _svg_bars(series: pd.Series, w=420, h=102, color="#1F3A5F", hi_color="#B3352C", highlight_last=3) -> str:
    if series is None or len(series) < 2:
        return ""
    lab_h = 12
    ph = h - lab_h
    y = series.values.astype(float)
    n = len(y)
    pad = 6
    top = float(y.max()) or 1.0
    bw = (w - 2 * pad) / n
    out = [f'<svg viewBox="0 0 {w} {h}" width="100%" height="{h}" preserveAspectRatio="none">']
    avg = float(y[:-highlight_last].mean()) if n > highlight_last else float(y.mean())
    ya = pad + (ph - 2 * pad) * (1 - avg / top)
    out.append(f'<line x1="{pad}" y1="{ya:.1f}" x2="{w-pad}" y2="{ya:.1f}" stroke="#9AA7B2" stroke-dasharray="3 3"/>')
    for i, v in enumerate(y):
        bh = (ph - 2 * pad) * v / top
        c = hi_color if i >= n - highlight_last else color
        out.append(f'<rect x="{pad + i*bw + 1:.1f}" y="{ph - pad - bh:.1f}" width="{max(bw-2,1):.1f}" height="{bh:.1f}" fill="{c}" opacity="{0.95 if i >= n-highlight_last else 0.75}"/>')
    out.append(f'<text x="{pad}" y="12" font-size="9" font-family="Consolas,monospace" fill="#8A97A3">peak ${top:.0f}bn</text>')
    out.append(_axis_labels(series, w, h, pad))
    out.append("</svg>")
    return "".join(out)


def render_html(sloos: dict, hy: dict, market: dict, composite: str, spread: pd.Series,
                issuance: pd.Series | None, sifma_file: str | None) -> str:
    import os, html as H
    today = date.today().isoformat()
    sif_note = ""
    if sifma_file and os.path.exists(sifma_file):
        age = (date.today() - date.fromtimestamp(os.path.getmtime(sifma_file))).days
        sif_note = f"SIFMA file dated {date.fromtimestamp(os.path.getmtime(sifma_file)).isoformat()} ({age} days old)"
        if age > 40:
            sif_note += " &mdash; <b>likely stale, download the new workbook</b>"
    lamp = lambda st: f'<span class="lamp" style="background:{_COL[st]}"></span>'

    # SLOOS cards
    sl_rows = []
    for name, d in sloos["indicators"].items():
        if "error" in d:
            sl_rows.append(f'<div class="row"><div class="k">{H.escape(name)}</div><div class="v err">{H.escape(d["error"])}</div></div>')
            continue
        ser = sloos["_series"].get(name)
        ser = ser[ser.index >= ser.index[-1] - pd.DateOffset(years=12)] if ser is not None else None
        spark = _svg_line(ser, color="#1F3A5F", band=[(SLOOS_AMBER, _COL["AMBER"]), (SLOOS_RED, _COL["RED"])])
        trend = "&#9650; rising" if d["rising_consecutive_qtrs"] else ""
        sl_rows.append(
            f'<div class="row"><div class="k">{lamp(d["status"])}{H.escape(name)}'
            f'<span class="sub">{d["as_of"]}</span></div>'
            f'<div class="num">{d["net_pct_tightening"]:+.1f}%<span class="sub">prev {d["prev_quarter"]:+.1f}% {trend}</span></div>'
            f'<div class="spark">{spark}</div></div>')

    # HY cards
    sp = hy["spread"]
    sp3 = spread[spread.index >= spread.index[-1] - pd.DateOffset(years=3)]
    sp_spark = _svg_line(sp3, color="#1F3A5F")
    iss = hy.get("issuance", {})
    if "error" in iss:
        iss_html = f'<div class="row"><div class="k">{lamp("UNKNOWN")}HY issuance</div><div class="v err">{H.escape(iss["error"])}</div></div>'
    else:
        extra = ""
        if iss.get("same_month_ly_bn") is not None:
            extra += f" &middot; same month LY ${iss['same_month_ly_bn']:.1f}bn"
        if iss.get("ytd_yoy") is not None:
            extra += f" &middot; YTD {iss['ytd_yoy']:+.1%} y/y"
        qs = '<div class="flag">QUIET SHUTDOWN: issuance drying up while spreads look benign</div>' if iss["quiet_shutdown_flag"] else ""
        iss_html = (
            f'<div class="row"><div class="k">{lamp(iss["status"])}HY issuance, 3m vs prior 12m'
            f'<span class="sub">as of {iss["as_of"]}</span></div>'
            f'<div class="num">{iss["ratio"]:.2f}&times;<span class="sub">${iss["last_3m_avg_bn"]:.1f}bn vs ${iss["prior_12m_avg_bn"]:.1f}bn</span></div>'
            f'<div class="spark">{_svg_bars(issuance.tail(28)) if issuance is not None else ""}</div></div>'
            f'<div class="note">Last month ${iss["last_month_bn"]:.1f}bn{extra}. Red bars = last 3 months; dashed = prior-12m average.</div>{qs}')
    sp_html = (
        f'<div class="row"><div class="k">{lamp("GREEN" if sp["benign"] else "AMBER")}HY OAS'
        f'<span class="sub">{sp["as_of"]}</span></div>'
        f'<div class="num">{sp["oas_pct"]:.2f}%<span class="sub">{sp["pctl_3y"]:.0%} pctl of 3y &middot; 1m {sp["chg_1m_bp"]:+.0f}bp</span></div>'
        f'<div class="spark">{sp_spark}</div></div>')

    ccc = market.get("ccc", {})
    if "error" in ccc:
        ccc_html = f'<div class="row"><div class="k">{lamp("UNKNOWN")}CCC OAS</div><div class="v err">{H.escape(ccc["error"])}</div></div>'
    else:
        cser = market["_series"].get("ccc")
        cser = cser[cser.index >= cser.index[-1] - pd.DateOffset(years=3)] if cser is not None else None
        ccc_html = (
            f'<div class="row"><div class="k">{lamp(ccc["status"])}CCC &amp; lower OAS'
            f'<span class="sub">{ccc["as_of"]}</span></div>'
            f'<div class="num">{ccc["oas_pct"]:.2f}%<span class="sub">{ccc["pctl_3y"]:.0%} pctl 3y &middot; 1m {ccc["chg_1m_bp"]:+.0f}bp &middot; vs HY {ccc["diff_vs_hy_pct"]:.2f}% ({ccc["diff_chg_1m_bp"]:+.0f}bp)</span></div>'
            f'<div class="spark">{_svg_line(cser, color="#1F3A5F")}</div></div>')
    rt = market.get("ratio", {})
    if "error" in rt:
        ratio_html = f'<div class="row"><div class="k">{lamp("UNKNOWN")}HYG / LQD</div><div class="v err">{H.escape(rt["error"])}</div></div>'
    else:
        rser = market["_series"].get("ratio")
        drv = f'<div class="note">{H.escape(rt["driver"])}</div>' if rt.get("driver") else ""
        ratio_html = (
            f'<div class="row"><div class="k">{lamp(rt["status"])}HYG / LQD ratio'
            f'<span class="sub">{rt["as_of"]}</span></div>'
            f'<div class="num">{rt["chg_1m"]:+.1%}<span class="sub">1m &middot; HYG itself {rt["hyg_chg_1m"]:+.1%}</span></div>'
            f'<div class="spark">{_svg_line(rser, color="#1F3A5F")}</div></div>{drv}')
    both = sloos["overall"] != "GREEN" and hy["overall"] not in ("GREEN", "UNKNOWN")
    verdict = {
        "GREEN": "Neither warning is flashing. Credit is open; stay with the trend.",
        "AMBER": "One leg is stirring. Not a signal yet; tighten the watch.",
        "RED": "A warning is live. Check internals and size accordingly.",
    }[composite] + (" Both legs are flashing &mdash; that is the combination that matters." if both else "")

    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Credit warning monitor &mdash; {today}</title>
<style>
:root{{--bg:#EAF0F4;--ink:#14212E;--mute:#5B6B78;--card:#FFFFFF;--line:#CFD9E1;--mono:"Cascadia Mono",Consolas,"SF Mono",Menlo,monospace;--sans:"Segoe UI",system-ui,-apple-system,sans-serif}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--bg);color:var(--ink);font-family:var(--sans);line-height:1.4}}
.wrap{{max-width:980px;margin:0 auto;padding:28px 20px 60px}}
header{{display:flex;justify-content:space-between;align-items:baseline;gap:12px;flex-wrap:wrap;margin-bottom:18px}}
h1{{font-size:15px;letter-spacing:.12em;text-transform:uppercase;font-weight:600;margin:0;color:var(--mute)}}
.when{{font-family:var(--mono);font-size:12px;color:var(--mute)}}
.banner{{background:{_COL[composite]};color:#fff;border-radius:10px;padding:22px 24px;margin-bottom:22px;display:grid;grid-template-columns:auto 1fr;gap:20px;align-items:center}}
.banner .big{{font-family:var(--mono);font-size:42px;font-weight:700;letter-spacing:.06em;line-height:1}}
.banner .txt{{font-size:16px;opacity:.95}}
.grid{{display:grid;grid-template-columns:1fr 1fr;gap:18px}}
@media(max-width:760px){{.grid{{grid-template-columns:1fr}}.row{{grid-template-columns:1fr}}.spark{{grid-column:1/-1}}}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:18px 18px 10px}}
.card h2{{font-size:13px;letter-spacing:.1em;text-transform:uppercase;margin:0 0 4px;display:flex;align-items:center;gap:8px}}
.card .lead{{font-size:12px;color:var(--mute);margin:0 0 12px}}
.row{{display:grid;grid-template-columns:1fr 1fr;gap:8px 14px;padding:12px 0;border-top:1px solid var(--line)}}
.row .k{{font-size:14px;font-weight:600;display:flex;align-items:center;gap:8px;flex-wrap:wrap}}
.row .num{{font-family:var(--mono);font-size:22px;text-align:right}}
.sub{{display:block;font-family:var(--mono);font-size:11px;color:var(--mute);font-weight:400;margin-top:2px}}
.spark{{grid-column:1/-1}}
.lamp{{display:inline-block;width:11px;height:11px;border-radius:50%;box-shadow:0 0 0 2px #fff inset}}
.note{{font-size:12px;color:var(--mute);padding:0 0 10px}}
.flag{{background:#FBEAE8;color:#7A1F18;border-left:4px solid #B3352C;padding:8px 10px;font-size:13px;font-weight:600;margin-bottom:10px;border-radius:4px}}
.err{{color:#B3352C;font-family:var(--mono);font-size:12px}}
footer{{margin-top:22px;font-size:11px;color:var(--mute);font-family:var(--mono)}}
</style></head><body><div class="wrap">
<header><h1>Credit warning monitor</h1><div class="when">generated {today}</div></header>
<div class="banner"><div class="big">{composite}</div><div class="txt">{verdict}</div></div>
<div class="grid">
<div class="card"><h2>{lamp(sloos["overall"])}Bank lending standards</h2>
<p class="lead">Fed SLOOS, net % of banks tightening. Amber &ge; {SLOOS_AMBER}, red &ge; {SLOOS_RED}. Leads HY spreads by quarters.</p>
{''.join(sl_rows)}</div>
<div class="card"><h2>{lamp(hy["overall"])}High-yield primary market</h2>
<p class="lead">Issuance dries up before spreads widen. Amber &lt; {ISS_AMBER_RATIO:.0%}, red &lt; {ISS_RED_RATIO:.0%} of the prior-12m run rate.</p>
{iss_html}{sp_html}</div>
<div class="card"><h2>{lamp(market["overall"])}Credit cross-checks</h2>
<p class="lead">CCC cracks first; HYG/LQD separates credit weakness from rates weakness. Ratio alone caps at amber.</p>
{ccc_html}{ratio_html}</div>
</div>
<footer>Sources: FRED (DRTSCILM, DRTSCIS, DRTSCLCC, BAMLH0A0HYM2, BAMLH0A3HYC) &middot; Stooq (HYG, LQD) &middot; SIFMA US Corporate Bond Statistics. {sif_note}</footer>
</div></body></html>"""

# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--hy-csv", help="local monthly HY issuance CSV (date,volume)")
    ap.add_argument("--sifma-file", help="locally downloaded SIFMA corporate-bond statistics xls/xlsx")
    ap.add_argument("--json", help="write results to this JSON file")
    ap.add_argument("--html", help="write an HTML report to this file")
    ap.add_argument("--open", action="store_true", help="open the HTML report in your browser after writing")
    args = ap.parse_args()

    print(f"Credit warning monitor  -  {date.today()}\n")

    # 1) Lending standards
    sloos = eval_sloos()
    print("1) BANK LENDING STANDARDS (SLOOS, net % tightening)")
    for name, d in sloos["indicators"].items():
        if "error" in d:
            print(f"   {name:<30} ERROR {d['error']}")
            continue
        arrow = "^" if d["rising_consecutive_qtrs"] else " "
        print(f"   {name:<30} {d['net_pct_tightening']:>6.1f}  (prev {d['prev_quarter']:>5.1f}) {arrow}  [{d['status']}]  as of {d['as_of']}")
    print(f"   -> overall: {sloos['overall']}\n")

    # 2) HY issuance + spreads
    print("2) HIGH-YIELD ISSUANCE + SPREADS")
    issuance = None
    try:
        if args.hy_csv:
            issuance = hy_issuance_from_csv(args.hy_csv)
        elif args.sifma_file:
            issuance = hy_issuance_from_sifma_file(args.sifma_file)
        else:
            issuance = hy_issuance_from_sifma()
    except Exception as e:
        print(f"   issuance source failed: {e}")
        print("   (download the xls from SIFMA and pass --sifma-file, or --hy-csv with date,volume)")
    try:
        spread = fred(HY_SPREAD_SERIES)
    except Exception as e:
        print(f"   FRED spread fetch failed: {e}")
        sys.exit(1)
    hy = eval_hy(issuance, spread)
    iss_series = issuance["monthly"] if isinstance(issuance, dict) else issuance
    sp = hy["spread"]
    print(f"   HY OAS {sp['oas_pct']:.2f}%  ({sp['pctl_3y']:.0%} pctl of 3y, 1m chg {sp['chg_1m_bp']:+.0f}bp)  benign={sp['benign']}")
    if "error" in hy["issuance"]:
        print(f"   issuance: {hy['issuance']['error']}")
    else:
        i = hy["issuance"]
        print(f"   issuance last 3m avg ${i['last_3m_avg_bn']:.1f}bn vs prior 12m avg ${i['prior_12m_avg_bn']:.1f}bn  ratio {i['ratio']:.2f}  [{i['status']}]  as of {i['as_of']}")
        extra = f"   last month ${i['last_month_bn']:.1f}bn"
        if i.get("same_month_ly_bn") is not None:
            extra += f" (same month last year ${i['same_month_ly_bn']:.1f}bn)"
        if i.get("ytd_yoy") is not None:
            extra += f"   YTD HY issuance {i['ytd_yoy']:+.1%} y/y"
        print(extra)
        if i["quiet_shutdown_flag"]:
            print("   !! QUIET SHUTDOWN: issuance drying up while spreads still look benign")
    print(f"   -> overall: {hy['overall']}\n")

    # 3) Cross-checks
    print("3) CREDIT CROSS-CHECKS")
    market = eval_market(spread)
    c = market.get("ccc", {})
    if "error" in c:
        print(f"   CCC OAS fetch failed: {c['error']}")
    else:
        print(f"   CCC OAS {c['oas_pct']:.2f}%  ({c['pctl_3y']:.0%} pctl 3y, 1m {c['chg_1m_bp']:+.0f}bp)  vs HY {c['diff_vs_hy_pct']:.2f}% ({c['diff_chg_1m_bp']:+.0f}bp 1m)  [{c['status']}]")
    r = market.get("ratio", {})
    if "error" in r:
        print(f"   HYG/LQD fetch failed: {r['error']}")
    else:
        line = f"   HYG/LQD 1m {r['chg_1m']:+.1%}  (HYG {r['hyg_chg_1m']:+.1%})  [{r['status']}]"
        if r.get("driver"):
            line += f"  -- {r['driver']}"
        print(line)
    print(f"   -> overall: {market['overall']}\n")

    # 4) Composite
    rank = {"GREEN": 0, "AMBER": 1, "RED": 2, "UNKNOWN": 0}
    composite = ["GREEN", "AMBER", "RED"][max(rank[sloos["overall"]], rank[hy["overall"]], rank[market["overall"]])]
    both = rank[sloos["overall"]] >= 1 and rank[hy["overall"]] >= 1
    print(f"COMPOSITE: {composite}" + ("  (both legs flashing -- that's the combination that matters)" if both else ""))

    if args.json:
        sl_json = {k: v for k, v in sloos.items() if not k.startswith("_")}
        with open(args.json, "w") as f:
            mk_json = {k: v for k, v in market.items() if not k.startswith("_")}
            json.dump({"date": date.today().isoformat(), "sloos": sl_json, "hy": hy,
                       "cross_checks": mk_json, "composite": composite}, f, indent=2, default=str)
        print(f"\nwritten: {args.json}")
    if args.html:
        page = render_html(sloos, hy, market, composite, spread, iss_series, args.sifma_file)
        with open(args.html, "w", encoding="utf-8") as f:
            f.write(page)
        print(f"written: {args.html}")
        if args.open:
            import webbrowser, os
            webbrowser.open("file://" + os.path.abspath(args.html))


if __name__ == "__main__":
    main()
