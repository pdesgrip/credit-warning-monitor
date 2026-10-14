"""
Offline tests -- no network. Run with:  python -m pytest -q

Covers the parts most likely to break silently:
  * the SIFMA workbook parser (mixed annual / YTD / quarterly / monthly rows)
  * the issuance scoring and the "quiet shutdown" flag
  * the CCC / HYG-LQD cross-checks
  * sparkline date-axis formatting
"""
import re
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import credit_warning_monitor as m  # noqa: E402


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------
@pytest.fixture
def sifma_workbook(tmp_path):
    """Synthetic workbook mimicking SIFMA's 'Issuance' sheet layout:
    metadata rows, a two-row header, then annual / YTD / quarterly / monthly
    rows all sharing the first column. Real SIFMA data is not redistributed."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Issuance"
    ws.append(["Security:", "US Corporate Bonds"])
    ws.append(["Units:", "$ Billion"])
    ws.append([None])
    ws.append([None, "Nonconvertible"])
    ws.append([None, "Investment Grade", "High Yield", "Total"])
    for yr in (2023, 2024, 2025):                          # annual -> ignored
        ws.append([yr, 1300.0, 300.0, 1600.0])
    ws.append([None])
    ws.append(["YTD 2025", 1000.0, 200.0, 1200.0])
    ws.append(["YTD 2026", 1100.0, 220.0, 1320.0])
    ws.append([None])
    for q in ("1Q25", "2Q25", "3Q25", "4Q25", "1Q26", "2Q26"):  # quarterly
        ws.append([q, 400.0, 90.0, 490.0])                     # 30/month
    ws.append([None])
    months = pd.date_range("2025-08-31", periods=13, freq="ME")
    vals = [30.0] * 10 + [12.0, 10.0, 8.0]                     # last 3 collapse
    for d, v in zip(months, vals):
        ws.append([d.to_pydatetime(), 150.0, v, 150.0 + v])
    path = tmp_path / "sifma.xlsx"
    wb.save(path)
    return path


def daily(values, end="2026-10-02"):
    idx = pd.bdate_range(end=end, periods=len(values))
    return pd.Series(values, index=idx, dtype=float)


# ---------------------------------------------------------------------------
# SIFMA parser
# ---------------------------------------------------------------------------
def test_parser_reads_each_row_type(sifma_workbook):
    d = m.hy_issuance_from_sifma_file(str(sifma_workbook))
    assert d["ytd"] == {2025: 200.0, 2026: 220.0}
    assert len(d["quarterly"]) == 6
    assert len(d["monthly_raw"]) == 13
    # annual rows must not leak in as bogus dates
    assert d["monthly"].index.min() >= pd.Timestamp("2025-01-01")


def test_parser_prefers_monthly_over_quarterly(sifma_workbook):
    d = m.hy_issuance_from_sifma_file(str(sifma_workbook))
    full = d["monthly"]
    assert full[pd.Timestamp("2026-08-01")] == pytest.approx(8.0)    # real monthly
    assert full[pd.Timestamp("2025-02-01")] == pytest.approx(30.0)   # quarterly / 3
    assert full.index.is_monotonic_increasing


# ---------------------------------------------------------------------------
# Issuance scoring
# ---------------------------------------------------------------------------
def test_quiet_shutdown_fires_when_issuance_collapses_and_spreads_benign(sifma_workbook):
    d = m.hy_issuance_from_sifma_file(str(sifma_workbook))
    spread = daily(np.linspace(4.0, 2.7, 800))          # tightening -> benign
    out = m.eval_hy(d, spread)
    iss = out["issuance"]
    assert iss["ratio"] == pytest.approx(10.0 / 30.0, abs=0.01)
    assert iss["status"] == "RED"
    assert iss["quiet_shutdown_flag"] is True
    assert iss["ytd_yoy"] == pytest.approx(0.10)


def test_short_history_is_unknown_not_green():
    spread = daily(np.full(300, 3.0))
    out = m.eval_hy(pd.Series([30.0] * 5, index=pd.date_range("2026-01-01", periods=5, freq="MS")), spread)
    assert out["overall"] == "UNKNOWN"


# ---------------------------------------------------------------------------
# Cross-checks
# ---------------------------------------------------------------------------
def test_ccc_decompression_turns_red(monkeypatch):
    hy = daily(np.full(800, 3.0))
    ccc_vals = np.r_[np.full(770, 9.0), np.linspace(9.0, 11.0, 30)]   # +200bp
    monkeypatch.setattr(m, "fred", lambda sid: daily(ccc_vals))
    flat = daily(np.full(300, 80.0))
    monkeypatch.setattr(m, "stooq", lambda sym: flat)
    out = m.eval_market(hy)
    assert out["ccc"]["status"] == "RED"
    assert out["ratio"]["status"] == "GREEN"
    assert out["overall"] == "RED"


def test_rates_driven_label(monkeypatch):
    hy = daily(np.full(800, 3.0))
    monkeypatch.setattr(m, "fred", lambda sid: daily(np.full(800, 9.0)))
    falling = np.r_[np.full(270, 80.0), np.linspace(80.0, 76.0, 30)]  # both -5%
    prices = {"hyg.us": daily(falling), "lqd.us": daily(falling * 1.3)}
    monkeypatch.setattr(m, "stooq", lambda sym: prices[sym])
    out = m.eval_market(hy)
    assert out["ratio"]["driver"].startswith("rates-driven")


# ---------------------------------------------------------------------------
# Chart labels
# ---------------------------------------------------------------------------
def _labels(svg):
    return re.findall(r">([^<]+)</text>", svg)[-3:]


def test_long_span_axis_uses_years():
    s = pd.Series(np.arange(49.0), index=pd.date_range("2014-07-01", periods=49, freq="QS"))
    assert _labels(m._svg_line(s)) == ["2014", "2020", "2026"]


def test_short_span_axis_uses_month_and_year():
    s = pd.Series(np.arange(13.0), index=pd.date_range("2025-09-01", periods=13, freq="MS"))
    assert _labels(m._svg_line(s)) == ["Sep '25", "Mar '26", "Sep '26"]
