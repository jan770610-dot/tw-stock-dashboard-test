# -*- coding: utf-8 -*-
"""Build the compact next-session baseline used by intraday_live.py.

v3.3.1 addition:
- publish exact latest-formal-day all-market RS as ``formal_rs`` so the
  dashboard can show RS for ordinary stocks too, not only stocks already in
  the strong/recovery lists.

The formal RS definition matches automation/rs/rs_breadth.py:
250-trading-day return ranked cross-sectionally on the latest formal date.
"""
from __future__ import annotations

from pathlib import Path
import json
import os
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
STORE = ROOT / "automation" / "rs" / "cache" / "recent_market_store.pkl.gz"
OUT = ROOT / "intraday_baseline.pkl.gz"
META = ROOT / "intraday_baseline_status.json"


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s, errors="coerce")


def _window_sum(s: pd.Series, n: int):
    x = _num(s).dropna()
    return float(x.tail(n).sum()) if len(x) >= n else None


def _window_max(s: pd.Series, n: int):
    x = _num(s).dropna()
    return float(x.tail(n).max()) if len(x) >= n else None


def _window_min(s: pd.Series, n: int):
    x = _num(s).dropna()
    return float(x.tail(n).min()) if len(x) >= n else None


def build() -> pd.DataFrame:
    if not STORE.exists():
        raise FileNotFoundError(f"RS recent market store not found: {STORE}")

    data = pd.read_pickle(STORE, compression="gzip")
    required = {"date", "code", "name", "market", "close", "high", "low"}
    missing = required.difference(data.columns)
    if missing:
        raise ValueError(f"recent market store missing columns: {sorted(missing)}")

    x = data.copy()
    x["date"] = pd.to_datetime(x["date"], errors="coerce")
    x = x.dropna(subset=["date", "code", "close"]).sort_values(["code", "date"])
    x["code"] = x["code"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    latest_date = x["date"].max()
    if pd.isna(latest_date):
        raise ValueError("recent market store contains no valid date")
    latest_iso = str(pd.Timestamp(latest_date).date())

    rows = []
    for code, g in x.groupby("code", sort=False, observed=True):
        g = g.sort_values("date").drop_duplicates("date", keep="last")
        closes = _num(g["close"])
        highs = _num(g["high"])
        lows = _num(g["low"])
        valid_closes = closes.dropna()
        if valid_closes.empty:
            continue

        last = g.iloc[-1]
        base_date = pd.Timestamp(last["date"]).date().isoformat()

        # Exact formal-day 250-session return used by the RS system:
        # current formal close / close shifted by 250 rows - 1.
        # Only assign on the market's latest formal date; a suspended stock whose
        # last row is older is not ranked into that day's cross-section.
        formal_ret250 = None
        if base_date == latest_iso and len(valid_closes) >= 251:
            base_formal = float(valid_closes.iloc[-251])
            cur_formal = float(valid_closes.iloc[-1])
            if base_formal != 0:
                formal_ret250 = cur_formal / base_formal - 1.0

        # When today's live quote is appended, the rolling windows contain
        # current quote + previous (window-1) formal closes/highs/lows.
        row = {
            "base_date": base_date,
            "code": str(code),
            "name": str(last.get("name", "")),
            "market": str(last.get("market", "")),
            "prev_close": float(valid_closes.iloc[-1]),
            "history_rows": int(valid_closes.size),
            "sum_close_19": _window_sum(closes, 19),
            "sum_close_49": _window_sum(closes, 49),
            "sum_close_59": _window_sum(closes, 59),
            "sum_close_199": _window_sum(closes, 199),
            "ma20_prev": float(valid_closes.tail(20).mean()) if valid_closes.size >= 20 else None,
            "ma50_prev": float(valid_closes.tail(50).mean()) if valid_closes.size >= 50 else None,
            "ma60_prev": float(valid_closes.tail(60).mean()) if valid_closes.size >= 60 else None,
            "ma200_prev": float(valid_closes.tail(200).mean()) if valid_closes.size >= 200 else None,
            # For the *next live day* t, t-250 is 249 formal sessions before t-1,
            # hence -250 here. This existing field is intentionally unchanged.
            "rs_base_close": float(valid_closes.iloc[-250]) if valid_closes.size >= 250 else None,
            "formal_ret250": formal_ret250,
            "max_high_249": _window_max(highs, 249),
            "min_low_249": _window_min(lows, 249),
            "max_close_249": _window_max(closes, 249),
            "min_close_249": _window_min(closes, 249),
        }
        rows.append(row)

    out = pd.DataFrame(rows)
    if out.empty:
        raise ValueError("intraday baseline is empty")

    # Cross-sectional formal RS for *all* stocks with enough history on latest day.
    out["formal_rs"] = pd.NA
    active = out["base_date"].astype(str).eq(latest_iso)
    valid = active & pd.to_numeric(out["formal_ret250"], errors="coerce").notna()
    if valid.any():
        ranks = pd.to_numeric(out.loc[valid, "formal_ret250"], errors="coerce").rank(
            method="average", pct=True
        ) * 100.0
        out.loc[valid, "formal_rs"] = ranks.round(4)
    out["formal_rs"] = pd.to_numeric(out["formal_rs"], errors="coerce")

    tmp = OUT.with_name(OUT.name + ".tmp")
    out.to_pickle(tmp, compression="gzip")
    os.replace(tmp, OUT)

    meta = {
        "status": "success",
        "base_date": latest_iso,
        "stocks": int(len(out)),
        "eligible_250d": int(out["rs_base_close"].notna().sum()),
        "formal_rs_count": int(out["formal_rs"].notna().sum()),
        "output": OUT.name,
    }
    META.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[INTRADAY BASELINE] {meta}", flush=True)
    return out


if __name__ == "__main__":
    build()
