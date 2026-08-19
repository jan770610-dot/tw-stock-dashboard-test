# -*- coding: utf-8 -*-
"""Build the compact next-session baseline used by intraday_live.py.

This script is intentionally run *after* the formal RS/Duckbill daily update.
It reads the existing RS recent-market cache and publishes only the rolling
values needed to recompute today's indicators with a live quote.  It does not
modify rs_latest.xlsx or duck_latest.xlsx.
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

    rows = []
    for code, g in x.groupby("code", sort=False, observed=True):
        g = g.sort_values("date").drop_duplicates("date", keep="last")
        closes = _num(g["close"])
        highs = _num(g["high"])
        lows = _num(g["low"])
        if closes.dropna().empty:
            continue

        last = g.iloc[-1]
        # When today's live quote is appended, the rolling windows contain
        # current quote + previous (window-1) formal closes/highs/lows.
        row = {
            "base_date": pd.Timestamp(last["date"]).date().isoformat(),
            "code": str(code),
            "name": str(last.get("name", "")),
            "market": str(last.get("market", "")),
            "prev_close": float(closes.iloc[-1]),
            "history_rows": int(closes.notna().sum()),
            "sum_close_19": _window_sum(closes, 19),
            "sum_close_49": _window_sum(closes, 49),
            "sum_close_59": _window_sum(closes, 59),
            "sum_close_199": _window_sum(closes, 199),
            "ma20_prev": float(closes.tail(20).mean()) if closes.notna().sum() >= 20 else None,
            "ma50_prev": float(closes.tail(50).mean()) if closes.notna().sum() >= 50 else None,
            "ma60_prev": float(closes.tail(60).mean()) if closes.notna().sum() >= 60 else None,
            "ma200_prev": float(closes.tail(200).mean()) if closes.notna().sum() >= 200 else None,
            "rs_base_close": float(closes.dropna().iloc[-250]) if closes.notna().sum() >= 250 else None,
            "max_high_249": _window_max(highs, 249),
            "min_low_249": _window_min(lows, 249),
            "max_close_249": _window_max(closes, 249),
            "min_close_249": _window_min(closes, 249),
        }
        rows.append(row)

    out = pd.DataFrame(rows)
    if out.empty:
        raise ValueError("intraday baseline is empty")

    tmp = OUT.with_name(OUT.name + ".tmp")
    out.to_pickle(tmp, compression="gzip")
    os.replace(tmp, OUT)

    meta = {
        "status": "success",
        "base_date": str(latest_date.date()),
        "stocks": int(len(out)),
        "eligible_250d": int(out["rs_base_close"].notna().sum()),
        "output": OUT.name,
    }
    META.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[INTRADAY BASELINE] {meta}", flush=True)
    return out


if __name__ == "__main__":
    build()
