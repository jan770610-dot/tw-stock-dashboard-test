# -*- coding: utf-8 -*-
"""Compare exit rules for the Taiwan-stock Duckbill / RS decision system.

Design goals
------------
1. Reconstruct only point-in-time technical signals from official TWSE/TPEx daily
   cache files. No future data is used in signal generation.
2. Entry signal is known after day T close; entry is executed on T+1 open.
3. Exit signal is known after day T close; exit is executed on T+1 open.
4. Compare many existing-system exit indicators with the *same* entry events.
5. Keep a common 60-session future window for peak-capture / give-back metrics.

The script reads automation/rs/cache/YYYYMMDD.csv. Missing weekday files can be
fetched with the existing RS official-data fetcher when --fetch-missing is used.
Outputs are written at repository root:
  - exit_backtest_latest.xlsx
  - exit_backtest_summary.csv
  - exit_backtest_status.json

This is research tooling, not an order-execution engine.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RS_DIR = ROOT / "automation" / "rs"
CACHE_DIR = RS_DIR / "cache"
HOLIDAY_FILE = CACHE_DIR / "nontrading_dates.txt"
STATUS_FILE = ROOT / "update_status.json"
OUT_XLSX = ROOT / "exit_backtest_latest.xlsx"
OUT_SUMMARY = ROOT / "exit_backtest_summary.csv"
OUT_STATUS = ROOT / "exit_backtest_status.json"

DEFAULT_START = dt.date(2023, 1, 3)
DEFAULT_MAX_HOLD = 60
WARMUP_CALENDAR_DAYS = 430  # comfortably covers 250 trading sessions


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------
def _num(v, default=None):
    try:
        if pd.isna(v):
            return default
        x = float(v)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _date(v: str) -> dt.date:
    return dt.datetime.strptime(str(v), "%Y-%m-%d").date()


def _latest_formal_date() -> dt.date:
    try:
        obj = json.loads(STATUS_FILE.read_text(encoding="utf-8"))
        for k in ("target_date", "rs_date", "duck_date"):
            s = str(obj.get(k) or "").strip()
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
                return _date(s)
    except Exception:
        pass
    now = dt.datetime.now(dt.timezone(dt.timedelta(hours=8)))
    d = now.date() if now.hour >= 16 else now.date() - dt.timedelta(days=1)
    while d.weekday() >= 5:
        d -= dt.timedelta(days=1)
    return d


def _known_nontrading_days() -> set[dt.date]:
    out: set[dt.date] = set()
    if not HOLIDAY_FILE.exists():
        return out
    try:
        for raw in HOLIDAY_FILE.read_text(encoding="utf-8").splitlines():
            s = raw.strip()
            if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
                out.add(_date(s))
    except Exception:
        pass
    return out


def _cache_path(day: dt.date) -> Path:
    return CACHE_DIR / f"{day.strftime('%Y%m%d')}.csv"


def _business_weekdays(start: dt.date, end: dt.date) -> List[dt.date]:
    if end < start:
        return []
    return [
        d.date()
        for d in pd.date_range(start, end, freq="D")
        if d.weekday() < 5
    ]


# ---------------------------------------------------------------------------
# Cache preparation
# ---------------------------------------------------------------------------
def ensure_history_cache(start: dt.date, end: dt.date, fetch_missing: bool, workers: int) -> dict:
    """Ensure daily all-market cache exists over warmup..end.

    Existing official cache files are reused. If requested, missing weekdays are
    fetched through automation/rs/rs_breadth.py, which already validates TWSE +
    TPEx completeness and records non-trading days.
    """
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    known_holidays = _known_nontrading_days()
    weekdays = _business_weekdays(start, end)
    missing = [d for d in weekdays if d not in known_holidays and not _cache_path(d).exists()]
    report = {
        "requested_weekdays": len(weekdays),
        "known_nontrading_before": len([d for d in weekdays if d in known_holidays]),
        "missing_before": len(missing),
        "fetched_ok": 0,
        "new_nontrading": 0,
        "failed": [],
    }
    if not missing or not fetch_missing:
        return report

    # Import only when network backfill is actually requested.
    if str(RS_DIR) not in sys.path:
        sys.path.insert(0, str(RS_DIR))
    import rs_breadth  # type: ignore

    workers = max(1, min(4, int(workers)))
    print(f"[history] missing weekdays={len(missing)}; fetch workers={workers}", flush=True)

    def task(day: dt.date):
        try:
            d, frame, status = rs_breadth.fetch_day(day, analysis_only=False)
            return d, frame is not None and not frame.empty, str(status), None
        except Exception as e:  # pragma: no cover - network path
            return day, False, "exception", f"{type(e).__name__}: {e}"

    completed = 0
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(task, d): d for d in missing}
        for fut in as_completed(futs):
            day = futs[fut]
            d, ok, status, err = fut.result()
            completed += 1
            if ok and _cache_path(d).exists():
                report["fetched_ok"] += 1
            elif status == "holiday":
                report["new_nontrading"] += 1
            else:
                report["failed"].append({"date": d.isoformat(), "status": status, "error": err or ""})
            if completed == 1 or completed % 25 == 0 or completed == len(missing):
                print(
                    f"[history] {completed}/{len(missing)} | ok={report['fetched_ok']} "
                    f"holiday={report['new_nontrading']} fail={len(report['failed'])}",
                    flush=True,
                )

    # One conservative retry for transient/incomplete failures.
    retry_days = []
    for x in report["failed"]:
        try:
            d = _date(x["date"])
            if not _cache_path(d).exists() and d not in _known_nontrading_days():
                retry_days.append(d)
        except Exception:
            pass
    if retry_days:
        print(f"[history] retry {len(retry_days)} dates sequentially", flush=True)
        still = []
        for i, d in enumerate(retry_days, 1):
            try:
                day, frame, status = rs_breadth.fetch_day(d, analysis_only=False)
                if frame is not None and not frame.empty and _cache_path(day).exists():
                    report["fetched_ok"] += 1
                elif status == "holiday":
                    report["new_nontrading"] += 1
                else:
                    still.append({"date": d.isoformat(), "status": str(status), "error": "retry not complete"})
            except Exception as e:  # pragma: no cover
                still.append({"date": d.isoformat(), "status": "retry exception", "error": f"{type(e).__name__}: {e}"})
            if i % 10 == 0:
                time.sleep(0.5)
        report["failed"] = still
    return report


def load_market_cache(start: dt.date, end: dt.date) -> Tuple[pd.DataFrame, dict]:
    files: List[Path] = []
    for p in CACHE_DIR.glob("20??????.csv"):
        try:
            d = dt.datetime.strptime(p.stem, "%Y%m%d").date()
        except Exception:
            continue
        if start <= d <= end:
            files.append(p)
    files.sort()
    if not files:
        raise RuntimeError(f"找不到 {start}~{end} 的 RS 每日行情快取：{CACHE_DIR}")

    frames = []
    errors = []
    for i, p in enumerate(files, 1):
        try:
            z = pd.read_csv(p, dtype={"code": str})
            if z.empty or "code" not in z.columns or "close" not in z.columns:
                continue
            z["date"] = pd.Timestamp(dt.datetime.strptime(p.stem, "%Y%m%d").date())
            keep = [c for c in ["date", "code", "name", "market", "open", "high", "low", "close"] if c in z.columns]
            z = z[keep].copy()
            frames.append(z)
        except Exception as e:
            errors.append(f"{p.name}:{type(e).__name__}:{e}")
        if i == 1 or i % 100 == 0 or i == len(files):
            print(f"[load] {i}/{len(files)} daily cache files", flush=True)

    if not frames:
        raise RuntimeError("行情快取存在，但沒有可解析的股票資料")
    x = pd.concat(frames, ignore_index=True, sort=False)
    x["code"] = x["code"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    for c in ["open", "high", "low", "close"]:
        if c not in x.columns:
            x[c] = np.nan
        x[c] = pd.to_numeric(x[c], errors="coerce").astype("float32")
    x = x.dropna(subset=["date", "code", "close"])
    x = x[x["close"] > 0].drop_duplicates(["date", "code"], keep="last")
    x["open"] = x["open"].where(x["open"] > 0, x["close"])
    x["high"] = x["high"].where(x["high"] > 0, x[["open", "close"]].max(axis=1))
    x["low"] = x["low"].where(x["low"] > 0, x[["open", "close"]].min(axis=1))
    if "name" not in x.columns:
        x["name"] = ""
    if "market" not in x.columns:
        x["market"] = ""
    x["name"] = x["name"].fillna("").astype("category")
    x["market"] = x["market"].fillna("").astype("category")
    x["code"] = x["code"].astype("category")
    x = x.sort_values(["code", "date"]).reset_index(drop=True)
    meta = {
        "cache_files": len(files),
        "rows": int(len(x)),
        "stocks": int(x["code"].nunique()),
        "first_date": x["date"].min().date().isoformat(),
        "last_date": x["date"].max().date().isoformat(),
        "read_errors": errors[:20],
    }
    return x, meta


# ---------------------------------------------------------------------------
# Reconstruct current-system technical / RS / right-side / heat signals
# ---------------------------------------------------------------------------
def compute_signals(df: pd.DataFrame) -> pd.DataFrame:
    x = df.copy().sort_values(["code", "date"]).reset_index(drop=True)
    grp = x.groupby("code", group_keys=False, observed=True)

    x["ma20"] = grp["close"].transform(lambda s: s.rolling(20, min_periods=20).mean()).astype("float32")
    x["ma60"] = grp["close"].transform(lambda s: s.rolling(60, min_periods=60).mean()).astype("float32")
    x["ma20_change"] = grp["ma20"].diff().astype("float32")
    x["ma60_change"] = grp["ma60"].diff().astype("float32")
    x["spread"] = (x["ma20"] - x["ma60"]).astype("float32")
    x["spread_change"] = x.groupby("code", observed=True)["spread"].diff().astype("float32")
    x["bias20"] = ((x["close"] / x["ma20"] - 1.0) * 100.0).astype("float32")

    prev250 = grp["close"].shift(250)
    x["ret250"] = (x["close"].astype("float64") / prev250.astype("float64") - 1.0)
    x["rs"] = (
        x.groupby("date", observed=True)["ret250"].rank(method="average", pct=True) * 100.0
    ).astype("float32")

    x["cond_price"] = (x["close"] > x["ma20"]) & (x["close"] > x["ma60"])
    x["cond_cross"] = x["ma20"] > x["ma60"]
    x["cond_ma20_up"] = x["ma20_change"] > 0
    x["cond_ma60_up"] = x["ma60_change"] > 0
    x["cond_spread_up"] = x["spread_change"] > 0
    conds = ["cond_price", "cond_cross", "cond_ma20_up", "cond_ma60_up", "cond_spread_up"]
    x["complete"] = (x[conds].sum(axis=1).astype("int8") * 20).astype("int16")
    x["duck"] = x[conds].all(axis=1)
    x["duck_prev"] = x.groupby("code", observed=True)["duck"].shift(1).astype("boolean").fillna(False).astype(bool)
    x["new"] = x["duck"] & ~x["duck_prev"]
    x["duck_exit"] = ~x["duck"] & x["duck_prev"]

    # Current duck segment's overheat history (same concept as stock_duckbill.py).
    x["raw_hot"] = x["duck"] & (x["bias20"] >= 8.0)
    false_count = (~x["duck"]).groupby(x["code"], observed=True).cumsum()
    x["had_hot"] = x["raw_hot"].groupby([x["code"], false_count], observed=True).cummax() & x["duck"]

    # Heat score: copy the current dashboard definition.
    bias = pd.to_numeric(x["bias20"], errors="coerce").fillna(0.0)
    heat = (bias.clip(lower=0.0) * 5.0).clip(lower=0.0, upper=70.0).astype("float64")
    rs = pd.to_numeric(x["rs"], errors="coerce")
    heat += np.select([rs >= 95, rs >= 85], [12.0, 8.0], default=0.0)
    heat += np.where(x["duck"] | (x["complete"] >= 100), 8.0, 0.0)
    raw_hot = x["raw_hot"].to_numpy(dtype=bool)
    heat = np.where(raw_hot, np.maximum(70.0, heat + 18.0), heat)
    heat += np.where(x["had_hot"].to_numpy(dtype=bool), 8.0, 0.0)
    x["heat"] = np.clip(heat, 0.0, 100.0).astype("float32")

    # Right-side score: copy the current dashboard definition for technical rows.
    right = (x["complete"].astype("float64") * 0.55).clip(lower=0.0, upper=55.0)
    right += np.select([rs >= 85, rs >= 70, rs >= 50], [25.0, 18.0, 10.0], default=0.0)
    right += np.where(x["new"], 12.0, np.where(x["duck"], 8.0, 0.0))
    sync_up = (x["ma20_change"] > 0) & (x["ma60_change"] > 0) & (x["spread_change"] > 0)
    right += np.where(sync_up, 6.0, 0.0)
    right -= np.where(x["duck_exit"], 30.0, 0.0)
    x["right"] = np.clip(right, 0.0, 100.0).astype("float32")

    # Minimal data retained for trade simulation.
    keep = [
        "date", "code", "name", "market", "open", "high", "low", "close",
        "ma20", "ma60", "ma20_change", "ma60_change", "spread_change", "bias20",
        "rs", "complete", "duck", "new", "duck_exit", "raw_hot", "had_hot", "heat", "right",
    ]
    return x[keep].copy()


# ---------------------------------------------------------------------------
# Exit rule engine
# ---------------------------------------------------------------------------
STRATEGIES = [
    ("過熱分數≥70", "heat70", "過熱"),
    ("過熱分數≥80", "heat80", "過熱"),
    ("過熱分數≥85", "heat85", "過熱"),
    ("過熱分數≥90", "heat90", "過熱"),
    ("月線乖離≥8%", "bias8", "過熱"),
    ("RS曾≥85後跌破85", "rs_cross85", "RS"),
    ("RS自持有高點回落5", "rs_drop5", "RS"),
    ("RS自持有高點回落10", "rs_drop10", "RS"),
    ("MA20轉下", "ma20_down", "均線"),
    ("MA60轉下", "ma60_down", "均線"),
    ("MA20-MA60開口轉縮", "spread_down", "均線"),
    ("跌破MA20", "below_ma20", "價格"),
    ("跌破MA60", "below_ma60", "價格"),
    ("鴨嘴正式失效", "duck_exit", "鴨嘴"),
    ("右側曾≥70後跌破70", "right_cross70", "右側"),
    ("右側曾≥60後跌破60", "right_cross60", "右側"),
    ("右側曾≥50後跌破50", "right_cross50", "右側"),
    ("過熱70後開口轉縮", "hot70_spread", "組合"),
    ("過熱70後MA20轉下", "hot70_ma20", "組合"),
    ("過熱70後RS回落5", "hot70_rs5", "組合"),
    ("過熱70後跌破MA20", "hot70_below_ma20", "組合"),
    ("過熱70後至少2項轉弱", "hot70_2weak", "組合"),
]

BENCHMARKS = [
    ("固定持有20日", "hold20", "基準"),
    ("固定持有40日", "hold40", "基準"),
    ("固定持有60日", "hold60", "基準"),
]


def _exit_hit(key: str, row: pd.Series, state: dict, held_index: int) -> bool:
    heat = _num(row.get("heat"), 0.0) or 0.0
    rs = _num(row.get("rs"))
    right = _num(row.get("right"), 0.0) or 0.0
    ma20chg = _num(row.get("ma20_change"))
    ma60chg = _num(row.get("ma60_change"))
    spreadchg = _num(row.get("spread_change"))
    close = _num(row.get("close"))
    ma20 = _num(row.get("ma20"))
    ma60 = _num(row.get("ma60"))

    if rs is not None:
        state["peak_rs"] = max(_num(state.get("peak_rs"), rs) or rs, rs)
        if rs >= 85:
            state["rs85_armed"] = True
    state["peak_right"] = max(_num(state.get("peak_right"), right) or right, right)
    if right >= 70:
        state["right70_armed"] = True
    if right >= 60:
        state["right60_armed"] = True
    if right >= 50:
        state["right50_armed"] = True
    if heat >= 70:
        state["hot70_armed"] = True

    peak_rs = _num(state.get("peak_rs"))
    rs_drop5 = rs is not None and peak_rs is not None and rs <= peak_rs - 5.0
    rs_drop10 = rs is not None and peak_rs is not None and rs <= peak_rs - 10.0
    ma20_down = ma20chg is not None and ma20chg <= 0
    ma60_down = ma60chg is not None and ma60chg <= 0
    spread_down = spreadchg is not None and spreadchg <= 0
    below_ma20 = close is not None and ma20 is not None and close < ma20
    below_ma60 = close is not None and ma60 is not None and close < ma60
    hot_armed = bool(state.get("hot70_armed"))

    if key == "heat70": return heat >= 70
    if key == "heat80": return heat >= 80
    if key == "heat85": return heat >= 85
    if key == "heat90": return heat >= 90
    if key == "bias8": return (_num(row.get("bias20"), -999) or -999) >= 8
    if key == "rs_cross85": return bool(state.get("rs85_armed")) and rs is not None and rs < 85
    if key == "rs_drop5": return bool(rs_drop5)
    if key == "rs_drop10": return bool(rs_drop10)
    if key == "ma20_down": return bool(ma20_down)
    if key == "ma60_down": return bool(ma60_down)
    if key == "spread_down": return bool(spread_down)
    if key == "below_ma20": return bool(below_ma20)
    if key == "below_ma60": return bool(below_ma60)
    if key == "duck_exit": return not bool(row.get("duck", False))
    if key == "right_cross70": return bool(state.get("right70_armed")) and right < 70
    if key == "right_cross60": return bool(state.get("right60_armed")) and right < 60
    if key == "right_cross50": return bool(state.get("right50_armed")) and right < 50
    if key == "hot70_spread": return hot_armed and bool(spread_down)
    if key == "hot70_ma20": return hot_armed and bool(ma20_down)
    if key == "hot70_rs5": return hot_armed and bool(rs_drop5)
    if key == "hot70_below_ma20": return hot_armed and bool(below_ma20)
    if key == "hot70_2weak":
        weak = sum([
            bool(spread_down), bool(ma20_down), bool(rs_drop5), bool(below_ma20), right < 60,
        ])
        return hot_armed and weak >= 2
    if key == "hold20": return held_index >= 19
    if key == "hold40": return held_index >= 39
    if key == "hold60": return held_index >= 59
    return False


def _entry_cohorts(row: pd.Series) -> List[str]:
    if not bool(row.get("new", False)):
        return []
    out = ["全部鴨嘴新進"]
    heat = _num(row.get("heat"), 100) or 100
    right = _num(row.get("right"), 0) or 0
    rs = _num(row.get("rs"))
    if heat < 70 and right >= 58:
        out.append("新進且可交易")
    if heat < 70 and rs is not None and rs >= 85:
        out.append("新進+RS85+未過熱")
    return out


def simulate_trades(signals: pd.DataFrame, start: dt.date, end: dt.date, max_hold: int) -> Tuple[pd.DataFrame, dict]:
    groups = {str(code): g.sort_values("date").reset_index(drop=True) for code, g in signals.groupby("code", observed=True)}

    entries = []
    skipped_future = 0
    for code, g in groups.items():
        for i in np.flatnonzero(g["new"].to_numpy(dtype=bool)):
            sig_day = g.at[i, "date"].date()
            if sig_day < start or sig_day > end:
                continue
            cohorts = _entry_cohorts(g.iloc[i])
            if not cohorts:
                continue
            entry_idx = i + 1
            # Full common horizon is required for fair peak/giveback comparison.
            if entry_idx >= len(g) or entry_idx + max_hold >= len(g):
                skipped_future += 1
                continue
            entries.append((code, i, entry_idx, cohorts))

    print(f"[trade] entry events with full {max_hold}-session horizon={len(entries)}; skipped recent={skipped_future}", flush=True)
    rows: List[dict] = []
    benchmark_rules = [r for r in BENCHMARKS if int(r[1].replace("hold", "")) <= max_hold]
    if max_hold not in {20, 40, 60}:
        benchmark_rules.append((f"固定持有{max_hold}日", f"hold{max_hold}", "基準"))
    all_rules = STRATEGIES + benchmark_rules

    def first_hit(mask: np.ndarray) -> Optional[int]:
        idx = np.flatnonzero(mask)
        return int(idx[0]) if idx.size else None

    for n, (code, sig_idx, entry_idx, cohorts) in enumerate(entries, 1):
        g = groups[code]
        sig = g.iloc[sig_idx]
        ent = g.iloc[entry_idx]
        entry_price = _num(ent.get("open")) or _num(ent.get("close"))
        if entry_price is None or entry_price <= 0:
            continue

        horizon_end = entry_idx + max_hold - 1
        path = g.iloc[entry_idx:horizon_end + 1].copy()
        opens = pd.to_numeric(path["open"], errors="coerce").to_numpy(dtype=float)
        closes = pd.to_numeric(path["close"], errors="coerce").to_numpy(dtype=float)
        highs = pd.to_numeric(path["high"], errors="coerce").to_numpy(dtype=float)
        lows = pd.to_numeric(path["low"], errors="coerce").to_numpy(dtype=float)
        heat = pd.to_numeric(path["heat"], errors="coerce").fillna(-999).to_numpy(dtype=float)
        bias = pd.to_numeric(path["bias20"], errors="coerce").fillna(-999).to_numpy(dtype=float)
        rs = pd.to_numeric(path["rs"], errors="coerce").to_numpy(dtype=float)
        right = pd.to_numeric(path["right"], errors="coerce").fillna(0).to_numpy(dtype=float)
        ma20chg = pd.to_numeric(path["ma20_change"], errors="coerce").to_numpy(dtype=float)
        ma60chg = pd.to_numeric(path["ma60_change"], errors="coerce").to_numpy(dtype=float)
        spreadchg = pd.to_numeric(path["spread_change"], errors="coerce").to_numpy(dtype=float)
        ma20 = pd.to_numeric(path["ma20"], errors="coerce").to_numpy(dtype=float)
        ma60 = pd.to_numeric(path["ma60"], errors="coerce").to_numpy(dtype=float)
        duck = path["duck"].to_numpy(dtype=bool)

        rs_for_peak = np.where(np.isfinite(rs), rs, -np.inf)
        peak_rs = np.maximum.accumulate(rs_for_peak)
        rs85_arm = np.maximum.accumulate(rs_for_peak >= 85)
        right70_arm = np.maximum.accumulate(right >= 70)
        right60_arm = np.maximum.accumulate(right >= 60)
        right50_arm = np.maximum.accumulate(right >= 50)
        hot70_arm = np.maximum.accumulate(heat >= 70)

        rs_drop5 = np.isfinite(rs) & (rs <= peak_rs - 5.0)
        rs_drop10 = np.isfinite(rs) & (rs <= peak_rs - 10.0)
        ma20_down = np.isfinite(ma20chg) & (ma20chg <= 0)
        ma60_down = np.isfinite(ma60chg) & (ma60chg <= 0)
        spread_down = np.isfinite(spreadchg) & (spreadchg <= 0)
        below_ma20 = np.isfinite(ma20) & (closes < ma20)
        below_ma60 = np.isfinite(ma60) & (closes < ma60)
        weak_count = (
            spread_down.astype(int) + ma20_down.astype(int) + rs_drop5.astype(int)
            + below_ma20.astype(int) + (right < 60).astype(int)
        )

        masks: Dict[str, np.ndarray] = {
            "heat70": heat >= 70,
            "heat80": heat >= 80,
            "heat85": heat >= 85,
            "heat90": heat >= 90,
            "bias8": bias >= 8,
            "rs_cross85": rs85_arm & np.isfinite(rs) & (rs < 85),
            "rs_drop5": rs_drop5,
            "rs_drop10": rs_drop10,
            "ma20_down": ma20_down,
            "ma60_down": ma60_down,
            "spread_down": spread_down,
            "below_ma20": below_ma20,
            "below_ma60": below_ma60,
            "duck_exit": ~duck,
            "right_cross70": right70_arm & (right < 70),
            "right_cross60": right60_arm & (right < 60),
            "right_cross50": right50_arm & (right < 50),
            "hot70_spread": hot70_arm & spread_down,
            "hot70_ma20": hot70_arm & ma20_down,
            "hot70_rs5": hot70_arm & rs_drop5,
            "hot70_below_ma20": hot70_arm & below_ma20,
            "hot70_2weak": hot70_arm & (weak_count >= 2),
        }
        for fixed in sorted({int(r[1].replace("hold", "")) for r in benchmark_rules}):
            k = f"hold{fixed}"
            m = np.zeros(max_hold, dtype=bool)
            m[fixed - 1] = True
            masks[k] = m

        oracle_high = max(float(np.nanmax(highs)), float(entry_price))
        oracle_low = min(float(np.nanmin(lows)), float(entry_price))
        oracle_mfe = (oracle_high / entry_price - 1.0) * 100.0
        oracle_mae = (oracle_low / entry_price - 1.0) * 100.0
        peak_local = int(np.nanargmax(highs)) if np.isfinite(highs).any() else 0
        peak_day = path.iloc[peak_local]["date"].date().isoformat()
        days_to_peak = peak_local + 1

        for rule_name, key, rule_type in all_rules:
            offset = first_hit(masks[key])
            forced = offset is None
            if forced:
                offset = max_hold - 1
            hit_idx = entry_idx + int(offset)
            exit_exec_idx = hit_idx + 1
            if exit_exec_idx >= len(g):
                continue
            hit = g.iloc[hit_idx]
            outrow = g.iloc[exit_exec_idx]
            exit_price = _num(outrow.get("open")) or _num(outrow.get("close"))
            if exit_price is None or exit_price <= 0:
                continue

            upto = int(offset) + 1
            local_high = max(float(np.nanmax(highs[:upto])), float(exit_price), float(entry_price))
            local_low = min(float(np.nanmin(lows[:upto])), float(exit_price), float(entry_price))
            ret = (exit_price / entry_price - 1.0) * 100.0
            mfe_to_exit = (local_high / entry_price - 1.0) * 100.0
            mae_to_exit = (local_low / entry_price - 1.0) * 100.0
            giveback60 = oracle_mfe - ret
            capture60 = (ret / oracle_mfe * 100.0) if oracle_mfe > 0.001 else np.nan
            if pd.notna(capture60):
                capture60 = float(np.clip(capture60, -300, 120))

            post5 = post10 = np.nan
            if exit_exec_idx + 5 < len(g):
                p5 = _num(g.iloc[exit_exec_idx + 5].get("close"))
                if p5 is not None: post5 = (p5 / exit_price - 1.0) * 100.0
            if exit_exec_idx + 10 < len(g):
                p10 = _num(g.iloc[exit_exec_idx + 10].get("close"))
                if p10 is not None: post10 = (p10 / exit_price - 1.0) * 100.0

            base = {
                "代號": code,
                "名稱": str(ent.get("name") or sig.get("name") or ""),
                "市場": str(ent.get("market") or sig.get("market") or ""),
                "進場訊號日": sig["date"].date().isoformat(),
                "進場日": ent["date"].date().isoformat(),
                "進場價": round(float(entry_price), 4),
                "進場RS": round(_num(sig.get("rs"), np.nan), 2) if _num(sig.get("rs")) is not None else np.nan,
                "進場右側": round(_num(sig.get("right"), np.nan), 2) if _num(sig.get("right")) is not None else np.nan,
                "進場過熱": round(_num(sig.get("heat"), np.nan), 2) if _num(sig.get("heat")) is not None else np.nan,
                "策略": rule_name,
                "策略代碼": key,
                "類型": rule_type,
                "出場訊號日": hit["date"].date().isoformat(),
                "出場日": outrow["date"].date().isoformat(),
                "出場價": round(float(exit_price), 4),
                "持有交易日": int(exit_exec_idx - entry_idx),
                "策略有觸發": not forced,
                "報酬%": round(float(ret), 4),
                "出場前MFE%": round(float(mfe_to_exit), 4),
                "出場前MAE%": round(float(mae_to_exit), 4),
                f"{max_hold}日最佳可能漲幅%": round(float(oracle_mfe), 4),
                f"{max_hold}日最大不利%": round(float(oracle_mae), 4),
                f"{max_hold}日峰值捕捉率%": round(float(capture60), 4) if pd.notna(capture60) else np.nan,
                f"相對{max_hold}日高點回吐%": round(float(giveback60), 4),
                "固定窗高點日": peak_day,
                "進場後高點第幾日": int(days_to_peak),
                "出場後5日%": round(float(post5), 4) if pd.notna(post5) else np.nan,
                "出場後10日%": round(float(post10), 4) if pd.notna(post10) else np.nan,
            }
            base["群組_可交易"] = "新進且可交易" in cohorts
            base["群組_RS85未過熱"] = "新進+RS85+未過熱" in cohorts
            rows.append(base)

        if n == 1 or n % 250 == 0 or n == len(entries):
            print(f"[trade] {n}/{len(entries)} entry events simulated", flush=True)

    trades = pd.DataFrame(rows)
    return trades, {"entry_events": len(entries), "skipped_recent_no_full_horizon": skipped_future}


# ---------------------------------------------------------------------------
# Summary / ranking
# ---------------------------------------------------------------------------
def _cohort_views(trades: pd.DataFrame):
    yield "全部鴨嘴新進", trades
    if "群組_可交易" in trades.columns:
        yield "新進且可交易", trades[trades["群組_可交易"].fillna(False).astype(bool)]
    if "群組_RS85未過熱" in trades.columns:
        yield "新進+RS85+未過熱", trades[trades["群組_RS85未過熱"].fillna(False).astype(bool)]


def summarize_trades(trades: pd.DataFrame, max_hold: int) -> pd.DataFrame:
    if trades.empty:
        return pd.DataFrame()
    peak_col = f"{max_hold}日峰值捕捉率%"
    give_col = f"相對{max_hold}日高點回吐%"
    rows = []
    for cohort, view in _cohort_views(trades):
        if view.empty:
            continue
        for (strategy, key, kind), g in view.groupby(["策略", "策略代碼", "類型"], dropna=False):
            ret = pd.to_numeric(g["報酬%"], errors="coerce")
            cap = pd.to_numeric(g[peak_col], errors="coerce")
            give = pd.to_numeric(g[give_col], errors="coerce")
            mae = pd.to_numeric(g["出場前MAE%"], errors="coerce")
            p5 = pd.to_numeric(g["出場後5日%"], errors="coerce")
            p10 = pd.to_numeric(g["出場後10日%"], errors="coerce")
            hold = pd.to_numeric(g["持有交易日"], errors="coerce")
            trigger = g["策略有觸發"].astype(bool)
            rows.append({
                "進場群組": cohort,
                "策略": strategy,
                "策略代碼": key,
                "類型": kind,
                "樣本數": int(ret.notna().sum()),
                "有效觸發率%": round(float(trigger.mean() * 100.0), 2),
                "平均報酬%": round(float(ret.mean()), 3),
                "中位報酬%": round(float(ret.median()), 3),
                "勝率%": round(float((ret > 0).mean() * 100.0), 2),
                "平均持有日": round(float(hold.mean()), 2),
                "中位持有日": round(float(hold.median()), 1),
                "平均峰值捕捉率%": round(float(cap.mean()), 2),
                "中位峰值捕捉率%": round(float(cap.median()), 2),
                "平均高點回吐%": round(float(give.mean()), 3),
                "中位高點回吐%": round(float(give.median()), 3),
                "平均MAE%": round(float(mae.mean()), 3),
                "平均出場後5日%": round(float(p5.mean()), 3),
                "平均出場後10日%": round(float(p10.mean()), 3),
            })
    out = pd.DataFrame(rows)
    if out.empty:
        return out

    # Transparent composite score. Benchmarks are shown but not given a recommendation rank.
    out["綜合分數"] = np.nan
    out["群組排名"] = np.nan
    for cohort, idx in out.groupby("進場群組").groups.items():
        sub = out.loc[idx].copy()
        cand = sub["類型"] != "基準"
        cidx = sub.index[cand]
        if len(cidx) == 0:
            continue
        def pct_rank(series: pd.Series, higher=True):
            ss = pd.to_numeric(series, errors="coerce")
            if not higher:
                ss = -ss
            return ss.rank(pct=True, method="average").fillna(0.0) * 100.0
        comp = (
            0.30 * pct_rank(out.loc[cidx, "中位報酬%"], True)
            + 0.25 * pct_rank(out.loc[cidx, "中位峰值捕捉率%"], True)
            + 0.20 * pct_rank(out.loc[cidx, "平均高點回吐%"], False)
            + 0.15 * pct_rank(out.loc[cidx, "平均出場後10日%"], False)
            + 0.10 * pct_rank(out.loc[cidx, "平均MAE%"], True)
        )
        trig = pd.to_numeric(out.loc[cidx, "有效觸發率%"], errors="coerce").fillna(0)
        reliability_factor = np.where(trig >= 40, 1.0, np.where(trig >= 20, 0.92, 0.82))
        comp = comp * reliability_factor
        out.loc[cidx, "綜合分數"] = comp.round(2)
        eligible = cidx[pd.to_numeric(out.loc[cidx, "有效觸發率%"], errors="coerce").fillna(0).to_numpy() >= 25.0]
        if len(eligible):
            out.loc[eligible, "群組排名"] = out.loc[eligible, "綜合分數"].rank(ascending=False, method="min").astype(float)

    out["是否進入綜合排名"] = np.where(
        (out["類型"] != "基準") & (pd.to_numeric(out["有效觸發率%"], errors="coerce").fillna(0) >= 25.0),
        "是", "否"
    )

    def reliability(r):
        n = int(r.get("樣本數", 0) or 0)
        tr = float(r.get("有效觸發率%", 0) or 0)
        if n >= 150 and tr >= 40: return "高"
        if n >= 70 and tr >= 25: return "中"
        return "需擴充樣本"
    out["可靠度"] = out.apply(reliability, axis=1)
    return out.sort_values(["進場群組", "類型", "群組排名", "綜合分數"], ascending=[True, True, True, False], na_position="last")


def annual_summary(trades: pd.DataFrame, max_hold: int) -> pd.DataFrame:
    if trades.empty:
        return pd.DataFrame()
    peak_col = f"{max_hold}日峰值捕捉率%"
    give_col = f"相對{max_hold}日高點回吐%"
    parts = []
    for cohort, view in _cohort_views(trades):
        if view.empty:
            continue
        x = view.copy()
        x["年度"] = pd.to_datetime(x["進場日"], errors="coerce").dt.year
        q = (
            x.groupby(["策略", "類型", "年度"], dropna=False)
            .agg(
                樣本數=("報酬%", "count"),
                中位報酬=("報酬%", "median"),
                勝率=("報酬%", lambda ss: float((pd.to_numeric(ss, errors="coerce") > 0).mean() * 100.0)),
                中位峰值捕捉=(peak_col, "median"),
                平均高點回吐=(give_col, "mean"),
                平均持有日=("持有交易日", "mean"),
            )
            .reset_index()
        )
        q.insert(0, "進場群組", cohort)
        parts.append(q)
    if not parts:
        return pd.DataFrame()
    out = pd.concat(parts, ignore_index=True)
    return out.rename(columns={
        "中位報酬": "中位報酬%",
        "勝率": "勝率%",
        "中位峰值捕捉": "中位峰值捕捉率%",
        "平均高點回吐": "平均高點回吐%",
    }).round(3)


def best_by_metric(summary: pd.DataFrame) -> pd.DataFrame:
    if summary.empty:
        return pd.DataFrame()
    rows = []
    base = summary[(summary["類型"] != "基準") & (pd.to_numeric(summary["有效觸發率%"], errors="coerce").fillna(0) >= 25.0)].copy()
    for cohort, g in base.groupby("進場群組"):
        if g.empty:
            continue
        specs = [
            ("綜合第一", "綜合分數", True),
            ("中位報酬最高", "中位報酬%", True),
            ("峰值捕捉最好", "中位峰值捕捉率%", True),
            ("高點回吐最小", "平均高點回吐%", False),
            ("出場後10日最弱", "平均出場後10日%", False),
            ("持有期風險最低", "平均MAE%", True),
        ]
        for label, col, higher in specs:
            z = g.dropna(subset=[col])
            if z.empty:
                continue
            r = z.loc[z[col].idxmax() if higher else z[col].idxmin()]
            rows.append({
                "進場群組": cohort,
                "評比": label,
                "策略": r["策略"],
                "數值": r[col],
                "樣本數": r["樣本數"],
                "有效觸發率%": r["有效觸發率%"],
                "可靠度": r["可靠度"],
            })
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Heat-event study (supplementary)
# ---------------------------------------------------------------------------
def heat_event_study(signals: pd.DataFrame, start: dt.date, end: dt.date) -> Tuple[pd.DataFrame, pd.DataFrame]:
    events = []
    for code, g in signals.groupby("code", observed=True):
        g = g.sort_values("date").reset_index(drop=True)
        starts = np.flatnonzero(g["new"].to_numpy(dtype=bool))
        for sidx in starts:
            sd = g.at[sidx, "date"].date()
            if sd < start or sd > end:
                continue
            # Current duck segment ends at the first non-duck row after start.
            eidx = sidx
            while eidx + 1 < len(g) and bool(g.at[eidx + 1, "duck"]):
                eidx += 1
            seg = g.iloc[sidx:eidx + 1]
            if seg.empty:
                continue
            for th in [70, 80, 85, 90]:
                hits = seg.index[pd.to_numeric(seg["heat"], errors="coerce") >= th].tolist()
                if not hits:
                    continue
                first_idx = hits[0]
                first_pos = first_idx - sidx
                # Consecutive threshold duration from first hit.
                duration = 0
                j = first_idx
                while j <= eidx and _num(g.at[j, "heat"], -1) >= th:
                    duration += 1
                    j += 1
                tail = g.iloc[first_idx:eidx + 1]
                highs = pd.to_numeric(tail["high"], errors="coerce")
                if highs.dropna().empty:
                    continue
                peak_rel = int(np.nanargmax(highs.to_numpy(dtype=float)))
                peak_idx = first_idx + peak_rel
                first_close = _num(g.at[first_idx, "close"])
                peak_high = _num(g.at[peak_idx, "high"])
                end_close = _num(g.at[eidx, "close"])
                if first_close is None or peak_high is None:
                    continue
                upside = (peak_high / first_close - 1.0) * 100.0
                end_ret = (end_close / first_close - 1.0) * 100.0 if end_close is not None else np.nan
                events.append({
                    "代號": str(code),
                    "名稱": str(g.at[sidx, "name"] or ""),
                    "鴨嘴新進日": sd.isoformat(),
                    "門檻": f"過熱≥{th}",
                    "首次達標日": g.at[first_idx, "date"].date().isoformat(),
                    "新進後第幾日達標": int(first_pos + 1),
                    "首次達標後連續天數": int(duration),
                    "首次達標後高點第幾日": int(peak_rel + 1),
                    "首次達標後至段內高點漲幅%": round(float(upside), 4),
                    "首次達標至鴨嘴段末報酬%": round(float(end_ret), 4) if pd.notna(end_ret) else np.nan,
                    "鴨嘴段總天數": int(eidx - sidx + 1),
                })
    detail = pd.DataFrame(events)
    if detail.empty:
        return detail, pd.DataFrame()
    summary = (
        detail.groupby("門檻")
        .agg(
            事件數=("代號", "count"),
            中位達標日=("新進後第幾日達標", "median"),
            中位連續天數=("首次達標後連續天數", "median"),
            中位高點日在達標後第幾日=("首次達標後高點第幾日", "median"),
            高點日25百分位=("首次達標後高點第幾日", lambda s: float(pd.Series(s).quantile(0.25))),
            高點日75百分位=("首次達標後高點第幾日", lambda s: float(pd.Series(s).quantile(0.75))),
            平均達標後續漲幅=("首次達標後至段內高點漲幅%", "mean"),
            中位達標後續漲幅=("首次達標後至段內高點漲幅%", "median"),
            中位至鴨嘴段末報酬=("首次達標至鴨嘴段末報酬%", "median"),
        )
        .reset_index()
        .rename(columns={
            "平均達標後續漲幅": "平均達標後續漲幅%",
            "中位達標後續漲幅": "中位達標後續漲幅%",
            "中位至鴨嘴段末報酬": "中位至鴨嘴段末報酬%",
        })
        .round(3)
    )
    return detail, summary


# ---------------------------------------------------------------------------
# Excel / status output
# ---------------------------------------------------------------------------
def write_outputs(
    summary: pd.DataFrame,
    best: pd.DataFrame,
    annual: pd.DataFrame,
    heat_summary: pd.DataFrame,
    heat_detail: pd.DataFrame,
    trades: pd.DataFrame,
    status: dict,
) -> None:
    summary.to_csv(OUT_SUMMARY, index=False, encoding="utf-8-sig")
    OUT_STATUS.write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")

    notes = pd.DataFrame({
        "說明": [
            "第一階段只比較可逐日、Point-in-time 重建的技術/RS/右側/過熱出場條件；歷史基本面與估值不硬塞入回測。",
            "鴨嘴條件：收盤>MA20/MA60、MA20>MA60、MA20上揚、MA60上揚、MA20-MA60開口擴大。",
            "RS：250交易日報酬在同一交易日全市場做百分位排名。",
            "進場：T日收盤確認『鴨嘴今日新進』，T+1日開盤成交；出場亦為T日收盤確認、T+1日開盤成交，避免收盤訊號偷看未來。",
            "峰值捕捉率與高點回吐都使用相同的最大持有窗，避免不同策略因觀察期間不同而失真。",
            "綜合分數：30%中位報酬、25%峰值捕捉、20%低回吐、15%出場後10日不再續漲、10%低MAE；觸發率太低會小幅折扣。",
            "綜合分數只用來排序研究優先級，不代表已找到永遠有效的最佳參數；需再看年度穩定度與較長歷史。",
            "最近進場事件若沒有完整最大持有窗，會排除在公平比較之外。",
        ]
    })

    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
        book = writer.book
        hdr = book.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center"})
        good = book.add_format({"bg_color": "#E2F0D9"})
        warn = book.add_format({"bg_color": "#FFF2CC"})
        hot = book.add_format({"bg_color": "#F4CCCC"})

        sheets = [
            ("出場策略排行榜", summary),
            ("各指標最佳", best),
            ("年度穩定度", annual),
            ("過熱事件統計", heat_summary),
            ("過熱事件明細", heat_detail),
            ("交易明細", trades),
            ("說明", notes),
        ]
        for name, df in sheets:
            data = df if df is not None else pd.DataFrame()
            data.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.freeze_panes(1, 0)
            if len(data.columns):
                ws.autofilter(0, 0, max(0, len(data)), len(data.columns) - 1)
            for c, col in enumerate(data.columns):
                ws.write(0, c, col, hdr)
                width = min(34, max(11, len(str(col)) * 2 + 3))
                if col in {"策略", "評比", "進場群組"}: width = max(width, 22)
                ws.set_column(c, c, width)
            if name == "出場策略排行榜" and not data.empty:
                if "群組排名" in data.columns:
                    c = data.columns.get_loc("群組排名")
                    ws.conditional_format(1, c, len(data), c, {"type": "cell", "criteria": "==", "value": 1, "format": good})
                if "有效觸發率%" in data.columns:
                    c = data.columns.get_loc("有效觸發率%")
                    ws.conditional_format(1, c, len(data), c, {"type": "cell", "criteria": "<", "value": 20, "format": warn})
                if "平均高點回吐%" in data.columns:
                    c = data.columns.get_loc("平均高點回吐%")
                    ws.conditional_format(1, c, len(data), c, {"type": "cell", "criteria": ">", "value": 12, "format": hot})
            if name == "說明":
                ws.set_column(0, 0, 110)


def run(start: dt.date, end: dt.date, max_hold: int, fetch_missing: bool, workers: int) -> int:
    warmup_start = start - dt.timedelta(days=WARMUP_CALENDAR_DAYS)
    print("=== 個股出場策略回測 ===", flush=True)
    print(f"正式研究區間: {start} ~ {end}", flush=True)
    print(f"RS/均線暖機: {warmup_start} 起", flush=True)
    print(f"最大共同持有窗: {max_hold} 交易日", flush=True)

    cache_report = ensure_history_cache(warmup_start, end, fetch_missing=fetch_missing, workers=workers)
    raw, cache_meta = load_market_cache(warmup_start, end)
    print(f"[data] {cache_meta}", flush=True)
    signals = compute_signals(raw)
    valid_rs = signals[(signals["date"].dt.date >= start) & signals["rs"].notna()]
    print(
        f"[signal] rows={len(signals):,}; research RS rows={len(valid_rs):,}; "
        f"duck new={int(signals[(signals['date'].dt.date>=start)&(signals['date'].dt.date<=end)]['new'].sum())}",
        flush=True,
    )

    trades, trade_meta = simulate_trades(signals, start, end, max_hold=max_hold)
    summary = summarize_trades(trades, max_hold=max_hold)
    annual = annual_summary(trades, max_hold=max_hold)
    best = best_by_metric(summary)
    heat_detail, heat_summary = heat_event_study(signals, start, end)

    status = {
        "status": "success" if not summary.empty else "no_results",
        "generated_at_taipei": dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).strftime("%Y-%m-%d %H:%M:%S"),
        "research_start": start.isoformat(),
        "research_end": end.isoformat(),
        "warmup_start": warmup_start.isoformat(),
        "max_hold_sessions": int(max_hold),
        "cache": cache_report,
        "data": cache_meta,
        "trades": trade_meta,
        "trade_rows": int(len(trades)),
        "summary_rows": int(len(summary)),
        "heat_events": int(len(heat_detail)),
        "method": "T close signal -> T+1 open execution; technical/RS point-in-time reconstruction",
        "outputs": [OUT_XLSX.name, OUT_SUMMARY.name, OUT_STATUS.name],
    }
    if not summary.empty:
        leaders = (
            summary[(summary["類型"] != "基準") & (summary["群組排名"] == 1)]
            [["進場群組", "策略", "綜合分數", "中位報酬%", "中位峰值捕捉率%", "平均高點回吐%", "樣本數", "有效觸發率%"]]
            .to_dict("records")
        )
        status["current_leaders"] = leaders
    write_outputs(summary, best, annual, heat_summary, heat_detail, trades, status)
    print(f"[done] {OUT_XLSX}", flush=True)
    print(f"[done] {OUT_SUMMARY}", flush=True)
    print(f"[done] {OUT_STATUS}", flush=True)
    if status.get("current_leaders"):
        print("[leaders]", json.dumps(status["current_leaders"], ensure_ascii=False), flush=True)
    return 0 if status["status"] == "success" else 2


def main() -> int:
    ap = argparse.ArgumentParser(description="Backtest exit indicators for Duckbill/RS stock entries")
    ap.add_argument("--start", default=DEFAULT_START.isoformat(), help="research start YYYY-MM-DD")
    ap.add_argument("--end", default="", help="research end YYYY-MM-DD; blank=latest formal update")
    ap.add_argument("--max-hold", type=int, default=DEFAULT_MAX_HOLD, help="common future window / max hold sessions")
    ap.add_argument("--fetch-missing", action="store_true", help="download missing official daily cache files")
    ap.add_argument("--workers", type=int, default=2, help="official history fetch workers (1-4)")
    args = ap.parse_args()
    start = _date(args.start)
    end = _date(args.end) if str(args.end).strip() else _latest_formal_date()
    if end <= start:
        raise SystemExit("end date must be after start date")
    return run(start, end, max(20, min(120, int(args.max_hold))), bool(args.fetch_missing), args.workers)


if __name__ == "__main__":
    raise SystemExit(main())
