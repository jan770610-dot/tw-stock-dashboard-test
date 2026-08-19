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
  - exit_backtest_validation.csv
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
OUT_VALIDATION = ROOT / "exit_backtest_validation.csv"
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
    """Optimized event simulator.

    v1 created one Python dict for every event x rule while repeatedly slicing
    pandas DataFrames. With ~40k duck-new events and ~25 rules that expands to
    >1M strategy-trade rows and is unnecessarily slow.  v1.1 converts each
    stock to NumPy arrays once, evaluates every rule for an event as a compact
    boolean matrix, stores numeric results in preallocated arrays, then builds
    one compact DataFrame at the end.
    """
    benchmark_rules = [r for r in BENCHMARKS if int(r[1].replace("hold", "")) <= max_hold]
    if max_hold not in {20, 40, 60}:
        benchmark_rules.append((f"固定持有{max_hold}日", f"hold{max_hold}", "基準"))
    all_rules = STRATEGIES + benchmark_rules
    rule_names = np.array([r[0] for r in all_rules], dtype=object)
    rule_keys = np.array([r[1] for r in all_rules], dtype=object)
    rule_types = np.array([r[2] for r in all_rules], dtype=object)
    nr = len(all_rules)

    stock_blocks: Dict[str, dict] = {}
    entries: List[Tuple[str, int, int, bool, bool]] = []
    skipped_future = 0

    needed_num = [
        "open", "high", "low", "close", "heat", "bias20", "rs", "right",
        "ma20_change", "ma60_change", "spread_change", "ma20", "ma60",
    ]
    for code, g in signals.groupby("code", observed=True, sort=False):
        g = g.sort_values("date").reset_index(drop=True)
        block = {
            "date": pd.to_datetime(g["date"], errors="coerce").to_numpy(dtype="datetime64[ns]"),
            "duck": g["duck"].fillna(False).to_numpy(dtype=bool),
            "new": g["new"].fillna(False).to_numpy(dtype=bool),
            "name": g.get("name", pd.Series("", index=g.index)).fillna("").astype(str).to_numpy(dtype=object),
            "market": g.get("market", pd.Series("", index=g.index)).fillna("").astype(str).to_numpy(dtype=object),
        }
        for c in needed_num:
            block[c] = pd.to_numeric(g[c], errors="coerce").to_numpy(dtype=np.float32)
        stock_blocks[str(code)] = block

        for i in np.flatnonzero(block["new"]):
            sig_day = pd.Timestamp(block["date"][i]).date()
            if sig_day < start or sig_day > end:
                continue
            heat0 = _num(block["heat"][i], 100) or 100
            right0 = _num(block["right"][i], 0) or 0
            rs0 = _num(block["rs"][i])
            cohort_tradeable = bool(heat0 < 70 and right0 >= 58)
            cohort_rs85 = bool(heat0 < 70 and rs0 is not None and rs0 >= 85)
            entry_idx = int(i) + 1
            # Need max_hold signal days plus the T+1 exit execution row.
            if entry_idx >= len(block["close"]) or entry_idx + max_hold >= len(block["close"]):
                skipped_future += 1
                continue
            entries.append((str(code), int(i), entry_idx, cohort_tradeable, cohort_rs85))

    n_events = len(entries)
    print(f"[trade] entry events with full {max_hold}-session horizon={n_events}; skipped recent={skipped_future}", flush=True)
    if n_events == 0:
        return pd.DataFrame(), {"entry_events": 0, "skipped_recent_no_full_horizon": skipped_future, "rule_count": nr}

    shape = (n_events, nr)
    f32 = lambda: np.full(shape, np.nan, dtype=np.float32)
    ret_mat = f32(); cap_mat = f32(); give_mat = f32(); mae_mat = f32(); post5_mat = f32(); post10_mat = f32()
    exit_price_mat = f32()
    hold_mat = np.zeros(shape, dtype=np.int16)
    trigger_mat = np.zeros(shape, dtype=bool)
    exit_sig_dates = np.full(shape, np.datetime64("NaT"), dtype="datetime64[ns]")
    exit_dates = np.full(shape, np.datetime64("NaT"), dtype="datetime64[ns]")

    event_code = np.empty(n_events, dtype=object)
    event_name = np.empty(n_events, dtype=object)
    event_market = np.empty(n_events, dtype=object)
    sig_dates = np.empty(n_events, dtype="datetime64[ns]")
    ent_dates = np.empty(n_events, dtype="datetime64[ns]")
    entry_price_arr = np.full(n_events, np.nan, dtype=np.float32)
    entry_rs_arr = np.full(n_events, np.nan, dtype=np.float32)
    entry_right_arr = np.full(n_events, np.nan, dtype=np.float32)
    entry_heat_arr = np.full(n_events, np.nan, dtype=np.float32)
    cohort_tradeable_arr = np.zeros(n_events, dtype=bool)
    cohort_rs85_arr = np.zeros(n_events, dtype=bool)

    strategy_key_to_col = {k: j for j, k in enumerate(rule_keys.tolist())}

    for n, (code, sig_idx, entry_idx, cohort_tradeable, cohort_rs85) in enumerate(entries):
        b = stock_blocks[code]
        opens_all = b["open"]; closes_all = b["close"]
        entry_price = _num(opens_all[entry_idx]) or _num(closes_all[entry_idx])
        if entry_price is None or entry_price <= 0:
            continue
        entry_price = float(entry_price)

        sl = slice(entry_idx, entry_idx + max_hold)
        closes = b["close"][sl].astype(np.float64, copy=False)
        highs = b["high"][sl].astype(np.float64, copy=False)
        lows = b["low"][sl].astype(np.float64, copy=False)
        heat = b["heat"][sl].astype(np.float64, copy=False)
        bias = b["bias20"][sl].astype(np.float64, copy=False)
        rs = b["rs"][sl].astype(np.float64, copy=False)
        right = b["right"][sl].astype(np.float64, copy=False)
        ma20chg = b["ma20_change"][sl].astype(np.float64, copy=False)
        ma60chg = b["ma60_change"][sl].astype(np.float64, copy=False)
        spreadchg = b["spread_change"][sl].astype(np.float64, copy=False)
        ma20 = b["ma20"][sl].astype(np.float64, copy=False)
        ma60 = b["ma60"][sl].astype(np.float64, copy=False)
        duck = b["duck"][sl]

        # Replace missing OHLC defensively with close/entry values for path metrics.
        safe_close = np.where(np.isfinite(closes) & (closes > 0), closes, entry_price)
        safe_high = np.where(np.isfinite(highs) & (highs > 0), highs, safe_close)
        safe_low = np.where(np.isfinite(lows) & (lows > 0), lows, safe_close)

        rs_for_peak = np.where(np.isfinite(rs), rs, -np.inf)
        peak_rs = np.maximum.accumulate(rs_for_peak)
        rs85_arm = np.maximum.accumulate(rs_for_peak >= 85)
        right70_arm = np.maximum.accumulate(np.where(np.isfinite(right), right, 0) >= 70)
        right60_arm = np.maximum.accumulate(np.where(np.isfinite(right), right, 0) >= 60)
        right50_arm = np.maximum.accumulate(np.where(np.isfinite(right), right, 0) >= 50)
        hot70_arm = np.maximum.accumulate(np.where(np.isfinite(heat), heat, -999) >= 70)

        rs_drop5 = np.isfinite(rs) & (rs <= peak_rs - 5.0)
        rs_drop10 = np.isfinite(rs) & (rs <= peak_rs - 10.0)
        ma20_down = np.isfinite(ma20chg) & (ma20chg <= 0)
        ma60_down = np.isfinite(ma60chg) & (ma60chg <= 0)
        spread_down = np.isfinite(spreadchg) & (spreadchg <= 0)
        below_ma20 = np.isfinite(ma20) & (safe_close < ma20)
        below_ma60 = np.isfinite(ma60) & (safe_close < ma60)
        weak_count = (
            spread_down.astype(np.int8) + ma20_down.astype(np.int8) + rs_drop5.astype(np.int8)
            + below_ma20.astype(np.int8) + (np.where(np.isfinite(right), right, 0) < 60).astype(np.int8)
        )

        masks = np.zeros((max_hold, nr), dtype=bool)
        base_masks = {
            "heat70": np.where(np.isfinite(heat), heat, -999) >= 70,
            "heat80": np.where(np.isfinite(heat), heat, -999) >= 80,
            "heat85": np.where(np.isfinite(heat), heat, -999) >= 85,
            "heat90": np.where(np.isfinite(heat), heat, -999) >= 90,
            "bias8": np.where(np.isfinite(bias), bias, -999) >= 8,
            "rs_cross85": rs85_arm & np.isfinite(rs) & (rs < 85),
            "rs_drop5": rs_drop5,
            "rs_drop10": rs_drop10,
            "ma20_down": ma20_down,
            "ma60_down": ma60_down,
            "spread_down": spread_down,
            "below_ma20": below_ma20,
            "below_ma60": below_ma60,
            "duck_exit": ~duck,
            "right_cross70": right70_arm & (np.where(np.isfinite(right), right, 0) < 70),
            "right_cross60": right60_arm & (np.where(np.isfinite(right), right, 0) < 60),
            "right_cross50": right50_arm & (np.where(np.isfinite(right), right, 0) < 50),
            "hot70_spread": hot70_arm & spread_down,
            "hot70_ma20": hot70_arm & ma20_down,
            "hot70_rs5": hot70_arm & rs_drop5,
            "hot70_below_ma20": hot70_arm & below_ma20,
            "hot70_2weak": hot70_arm & (weak_count >= 2),
        }
        for k, m in base_masks.items():
            j = strategy_key_to_col[k]
            masks[:, j] = m
        for fixed in sorted({int(r[1].replace("hold", "")) for r in benchmark_rules}):
            j = strategy_key_to_col[f"hold{fixed}"]
            masks[fixed - 1, j] = True

        has_hit = masks.any(axis=0)
        offsets = masks.argmax(axis=0).astype(np.int16)
        offsets[~has_hit] = max_hold - 1
        hit_idx = entry_idx + offsets.astype(np.int64)
        exit_idx = hit_idx + 1

        exit_open = opens_all[exit_idx].astype(np.float64, copy=False)
        exit_close = closes_all[exit_idx].astype(np.float64, copy=False)
        exit_price = np.where(np.isfinite(exit_open) & (exit_open > 0), exit_open, exit_close)
        valid_exit = np.isfinite(exit_price) & (exit_price > 0)

        cum_high = np.maximum.accumulate(np.maximum(safe_high, entry_price))
        cum_low = np.minimum.accumulate(np.minimum(safe_low, entry_price))
        local_high = np.maximum(cum_high[offsets], np.where(valid_exit, exit_price, entry_price))
        local_low = np.minimum(cum_low[offsets], np.where(valid_exit, exit_price, entry_price))
        oracle_high = max(float(cum_high[-1]), entry_price)
        oracle_mfe = (oracle_high / entry_price - 1.0) * 100.0

        ret = np.full(nr, np.nan, dtype=np.float64)
        ret[valid_exit] = (exit_price[valid_exit] / entry_price - 1.0) * 100.0
        mae = (local_low / entry_price - 1.0) * 100.0
        give = oracle_mfe - ret
        cap = np.full(nr, np.nan, dtype=np.float64)
        if oracle_mfe > 0.001:
            cap[valid_exit] = np.clip(ret[valid_exit] / oracle_mfe * 100.0, -300.0, 120.0)

        post5 = np.full(nr, np.nan, dtype=np.float64)
        post10 = np.full(nr, np.nan, dtype=np.float64)
        idx5 = exit_idx + 5
        ok5 = valid_exit & (idx5 < len(closes_all))
        if ok5.any():
            p5 = closes_all[idx5[ok5]].astype(np.float64, copy=False)
            good = np.isfinite(p5) & (p5 > 0)
            tmp = np.flatnonzero(ok5)
            post5[tmp[good]] = (p5[good] / exit_price[tmp[good]] - 1.0) * 100.0
        idx10 = exit_idx + 10
        ok10 = valid_exit & (idx10 < len(closes_all))
        if ok10.any():
            p10 = closes_all[idx10[ok10]].astype(np.float64, copy=False)
            good = np.isfinite(p10) & (p10 > 0)
            tmp = np.flatnonzero(ok10)
            post10[tmp[good]] = (p10[good] / exit_price[tmp[good]] - 1.0) * 100.0

        ret_mat[n] = ret.astype(np.float32)
        cap_mat[n] = cap.astype(np.float32)
        give_mat[n] = give.astype(np.float32)
        mae_mat[n] = mae.astype(np.float32)
        post5_mat[n] = post5.astype(np.float32)
        post10_mat[n] = post10.astype(np.float32)
        exit_price_mat[n] = np.where(valid_exit, exit_price, np.nan).astype(np.float32)
        hold_mat[n] = (exit_idx - entry_idx).astype(np.int16)
        trigger_mat[n] = has_hit
        exit_sig_dates[n] = b["date"][hit_idx]
        exit_dates[n] = b["date"][exit_idx]

        event_code[n] = code
        event_name[n] = str(b["name"][sig_idx] or "")
        event_market[n] = str(b["market"][sig_idx] or "")
        sig_dates[n] = b["date"][sig_idx]
        ent_dates[n] = b["date"][entry_idx]
        entry_price_arr[n] = np.float32(entry_price)
        entry_rs_arr[n] = np.float32(b["rs"][sig_idx]) if np.isfinite(b["rs"][sig_idx]) else np.nan
        entry_right_arr[n] = np.float32(b["right"][sig_idx]) if np.isfinite(b["right"][sig_idx]) else np.nan
        entry_heat_arr[n] = np.float32(b["heat"][sig_idx]) if np.isfinite(b["heat"][sig_idx]) else np.nan
        cohort_tradeable_arr[n] = cohort_tradeable
        cohort_rs85_arr[n] = cohort_rs85

        done = n + 1
        if done == 1 or done % 1000 == 0 or done == n_events:
            print(f"[trade] {done}/{n_events} entry events simulated", flush=True)

    # Flatten compact numeric matrices. Categories prevent repeated strategy/code/name
    # strings from ballooning memory while keeping downstream groupby semantics.
    rep_event = np.repeat(np.arange(n_events), nr)
    tiled_rule = np.tile(np.arange(nr), n_events)
    trades = pd.DataFrame({
        "代號": pd.Categorical(event_code[rep_event]),
        "名稱": pd.Categorical(event_name[rep_event]),
        "市場": pd.Categorical(event_market[rep_event]),
        "進場訊號日": sig_dates[rep_event],
        "進場日": ent_dates[rep_event],
        "進場價": entry_price_arr[rep_event],
        "進場RS": entry_rs_arr[rep_event],
        "進場右側": entry_right_arr[rep_event],
        "進場過熱": entry_heat_arr[rep_event],
        "策略": pd.Categorical(rule_names[tiled_rule]),
        "策略代碼": pd.Categorical(rule_keys[tiled_rule]),
        "類型": pd.Categorical(rule_types[tiled_rule]),
        "出場訊號日": exit_sig_dates.reshape(-1),
        "出場日": exit_dates.reshape(-1),
        "出場價": exit_price_mat.reshape(-1),
        "持有交易日": hold_mat.reshape(-1),
        "策略有觸發": trigger_mat.reshape(-1),
        "報酬%": ret_mat.reshape(-1),
        "出場前MAE%": mae_mat.reshape(-1),
        f"{max_hold}日峰值捕捉率%": cap_mat.reshape(-1),
        f"相對{max_hold}日高點回吐%": give_mat.reshape(-1),
        "出場後5日%": post5_mat.reshape(-1),
        "出場後10日%": post10_mat.reshape(-1),
        "群組_可交易": cohort_tradeable_arr[rep_event],
        "群組_RS85未過熱": cohort_rs85_arr[rep_event],
    })
    trades = trades[pd.to_numeric(trades["出場價"], errors="coerce").notna()].reset_index(drop=True)
    return trades, {
        "entry_events": n_events,
        "skipped_recent_no_full_horizon": skipped_future,
        "rule_count": nr,
        "optimization": "v1.2 numpy-matrix + triggered-only validation",
    }


# ---------------------------------------------------------------------------
# Summary / ranking
# ---------------------------------------------------------------------------
def _cohort_views(trades: pd.DataFrame):
    yield "全部鴨嘴新進", trades
    if "群組_可交易" in trades.columns:
        yield "新進且可交易", trades[trades["群組_可交易"].fillna(False).astype(bool)]
    if "群組_RS85未過熱" in trades.columns:
        yield "新進+RS85+未過熱", trades[trades["群組_RS85未過熱"].fillna(False).astype(bool)]


def _metric_pack(g: pd.DataFrame, max_hold: int, prefix: str = "") -> dict:
    """Return a compact metric set for one policy/trigger subset."""
    peak_col = f"{max_hold}日峰值捕捉率%"
    give_col = f"相對{max_hold}日高點回吐%"
    ret = pd.to_numeric(g.get("報酬%"), errors="coerce")
    cap = pd.to_numeric(g.get(peak_col), errors="coerce")
    give = pd.to_numeric(g.get(give_col), errors="coerce")
    mae = pd.to_numeric(g.get("出場前MAE%"), errors="coerce")
    p5 = pd.to_numeric(g.get("出場後5日%"), errors="coerce")
    p10 = pd.to_numeric(g.get("出場後10日%"), errors="coerce")
    hold = pd.to_numeric(g.get("持有交易日"), errors="coerce")
    def f(s, fn, digits=3):
        try:
            v = getattr(s, fn)()
            return round(float(v), digits) if pd.notna(v) else np.nan
        except Exception:
            return np.nan
    return {
        f"{prefix}平均報酬%": f(ret, "mean"),
        f"{prefix}中位報酬%": f(ret, "median"),
        f"{prefix}勝率%": round(float((ret > 0).mean() * 100.0), 2) if len(ret) else np.nan,
        f"{prefix}平均持有日": f(hold, "mean", 2),
        f"{prefix}中位持有日": f(hold, "median", 1),
        f"{prefix}平均峰值捕捉率%": f(cap, "mean", 2),
        f"{prefix}中位峰值捕捉率%": f(cap, "median", 2),
        f"{prefix}平均高點回吐%": f(give, "mean"),
        f"{prefix}中位高點回吐%": f(give, "median"),
        f"{prefix}平均MAE%": f(mae, "mean"),
        f"{prefix}平均出場後5日%": f(p5, "mean"),
        f"{prefix}平均出場後10日%": f(p10, "mean"),
    }


def summarize_trades(trades: pd.DataFrame, max_hold: int) -> pd.DataFrame:
    """v1.2: separate actual signal hits from max-hold fallback outcomes.

    The old all-event policy metrics are preserved for comparability.  New
    ``觸發_`` columns describe only cases where the exit signal really fired;
    ``未觸發_`` columns show the cohort that simply reached the common max hold.
    """
    if trades.empty:
        return pd.DataFrame()
    rows = []
    for cohort, view in _cohort_views(trades):
        if view.empty:
            continue
        for (strategy, key, kind), g in view.groupby(["策略", "策略代碼", "類型"], dropna=False, observed=True):
            trigger = g["策略有觸發"].fillna(False).astype(bool)
            hit = g[trigger].copy()
            miss = g[~trigger].copy()
            row = {
                "進場群組": cohort,
                "策略": strategy,
                "策略代碼": key,
                "類型": kind,
                "樣本數": int(len(g)),
                "觸發樣本數": int(len(hit)),
                "未觸發樣本數": int(len(miss)),
                "有效觸發率%": round(float(trigger.mean() * 100.0), 2) if len(g) else np.nan,
            }
            # Existing columns = policy performance, including max-hold fallback.
            row.update(_metric_pack(g, max_hold=max_hold, prefix=""))
            # Actual signal quality only.
            row.update(_metric_pack(hit, max_hold=max_hold, prefix="觸發_"))
            # What happened when the rule never fired within the common window.
            row.update(_metric_pack(miss, max_hold=max_hold, prefix="未觸發_"))
            rows.append(row)
    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out["政策分數"] = np.nan
    out["訊號分數"] = np.nan

    def pct_rank(series: pd.Series, higher=True):
        ss = pd.to_numeric(series, errors="coerce")
        if not higher:
            ss = -ss
        return ss.rank(pct=True, method="average").fillna(0.0) * 100.0

    for cohort, idx in out.groupby("進場群組").groups.items():
        cidx = pd.Index([i for i in idx if str(out.at[i, "類型"]) != "基準"])
        if len(cidx) == 0:
            continue
        policy = (
            0.30 * pct_rank(out.loc[cidx, "中位報酬%"], True)
            + 0.25 * pct_rank(out.loc[cidx, "中位峰值捕捉率%"], True)
            + 0.20 * pct_rank(out.loc[cidx, "平均高點回吐%"], False)
            + 0.15 * pct_rank(out.loc[cidx, "平均出場後10日%"], False)
            + 0.10 * pct_rank(out.loc[cidx, "平均MAE%"], True)
        )
        out.loc[cidx, "政策分數"] = policy.round(2)

        signal = (
            0.30 * pct_rank(out.loc[cidx, "觸發_中位報酬%"], True)
            + 0.25 * pct_rank(out.loc[cidx, "觸發_中位峰值捕捉率%"], True)
            + 0.20 * pct_rank(out.loc[cidx, "觸發_平均高點回吐%"], False)
            + 0.15 * pct_rank(out.loc[cidx, "觸發_平均出場後10日%"], False)
            + 0.10 * pct_rank(out.loc[cidx, "觸發_平均MAE%"], True)
        )
        tr = pd.to_numeric(out.loc[cidx, "有效觸發率%"], errors="coerce").fillna(0)
        ntr = pd.to_numeric(out.loc[cidx, "觸發樣本數"], errors="coerce").fillna(0)
        reliability_factor = np.where(
            (tr >= 40) & (ntr >= 150), 1.0,
            np.where((tr >= 25) & (ntr >= 70), 0.92, 0.80)
        )
        out.loc[cidx, "訊號分數"] = (signal * reliability_factor).round(2)

    # Final validation rank is added after annual stability is known.
    out["年度穩定分數"] = np.nan
    out["綜合分數"] = np.nan
    out["群組排名"] = np.nan
    out["是否進入綜合排名"] = "否"
    out["可靠度"] = "待年度驗證"
    return out

def annual_summary(trades: pd.DataFrame, max_hold: int) -> pd.DataFrame:
    """Year-by-year policy and actual-trigger results."""
    if trades.empty:
        return pd.DataFrame()
    rows = []
    for cohort, view in _cohort_views(trades):
        if view.empty:
            continue
        x = view.copy()
        x["年度"] = pd.to_datetime(x["進場日"], errors="coerce").dt.year
        for (strategy, key, kind, year), g in x.groupby(["策略", "策略代碼", "類型", "年度"], dropna=False, observed=True):
            trigger = g["策略有觸發"].fillna(False).astype(bool)
            hit = g[trigger]
            row = {
                "進場群組": cohort,
                "策略": strategy,
                "策略代碼": key,
                "類型": kind,
                "年度": int(year) if pd.notna(year) else year,
                "樣本數": int(len(g)),
                "觸發樣本數": int(len(hit)),
                "未觸發樣本數": int((~trigger).sum()),
                "有效觸發率%": round(float(trigger.mean() * 100.0), 2) if len(g) else np.nan,
            }
            row.update({f"全樣本_{k}": v for k, v in _metric_pack(g, max_hold, prefix="").items()})
            row.update(_metric_pack(hit, max_hold, prefix="觸發_"))
            rows.append(row)
    return pd.DataFrame(rows)


def annual_stability_summary(annual: pd.DataFrame, min_trigger_year_samples: int = 30) -> pd.DataFrame:
    if annual is None or annual.empty:
        return pd.DataFrame()
    rows = []
    for (cohort, strategy, key, kind), g in annual.groupby(["進場群組", "策略", "策略代碼", "類型"], observed=True):
        valid = g[pd.to_numeric(g["觸發樣本數"], errors="coerce").fillna(0) >= min_trigger_year_samples].copy()
        med = pd.to_numeric(valid.get("觸發_中位報酬%"), errors="coerce") if not valid.empty else pd.Series(dtype=float)
        pos = int((med > 0).sum()) if len(med) else 0
        nyr = int(len(valid))
        rows.append({
            "進場群組": cohort,
            "策略": strategy,
            "策略代碼": key,
            "類型": kind,
            "年度有效年數": nyr,
            "年度正中位報酬年數": pos,
            "年度正報酬率%": round(pos / nyr * 100.0, 1) if nyr else np.nan,
            "跨年中位報酬%": round(float(med.median()), 3) if len(med) else np.nan,
            "最差年度中位報酬%": round(float(med.min()), 3) if len(med) else np.nan,
            "最佳年度中位報酬%": round(float(med.max()), 3) if len(med) else np.nan,
            "最低年度觸發率%": round(float(pd.to_numeric(valid["有效觸發率%"], errors="coerce").min()), 2) if nyr else np.nan,
            "年度穩定分數": round(pos / nyr * 100.0, 2) if nyr else np.nan,
        })
    return pd.DataFrame(rows)


def apply_validation_ranking(summary: pd.DataFrame, stability: pd.DataFrame) -> pd.DataFrame:
    """Combine policy, actual-trigger quality, and year consistency.

    This is a research-priority score, not an optimized trading parameter.
    """
    if summary is None or summary.empty:
        return summary
    out = summary.copy()
    if stability is not None and not stability.empty:
        keep = [
            "進場群組", "策略代碼", "年度有效年數", "年度正中位報酬年數",
            "年度正報酬率%", "跨年中位報酬%", "最差年度中位報酬%",
            "最佳年度中位報酬%", "最低年度觸發率%", "年度穩定分數",
        ]
        out = out.drop(columns=[c for c in keep[2:] if c in out.columns], errors="ignore")
        out = out.merge(stability[keep], on=["進場群組", "策略代碼"], how="left")
    else:
        out["年度有效年數"] = 0
        out["年度正中位報酬年數"] = 0
        out["年度正報酬率%"] = np.nan
        out["年度穩定分數"] = np.nan

    out["綜合分數"] = np.nan
    out["群組排名"] = np.nan
    out["是否進入綜合排名"] = "否"

    for cohort, idx in out.groupby("進場群組").groups.items():
        cidx = pd.Index([i for i in idx if str(out.at[i, "類型"]) != "基準"])
        if len(cidx) == 0:
            continue
        p = pd.to_numeric(out.loc[cidx, "政策分數"], errors="coerce").fillna(0)
        s = pd.to_numeric(out.loc[cidx, "訊號分數"], errors="coerce").fillna(0)
        y = pd.to_numeric(out.loc[cidx, "年度穩定分數"], errors="coerce").fillna(0)
        score = 0.40 * p + 0.45 * s + 0.15 * y
        out.loc[cidx, "綜合分數"] = score.round(2)

        tr = pd.to_numeric(out.loc[cidx, "有效觸發率%"], errors="coerce").fillna(0)
        ntr = pd.to_numeric(out.loc[cidx, "觸發樣本數"], errors="coerce").fillna(0)
        yrs = pd.to_numeric(out.loc[cidx, "年度有效年數"], errors="coerce").fillna(0)
        eligible_mask = (tr >= 25) & (ntr >= 70) & (yrs >= 2)
        eligible = cidx[eligible_mask.to_numpy()]
        if len(eligible):
            out.loc[eligible, "群組排名"] = out.loc[eligible, "綜合分數"].rank(ascending=False, method="min").astype(float)
            out.loc[eligible, "是否進入綜合排名"] = "是"

    def reliability(r):
        n = int(_num(r.get("觸發樣本數"), 0) or 0)
        tr = float(_num(r.get("有效觸發率%"), 0) or 0)
        yrs = int(_num(r.get("年度有效年數"), 0) or 0)
        if n >= 150 and tr >= 40 and yrs >= 3: return "高"
        if n >= 70 and tr >= 25 and yrs >= 2: return "中"
        return "需擴充樣本"
    out["可靠度"] = out.apply(reliability, axis=1)
    return out.sort_values(["進場群組", "類型", "群組排名", "綜合分數"], ascending=[True, True, True, False], na_position="last")

def best_by_metric(summary: pd.DataFrame) -> pd.DataFrame:
    if summary is None or summary.empty:
        return pd.DataFrame()
    rows = []
    base = summary[
        (summary["類型"] != "基準")
        & (summary["是否進入綜合排名"].astype(str) == "是")
    ].copy()
    for cohort, g in base.groupby("進場群組", observed=True):
        if g.empty:
            continue
        specs = [
            ("v1.2驗證第一", "綜合分數", True),
            ("實際觸發中位報酬最高", "觸發_中位報酬%", True),
            ("實際觸發峰值捕捉最好", "觸發_中位峰值捕捉率%", True),
            ("實際觸發高點回吐最小", "觸發_平均高點回吐%", False),
            ("實際觸發後10日最弱", "觸發_平均出場後10日%", False),
            ("年度一致性最好", "年度穩定分數", True),
            ("最差年度表現最好", "最差年度中位報酬%", True),
        ]
        for label, col, higher in specs:
            if col not in g.columns:
                continue
            z = g.dropna(subset=[col])
            if z.empty:
                continue
            r = z.loc[z[col].idxmax() if higher else z[col].idxmin()]
            rows.append({
                "進場群組": cohort,
                "評比": label,
                "策略": r["策略"],
                "數值": r[col],
                "觸發樣本數": r.get("觸發樣本數"),
                "有效觸發率%": r.get("有效觸發率%"),
                "年度正報酬率%": r.get("年度正報酬率%"),
                "可靠度": r.get("可靠度"),
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
def select_excel_trade_detail(trades: pd.DataFrame, summary: pd.DataFrame, max_rows: int = 320000) -> pd.DataFrame:
    """Keep Excel readable and below the one-sheet row ceiling.

    Full strategy-event results remain in memory for ranking.  Excel receives
    only the #1 strategy from each cohort plus fixed-hold benchmarks.  This is
    enough to audit the recommendation while avoiding ~1M detail rows.
    """
    if trades is None or trades.empty:
        return pd.DataFrame()
    keys = set()
    if summary is not None and not summary.empty:
        top = summary[(summary["類型"] != "基準") & (pd.to_numeric(summary["群組排名"], errors="coerce") == 1)]
        keys.update(top["策略代碼"].astype(str).tolist())
    keys.update([k for k in trades[trades["類型"].astype(str) == "基準"]["策略代碼"].astype(str).unique().tolist()])
    detail = trades[trades["策略代碼"].astype(str).isin(keys)].copy()
    detail = detail.sort_values(["進場日", "代號", "策略"]).reset_index(drop=True)
    if len(detail) > max_rows:
        # Prefer all leader rows; thin benchmark detail deterministically if needed.
        lead = detail[detail["類型"].astype(str) != "基準"]
        bench = detail[detail["類型"].astype(str) == "基準"]
        room = max(0, max_rows - len(lead))
        if room < len(bench):
            step = max(1, int(math.ceil(len(bench) / max(1, room))))
            bench = bench.iloc[::step].head(room)
        detail = pd.concat([lead, bench], ignore_index=True).head(max_rows)
    return detail


def write_outputs(
    summary: pd.DataFrame,
    best: pd.DataFrame,
    annual: pd.DataFrame,
    stability: pd.DataFrame,
    heat_summary: pd.DataFrame,
    heat_detail: pd.DataFrame,
    trades: pd.DataFrame,
    status: dict,
) -> None:
    summary.to_csv(OUT_SUMMARY, index=False, encoding="utf-8-sig")
    validation_cols = [c for c in [
        "進場群組","策略","類型","群組排名","綜合分數","政策分數","訊號分數","年度穩定分數",
        "樣本數","觸發樣本數","未觸發樣本數","有效觸發率%",
        "觸發_中位報酬%","觸發_勝率%","觸發_中位持有日","觸發_中位峰值捕捉率%",
        "觸發_平均高點回吐%","觸發_平均出場後10日%",
        "年度有效年數","年度正中位報酬年數","年度正報酬率%","跨年中位報酬%","最差年度中位報酬%",
        "可靠度"
    ] if c in summary.columns]
    summary[validation_cols].to_csv(OUT_VALIDATION, index=False, encoding="utf-8-sig")
    notes = pd.DataFrame({
        "說明": [
            "第一階段只比較可逐日、Point-in-time 重建的技術/RS/右側/過熱出場條件；歷史基本面與估值不硬塞入回測。",
            "鴨嘴條件：收盤>MA20/MA60、MA20>MA60、MA20上揚、MA60上揚、MA20-MA60開口擴大。",
            "RS：250交易日報酬在同一交易日全市場做百分位排名。",
            "進場：T日收盤確認『鴨嘴今日新進』，T+1日開盤成交；出場亦為T日收盤確認、T+1日開盤成交，避免收盤訊號偷看未來。",
            "峰值捕捉率與高點回吐都使用相同的最大持有窗，避免不同策略因觀察期間不同而失真。",
            "政策分數沿用v1.1全事件評分；訊號分數只看真正觸發案例；v1.2最終綜合分數=40%政策分數＋45%訊號分數＋15%年度穩定分數。",
            "綜合分數只用來排序研究優先級，不代表已找到永遠有效的最佳參數；需再看年度穩定度與較長歷史。",
            "最近進場事件若沒有完整最大持有窗，會排除在公平比較之外。",
            "v1.1：排行榜仍使用全部策略×全部事件；Excel 的交易明細只保留各進場群組第1名策略與固定持有基準，避免明細超過單表合理容量。",
            "v1.2：將『實際觸發』與『60日內未觸發而強制結算』完全拆開；最終驗證分數同時看政策結果、實際觸發品質與年度一致性。",
            "年度穩定：單一年度至少30筆實際觸發才列入年度一致性；2026若資料尚未完整，視為部分年度，解讀時需保守。",
        ]
    })

    trade_detail = select_excel_trade_detail(trades, summary)
    status["excel_trade_detail_rows"] = int(len(trade_detail))
    OUT_STATUS.write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")

    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
        book = writer.book
        hdr = book.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center"})
        good = book.add_format({"bg_color": "#E2F0D9"})
        warn = book.add_format({"bg_color": "#FFF2CC"})
        hot = book.add_format({"bg_color": "#F4CCCC"})

        sheets = [
            ("出場策略排行榜", summary),
            ("各指標最佳", best),
            ("年度逐年驗證", annual),
            ("年度穩定摘要", stability),
            ("過熱事件統計", heat_summary),
            ("過熱事件明細", heat_detail),
            ("交易明細_精選", trade_detail),
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
    print("=== 個股出場策略回測 v1.2（觸發/未觸發分離＋年度驗證）===", flush=True)
    print(f"正式研究區間: {start} ~ {end}", flush=True)
    print(f"RS/均線暖機: {warmup_start} 起", flush=True)
    print(f"最大共同持有窗: {max_hold} 交易日", flush=True)

    cache_report = ensure_history_cache(warmup_start, end, fetch_missing=fetch_missing, workers=workers)
    raw, cache_meta = load_market_cache(warmup_start, end)
    print(f"[data] {cache_meta}", flush=True)
    cache_last = _date(cache_meta["last_date"])
    effective_end = min(end, cache_last)
    if effective_end < end:
        print(f"[WARN] requested end={end}, available cache ends={cache_last}; effective end={effective_end}", flush=True)
    signals = compute_signals(raw)
    del raw
    valid_rs = signals[(signals["date"].dt.date >= start) & signals["rs"].notna()]
    print(
        f"[signal] rows={len(signals):,}; research RS rows={len(valid_rs):,}; "
        f"duck new={int(signals[(signals['date'].dt.date>=start)&(signals['date'].dt.date<=end)]['new'].sum())}",
        flush=True,
    )

    print("[phase] simulate exit rules", flush=True)
    trades, trade_meta = simulate_trades(signals, start, effective_end, max_hold=max_hold)
    print(f"[phase] compact trade rows={len(trades):,}; summarize", flush=True)
    summary = summarize_trades(trades, max_hold=max_hold)
    annual = annual_summary(trades, max_hold=max_hold)
    stability = annual_stability_summary(annual)
    summary = apply_validation_ranking(summary, stability)
    best = best_by_metric(summary)
    print("[phase] heat-event study", flush=True)
    heat_detail, heat_summary = heat_event_study(signals, start, effective_end)

    status = {
        "status": "success" if not summary.empty else "no_results",
        "generated_at_taipei": dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).strftime("%Y-%m-%d %H:%M:%S"),
        "research_start": start.isoformat(),
        "research_end_requested": end.isoformat(),
        "research_end_effective": effective_end.isoformat(),
        "warmup_start": warmup_start.isoformat(),
        "max_hold_sessions": int(max_hold),
        "cache": cache_report,
        "data": cache_meta,
        "trades": trade_meta,
        "validation": "v1.2 actual-trigger separated; annual stability",
        "trade_rows": int(len(trades)),
        "summary_rows": int(len(summary)),
        "heat_events": int(len(heat_detail)),
        "method": "T close signal -> T+1 open execution; technical/RS point-in-time reconstruction",
        "outputs": [OUT_XLSX.name, OUT_SUMMARY.name, OUT_VALIDATION.name, OUT_STATUS.name],
    }
    if not summary.empty:
        leaders = (
            summary[(summary["類型"] != "基準") & (summary["群組排名"] == 1)]
            [["進場群組", "策略", "綜合分數", "觸發_中位報酬%", "觸發_勝率%", "觸發_中位持有日", "有效觸發率%", "年度正報酬率%", "可靠度"]]
            .to_dict("records")
        )
        status["current_leaders"] = leaders
    write_outputs(summary, best, annual, stability, heat_summary, heat_detail, trades, status)
    print(f"[done] {OUT_XLSX}", flush=True)
    print(f"[done] {OUT_SUMMARY}", flush=True)
    print(f"[done] {OUT_VALIDATION}", flush=True)
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
