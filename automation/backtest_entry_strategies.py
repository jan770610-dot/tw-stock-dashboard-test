# -*- coding: utf-8 -*-
"""Backtest entry timing for the Taiwan-stock Duckbill / RS decision system.

Research goal
-------------
Use only point-in-time information already available in the current project:
- TWSE/TPEx daily all-market cache
- MA20 / MA60 / Duckbill structure
- 250-session cross-sectional RS
- current right-side score definition
- current heat score definition
- market breadth reconstructed from the same daily cache

No current fundamental / valuation snapshot is used historically because the
repository does not retain point-in-time fundamentals for every historical day.
That avoids look-ahead bias.

Signal convention
-----------------
A signal is known after day T close.
Entry is executed on T+1 open.
Entry quality is evaluated with fixed forward windows so that all entry rules
are compared independently of exit-rule choice.

Outputs at repository root:
- entry_backtest_latest.xlsx
- entry_backtest_summary.csv
- entry_backtest_annual.csv
- entry_backtest_status.json
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RS_DIR = ROOT / "automation" / "rs"
CACHE_DIR = RS_DIR / "cache"
HOLIDAY_FILE = CACHE_DIR / "nontrading_dates.txt"
STATUS_FILE = ROOT / "update_status.json"

OUT_XLSX = ROOT / "entry_backtest_latest.xlsx"
OUT_SUMMARY = ROOT / "entry_backtest_summary.csv"
OUT_ANNUAL = ROOT / "entry_backtest_annual.csv"
OUT_STATUS = ROOT / "entry_backtest_status.json"

DEFAULT_START = dt.date(2023, 1, 3)
DEFAULT_MAX_FORWARD = 60
WARMUP_CALENDAR_DAYS = 430


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
    return [
        d.date()
        for d in pd.date_range(start, end, freq="D")
        if d.weekday() < 5
    ]


def ensure_history_cache(start: dt.date, end: dt.date, fetch_missing: bool, workers: int) -> dict:
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

    if str(RS_DIR) not in sys.path:
        sys.path.insert(0, str(RS_DIR))
    import rs_breadth  # type: ignore

    workers = max(1, min(4, int(workers)))
    print(f"[history] missing weekdays={len(missing)}; fetch workers={workers}", flush=True)

    def task(day: dt.date):
        try:
            d, frame, status = rs_breadth.fetch_day(day, analysis_only=False)
            return d, frame is not None and not frame.empty, str(status), None
        except Exception as e:
            return day, False, "exception", f"{type(e).__name__}: {e}"

    completed = 0
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(task, d): d for d in missing}
        for fut in as_completed(futs):
            d, ok, status, err = fut.result()
            completed += 1
            if ok and _cache_path(d).exists():
                report["fetched_ok"] += 1
            elif status == "holiday":
                report["new_nontrading"] += 1
            else:
                report["failed"].append({
                    "date": d.isoformat(),
                    "status": status,
                    "error": err or "",
                })
            if completed == 1 or completed % 25 == 0 or completed == len(missing):
                print(
                    f"[history] {completed}/{len(missing)} | ok={report['fetched_ok']} "
                    f"holiday={report['new_nontrading']} fail={len(report['failed'])}",
                    flush=True,
                )

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
                    still.append({"date": d.isoformat(), "status": str(status), "error": "retry incomplete"})
            except Exception as e:
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
            frames.append(z[keep].copy())
        except Exception as e:
            errors.append(f"{p.name}:{type(e).__name__}:{e}")
        if i == 1 or i % 100 == 0 or i == len(files):
            print(f"[load] {i}/{len(files)} daily cache files", flush=True)

    if not frames:
        raise RuntimeError("行情快取存在，但沒有可解析資料")

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


def compute_signals(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    x = df.copy().sort_values(["code", "date"]).reset_index(drop=True)
    grp = x.groupby("code", group_keys=False, observed=True)

    x["prev_close"] = grp["close"].shift(1)
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
    x["duck_prev"] = grp["duck"].shift(1).astype("boolean").fillna(False).astype(bool)
    x["new"] = x["duck"] & ~x["duck_prev"]

    x["raw_hot"] = x["duck"] & (x["bias20"] >= 8.0)
    false_count = (~x["duck"]).groupby(x["code"], observed=True).cumsum()
    x["had_hot"] = x["raw_hot"].groupby([x["code"], false_count], observed=True).cummax() & x["duck"]

    rs = pd.to_numeric(x["rs"], errors="coerce")
    bias = pd.to_numeric(x["bias20"], errors="coerce").fillna(0.0)
    heat = (bias.clip(lower=0.0) * 5.0).clip(lower=0.0, upper=70.0).astype("float64")
    heat += np.select([rs >= 95, rs >= 85], [12.0, 8.0], default=0.0)
    heat += np.where(x["duck"] | (x["complete"] >= 100), 8.0, 0.0)
    heat = np.where(x["raw_hot"].to_numpy(bool), np.maximum(70.0, heat + 18.0), heat)
    heat += np.where(x["had_hot"].to_numpy(bool), 8.0, 0.0)
    x["heat"] = np.clip(heat, 0.0, 100.0).astype("float32")

    right = (x["complete"].astype("float64") * 0.55).clip(lower=0.0, upper=55.0)
    right += np.select([rs >= 85, rs >= 70, rs >= 50], [25.0, 18.0, 10.0], default=0.0)
    right += np.where(x["new"], 12.0, np.where(x["duck"], 8.0, 0.0))
    sync_up = (x["ma20_change"] > 0) & (x["ma60_change"] > 0) & (x["spread_change"] > 0)
    right += np.where(sync_up, 6.0, 0.0)
    x["right"] = np.clip(right, 0.0, 100.0).astype("float32")

    x["prev_complete"] = grp["complete"].shift(1)
    x["prev_right"] = grp["right"].shift(1)
    x["prev_rs"] = grp["rs"].shift(1)

    x["advance"] = x["close"] > x["prev_close"]
    x["above_ma20"] = x["close"] > x["ma20"]
    x["above_ma60"] = x["close"] > x["ma60"]
    x["rs85"] = x["rs"] >= 85

    market = (
        x.groupby("date", observed=True)
        .agg(
            股票數=("code", "count"),
            上漲比例=("advance", "mean"),
            MA20上方比例=("above_ma20", "mean"),
            MA60上方比例=("above_ma60", "mean"),
            RS85比例=("rs85", "mean"),
        )
        .reset_index()
    )
    for c in ["上漲比例", "MA20上方比例", "MA60上方比例", "RS85比例"]:
        market[c] = pd.to_numeric(market[c], errors="coerce") * 100.0
    market["RS85比例5日變化"] = market["RS85比例"].diff(5)
    market["MA20上方比例5日變化"] = market["MA20上方比例"].diff(5)
    market["市場廣度改善"] = (
        (market["上漲比例"] >= 50)
        & (market["MA20上方比例"] >= 50)
        & ((market["RS85比例5日變化"] > 0) | (market["MA20上方比例5日變化"] > 0))
    )
    market["市場廣度強"] = (
        (market["上漲比例"] >= 55)
        & (market["MA20上方比例"] >= 55)
        & (market["MA60上方比例"] >= 50)
    )
    x = x.merge(
        market[["date", "上漲比例", "MA20上方比例", "MA60上方比例", "RS85比例",
                "RS85比例5日變化", "MA20上方比例5日變化", "市場廣度改善", "市場廣度強"]],
        on="date", how="left",
    )
    return x, market


STRATEGIES = [
    ("鴨嘴正式新進", "duck_new", "鴨嘴"),
    ("鴨嘴新進＋未過熱", "duck_new_unhot", "鴨嘴"),
    ("鴨嘴新進＋RS≥70", "duck_new_rs70", "鴨嘴×RS"),
    ("鴨嘴新進＋RS≥85", "duck_new_rs85", "鴨嘴×RS"),
    ("預備80首次出現", "pre80", "預備"),
    ("預備80＋RS≥70＋未過熱", "pre80_rs70", "預備×RS"),
    ("預備80＋RS≥85＋未過熱", "pre80_rs85", "預備×RS"),
    ("右側首次≥58＋未過熱", "right58", "右側"),
    ("右側首次≥65＋未過熱", "right65", "右側"),
    ("右側首次≥70＋RS≥70", "right70_rs70", "右側×RS"),
    ("RS突破70＋結構≥80", "rs70_struct", "RS"),
    ("RS突破85＋結構≥80", "rs85_struct", "RS"),
    ("正式鴨嘴回踩MA20守住", "pullback_ma20", "回踩"),
    ("鴨嘴新進＋市場廣度改善", "duck_new_market", "市場組合"),
    ("鴨嘴新進＋RS≥70＋市場廣度改善", "duck_new_rs70_market", "市場組合"),
    ("預備80＋RS≥70＋市場廣度改善", "pre80_rs70_market", "市場組合"),
    ("右側首次≥65＋市場廣度改善", "right65_market", "市場組合"),
]


def _rising_edge(mask: pd.Series, codes: pd.Series) -> pd.Series:
    prev = mask.groupby(codes, observed=True).shift(1).astype("boolean").fillna(False)
    return mask.astype(bool) & ~prev.astype(bool)


def build_strategy_masks(x: pd.DataFrame) -> Dict[str, pd.Series]:
    heat_ok = pd.to_numeric(x["heat"], errors="coerce").fillna(999) < 70
    rs = pd.to_numeric(x["rs"], errors="coerce")
    complete = pd.to_numeric(x["complete"], errors="coerce").fillna(0)
    right = pd.to_numeric(x["right"], errors="coerce").fillna(0)

    pre80_state = (complete >= 80) & (complete < 100)
    pre80 = _rising_edge(pre80_state, x["code"])

    right58 = (right >= 58) & (pd.to_numeric(x["prev_right"], errors="coerce").fillna(-1) < 58) & heat_ok
    right65 = (right >= 65) & (pd.to_numeric(x["prev_right"], errors="coerce").fillna(-1) < 65) & heat_ok
    right70_rs70 = (
        (right >= 70)
        & (pd.to_numeric(x["prev_right"], errors="coerce").fillna(-1) < 70)
        & (rs >= 70)
        & heat_ok
    )
    rs70_struct = (
        (rs >= 70)
        & (pd.to_numeric(x["prev_rs"], errors="coerce") < 70)
        & (complete >= 80)
        & heat_ok
    )
    rs85_struct = (
        (rs >= 85)
        & (pd.to_numeric(x["prev_rs"], errors="coerce") < 85)
        & (complete >= 80)
        & heat_ok
    )
    ma20 = pd.to_numeric(x["ma20"], errors="coerce")
    low = pd.to_numeric(x["low"], errors="coerce")
    close = pd.to_numeric(x["close"], errors="coerce")
    ma20chg = pd.to_numeric(x["ma20_change"], errors="coerce")
    pullback_state = (
        x["duck"].astype(bool)
        & (right >= 65)
        & heat_ok
        & (ma20chg > 0)
        & (low <= ma20 * 1.01)
        & (close >= ma20)
    )
    pullback = _rising_edge(pullback_state, x["code"])

    mkt = x["市場廣度改善"].fillna(False).astype(bool)

    return {
        "duck_new": x["new"].astype(bool),
        "duck_new_unhot": x["new"].astype(bool) & heat_ok,
        "duck_new_rs70": x["new"].astype(bool) & heat_ok & (rs >= 70),
        "duck_new_rs85": x["new"].astype(bool) & heat_ok & (rs >= 85),
        "pre80": pre80,
        "pre80_rs70": pre80 & heat_ok & (rs >= 70),
        "pre80_rs85": pre80 & heat_ok & (rs >= 85),
        "right58": right58,
        "right65": right65,
        "right70_rs70": right70_rs70,
        "rs70_struct": rs70_struct,
        "rs85_struct": rs85_struct,
        "pullback_ma20": pullback,
        "duck_new_market": x["new"].astype(bool) & heat_ok & mkt,
        "duck_new_rs70_market": x["new"].astype(bool) & heat_ok & (rs >= 70) & mkt,
        "pre80_rs70_market": pre80 & heat_ok & (rs >= 70) & mkt,
        "right65_market": right65 & mkt,
    }


def _cooldown_positions(pos: np.ndarray, cooldown: int) -> np.ndarray:
    if len(pos) <= 1 or cooldown <= 0:
        return pos
    keep = [int(pos[0])]
    last = int(pos[0])
    for p in pos[1:]:
        p = int(p)
        if p - last > cooldown:
            keep.append(p)
            last = p
    return np.asarray(keep, dtype=int)


def simulate_entries(
    x: pd.DataFrame,
    start: dt.date,
    end: dt.date,
    max_forward: int,
    cooldown: int,
) -> Tuple[pd.DataFrame, dict]:
    masks = build_strategy_masks(x)
    strategy_meta = {key: (name, kind) for name, key, kind in STRATEGIES}
    rows: List[dict] = []
    skipped_recent = 0
    raw_events = 0

    groups = {str(code): g.sort_values("date").reset_index(drop=True)
              for code, g in x.groupby("code", observed=True)}

    for gi, (code, g) in enumerate(groups.items(), 1):
        # Build local masks by reindexing from original row ids preserved through values.
        # Recompute masks locally for exact alignment.
        gmasks = build_strategy_masks(g)
        dates = pd.to_datetime(g["date"], errors="coerce")
        opens = pd.to_numeric(g["open"], errors="coerce").to_numpy(float)
        highs = pd.to_numeric(g["high"], errors="coerce").to_numpy(float)
        lows = pd.to_numeric(g["low"], errors="coerce").to_numpy(float)
        closes = pd.to_numeric(g["close"], errors="coerce").to_numpy(float)

        for key, mask in gmasks.items():
            if key not in strategy_meta:
                continue
            pos = np.flatnonzero(mask.to_numpy(dtype=bool))
            if len(pos) == 0:
                continue
            pos = _cooldown_positions(pos, cooldown)
            raw_events += len(pos)
            name, kind = strategy_meta[key]

            for sig_idx in pos:
                sig_day = dates.iloc[sig_idx].date()
                if sig_day < start or sig_day > end:
                    continue
                entry_idx = sig_idx + 1
                if entry_idx >= len(g) or entry_idx + max_forward >= len(g):
                    skipped_recent += 1
                    continue
                entry_price = opens[entry_idx]
                if not np.isfinite(entry_price) or entry_price <= 0:
                    entry_price = closes[entry_idx]
                if not np.isfinite(entry_price) or entry_price <= 0:
                    continue

                path_slice = slice(entry_idx, entry_idx + max_forward)
                ph = highs[path_slice]
                pl = lows[path_slice]
                pc = closes[path_slice]

                out = {
                    "策略": name,
                    "策略代碼": key,
                    "類型": kind,
                    "代號": code,
                    "名稱": str(g.iloc[sig_idx].get("name") or ""),
                    "市場": str(g.iloc[sig_idx].get("market") or ""),
                    "訊號日": sig_day.isoformat(),
                    "進場日": dates.iloc[entry_idx].date().isoformat(),
                    "進場價": round(float(entry_price), 4),
                    "訊號RS": round(_num(g.iloc[sig_idx].get("rs"), np.nan), 2),
                    "訊號右側": round(_num(g.iloc[sig_idx].get("right"), np.nan), 2),
                    "訊號過熱": round(_num(g.iloc[sig_idx].get("heat"), np.nan), 2),
                    "訊號完成度": round(_num(g.iloc[sig_idx].get("complete"), np.nan), 1),
                    "市場上漲比例": round(_num(g.iloc[sig_idx].get("上漲比例"), np.nan), 2),
                    "市場MA20上方比例": round(_num(g.iloc[sig_idx].get("MA20上方比例"), np.nan), 2),
                    "市場RS85比例5日變化": round(_num(g.iloc[sig_idx].get("RS85比例5日變化"), np.nan), 3),
                }

                for h in [5, 10, 20, 40, 60]:
                    if h <= max_forward:
                        px = pc[h - 1]
                        ret = (px / entry_price - 1.0) * 100.0 if np.isfinite(px) else np.nan
                        hh = np.nanmax(ph[:h]) if np.isfinite(ph[:h]).any() else np.nan
                        ll = np.nanmin(pl[:h]) if np.isfinite(pl[:h]).any() else np.nan
                        mfe = (hh / entry_price - 1.0) * 100.0 if np.isfinite(hh) else np.nan
                        mae = (ll / entry_price - 1.0) * 100.0 if np.isfinite(ll) else np.nan
                        out[f"{h}日報酬%"] = round(float(ret), 4) if np.isfinite(ret) else np.nan
                        out[f"{h}日MFE%"] = round(float(mfe), 4) if np.isfinite(mfe) else np.nan
                        out[f"{h}日MAE%"] = round(float(mae), 4) if np.isfinite(mae) else np.nan

                h20 = ph[:min(20, max_forward)]
                l20 = pl[:min(20, max_forward)]
                out["20日達+5%"] = bool(np.isfinite(h20).any() and np.nanmax(h20) >= entry_price * 1.05)
                out["20日曾跌-5%"] = bool(np.isfinite(l20).any() and np.nanmin(l20) <= entry_price * 0.95)
                h40 = ph[:min(40, max_forward)]
                l40 = pl[:min(40, max_forward)]
                out["40日達+10%"] = bool(np.isfinite(h40).any() and np.nanmax(h40) >= entry_price * 1.10)
                out["40日曾跌-8%"] = bool(np.isfinite(l40).any() and np.nanmin(l40) <= entry_price * 0.92)

                if np.isfinite(h20).any():
                    out["20日高點第幾日"] = int(np.nanargmax(h20) + 1)
                else:
                    out["20日高點第幾日"] = np.nan

                rows.append(out)

        if gi == 1 or gi % 250 == 0 or gi == len(groups):
            print(f"[simulate] {gi}/{len(groups)} stocks | rows={len(rows)}", flush=True)

    return pd.DataFrame(rows), {
        "raw_signal_events_after_cooldown": int(raw_events),
        "usable_event_rows": int(len(rows)),
        "skipped_recent_no_full_horizon": int(skipped_recent),
        "strategy_count": len(STRATEGIES),
        "cooldown_sessions": int(cooldown),
    }


def annual_summary(events: pd.DataFrame) -> pd.DataFrame:
    if events.empty:
        return pd.DataFrame()
    x = events.copy()
    x["年度"] = pd.to_datetime(x["進場日"], errors="coerce").dt.year
    rows = []
    for (strategy, key, kind, year), g in x.groupby(["策略", "策略代碼", "類型", "年度"], dropna=False):
        r20 = pd.to_numeric(g.get("20日報酬%"), errors="coerce")
        r10 = pd.to_numeric(g.get("10日報酬%"), errors="coerce")
        mae20 = pd.to_numeric(g.get("20日MAE%"), errors="coerce")
        rows.append({
            "策略": strategy,
            "策略代碼": key,
            "類型": kind,
            "年度": int(year),
            "樣本數": int(len(g)),
            "10日中位報酬%": round(float(r10.median()), 3),
            "20日中位報酬%": round(float(r20.median()), 3),
            "20日勝率%": round(float((r20 > 0).mean() * 100.0), 2),
            "20日中位MAE%": round(float(mae20.median()), 3),
            "20日達+5率%": round(float(g["20日達+5%"].astype(bool).mean() * 100.0), 2),
        })
    return pd.DataFrame(rows)


def summarize(events: pd.DataFrame, annual: pd.DataFrame) -> pd.DataFrame:
    if events.empty:
        return pd.DataFrame()

    rows = []
    for (strategy, key, kind), g in events.groupby(["策略", "策略代碼", "類型"], dropna=False):
        d = {
            "策略": strategy,
            "策略代碼": key,
            "類型": kind,
            "樣本數": int(len(g)),
        }
        for h in [5, 10, 20, 40, 60]:
            c = f"{h}日報酬%"
            if c in g.columns:
                s = pd.to_numeric(g[c], errors="coerce")
                d[f"{h}日平均報酬%"] = round(float(s.mean()), 3)
                d[f"{h}日中位報酬%"] = round(float(s.median()), 3)
                d[f"{h}日勝率%"] = round(float((s > 0).mean() * 100.0), 2)
        for h in [20, 40]:
            mfe = pd.to_numeric(g.get(f"{h}日MFE%"), errors="coerce")
            mae = pd.to_numeric(g.get(f"{h}日MAE%"), errors="coerce")
            d[f"{h}日中位MFE%"] = round(float(mfe.median()), 3)
            d[f"{h}日中位MAE%"] = round(float(mae.median()), 3)
        d["20日達+5率%"] = round(float(g["20日達+5%"].astype(bool).mean() * 100.0), 2)
        d["20日跌-5率%"] = round(float(g["20日曾跌-5%"].astype(bool).mean() * 100.0), 2)
        d["40日達+10率%"] = round(float(g["40日達+10%"].astype(bool).mean() * 100.0), 2)
        d["40日跌-8率%"] = round(float(g["40日曾跌-8%"].astype(bool).mean() * 100.0), 2)
        d["20日中位高點日在第幾日"] = round(float(pd.to_numeric(g["20日高點第幾日"], errors="coerce").median()), 1)

        ya = annual[annual["策略代碼"] == key].copy() if not annual.empty else pd.DataFrame()
        ya = ya[ya["樣本數"] >= 30] if not ya.empty else ya
        if ya.empty:
            d["年度有效年數"] = 0
            d["年度正報酬率%"] = np.nan
            d["跨年20日中位報酬%"] = np.nan
            d["最差年度20日中位報酬%"] = np.nan
        else:
            vals = pd.to_numeric(ya["20日中位報酬%"], errors="coerce")
            d["年度有效年數"] = int(len(ya))
            d["年度正報酬率%"] = round(float((vals > 0).mean() * 100.0), 2)
            d["跨年20日中位報酬%"] = round(float(vals.median()), 3)
            d["最差年度20日中位報酬%"] = round(float(vals.min()), 3)
        rows.append(d)

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    def rank_pct(series, higher=True):
        s = pd.to_numeric(series, errors="coerce")
        if not higher:
            s = -s
        return s.rank(pct=True, method="average").fillna(0) * 100.0

    score = (
        0.20 * rank_pct(out["10日中位報酬%"], True)
        + 0.25 * rank_pct(out["20日中位報酬%"], True)
        + 0.15 * rank_pct(out["20日勝率%"], True)
        + 0.15 * rank_pct(out["20日中位MFE%"], True)
        + 0.15 * rank_pct(out["20日中位MAE%"], True)  # closer to zero is higher
        + 0.10 * rank_pct(out["年度正報酬率%"], True)
    )
    n = pd.to_numeric(out["樣本數"], errors="coerce").fillna(0)
    reliability_factor = np.where(n >= 1000, 1.0, np.where(n >= 300, 0.96, np.where(n >= 100, 0.90, 0.82)))
    out["綜合分數"] = (score * reliability_factor).round(2)
    out["排名"] = out["綜合分數"].rank(ascending=False, method="min").astype(int)
    out["可靠度"] = np.select(
        [n >= 1000, n >= 300, n >= 100],
        ["高", "中高", "中"],
        default="需擴充樣本",
    )
    return out.sort_values(["排名", "樣本數"], ascending=[True, False]).reset_index(drop=True)


def write_outputs(
    summary: pd.DataFrame,
    annual: pd.DataFrame,
    events: pd.DataFrame,
    market: pd.DataFrame,
    status: dict,
) -> None:
    summary.to_csv(OUT_SUMMARY, index=False, encoding="utf-8-sig")
    annual.to_csv(OUT_ANNUAL, index=False, encoding="utf-8-sig")
    OUT_STATUS.write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")

    # Keep detailed Excel manageable: all summary/annual + event rows for top 5 strategies.
    top_keys = summary.head(5)["策略代碼"].astype(str).tolist() if not summary.empty else []
    detail = events[events["策略代碼"].astype(str).isin(top_keys)].copy() if top_keys else events.head(0).copy()
    if len(detail) > 300000:
        detail = detail.head(300000).copy()

    notes = pd.DataFrame({
        "說明": [
            "訊號在 T 日收盤後成立，統一使用 T+1 開盤作為進場價。",
            "第一輪只比較『進場品質』，用固定 5/10/20/40/60 日報酬、MFE、MAE、命中率評分，不把出場策略混進來。",
            "同一股票同一策略採 rising-edge 事件，並使用 cooldown 避免短期反覆觸發造成樣本灌水。",
            "歷史基本面/估值沒有逐日 point-in-time 快照，因此本回測不使用目前的三率、營收、EPS、折價率作歷史條件，以免偷看未來。",
            "市場廣度條件是用同一份每日全市場行情即時計算：上漲比例、MA20/MA60 上方比例、RS85 比例及 5 日變化。",
            "排名主要看 10/20 日中位報酬、20 日勝率、20 日 MFE/MAE 與跨年度穩定度。",
            "找到前 3 名進場訊號後，下一輪再與已驗證的 v1.2 出場規則組合，做完整進出場績效。",
        ]
    })

    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="進場策略排行榜", index=False)
        annual.to_excel(writer, sheet_name="年度驗證", index=False)
        detail.to_excel(writer, sheet_name="前5名事件明細", index=False)
        market.to_excel(writer, sheet_name="市場廣度歷史", index=False)
        notes.to_excel(writer, sheet_name="說明", index=False)

        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        pct_fmt = wb.add_format({"num_format": "0.00"})
        for sheet_name, df in [
            ("進場策略排行榜", summary),
            ("年度驗證", annual),
            ("前5名事件明細", detail),
            ("市場廣度歷史", market),
            ("說明", notes),
        ]:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(0, len(df)), max(0, len(df.columns) - 1))
            for j, col in enumerate(df.columns):
                ws.write(0, j, col, header_fmt)
                width = min(34, max(10, len(str(col)) + 2))
                if col in {"策略", "類型", "名稱"}:
                    width = max(width, 18)
                if sheet_name == "說明":
                    width = 100
                ws.set_column(j, j, width, pct_fmt if str(col).endswith("%") else None)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", default=DEFAULT_START.isoformat())
    ap.add_argument("--end", default="")
    ap.add_argument("--max-forward", type=int, default=DEFAULT_MAX_FORWARD)
    ap.add_argument("--cooldown", type=int, default=10)
    ap.add_argument("--workers", type=int, default=2)
    ap.add_argument("--fetch-missing", action="store_true")
    args = ap.parse_args()

    start = _date(args.start)
    requested_end = _date(args.end) if args.end else _latest_formal_date()
    warmup = start - dt.timedelta(days=WARMUP_CALENDAR_DAYS)
    max_forward = max(20, min(80, int(args.max_forward)))
    cooldown = max(0, min(30, int(args.cooldown)))

    print("=== 個股進場時機回測 ===", flush=True)
    print(f"研究區間 requested: {start} ~ {requested_end}", flush=True)
    print(f"暖機: {warmup} 起", flush=True)
    print(f"最大固定觀察窗: {max_forward} 交易日", flush=True)
    print(f"同策略同股票 cooldown: {cooldown} 交易日", flush=True)

    cache_report = ensure_history_cache(warmup, requested_end, args.fetch_missing, args.workers)
    raw, meta = load_market_cache(warmup, requested_end)
    effective_end = min(requested_end, pd.to_datetime(meta["last_date"]).date())
    signals, market = compute_signals(raw)

    print(f"[data] {meta}", flush=True)
    print("[signals] computed", flush=True)

    events, sim_meta = simulate_entries(signals, start, effective_end, max_forward, cooldown)
    if events.empty:
        raise RuntimeError("沒有可用的進場事件")

    annual = annual_summary(events)
    summary = summarize(events, annual)
    if summary.empty:
        raise RuntimeError("無法建立進場策略摘要")

    leaders = []
    for _, r in summary.head(8).iterrows():
        leaders.append({
            "排名": int(r["排名"]),
            "策略": str(r["策略"]),
            "類型": str(r["類型"]),
            "綜合分數": float(r["綜合分數"]),
            "樣本數": int(r["樣本數"]),
            "10日中位報酬%": _num(r.get("10日中位報酬%")),
            "20日中位報酬%": _num(r.get("20日中位報酬%")),
            "20日勝率%": _num(r.get("20日勝率%")),
            "20日中位MFE%": _num(r.get("20日中位MFE%")),
            "20日中位MAE%": _num(r.get("20日中位MAE%")),
            "年度正報酬率%": _num(r.get("年度正報酬率%")),
            "可靠度": str(r.get("可靠度")),
        })

    status = {
        "status": "success",
        "generated_at_taipei": dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).strftime("%Y-%m-%d %H:%M:%S"),
        "research_start": start.isoformat(),
        "research_end_requested": requested_end.isoformat(),
        "research_end_effective": effective_end.isoformat(),
        "warmup_start": warmup.isoformat(),
        "max_forward_sessions": max_forward,
        "cooldown_sessions": cooldown,
        "cache": cache_report,
        "data": meta,
        "simulation": sim_meta,
        "event_rows": int(len(events)),
        "summary_rows": int(len(summary)),
        "annual_rows": int(len(annual)),
        "method": "T close signal -> T+1 open entry; fixed forward windows; point-in-time technical/RS/breadth reconstruction",
        "fundamental_history_used": False,
        "fundamental_history_note": "No point-in-time historical fundamental/valuation snapshots in repo; excluded to avoid look-ahead bias.",
        "current_leaders": leaders,
        "outputs": [
            OUT_XLSX.name,
            OUT_SUMMARY.name,
            OUT_ANNUAL.name,
            OUT_STATUS.name,
        ],
    }
    write_outputs(summary, annual, events, market, status)

    print("[done] top entry strategies:", flush=True)
    show_cols = ["排名", "策略", "類型", "樣本數", "10日中位報酬%", "20日中位報酬%", "20日勝率%", "20日中位MFE%", "20日中位MAE%", "年度正報酬率%", "綜合分數", "可靠度"]
    print(summary[show_cols].head(12).to_string(index=False), flush=True)


if __name__ == "__main__":
    main()
