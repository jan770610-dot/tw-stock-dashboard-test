# -*- coding: utf-8 -*-
"""Build a causality-safe historical Wade market-state database and backtest it.

This script intentionally reuses the project's formal RS/raw-market functions from
``automation/rs/rs_breadth.py`` so historical results stay aligned with the live
system.  The important difference is that historical water-level percentiles are
computed *causally* (only information available up to each date), avoiding the
look-ahead bias that would occur if a 2019 signal used the 2018-2026 full-sample
percentile distribution.

Outputs (default: data/wade_history):
- wade_market_history.csv.gz       daily causal database
- wade_signal_events.csv           de-duplicated signal/event starts
- wade_backtest_summary.csv        event-level aggregate statistics
- wade_backtest_by_year.csv        annual stability table
- wade_history_status.json         coverage / missing-date audit
- wade_market_history.parquet      optional, when parquet engine is available

Signal timing convention:
A signal is known only after day D closes.  Backtest entry is therefore D+1 close,
and N-day forward return is measured from D+1 close to N trading sessions later.
MFE/MAE are close-to-close path statistics, not intraday high/low excursions.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Iterable

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parent
RS_DIR = HERE / "rs"
OUT_DIR_DEFAULT = REPO_ROOT / "data" / "wade_history"


def _load_rs_module():
    if str(RS_DIR) not in sys.path:
        sys.path.insert(0, str(RS_DIR))
    import rs_breadth as rs  # type: ignore
    return rs


def _parse_day(value: str, *, default: dt.date | None = None) -> dt.date:
    s = str(value or "").strip().lower()
    if s in {"", "auto", "today"}:
        return default or dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).date()
    return dt.datetime.strptime(s, "%Y-%m-%d").date()


def _business_days(start: dt.date, end: dt.date) -> list[dt.date]:
    return [x.date() for x in pd.date_range(start, end, freq="B")]


def _expanding_percentile(series: pd.Series, min_periods: int = 252) -> pd.Series:
    """Causal percentile: each row compares only with rows at/before that date."""
    vals = pd.to_numeric(series, errors="coerce").to_numpy(dtype=float)
    out = np.full(len(vals), np.nan, dtype=float)
    history: list[float] = []
    for i, v in enumerate(vals):
        if not np.isfinite(v):
            continue
        history.append(float(v))
        if len(history) < min_periods:
            continue
        a = np.asarray(history, dtype=float)
        out[i] = float((a <= v).sum() / len(a) * 100.0)
    return pd.Series(out, index=series.index, dtype="float64")


def _ratio_strength_score(v) -> float:
    try:
        if pd.isna(v):
            return 50.0
        return float(np.clip((float(v) - 35.0) / 30.0 * 100.0, 0.0, 100.0))
    except Exception:
        return 50.0


def _balance_score(pos, neg) -> float:
    try:
        p = max(float(pos), 0.0) if not pd.isna(pos) else 0.0
        n = max(float(neg), 0.0) if not pd.isna(neg) else 0.0
        return 50.0 if p + n <= 0 else p / (p + n) * 100.0
    except Exception:
        return 50.0


def _build_causal_history(raw: pd.DataFrame, start: dt.date, end: dt.date, cfg, rs) -> pd.DataFrame:
    """Rebuild Wade v0.2 using causal water-level percentiles."""
    d = raw.copy().sort_values("date").drop_duplicates("date", keep="last")
    d["date"] = pd.to_datetime(d["date"], errors="coerce").dt.normalize()
    d = d.dropna(subset=["date"]).reset_index(drop=True)

    # Core RS breadth path.  Keep warm-up rows until after all rolling features exist.
    d["eligible_count"] = pd.to_numeric(d["eligible_count"], errors="coerce")
    d["strong_count"] = pd.to_numeric(d["strong_count"], errors="coerce")
    d["strong_pct"] = np.where(
        d["eligible_count"] > 0,
        d["strong_count"] / d["eligible_count"] * 100.0,
        np.nan,
    )
    d["change"] = d["strong_count"].diff()
    d["ma5"] = d["strong_count"].rolling(5, min_periods=5).mean()
    d["ma20"] = d["strong_count"].rolling(20, min_periods=20).mean()
    d["ma60"] = d["strong_count"].rolling(60, min_periods=60).mean()
    d["pct_ma5"] = d["strong_pct"].rolling(5, min_periods=5).mean()
    d["pct_ma20"] = d["strong_pct"].rolling(20, min_periods=20).mean()
    d["pct_ma60"] = d["strong_pct"].rolling(60, min_periods=60).mean()
    d["pct_ma5_prev5"] = d["pct_ma5"].shift(5)
    d["pct_ma20_prev5"] = d["pct_ma20"].shift(5)
    d["breadth_spread"] = d["pct_ma5"] - d["pct_ma20"]
    d["breadth_spread_change5"] = d["breadth_spread"] - d["breadth_spread"].shift(5)
    d["breadth_speed"] = [
        rs.breadth_speed_label(a, b)
        for a, b in zip(d["breadth_spread"], d["breadth_spread_change5"])
    ]
    d["trend"] = [
        rs.trend_label(a, b, c, e, f)
        for a, b, c, e, f in zip(
            d["pct_ma5"], d["pct_ma20"], d["pct_ma60"],
            d["pct_ma5_prev5"], d["pct_ma20_prev5"],
        )
    ]

    # Keep both research views.  retrospective_pct_rank matches today's historical
    # display concept; causal_pct_rank is the only one allowed to drive backtests.
    d["retrospective_pct_rank"] = d["strong_pct"].rank(method="average", pct=True) * 100.0
    d["causal_pct_rank"] = _expanding_percentile(d["strong_pct"], min_periods=252)
    rolling_days = int(cfg.get("trend", "rolling_percentile_days", fallback="756"))
    rolling_min = int(cfg.get("trend", "rolling_percentile_min_periods", fallback="252"))
    d["rolling3y_pct_rank"] = rs.rolling_percentile(d["strong_pct"], rolling_days, rolling_min)

    # If causal full-history percentile is not mature yet, the 3Y rolling percentile
    # is allowed only when it has independently reached its own minimum sample size.
    d["signal_pct_rank"] = d["causal_pct_rank"].where(
        d["causal_pct_rank"].notna(), d["rolling3y_pct_rank"]
    )
    d["water"] = d["signal_pct_rank"].map(rs.water_label)
    d["extreme"] = d["signal_pct_rank"].map(rs.extreme_signal)
    d["market_state"] = [rs.market_state(w, t) for w, t in zip(d["water"], d["trend"])]

    # Index trend fields used for research / later dashboard integration.
    for prefix in ["twii", "twoii"]:
        c = f"{prefix}_close"
        if c not in d.columns:
            d[c] = np.nan
        close = pd.to_numeric(d[c], errors="coerce")
        d[f"{prefix}_ma20"] = close.rolling(20, min_periods=10).mean()
        d[f"{prefix}_ma60"] = close.rolling(60, min_periods=30).mean()
        d[f"{prefix}_ma20_change"] = d[f"{prefix}_ma20"].diff()
        d[f"{prefix}_ma60_change"] = d[f"{prefix}_ma60"].diff()

    d["amount_ma20"] = pd.to_numeric(d.get("total_amount"), errors="coerce").rolling(20, min_periods=10).mean()
    calc_ratio = np.where(
        d["amount_ma20"] > 0,
        pd.to_numeric(d.get("total_amount"), errors="coerce") / d["amount_ma20"] * 100.0,
        np.nan,
    )
    if "amount_ratio20" not in d.columns:
        d["amount_ratio20"] = calc_ratio
    else:
        d["amount_ratio20"] = pd.to_numeric(d["amount_ratio20"], errors="coerce").fillna(
            pd.Series(calc_ratio, index=d.index)
        )

    idx_sync: list[float] = []
    vol_eff: list[float] = []
    leadership: list[float] = []
    tpex_rel: list[float] = []
    for r in d.itertuples(index=False):
        ar = getattr(r, "advance_ratio", np.nan)
        vr = getattr(r, "amount_ratio20", np.nan)
        mr = getattr(r, "mean_return", np.nan)
        twii = getattr(r, "twii_ret", np.nan)
        twse_eq = getattr(r, "twse_mean_return", np.nan)
        twoii = getattr(r, "twoii_ret", np.nan)
        tpex_eq = getattr(r, "tpex_mean_return", np.nan)

        breadth = _ratio_strength_score(ar)
        activity = 50.0 if pd.isna(vr) else float(np.clip(50.0 + (float(vr) - 100.0) * 0.6, 0.0, 100.0))
        vol_eff.append(round(0.7 * breadth + 0.3 * (activity if (pd.isna(mr) or float(mr) >= 0) else 100.0 - activity), 1))

        if not pd.isna(twii) and not pd.isna(twse_eq):
            gap = abs(float(twii) - float(twse_eq))
            same = (float(twii) >= 0) == (float(twse_eq) >= 0)
            idx_sync.append(round(float(np.clip((80.0 if same else 25.0) - gap * 8.0, 0.0, 100.0)), 1))
        else:
            idx_sync.append(round(_ratio_strength_score(getattr(r, "twse_advance_ratio", np.nan)), 1))

        if not pd.isna(twoii) and not pd.isna(twii):
            tpex_rel.append(round(float(np.clip(50.0 + (float(twoii) - float(twii)) * 12.0, 0.0, 100.0)), 1))
        elif not pd.isna(tpex_eq) and not pd.isna(twse_eq):
            tpex_rel.append(round(float(np.clip(50.0 + (float(tpex_eq) - float(twse_eq)) * 12.0, 0.0, 100.0)), 1))
        else:
            tpex_rel.append(round(_ratio_strength_score(getattr(r, "tpex_advance_ratio", np.nan)), 1))

        leadership.append(round(_ratio_strength_score(getattr(r, "leader_retention_ratio", np.nan)), 1))

    d["volume_efficiency_score"] = vol_eff
    d["index_sync_score"] = idx_sync
    d["tpex_relative_score"] = tpex_rel
    d["leadership_score"] = leadership

    dir_score_map = {"明顯增加": 100.0, "增加": 75.0, "持平": 50.0, "減少": 25.0, "明顯減少": 0.0}
    scores: list[float] = []
    for r in d.itertuples(index=False):
        adv = _ratio_strength_score(getattr(r, "advance_ratio", np.nan))
        rs_s = getattr(r, "signal_pct_rank", np.nan)
        rs_s = 50.0 if pd.isna(rs_s) else float(rs_s)
        nh = _balance_score(getattr(r, "new_high_count", np.nan), getattr(r, "new_low_count", np.nan))
        lim = _balance_score(getattr(r, "limit_up_count", np.nan), getattr(r, "limit_down_count", np.nan))
        ds = dir_score_map.get(str(getattr(r, "trend", "")), 50.0)
        ve = float(getattr(r, "volume_efficiency_score", 50.0))
        sync = float(getattr(r, "index_sync_score", 50.0))
        trel = float(getattr(r, "tpex_relative_score", 50.0))
        lead = float(getattr(r, "leadership_score", 50.0))
        scores.append(round(
            0.20 * adv + 0.20 * rs_s + 0.10 * nh + 0.07 * lim + 0.12 * ds
            + 0.10 * ve + 0.08 * sync + 0.06 * trel + 0.07 * lead,
            1,
        ))
    d["wade_score"] = scores
    if "advance_count" in d.columns:
        d.loc[pd.to_numeric(d["advance_count"], errors="coerce").isna(), "wade_score"] = np.nan
    d["wade_score_change5"] = d["wade_score"] - d["wade_score"].shift(5)
    d["wade_state"] = [rs._wade_state(s, w, t) for s, w, t in zip(d["wade_score"], d["water"], d["trend"])]

    left: list[str] = []
    early: list[str] = []
    for r in d.itertuples(index=False):
        score = getattr(r, "wade_score", np.nan)
        water = str(getattr(r, "water", ""))
        trend = str(getattr(r, "trend", ""))
        ar = getattr(r, "advance_ratio", np.nan)
        tr = getattr(r, "tpex_advance_ratio", np.nan)
        ch5 = getattr(r, "wade_score_change5", np.nan)
        high = water in {"偏高水位", "極高水位"}
        weakening = trend in {"減少", "明顯減少"}
        if high and weakening:
            left.append("🟠 左側減碼觀察")
        elif high and not pd.isna(score) and float(score) < 50.0:
            left.append("🟠 高檔內部轉弱")
        elif high and not pd.isna(ch5) and float(ch5) <= -12.0:
            left.append("🟡 強度快速降溫")
        elif not pd.isna(ar) and not pd.isna(tr) and float(ar) < 45.0 and float(tr) < 45.0 and weakening:
            left.append("🟡 盤面廣度偏弱")
        else:
            left.append("—")

        nh = getattr(r, "new_high_count", np.nan)
        nl = getattr(r, "new_low_count", np.nan)
        low_mid = water in {"極低水位", "偏低水位", "正常水位"}
        improving = trend in {"增加", "明顯增加"}
        if (
            low_mid and improving
            and not pd.isna(ar) and float(ar) >= 55.0
            and not pd.isna(tr) and float(tr) >= 50.0
            and (pd.isna(nh) or pd.isna(nl) or float(nh) >= float(nl))
        ):
            early.append("🟢 早期轉強")
        else:
            early.append("—")
    d["left_reduce_alert"] = left
    d["early_strength_signal"] = early

    # An equal-weight market proxy keeps the database researchable even when Yahoo
    # index history is temporarily unavailable.  It is not an investable index.
    mr = pd.to_numeric(d.get("mean_return"), errors="coerce").fillna(0.0)
    d["equal_weight_proxy"] = (1.0 + mr / 100.0).cumprod() * 100.0

    d["month"] = d["date"].dt.to_period("M").astype(str)
    d["year"] = d["date"].dt.year
    mask = (d["date"].dt.date >= start) & (d["date"].dt.date <= end)
    return d.loc[mask].copy().reset_index(drop=True)


def _choose_benchmark(d: pd.DataFrame) -> tuple[str, str]:
    twii = pd.to_numeric(d.get("twii_close"), errors="coerce")
    coverage = float(twii.notna().mean()) if len(d) else 0.0
    if coverage >= 0.80:
        return "twii_close", "加權指數收盤"
    return "equal_weight_proxy", "全市場等權代理"


def _event_start_mask(s: pd.Series) -> pd.Series:
    cur = s.fillna("—").astype(str)
    active = cur.ne("—") & cur.ne("") & cur.ne("資料累積中") & cur.ne("資料不足")
    return active & cur.ne(cur.shift(1).fillna("—"))


def _state_transition_mask(s: pd.Series) -> pd.Series:
    cur = s.fillna("資料不足").astype(str)
    valid = ~cur.isin(["資料不足", "資料累積中", ""])
    return valid & cur.ne(cur.shift(1))


def _forward_path_stats(values: np.ndarray, signal_i: int, horizon: int) -> tuple[float, float, float] | tuple[float, float, float]:
    """Entry = D+1 close; exit = horizon sessions after entry."""
    entry_i = signal_i + 1
    exit_i = entry_i + horizon
    if entry_i >= len(values) or exit_i >= len(values):
        return (np.nan, np.nan, np.nan)
    entry = values[entry_i]
    if not np.isfinite(entry) or entry <= 0:
        return (np.nan, np.nan, np.nan)
    path = values[entry_i:exit_i + 1]
    path = path[np.isfinite(path)]
    if len(path) < 2:
        return (np.nan, np.nan, np.nan)
    ret = (path[-1] / entry - 1.0) * 100.0
    mfe = (np.max(path) / entry - 1.0) * 100.0
    mae = (np.min(path) / entry - 1.0) * 100.0
    return (float(ret), float(mfe), float(mae))


def _build_events(d: pd.DataFrame) -> pd.DataFrame:
    benchmark_col, benchmark_label = _choose_benchmark(d)
    values = pd.to_numeric(d[benchmark_col], errors="coerce").to_numpy(dtype=float)
    families: list[tuple[str, str, pd.Series]] = [
        ("市場階段", "market_state", _state_transition_mask(d["market_state"])),
        ("減碼警示", "left_reduce_alert", _event_start_mask(d["left_reduce_alert"])),
        ("早期轉強", "early_strength_signal", _event_start_mask(d["early_strength_signal"])),
    ]
    horizons = [1, 3, 5, 10, 20, 40, 60]
    rows: list[dict] = []
    for family, col, mask in families:
        for i in np.flatnonzero(mask.to_numpy()):
            r = d.iloc[i]
            item = {
                "signal_family": family,
                "signal_label": str(r[col]),
                "signal_date": pd.Timestamp(r["date"]).date().isoformat(),
                "signal_year": int(pd.Timestamp(r["date"]).year),
                "benchmark": benchmark_label,
                "market_state": r.get("market_state"),
                "water": r.get("water"),
                "trend": r.get("trend"),
                "wade_score": r.get("wade_score"),
                "wade_state": r.get("wade_state"),
                "strong_pct": r.get("strong_pct"),
                "signal_pct_rank": r.get("signal_pct_rank"),
                "advance_ratio": r.get("advance_ratio"),
                "tpex_advance_ratio": r.get("tpex_advance_ratio"),
                "new_high_count": r.get("new_high_count"),
                "new_low_count": r.get("new_low_count"),
            }
            entry_i = i + 1
            item["entry_date"] = (
                pd.Timestamp(d.iloc[entry_i]["date"]).date().isoformat()
                if entry_i < len(d) else None
            )
            item["entry_value"] = float(values[entry_i]) if entry_i < len(values) and np.isfinite(values[entry_i]) else np.nan
            for h in horizons:
                ret, mfe, mae = _forward_path_stats(values, i, h)
                item[f"ret_{h}d_pct"] = ret
                if h in {20, 60}:
                    item[f"mfe_{h}d_pct"] = mfe
                    item[f"mae_{h}d_pct"] = mae
            rows.append(item)
    return pd.DataFrame(rows).sort_values(["signal_date", "signal_family"]).reset_index(drop=True) if rows else pd.DataFrame()


def _summary_tables(events: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if events.empty:
        return pd.DataFrame(), pd.DataFrame()
    horizons = [1, 3, 5, 10, 20, 40, 60]

    def summarize(group: pd.DataFrame) -> dict:
        fam = str(group["signal_family"].iloc[0])
        out: dict = {"events": int(len(group))}
        for h in horizons:
            c = f"ret_{h}d_pct"
            s = pd.to_numeric(group[c], errors="coerce").dropna()
            out[f"avg_{h}d_pct"] = float(s.mean()) if len(s) else np.nan
            out[f"median_{h}d_pct"] = float(s.median()) if len(s) else np.nan
            out[f"positive_rate_{h}d_pct"] = float((s > 0).mean() * 100.0) if len(s) else np.nan
            out[f"negative_rate_{h}d_pct"] = float((s < 0).mean() * 100.0) if len(s) else np.nan
            if fam == "早期轉強":
                out[f"signal_success_rate_{h}d_pct"] = out[f"positive_rate_{h}d_pct"]
            elif fam == "減碼警示":
                out[f"signal_success_rate_{h}d_pct"] = out[f"negative_rate_{h}d_pct"]
            else:
                out[f"signal_success_rate_{h}d_pct"] = out[f"positive_rate_{h}d_pct"]
        for h in [20, 60]:
            for kind in ["mfe", "mae"]:
                s = pd.to_numeric(group[f"{kind}_{h}d_pct"], errors="coerce").dropna()
                out[f"avg_{kind}_{h}d_pct"] = float(s.mean()) if len(s) else np.nan
                out[f"median_{kind}_{h}d_pct"] = float(s.median()) if len(s) else np.nan
        return out

    rows = []
    for (family, label), g in events.groupby(["signal_family", "signal_label"], dropna=False):
        row = {"signal_family": family, "signal_label": label}
        row.update(summarize(g))
        rows.append(row)
    summary = pd.DataFrame(rows).sort_values(["signal_family", "events"], ascending=[True, False]).reset_index(drop=True)

    yearly_rows = []
    for (family, label, year), g in events.groupby(["signal_family", "signal_label", "signal_year"], dropna=False):
        row = {"signal_family": family, "signal_label": label, "year": int(year)}
        row.update(summarize(g))
        yearly_rows.append(row)
    yearly = pd.DataFrame(yearly_rows).sort_values(["signal_family", "signal_label", "year"]).reset_index(drop=True)
    return summary, yearly


def _fetch_all(rs, dates: Iterable[dt.date], workers: int, retry_rounds: int) -> tuple[list[dt.date], list[tuple[dt.date, str]]]:
    dates = list(dates)
    completed: list[dt.date] = []
    failed: list[tuple[dt.date, str]] = []
    total = len(dates)
    t0 = time.time()

    def one(day):
        return rs.fetch_day(day, False)

    if workers <= 1:
        iterator = enumerate(dates, 1)
        for i, day in iterator:
            d, frame, status = one(day)
            if frame is not None and not frame.empty:
                completed.append(d)
            elif status not in {"weekend", "holiday", "holiday-cache"}:
                failed.append((d, status))
            if i % 25 == 0 or i == total:
                print(f"[BACKFILL] {i}/{total}｜成功cache/下載 {len(completed)}｜待重試 {len(failed)}｜{(time.time()-t0)/60:.1f} 分", flush=True)
    else:
        with ThreadPoolExecutor(max_workers=max(1, min(int(workers), 4))) as ex:
            futs = {ex.submit(one, day): day for day in dates}
            done = 0
            for fut in as_completed(futs):
                done += 1
                d, frame, status = fut.result()
                if frame is not None and not frame.empty:
                    completed.append(d)
                elif status not in {"weekend", "holiday", "holiday-cache"}:
                    failed.append((d, status))
                if done % 25 == 0 or done == total:
                    print(f"[BACKFILL] {done}/{total}｜成功cache/下載 {len(completed)}｜待重試 {len(failed)}｜{(time.time()-t0)/60:.1f} 分", flush=True)

    for round_no in range(1, max(0, retry_rounds) + 1):
        if not failed:
            break
        print(f"[RETRY] 第 {round_no} 輪，共 {len(failed)} 日", flush=True)
        retry = failed
        failed = []
        for i, (day, _) in enumerate(retry, 1):
            d, frame, status = one(day)
            if frame is not None and not frame.empty:
                completed.append(d)
            elif status not in {"weekend", "holiday", "holiday-cache"}:
                failed.append((d, status))
            if i % 20 == 0:
                time.sleep(0.5)
    return sorted(set(completed)), sorted(failed)


def _load_cached_market(rs, start: dt.date, end: dt.date) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    days = _business_days(start, end)
    for i, day in enumerate(days, 1):
        z = rs.load_cached(day)
        if z is None or z.empty:
            continue
        z = z.copy()
        z["date"] = day.isoformat()
        frames.append(z)
        if i % 250 == 0:
            print(f"[LOAD CACHE] {i}/{len(days)}", flush=True)
    if not frames:
        return pd.DataFrame()
    x = pd.concat(frames, ignore_index=True)
    x["date"] = pd.to_datetime(x["date"], errors="coerce")
    x["code"] = x["code"].astype(str)
    return x.dropna(subset=["date", "code"]).drop_duplicates(["date", "code"], keep="last")


def _write_outputs(out_dir: Path, history: pd.DataFrame, events: pd.DataFrame, summary: pd.DataFrame, yearly: pd.DataFrame, status: dict) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    history.to_csv(out_dir / "wade_market_history.csv.gz", index=False, encoding="utf-8-sig", compression="gzip")
    events.to_csv(out_dir / "wade_signal_events.csv", index=False, encoding="utf-8-sig")
    summary.to_csv(out_dir / "wade_backtest_summary.csv", index=False, encoding="utf-8-sig")
    yearly.to_csv(out_dir / "wade_backtest_by_year.csv", index=False, encoding="utf-8-sig")
    (out_dir / "wade_history_status.json").write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")
    try:
        history.to_parquet(out_dir / "wade_market_history.parquet", index=False)
        status["parquet"] = "success"
        (out_dir / "wade_history_status.json").write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        status["parquet"] = f"skipped: {type(e).__name__}: {e}"
        (out_dir / "wade_history_status.json").write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Backfill and backtest causal Wade market signals")
    parser.add_argument("--start", default="2018-01-01", help="research start YYYY-MM-DD")
    parser.add_argument("--end", default="auto", help="research end YYYY-MM-DD or auto")
    parser.add_argument("--warmup-start", default="2016-01-01", help="raw-data warmup start; 2016 recommended")
    parser.add_argument("--workers", type=int, default=1, help="historical date concurrency; 1 safest, max 4")
    parser.add_argument("--retry-rounds", type=int, default=2)
    parser.add_argument("--analysis-only", action="store_true", help="do not download; compute from existing daily caches only")
    parser.add_argument("--out-dir", default=str(OUT_DIR_DEFAULT))
    args = parser.parse_args()

    rs = _load_rs_module()
    cfg = rs.cfg_get()
    today = dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).date()
    start = _parse_day(args.start)
    end = _parse_day(args.end, default=today)
    warmup_start = _parse_day(args.warmup_start)
    if warmup_start >= start:
        raise SystemExit("warmup-start 必須早於 research start；建議 2016-01-01 -> 2018-01-01")
    if end < start:
        raise SystemExit("end 不可早於 start")

    days = _business_days(warmup_start, end)
    print("=" * 76)
    print("Wade 歷史資料庫 / 無偷看未來回測 v1.0")
    print(f"原始暖機：{warmup_start} ~ {end}｜研究期：{start} ~ {end}")
    print(f"待檢查平日：{len(days):,}｜workers={args.workers}｜analysis_only={args.analysis_only}")
    print("訊號在 D 收盤確認；回測最早由 D+1 收盤開始。")
    print("=" * 76, flush=True)

    failed: list[tuple[dt.date, str]] = []
    if not args.analysis_only:
        _, failed = _fetch_all(rs, days, args.workers, args.retry_rounds)

    market = _load_cached_market(rs, warmup_start, end)
    if market.empty:
        raise SystemExit("沒有可用歷史市場 cache。請先取消 --analysis-only 執行一次完整回補。")

    print(f"[COMPUTE] 載入 {len(market):,} 筆股票日資料，開始重建每日 RS / 廣度...", flush=True)
    raw, _ = rs._daily_raw_from_market(market, cfg)

    # Add TWII/TWOII.  Existing function is fail-soft: if Yahoo is unavailable,
    # Wade's index components fall back to internal market participation proxies.
    try:
        idx = rs._update_index_store(warmup_start, end, args.analysis_only)
    except Exception as e:
        print(f"[WARN] 指數資料補取失敗：{e}", flush=True)
        idx = pd.DataFrame()
    if not raw.empty and idx is not None and not idx.empty:
        raw = raw.copy()
        idx = idx.copy()
        raw["date"] = pd.to_datetime(raw["date"], errors="coerce").dt.normalize()
        idx["date"] = pd.to_datetime(idx["date"], errors="coerce").dt.normalize()
        cols = [c for c in ["date", "twii_close", "twii_ret", "twoii_close", "twoii_ret"] if c in idx.columns]
        raw = raw.merge(idx[cols], on="date", how="left")

    history = _build_causal_history(raw, start, end, cfg, rs)
    if history.empty:
        raise SystemExit("歷史每日資料建立失敗：研究期間沒有資料。")
    events = _build_events(history)
    summary, yearly = _summary_tables(events)

    expected = set(_business_days(warmup_start, end))
    cached = {pd.Timestamp(x).date() for x in market["date"].dropna().unique()}
    known_holidays = set(rs._known_nontrading_days())
    missing = sorted(expected - cached - known_holidays)
    status = {
        "version": "1.0-causal",
        "generated_at_taipei": dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).isoformat(),
        "research_start": start.isoformat(),
        "research_end": end.isoformat(),
        "warmup_start": warmup_start.isoformat(),
        "market_rows": int(len(market)),
        "history_days": int(len(history)),
        "event_rows": int(len(events)),
        "failed_after_retry": [{"date": d.isoformat(), "reason": reason} for d, reason in failed],
        "missing_weekdays_not_marked_holiday": [d.isoformat() for d in missing],
        "benchmark": _choose_benchmark(history)[1],
        "lookahead_safe": True,
        "signal_timing": "D close confirms signal; D+1 close is first backtest entry",
    }
    _write_outputs(Path(args.out_dir), history, events, summary, yearly, status)

    print("=" * 76)
    print(f"完成：{Path(args.out_dir).resolve()}")
    print(f"研究交易日：{len(history):,}｜事件：{len(events):,}｜未標記缺日：{len(missing)}")
    if not summary.empty:
        focus = summary[summary["signal_family"].isin(["早期轉強", "減碼警示"])]
        print(focus[[c for c in ["signal_family", "signal_label", "events", "avg_20d_pct", "signal_success_rate_20d_pct", "avg_mae_20d_pct", "avg_mfe_20d_pct"] if c in focus.columns]].to_string(index=False))
    print("=" * 76)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
