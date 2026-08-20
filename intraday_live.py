# -*- coding: utf-8 -*-
"""Intraday trial analysis for the Taiwan-stock Streamlit dashboard.

Design goals
------------
1. Never overwrite formal daily files (rs_latest.xlsx / duck_latest.xlsx).
2. Reuse the previous close's compact rolling baseline.
3. Fetch TWSE/TPEx MIS quotes during the session and recompute the same RS
   strong-stock core plus a Wade-style *intraday trial* score.
4. Make trial/estimated fields explicit in the UI.
5. Run full-market scans in a background worker so automatic updates never block the UI.

The quote transport intentionally uses only Python stdlib so the dashboard does
not need a new package merely for the intraday feature.
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo
import datetime as dt
import json
import math
import os
import time
import threading

import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
BASELINE = APP_DIR / "intraday_baseline.pkl.gz"
BASELINE_STATUS = APP_DIR / "intraday_baseline_status.json"
TRADE_STATE_FILE = APP_DIR / "intraday_trade_state_today.json"
MIS_URL = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp"
TZ = ZoneInfo("Asia/Taipei")
ENGINE_VERSION = "3.6.3-mis-resilience-1"

RS_THRESHOLD = 85.0
AMOUNT_MIN = 30_000_000.0
MAX_BELOW_52W_HIGH = 0.25
MIN_ABOVE_52W_LOW = 0.30


def _f(v, default=None):
    try:
        if v in (None, "", "-", "--"):
            return default
        x = float(str(v).replace(",", ""))
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _norm_code(v) -> str:
    return str(v).replace(".0", "").strip()


def _first_book_price(v):
    """Return the first usable price from MIS five-level book fields."""
    if v in (None, "", "-", "--"):
        return None
    for part in str(v).split("_"):
        x = _f(part)
        if x is not None and x > 0:
            return x
    return None


def _channel(code: str, market: str) -> str:
    m = str(market).strip().upper()
    prefix = "otc" if m in {"TPEX", "OTC", "上櫃"} else "tse"
    return f"{prefix}_{code}.tw"


def _fetch_batch(channels: list[str], timeout: int = 10, attempts: int = 3) -> tuple[list[dict], str | None]:
    """Fetch one MIS batch with short retry/backoff for transient network failures."""
    if not channels:
        return [], None
    last_err=None
    for attempt in range(max(1,int(attempts))):
        query = urlencode({"ex_ch": "|".join(channels), "json": "1", "delay": "0", "_": str(int(time.time() * 1000))})
        req = Request(
            f"{MIS_URL}?{query}",
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/138 Safari/537.36",
                "Referer": "https://mis.twse.com.tw/stock/index.jsp",
                "Accept": "application/json,text/plain,*/*",
            },
        )
        try:
            with urlopen(req, timeout=timeout) as resp:
                payload = json.loads(resp.read().decode("utf-8", errors="replace"))
            rows=payload.get("msgArray", []) or []
            if rows:
                return rows, None
            last_err="MIS 回傳空批次"
        except Exception as e:
            last_err=f"{type(e).__name__}: {e}"
        if attempt < max(1,int(attempts))-1:
            time.sleep(0.35 * (2 ** attempt))
    return [], last_err


def _fetch_mis_quotes_raw(universe_json: str, batch_size: int = 80, workers: int = 4) -> tuple[pd.DataFrame, list[str]]:
    """Pure-Python quote fetcher. Safe to call from the background worker (no Streamlit API)."""
    universe = json.loads(universe_json)
    channels = [_channel(str(r["code"]), str(r["market"])) for r in universe]
    batches = [channels[i:i + batch_size] for i in range(0, len(channels), batch_size)]
    raw: list[dict] = []
    errors: list[str] = []
    with ThreadPoolExecutor(max_workers=max(1, min(workers, 8))) as ex:
        futs = {ex.submit(_fetch_batch, b): b for b in batches}
        for fut in as_completed(futs):
            rows, err = fut.result()
            raw.extend(rows)
            if err:
                errors.append(err)

    rows = []
    for m in raw:
        code = _norm_code(m.get("c") or "")
        if not code:
            ch = str(m.get("ch") or "")
            code = ch.split(".")[0].split("_")[-1]
        if not code:
            continue
        prev = _f(m.get("y"))
        z_last = _f(m.get("z"))
        best_bid = _first_book_price(m.get("b"))
        best_ask = _first_book_price(m.get("a"))

        # v3.5.2：只有 MIS z 才當「最新成交價」。
        # z 暫時沒有值時，不再使用 pz，也不再取買一/賣一中間價，
        # 避免出現像 245.25 這種普通股票實際不可成交的價位。
        # 試算若需要價格，改採真實存在於委託簿上的買一（較保守），
        # 其次賣一，最後才以前收作備援；UI 一律標示為參考價而非成交價。
        if z_last is not None:
            last = z_last
            price_source = "MIS z 最新成交"
            estimated = False
        elif best_bid is not None:
            last = best_bid
            price_source = "最佳買價參考（無最新成交）"
            estimated = True
        elif best_ask is not None:
            last = best_ask
            price_source = "最佳賣價參考（無最新成交）"
            estimated = True
        else:
            last = prev
            price_source = "前收備援（尚無盤中成交/五檔）"
            estimated = True

        if last is None:
            continue
        traded = z_last is not None
        high = _f(m.get("h"), last)
        low = _f(m.get("l"), last)
        open_ = _f(m.get("o"), last)
        volume_lots = _f(m.get("v"), 0.0) or 0.0
        rows.append({
            "code": code,
            "name_live": str(m.get("n") or ""),
            "last": last,
            "prev_close_live": prev,
            "open_live": open_,
            "high_live": high,
            "low_live": low,
            "best_bid": best_bid,
            "best_ask": best_ask,
            "price_source": price_source,
            "price_estimated": estimated,
            "volume_lots": volume_lots,
            "traded": traded,
            "quote_time": str(m.get("t") or ""),
            "quote_date": str(m.get("d") or ""),
            "quote_long": str(m.get("tlong") or ""),
        })
    q = pd.DataFrame(rows)
    if not q.empty:
        q = q.drop_duplicates("code", keep="last")
    return q, errors


def _fetch_mis_quotes_resilient(universe_json: str) -> tuple[pd.DataFrame, list[str]]:
    """Two-stage MIS fetch for full-market scans.

    First pass stays moderately parallel.  If some/most batches fail, retry only
    the missing symbols in smaller batches with one worker.  This reduces the
    chance that a temporary MIS throttle turns into a complete 2,000-stock scan
    failure on Streamlit Cloud.
    """
    universe = json.loads(universe_json)
    total = len(universe)
    q1, err1 = _fetch_mis_quotes_raw(universe_json, batch_size=80, workers=3)

    have = set()
    if q1 is not None and not q1.empty and "code" in q1.columns:
        have = set(q1["code"].astype(str).map(_norm_code))

    # If first pass is already very healthy, avoid extra network calls.
    coverage1 = (len(have) / total * 100.0) if total else 100.0
    if coverage1 >= 92.0 and not err1:
        return q1, err1

    missing = [
        {"code": str(r.get("code", "")), "market": str(r.get("market", ""))}
        for r in universe
        if _norm_code(r.get("code")) not in have
    ]
    q2 = pd.DataFrame()
    err2 = []
    if missing:
        time.sleep(0.45)
        missing_json = json.dumps(missing, ensure_ascii=False, separators=(",", ":"))
        q2, err2 = _fetch_mis_quotes_raw(missing_json, batch_size=45, workers=1)

    parts = [q for q in [q1, q2] if q is not None and not q.empty]
    if parts:
        q = pd.concat(parts, ignore_index=True)
        if "code" in q.columns:
            q["code"] = q["code"].astype(str).map(_norm_code)
            q = q.drop_duplicates("code", keep="last")
    else:
        q = pd.DataFrame()

    errors = list(err1 or []) + list(err2 or [])
    return q, errors


@st.cache_data(ttl=25, show_spinner=False)
def fetch_mis_quotes(universe_json: str, batch_size: int = 80, workers: int = 4) -> tuple[pd.DataFrame, list[str]]:
    return _fetch_mis_quotes_raw(universe_json, batch_size=batch_size, workers=workers)


def _load_baseline_raw(path_s: str) -> pd.DataFrame:
    return pd.read_pickle(path_s, compression="gzip")


@st.cache_data(ttl=300, show_spinner=False)
def load_baseline(path_s: str, mtime_ns: int) -> pd.DataFrame:
    return _load_baseline_raw(path_s)


def _ratio_strength(v) -> float:
    x = _f(v)
    if x is None:
        return 50.0
    return float(max(0.0, min(100.0, (x - 35.0) / 30.0 * 100.0)))


def _balance_score(pos, neg) -> float:
    p = max(_f(pos, 0.0) or 0.0, 0.0)
    n = max(_f(neg, 0.0) or 0.0, 0.0)
    return 50.0 if p + n <= 0 else p / (p + n) * 100.0


def _water_label(rank: float) -> str:
    if rank <= 10: return "極低水位"
    if rank <= 25: return "偏低水位"
    if rank < 75: return "正常水位"
    if rank < 90: return "偏高水位"
    return "極高水位"


def _market_state(water: str, trend: str) -> str:
    low = water in {"極低水位", "偏低水位"}
    high = water in {"極高水位", "偏高水位"}
    rising = trend in {"明顯增加", "增加"}
    falling = trend in {"明顯減少", "減少"}
    if low:
        return "低檔回升" if rising else "低檔續弱" if falling else "低檔整理"
    if high:
        return "高檔續強" if rising else "高檔轉弱" if falling else "高檔整理"
    return "中段回升" if rising else "中段走弱" if falling else "中段整理"


def _trend_from_history(daily: pd.DataFrame, live_pct: float) -> tuple[str, float, float, float]:
    hist = pd.to_numeric(daily.get("強勢股占比%", pd.Series(dtype=float)), errors="coerce").dropna().tolist()
    s = pd.Series(hist + [live_pct], dtype="float64")
    if len(s) < 60:
        return "資料累積中", float("nan"), float("nan"), float("nan")
    ma5 = float(s.tail(5).mean())
    ma20 = float(s.tail(20).mean())
    ma60 = float(s.tail(60).mean())
    ma5_prev5 = float(s.iloc[:-5].tail(5).mean()) if len(s) >= 10 else float("nan")
    ma20_prev5 = float(s.iloc[:-5].tail(20).mean()) if len(s) >= 25 else float("nan")
    vals = [ma5, ma20, ma60, ma5_prev5, ma20_prev5]
    if any(pd.isna(x) for x in vals):
        trend = "資料累積中"
    else:
        tol = max(0.05, abs(ma20) * 0.03)
        if ma5 > ma20 > ma60 and ma5 > ma5_prev5 and ma20 >= ma20_prev5:
            trend = "明顯增加"
        elif ma5 < ma20 < ma60 and ma5 < ma5_prev5 and ma20 <= ma20_prev5:
            trend = "明顯減少"
        elif ma5 > ma20 + tol:
            trend = "增加"
        elif ma5 < ma20 - tol:
            trend = "減少"
        else:
            trend = "持平"
    return trend, ma5, ma20, ma60


def _percentile_with_live(daily: pd.DataFrame, live_pct: float) -> float:
    hist = pd.to_numeric(daily.get("強勢股占比%", pd.Series(dtype=float)), errors="coerce").dropna()
    s = pd.concat([hist, pd.Series([live_pct])], ignore_index=True)
    return float(s.rank(method="average", pct=True).iloc[-1] * 100.0) if len(s) else 50.0


def _wade_state(score: float, water: str, trend: str) -> str:
    improving = trend in {"增加", "明顯增加"}
    weakening = trend in {"減少", "明顯減少"}
    if score >= 75: return "全面強勢"
    if score >= 65: return "偏多健康"
    if score >= 55: return "輪動偏多"
    if score >= 45: return "高檔分化" if water in {"偏高水位", "極高水位"} else ("改善中" if improving else "中性整理")
    if score >= 35: return "低檔修復中" if improving and water in {"極低水位", "偏低水位"} else ("內部轉弱" if weakening else "偏弱整理")
    if improving and water in {"極低水位", "偏低水位"}: return "低檔修復中"
    return "全面弱勢"


def _code_set(df: pd.DataFrame) -> set[str]:
    if df is None or df.empty or "代號" not in df.columns:
        return set()
    return set(df["代號"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip())


def build_snapshot(daily: pd.DataFrame, strong: pd.DataFrame, use_streamlit_cache: bool = True) -> dict:
    if use_streamlit_cache:
        base = load_baseline(str(BASELINE), BASELINE.stat().st_mtime_ns).copy()
    else:
        base = _load_baseline_raw(str(BASELINE)).copy()
    base["code"] = base["code"].astype(str)
    universe = base[["code", "market"]].drop_duplicates("code").to_dict("records")
    universe_json = json.dumps(universe, ensure_ascii=False, separators=(",", ":"))
    if use_streamlit_cache:
        q, errors = fetch_mis_quotes(universe_json)
    else:
        q, errors = _fetch_mis_quotes_resilient(universe_json)
    if q.empty:
        detail = "；".join(dict.fromkeys([str(e) for e in (errors or []) if str(e)]))[:500]
        raise RuntimeError(
            "MIS 全市場兩階段抓取仍無可用行情"
            + (f"｜{detail}" if detail else "")
        )
    raw_coverage = len(q) / len(universe) * 100.0 if universe else 100.0
    # Too little fresh coverage would make cross-sectional RS/Wade misleading.
    # Keep the previous successful snapshot instead of publishing a bad new one.
    if not use_streamlit_cache and raw_coverage < 70.0:
        detail = "；".join(dict.fromkeys([str(e) for e in (errors or []) if str(e)]))[:350]
        raise RuntimeError(
            f"MIS 全市場有效覆蓋僅 {raw_coverage:.1f}%（低於70%安全門檻）"
            + (f"｜{detail}" if detail else "")
        )

    x = base.merge(q, on="code", how="left")
    x["last"] = pd.to_numeric(x["last"], errors="coerce")
    x = x[x["last"].notna() & (x["last"] > 0)].copy()
    x["prev_close"] = pd.to_numeric(x["prev_close"], errors="coerce")
    x["ret_pct"] = (x["last"] / x["prev_close"] - 1.0) * 100.0

    for col in ["sum_close_19", "sum_close_49", "sum_close_59", "sum_close_199", "rs_base_close", "max_high_249", "min_low_249", "max_close_249", "min_close_249"]:
        x[col] = pd.to_numeric(x[col], errors="coerce")
    x["ma20_live"] = (x["sum_close_19"] + x["last"]) / 20.0
    x["ma50_live"] = (x["sum_close_49"] + x["last"]) / 50.0
    x["ma60_live"] = (x["sum_close_59"] + x["last"]) / 60.0
    x["ma200_live"] = (x["sum_close_199"] + x["last"]) / 200.0
    x["ret250"] = x["last"] / x["rs_base_close"] - 1.0
    x["eligible"] = x[["ma200_live", "rs_base_close", "max_high_249", "min_low_249"]].notna().all(axis=1)
    x["rs_live"] = float("nan")
    elig = x["eligible"] & x["ret250"].notna()
    if elig.any():
        x.loc[elig, "rs_live"] = x.loc[elig, "ret250"].rank(method="average", pct=True) * 100.0

    live_high = pd.to_numeric(x["high_live"], errors="coerce").fillna(x["last"])
    live_low = pd.to_numeric(x["low_live"], errors="coerce").fillna(x["last"])
    x["high250_live"] = pd.concat([x["max_high_249"], live_high], axis=1).max(axis=1)
    x["low250_live"] = pd.concat([x["min_low_249"], live_low], axis=1).min(axis=1)
    x["close_high250_live"] = pd.concat([x["max_close_249"], x["last"]], axis=1).max(axis=1)
    x["close_low250_live"] = pd.concat([x["min_close_249"], x["last"]], axis=1).min(axis=1)

    # MIS 'v' is displayed in board-lot style on MIS; use x1000 as an
    # intraday amount estimate.  The UI labels amount-dependent output as trial.
    x["amount_est"] = x["last"] * pd.to_numeric(x["volume_lots"], errors="coerce").fillna(0) * 1000.0
    x["rs85"] = x["eligible"] & (x["rs_live"] > RS_THRESHOLD)
    x["strong_live"] = (
        x["rs85"]
        & (x["last"] > x["ma200_live"])
        & (x["ma50_live"] > x["ma200_live"])
        & (x["amount_est"] > AMOUNT_MIN)
        & (x["last"] >= x["high250_live"] * (1.0 - MAX_BELOW_52W_HIGH))
        & (x["last"] >= x["low250_live"] * (1.0 + MIN_ABOVE_52W_LOW))
    )
    x["advance"] = x["ret_pct"] > 0
    x["decline"] = x["ret_pct"] < 0
    x["flat"] = x["ret_pct"].abs() < 1e-12
    x["limit_up"] = x["ret_pct"] >= 9.5
    x["limit_down"] = x["ret_pct"] <= -9.5
    x["new_high"] = x["last"] >= x["close_high250_live"] * 0.999999
    x["new_low"] = x["last"] <= x["close_low250_live"] * 1.000001

    eligible_count = int(x["eligible"].sum())
    strong_count = int(x["strong_live"].sum())
    strong_pct = strong_count / eligible_count * 100.0 if eligible_count else float("nan")
    advance_count = int(x["advance"].sum())
    decline_count = int(x["decline"].sum())
    breadth_den = advance_count + decline_count
    advance_ratio = advance_count / breadth_den * 100.0 if breadth_den else 50.0
    ma20_valid = pd.to_numeric(x["ma20_live"], errors="coerce").notna()
    above_ma20_ratio = (
        float((x.loc[ma20_valid, "last"] > x.loc[ma20_valid, "ma20_live"]).mean() * 100.0)
        if ma20_valid.any() else float("nan")
    )

    def _market_ratio(market: str) -> float:
        z = x[x["market"].astype(str).str.upper() == market]
        a, d = int(z["advance"].sum()), int(z["decline"].sum())
        return a / (a + d) * 100.0 if a + d else 50.0

    twse_adv = _market_ratio("TWSE")
    tpex_adv = _market_ratio("TPEX")
    # tolerate TPEx capitalization used in the cache
    if abs(tpex_adv - 50.0) < 1e-9:
        z = x[x["market"].astype(str).str.upper().isin(["TPEX", "OTC"])]
        a, d = int(z["advance"].sum()), int(z["decline"].sum())
        tpex_adv = a / (a + d) * 100.0 if a + d else 50.0

    full_rank = _percentile_with_live(daily, strong_pct) if not pd.isna(strong_pct) else 50.0
    water = _water_label(full_rank)
    trend, ma5_pct, ma20_pct, ma60_pct = _trend_from_history(daily, strong_pct)
    stage = _market_state(water, trend)

    prev_strong = _code_set(strong)
    cur_strong = set(x.loc[x["strong_live"], "code"].astype(str))
    retention = len(cur_strong & prev_strong) / len(prev_strong) * 100.0 if prev_strong else float("nan")
    new_leaders = cur_strong - prev_strong
    dropped = prev_strong - cur_strong

    total_amount = float(x["amount_est"].sum())
    hist_amount = pd.to_numeric(daily.get("總成交金額", pd.Series(dtype=float)), errors="coerce").dropna().tolist()
    amount_s = pd.Series(hist_amount[-19:] + [total_amount], dtype="float64")
    amount_ma20 = float(amount_s.mean()) if len(amount_s) >= 10 else float("nan")
    amount_ratio20 = total_amount / amount_ma20 * 100.0 if amount_ma20 and not pd.isna(amount_ma20) else float("nan")
    mean_return = float(pd.to_numeric(x["ret_pct"], errors="coerce").mean())
    twse_ret = float(pd.to_numeric(x.loc[x["market"].astype(str).str.upper() == "TWSE", "ret_pct"], errors="coerce").mean())
    tpex_mask = x["market"].astype(str).str.upper().isin(["TPEX", "OTC"])
    tpex_ret = float(pd.to_numeric(x.loc[tpex_mask, "ret_pct"], errors="coerce").mean())

    breadth_score = _ratio_strength(advance_ratio)
    activity = 50.0 if pd.isna(amount_ratio20) else max(0.0, min(100.0, 50.0 + (amount_ratio20 - 100.0) * 0.6))
    volume_eff = 0.7 * breadth_score + 0.3 * (activity if mean_return >= 0 else 100.0 - activity)
    # Same fallback spirit as formal RS when a market-index quote is absent.
    index_sync = _ratio_strength(twse_adv)
    tpex_relative = max(0.0, min(100.0, 50.0 + (tpex_ret - twse_ret) * 12.0)) if not (pd.isna(tpex_ret) or pd.isna(twse_ret)) else _ratio_strength(tpex_adv)
    leadership_score = _ratio_strength(retention)
    dir_score = {"明顯增加": 100.0, "增加": 75.0, "持平": 50.0, "減少": 25.0, "明顯減少": 0.0}.get(trend, 50.0)
    new_high = int(x["new_high"].sum())
    new_low = int(x["new_low"].sum())
    limit_up = int(x["limit_up"].sum())
    limit_down = int(x["limit_down"].sum())
    wade = round(
        0.20 * breadth_score
        + 0.20 * full_rank
        + 0.10 * _balance_score(new_high, new_low)
        + 0.07 * _balance_score(limit_up, limit_down)
        + 0.12 * dir_score
        + 0.10 * volume_eff
        + 0.08 * index_sync
        + 0.06 * tpex_relative
        + 0.07 * leadership_score,
        1,
    )
    wade_state = _wade_state(wade, water, trend)

    high_water = water in {"偏高水位", "極高水位"}
    weakening = trend in {"減少", "明顯減少"}
    improving = trend in {"增加", "明顯增加"}
    prev_wade5 = None
    if "Wade內部強度分數" in daily.columns and len(daily) >= 5:
        prev_wade5 = _f(pd.to_numeric(daily["Wade內部強度分數"], errors="coerce").iloc[-5])
    wade_change5 = wade - prev_wade5 if prev_wade5 is not None else None

    if high_water and weakening:
        risk = "🟠 盤中向上減碼觀察"
    elif high_water and wade < 50:
        risk = "🟠 盤中高檔內部轉弱"
    elif high_water and wade_change5 is not None and wade_change5 <= -12:
        risk = "🟡 盤中強度快速降溫"
    elif advance_ratio < 45 and tpex_adv < 45 and weakening:
        risk = "🟡 盤中廣度偏弱"
    else:
        risk = "未觸發"

    low_mid = water in {"極低水位", "偏低水位", "正常水位"}
    early = "🟢 盤中早期轉強" if low_mid and improving and advance_ratio >= 55 and tpex_adv >= 50 and new_high >= new_low else "未觸發"
    if risk.startswith("🟠"):
        action = "停止追高／汰弱留強；等收盤確認是否分批減碼"
    elif risk.startswith("🟡"):
        action = "停止擴大曝險，觀察是否惡化"
    elif early.startswith("🟢") or stage == "低檔回升":
        action = "提高關注／小部位分批試單；收盤再確認"
    elif stage in {"中段回升", "高檔續強"} and wade >= 60:
        action = "偏多持有／強股續抱；盤中不追價"
    elif stage in {"中段走弱", "低檔續弱"} and wade < 45:
        action = "防守降低曝險"
    else:
        action = "觀望等待"

    quote_coverage = len(x) / len(base) * 100.0 if len(base) else 0.0
    # 報價覆蓋率包含買一/賣一/前收備援；真正 MIS z 成交另外統計，
    # 避免把「有參考價」誤說成「有即時成交」。
    real_trade_coverage = (
        float(x.get("traded", pd.Series(False, index=x.index)).fillna(False).astype(bool).mean() * 100.0)
        if len(x) else 0.0
    )
    newest_time = "—"
    qt = x.get("quote_time")
    if qt is not None and qt.astype(str).str.len().gt(0).any():
        newest_time = qt.astype(str).replace("", pd.NA).dropna().max()

    # Per-stock diagnostics used by the background change radar and manual lookup.
    # These are still intraday trial values; formal inclusion/exit is confirmed after close.
    x["cond_rs"] = x["rs_live"] > RS_THRESHOLD
    x["cond_price_ma200"] = x["last"] > x["ma200_live"]
    x["cond_ma50_ma200"] = x["ma50_live"] > x["ma200_live"]
    x["cond_amount"] = x["amount_est"] > AMOUNT_MIN
    x["cond_52high"] = x["last"] >= x["high250_live"] * (1.0 - MAX_BELOW_52W_HIGH)
    x["cond_52low"] = x["last"] >= x["low250_live"] * (1.0 + MIN_ABOVE_52W_LOW)
    cond_cols = ["cond_rs", "cond_price_ma200", "cond_ma50_ma200", "cond_amount", "cond_52high", "cond_52low"]
    x["strong_cond_count"] = x[cond_cols].fillna(False).astype(int).sum(axis=1)

    def _missing_conditions(r):
        items = []
        if not bool(r.get("cond_rs")): items.append("RS>85")
        if not bool(r.get("cond_price_ma200")): items.append("股價>MA200")
        if not bool(r.get("cond_ma50_ma200")): items.append("MA50>MA200")
        if not bool(r.get("cond_amount")): items.append("成交額>3000萬")
        if not bool(r.get("cond_52high")): items.append("距52週高點≤25%")
        if not bool(r.get("cond_52low")): items.append("高於52週低點≥30%")
        return "、".join(items) if items else "已全部符合"

    x["strong_missing"] = x.apply(_missing_conditions, axis=1)
    # "即將符合"：只差一項正式盤中強勢條件，而且所有門檻都已接近。
    x["near_strong"] = (
        (~x["strong_live"])
        & (x["strong_cond_count"] >= 5)
        & (x["rs_live"] >= 80.0)
        & (x["amount_est"] >= 20_000_000.0)
        & (x["last"] > x["ma200_live"])
        & (x["ma50_live"] > x["ma200_live"])
        & (x["last"] >= x["high250_live"] * 0.70)
        & (x["last"] >= x["low250_live"] * 1.25)
    )
    x["duck_price_live"] = (x["last"] > x["ma20_live"]) & (x["ma20_live"] > x["ma60_live"])
    x["ma20_gap_pct"] = (x["last"] / x["ma20_live"] - 1.0) * 100.0

    live_cols = [
        "code", "name", "market", "last", "ret_pct", "rs_live", "ret250",
        "ma20_live", "ma50_live", "ma60_live", "ma200_live", "amount_est",
        "strong_live", "strong_cond_count", "strong_missing", "near_strong",
        "duck_price_live", "ma20_gap_pct", "best_bid", "best_ask",
        "price_source", "price_estimated", "traded", "quote_time", "quote_date",
    ]
    live_stocks = x[live_cols].copy()
    live_stocks = live_stocks.rename(columns={
        "code": "代號", "name": "名稱", "market": "市場", "last": "盤中價", "ret_pct": "漲跌幅%", "rs_live": "盤中RS", "ret250": "250日報酬試算",
        "ma20_live": "盤中MA20", "ma50_live": "盤中MA50", "ma60_live": "盤中MA60", "ma200_live": "盤中MA200", "amount_est": "盤中成交金額估算",
        "strong_live": "盤中強勢", "strong_cond_count": "強勢條件通過數", "strong_missing": "強勢尚缺條件", "near_strong": "即將強勢",
        "duck_price_live": "盤中鴨嘴價格結構", "ma20_gap_pct": "月線乖離率%",
        "best_bid": "最佳買價", "best_ask": "最佳賣價",
        "price_source": "價格來源", "price_estimated": "價格為估計", "traded": "MIS真實成交",
        "quote_time": "MIS報價時間", "quote_date": "MIS報價日期",
    })

    return {
        "base_date": str(base["base_date"].dropna().max()) if "base_date" in base.columns else "—",
        "snapshot_time": dt.datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "quote_time": newest_time,
        "coverage": quote_coverage,
        "real_trade_coverage": real_trade_coverage,
        "engine_version": ENGINE_VERSION,
        "errors": errors,
        "strong_count": strong_count,
        "eligible_count": eligible_count,
        "strong_pct": strong_pct,
        "full_rank": full_rank,
        "water": water,
        "trend": trend,
        "stage": stage,
        "ma5_pct": ma5_pct,
        "ma20_pct": ma20_pct,
        "ma60_pct": ma60_pct,
        "advance_count": advance_count,
        "decline_count": decline_count,
        "advance_ratio": advance_ratio,
        "above_ma20_ratio": above_ma20_ratio,
        "twse_adv": twse_adv,
        "tpex_adv": tpex_adv,
        "new_high": new_high,
        "new_low": new_low,
        "limit_up": limit_up,
        "limit_down": limit_down,
        "retention": retention,
        "new_leaders": new_leaders,
        "dropped": dropped,
        "amount_ratio20": amount_ratio20,
        "volume_eff": volume_eff,
        "index_sync": index_sync,
        "tpex_relative": tpex_relative,
        "leadership_score": leadership_score,
        "wade": wade,
        "wade_state": wade_state,
        "wade_change5": wade_change5,
        "risk": risk,
        "early": early,
        "action": action,
        "stocks": live_stocks,
    }



def _codes_from_live(snapshot: dict, col: str, watch_codes: set[str] | None = None) -> set[str]:
    stocks = snapshot.get("stocks") if snapshot else None
    if stocks is None or stocks.empty or col not in stocks.columns:
        return set()
    mask = stocks[col].fillna(False).astype(bool)
    if watch_codes is not None:
        mask = mask & stocks["代號"].astype(str).isin(watch_codes)
    return set(stocks.loc[mask, "代號"].astype(str))




def _plain_json_value(v):
    """Convert pandas/numpy-ish scalars into small JSON-safe values."""
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    if isinstance(v, (str, int, float, bool)) or v is None:
        return v
    try:
        x = v.item()
        if isinstance(x, (str, int, float, bool)) or x is None:
            return x
    except Exception:
        pass
    return str(v)


def _trade_state_transition_reason(old_state: str, new_state: str, row: dict) -> str:
    """Describe why a touched stock changed state, using the same live diagnostics shown in the UI."""
    old_state = str(old_state or "⏳ 等待")
    new_state = str(new_state or "⏳ 等待")
    if old_state == new_state:
        return "狀態延續"
    if old_state == "⏳ 等待" and new_state != "⏳ 等待":
        return f"即時條件成立：{new_state}"
    if old_state != "⏳ 等待" and new_state != "⏳ 等待":
        return f"分類轉換：{old_state} → {new_state}"

    right = _f(row.get("右側分數"))
    rs = _f(row.get("盤中RS"))
    heat = _f(row.get("過熱程度"))
    complete = _f(row.get("鴨嘴完成度%"))
    gate = str(row.get("市場閘門") or "")
    live_duck = bool(row.get("盤中正式鴨嘴", False))
    reasons = []
    if old_state.startswith("📈"):
        if right is not None and right < 65:
            reasons.append(f"右側 {right:.1f}<65")
        if "未確認" in gate:
            reasons.append("市場廣度閘門未確認")
        if heat is not None and heat >= 70:
            reasons.append(f"過熱 {heat:.1f}，離開進場區")
    elif old_state.startswith("🌱"):
        if complete is not None and complete < 80:
            reasons.append(f"鴨嘴完成度 {complete:.0f}%<80%")
        if rs is not None and rs < 70:
            reasons.append(f"RS {rs:.1f}<70")
        if "未確認" in gate:
            reasons.append("市場廣度閘門未確認")
        if heat is not None and heat >= 70:
            reasons.append(f"過熱 {heat:.1f}，離開試單區")
    elif old_state.startswith("🚀"):
        if not live_duck:
            reasons.append("正式鴨嘴結構暫時不成立")
        if right is not None and right < 65:
            reasons.append(f"右側 {right:.1f}<65")
    elif old_state.startswith("🟡"):
        if heat is not None and heat < 70:
            reasons.append(f"過熱降至 {heat:.1f}<70")
    elif old_state.startswith("🟠"):
        threshold = 90 if rs is not None and rs >= 85 else 80
        if heat is not None and heat < threshold:
            reasons.append(f"過熱降至 {heat:.1f}<{threshold}")
    elif old_state.startswith("🔴"):
        reasons.append("贏家退潮條件目前解除")
    return "；".join(reasons[:3]) if reasons else f"{old_state} → {new_state}"


def _trade_state_path_text(path) -> str:
    vals = [str(x) for x in (path or []) if str(x).strip()]
    return " → ".join(vals[-10:]) if vals else "—"


class IntradayBackgroundManager:
    """Process-shared scanner: full-market pulse + reactive intraday watch universe.

    v3.5.5 keeps every stock currently used by a盤中/即時 panel in the fast
    watch universe.  The full market is rescanned on a slower cadence to keep
    breadth/RS/Wade and discovery honest; active/event/manual stocks are then
    refreshed much more frequently in MIS batches.
    """

    def __init__(self):
        self._lock = threading.RLock()
        self._wake = threading.Event()
        self._stop = threading.Event()
        self._daily: pd.DataFrame | None = None
        self._strong: pd.DataFrame | None = None
        self._official_strong: set[str] = set()
        self._official_duck: set[str] = set()
        self._duck_watch: set[str] = set()
        self._formal_date = "—"
        self._config_sig = None
        self._interval = 30.0
        self._enabled = False
        self._force_once = False
        self._snapshot: dict | None = None
        self._previous_snapshot: dict | None = None
        self._events = {"new_strong": set(), "out_strong": set(), "new_duck": set(), "near": set(), "compare_label": ""}
        self._event_log: list[dict] = []
        self._version = 0
        self._running = False
        self._last_started: dt.datetime | None = None
        self._last_completed: dt.datetime | None = None
        self._last_error: str | None = None
        self._full_fail_streak = 0

        # v3.5.5: every stock shown by a盤中/即時 decision/event/manual view is
        # eligible for the reactive watch universe, not only the first 120 names.
        self._fast_watch_codes: list[str] = []
        self._fast_interval = 5.0
        self._fast_enabled = False
        self._fast_rows: pd.DataFrame | None = None
        self._fast_version = 0
        self._fast_running = False
        self._fast_last_started: dt.datetime | None = None
        self._fast_last_completed: dt.datetime | None = None
        self._fast_last_error: str | None = None
        self._fast_last_returned_count = 0
        self._fast_cycle = 0
        self._fast_fail_streak = 0

        # v3.5.8: process-shared per-stock state machine.  Once a stock has entered
        # any actionable radar state, it remains traceable for the rest of the day
        # even after it returns to "等待" or moves to another bucket.
        self._trade_state_day = dt.datetime.now(TZ).strftime("%Y-%m-%d")
        self._trade_states: dict[str, dict] = {}
        self._trade_state_events: list[dict] = []
        self._trade_state_version = 0
        self._trade_state_last_save = 0.0
        self._trade_state_persist_error: str | None = None
        self._load_trade_state_file_locked()

        self._thread = threading.Thread(target=self._loop, name="tw-intraday-background", daemon=True)
        self._thread.start()

    def configure(
        self,
        daily: pd.DataFrame,
        strong: pd.DataFrame,
        official_strong_codes: set[str] | None = None,
        official_duck_codes: set[str] | None = None,
        duck_watch_codes: set[str] | None = None,
        formal_date: str = "—",
        interval_seconds: int = 60,
        enabled: bool = True,
        ensure_once: bool = False,
    ) -> None:
        official_strong_codes = set(official_strong_codes or set())
        official_duck_codes = set(official_duck_codes or set())
        duck_watch_codes = set(duck_watch_codes or set())
        sig = (
            str(formal_date), len(daily), len(strong), len(official_strong_codes),
            len(official_duck_codes), len(duck_watch_codes),
            BASELINE.stat().st_mtime_ns if BASELINE.exists() else 0,
        )
        should_wake = False
        with self._lock:
            if sig != self._config_sig:
                old_date = self._formal_date
                self._daily = daily.copy(deep=True)
                self._strong = strong.copy(deep=True)
                self._official_strong = official_strong_codes
                self._official_duck = official_duck_codes
                self._duck_watch = duck_watch_codes
                self._formal_date = str(formal_date)
                self._config_sig = sig
                if old_date != "—" and old_date != self._formal_date:
                    self._snapshot = None
                    self._previous_snapshot = None
                    self._events = {"new_strong": set(), "out_strong": set(), "new_duck": set(), "near": set(), "compare_label": ""}
                    self._event_log = []
                    self._version = 0
                    self._fast_watch_codes = []
                    self._fast_rows = None
                    self._fast_version = 0
                    self._fast_last_completed = None
                    self._fast_last_started = None
                    self._fast_last_returned_count = 0
                    self._fast_cycle = 0
                should_wake = True
            new_interval = float(max(60, int(interval_seconds)))
            if new_interval != self._interval or bool(enabled) != self._enabled:
                self._interval = new_interval
                self._enabled = bool(enabled)
                should_wake = True
            if not self._enabled:
                self._fast_enabled = False
            if (ensure_once or self._enabled) and self._snapshot is None:
                self._force_once = True
                should_wake = True
        if should_wake:
            self._wake.set()

    def _reset_trade_state_day_locked(self, day: str | None = None) -> None:
        day = day or dt.datetime.now(TZ).strftime("%Y-%m-%d")
        if day == self._trade_state_day:
            return
        self._trade_state_day = day
        self._trade_states = {}
        self._trade_state_events = []
        self._trade_state_version += 1
        self._trade_state_persist_error = None
        try:
            if TRADE_STATE_FILE.exists():
                TRADE_STATE_FILE.unlink()
        except Exception:
            pass

    def _load_trade_state_file_locked(self) -> None:
        """Best-effort restore for browser/session refreshes and same-container restarts."""
        try:
            if not TRADE_STATE_FILE.exists():
                return
            payload = json.loads(TRADE_STATE_FILE.read_text(encoding="utf-8"))
            today = dt.datetime.now(TZ).strftime("%Y-%m-%d")
            if str(payload.get("day") or "") != today:
                try:
                    TRADE_STATE_FILE.unlink()
                except Exception:
                    pass
                return
            states = payload.get("states") or {}
            events = payload.get("events") or []
            if isinstance(states, dict):
                self._trade_states = {str(k): dict(v) for k, v in states.items() if isinstance(v, dict)}
            if isinstance(events, list):
                self._trade_state_events = [dict(x) for x in events[-2000:] if isinstance(x, dict)]
            self._trade_state_day = today
        except Exception as e:
            self._trade_state_persist_error = f"讀取狀態檔失敗：{type(e).__name__}: {e}"

    def _save_trade_state_file_locked(self, force: bool = False) -> None:
        now_ts = time.time()
        if not force and now_ts - self._trade_state_last_save < 20.0:
            return
        try:
            # Only touched stocks need durable continuity; untouched waiting stocks can
            # always be reconstructed from the next full-market snapshot.
            touched = {
                code: rec for code, rec in self._trade_states.items()
                if bool(rec.get("今日曾觸發", False))
            }
            payload = {
                "day": self._trade_state_day,
                "engine_version": ENGINE_VERSION,
                "saved_at": dt.datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
                "states": touched,
                "events": self._trade_state_events[-2000:],
            }
            tmp = TRADE_STATE_FILE.with_name(TRADE_STATE_FILE.name + ".tmp")
            tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            os.replace(tmp, TRADE_STATE_FILE)
            self._trade_state_last_save = now_ts
            self._trade_state_persist_error = None
        except Exception as e:
            self._trade_state_persist_error = f"寫入狀態檔失敗：{type(e).__name__}: {e}"

    def sync_trade_states(self, rows) -> None:
        """Synchronize app-classified live rows into one continuous state per stock.

        `rows` can be a DataFrame or a list of dicts.  Data-quality problems do not
        force a trading-state transition: the last valid state is kept and the row is
        marked non-actionable until a fresh MIS z quote returns.
        """
        if rows is None:
            return
        if isinstance(rows, pd.DataFrame):
            records = rows.to_dict("records")
        else:
            records = list(rows or [])
        if not records:
            return

        now = dt.datetime.now(TZ)
        day = now.strftime("%Y-%m-%d")
        now_text = now.strftime("%Y-%m-%d %H:%M:%S")
        clock = now.strftime("%H:%M:%S")
        changed = False
        with self._lock:
            self._reset_trade_state_day_locked(day)
            for raw in records:
                if not isinstance(raw, dict):
                    continue
                row = {str(k): _plain_json_value(v) for k, v in raw.items()}
                code = _norm_code(row.get("代號"))
                if not code:
                    continue
                candidate = str(row.get("狀態") or "⏳ 等待")
                data_issue = str(row.get("資料問題") or "").strip()
                price = _f(row.get("盤中價"))
                rec = self._trade_states.get(code)

                # Do not create 2,000 permanent waiting rows.  A stock enters the state
                # machine only after its first real actionable signal; from then on it
                # never disappears for the rest of that trading day.
                if rec is None:
                    if data_issue or candidate == "⏳ 等待":
                        continue
                    path = [candidate]
                    rec = {
                        "日期": day,
                        "代號": code,
                        "名稱": str(row.get("名稱") or ""),
                        "市場": str(row.get("市場") or ""),
                        "目前狀態": candidate,
                        "前一狀態": "—",
                        "今日路徑清單": path,
                        "今日路徑": _trade_state_path_text(path),
                        "今日曾觸發": True,
                        "首次觸發分類": candidate,
                        "首次觸發時間": clock,
                        "首次觸發價": price,
                        "本狀態觸發時間": clock,
                        "本狀態觸發價": price,
                        "狀態起始時間": clock,
                        "最近變更時間": clock,
                        "最近變更原因": f"首次進入：{candidate}",
                        "狀態變更次數": 0,
                    }
                    self._trade_states[code] = rec
                    self._trade_state_events.append({
                        "時間": clock, "代號": code, "名稱": rec["名稱"],
                        "前一狀態": "—", "目前狀態": candidate,
                        "價格": price, "原因": rec["最近變更原因"],
                    })
                    changed = True
                else:
                    rec["名稱"] = str(row.get("名稱") or rec.get("名稱") or "")
                    rec["市場"] = str(row.get("市場") or rec.get("市場") or "")

                # Always refresh diagnostics/data-health, even while a quote is stale.
                rec["最後同步時間"] = now_text
                rec["資料狀態"] = str(row.get("資料狀態") or "—")
                rec["資料層級"] = str(row.get("資料層級") or "—")
                rec["資料更新時間"] = str(row.get("資料更新時間") or "—")
                rec["MIS報價時間"] = str(row.get("MIS報價時間") or "—")
                rec["資料問題"] = data_issue
                rec["可即時操作"] = bool(not data_issue and candidate != "⏳ 等待")
                rec["下一步條件"] = str(row.get("下一步條件") or rec.get("下一步條件") or "—")
                for k in ["盤中RS", "右側分數", "過熱程度", "鴨嘴完成度%", "市場閘門", "盤中操作", "盤中正式鴨嘴"]:
                    if k in row:
                        rec[k] = row.get(k)

                # Keep last real price separate from an estimated book/previous-close reference.
                if not bool(row.get("價格為估計", False)) and price is not None:
                    rec["最新價"] = price
                    rec["最新真實價時間"] = now_text
                elif "最新價" not in rec and price is not None:
                    rec["最新價"] = price

                # A temporary stale/estimated quote must not make a stock look as if it
                # left the strategy.  Wait for the next fresh real quote before changing state.
                if data_issue:
                    rec["顯示狀態"] = f"{rec.get('目前狀態','⏳ 等待')}｜⚠️ 行情待刷新"
                    continue

                old_state = str(rec.get("目前狀態") or "⏳ 等待")
                if candidate != old_state:
                    reason = _trade_state_transition_reason(old_state, candidate, row)
                    rec["前一狀態"] = old_state
                    rec["目前狀態"] = candidate
                    path = list(rec.get("今日路徑清單") or [])
                    if not path or path[-1] != candidate:
                        path.append(candidate)
                    rec["今日路徑清單"] = path[-12:]
                    rec["今日路徑"] = _trade_state_path_text(rec["今日路徑清單"])
                    rec["狀態起始時間"] = clock
                    rec["本狀態觸發時間"] = clock
                    rec["本狀態觸發價"] = price
                    rec["最近變更時間"] = clock
                    rec["最近變更原因"] = reason
                    rec["狀態變更次數"] = int(rec.get("狀態變更次數", 0) or 0) + 1
                    if candidate != "⏳ 等待":
                        rec["今日曾觸發"] = True
                    self._trade_state_events.append({
                        "時間": clock, "代號": code, "名稱": rec.get("名稱", ""),
                        "前一狀態": old_state, "目前狀態": candidate,
                        "價格": price, "原因": reason,
                    })
                    changed = True

                rec["顯示狀態"] = str(rec.get("目前狀態") or "⏳ 等待")
                p0 = _f(rec.get("首次觸發價"))
                ps = _f(rec.get("本狀態觸發價"))
                pnow = _f(rec.get("最新價"))
                rec["首次觸發後漲跌%"] = ((pnow / p0 - 1.0) * 100.0) if p0 not in (None, 0) and pnow is not None else None
                rec["本狀態漲跌%"] = ((pnow / ps - 1.0) * 100.0) if ps not in (None, 0) and pnow is not None else None

            self._trade_state_events = self._trade_state_events[-3000:]
            self._trade_state_version += 1
            self._save_trade_state_file_locked(force=changed)

    def get_trade_state_table(self) -> pd.DataFrame:
        with self._lock:
            rows = [dict(v) for v in self._trade_states.values() if bool(v.get("今日曾觸發", False))]
        return pd.DataFrame(rows)

    def get_trade_state(self, code: str) -> dict | None:
        code = _norm_code(code)
        with self._lock:
            rec = self._trade_states.get(code)
            return dict(rec) if rec and bool(rec.get("今日曾觸發", False)) else None

    def set_fast_watch_codes(self, codes, interval_seconds: int = 5, enabled: bool = True, max_codes: int = 2500) -> None:
        ordered = []
        seen = set()
        for code in list(codes or []):
            code = _norm_code(code)
            if code and code not in seen:
                ordered.append(code); seen.add(code)
            if len(ordered) >= max(1, int(max_codes)):
                break
        interval = float(max(5, int(interval_seconds)))
        should_wake = False
        with self._lock:
            active = bool(enabled and self._enabled and ordered)
            if ordered != self._fast_watch_codes:
                self._fast_watch_codes = ordered
                # 不再因候選名單增減就把整張快速行情清空。
                # 保留仍在追蹤名單中的上一筆快刷資料，下一輪再逐碼覆蓋。
                if self._fast_rows is not None and not self._fast_rows.empty and "代號" in self._fast_rows.columns:
                    keep = self._fast_rows.copy()
                    keep["代號"] = keep["代號"].astype(str).map(_norm_code)
                    self._fast_rows = keep[keep["代號"].isin(set(ordered))].copy()
                else:
                    self._fast_rows = None
                self._fast_last_completed = None
                self._fast_last_started = None
                should_wake = True
            if interval != self._fast_interval or active != self._fast_enabled:
                self._fast_interval = interval
                self._fast_enabled = active
                should_wake = True
        if should_wake:
            self._wake.set()

    def request_refresh(self) -> None:
        with self._lock:
            self._force_once = True
        self._wake.set()

    def get_state(self) -> dict:
        with self._lock:
            return {
                "snapshot": self._snapshot,
                "previous_snapshot": self._previous_snapshot,
                "events": {
                    "new_strong": set(self._events.get("new_strong", set())),
                    "out_strong": set(self._events.get("out_strong", set())),
                    "new_duck": set(self._events.get("new_duck", set())),
                    "near": set(self._events.get("near", set())),
                    "compare_label": self._events.get("compare_label", ""),
                },
                "event_log": list(self._event_log),
                "version": self._version,
                "running": self._running,
                "enabled": self._enabled,
                "interval_seconds": self._interval,
                "last_started": self._last_started.strftime("%Y-%m-%d %H:%M:%S") if self._last_started else None,
                "last_completed": self._last_completed.strftime("%Y-%m-%d %H:%M:%S") if self._last_completed else None,
                "last_error": self._last_error,
                "full_fail_streak": int(getattr(self, "_full_fail_streak", 0)),
                "fast_rows": None if self._fast_rows is None else self._fast_rows.copy(deep=False),
                "fast_watch_codes": list(self._fast_watch_codes),
                "fast_watch_count": len(self._fast_watch_codes),
                "fast_returned_count": int(self._fast_last_returned_count),
                "fast_coverage": (float(self._fast_last_returned_count) / len(self._fast_watch_codes) * 100.0) if self._fast_watch_codes else 0.0,
                "fast_cycle": int(self._fast_cycle),
                "engine_version": ENGINE_VERSION,
                "fast_version": self._fast_version,
                "fast_running": self._fast_running,
                "fast_enabled": self._fast_enabled,
                "fast_interval_seconds": self._fast_interval,
                "fast_last_started": self._fast_last_started.strftime("%Y-%m-%d %H:%M:%S") if self._fast_last_started else None,
                "fast_last_completed": self._fast_last_completed.strftime("%Y-%m-%d %H:%M:%S") if self._fast_last_completed else None,
                "fast_last_error": self._fast_last_error,
                "fast_fail_streak": int(getattr(self, "_fast_fail_streak", 0)),
                "trade_state_day": self._trade_state_day,
                "trade_state_version": int(self._trade_state_version),
                "trade_state_codes": [code for code, rec in self._trade_states.items() if bool(rec.get("今日曾觸發", False))],
                "trade_state_count": sum(1 for rec in self._trade_states.values() if bool(rec.get("今日曾觸發", False))),
                "trade_state_events": list(self._trade_state_events[-500:]),
                "trade_state_persist_error": self._trade_state_persist_error,
            }

    def _accept_snapshot(self, snap: dict) -> None:
        prev = self._snapshot
        current_strong = _codes_from_live(snap, "盤中強勢")
        current_duck = _codes_from_live(snap, "盤中鴨嘴價格結構", self._duck_watch)
        near = _codes_from_live(snap, "即將強勢")
        if prev is None:
            compare_strong = set(self._official_strong)
            compare_duck = set(self._official_duck)
            compare_label = f"相對 {self._formal_date} 正式"
        else:
            compare_strong = _codes_from_live(prev, "盤中強勢")
            compare_duck = _codes_from_live(prev, "盤中鴨嘴價格結構", self._duck_watch)
            compare_label = f"相對上一批約 {int(self._interval)} 秒前"
        new_strong = current_strong - compare_strong
        out_strong = compare_strong - current_strong
        new_duck = current_duck - compare_duck

        stocks = snap.get("stocks")
        name_map = {}
        if stocks is not None and not stocks.empty:
            name_map = dict(zip(stocks["代號"].astype(str), stocks.get("名稱", pd.Series("", index=stocks.index)).astype(str)))
        t = str(snap.get("snapshot_time") or "")[-8:]
        for typ, codes in [
            ("🔥 新進強勢", new_strong),
            ("⚠️ 暫時退出", out_strong),
            ("🦆 鴨嘴價格結構新進", new_duck),
        ]:
            for code in sorted(codes):
                self._event_log.append({"時間": t, "類型": typ, "代號": code, "名稱": name_map.get(code, "")})
        self._event_log = self._event_log[-300:]
        self._previous_snapshot = prev
        self._snapshot = snap
        self._events = {
            "new_strong": new_strong,
            "out_strong": out_strong,
            "new_duck": new_duck,
            "near": near,
            "compare_label": compare_label,
        }
        self._version += 1
        # 全市場重掃完成後，保留既有快速行情，避免每30秒瞬間退回舊快照。
        # 只把 fast anchor 清掉，讓下一個 loop 立刻用新全市場基準重刷追蹤股。
        self._fast_last_completed = None
        self._fast_last_started = None

    def _loop(self) -> None:
        while not self._stop.is_set():
            with self._lock:
                ready = self._daily is not None and self._strong is not None and BASELINE.exists()
                enabled = self._enabled
                force = self._force_once
                interval = self._interval
                last_completed = self._last_completed
                last_started = self._last_started
                daily = self._daily
                strong = self._strong
                fast_enabled = self._fast_enabled
                fast_interval = self._fast_interval
                fast_last_completed = self._fast_last_completed
                fast_last_started = self._fast_last_started
                fast_codes = list(self._fast_watch_codes)
                snap_for_fast = self._snapshot
            now = dt.datetime.now(TZ)
            full_anchors=[x for x in [last_completed,last_started] if x is not None]
            anchor=max(full_anchors) if full_anchors else None
            due = anchor is None or (now - anchor).total_seconds() >= interval

            # Full scan has priority when due; it refreshes the market gate and candidate discovery universe.
            if ready and (force or (enabled and due)):
                with self._lock:
                    self._force_once = False
                    self._running = True
                    self._last_started = dt.datetime.now(TZ)
                    self._last_error = None
                    daily_local = daily.copy(deep=False)
                    strong_local = strong.copy(deep=False)
                try:
                    snap = build_snapshot(daily_local, strong_local, use_streamlit_cache=False)
                except Exception as e:
                    with self._lock:
                        self._full_fail_streak += 1
                        self._last_error = f"{type(e).__name__}: {e}"
                        self._running = False
                else:
                    with self._lock:
                        self._accept_snapshot(snap)
                        self._last_completed = dt.datetime.now(TZ)
                        self._full_fail_streak = 0
                        self._last_error = None
                        self._running = False
                continue

            fast_anchors=[x for x in [fast_last_completed,fast_last_started] if x is not None]
            fast_anchor=max(fast_anchors) if fast_anchors else None
            fast_due = fast_anchor is None or (now - fast_anchor).total_seconds() >= fast_interval
            if ready and enabled and fast_enabled and fast_codes and snap_for_fast is not None and fast_due:
                with self._lock:
                    self._fast_running = True
                    self._fast_last_started = dt.datetime.now(TZ)
                    self._fast_last_error = None
                    fast_snapshot_local = snap_for_fast
                    fast_codes_local = list(fast_codes)
                try:
                    rows = refresh_watchlist(fast_snapshot_local, fast_codes_local, max_codes=2500)
                except Exception as e:
                    with self._lock:
                        self._fast_fail_streak += 1
                        self._fast_last_error = f"{type(e).__name__}: {e}"
                        self._fast_running = False
                else:
                    with self._lock:
                        now_done = dt.datetime.now(TZ)
                        self._fast_fail_streak = 0
                        self._fast_last_error = None
                        # MIS 批次可能偶爾少回個別股票；不能因此讓該股退回30秒甚至更舊的全市場列。
                        # 將本輪成功回傳逐碼合併到上一輪快速行情，未回傳者保留上一筆並由 UI 顯示資料年齡。
                        if rows is not None and not rows.empty:
                            nr = rows.copy()
                            nr["代號"] = nr["代號"].astype(str).map(_norm_code)
                            nr["快速追蹤批次"] = self._fast_cycle + 1
                            nr["快速抓取完成時間"] = now_done.strftime("%Y-%m-%d %H:%M:%S")
                            if self._fast_rows is None or self._fast_rows.empty:
                                merged = nr
                            else:
                                old = self._fast_rows.copy()
                                old["代號"] = old["代號"].astype(str).map(_norm_code)
                                merged = pd.concat([old, nr], ignore_index=True)
                                merged = merged.drop_duplicates("代號", keep="last")
                            if self._fast_watch_codes:
                                merged = merged[merged["代號"].isin(set(self._fast_watch_codes))].copy()
                            self._fast_rows = merged
                            self._fast_last_returned_count = int(len(nr))
                        else:
                            self._fast_last_returned_count = 0
                        self._fast_cycle += 1
                        self._fast_version += 1
                        self._fast_last_completed = now_done
                        self._fast_running = False
                continue

            waits = [15.0]
            if enabled and anchor is not None:
                waits.append(max(1.0, interval - (now - anchor).total_seconds()))
            if enabled and fast_enabled and fast_codes and snap_for_fast is not None and fast_anchor is not None:
                waits.append(max(0.5, fast_interval - (now - fast_anchor).total_seconds()))
            wait_s = max(0.5, min(15.0, min(waits)))
            self._wake.clear()
            self._wake.wait(wait_s)


def refresh_watchlist(snapshot: dict, codes, max_codes: int = 2500) -> pd.DataFrame:
    """Batch-refresh a small actionable watchlist without rescanning the whole market.

    The full-market snapshot remains the ranking anchor. Updated watched stocks are
    reranked together against the unchanged 250-day returns of the rest of the market.
    The caller can pass the complete intraday-active universe; requests are still batched
    (80 symbols per MIS request) rather than one request per stock.  The app dynamically
    slows the fast interval when the active universe becomes very large.
    """
    if not snapshot or snapshot.get("stocks") is None:
        raise RuntimeError("尚未有背景市場快照")
    stocks = snapshot["stocks"].copy()
    if stocks.empty or "代號" not in stocks.columns:
        return pd.DataFrame()
    stocks["代號"] = stocks["代號"].astype(str).map(_norm_code)

    ordered = []
    seen = set()
    for c in list(codes or []):
        c = _norm_code(c)
        if c and c not in seen:
            ordered.append(c); seen.add(c)
        if len(ordered) >= max(1, int(max_codes)):
            break
    if not ordered:
        return pd.DataFrame()

    hit = stocks[stocks["代號"].isin(ordered)].copy()
    if hit.empty:
        return pd.DataFrame()
    market_map = dict(zip(hit["代號"].astype(str), hit.get("市場", pd.Series("TWSE", index=hit.index)).astype(str)))
    universe = [{"code": c, "market": market_map.get(c, "TWSE")} for c in ordered if c in market_map]
    universe_json = json.dumps(universe, ensure_ascii=False, separators=(",", ":"))
    q, errors = _fetch_mis_quotes_raw(universe_json, batch_size=80, workers=3)
    if q.empty:
        raise RuntimeError(errors[0] if errors else "MIS 沒有回傳快速追蹤行情")
    q["code"] = q["code"].astype(str).map(_norm_code)
    q = q.drop_duplicates("code", keep="last").set_index("code")

    base = _load_baseline_raw(str(BASELINE)).copy()
    base["code"] = base["code"].astype(str).map(_norm_code)
    base = base.drop_duplicates("code", keep="last").set_index("code")

    # Rebuild watched ret250 first, then rank them against the last full-market snapshot.
    ret_all = pd.to_numeric(stocks.get("250日報酬試算"), errors="coerce")
    ret_all.index = stocks["代號"].astype(str)
    updated_ret = {}
    for code in ordered:
        if code not in q.index or code not in base.index:
            continue
        last = _f(q.loc[code].get("last"))
        rs_base = _f(base.loc[code].get("rs_base_close"))
        if last is not None and rs_base not in (None, 0):
            updated_ret[code] = last / rs_base - 1.0
    for code, ret in updated_ret.items():
        ret_all.loc[code] = ret
    rs_rank = ret_all.rank(method="average", pct=True) * 100.0

    old_map = {str(r.get("代號")): r for r in stocks.to_dict("records")}
    rows = []
    for code in ordered:
        if code not in q.index or code not in base.index or code not in old_map:
            continue
        qr = q.loc[code]
        b = base.loc[code]
        old = old_map[code]
        last = _f(qr.get("last"))
        prev_close = _f(b.get("prev_close"))
        if last is None or prev_close in (None, 0):
            continue
        ma20 = (_f(b.get("sum_close_19"), 0) + last) / 20.0
        ma50 = (_f(b.get("sum_close_49"), 0) + last) / 50.0
        ma60 = (_f(b.get("sum_close_59"), 0) + last) / 60.0
        ma200 = (_f(b.get("sum_close_199"), 0) + last) / 200.0
        ret250 = updated_ret.get(code)
        rs_live = _f(rs_rank.get(code)) if code in rs_rank.index else _f(old.get("盤中RS"))

        high_live = _f(qr.get("high_live"), last)
        low_live = _f(qr.get("low_live"), last)
        high250 = max(_f(b.get("max_high_249"), last), high_live)
        low250 = min(_f(b.get("min_low_249"), last), low_live)
        amount_est = last * (_f(qr.get("volume_lots"), 0.0) or 0.0) * 1000.0
        conds = {
            "RS>85": rs_live is not None and rs_live > RS_THRESHOLD,
            "股價>MA200": last > ma200,
            "MA50>MA200": ma50 > ma200,
            "成交額>3000萬": amount_est > AMOUNT_MIN,
            "距52週高點≤25%": last >= high250 * (1.0 - MAX_BELOW_52W_HIGH),
            "高於52週低點≥30%": last >= low250 * (1.0 + MIN_ABOVE_52W_LOW),
        }
        count = sum(bool(v) for v in conds.values())
        missing = [k for k, v in conds.items() if not v]
        strong_live = count == 6
        near = (
            (not strong_live) and count >= 5 and (rs_live or 0) >= 80.0
            and amount_est >= 20_000_000.0 and last > ma200 and ma50 > ma200
            and last >= high250 * 0.70 and last >= low250 * 1.25
        )
        rows.append({
            "代號": code,
            "名稱": str(old.get("名稱") or qr.get("name_live") or ""),
            "市場": str(old.get("市場") or market_map.get(code) or "TWSE"),
            "盤中價": last,
            "漲跌幅%": (last / prev_close - 1.0) * 100.0,
            "最佳買價": _f(qr.get("best_bid")),
            "最佳賣價": _f(qr.get("best_ask")),
            "價格來源": str(qr.get("price_source") or "—"),
            "價格為估計": bool(qr.get("price_estimated", False)),
            "MIS報價時間": str(qr.get("quote_time") or "—"),
            "MIS報價日期": str(qr.get("quote_date") or "—"),
            "盤中RS": rs_live,
            "250日報酬試算": ret250,
            "盤中MA20": ma20,
            "盤中MA50": ma50,
            "盤中MA60": ma60,
            "盤中MA200": ma200,
            "盤中成交金額估算": amount_est,
            "盤中強勢": strong_live,
            "強勢條件通過數": count,
            "強勢尚缺條件": "、".join(missing) if missing else "已全部符合",
            "即將強勢": near,
            "盤中鴨嘴價格結構": bool(last > ma20 > ma60),
            "月線乖離率%": (last / ma20 - 1.0) * 100.0 if ma20 else None,
            "快速追蹤時間": dt.datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
            "快速追蹤epoch": time.time(),
            "行情引擎版本": ENGINE_VERSION,
        })
    return pd.DataFrame(rows)


@st.cache_resource(show_spinner=False)
def get_background_manager(engine_generation: str = ENGINE_VERSION) -> IntradayBackgroundManager:
    # generation 參數是 cache key 的一部分；部署新版時強制建立新的 manager/thread，
    # 避免 Streamlit 沿用舊版 cache_resource 背景執行緒。
    return IntradayBackgroundManager()


def refresh_single_stock(snapshot: dict, code: str) -> dict:
    """Fetch one stock using the same batch-refresh core used by the fast radar."""
    code = _norm_code(code)
    rows = refresh_watchlist(snapshot, [code], max_codes=1)
    if rows.empty:
        raise RuntimeError(f"MIS 沒有回傳股票 {code} 的可用行情")
    r = rows.iloc[0].to_dict()
    r["單檔抓取時間"] = r.get("快速追蹤時間") or dt.datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
    return r


def _fmt(v, digits=1, suffix=""):
    x = _f(v)
    return "—" if x is None else f"{x:,.{digits}f}{suffix}"


def _formal_num(row, name):
    try:
        return _f(row.get(name))
    except Exception:
        return None


def _render_once(daily: pd.DataFrame, strong: pd.DataFrame, all_ok: pd.DataFrame, pre: pd.DataFrame, rs_latest: pd.Series, snap_override: dict | None = None):
    now = dt.datetime.now(TZ)
    if now.weekday() >= 5 or not (dt.time(8, 30) <= now.time() <= dt.time(14, 0)):
        st.info("目前不在台股一般交易附近時段；盤中雷達仍可手動讀取 MIS 的最新可用行情，但正式判讀請以收盤資料為準。")

    if snap_override is not None:
        snap = snap_override
    else:
        try:
            snap = build_snapshot(daily, strong)
        except Exception as e:
            st.error(f"盤中行情暫時無法建立：{e}")
            st.caption("正式盤後系統不受影響。若剛部署此功能，先手動執行一次『每日台股自動更新』，讓 Actions 產生 intraday_baseline.pkl.gz。")
            return

    if snap["coverage"] < 80:
        st.warning(f"盤中報價覆蓋率只有 {snap['coverage']:.1f}%，這次訊號僅供參考，不建議用來調整部位。")
    elif snap["errors"]:
        st.warning(f"有 {len(snap['errors'])} 個行情批次失敗，但目前整體覆蓋率仍為 {snap['coverage']:.1f}%。")

    formal_strong = _formal_num(rs_latest, "強勢股檔數")
    formal_wade = _formal_num(rs_latest, "Wade內部強度分數")
    formal_adv = _formal_num(rs_latest, "上漲比例%")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("盤中 RS 強勢股", f"{snap['strong_count']:,} 檔", None if formal_strong is None else f"較前收 {snap['strong_count']-formal_strong:+.0f}")
    c2.metric("盤中 Wade 試算", f"{snap['wade']:.1f}", None if formal_wade is None else f"較前收 {snap['wade']-formal_wade:+.1f}")
    c3.metric("上漲比例", f"{snap['advance_ratio']:.1f}%", None if formal_adv is None else f"較前收 {snap['advance_ratio']-formal_adv:+.1f}pct")
    c4.metric("52週新高／新低", f"{snap['new_high']}／{snap['new_low']}")
    c5.metric("上市／上櫃上漲", f"{snap['twse_adv']:.1f}%／{snap['tpex_adv']:.1f}%")
    c6.metric("報價覆蓋率", f"{snap['coverage']:.1f}%")

    st.markdown(
        f"**📡 盤中試算：{snap['stage']}｜{snap['water']}｜強勢股方向 {snap['trend']}**  \n"
        f"Wade：**{snap['wade']:.1f} / 100（{snap['wade_state']}）**　｜　"
        f"早期轉強：**{snap['early']}**　｜　風險：**{snap['risk']}**  \n"
        f"**操作語言：{snap['action']}**"
    )
    st.caption(
        f"基準日 {snap['base_date']}｜抓取 {snap['snapshot_time']}｜MIS 最新欄位時間 {snap['quote_time']}｜"
        "所有盤中訊號皆為試算；13:30 收盤與盤後正式更新後才寫入歷史。盤中成交金額依 MIS 累計量換算，屬估算欄位。"
    )

    a, b, c, d = st.columns(4)
    a.metric("漲停／跌停近似", f"{snap['limit_up']}／{snap['limit_down']}")
    b.metric("主流延續率", _fmt(snap['retention'], 1, "%"), f"新領頭 {len(snap['new_leaders'])} 檔")
    c.metric("成交金額／20日", _fmt(snap['amount_ratio20'], 1, "%"), f"量價效率 {_fmt(snap['volume_eff'],1)}")
    d.metric("5日 Wade 變化", _fmt(snap['wade_change5'], 1), f"RS 水位 P{snap['full_rank']:.1f}")

    tab1, tab2, tab3 = st.tabs(["盤中新進／退出強勢", "盤中個股技術", "鴨嘴候選即時價格"])
    stocks = snap["stocks"].copy()
    strong_now = set(stocks.loc[stocks["盤中強勢"].fillna(False), "代號"].astype(str))
    prev = _code_set(strong)
    with tab1:
        new_codes = strong_now - prev
        out_codes = prev - strong_now
        cc1, cc2 = st.columns(2)
        with cc1:
            st.write(f"**盤中新增強勢：{len(new_codes)} 檔**")
            z = stocks[stocks["代號"].astype(str).isin(new_codes)].sort_values("盤中RS", ascending=False)
            st.dataframe(z.head(80), hide_index=True, use_container_width=True, height=420)
        with cc2:
            st.write(f"**前收強勢、盤中暫時退出：{len(out_codes)} 檔**")
            z = stocks[stocks["代號"].astype(str).isin(out_codes)].sort_values("漲跌幅%")
            st.dataframe(z.head(80), hide_index=True, use_container_width=True, height=420)
    with tab2:
        q = st.text_input("搜尋代號／名稱（盤中）", key="intraday_stock_search")
        z = stocks.copy()
        if q.strip():
            qq = q.strip()
            z = z[z["代號"].astype(str).str.contains(qq, case=False, na=False) | z["名稱"].astype(str).str.contains(qq, case=False, na=False)]
        z = z.sort_values(["盤中強勢", "盤中RS", "漲跌幅%"], ascending=[False, False, False], na_position="last")
        st.dataframe(z.head(300), hide_index=True, use_container_width=True, height=600)
    with tab3:
        candidate_codes = _code_set(all_ok) | _code_set(pre)
        z = stocks[stocks["代號"].astype(str).isin(candidate_codes)].copy()
        if not z.empty:
            if "月線乖離率%" not in z.columns:
                z["月線乖離率%"] = (pd.to_numeric(z["盤中價"], errors="coerce") / pd.to_numeric(z["盤中MA20"], errors="coerce") - 1) * 100
            if "盤中鴨嘴價格結構" not in z.columns:
                z["盤中鴨嘴價格結構"] = (pd.to_numeric(z["盤中價"], errors="coerce") > pd.to_numeric(z["盤中MA20"], errors="coerce")) & (pd.to_numeric(z["盤中MA20"], errors="coerce") > pd.to_numeric(z["盤中MA60"], errors="coerce"))
            z = z.sort_values(["盤中鴨嘴價格結構", "月線乖離率%"], ascending=[False, False])
        st.caption("這裡即時計算價格、MA20/MA60 與乖離；正式『新進／退出』仍以鴨嘴完整條件在盤後確認。")
        show = [c for c in ["代號", "名稱", "市場", "盤中價", "漲跌幅%", "盤中MA20", "盤中MA60", "月線乖離率%", "盤中鴨嘴價格結構", "盤中RS"] if c in z.columns]
        st.dataframe(z[show].head(250), hide_index=True, use_container_width=True, height=600)


def render_intraday_panel(daily: pd.DataFrame, strong: pd.DataFrame, all_ok: pd.DataFrame, pre: pd.DataFrame, rs_latest: pd.Series):
    st.subheader("📡 盤中即時市場雷達")
    st.caption("v3.5.5：全市場約30秒更新廣度/RS/Wade；盤中活躍個股另以動態5–20秒批次追蹤。")
    if not BASELINE.exists():
        st.warning("尚未找到 intraday_baseline.pkl.gz。部署這個版本後，請手動跑一次『每日台股自動更新』。")
        return

    auto = st.toggle("背景自動掃描", value=True, key="intraday_auto_refresh")
    seconds = st.select_slider("背景掃描頻率", options=[30, 45, 60, 90], value=30, format_func=lambda x: f"{x} 秒")
    manager = get_background_manager()
    official_strong = _code_set(strong)
    official_duck = _code_set(all_ok)
    duck_watch = official_duck | _code_set(pre)
    formal_date = str(rs_latest.get("日期", "—")) if hasattr(rs_latest, "get") else "—"
    manager.configure(
        daily, strong,
        official_strong_codes=official_strong,
        official_duck_codes=official_duck,
        duck_watch_codes=duck_watch,
        formal_date=formal_date,
        interval_seconds=seconds,
        enabled=auto,
        ensure_once=True,
    )
    if st.button("🔄 背景立即重掃", type="primary", use_container_width=True):
        manager.request_refresh()
        st.toast("已要求背景重掃；目前畫面繼續使用上一批完成資料。")

    def _show_state():
        state = manager.get_state()
        snap = state.get("snapshot")
        if not snap:
            if state.get("running"):
                st.info("背景正在建立第一批行情；你可以繼續操作，不需要等待。")
            elif state.get("last_error"):
                st.error(f"背景行情建立失敗：{state.get('last_error')}")
            else:
                st.info("背景掃描已啟動。")
            return
        if state.get("running"):
            st.caption(f"🔄 下一批背景處理中｜目前顯示：{state.get('last_completed') or snap.get('snapshot_time','—')}")
        else:
            st.caption(f"🟢 最近完成：{state.get('last_completed') or snap.get('snapshot_time','—')}")
        _render_once(daily, strong, all_ok, pre, rs_latest, snap_override=snap)

    if hasattr(st, "fragment") and auto:
        @st.fragment(run_every="30s")
        def _frag():
            _show_state()
        _frag()
    else:
        _show_state()

