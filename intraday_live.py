# -*- coding: utf-8 -*-
"""Intraday trial analysis for the Taiwan-stock Streamlit dashboard.

Design goals
------------
1. Never overwrite formal daily files (rs_latest.xlsx / duck_latest.xlsx).
2. Reuse the previous close's compact rolling baseline.
3. Fetch TWSE/TPEx MIS quotes during the session and recompute the same RS
   strong-stock core plus a Wade-style *intraday trial* score.
4. Make trial/estimated fields explicit in the UI.

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
import time

import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
BASELINE = APP_DIR / "intraday_baseline.pkl.gz"
BASELINE_STATUS = APP_DIR / "intraday_baseline_status.json"
MIS_URL = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp"
TZ = ZoneInfo("Asia/Taipei")

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


def _channel(code: str, market: str) -> str:
    m = str(market).strip().upper()
    prefix = "otc" if m in {"TPEX", "OTC", "上櫃"} else "tse"
    return f"{prefix}_{code}.tw"


def _fetch_batch(channels: list[str], timeout: int = 8) -> tuple[list[dict], str | None]:
    if not channels:
        return [], None
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
        return payload.get("msgArray", []) or [], None
    except Exception as e:
        return [], f"{type(e).__name__}: {e}"


@st.cache_data(ttl=25, show_spinner=False)
def fetch_mis_quotes(universe_json: str, batch_size: int = 80, workers: int = 6) -> tuple[pd.DataFrame, list[str]]:
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
            # ch often looks like 2330.tw; keep a conservative fallback.
            code = ch.split(".")[0].split("_")[-1]
        if not code:
            continue
        prev = _f(m.get("y"))
        last = _f(m.get("z"))
        traded = last is not None
        if last is None:
            # No trade yet: use the reference close only so breadth doesn't
            # falsely classify the stock as missing.  Mark it as not traded.
            last = prev
        if last is None:
            continue
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


@st.cache_data(ttl=300, show_spinner=False)
def load_baseline(path_s: str, mtime_ns: int) -> pd.DataFrame:
    return pd.read_pickle(path_s, compression="gzip")


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


def build_snapshot(daily: pd.DataFrame, strong: pd.DataFrame) -> dict:
    base = load_baseline(str(BASELINE), BASELINE.stat().st_mtime_ns).copy()
    base["code"] = base["code"].astype(str)
    universe = base[["code", "market"]].drop_duplicates("code").to_dict("records")
    q, errors = fetch_mis_quotes(json.dumps(universe, ensure_ascii=False, separators=(",", ":")))
    if q.empty:
        raise RuntimeError("TWSE MIS did not return any usable quotes")

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
    newest_time = "—"
    qt = x.get("quote_time")
    if qt is not None and qt.astype(str).str.len().gt(0).any():
        newest_time = qt.astype(str).replace("", pd.NA).dropna().max()

    live_cols = ["code", "name", "market", "last", "ret_pct", "rs_live", "ma20_live", "ma50_live", "ma60_live", "ma200_live", "amount_est", "strong_live"]
    live_stocks = x[live_cols].copy()
    live_stocks = live_stocks.rename(columns={
        "code": "代號", "name": "名稱", "market": "市場", "last": "盤中價", "ret_pct": "漲跌幅%", "rs_live": "盤中RS",
        "ma20_live": "盤中MA20", "ma50_live": "盤中MA50", "ma60_live": "盤中MA60", "ma200_live": "盤中MA200", "amount_est": "盤中成交金額估算", "strong_live": "盤中強勢",
    })

    return {
        "base_date": str(base["base_date"].dropna().max()) if "base_date" in base.columns else "—",
        "snapshot_time": dt.datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S"),
        "quote_time": newest_time,
        "coverage": quote_coverage,
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


def _fmt(v, digits=1, suffix=""):
    x = _f(v)
    return "—" if x is None else f"{x:,.{digits}f}{suffix}"


def _formal_num(row, name):
    try:
        return _f(row.get(name))
    except Exception:
        return None


def _render_once(daily: pd.DataFrame, strong: pd.DataFrame, all_ok: pd.DataFrame, pre: pd.DataFrame, rs_latest: pd.Series):
    now = dt.datetime.now(TZ)
    if now.weekday() >= 5 or not (dt.time(8, 30) <= now.time() <= dt.time(14, 0)):
        st.info("目前不在台股一般交易附近時段；盤中雷達仍可手動讀取 MIS 的最新可用行情，但正式判讀請以收盤資料為準。")

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
            z["月線乖離率%"] = (pd.to_numeric(z["盤中價"], errors="coerce") / pd.to_numeric(z["盤中MA20"], errors="coerce") - 1) * 100
            z["盤中鴨嘴價格結構"] = (pd.to_numeric(z["盤中價"], errors="coerce") > pd.to_numeric(z["盤中MA20"], errors="coerce")) & (pd.to_numeric(z["盤中MA20"], errors="coerce") > pd.to_numeric(z["盤中MA60"], errors="coerce"))
            z = z.sort_values(["盤中鴨嘴價格結構", "月線乖離率%"], ascending=[False, False])
        st.caption("這裡即時計算價格、MA20/MA60 與乖離；正式『新進／退出』仍以鴨嘴完整條件在盤後確認。")
        show = [c for c in ["代號", "名稱", "市場", "盤中價", "漲跌幅%", "盤中MA20", "盤中MA60", "月線乖離率%", "盤中鴨嘴價格結構", "盤中RS"] if c in z.columns]
        st.dataframe(z[show].head(250), hide_index=True, use_container_width=True, height=600)


def render_intraday_panel(daily: pd.DataFrame, strong: pd.DataFrame, all_ok: pd.DataFrame, pre: pd.DataFrame, rs_latest: pd.Series):
    st.subheader("📡 盤中即時市場雷達")
    st.caption("盤中＝試算；收盤＝正式。這一頁不寫回 rs_latest.xlsx / duck_latest.xlsx。")
    if not BASELINE.exists():
        st.warning("尚未找到 intraday_baseline.pkl.gz。部署這個版本後，請手動跑一次『每日台股自動更新』；新的 workflow 會自動建立盤中基準檔。")
        return

    auto = st.toggle("盤中自動更新", value=True, key="intraday_auto_refresh")
    seconds = st.select_slider("刷新頻率", options=[30, 45, 60, 90, 120], value=60, format_func=lambda x: f"{x} 秒")
    if auto and hasattr(st, "fragment"):
        @st.fragment(run_every=f"{seconds}s")
        def _frag():
            _render_once(daily, strong, all_ok, pre, rs_latest)
        _frag()
    else:
        if not hasattr(st, "fragment"):
            st.info("目前 Streamlit 版本不支援局部自動刷新；可用下方按鈕手動更新，或升級 Streamlit 後自動啟用。")
        if st.button("立即刷新盤中行情", type="primary", use_container_width=True):
            fetch_mis_quotes.clear()
        _render_once(daily, strong, all_ok, pre, rs_latest)
