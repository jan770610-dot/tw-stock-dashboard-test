# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import configparser
import datetime as dt
import math
import os
import re
import ssl
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import requests
from requests.adapters import HTTPAdapter

APP_DIR = Path(__file__).resolve().parent
CACHE_DIR = APP_DIR / "cache"
OUT_DIR = APP_DIR / "output"
LOG_DIR = APP_DIR / "log"
CFG_PATH = APP_DIR / "settings.ini"
for p in (CACHE_DIR, OUT_DIR, LOG_DIR):
    p.mkdir(exist_ok=True)

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/138.0 Safari/537.36"


class _TeeWriter:
    def __init__(self, *targets):
        self.targets = targets
    def write(self, data):
        for t in self.targets:
            try:
                t.write(data)
                t.flush()
            except Exception:
                pass
        return len(data)
    def flush(self):
        for t in self.targets:
            try:
                t.flush()
            except Exception:
                pass


def install_console_log() -> None:
    try:
        name = "recalculate_console.txt" if "--analysis-only" in sys.argv else "run_console.txt"
        f = open(LOG_DIR / name, "w", encoding="utf-8", buffering=1)
        sys.stdout = _TeeWriter(sys.__stdout__, f)
        sys.stderr = _TeeWriter(sys.__stderr__, f)
    except Exception:
        pass


def _compat_ssl_context() -> ssl.SSLContext:
    ctx = ssl.create_default_context()
    strict = getattr(ssl, "VERIFY_X509_STRICT", None)
    if strict is not None:
        ctx.verify_flags &= ~strict
    return ctx


class _CompatSSLAdapter(HTTPAdapter):
    def init_poolmanager(self, connections, maxsize, block=False, **pool_kwargs):
        pool_kwargs["ssl_context"] = _compat_ssl_context()
        return super().init_poolmanager(connections, maxsize, block=block, **pool_kwargs)

    def proxy_manager_for(self, proxy, **proxy_kwargs):
        proxy_kwargs["ssl_context"] = _compat_ssl_context()
        return super().proxy_manager_for(proxy, **proxy_kwargs)


SESSION = requests.Session()
SESSION.mount("https://", _CompatSSLAdapter())
SESSION.headers.update({"User-Agent": UA, "Accept": "application/json,text/plain,*/*", "Connection": "close"})


def cfg_get() -> configparser.ConfigParser:
    c = configparser.ConfigParser()
    c.read(CFG_PATH, encoding="utf-8-sig")
    return c


def clean_num(v: Any) -> Optional[float]:
    if v is None:
        return None
    s = str(v).strip().replace(",", "").replace("--", "").replace("---", "")
    if s in {"", "-", "N/A", "nan", "None"}:
        return None
    s = s.replace("X", "").replace("*", "").replace("#", "").replace("=", "")
    try:
        return float(s)
    except Exception:
        m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
        return float(m.group()) if m else None


def is_common_stock(code: str, name: str) -> bool:
    code, name = str(code).strip(), str(name).strip()
    if not re.fullmatch(r"\d{4}", code):
        return False
    if int(code) < 1000:
        return False
    if "特別" in name or re.search(r"[甲乙丙丁]特$", name):
        return False
    upper_name = name.upper()
    if "-DR" in upper_name or "TDR" in upper_name or "存託憑證" in name:
        return False
    return True


def field_index(fields: Iterable[str], candidates: Iterable[str]) -> Optional[int]:
    fs = [str(x).replace(" ", "") for x in fields]
    for cand in candidates:
        c = cand.replace(" ", "")
        for i, f in enumerate(fs):
            if c == f or c in f:
                return i
    return None


def normalize_table(fields: List[str], rows: List[list], market: str) -> List[dict]:
    idx_code = field_index(fields, ["證券代號", "股票代號", "代號", "Code"])
    idx_name = field_index(fields, ["證券名稱", "股票名稱", "名稱", "Name"])
    idx_close = field_index(fields, ["收盤價", "收盤", "Close", "ClosingPrice"])
    idx_open = field_index(fields, ["開盤價", "開盤", "Open", "OpeningPrice"])
    idx_high = field_index(fields, ["最高價", "最高", "High", "HighestPrice"])
    idx_low = field_index(fields, ["最低價", "最低", "Low", "LowestPrice"])
    idx_vol = field_index(fields, ["成交股數", "成交量", "TradeVolume", "TradingShares"])
    if idx_code is None or idx_name is None or idx_close is None:
        return []
    out = []
    for row in rows:
        try:
            code, name = str(row[idx_code]).strip(), str(row[idx_name]).strip()
        except Exception:
            continue
        if not is_common_stock(code, name):
            continue
        close = clean_num(row[idx_close])
        if close is None or close <= 0:
            continue
        high = clean_num(row[idx_high]) if idx_high is not None else close
        low = clean_num(row[idx_low]) if idx_low is not None else close
        out.append({
            "code": code, "name": name, "market": market,
            "open": clean_num(row[idx_open]) if idx_open is not None else None,
            "high": high if high and high > 0 else close,
            "low": low if low and low > 0 else close,
            "close": close,
            "volume": clean_num(row[idx_vol]) if idx_vol is not None else None,
        })
    return out


def request_json(url: str, params: Optional[dict] = None, timeout: int = 12, tries: int = 3) -> Any:
    last = None
    for i in range(tries):
        try:
            r = SESSION.get(url, params=params, timeout=timeout)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            last = e
            if i < tries - 1:
                time.sleep(1.0 + i * 1.5)
    raise RuntimeError(f"{url} | {last}")


def _parse_twse(js: Any) -> List[dict]:
    if isinstance(js, dict) and str(js.get("stat", "")).upper() not in {"", "OK"}:
        return []
    cand = []
    if isinstance(js, dict):
        for t in js.get("tables", []) or []:
            f = t.get("fields") or []
            d = t.get("data") or []
            if f and d:
                cand.append((f, d))
        for k, v in list(js.items()):
            if k.startswith("fields") and isinstance(v, list):
                d = js.get("data" + k[6:])
                if isinstance(d, list):
                    cand.append((v, d))
    best = []
    for f, d in cand:
        p = normalize_table(f, d, "TWSE")
        if len(p) > len(best):
            best = p
    return best


def fetch_twse(day: dt.date) -> List[dict]:
    params = {"response": "json", "date": day.strftime("%Y%m%d"), "type": "ALLBUT0999"}
    errs = []
    for u in [
        "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX",
        "https://www.twse.com.tw/exchangeReport/MI_INDEX",
    ]:
        try:
            js = request_json(u, params, timeout=15, tries=2)
            rows = _parse_twse(js)
            if rows:
                return rows
            if isinstance(js, dict) and str(js.get("stat", "")).upper() not in {"", "OK"}:
                return []
        except Exception as e:
            errs.append(str(e))
    if errs:
        raise RuntimeError(" | ".join(errs))
    return []


def roc_date(day: dt.date) -> str:
    return f"{day.year - 1911}/{day.month:02d}/{day.day:02d}"


def fetch_tpex(day: dt.date) -> List[dict]:
    eps = [
        ("https://www.tpex.org.tw/www/zh-tw/afterTrading/dailyQuotes", {"date": day.strftime("%Y/%m/%d"), "id": "", "response": "json"}),
        ("https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php", {"l": "zh-tw", "o": "json", "d": roc_date(day), "s": "0,asc,0"}),
    ]
    errs = []
    for u, p in eps:
        try:
            js = request_json(u, p, timeout=15, tries=2)
            cand = []
            if isinstance(js, dict):
                for t in js.get("tables", []) or []:
                    f = t.get("fields") or t.get("columns") or []
                    d = t.get("data") or []
                    if f and d:
                        cand.append((f, d))
                if cand:
                    best = []
                    for f, d in cand:
                        z = normalize_table(f, d, "TPEx")
                        if len(z) > len(best):
                            best = z
                    if best:
                        return best
                aa = js.get("aaData") or js.get("data")
                if isinstance(aa, list) and aa and isinstance(aa[0], list):
                    out = []
                    for row in aa:
                        if len(row) < 9:
                            continue
                        code, name = str(row[0]).strip(), str(row[1]).strip()
                        if not is_common_stock(code, name):
                            continue
                        close = clean_num(row[2])
                        if close is None or close <= 0:
                            continue
                        high = clean_num(row[5]) or close
                        low = clean_num(row[6]) or close
                        out.append({
                            "code": code, "name": name, "market": "TPEx", "close": close,
                            "open": clean_num(row[4]), "high": high, "low": low,
                            "volume": clean_num(row[8]),
                        })
                    if out:
                        return out
        except Exception as e:
            errs.append(str(e))
    if errs:
        raise RuntimeError(" | ".join(errs))
    return []


def cache_path(day: dt.date) -> Path:
    return CACHE_DIR / f"{day.strftime('%Y%m%d')}.csv"


def load_cached(day: dt.date) -> Optional[pd.DataFrame]:
    p = cache_path(day)
    if not p.exists():
        return None
    try:
        df = pd.read_csv(p, dtype={"code": str})
        if {"TWSE", "TPEx"}.issubset(set(df["market"].dropna().astype(str).unique())):
            return df
    except Exception:
        pass
    return None


def save_cache(day: dt.date, rows: List[dict]) -> None:
    pd.DataFrame(rows).to_csv(cache_path(day), index=False, encoding="utf-8-sig")


def fetch_day(day: dt.date, analysis_only: bool = False) -> Tuple[dt.date, Optional[pd.DataFrame], str]:
    if day.weekday() >= 5:
        return day, None, "weekend"
    c = load_cached(day)
    if c is not None:
        return day, c, "cache"
    if analysis_only:
        return day, None, "missing-cache"
    try:
        a = fetch_twse(day)
        b = fetch_tpex(day)
        if not a and not b:
            return day, None, "holiday"
        if not a or not b:
            return day, None, f"partial TWSE={len(a)} TPEx={len(b)}"
        rows = a + b
        save_cache(day, rows)
        return day, pd.DataFrame(rows), "download"
    except Exception as e:
        return day, None, f"error {e}"


def latest_target() -> dt.date:
    now = dt.datetime.now()
    d = now.date() if now.time() >= dt.time(17, 45) else now.date() - dt.timedelta(days=1)
    while d.weekday() >= 5:
        d -= dt.timedelta(days=1)
    return d


def trend_label(ma5: float, ma20: float, ma60: float, ma5_prev5: float, ma20_prev5: float) -> str:
    vals = [ma5, ma20, ma60, ma5_prev5, ma20_prev5]
    if any(pd.isna(x) for x in vals):
        return "資料累積中"
    tol = max(0.05, abs(ma20) * 0.03)  # 百分點容許帶；避免極小差異被誤判
    if ma5 > ma20 > ma60 and ma5 > ma5_prev5 and ma20 >= ma20_prev5:
        return "明顯增加"
    if ma5 < ma20 < ma60 and ma5 < ma5_prev5 and ma20 <= ma20_prev5:
        return "明顯減少"
    if ma5 > ma20 + tol:
        return "增加"
    if ma5 < ma20 - tol:
        return "減少"
    return "持平"


def breadth_speed_label(spread: float, spread_change5: float) -> str:
    if pd.isna(spread) or pd.isna(spread_change5):
        return "資料累積中"
    # spread = 5日平均占比 - 20日平均占比；change5 = 該差值五個交易日的變化。
    # 以 0.20 個百分點做緩衝，避免微小波動造成標籤頻繁切換。
    eps = 0.20
    if spread > eps:
        if spread_change5 > eps:
            return "增加加速"
        if spread_change5 < -eps:
            return "增加放緩"
        return "增加平穩"
    if spread < -eps:
        if spread_change5 < -eps:
            return "減少加速"
        if spread_change5 > eps:
            return "減少放緩"
        return "減少平穩"
    return "變化不大"


def water_label(full_pct_rank: float) -> str:
    if pd.isna(full_pct_rank):
        return "資料不足"
    if full_pct_rank <= 10:
        return "極低水位"
    if full_pct_rank <= 25:
        return "偏低水位"
    if full_pct_rank < 75:
        return "正常水位"
    if full_pct_rank < 90:
        return "偏高水位"
    return "極高水位"


def extreme_signal(full_pct_rank: float) -> str:
    if pd.isna(full_pct_rank):
        return ""
    if full_pct_rank <= 5:
        return "歷史冰點區"
    if full_pct_rank <= 10:
        return "低檔極值區"
    if full_pct_rank >= 95:
        return "歷史高熱區"
    if full_pct_rank >= 90:
        return "高檔極值區"
    return ""


def market_state(water: str, trend: str) -> str:
    low = water in {"極低水位", "偏低水位"}
    high = water in {"極高水位", "偏高水位"}
    rising = trend in {"明顯增加", "增加"}
    falling = trend in {"明顯減少", "減少"}
    flat = trend == "持平"
    if trend == "資料累積中":
        return "資料累積中"
    if low:
        if rising:
            return "低檔回升"
        if falling:
            return "低檔續弱"
        return "低檔整理"
    if high:
        if rising:
            return "高檔續強"
        if falling:
            return "高檔轉弱"
        return "高檔整理"
    if rising:
        return "中段回升"
    if falling:
        return "中段走弱"
    if flat:
        return "中段整理"
    return "中段整理"


def rolling_percentile(s: pd.Series, window: int = 756, min_periods: int = 252) -> pd.Series:
    vals = s.to_numpy(dtype=float)
    out = np.full(len(vals), np.nan, dtype=float)
    for i, v in enumerate(vals):
        if np.isnan(v):
            continue
        left = max(0, i - window + 1)
        a = vals[left:i + 1]
        a = a[~np.isnan(a)]
        if len(a) < min_periods:
            continue
        out[i] = float((a <= v).sum() / len(a) * 100.0)
    return pd.Series(out, index=s.index)


def _downcast_market_data(all_df: pd.DataFrame) -> pd.DataFrame:
    keep = [c for c in ["date", "code", "name", "market", "close", "high", "low", "volume"] if c in all_df.columns]
    x = all_df[keep].copy()
    x["code"] = x["code"].astype(str)
    x["market"] = x["market"].astype("category")
    x["name"] = x["name"].astype("category")
    for c in ["close", "high", "low", "volume"]:
        x[c] = pd.to_numeric(x[c], errors="coerce", downcast="float")
    return x


def compute_counts(all_df: pd.DataFrame, start: dt.date, end: dt.date, cfg: configparser.ConfigParser):
    g = cfg["strategy"]
    rs_cut = float(g.get("rs_threshold", 85))
    amount_cut = float(g.get("amount_min", 30000000))
    high_dist = float(g.get("max_below_52w_high_pct", 25)) / 100.0
    low_above = float(g.get("min_above_52w_low_pct", 30)) / 100.0
    rs_days = int(g.get("rs_days", 250))
    ma50 = int(g.get("ma50_days", 50))
    ma200 = int(g.get("ma200_days", 200))
    hl_days = int(g.get("high_low_days", 250))

    all_df = _downcast_market_data(all_df)
    all_df["date"] = pd.to_datetime(all_df["date"])
    all_df = all_df.sort_values(["code", "date"])
    all_df["amount"] = (all_df["close"].astype("float64") * all_df["volume"].fillna(0).astype("float64"))
    grp = all_df.groupby("code", group_keys=False, observed=True)
    all_df["ma50"] = grp["close"].transform(lambda s: s.rolling(ma50, min_periods=ma50).mean()).astype("float32")
    all_df["ma200"] = grp["close"].transform(lambda s: s.rolling(ma200, min_periods=ma200).mean()).astype("float32")
    all_df["high250"] = grp["high"].transform(lambda s: s.rolling(hl_days, min_periods=hl_days).max()).astype("float32")
    all_df["low250"] = grp["low"].transform(lambda s: s.rolling(hl_days, min_periods=hl_days).min()).astype("float32")
    all_df["ret250"] = grp["close"].transform(lambda s: s / s.shift(rs_days) - 1).astype("float32")
    all_df["rs"] = (all_df.groupby("date", observed=True)["ret250"].rank(method="average", pct=True) * 100).astype("float32")
    all_df["eligible"] = all_df[["ma200", "high250", "low250", "ret250"]].notna().all(axis=1)
    all_df["rs85"] = all_df["eligible"] & (all_df["rs"] > rs_cut)
    all_df["strong"] = (
        all_df["rs85"]
        & (all_df["close"] > all_df["ma200"])
        & (all_df["ma50"] > all_df["ma200"])
        & (all_df["amount"] > amount_cut)
        & (all_df["close"] >= all_df["high250"] * (1 - high_dist))
        & (all_df["close"] >= all_df["low250"] * (1 + low_above))
    )

    # 先對所有已抓歷史日計算市場廣度，再切到 2018 起輸出；這樣 2018/1 月就有 60 日趨勢基準。
    d_all = all_df.groupby("date", observed=True).agg(
        eligible_count=("eligible", "sum"),
        rs85_count=("rs85", "sum"),
        strong_count=("strong", "sum"),
    ).reset_index().sort_values("date")
    d_all = d_all[d_all["eligible_count"] > 0].copy()
    d_all["strong_pct"] = np.where(d_all["eligible_count"] > 0, d_all["strong_count"] / d_all["eligible_count"] * 100, np.nan)
    d_all["change"] = d_all["strong_count"].diff()
    d_all["ma5"] = d_all["strong_count"].rolling(5, min_periods=5).mean()
    d_all["ma20"] = d_all["strong_count"].rolling(20, min_periods=20).mean()
    d_all["ma60"] = d_all["strong_count"].rolling(60, min_periods=60).mean()
    d_all["pct_ma5"] = d_all["strong_pct"].rolling(5, min_periods=5).mean()
    d_all["pct_ma20"] = d_all["strong_pct"].rolling(20, min_periods=20).mean()
    d_all["pct_ma60"] = d_all["strong_pct"].rolling(60, min_periods=60).mean()
    d_all["pct_ma5_prev5"] = d_all["pct_ma5"].shift(5)
    d_all["pct_ma20_prev5"] = d_all["pct_ma20"].shift(5)
    d_all["breadth_spread"] = d_all["pct_ma5"] - d_all["pct_ma20"]
    d_all["breadth_spread_change5"] = d_all["breadth_spread"] - d_all["breadth_spread"].shift(5)
    d_all["breadth_speed"] = [
        breadth_speed_label(r.breadth_spread, r.breadth_spread_change5)
        for r in d_all.itertuples(index=False)
    ]
    rolling_days = int(cfg.get("trend", "rolling_percentile_days", fallback="756"))
    rolling_min = int(cfg.get("trend", "rolling_percentile_min_periods", fallback="252"))
    d_all["rolling3y_pct_rank"] = rolling_percentile(d_all["strong_pct"], rolling_days, rolling_min)
    d_all["trend"] = [
        trend_label(r.pct_ma5, r.pct_ma20, r.pct_ma60, r.pct_ma5_prev5, r.pct_ma20_prev5)
        for r in d_all.itertuples(index=False)
    ]

    mask = (d_all["date"].dt.date >= start) & (d_all["date"].dt.date <= end)
    d = d_all.loc[mask].copy().reset_index(drop=True)
    if d.empty:
        return d, pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # 全期回顧百分位：用 2018~目前全部歷史判斷目前位於長期分布的哪裡。
    d["full_pct_rank"] = d["strong_pct"].rank(method="average", pct=True) * 100
    # 近3年滾動百分位：只使用當日以前的可得資料，較適合觀察近期市場環境。
    # 已在 d_all 上先計算，因此 2018 初期可利用暖機資料，不會從第一天就出現 P100 假象。
    d["water"] = d["full_pct_rank"].map(water_label)
    d["extreme"] = d["full_pct_rank"].map(extreme_signal)
    d["market_state"] = [market_state(w, t) for w, t in zip(d["water"], d["trend"])]
    d["month"] = d["date"].dt.to_period("M").astype(str)

    # 極值的「主基準」使用占比；家數則依最新有效樣本動態換算，避免不同年份總股票數不同造成誤判。
    ps = [5, 10, 25, 50, 75, 90, 95]
    pct_vals = np.nanpercentile(d["strong_pct"], ps)
    latest_eligible = int(d["eligible_count"].iloc[-1])
    equiv_counts = np.rint(pct_vals / 100.0 * latest_eligible).astype(int)
    extreme_df = pd.DataFrame({
        "percentile": [f"P{p}" for p in ps],
        "strong_pct": pct_vals,
        "equiv_count": equiv_counts,
        "meaning": [
            "歷史冰點參考", "低檔極值線", "偏低水位線", "歷史中位數", "偏高水位線", "高檔極值線", "歷史高熱參考"
        ],
    })

    m = d.groupby("month", observed=True).agg(
        avg_strong=("strong_count", "mean"),
        high=("strong_count", "max"),
        low=("strong_count", "min"),
        end_count=("strong_count", "last"),
        avg_pct=("strong_pct", "mean"),
        end_pct=("strong_pct", "last"),
        end_percentile=("full_pct_rank", "last"),
        end_rolling3y=("rolling3y_pct_rank", "last"),
        end_water=("water", "last"),
        end_trend=("trend", "last"),
        end_state=("market_state", "last"),
        end_spread=("breadth_spread", "last"),
        end_speed=("breadth_speed", "last"),
    ).reset_index()
    m["end_change_vs_prev_month"] = m["end_count"].diff()

    # 若 end 是假日/尚未有完整行情，最新名單改抓輸出資料中的最後一個有效交易日。
    latest_day = d["date"].max()
    latest_list = all_df[(all_df["date"] == latest_day) & all_df["strong"]].copy()
    latest_list = latest_list[["code", "name", "market", "close", "rs", "ma50", "ma200", "high250", "low250", "amount"]].sort_values("rs", ascending=False)
    return d, m, latest_list, extreme_df


def build_dashboard_text(latest_row, extreme_df: pd.DataFrame) -> str:
    p10 = extreme_df.loc[extreme_df["percentile"] == "P10", "strong_pct"].iloc[0]
    p90 = extreme_df.loc[extreme_df["percentile"] == "P90", "strong_pct"].iloc[0]
    state = latest_row["market_state"]
    messages = {
        "低檔回升": "強勢股仍在歷史低檔區，但正在增加，屬低檔回升；可觀察是否延續成更完整的底部修復。",
        "低檔整理": "強勢股仍少，但短期變化不大；目前偏低檔整理，尚未出現明確回升或續弱。",
        "低檔續弱": "強勢股已經偏少，而且仍在減少；目前屬低檔續弱，不能只因水位低就視為見底。",
        "中段回升": "市場廣度位於中間區，但強勢股正在增加；屬中段回升，內部結構改善。",
        "中段整理": "市場廣度位於中間區，強勢股數量變化不大；目前屬中段整理。",
        "中段走弱": "市場廣度位於中間區，但強勢股正在減少；屬中段走弱，需留意後續是否跌入低水位。",
        "高檔續強": "強勢股處在歷史高水位，而且仍在增加；屬高檔續強，不應只因水位高就直接判定見頂。",
        "高檔整理": "強勢股處在歷史高水位，但短期變化不大；屬高檔整理，重點看後續往上擴散或轉弱。",
        "高檔轉弱": "強勢股仍在高水位，但已開始減少；屬高檔轉弱，這是比單純『很熱』更值得注意的風險訊號。",
    }
    core = messages.get(state, "目前市場廣度仍在資料累積或中性區。")
    r3 = latest_row.get("rolling3y_pct_rank", np.nan)
    r3txt = f"近3年 P{r3:.1f}" if not pd.isna(r3) else "近3年資料不足"
    spread = latest_row.get("breadth_spread", np.nan)
    speed = latest_row.get("breadth_speed", "")
    spread_txt = f"5日－20日廣度差 {spread:+.2f} 個百分點（{speed}）" if not pd.isna(spread) else "廣度差資料不足"
    return (f"{core} 目前強勢股占比 {latest_row['strong_pct']:.2f}%；全期回顧 P{latest_row['full_pct_rank']:.1f}，{r3txt}；"
            f"{spread_txt}。歷史 P10={p10:.2f}%，P90={p90:.2f}%。")


def write_excel(d: pd.DataFrame, m: pd.DataFrame, latest: pd.DataFrame, extreme_df: pd.DataFrame, cfg, start, end) -> Path:
    out = OUT_DIR / f"{start.strftime('%Y%m%d')}_{end.strftime('%Y%m%d')}_RS強勢股市場廣度_極值趨勢.xlsx"
    with pd.ExcelWriter(out, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as w:
        book = w.book
        fmt_title = book.add_format({"bold": True, "font_size": 16, "font_color": "white", "bg_color": "#17365D", "align": "center", "valign": "vcenter"})
        fmt_head = book.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
        fmt_kpi = book.add_format({"bold": True, "font_size": 13, "bg_color": "#D9EAF7", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_kpi_value = book.add_format({"bold": True, "font_size": 15, "border": 1, "align": "center", "valign": "vcenter"})
        fmt_note = book.add_format({"text_wrap": True, "valign": "top", "bg_color": "#FFF2CC", "border": 1})
        fmt_pct = book.add_format({"num_format": "0.00"})
        fmt_dec = book.add_format({"num_format": "0.00"})
        fmt_int = book.add_format({"num_format": "0"})

        # 每日表
        d2 = d[[
            "date", "eligible_count", "rs85_count", "strong_count", "change", "strong_pct",
            "ma5", "ma20", "ma60", "pct_ma5", "pct_ma20", "pct_ma60",
            "full_pct_rank", "rolling3y_pct_rank", "water", "extreme", "trend", "market_state",
            "breadth_spread", "breadth_spread_change5", "breadth_speed", "month"
        ]].copy()
        d2.columns = [
            "日期", "有效樣本數", "RS強勢母池檔數", "強勢股檔數", "較前日增減", "強勢股占比%",
            "5日均量(檔)", "20日均量(檔)", "60日均量(檔)", "5日平均占比%", "20日平均占比%", "60日平均占比%",
            "全期回顧百分位%", "近3年滾動百分位%", "歷史水位", "極值訊號", "強勢股方向", "市場階段",
            "5日-20日廣度差", "廣度差5日變化", "變化速度", "月份"
        ]
        d2.to_excel(w, index=False, sheet_name="每日強勢股數量")
        sh = w.sheets["每日強勢股數量"]
        sh.freeze_panes(1, 0)
        sh.autofilter(0, 0, len(d2), len(d2.columns) - 1)
        sh.set_row(0, 34, fmt_head)
        sh.set_column("A:A", 12)
        sh.set_column("B:E", 13)
        sh.set_column("F:N", 15)
        sh.set_column("O:R", 15)
        sh.set_column("S:U", 16)
        sh.set_column("V:V", 10)
        if len(d2) > 0:
            sh.conditional_format(1, 3, len(d2), 3, {"type": "data_bar", "bar_color": "#4F81BD"})
            sh.conditional_format(1, 12, len(d2), 12, {"type": "3_color_scale", "min_color": "#5B9BD5", "mid_color": "#FFF2CC", "max_color": "#C00000"})
            sh.conditional_format(1, 13, len(d2), 13, {"type": "3_color_scale", "min_color": "#5B9BD5", "mid_color": "#FFF2CC", "max_color": "#C00000"})
            sh.conditional_format(1, 18, len(d2), 19, {"type": "3_color_scale", "min_color": "#C00000", "mid_color": "#FFF2CC", "max_color": "#70AD47"})

        # 圖1：家數 + 5/20/60均量
        n = len(d2)
        chart = book.add_chart({"type": "line"})
        col_idx = {name: i for i, name in enumerate(d2.columns)}
        for name, color, width in [
            ("強勢股檔數", "#1F4E78", 1.5), ("5日均量(檔)", "#70AD47", 1.25),
            ("20日均量(檔)", "#C55A11", 1.25), ("60日均量(檔)", "#7030A0", 1.25),
        ]:
            c = col_idx[name]
            chart.add_series({"name": name, "categories": ["每日強勢股數量", 1, 0, n, 0], "values": ["每日強勢股數量", 1, c, n, c], "line": {"color": color, "width": width}})
        chart.set_title({"name": "強勢股家數與 5/20/60 日趨勢"})
        chart.set_x_axis({"name": "日期", "date_axis": True})
        chart.set_y_axis({"name": "檔數", "major_gridlines": {"visible": True}})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 900, "height": 420})
        sh.insert_chart("X2", chart)

        # 圖2：強勢股占比 + 5/20/60日平均占比
        chart2 = book.add_chart({"type": "line"})
        for name, color, width in [
            ("強勢股占比%", "#1F4E78", 1.4), ("5日平均占比%", "#70AD47", 1.2),
            ("20日平均占比%", "#C55A11", 1.2), ("60日平均占比%", "#7030A0", 1.2),
        ]:
            c = col_idx[name]
            chart2.add_series({"name": name, "categories": ["每日強勢股數量", 1, 0, n, 0], "values": ["每日強勢股數量", 1, c, n, c], "line": {"color": color, "width": width}})
        chart2.set_title({"name": "強勢股占比（跨年度比較的主要指標）"})
        chart2.set_x_axis({"name": "日期", "date_axis": True})
        chart2.set_y_axis({"name": "占有效樣本 %", "major_gridlines": {"visible": True}})
        chart2.set_legend({"position": "bottom"})
        chart2.set_size({"width": 900, "height": 420})
        sh.insert_chart("X24", chart2)

        # 極值基準
        e2 = extreme_df[["percentile", "strong_pct", "equiv_count", "meaning"]].copy()
        e2.columns = ["百分位", "強勢股占比%", "依最新有效樣本換算(檔)", "意義"]
        e2.to_excel(w, index=False, sheet_name="極值基準")
        se = w.sheets["極值基準"]
        se.set_row(0, 26, fmt_head)
        se.set_column("A:A", 12)
        se.set_column("B:C", 22)
        se.set_column("D:D", 24)
        se.write(len(e2) + 2, 0, "判讀原則", fmt_head)
        se.merge_range(len(e2) + 3, 0, len(e2) + 5, 3,
                       "跨年度判斷『到頂／到底』以強勢股占比的歷史百分位為主。家數欄改成依『最新有效樣本數』動態換算，因此不再使用跨年度固定家數。P10 以下屬低檔極值區，P90 以上屬高檔極值區；高水位不等於立即到頂、低水位也不等於立即到底，還要看強勢股方向與市場階段。",
                       fmt_note)

        # 月度摘要
        m2 = m[["month", "avg_strong", "high", "low", "end_count", "end_change_vs_prev_month", "avg_pct", "end_pct", "end_percentile", "end_rolling3y", "end_water", "end_trend", "end_state", "end_spread", "end_speed"]].copy()
        m2.columns = ["月份", "平均強勢股數", "月高", "月低", "月底強勢股數", "月底較前月", "平均強勢股占比%", "月底強勢股占比%", "月底全期百分位%", "月底近3年百分位%", "月底水位", "月底強勢股方向", "月底市場階段", "月底5日-20日廣度差", "月底變化速度"]
        m2.to_excel(w, index=False, sheet_name="月度摘要")
        sm = w.sheets["月度摘要"]
        sm.freeze_panes(1, 0)
        sm.set_row(0, 30, fmt_head)
        sm.set_column("A:A", 10)
        sm.set_column("B:I", 16)
        sm.set_column("J:O", 17)

        # 最新強勢股名單
        latest2 = latest.copy()
        latest2.columns = ["代號", "名稱", "市場", "收盤", "RS", "MA50", "MA200", "52週高點", "52週低點", "成交金額"]
        latest2.to_excel(w, index=False, sheet_name="最新強勢股名單")
        sl = w.sheets["最新強勢股名單"]
        sl.freeze_panes(1, 0)
        sl.set_row(0, 26, fmt_head)
        sl.set_column("A:C", 14)
        sl.set_column("D:J", 14)

        # 設定與說明
        setsh = book.add_worksheet("設定與說明")
        setsh.merge_range("A1:D1", "RS 強勢股廣度—極值、方向與市場階段說明", fmt_title)
        rows = [
            ["項目", "本版設定", "簡單解釋", "判讀重點"],
            ["統計期間", f"{start} ~ {end}", "用 2018 至今建立極值分布", "至少涵蓋多空循環；日後每日自動延長"],
            ["極值主指標", "強勢股占有效樣本比例", "避免上市櫃總家數逐年增加造成固定家數失真", "家數照樣顯示，但水位以占比為主"],
            ["極低水位", "≤ 歷史 P10", "歷史最弱約 10% 的日子", "不是直接抄底訊號"],
            ["偏低水位", "P10 ~ P25", "市場強勢股偏少", "觀察是否開始擴散"],
            ["正常水位", "P25 ~ P75", "市場一般區間", "重點看方向"],
            ["偏高水位", "P75 ~ P90", "多頭廣度偏強", "不代表一定見頂"],
            ["極高水位", "≥ 歷史 P90", "歷史最熱約 10% 的日子", "若同時收縮才是高檔退潮警訊"],
            ["明顯增加", "5日占比 > 20日 > 60日，且5/20日都上升", "強勢股明顯持續增加", "最強的增加型態"],
            ["增加", "5日平均占比明顯高於20日", "近期強勢股比中期多", "市場廣度改善"],
            ["持平", "5日與20日平均占比接近", "強勢股數量變化不大", "方向不明顯"],
            ["減少", "5日平均占比明顯低於20日", "近期強勢股比中期少", "市場廣度轉弱"],
            ["明顯減少", "5日 < 20日 < 60日，且5/20日都下降", "強勢股明顯持續減少", "最弱的減少型態"],
            ["低檔回升", "低水位 + 增加/明顯增加", "強勢股仍少，但開始回升", "觀察是否形成底部修復"],
            ["低檔整理", "低水位 + 持平", "強勢股仍少，短期變化不大", "等待方向"],
            ["低檔續弱", "低水位 + 減少/明顯減少", "強勢股很少而且還在減少", "尚未止穩"],
            ["中段回升", "正常水位 + 增加", "市場廣度由中段往上改善", "偏多"],
            ["中段整理", "正常水位 + 持平", "市場廣度中性", "無明顯方向"],
            ["中段走弱", "正常水位 + 減少", "市場廣度由中段往下惡化", "偏空"],
            ["高檔續強", "高水位 + 增加/明顯增加", "市場很熱但強勢股仍在增加", "不等於立即見頂"],
            ["高檔整理", "高水位 + 持平", "市場仍熱但短期沒有明顯方向", "等待變化"],
            ["高檔轉弱", "高水位 + 減少/明顯減少", "市場仍熱但強勢股開始退潮", "高檔風險訊號"],
            ["5日-20日廣度差", "5日平均占比－20日平均占比", "正值代表近期比中期強；負值代表近期比中期弱", "看方向強弱"],
            ["變化速度", "廣度差相較5日前的變化", "增加加速/放緩、減少加速/放緩", "看強弱變化是否加速"],
            ["近3年滾動百分位", "最近756交易日；至少252日才計算", "反映近期市場環境中的冷熱位置", "搭配全期百分位一起看"],
            ["年RS門檻", cfg["strategy"].get("rs_threshold", "85"), "250交易日報酬的全市場百分位", "尼克萊公開主策略近似"],
            ["成交金額門檻", cfg["strategy"].get("amount_min", "30000000"), "日成交金額", "公開：>3000萬"],
            ["52週高低點", f"高點內 {cfg['strategy'].get('max_below_52w_high_pct','25')}%；低點上 {cfg['strategy'].get('min_above_52w_low_pct','30')}%", "公開未揭露精確百分比，因此採可調近似", "可修改 settings.ini 後重算"],
            ["TWSE來源", "https://www.twse.com.tw/", "上市每日行情", "官方"],
            ["TPEx來源", "https://www.tpex.org.tw/", "上櫃每日行情", "官方"],
            ["重要限制", "公開策略近似版", "尼克萊/CMoney 私有 RS 細節與高低點精確門檻未完整公開", "不宣稱與付費 APP 每日名單完全一致"],
        ]
        for r, row in enumerate(rows, 2):
            for c, val in enumerate(row):
                setsh.write(r - 1, c, val, fmt_head if r == 2 else None)
        setsh.set_column("A:A", 20)
        setsh.set_column("B:B", 38)
        setsh.set_column("C:D", 48)
        setsh.freeze_panes(2, 0)

        # 首頁 Dashboard 最後建立，並移到最前
        dash = book.add_worksheet("市場廣度總覽")
        dash.merge_range("A1:J1", "RS 強勢股市場廣度｜極值＋方向＋市場階段", fmt_title)
        if not d.empty:
            r = d.iloc[-1]
            p10 = extreme_df.loc[extreme_df["percentile"] == "P10"].iloc[0]
            p90 = extreme_df.loc[extreme_df["percentile"] == "P90"].iloc[0]
            p5 = extreme_df.loc[extreme_df["percentile"] == "P5"].iloc[0]
            p95 = extreme_df.loc[extreme_df["percentile"] == "P95"].iloc[0]
            kpis = [
                ("最新交易日", r["date"].strftime("%Y-%m-%d")),
                ("強勢股檔數", int(r["strong_count"])),
                ("強勢股占比", f"{r['strong_pct']:.2f}%"),
                ("全期回顧百分位", f"P{r['full_pct_rank']:.1f}"),
                ("近3年百分位", f"P{r['rolling3y_pct_rank']:.1f}" if not pd.isna(r['rolling3y_pct_rank']) else "—"),
                ("目前水位", r["water"]),
                ("強勢股方向", r["trend"]),
                ("市場階段", r["market_state"]),
                ("5日-20日廣度差", f"{r['breadth_spread']:+.2f} 個百分點" if not pd.isna(r['breadth_spread']) else "—"),
                ("變化速度", r["breadth_speed"]),
            ]
            for i, (label, val) in enumerate(kpis):
                row = 3 + (i // 5) * 2
                col = (i % 5) * 2
                dash.write(row - 1, col, label, fmt_kpi)
                dash.write(row - 1, col + 1, val, fmt_kpi_value)
            latest_n = int(r["eligible_count"])
            dash.merge_range("A8:B8", "歷史極值參考（家數依今日有效樣本動態換算）", fmt_head)
            dash.write("A9", "P5 冰點")
            dash.write("B9", f"{p5['strong_pct']:.2f}% / 今日約 {int(round(p5['strong_pct'] / 100 * latest_n))} 檔")
            dash.write("A10", "P10 低檔線")
            dash.write("B10", f"{p10['strong_pct']:.2f}% / 今日約 {int(round(p10['strong_pct'] / 100 * latest_n))} 檔")
            dash.write("D9", "P90 高檔線")
            dash.write("E9", f"{p90['strong_pct']:.2f}% / 今日約 {int(round(p90['strong_pct'] / 100 * latest_n))} 檔")
            dash.write("D10", "P95 高熱")
            dash.write("E10", f"{p95['strong_pct']:.2f}% / 今日約 {int(round(p95['strong_pct'] / 100 * latest_n))} 檔")
            dash.merge_range("A12:J14", build_dashboard_text(r, extreme_df), fmt_note)
            dash.merge_range("A16:J18", "最直觀的看法：①『目前水位』看現在偏冷還是偏熱；②『強勢股方向』直接看家數是在增加、持平還是減少；③『市場階段』把位置和方向合在一起：低檔回升＝低檔開始改善，高檔轉弱＝高檔開始退潮。再看『5日-20日廣度差』與『變化速度』判斷改善/惡化是否正在加速。", fmt_note)
        dash.set_column("A:J", 18)
        dash.set_row(0, 30)

        # xlsxwriter 依建立順序排工作表；把 Dashboard 設為使用者開啟時看到的頁面。
        dash.set_first_sheet()
        dash.activate()
        dash.select()

    return out


def load_or_download_frames(dates: list[dt.date], workers: int, analysis_only: bool) -> tuple[list[pd.DataFrame], list[tuple[dt.date, str]]]:
    frames = []
    failed = []
    done = 0
    started = time.time()
    with ThreadPoolExecutor(max_workers=max(1, min(workers, 4))) as ex:
        futs = {ex.submit(fetch_day, d, analysis_only): d for d in dates}
        for fut in as_completed(futs):
            try:
                day, df, status = fut.result()
            except Exception as e:
                day = futs[fut]
                df = None
                status = f"unexpected {type(e).__name__}: {e}"
            done += 1
            if df is not None and not df.empty:
                df = df.copy()
                df["date"] = day.isoformat()
                frames.append(df)
            elif status not in {"weekend", "holiday"}:
                failed.append((day, status))
            if done == 1 or done % 20 == 0 or done == len(dates):
                elapsed = max(time.time() - started, 0.1)
                rate = done / elapsed
                eta_min = (len(dates) - done) / rate / 60 if rate > 0 else 0
                print(f"[{done}/{len(dates)}] {day} {status} | ETA約 {eta_min:.1f} 分鐘", flush=True)
    return frames, failed


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--analysis-only", action="store_true", help="只讀 cache 重算，不連網補缺日")
    parser.add_argument("--date", help="指定分析日期 YYYY-MM-DD；供雲端排程使用")
    args = parser.parse_args()

    cfg = cfg_get()
    start = dt.datetime.strptime(cfg["general"].get("start_date", "2018-01-01"), "%Y-%m-%d").date()
    end_s = (args.date or os.environ.get("RS_TARGET_DATE") or cfg["general"].get("end_date", "auto")).strip().lower()
    end = latest_target() if end_s in {"", "auto"} else dt.datetime.strptime(end_s, "%Y-%m-%d").date()
    back_days = int(cfg["general"].get("history_calendar_days", "550"))
    hist_start = start - dt.timedelta(days=back_days)
    dates = [hist_start + dt.timedelta(days=i) for i in range((end - hist_start).days + 1) if (hist_start + dt.timedelta(days=i)).weekday() < 5]
    workers = int(cfg["general"].get("workers", "1"))

    print("=" * 68)
    print("RS 強勢股市場廣度 v5｜2018至今・動態極值＋直觀方向版")
    print(f"統計主期間：{start} ~ {end}")
    print(f"技術指標暖機資料：{hist_start} 起（約 {back_days} 個日曆日）")
    print(f"需檢查 {len(dates)} 個平日候選；已有 cache 的日期不重抓。")
    print("第一次完整建立可能很久，之後每日只補新日期。")
    if args.analysis_only:
        print("模式：只讀 cache 重算（不連網）")
    print("=" * 68)

    frames, failed = load_or_download_frames(dates, workers, args.analysis_only)

    if failed and not args.analysis_only:
        print(f"第一次抓取有 {len(failed)} 個日期失敗，開始逐日重試...", flush=True)
        still = []
        for i, (day, _) in enumerate(sorted(failed), 1):
            day, df, status = fetch_day(day, False)
            if df is not None and not df.empty:
                df = df.copy()
                df["date"] = day.isoformat()
                frames.append(df)
            else:
                still.append((day, status))
            print(f"[重試 {i}/{len(failed)}] {day} {status}", flush=True)
        failed = still

    if failed:
        with open(LOG_DIR / "missing_dates.txt", "w", encoding="utf-8") as f:
            for day, status in failed:
                f.write(f"{day}\t{status}\n")
        print(f"警告：仍有 {len(failed)} 日沒有完整 cache，已寫入 log\\missing_dates.txt。", flush=True)
        if args.analysis_only:
            print("analysis-only 不會補下載；若缺很多日，請改跑 01_FIRST_SETUP_RS.cmd。", flush=True)

    if not frames:
        raise SystemExit("沒有取得任何行情資料。")

    print("開始合併歷史資料並計算 RS / MA / 52週高低點...", flush=True)
    data = pd.concat(frames, ignore_index=True).drop_duplicates(["date", "code"], keep="last")
    d, m, latest, extreme_df = compute_counts(data, start, end, cfg)
    if d.empty:
        raise SystemExit("統計期間內沒有足夠資料可計算。")
    actual_end = d["date"].max().date()
    out = write_excel(d, m, latest, extreme_df, cfg, start, actual_end)
    r = d.iloc[-1]
    p10 = extreme_df.loc[extreme_df["percentile"] == "P10"].iloc[0]
    p90 = extreme_df.loc[extreme_df["percentile"] == "P90"].iloc[0]

    print("=" * 68)
    print(f"完成：{out}")
    print(f"最新 {r['date'].date()}：強勢股 {int(r['strong_count'])} 檔 / {r['strong_pct']:.2f}%")
    print(f"歷史水位：{r['water']}（P{r['full_pct_rank']:.1f}）")
    print(f"強勢股方向：{r['trend']}｜市場階段：{r['market_state']}｜變化速度：{r['breadth_speed']}｜極值：{r['extreme'] or '—'}")
    print(f"近3年百分位：P{r['rolling3y_pct_rank']:.1f}" if not pd.isna(r['rolling3y_pct_rank']) else "近3年百分位：資料不足")
    print(f"5日-20日廣度差：{r['breadth_spread']:+.2f} 個百分點")
    print(f"目前歷史 P10 低檔線：{p10['strong_pct']:.2f}% / 依今日樣本約 {int(p10['equiv_count'])} 檔")
    print(f"目前歷史 P90 高檔線：{p90['strong_pct']:.2f}% / 依今日樣本約 {int(p90['equiv_count'])} 檔")
    print("=" * 68)


if __name__ == "__main__":
    install_console_log()
    try:
        main()
    except KeyboardInterrupt:
        print("\n使用者中止執行。", flush=True)
        raise SystemExit(130)
    except Exception as e:
        import traceback
        LOG_DIR.mkdir(exist_ok=True)
        msg = traceback.format_exc()
        try:
            (LOG_DIR / "last_error.txt").write_text(msg, encoding="utf-8")
        except Exception:
            pass
        print("\n[ERROR] 執行失敗：", repr(e), flush=True)
        print(msg, flush=True)
        print("錯誤已寫入 log\\last_error.txt", flush=True)
        raise SystemExit(1)
