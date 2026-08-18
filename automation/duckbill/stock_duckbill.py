# -*- coding: utf-8 -*-
"""
台股鴨嘴型態篩選系統 v30（特殊申報期限＋真實缺漏判定版）
- 上市：TWSE 官方資料
- 上櫃：TPEx 官方資料
- 台指期：TAIFEX 官方每日行情（TX，一般交易時段）
- 儲存：SQLite
- 輸出：Excel (.xlsx)
- EPS：Yahoo Finance/yfinance 季度 EPS 優先（估值候選＋鴨嘴＋預備鴨嘴），MOPS/交易所隱含 EPS 備援
- 預備鴨嘴：尚未正式成立，但 5 條件已符合 3~4 條且剩餘條件接近門檻
- v23：主工作表改為「全市場整合篩選」，技術面/預備鴨嘴/營收/三率/EPS/估值同列，可直接用篩選器交叉查看

鴨嘴條件：
1) Close > MA20 and Close > MA60
2) MA20 > MA60
3) MA20(today) > MA20(prev)
4) MA60(today) > MA60(prev)
5) (MA20-MA60)(today) > (MA20-MA60)(prev)
"""
from __future__ import annotations

import argparse
import configparser
import calendar
import datetime as dt
import json
import logging
import csv
import io
import math
import os
import re
import sqlite3
import ssl
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests
from requests.adapters import HTTPAdapter

try:
    import yfinance as yf
except Exception:
    yf = None

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "stock.db"
OUT_DIR = APP_DIR / "output"
LOG_DIR = APP_DIR / "log"
CONFIG_PATH = APP_DIR / "settings.ini"
LEGACY_CONFIG_PATH = APP_DIR / "設定.ini"

OUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "update.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("duckbill")

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/138.0 Safari/537.36"
)
def _compat_ssl_context() -> ssl.SSLContext:
    """
    Python 3.13+ enables VERIFY_X509_STRICT by default.
    Some official Taiwan government endpoints currently present a certificate chain
    that OpenSSL rejects under STRICT with errors such as
    "Missing Subject Key Identifier".  We keep normal CA + hostname validation,
    but disable only the extra STRICT compatibility flag documented by Python.
    This is NOT verify=False.
    """
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
        # 讓公司代理/HTTPS proxy 也套用同一個相容 context。
        proxy_kwargs["ssl_context"] = _compat_ssl_context()
        return super().proxy_manager_for(proxy, **proxy_kwargs)


SESSION = requests.Session()
SESSION.mount("https://", _CompatSSLAdapter())
SESSION.headers.update({"User-Agent": UA, "Accept": "application/json,text/plain,*/*", "Connection": "close"})


def load_config() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH if CONFIG_PATH.exists() else LEGACY_CONFIG_PATH, encoding="utf-8-sig")
    if "general" not in cfg:
        cfg["general"] = {}
    return cfg


def parse_date(s: Optional[str]) -> dt.date:
    if s:
        return dt.datetime.strptime(s, "%Y-%m-%d").date()
    # 17:45 前不把尚未收盤/尚未完整更新的「今天」當正式資料日。
    # 早上執行時，預設分析前一個工作日；週末則往前退到週五。
    now = dt.datetime.now()
    day = now.date() if now.time() >= dt.time(17, 45) else now.date() - dt.timedelta(days=1)
    while day.weekday() >= 5:
        day -= dt.timedelta(days=1)
    return day


def request_json(url: str, params: Optional[dict] = None, timeout: int = 12, tries: int = 2) -> Any:
    last: Optional[Exception] = None
    for i in range(tries):
        try:
            r = SESSION.get(url, params=params, timeout=timeout)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            last = e
            if i < tries - 1:
                time.sleep(1.2 * (i + 1))
    raise RuntimeError(f"GET JSON 失敗: {url} params={params} | {last}")



def request_csv_rows(url: str, timeout: int = 15, tries: int = 3) -> List[dict]:
    """下載官方 CSV，回傳 dict rows。特別用於 MOPS OpenData，避免大型 JSON 被中途 reset。"""
    last: Optional[Exception] = None
    for i in range(tries):
        try:
            # 不共用 SESSION，避免某些政府站對 keep-alive/舊連線重設。
            r = SESSION.get(
                url,
                timeout=timeout,
                headers={
                    "User-Agent": UA,
                    "Accept": "text/csv,text/plain,*/*",
                    "Connection": "close",
                    "Referer": "https://mopsfin.twse.com.tw/",
                },
            )
            r.raise_for_status()
            raw = r.content
            # MOPS OpenData 多為 UTF-8-SIG；保留 Big5 fallback。
            text = None
            for enc in ("utf-8-sig", "utf-8", "cp950", "big5"):
                try:
                    text = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    pass
            if text is None:
                text = raw.decode("utf-8", errors="replace")
            rows = list(csv.DictReader(io.StringIO(text)))
            if not rows:
                raise RuntimeError("CSV 無資料列")
            return rows
        except Exception as e:
            last = e
            if i < tries - 1:
                time.sleep(1.5 * (i + 1))
    raise RuntimeError(f"GET CSV 失敗: {url} | {last}")


def request_rows_with_fallback(primary_url: str, fallback_url: Optional[str] = None,
                               primary_kind: str = "json", timeout: int = 20) -> Tuple[List[dict], str]:
    """取得資料列；primary 失敗時自動切官方備援。回傳(rows, used_url)。"""
    errors: List[str] = []
    urls = [(primary_url, primary_kind)]
    if fallback_url:
        urls.append((fallback_url, "csv" if fallback_url.lower().endswith(".csv") else "json"))
    for url, kind in urls:
        try:
            if kind == "csv":
                rows = request_csv_rows(url, timeout=timeout, tries=3)
            else:
                js = request_json(url, timeout=timeout, tries=3)
                if not isinstance(js, list):
                    raise RuntimeError(f"回傳格式非 list: {type(js).__name__}")
                rows = js
            if rows:
                return rows, url
            errors.append(f"{url}: 空資料")
        except Exception as e:
            errors.append(f"{url}: {e}")
    raise RuntimeError("；".join(errors))

def clean_num(v: Any) -> Optional[float]:
    if v is None:
        return None
    s = str(v).strip().replace(",", "")
    s = s.replace("--", "").replace("---", "")
    if s in {"", "-", "N/A", "nan", "None"}:
        return None
    s = s.replace("X", "").replace("*", "").replace("#", "").replace("=", "")
    try:
        return float(s)
    except Exception:
        m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
        return float(m.group()) if m else None


def is_common_stock(code: str, name: str) -> bool:
    code = str(code).strip()
    name = str(name).strip()
    # 台灣普通股代碼主要為四位數且 >= 1000；藉此排除 ETF、ETN、權證等。
    if not re.fullmatch(r"\d{4}", code):
        return False
    if int(code) < 1000:
        return False
    # v30：DR/TDR 是存託憑證，不是普通股。舊版只看四位數，9110 等會誤混入普通股母體。
    uname = name.upper()
    if "-DR" in uname or "TDR" in uname or "存託憑證" in name:
        return False
    # 排除常見特別股文字（保留 KY 第一上市普通股）
    if "特別" in name or re.search(r"[甲乙丙丁]特$", name):
        return False
    return True


# v30 特殊財報/營收時限與三率適用性。
# 2800~2899 為銀行、保險、金控等金融類，毛利率/營益率/淨利率「三率」不具一般產業可比性。
# 2207、2905 為官方綜合損益「異業」結構，亦不硬套一般業三率。
THREE_RATE_NA_CODES = {"2905", "5876", "5880"}

# 2026 起：保險業及具保險業子公司之公開發行公司，月營收最晚可延至次月15日。
# 採『最晚法定期限』做歷史 point-in-time，寧可晚用、不提前穿越。
INSURANCE_REVENUE_15_CODES = {
    "2832", "2850", "2851", "2852", "2867",
    "2880", "2881", "2882", "2883", "2885", "2886", "2887",
    "2891", "2892", "2905", "5880",
}


def _stock_identity_finance_like(stock_id: str, name: str = "") -> bool:
    sid = str(stock_id).strip()
    nm = str(name or "")
    try:
        n = int(sid)
    except Exception:
        n = -1
    if 2800 <= n <= 2899:
        return True
    if sid in THREE_RATE_NA_CODES:
        return True
    return any(k in nm for k in ["金控", "銀行", "銀行業", "證券", "產險", "人壽", "壽險", "保險"] )


def _stock_name_asof(conn: sqlite3.Connection, stock_id: str, target: Optional[dt.date] = None) -> str:
    if target is None:
        row = conn.execute("SELECT name FROM prices WHERE stock_id=? ORDER BY date DESC LIMIT 1", (str(stock_id),)).fetchone()
    else:
        row = conn.execute("SELECT name FROM prices WHERE stock_id=? AND date<=? ORDER BY date DESC LIMIT 1", (str(stock_id), target.isoformat())).fetchone()
    return str(row[0] or "") if row else ""


def _is_primary_foreign(stock_id: str, name: str = "") -> bool:
    # 第一上市外國公司在台股名稱通常以 -KY 標示；DR/TDR 已在 ordinary-stock 母體排除。
    return "-KY" in str(name or "").upper()


def _is_three_rate_na_stock(conn: sqlite3.Connection, stock_id: str, target: dt.date) -> bool:
    sid = str(stock_id)
    name = _stock_name_asof(conn, sid, target)
    if _stock_identity_finance_like(sid, name):
        return True
    # 若已抓到官方金融/保險/金控/證券端點，也視為三率不適用；異業若欄位齊全仍可計算。
    row = conn.execute("""SELECT source FROM fundamental_income_history
                          WHERE stock_id=? AND availability_date IS NOT NULL AND availability_date<=?
                          ORDER BY year DESC,quarter DESC LIMIT 1""", (sid, target.isoformat())).fetchone()
    src = str(row[0] or "").lower() if row else ""
    return any(tok in src for tok in ["_basi", "_bd", "_fh", "_ins", "金融", "證券", "金控", "保險"])


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        PRAGMA journal_mode=WAL;
        CREATE TABLE IF NOT EXISTS prices (
            date TEXT NOT NULL,
            stock_id TEXT NOT NULL,
            name TEXT,
            market TEXT NOT NULL,
            open REAL, high REAL, low REAL, close REAL, volume REAL,
            PRIMARY KEY(date, stock_id)
        );
        CREATE INDEX IF NOT EXISTS idx_prices_stock_date ON prices(stock_id, date);

        CREATE TABLE IF NOT EXISTS futures (
            date TEXT NOT NULL,
            contract TEXT NOT NULL,
            open REAL, high REAL, low REAL, close REAL,
            settlement REAL, volume REAL, open_interest REAL,
            PRIMARY KEY(date, contract)
        );

        CREATE TABLE IF NOT EXISTS source_status (
            run_date TEXT NOT NULL,
            source TEXT NOT NULL,
            data_date TEXT,
            success INTEGER NOT NULL,
            complete INTEGER NOT NULL,
            message TEXT,
            checked_at TEXT NOT NULL,
            PRIMARY KEY(run_date, source)
        );

        CREATE TABLE IF NOT EXISTS fundamental_income_snapshot (
            observed_date TEXT NOT NULL,
            stock_id TEXT NOT NULL,
            name TEXT, market TEXT,
            year INTEGER NOT NULL, quarter INTEGER NOT NULL,
            report_generated TEXT,
            revenue_ytd REAL, gross_profit_ytd REAL, operating_income_ytd REAL, net_income_ytd REAL, eps_ytd REAL,
            source TEXT,
            PRIMARY KEY(observed_date, stock_id)
        );
        CREATE INDEX IF NOT EXISTS idx_fund_income_snap_stock_obs ON fundamental_income_snapshot(stock_id, observed_date);

        CREATE TABLE IF NOT EXISTS fundamental_income_history (
            stock_id TEXT NOT NULL,
            year INTEGER NOT NULL, quarter INTEGER NOT NULL,
            revenue_ytd REAL, gross_profit_ytd REAL, operating_income_ytd REAL, net_income_ytd REAL, eps_ytd REAL,
            availability_date TEXT, availability_basis TEXT,
            source TEXT, updated_at TEXT,
            PRIMARY KEY(stock_id, year, quarter)
        );

        CREATE TABLE IF NOT EXISTS fundamental_revenue_snapshot (
            observed_date TEXT NOT NULL,
            stock_id TEXT NOT NULL,
            name TEXT, market TEXT, month_key INTEGER NOT NULL,
            revenue_mom_pct REAL, revenue_yoy_pct REAL, revenue_cum_yoy_pct REAL,
            source TEXT,
            PRIMARY KEY(observed_date, stock_id)
        );
        CREATE INDEX IF NOT EXISTS idx_fund_rev_snap_stock_obs ON fundamental_revenue_snapshot(stock_id, observed_date);

        CREATE TABLE IF NOT EXISTS fundamental_revenue_history (
            stock_id TEXT NOT NULL,
            month_key INTEGER NOT NULL,
            revenue REAL,
            revenue_mom_pct REAL, revenue_yoy_pct REAL, revenue_cum_yoy_pct REAL,
            availability_date TEXT, availability_basis TEXT,
            source TEXT, updated_at TEXT,
            PRIMARY KEY(stock_id, month_key)
        );
        CREATE INDEX IF NOT EXISTS idx_fund_rev_hist_avail ON fundamental_revenue_history(stock_id, availability_date);
        CREATE TABLE IF NOT EXISTS valuation_pe_history (
            stock_id TEXT NOT NULL,
            date TEXT NOT NULL,
            per REAL, pbr REAL, dividend_yield REAL,
            source TEXT,
            PRIMARY KEY(stock_id, date)
        );
        CREATE INDEX IF NOT EXISTS idx_valuation_pe_stock_date ON valuation_pe_history(stock_id, date);
        CREATE TABLE IF NOT EXISTS fundamental_backfill_nodata (
            dataset TEXT NOT NULL,
            stock_id TEXT NOT NULL,
            period TEXT NOT NULL,
            checked_at TEXT NOT NULL,
            PRIMARY KEY(dataset,stock_id,period)
        );

        CREATE TABLE IF NOT EXISTS yfinance_eps_snapshot (
            observed_date TEXT NOT NULL,
            stock_id TEXT NOT NULL,
            market TEXT,
            yahoo_symbol TEXT,
            period_end TEXT NOT NULL,
            basic_eps REAL,
            diluted_eps REAL,
            used_eps REAL,
            source TEXT,
            fetched_at TEXT NOT NULL,
            PRIMARY KEY(observed_date,stock_id,period_end)
        );
        CREATE INDEX IF NOT EXISTS idx_yf_eps_stock_obs ON yfinance_eps_snapshot(stock_id, observed_date);
        """
    )
    # v8 -> v9 SQLite migration：允許沿用既有 stock.db，不必重抓全部股價。
    income_cols = {r[1] for r in conn.execute("PRAGMA table_info(fundamental_income_history)").fetchall()}
    if "availability_date" not in income_cols:
        conn.execute("ALTER TABLE fundamental_income_history ADD COLUMN availability_date TEXT")
    if "availability_basis" not in income_cols:
        conn.execute("ALTER TABLE fundamental_income_history ADD COLUMN availability_basis TEXT")
    if "eps_ytd" not in income_cols:
        conn.execute("ALTER TABLE fundamental_income_history ADD COLUMN eps_ytd REAL")
    snap_cols = {r[1] for r in conn.execute("PRAGMA table_info(fundamental_income_snapshot)").fetchall()}
    if "eps_ytd" not in snap_cols:
        conn.execute("ALTER TABLE fundamental_income_snapshot ADD COLUMN eps_ytd REAL")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_fund_income_hist_avail ON fundamental_income_history(stock_id, availability_date)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_fund_rev_hist_avail ON fundamental_revenue_history(stock_id, availability_date)")
    conn.commit()


def record_source(conn: sqlite3.Connection, run_date: dt.date, source: str, data_date: Optional[str], success: bool, complete: bool, message: str) -> None:
    conn.execute(
        """INSERT OR REPLACE INTO source_status
           (run_date,source,data_date,success,complete,message,checked_at)
           VALUES (?,?,?,?,?,?,?)""",
        (run_date.isoformat(), source, data_date, int(success), int(complete), message, dt.datetime.now().isoformat(timespec="seconds")),
    )
    conn.commit()


def clear_run_source_status(conn: sqlite3.Connection, run_date: dt.date) -> None:
    """重跑同一天前先清除舊版本來源狀態，避免 v10/v11 的 FinMind 失敗列殘留。"""
    conn.execute("DELETE FROM source_status WHERE run_date=?", (run_date.isoformat(),))
    conn.commit()


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
    out: List[dict] = []
    for row in rows:
        try:
            code = str(row[idx_code]).strip()
            name = str(row[idx_name]).strip()
        except Exception:
            continue
        if not is_common_stock(code, name):
            continue
        close = clean_num(row[idx_close])
        if close is None or close <= 0:
            continue
        out.append({
            "stock_id": code,
            "name": name,
            "market": market,
            "open": clean_num(row[idx_open]) if idx_open is not None else None,
            "high": clean_num(row[idx_high]) if idx_high is not None else None,
            "low": clean_num(row[idx_low]) if idx_low is not None else None,
            "close": close,
            "volume": clean_num(row[idx_vol]) if idx_vol is not None else None,
        })
    return out


def _parse_twse_mi_index(js: Any) -> List[dict]:
    if isinstance(js, dict) and str(js.get("stat", "")).upper() not in {"", "OK"}:
        return []
    candidates: List[Tuple[List[str], List[list]]] = []
    if isinstance(js, dict):
        for t in js.get("tables", []) or []:
            fields = t.get("fields") or []
            rows = t.get("data") or []
            if fields and rows:
                candidates.append((fields, rows))
        for k, v in list(js.items()):
            if k.startswith("fields") and isinstance(v, list):
                suffix = k[6:]
                rows = js.get("data" + suffix)
                if isinstance(rows, list):
                    candidates.append((v, rows))
    best: List[dict] = []
    for fields, rows in candidates:
        parsed = normalize_table(fields, rows, "TWSE")
        if len(parsed) > len(best):
            best = parsed
    return best


def fetch_twse_date(day: dt.date, slow_retry: bool = False) -> List[dict]:
    """TWSE 官方歷史日期查詢。

    先走目前網站使用的 rwd/zh/afterTrading/MI_INDEX；若失敗再走舊 exchangeReport。
    初次掃描採短逾時避免單日卡住；第二輪補洞時 slow_retry=True 會延長逾時。
    """
    params = {"response": "json", "date": day.strftime("%Y%m%d"), "type": "ALLBUT0999"}
    timeout = 25 if slow_retry else 7
    tries = 2 if slow_retry else 1
    endpoints = [
        "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX",
        "https://www.twse.com.tw/exchangeReport/MI_INDEX",
    ]
    errors = []
    for url in endpoints:
        try:
            js = request_json(url, params, timeout=timeout, tries=tries)
            rows = _parse_twse_mi_index(js)
            # 非交易日會 stat != OK，rows=[]，這不是錯誤。
            if rows:
                return rows
            if isinstance(js, dict) and str(js.get("stat", "")).upper() not in {"", "OK"}:
                return []
        except Exception as e:
            errors.append(f"{url}: {e}")
    if errors:
        raise RuntimeError(" | ".join(errors))
    return []


def roc_date(day: dt.date) -> str:
    return f"{day.year-1911}/{day.month:02d}/{day.day:02d}"


def fetch_tpex_date(day: dt.date) -> List[dict]:
    """TPEx 官方歷史日期查詢。新版失敗時自動嘗試舊版。"""
    errors = []
    # 新版網站資料端點（官方）
    endpoints = [
        (
            "https://www.tpex.org.tw/www/zh-tw/afterTrading/dailyQuotes",
            {"date": day.strftime("%Y/%m/%d"), "id": "", "response": "json"},
        ),
        # 舊版官方 JSON 端點，仍保留作備援
        (
            "https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php",
            {"l": "zh-tw", "o": "json", "d": roc_date(day), "s": "0,asc,0"},
        ),
    ]
    for url, params in endpoints:
        try:
            js = request_json(url, params)
            # 新版 tables
            candidates: List[Tuple[List[str], List[list]]] = []
            if isinstance(js, dict):
                for t in js.get("tables", []) or []:
                    fields = t.get("fields") or t.get("columns") or []
                    rows = t.get("data") or []
                    if fields and rows:
                        candidates.append((fields, rows))
                if candidates:
                    best: List[dict] = []
                    for fields, rows in candidates:
                        p = normalize_table(fields, rows, "TPEx")
                        if len(p) > len(best):
                            best = p
                    if best:
                        return best
                # 舊版 aaData：欄位順序為 代號、名稱、收盤、漲跌、開盤、最高、最低、均價、成交股數...
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
                        out.append({
                            "stock_id": code, "name": name, "market": "TPEx",
                            "close": close, "open": clean_num(row[4]), "high": clean_num(row[5]),
                            "low": clean_num(row[6]), "volume": clean_num(row[8]),
                        })
                    if out:
                        return out
        except Exception as e:
            errors.append(f"{url}: {e}")
    if errors:
        raise RuntimeError(" | ".join(errors))
    return []


def fetch_twse_openapi_current() -> List[dict]:
    url = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"
    js = request_json(url)
    if not isinstance(js, list):
        return []
    out = []
    for x in js:
        code = str(x.get("Code", x.get("證券代號", ""))).strip()
        name = str(x.get("Name", x.get("證券名稱", ""))).strip()
        if not is_common_stock(code, name):
            continue
        close = clean_num(x.get("ClosingPrice", x.get("Close", x.get("收盤價"))))
        if close is None or close <= 0:
            continue
        out.append({
            "stock_id": code, "name": name, "market": "TWSE",
            "close": close,
            "open": clean_num(x.get("OpeningPrice", x.get("Open"))),
            "high": clean_num(x.get("HighestPrice", x.get("High"))),
            "low": clean_num(x.get("LowestPrice", x.get("Low"))),
            "volume": clean_num(x.get("TradeVolume", x.get("TradingShares"))),
        })
    return out


def fetch_tpex_openapi_current() -> List[dict]:
    urls = [
        "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes",
        "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_quotes",
    ]
    errors: List[str] = []
    for url in urls:
        try:
            js = request_json(url, timeout=20, tries=3)
            if not isinstance(js, list):
                continue
            out = []
            for x in js:
                if not isinstance(x, dict):
                    continue
                code = str(x.get("SecuritiesCompanyCode", x.get("SecuritiesCode", x.get("Code", x.get("股票代號", x.get("代號", "")))))).strip()
                name = str(x.get("CompanyName", x.get("SecuritiesName", x.get("Name", x.get("股票名稱", x.get("名稱", "")))))).strip()
                if not is_common_stock(code, name):
                    continue
                close = clean_num(x.get("Close", x.get("ClosingPrice", x.get("ClosePrice", x.get("收盤價", x.get("收盤"))))))
                if close is None or close <= 0:
                    continue
                out.append({
                    "stock_id": code, "name": name, "market": "TPEx", "close": close,
                    "open": clean_num(x.get("Open") or x.get("OpeningPrice") or x.get("OpenPrice") or x.get("開盤價") or x.get("開盤")),
                    "high": clean_num(x.get("High") or x.get("HighestPrice") or x.get("HighPrice") or x.get("最高價") or x.get("最高")),
                    "low": clean_num(x.get("Low") or x.get("LowestPrice") or x.get("LowPrice") or x.get("最低價") or x.get("最低")),
                    "volume": clean_num(x.get("TradingShares") or x.get("TradeVolume") or x.get("TradingVolume") or x.get("成交股數") or x.get("成交量")),
                })
            if out:
                return out
        except Exception as e:
            errors.append(f"{url}: {e}")
    if errors:
        raise RuntimeError(" | ".join(errors))
    return []


def upsert_prices(conn: sqlite3.Connection, day: dt.date, rows: List[dict]) -> int:
    if not rows:
        return 0
    values = [
        (day.isoformat(), r["stock_id"], r["name"], r["market"], r.get("open"), r.get("high"), r.get("low"), r.get("close"), r.get("volume"))
        for r in rows
    ]
    conn.executemany(
        """INSERT OR REPLACE INTO prices
           (date,stock_id,name,market,open,high,low,close,volume)
           VALUES (?,?,?,?,?,?,?,?,?)""",
        values,
    )
    conn.commit()
    return len(values)


def _price_market_minimums(conn: sqlite3.Connection, day: dt.date) -> Tuple[int, int]:
    min_twse, min_tpex = 700, 550
    row = conn.execute(
        "SELECT MAX(date) FROM prices WHERE date < ?", (day.isoformat(),)
    ).fetchone()
    if row and row[0]:
        prev = row[0]
        counts = dict(conn.execute(
            "SELECT market, COUNT(*) FROM prices WHERE date=? GROUP BY market", (prev,)
        ).fetchall())
        c1, c2 = int(counts.get("TWSE", 0)), int(counts.get("TPEx", 0))
        if c1 >= 700:
            min_twse = max(min_twse, int(c1 * 0.85))
        if c2 >= 550:
            min_tpex = max(min_tpex, int(c2 * 0.85))
    return min_twse, min_tpex


def fetch_and_store_day(conn: sqlite3.Connection, day: dt.date, allow_current_openapi: bool = True, slow_retry: bool = False) -> Tuple[int, int, List[str]]:
    errors: List[str] = []
    twse: List[dict] = []
    tpex: List[dict] = []
    try:
        twse = fetch_twse_date(day, slow_retry=slow_retry)
    except Exception as e:
        errors.append(f"TWSE歷史端點：{e}")
    try:
        tpex = fetch_tpex_date(day)
    except Exception as e:
        errors.append(f"TPEx歷史端點：{e}")

    min_twse, min_tpex = _price_market_minimums(conn, day)
    taipei_today = dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).date()
    if allow_current_openapi and day == taipei_today:
        if len(twse) < min_twse:
            try:
                x = fetch_twse_openapi_current()
                if len(x) > len(twse):
                    twse = x
                    errors.append("TWSE歷史端點不完整，已改用TWSE OpenAPI當日資料")
            except Exception as e:
                errors.append(f"TWSE OpenAPI：{e}")
        if len(tpex) < min_tpex:
            try:
                x = fetch_tpex_openapi_current()
                if len(x) > len(tpex):
                    tpex = x
                    errors.append("TPEx歷史端點不完整，已改用TPEx OpenAPI當日資料")
            except Exception as e:
                errors.append(f"TPEx OpenAPI：{e}")

    # 原版會把單邊市場先寫入 DB，造成殘缺交易日。新版採 atomic gate：兩邊完整才一次提交。
    if len(twse) < min_twse or len(tpex) < min_tpex:
        errors.append(
            f"完整性驗證失敗：TWSE={len(twse)}（至少{min_twse}）、TPEx={len(tpex)}（至少{min_tpex}）；本日不寫入資料庫"
        )
        return 0, 0, errors

    # 若先前版本留下殘缺同日資料，正式提交前先清掉該日再重建，避免混雜。
    conn.execute("DELETE FROM prices WHERE date=?", (day.isoformat(),))
    conn.commit()
    n1 = upsert_prices(conn, day, twse)
    n2 = upsert_prices(conn, day, tpex)
    return n1, n2, errors


def latest_price_date(conn: sqlite3.Connection) -> Optional[dt.date]:
    row = conn.execute("SELECT MAX(date) FROM prices").fetchone()
    return dt.date.fromisoformat(row[0]) if row and row[0] else None


def backfill_prices(conn: sqlite3.Connection, target: dt.date, calendar_days: int) -> None:
    latest = latest_price_date(conn)
    if latest is None:
        start = target - dt.timedelta(days=calendar_days)
        log.info("首次初始化：補抓 %s ~ %s", start, target)
    else:
        start = latest + dt.timedelta(days=1)
        if start > target:
            log.info("股票資料庫已更新到 %s", latest)
            return
        log.info("追加缺少日期：%s ~ %s", start, target)

    d = start
    success_days = 0
    retry_days: List[dt.date] = []
    while d <= target:
        if d.weekday() < 5:
            try:
                n1, n2, errors = fetch_and_store_day(conn, d, allow_current_openapi=True, slow_retry=False)
                if n1 or n2:
                    success_days += 1
                    log.info("%s 上市=%d 上櫃=%d", d, n1, n2)
                # 只要有任一市場歷史端點發生真正連線錯誤，就排入第二輪補洞。
                if any("歷史端點" in e and ("GET JSON 失敗" in e or "RuntimeError" in e or "timed out" in e.lower()) for e in errors):
                    retry_days.append(d)
                for e in errors:
                    log.warning("%s | %s", d, e)
            except Exception as e:
                retry_days.append(d)
                log.warning("%s 抓取失敗：%s", d, e)
            time.sleep(0.35)
        d += dt.timedelta(days=1)

    # 第二輪只重試第一輪出錯的日期；避免初始化被單日 12~30 秒卡住。
    if retry_days:
        uniq = sorted(set(retry_days))
        log.info("第一輪完成；開始補抓 %d 個疑似缺漏日期（延長逾時）", len(uniq))
        for d in uniq:
            try:
                n1, n2, errors = fetch_and_store_day(conn, d, allow_current_openapi=True, slow_retry=True)
                log.info("補抓 %s 上市=%d 上櫃=%d", d, n1, n2)
                for e in errors:
                    log.warning("補抓 %s | %s", d, e)
            except Exception as e:
                log.error("補抓 %s 仍失敗：%s", d, e)
            time.sleep(0.25)

    log.info("本次第一輪成功寫入 %d 個交易日", success_days)


def compute_signals(conn: sqlite3.Connection, target: dt.date, cfg: Optional[configparser.ConfigParser] = None) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df = pd.read_sql_query(
        "SELECT date,stock_id,name,market,close FROM prices WHERE date<=? ORDER BY stock_id,date",
        conn,
        params=(target.isoformat(),),
    )
    if df.empty:
        return (pd.DataFrame(),) * 6
    # v30：舊 stock.db 可能已存入 DR/TDR；即使不清 DB，也在計算前重新套普通股母體規則。
    df = df[df.apply(lambda r: is_common_stock(str(r.get("stock_id") or ""), str(r.get("name") or "")), axis=1)].copy()
    if df.empty:
        return (pd.DataFrame(),) * 6
    df["date"] = pd.to_datetime(df["date"])
    df["close"] = pd.to_numeric(df["close"], errors="coerce")
    g = df.groupby("stock_id", group_keys=False)
    df["MA20"] = g["close"].transform(lambda s: s.rolling(20, min_periods=20).mean())
    df["MA60"] = g["close"].transform(lambda s: s.rolling(60, min_periods=60).mean())
    df["MA20_prev"] = g["MA20"].shift(1)
    df["MA60_prev"] = g["MA60"].shift(1)
    df["MA20_change"] = df["MA20"] - df["MA20_prev"]
    df["MA60_change"] = df["MA60"] - df["MA60_prev"]
    df["spread"] = df["MA20"] - df["MA60"]
    df["spread_prev"] = g["spread"].shift(1)
    df["spread_change"] = df["spread"] - df["spread_prev"]
    df["bias20_pct"] = (df["close"] / df["MA20"] - 1) * 100
    df["duck"] = (
        (df["close"] > df["MA20"]) &
        (df["close"] > df["MA60"]) &
        (df["MA20"] > df["MA60"]) &
        (df["MA20_change"] > 0) &
        (df["MA60_change"] > 0) &
        (df["spread_change"] > 0)
    )
    df["duck_prev"] = g["duck"].shift(1).fillna(False).astype(bool)
    df["new"] = df["duck"] & ~df["duck_prev"]
    df["exit"] = ~df["duck"] & df["duck_prev"]

    target_ts = pd.Timestamp(target)
    today = df[df["date"] == target_ts].copy()
    if today.empty:
        return today, pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), today

    # 找目前鴨嘴段的首次過熱日。只針對當日仍成立者。
    first_overheat: Dict[str, Optional[str]] = {}
    had_overheat: Dict[str, bool] = {}
    for sid in today.loc[today["duck"], "stock_id"].tolist():
        hist = df[df["stock_id"] == sid].sort_values("date")
        idxs = hist.index.tolist()
        cur_idx = hist.index[-1]
        # 自最後一個 duck=False 之後視為目前這一段
        segment = hist.copy()
        false_pos = segment.index[~segment["duck"]].tolist()
        if false_pos:
            last_false = false_pos[-1]
            pos = segment.index.get_loc(last_false)
            segment = segment.iloc[pos+1:]
        oh = segment[segment["bias20_pct"] >= 8]
        had = not oh.empty
        had_overheat[sid] = had
        first_overheat[sid] = oh.iloc[0]["date"].date().isoformat() if had else None

    def status_row(r: pd.Series) -> str:
        if bool(r["new"]):
            base = "新進"
        elif bool(r["duck"]):
            base = "持續符合"
        elif bool(r["exit"]):
            return "退出"
        else:
            return ""
        bias = r["bias20_pct"]
        if pd.notna(bias) and bias >= 8:
            return base + "／不建議追價／過熱"
        if bool(r["duck"]) and had_overheat.get(r["stock_id"], False):
            if pd.notna(bias) and bias <= 5:
                return base + "／過熱解除"
            return base + "／曾過熱，降溫中"
        return base

    today["首次過熱日期"] = today["stock_id"].map(first_overheat)
    today["狀態"] = today.apply(status_row, axis=1)

    # ------------------------------------------------------------
    # v21 預備鴨嘴：在正式 5/5 成立前，先抓「只差 1~2 條且差距已很小」的股票。
    # 目的不是預測一定會成立，而是把觀察時間往前移，降低正式成立後才追高的情況。
    # ------------------------------------------------------------
    if cfg is None:
        cfg = load_config()
    enabled = cfg.getint("general", "preduck_enabled", fallback=1) == 1
    price_near_pct = max(0.0, cfg.getfloat("general", "preduck_price_near_pct", fallback=2.0))
    cross_near_pct = max(0.0, cfg.getfloat("general", "preduck_ma_cross_near_pct", fallback=1.0))
    ma20_slope_near_pct = max(0.0, cfg.getfloat("general", "preduck_ma20_slope_near_pct", fallback=0.20))
    ma60_slope_near_pct = max(0.0, cfg.getfloat("general", "preduck_ma60_slope_near_pct", fallback=0.10))
    spread_near_pct = max(0.0, cfg.getfloat("general", "preduck_spread_near_pct", fallback=0.10))
    min_pass = min(4, max(3, cfg.getint("general", "preduck_min_pass", fallback=3)))

    today["cond_price"] = (today["close"] > today["MA20"]) & (today["close"] > today["MA60"])
    today["cond_cross"] = today["MA20"] > today["MA60"]
    today["cond_ma20_up"] = today["MA20_change"] > 0
    today["cond_ma60_up"] = today["MA60_change"] > 0
    today["cond_spread_up"] = today["spread_change"] > 0
    cond_cols = ["cond_price","cond_cross","cond_ma20_up","cond_ma60_up","cond_spread_up"]
    today["鴨嘴完成條件數"] = today[cond_cols].sum(axis=1).astype(int)
    today["鴨嘴完成度%"] = today["鴨嘴完成條件數"] * 20
    today["距MA20%"] = (today["close"] / today["MA20"] - 1) * 100
    today["MA20距MA60%"] = (today["MA20"] / today["MA60"] - 1) * 100

    # 每個「尚未通過」條件，都必須已接近門檻，才列入預備名單。
    near_price = today["close"] >= today[["MA20","MA60"]].max(axis=1) * (1 - price_near_pct / 100.0)
    near_cross = today["MA20"] >= today["MA60"] * (1 - cross_near_pct / 100.0)
    near_ma20 = today["MA20_change"] >= -(today["MA20"].abs() * ma20_slope_near_pct / 100.0)
    near_ma60 = today["MA60_change"] >= -(today["MA60"].abs() * ma60_slope_near_pct / 100.0)
    near_spread = today["spread_change"] >= -(today["MA60"].abs() * spread_near_pct / 100.0)
    near_map = {
        "cond_price": near_price,
        "cond_cross": near_cross,
        "cond_ma20_up": near_ma20,
        "cond_ma60_up": near_ma60,
        "cond_spread_up": near_spread,
    }

    labels = {
        "cond_price": "收盤尚未同時站上MA20/MA60",
        "cond_cross": "MA20尚未站上MA60",
        "cond_ma20_up": "MA20尚未上揚",
        "cond_ma60_up": "MA60尚未上揚",
        "cond_spread_up": "MA20-MA60開口尚未擴大",
    }

    def preduck_row(r: pd.Series) -> Tuple[str, str, bool]:
        if not enabled or bool(r.get("duck", False)):
            return "", "", False
        passed = int(r.get("鴨嘴完成條件數", 0))
        if passed < min_pass or passed > 4:
            return "", "", False
        failed = [c for c in cond_cols if not bool(r.get(c, False))]
        # 缺少的每一條都要在「接近門檻」範圍內。
        idx = r.name
        for c in failed:
            nm = near_map[c]
            try:
                if not bool(nm.loc[idx]):
                    return "", "", False
            except Exception:
                return "", "", False
        # B級至少要有一個動能條件已轉正，避免只因均線距離很近就過早列入。
        if passed == 3 and not (bool(r.get("cond_ma20_up")) or bool(r.get("cond_ma60_up")) or bool(r.get("cond_spread_up"))):
            return "", "", False
        status = "A級：臨門一腳（4/5）" if passed == 4 else "B級：接近形成（3/5）"
        missing = "；".join(labels[c] for c in failed)
        return status, missing, True

    pre_info = today.apply(preduck_row, axis=1, result_type="expand")
    if not pre_info.empty:
        pre_info.columns = ["預備鴨嘴狀態","尚缺條件","preduck"]
        today[["預備鴨嘴狀態","尚缺條件","preduck"]] = pre_info
    else:
        today["預備鴨嘴狀態"] = ""
        today["尚缺條件"] = ""
        today["preduck"] = False
    today["preduck"] = today["preduck"].fillna(False).astype(bool)

    cols = ["date","stock_id","name","market","close","MA20","MA60","MA20_change","MA60_change","spread","spread_change","bias20_pct","首次過熱日期","狀態"]
    current = today[today["duck"]][cols].copy().sort_values(["bias20_pct","stock_id"], ascending=[False, True])
    new = today[today["new"]][cols].copy().sort_values("stock_id")
    exits = today[today["exit"]][cols].copy().sort_values("stock_id")
    hot = current[current["bias20_pct"] >= 8].copy()

    pre_cols = cols[:-2] + ["距MA20%","MA20距MA60%","鴨嘴完成條件數","鴨嘴完成度%","預備鴨嘴狀態","尚缺條件"]
    preduck = today[today["preduck"]][pre_cols].copy()
    if not preduck.empty:
        preduck = preduck.sort_values(["鴨嘴完成條件數","bias20_pct","stock_id"], ascending=[False, True, True], na_position="last")

    # v23：保留當日全市場技術狀態，之後和營收／基本面／估值合成一張主表。
    def all_missing(r: pd.Series) -> str:
        return "；".join(labels[c] for c in cond_cols if not bool(r.get(c, False)))
    today["鴨嘴未達條件"] = today.apply(all_missing, axis=1)
    def stage_row(r: pd.Series) -> str:
        if bool(r.get("duck", False)):
            return str(r.get("狀態") or "正式鴨嘴")
        if bool(r.get("preduck", False)):
            return str(r.get("預備鴨嘴狀態") or "預備鴨嘴")
        if bool(r.get("exit", False)):
            return "今日退出"
        return "未符合"
    today["鴨嘴階段"] = today.apply(stage_row, axis=1)
    return current, new, exits, hot, preduck, today



# ============================
# 基本面培育中心：Point-in-time 安全版
# ============================
# 三率三升：以「單季」毛利率、營業利益率、稅後純益率比較最新兩季。
# 官方 t187ap06 為累季損益，因此 Q2/Q3/Q4 必須先扣除前一累季，還原單季後再計算三率。
# 最新已申報季度：TWSE/TPEx 官方綜合損益表快照。
# 前期歷史：只使用 SQLite 已保存之官方資料；不足時明確標示，不以付費 API 補洞。
# 這樣可避免 8/14 才出現的 Q2 資料回頭污染 8/13 的歷史重跑。
# 營收三增：官方月營收最新快照；歷史重跑只使用 observed_date <= target 的快照。

# 基本面優先走 MOPS 官方 CSV；OpenAPI 當備援。
# 原因：TPEx OpenAPI 偶爾會對大型財報 JSON 發生 WinError 10054 / ConnectionResetError。
MOPS_TWSE_INCOME_CSV = "https://mopsfin.twse.com.tw/opendata/t187ap06_L_ci.csv"
MOPS_TPEX_INCOME_CSV = "https://mopsfin.twse.com.tw/opendata/t187ap06_O_ci.csv"
MOPS_TWSE_REVENUE_CSV = "https://mopsfin.twse.com.tw/opendata/t187ap05_L.csv"
MOPS_TPEX_REVENUE_CSV = "https://mopsfin.twse.com.tw/opendata/t187ap05_O.csv"
TWSE_INCOME_URL = "https://openapi.twse.com.tw/v1/opendata/t187ap06_L_ci"
TPEX_INCOME_URL = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_ci"
TWSE_REVENUE_URL = "https://openapi.twse.com.tw/v1/opendata/t187ap05_L"
TPEX_REVENUE_URL = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap05_O"


def _norm_key(v: Any) -> str:
    return re.sub(r"[\s　()（）%％/／_\-]", "", str(v or "")).lower()


def _row_value(row: dict, aliases: Iterable[str]) -> Any:
    if not isinstance(row, dict):
        return None
    nk = {_norm_key(k): k for k in row.keys()}
    for a in aliases:
        na = _norm_key(a)
        if na in nk:
            return row.get(nk[na])
    for a in aliases:
        na = _norm_key(a)
        for k0, real in nk.items():
            if na and (na in k0 or k0 in na):
                return row.get(real)
    return None


def _as_int(v: Any) -> Optional[int]:
    n = clean_num(v)
    return int(n) if n is not None else None


def _period_year_quarter(row: dict) -> Tuple[Optional[int], Optional[int]]:
    y = _as_int(_row_value(row, ["年度", "年", "Year"]))
    qv = _row_value(row, ["季別", "季", "Quarter"])
    q = _as_int(qv)
    if q is None and qv is not None:
        m = re.search(r"([1-4])", str(qv))
        q = int(m.group(1)) if m else None
    if y is not None and y < 1911:
        y += 1911
    return y, q


def _parse_month_key(v: Any) -> Optional[int]:
    if v is None:
        return None
    digits = re.sub(r"\D", "", str(v))
    if len(digits) >= 5:
        y, m = int(digits[:-2]), int(digits[-2:])
        if y < 1911:
            y += 1911
        if 1 <= m <= 12:
            return y * 100 + m
    return None


def _parse_roc_or_gregorian_date(v: Any) -> Optional[str]:
    if v is None:
        return None
    d = re.sub(r"\D", "", str(v))
    try:
        if len(d) == 7:  # 民國 yyyMMdd
            y, m, day = int(d[:3]) + 1911, int(d[3:5]), int(d[5:7])
        elif len(d) == 8:
            y, m, day = int(d[:4]), int(d[4:6]), int(d[6:8])
        else:
            return None
        return dt.date(y, m, day).isoformat()
    except Exception:
        return None


def _month_revenue_conservative_available(year: int, month: int) -> dt.date:
    """
    歷史月營收若沒有實際 create_time，採保守可得日：次月 10 日。
    這不會把尚未依法申報的資料提前用到過去日期；若 FinMind 有 create_time 則優先用實際日期。
    """
    if month == 12:
        return dt.date(year + 1, 1, 10)
    return dt.date(year, month + 1, 10)


def _income_conservative_available(year: int, quarter: int) -> dt.date:
    """
    歷史季報缺少實際申報日時的保守可得日。
    Q1/Q2/Q3 採季末後 45 日；Q4 年報採次年 3/31。
    目的不是猜實際公告日，而是避免 look-ahead；當官方每日快照存在時仍以官方快照為準。
    """
    if quarter == 4:
        return dt.date(year + 1, 3, 31)
    end_month = quarter * 3
    # 季末固定為 3/31、6/30、9/30
    end_day = 31 if end_month in (3,) else 30
    return dt.date(year, end_month, end_day) + dt.timedelta(days=45)


def _income_available_for_identity(stock_id: str, name: str, year: int, quarter: int, source: str = "") -> dt.date:
    """逐公司保守財報可得日。

    - 一般上市櫃：沿用 Q1/Q2/Q3 季末+45日，Q4 次年3/31。
    - 第一上市(KY)第二季：法定最晚為第二季終了後2個月，因此 6/30 -> 8/31。
    - 金融/保險/金控：第二季實務法定期限亦較一般業晚；為避免歷史穿越，Q2採8/31保守日。
    """
    src=str(source or "").lower()
    source_finance=any(tok in src for tok in ["_basi","_bd","_fh","_ins","金融","證券","金控","保險"])
    if int(quarter) == 2 and (_is_primary_foreign(stock_id, name) or _stock_identity_finance_like(stock_id, name) or source_finance):
        return dt.date(int(year), 8, 31)
    return _income_conservative_available(int(year), int(quarter))


def _income_available_for_stock(conn: sqlite3.Connection, stock_id: str, target: dt.date, year: int, quarter: int) -> dt.date:
    return _income_available_for_identity(str(stock_id), _stock_name_asof(conn, str(stock_id), target), int(year), int(quarter))


def _month_revenue_available_for_identity(stock_id: str, name: str, year: int, month: int) -> dt.date:
    sid = str(stock_id)
    # 115(2026)會計年度起，保險業/具保險子公司公發公司可延至次月15日。
    day = 15 if int(year) >= 2026 and sid in INSURANCE_REVENUE_15_CODES else 10
    if int(month) == 12:
        return dt.date(int(year)+1, 1, day)
    return dt.date(int(year), int(month)+1, day)


def _month_revenue_available_for_stock(conn: sqlite3.Connection, stock_id: str, target: dt.date, year: int, month: int) -> dt.date:
    return _month_revenue_available_for_identity(str(stock_id), _stock_name_asof(conn, str(stock_id), target), int(year), int(month))


def _available_quarters_for_stock(conn: sqlite3.Connection, stock_id: str, target: dt.date, n: int) -> List[Tuple[int,int]]:
    vals: List[Tuple[int,int]] = []
    q=(target.month-1)//3+1; y=target.year
    for _ in range(n+10):
        if _income_available_for_stock(conn,stock_id,target,y,q) <= target:
            vals.append((y,q))
            if len(vals)>=n:
                break
        q-=1
        if q==0:
            y-=1; q=4
    return sorted(vals)


def _three_rate_required_periods_for_stock(conn: sqlite3.Connection, stock_id: str, target: dt.date) -> List[Tuple[int,int]]:
    latest=_available_quarters_for_stock(conn,stock_id,target,1)
    if not latest:
        return []
    cy,cq=latest[-1]
    py,pq=(cy,cq-1) if cq>1 else (cy-1,4)
    req={(cy,cq),(py,pq)}
    if cq>1:
        req.add((cy,cq-1))
    if pq>1:
        req.add((py,pq-1))
    return sorted(req)


def _available_month_keys_for_stock(conn: sqlite3.Connection, stock_id: str, target: dt.date, n: int) -> List[int]:
    vals=[]; y,m=target.year,target.month
    for _ in range(n+8):
        m-=1
        if m==0:
            y-=1; m=12
        if _month_revenue_available_for_stock(conn,stock_id,target,y,m) <= target:
            vals.append(y*100+m)
            if len(vals)>=n:
                break
    return sorted(vals)


def _normalize_special_availability_dates(conn: sqlite3.Connection, target: dt.date) -> None:
    """修正舊版 DB 已用一般公司 deadline 寫入的保守 availability_date。

    只動 conservative 類 basis；真正 live snapshot 的觀測日不改。
    """
    # 季報：KY/金融類 Q2 若舊版標成 8/14，往後修到 8/31。
    rows=conn.execute("""SELECT stock_id,year,quarter,availability_date,availability_basis,source
                         FROM fundamental_income_history
                         WHERE quarter=2 AND year>=2021 AND availability_date IS NOT NULL""").fetchall()
    for sid,y,q,ad,basis,source in rows:
        b=str(basis or "").lower()
        if "conservative" not in b and "deadline" not in b:
            continue
        name=_stock_name_asof(conn,str(sid),target)
        newd=_income_available_for_identity(str(sid),name,int(y),int(q),str(source or "")).isoformat()
        if str(ad) < newd:
            conn.execute("UPDATE fundamental_income_history SET availability_date=? WHERE stock_id=? AND year=? AND quarter=?",(newd,str(sid),int(y),int(q)))
    # 月營收：2026起特殊公司採次月15日；同樣只調 conservative 類 basis。
    rows=conn.execute("""SELECT stock_id,month_key,availability_date,availability_basis
                         FROM fundamental_revenue_history
                         WHERE month_key>=202601 AND availability_date IS NOT NULL""").fetchall()
    for sid,mk,ad,basis in rows:
        if str(sid) not in INSURANCE_REVENUE_15_CODES:
            continue
        b=str(basis or "").lower()
        if "live_snapshot" in b:
            continue
        if not any(k in b for k in ["conservative","month10","safe_asof"]):
            continue
        y,m=int(mk)//100,int(mk)%100
        newd=_month_revenue_available_for_identity(str(sid),_stock_name_asof(conn,str(sid),target),y,m).isoformat()
        if str(ad) < newd:
            conn.execute("UPDATE fundamental_revenue_history SET availability_date=? WHERE stock_id=? AND month_key=?",(newd,str(sid),int(mk)))
    conn.commit()


def _save_income_snapshot(conn: sqlite3.Connection, observed: dt.date, rows: List[dict]) -> None:
    if not rows:
        return
    vals = []
    hist = []
    now = dt.datetime.now().isoformat(timespec="seconds")
    for r in rows:
        vals.append((observed.isoformat(), r["stock_id"], r.get("name"), r.get("market"), int(r["year"]), int(r["quarter"]),
                     r.get("report_generated"), r.get("revenue_ytd"), r.get("gross_profit_ytd"), r.get("operating_income_ytd"), r.get("net_income_ytd"), r.get("eps_ytd"), r.get("source")))
        hist.append((r["stock_id"], int(r["year"]), int(r["quarter"]), r.get("revenue_ytd"), r.get("gross_profit_ytd"),
                     r.get("operating_income_ytd"), r.get("net_income_ytd"), r.get("eps_ytd"), observed.isoformat(), "official_live_snapshot", r.get("source"), now))
    conn.executemany("""INSERT OR REPLACE INTO fundamental_income_snapshot
        (observed_date,stock_id,name,market,year,quarter,report_generated,revenue_ytd,gross_profit_ytd,operating_income_ytd,net_income_ytd,eps_ytd,source)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", vals)
    conn.executemany("""INSERT INTO fundamental_income_history
        (stock_id,year,quarter,revenue_ytd,gross_profit_ytd,operating_income_ytd,net_income_ytd,eps_ytd,availability_date,availability_basis,source,updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
        ON CONFLICT(stock_id,year,quarter) DO UPDATE SET
        revenue_ytd=excluded.revenue_ytd,gross_profit_ytd=excluded.gross_profit_ytd,
        operating_income_ytd=excluded.operating_income_ytd,net_income_ytd=excluded.net_income_ytd,eps_ytd=COALESCE(excluded.eps_ytd,fundamental_income_history.eps_ytd),
        availability_date=CASE
            WHEN fundamental_income_history.availability_date IS NULL THEN excluded.availability_date
            WHEN excluded.availability_basis='official_live_snapshot' AND excluded.availability_date < fundamental_income_history.availability_date THEN excluded.availability_date
            ELSE fundamental_income_history.availability_date END,
        availability_basis=CASE
            WHEN fundamental_income_history.availability_date IS NULL THEN excluded.availability_basis
            WHEN excluded.availability_basis='official_live_snapshot' AND excluded.availability_date <= fundamental_income_history.availability_date THEN excluded.availability_basis
            ELSE fundamental_income_history.availability_basis END,
        source=excluded.source,updated_at=excluded.updated_at""", hist)
    conn.commit()


def _latest_income_snapshot(conn: sqlite3.Connection, target: dt.date, stock_ids: Optional[List[str]] = None) -> pd.DataFrame:
    params: List[Any] = [target.isoformat(), target.isoformat()]
    stock_filter = ""
    if stock_ids:
        marks = ",".join("?" for _ in stock_ids)
        stock_filter = f" AND s.stock_id IN ({marks})"
        params.extend(stock_ids)
    q = f"""
        SELECT s.* FROM fundamental_income_snapshot s
        JOIN (
            SELECT stock_id, MAX(observed_date) obs
            FROM fundamental_income_snapshot
            WHERE observed_date<=?
            GROUP BY stock_id
        ) x ON s.stock_id=x.stock_id AND s.observed_date=x.obs
        WHERE s.observed_date<=? {stock_filter}
    """
    return pd.read_sql_query(q, conn, params=params)



INCOME_TYPES = ("ci", "basi", "bd", "fh", "ins", "mim")


def _parse_official_income_row(r: dict, market: str, source: str) -> Optional[dict]:
    code = str(_row_value(r, ["公司代號", "股票代號", "證券代號", "Code"]) or "").strip()
    name = str(_row_value(r, ["公司名稱", "名稱", "Name"]) or "").strip()
    if not is_common_stock(code, name):
        return None
    y, q = _period_year_quarter(r)
    revenue = clean_num(_row_value(r, [
        "營業收入", "營業收入合計", "營業收入淨額", "營業收入總額", "營業收益", "營業收益合計",
        "收益合計", "收益", "收入合計", "淨收益", "利息淨收益", "淨利息收益", "手續費淨收益",
        "Revenue", "TotalRevenue", "OperatingRevenue"
    ]))
    gross = clean_num(_row_value(r, [
        "營業毛利（毛損）淨額", "營業毛利(毛損)淨額", "營業毛利（毛損）", "營業毛利(毛損)",
        "營業毛利", "營業毛利淨額", "GrossProfit"
    ]))
    op = clean_num(_row_value(r, [
        "營業利益（損失）", "營業利益(損失)", "營業利益（損失）淨額", "營業利益(損失)淨額",
        "營業利益", "營業淨利", "營業利益淨額", "營業損益", "OperatingIncome", "OperatingIncomeLoss"
    ]))
    net = clean_num(_row_value(r, [
        "本期淨利（淨損）", "本期淨利(淨損)", "本期淨利", "本期損益", "本期稅後淨利（淨損）",
        "本期稅後淨利(淨損)", "本期稅後淨利", "歸屬於母公司業主之淨利（損）",
        "歸屬於母公司業主之淨利(損)", "歸屬母公司業主淨利（損）", "歸屬母公司業主淨利(損)",
        "NetIncome", "IncomeAfterTaxes", "NetIncomeLoss"
    ]))
    eps = clean_num(_row_value(r, ["基本每股盈餘（元）", "基本每股盈餘(元)", "基本每股盈餘", "EPS", "BasicEPS"]))
    if y is None or q is None:
        return None
    # v30：金融/保險/金控/證券/異業等報表結構與一般業不同，未必有毛利或「營業利益」欄。
    # 這些公司若至少有淨利與一個上層收益欄，仍保存該季，後續標成「三率不適用」，
    # 不再把「報表存在但科目不同」誤判成資料缺漏。一般業仍要求營收/營益/淨利完整。
    source_l = str(source).lower()
    finance_like = any(tok in source_l for tok in ["_basi", "_bd", "_fh", "_ins", "_mim", "金融", "證券", "金控", "保險", "異業"])
    if finance_like:
        if net is None or (revenue is None and op is None):
            return None
    elif revenue is None or op is None or net is None:
        return None
    return {
        "stock_id": code, "name": name, "market": market, "year": int(y), "quarter": int(q),
        "report_generated": _parse_roc_or_gregorian_date(_row_value(r, ["出表日期", "Date"])),
        "revenue_ytd": revenue, "gross_profit_ytd": gross, "operating_income_ytd": op,
        "net_income_ytd": net, "eps_ytd": eps, "source": source,
    }


def _fetch_official_income_bulk_latest(stock_ids: Optional[List[str]] = None) -> Tuple[List[dict], List[str]]:
    """一次抓完上市/上櫃六種產業型態的最新綜合損益表。

    v24 只抓一般業(ci)，會漏掉金融、證券期貨、金控、保險、異業。
    v25 改用 TWSE/TPEx OpenAPI 六類端點全部合併；單一端點失敗不拖垮其他類型。
    """
    wanted = set(str(x) for x in (stock_ids or []))
    rows: List[dict] = []
    errors: List[str] = []
    seen = set()
    # v30：除了目前上市(L)/上櫃(O)，再補「公發公司(X)」與「興櫃(U)」財報。
    # 新上市/轉板公司常見前一季仍掛在公發或興櫃端點；只抓 L/O 會把可取得的比較季誤判成缺漏。
    # 順序固定：目前市場主端點優先，其次財報資訊A，再來跨市場歷史身分；相同 stock/period 不覆蓋主端點。
    endpoint_sets = [
        ("TWSE", "https://openapi.twse.com.tw/v1/opendata/t187ap06_L_{}", "L"),
        ("TPEx", "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_{}", "O"),
        ("TPEx", "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_O_{}A", "OA"),
        ("TWSE", "https://openapi.twse.com.tw/v1/opendata/t187ap06_X_{}", "X"),
        ("TPEx", "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap06_U_{}", "U"),
    ]
    for market, template, variant in endpoint_sets:
        for typ in INCOME_TYPES:
            url = template.format(typ)
            try:
                js = request_json(url, timeout=18, tries=3)
                if not isinstance(js, list):
                    raise RuntimeError("OpenAPI 回傳非陣列")
                for raw in js:
                    if not isinstance(raw, dict):
                        continue
                    suffix = "A" if variant == "OA" else ""
                    rec = _parse_official_income_row(raw, market, f"{market} OpenAPI {variant} t187ap06_{typ}{suffix}")
                    if not rec:
                        continue
                    if wanted and rec["stock_id"] not in wanted:
                        continue
                    key = (rec["stock_id"], rec["year"], rec["quarter"])
                    if key in seen:
                        continue
                    seen.add(key); rows.append(rec)
                # 空端點不一定是錯（例如該市場可能沒有某類公司），不列為 failure。
            except Exception as e:
                errors.append(f"{market}-{typ}{variant}:{type(e).__name__}:{str(e)[:100]}")
    return rows, errors


def _upsert_income_record_with_basis(conn: sqlite3.Connection, r: dict, availability_date: dt.date,
                                     basis: str, source: Optional[str] = None) -> None:
    now = dt.datetime.now().isoformat(timespec="seconds")
    conn.execute("""INSERT INTO fundamental_income_history
        (stock_id,year,quarter,revenue_ytd,gross_profit_ytd,operating_income_ytd,net_income_ytd,eps_ytd,
         availability_date,availability_basis,source,updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(stock_id,year,quarter) DO UPDATE SET
        revenue_ytd=COALESCE(excluded.revenue_ytd,fundamental_income_history.revenue_ytd),
        gross_profit_ytd=COALESCE(excluded.gross_profit_ytd,fundamental_income_history.gross_profit_ytd),
        operating_income_ytd=COALESCE(excluded.operating_income_ytd,fundamental_income_history.operating_income_ytd),
        net_income_ytd=COALESCE(excluded.net_income_ytd,fundamental_income_history.net_income_ytd),
        eps_ytd=COALESCE(excluded.eps_ytd,fundamental_income_history.eps_ytd),
        availability_date=CASE WHEN fundamental_income_history.availability_date IS NULL OR excluded.availability_date < fundamental_income_history.availability_date
                               THEN excluded.availability_date ELSE fundamental_income_history.availability_date END,
        availability_basis=CASE WHEN fundamental_income_history.availability_date IS NULL OR excluded.availability_date <= fundamental_income_history.availability_date
                                THEN excluded.availability_basis ELSE fundamental_income_history.availability_basis END,
        source=COALESCE(excluded.source,fundamental_income_history.source),updated_at=excluded.updated_at""",
        (str(r["stock_id"]), int(r["year"]), int(r["quarter"]), r.get("revenue_ytd"), r.get("gross_profit_ytd"),
         r.get("operating_income_ytd"), r.get("net_income_ytd"), r.get("eps_ytd"), availability_date.isoformat(),
         basis, source or r.get("source"), now))


def _seed_bulk_income_for_target(conn: sqlite3.Connection, target: dt.date, stock_ids: List[str]) -> Tuple[int, List[str]]:
    """只補『當期應有且資料庫尚缺』的最新季，避免每天重抓同一季財報。"""
    if not stock_ids:
        return 0, []
    missing=[]
    for sid in dict.fromkeys(str(x) for x in stock_ids):
        if _is_three_rate_na_stock(conn,sid,target):
            continue
        latest=_available_quarters_for_stock(conn,sid,target,1)
        if not latest:
            continue
        y,q=latest[-1]
        got=conn.execute("""SELECT 1 FROM fundamental_income_history
                            WHERE stock_id=? AND year=? AND quarter=? AND availability_date<=? LIMIT 1""",
                         (sid,int(y),int(q),target.isoformat())).fetchone()
        if not got:
            missing.append(sid)
    if not missing:
        return 0, []
    rows, errors = _fetch_official_income_bulk_latest(missing)
    added = 0
    miss=set(missing)
    for r in rows:
        if str(r.get("stock_id")) not in miss:
            continue
        avail = _income_available_for_identity(str(r["stock_id"]), str(r.get("name") or _stock_name_asof(conn,str(r["stock_id"]),target)), int(r["year"]), int(r["quarter"]), str(r.get("source") or ""))
        if avail > target:
            continue
        _upsert_income_record_with_basis(conn, r, avail, "official_bulk_conservative_deadline", r.get("source"))
        added += 1
    if added:
        conn.commit()
    return added, errors

def fetch_official_income_snapshot(conn: sqlite3.Connection, target: dt.date, stock_ids: Optional[List[str]] = None) -> pd.DataFrame:
    # 歷史重跑禁止把「現在網路看到的最新財報」回填成過去日期。
    # 只有 target == 今天時才建立新快照；舊日期一律只讀當時已存下來的 snapshot。
    if target != dt.date.today():
        wanted=set(str(x) for x in (stock_ids or []))
        hist = _latest_income_history_asof(conn, target, stock_ids)
        # v30：最新應有季度改為『逐公司』判斷。KY/金融Q2不再被8/14的45日通則誤判為缺漏。
        latest_by_sid={}
        if not hist.empty:
            for _,r in hist.iterrows():
                latest_by_sid[str(r["stock_id"])]=(int(r["year"]),int(r["quarter"]))
        required_period={}
        exempt=[]
        for sid in wanted:
            if _is_three_rate_na_stock(conn,sid,target):
                exempt.append(sid)
                continue
            latest=_available_quarters_for_stock(conn,sid,target,1)
            if latest:
                required_period[sid]=latest[-1]
        missing=sorted(sid for sid,pq in required_period.items() if latest_by_sid.get(sid)!=pq)
        complete=(len(missing)==0)
        # 對需做一般業三率者只保留其 target 當日最新可得季；特殊業別可保留已有資料供其他欄位使用。
        if not hist.empty and required_period:
            hist=hist[hist.apply(lambda r: (str(r["stock_id"]) in exempt) or required_period.get(str(r["stock_id"]))==(int(r["year"]),int(r["quarter"])),axis=1)].copy()
        record_source(conn, target, "基本面Point-in-time", target.isoformat() if not hist.empty else None, bool(not hist.empty), complete,
                      f"逐公司法定時限：一般業需最新季 {len(required_period)} 檔，覆蓋 {len(required_period)-len(missing)}/{len(required_period)}；三率不適用/特殊業別 {len(exempt)} 檔" +
                      (f"；真正缺最新季 {len(missing)} 檔 ({','.join(missing[:20])})" if missing else "；當期最新季無實質缺漏"))
        return hist
    # 今日若資料庫已涵蓋每檔目前依法可得最新季，直接沿用，不為相同季度重抓。
    wanted=set(str(x) for x in (stock_ids or []))
    if wanted:
        hist=_latest_income_history_asof(conn,target,stock_ids)
        latest_by_sid={}
        if not hist.empty:
            for _,rr in hist.iterrows():
                latest_by_sid[str(rr["stock_id"])]=(int(rr["year"]),int(rr["quarter"]))
        required={}
        exempt=[]
        for sid in wanted:
            if _is_three_rate_na_stock(conn,sid,target):
                exempt.append(sid); continue
            latest=_available_quarters_for_stock(conn,sid,target,1)
            if latest: required[sid]=latest[-1]
        missing=[sid for sid,pq in required.items() if latest_by_sid.get(sid)!=pq]
        if not missing:
            if not hist.empty and required:
                hist=hist[hist.apply(lambda r: (str(r["stock_id"]) in exempt) or required.get(str(r["stock_id"]))==(int(r["year"]),int(r["quarter"])),axis=1)].copy()
            record_source(conn,target,"官方財報最新批次",target.isoformat(),True,True,
                          f"增量更新：當期最新季已存在，沿用 {len(hist)} 檔，不重抓相同季度")
            return hist
    rows_all, errors = _fetch_official_income_bulk_latest(stock_ids)
    if rows_all:
        # 每家公司只保留最新一期作為今天的官方快照。
        latest: Dict[str, dict] = {}
        for r in rows_all:
            sid = str(r["stock_id"])
            if sid not in latest or (int(r["year"]), int(r["quarter"])) > (int(latest[sid]["year"]), int(latest[sid]["quarter"])):
                latest[sid] = r
        latest_rows = list(latest.values())
        _save_income_snapshot(conn, target, latest_rows)
        periods = [int(r["year"]) * 10 + int(r["quarter"]) for r in latest_rows]
        label = f"{max(periods)//10}Q{max(periods)%10}" if periods else target.isoformat()
        record_source(conn, target, "官方財報最新批次", label, True, len(errors)==0,
                      f"TWSE/TPEx 六類綜合損益 OpenAPI；可解析 {len(latest_rows)} 檔" +
                      (f"；端點失敗 {len(errors)} 個：{' | '.join(errors[:4])}" if errors else "；全部端點完成"))
    else:
        record_source(conn, target, "官方財報最新批次", None, False, False,
                      "六類綜合損益 OpenAPI 均未取得可解析資料" + (f"；{' | '.join(errors[:4])}" if errors else ""))
    return _latest_income_snapshot(conn, target, stock_ids)



def _history_periods(conn: sqlite3.Connection, stock_id: str) -> pd.DataFrame:
    return pd.read_sql_query("""SELECT stock_id,year,quarter,revenue_ytd,gross_profit_ytd,operating_income_ytd,net_income_ytd,eps_ytd,
                                availability_date,availability_basis,source
                                FROM fundamental_income_history WHERE stock_id=? ORDER BY year,quarter""", conn, params=(stock_id,))


def _latest_income_history_asof(conn: sqlite3.Connection, target: dt.date, stock_ids: Optional[List[str]] = None) -> pd.DataFrame:
    """用已回補的歷史季報，挑出 target 當天『保守可得』的最新一期。"""
    params: List[Any] = [target.isoformat()]
    where = "availability_date IS NOT NULL AND availability_date<=?"
    if stock_ids:
        marks = ",".join("?" for _ in stock_ids)
        where += f" AND stock_id IN ({marks})"
        params.extend(stock_ids)
    hist = pd.read_sql_query(f"""SELECT * FROM fundamental_income_history WHERE {where}""", conn, params=params)
    if hist.empty:
        return hist
    hist["period_key"] = hist["year"].astype(int) * 10 + hist["quarter"].astype(int)
    hist = hist.sort_values(["stock_id", "period_key", "availability_date"]).groupby("stock_id", as_index=False).tail(1).copy()
    hist["observed_date"] = hist["availability_date"]
    hist["name"] = None
    hist["market"] = None
    hist["report_generated"] = hist["availability_date"]
    return hist[["observed_date","stock_id","name","market","year","quarter","report_generated",
                 "revenue_ytd","gross_profit_ytd","operating_income_ytd","net_income_ytd","eps_ytd","source","availability_basis"]]



# ============================
# v15：官方歷史基本面「強制補洞」
# ============================
# 月營收：使用 MOPS 官方歷史月營收批次頁（上市/上櫃、境內/境外），一次抓一個月份全市場。
# 季報：針對仍缺少的個股/季度，使用 MOPS 單一公司 IFRS 綜合損益表查詢補洞。
# 每次只抓 SQLite 缺少的月份/季度；中斷後下次可接續，不會重抓已完成資料。

MOPS_MONTHLY_HOSTS = ["https://mops.twse.com.tw", "https://mopsov.twse.com.tw"]
MOPS_INCOME_POST_HOSTS = ["https://mops.twse.com.tw", "https://mopsov.twse.com.tw"]


def _http_text(method: str, url: str, *, data: Optional[dict] = None, timeout: int = 15, tries: int = 2) -> str:
    last: Optional[Exception] = None
    for i in range(tries):
        try:
            if method.upper() == "POST":
                r = SESSION.post(url, data=data, timeout=timeout, headers={"User-Agent": UA, "Referer": url.rsplit('/',1)[0] + '/', "Connection":"close"})
            else:
                r = SESSION.get(url, timeout=timeout, headers={"User-Agent": UA, "Connection":"close"})
            r.raise_for_status()
            raw = r.content
            # MOPS 舊歷史頁以 Big5/CP950 為主，新頁多為 UTF-8。
            for enc in (r.encoding, "utf-8-sig", "utf-8", "cp950", "big5"):
                if not enc:
                    continue
                try:
                    return raw.decode(enc)
                except Exception:
                    pass
            return raw.decode("utf-8", errors="replace")
        except Exception as e:
            last = e
            if i < tries - 1:
                time.sleep(1.0 + i)
    raise RuntimeError(f"HTTP {method} 失敗: {url} | {last}")


def _flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    out=df.copy()
    cols=[]
    for c in out.columns:
        if isinstance(c, tuple):
            parts=[str(x).strip() for x in c if str(x).strip() and not str(x).startswith("Unnamed")]
            cols.append(parts[-1] if parts else str(c[-1]))
        else:
            cols.append(str(c).strip())
    out.columns=cols
    return out


def _df_col(df: pd.DataFrame, aliases: Iterable[str]) -> Optional[str]:
    norm={_norm_key(c):c for c in df.columns}
    for a in aliases:
        na=_norm_key(a)
        if na in norm:
            return norm[na]
    for a in aliases:
        na=_norm_key(a)
        for k,c in norm.items():
            if na and (na in k or k in na):
                return c
    return None


def _available_month_keys(target: dt.date, n: int) -> List[int]:
    vals=[]
    y,m=target.year,target.month
    # 從目標月往回找；只有依法最晚申報日 <= target 才納入。
    for _ in range(n+6):
        # 上一個月
        m-=1
        if m==0:
            y-=1; m=12
        if _month_revenue_conservative_available(y,m) <= target:
            vals.append(y*100+m)
            if len(vals)>=n:
                break
    return sorted(vals)


def _available_quarters(target: dt.date, n: int) -> List[Tuple[int,int]]:
    vals=[]
    # 從 target 所在季度往前逐季判斷保守可得日。
    q=(target.month-1)//3+1; y=target.year
    for _ in range(n+8):
        if _income_conservative_available(y,q) <= target:
            vals.append((y,q))
            if len(vals)>=n:
                break
        q-=1
        if q==0:
            y-=1; q=4
    return sorted(vals)


def _three_rate_required_periods(target: dt.date) -> List[Tuple[int,int]]:
    """回傳計算 target 當期「本季 vs 上季」三率真正需要的累季報表期別。

    Q2: Q1,Q2
    Q3: Q1,Q2,Q3
    Q4: Q2,Q3,Q4
    Q1: 前年Q3,Q4 + 當年Q1
    因此每日核心判斷不需要先補滿12季。
    """
    latest=_available_quarters(target,1)
    if not latest:
        return []
    cy,cq=latest[-1]
    py,pq=(cy,cq-1) if cq>1 else (cy-1,4)
    req={(cy,cq),(py,pq)}
    if cq>1:
        req.add((cy,cq-1))
    if pq>1:
        req.add((py,pq-1))
    return sorted(req)


def _parse_mops_monthly_table(html: str, month_key: int, market: str, source: str) -> List[dict]:
    rows=[]
    try:
        tables=pd.read_html(io.StringIO(html))
    except Exception:
        return rows
    for raw in tables:
        df=_flatten_columns(raw)
        c_code=_df_col(df,["公司代號","公司代碼"])
        c_name=_df_col(df,["公司名稱"])
        if not c_code or not c_name:
            continue
        c_rev=_df_col(df,["當月營收","本月營收"])
        c_prev=_df_col(df,["上月營收"])
        c_y1=_df_col(df,["去年當月營收","去年本月營收"])
        c_cum=_df_col(df,["當月累計營收","本年累計營收","累計營收"])
        c_cum1=_df_col(df,["去年累計營收","去年同期累計營收"])
        c_mom=_df_col(df,["上月比較增減(%)","上月比較增減％","月增率"])
        c_yoy=_df_col(df,["去年同月增減(%)","去年同月增減％","年增率"])
        c_cyoy=_df_col(df,["前期比較增減(%)","前期比較增減％","累計營收增","累計年增率"])
        for _,rr in df.iterrows():
            code=str(rr.get(c_code,"")).strip().replace(".0","")
            name=str(rr.get(c_name,"")).strip()
            if not is_common_stock(code,name):
                continue
            rev=clean_num(rr.get(c_rev)) if c_rev else None
            prev=clean_num(rr.get(c_prev)) if c_prev else None
            y1=clean_num(rr.get(c_y1)) if c_y1 else None
            cum=clean_num(rr.get(c_cum)) if c_cum else None
            cum1=clean_num(rr.get(c_cum1)) if c_cum1 else None
            mom=clean_num(rr.get(c_mom)) if c_mom else None
            yoy=clean_num(rr.get(c_yoy)) if c_yoy else None
            cyoy=clean_num(rr.get(c_cyoy)) if c_cyoy else None
            if mom is None and rev is not None and prev not in (None,0): mom=(rev/prev-1)*100
            if yoy is None and rev is not None and y1 not in (None,0): yoy=(rev/y1-1)*100
            if cyoy is None and cum is not None and cum1 not in (None,0): cyoy=(cum/cum1-1)*100
            # v30：成長率欄空白不代表月營收不存在（例如前期為0或特殊申報）。
            # 先保存原始營收，後續可用歷史月營收自行重算 MoM / YoY / 累計YoY。
            if rev is None and mom is None and yoy is None and cyoy is None:
                continue
            rows.append({"stock_id":code,"name":name,"market":market,"month_key":month_key,"revenue":rev,
                         "revenue_mom_pct":mom,"revenue_yoy_pct":yoy,"revenue_cum_yoy_pct":cyoy,"source":source})
    # 同一公司可能因多張表重複，最後一筆即可。
    return list({r["stock_id"]:r for r in rows}.values())


def _fetch_mops_month_revenue_history(month_key: int) -> Tuple[List[dict], List[str]]:
    y,m=month_key//100,month_key%100; roc=y-1911
    all_rows=[]; errors=[]
    for seg,market in (("sii","TWSE"),("otc","TPEx")):
        for company_type in (0,1):
            ok=False
            for host in MOPS_MONTHLY_HOSTS:
                url=f"{host}/nas/t21/{seg}/t21sc03_{roc}_{m}_{company_type}.html"
                try:
                    html=_http_text("GET",url,timeout=15,tries=2)
                    parsed=_parse_mops_monthly_table(html,month_key,market,f"MOPS歷史月營收[{seg}/{company_type}]")
                    all_rows.extend(parsed); ok=True; break
                except Exception as e:
                    last=e
            if not ok:
                errors.append(f"{seg}/{month_key}/{company_type}: {last}")
    return list({(r["stock_id"],r["month_key"]):r for r in all_rows}.values()), errors


def _upsert_revenue_history_rows(conn: sqlite3.Connection, rows: List[dict]) -> int:
    if not rows: return 0
    now=dt.datetime.now().isoformat(timespec="seconds"); vals=[]
    for r in rows:
        mk=int(r["month_key"]); y,m=mk//100,mk%100
        sid=str(r["stock_id"]); name=str(r.get("name") or "")
        avail=_month_revenue_available_for_identity(sid,name,y,m)
        vals.append((sid,mk,r.get("revenue"),r.get("revenue_mom_pct"),r.get("revenue_yoy_pct"),r.get("revenue_cum_yoy_pct"),
                     avail.isoformat(),"mops_historical_per_stock_deadline",r.get("source"),now))
    conn.executemany("""INSERT INTO fundamental_revenue_history
        (stock_id,month_key,revenue,revenue_mom_pct,revenue_yoy_pct,revenue_cum_yoy_pct,availability_date,availability_basis,source,updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?) ON CONFLICT(stock_id,month_key) DO UPDATE SET
        revenue=COALESCE(excluded.revenue,fundamental_revenue_history.revenue),
        revenue_mom_pct=COALESCE(excluded.revenue_mom_pct,fundamental_revenue_history.revenue_mom_pct),
        revenue_yoy_pct=COALESCE(excluded.revenue_yoy_pct,fundamental_revenue_history.revenue_yoy_pct),
        revenue_cum_yoy_pct=COALESCE(excluded.revenue_cum_yoy_pct,fundamental_revenue_history.revenue_cum_yoy_pct),
        availability_date=CASE WHEN fundamental_revenue_history.availability_date IS NULL OR excluded.availability_date < fundamental_revenue_history.availability_date THEN excluded.availability_date ELSE fundamental_revenue_history.availability_date END,
        availability_basis=CASE WHEN fundamental_revenue_history.availability_date IS NULL OR excluded.availability_date <= fundamental_revenue_history.availability_date THEN excluded.availability_basis ELSE fundamental_revenue_history.availability_basis END,
        source=COALESCE(excluded.source,fundamental_revenue_history.source),updated_at=excluded.updated_at""",vals)
    conn.commit(); return len(vals)


def _extract_statement_value(df: pd.DataFrame, labels: Iterable[str]) -> Optional[float]:
    d=_flatten_columns(df)
    # 逐列找科目名稱，取該列第一個可解析的數值欄；MOPS 表格通常先列本期累計。
    for _,row in d.iterrows():
        vals=list(row.values)
        if not vals: continue
        label=" ".join(str(x) for x in vals[:2] if pd.notna(x))
        nlabel=_norm_key(label)
        if not any(_norm_key(x) in nlabel for x in labels):
            continue
        for v in vals[1:]:
            n=clean_num(v)
            if n is not None:
                return n
    return None


def _fetch_mops_income_history_one(stock_id: str, year: int, quarter: int) -> Optional[dict]:
    payload={"encodeURIComponent":"1","step":"1","firstin":"1","off":"1","queryName":"co_id","inpuType":"co_id",
             "TYPEK":"all","isnew":"false","co_id":str(stock_id),"year":str(year-1911),"season":f"{quarter:02d}"}
    last=None
    for host in MOPS_INCOME_POST_HOSTS:
        url=f"{host}/mops/web/ajax_t164sb04"
        try:
            html=_http_text("POST",url,data=payload,timeout=18,tries=2)
            tables=pd.read_html(io.StringIO(html))
            best=None
            for df in tables:
                revenue=_extract_statement_value(df,["營業收入合計","營業收入","收益合計","淨收益"])
                gross=_extract_statement_value(df,["營業毛利（毛損）淨額","營業毛利毛損淨額","營業毛利"])
                op=_extract_statement_value(df,["營業利益（損失）","營業利益損失","營業利益","繼續營業單位稅前淨利","稅前淨利"])
                net=_extract_statement_value(df,["本期淨利（淨損）","本期淨利淨損","本期淨利","本期稅後淨利","淨利淨損"])
                eps=_extract_statement_value(df,["基本每股盈餘（元）","基本每股盈餘元","基本每股盈餘"])
                # 金融/保險業未必有「毛利」科目；仍保存季度，三率策略則標示資料不適用。
                if None not in (revenue,op,net):
                    best={"stock_id":str(stock_id),"year":year,"quarter":quarter,"revenue_ytd":revenue,"gross_profit_ytd":gross,
                          "operating_income_ytd":op,"net_income_ytd":net,"eps_ytd":eps,"source":"MOPS歷史綜合損益[t164sb04]"}
                    break
            if best is not None:
                return best
            compact=re.sub(r"\s+","",html)
            if any(x in compact for x in ["查無資料","無符合條件","沒有符合條件","無資料"]):
                return None
            # v24：不能把安全攔截/版型改變/解析失敗誤寫成 NODATA，否則之後永遠不會重抓。
            raise RuntimeError("MOPS 回傳頁面未解析到有效綜合損益表（不寫入 NODATA，保留下次重試）")
        except Exception as e:
            last=e
    if last:
        raise RuntimeError(str(last))
    return None


def _upsert_income_history_record(conn: sqlite3.Connection, r: dict) -> None:
    sid=str(r["stock_id"]); name=str(r.get("name") or _stock_name_asof(conn,sid,None))
    avail=_income_available_for_identity(sid,name,int(r["year"]),int(r["quarter"]),str(r.get("source") or "")); now=dt.datetime.now().isoformat(timespec="seconds")
    conn.execute("""INSERT INTO fundamental_income_history
        (stock_id,year,quarter,revenue_ytd,gross_profit_ytd,operating_income_ytd,net_income_ytd,eps_ytd,availability_date,availability_basis,source,updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(stock_id,year,quarter) DO UPDATE SET
        revenue_ytd=COALESCE(excluded.revenue_ytd,fundamental_income_history.revenue_ytd),
        gross_profit_ytd=COALESCE(excluded.gross_profit_ytd,fundamental_income_history.gross_profit_ytd),
        operating_income_ytd=COALESCE(excluded.operating_income_ytd,fundamental_income_history.operating_income_ytd),
        net_income_ytd=COALESCE(excluded.net_income_ytd,fundamental_income_history.net_income_ytd),
        eps_ytd=COALESCE(excluded.eps_ytd,fundamental_income_history.eps_ytd),
        availability_date=CASE WHEN fundamental_income_history.availability_date IS NULL OR excluded.availability_date < fundamental_income_history.availability_date THEN excluded.availability_date ELSE fundamental_income_history.availability_date END,
        availability_basis=CASE WHEN fundamental_income_history.availability_date IS NULL OR excluded.availability_date <= fundamental_income_history.availability_date THEN excluded.availability_basis ELSE fundamental_income_history.availability_basis END,
        source=COALESCE(excluded.source,fundamental_income_history.source),updated_at=excluded.updated_at""",
        (r["stock_id"],r["year"],r["quarter"],r.get("revenue_ytd"),r.get("gross_profit_ytd"),r.get("operating_income_ytd"),r.get("net_income_ytd"),r.get("eps_ytd"),
         avail.isoformat(),"mops_historical_conservative_deadline",r.get("source"),now))
    conn.commit()

def _first_trade_date(conn: sqlite3.Connection, stock_id: str, target: dt.date) -> Optional[dt.date]:
    row=conn.execute("SELECT MIN(date) FROM prices WHERE stock_id=? AND date<=?",(str(stock_id),target.isoformat())).fetchone()
    if not row or not row[0]:
        return None
    try:
        return dt.date.fromisoformat(str(row[0])[:10])
    except Exception:
        return None

def _quarter_end_date(year:int, quarter:int)->dt.date:
    m=quarter*3
    return dt.date(year,m,calendar.monthrange(year,m)[1])

def _eligible_quarters_for_stock(conn: sqlite3.Connection, stock_id:str, target:dt.date, expected:List[Tuple[int,int]])->List[Tuple[int,int]]:
    first=_first_trade_date(conn,stock_id,target)
    if first is None:
        return list(expected)
    # 上市/上櫃前的季度不列為「缺漏」。季度結束日在首個交易日前者免計。
    return [(y,q) for y,q in expected if _quarter_end_date(y,q) >= first]

def _month_end_date(month_key:int)->dt.date:
    y,m=month_key//100,month_key%100
    return dt.date(y,m,calendar.monthrange(y,m)[1])

def _eligible_months_for_stock(conn: sqlite3.Connection, stock_id:str, target:dt.date, expected:List[int])->List[int]:
    first=_first_trade_date(conn,stock_id,target)
    if first is None:
        return list(expected)
    # 首個交易日前已結束的月份不列為缺漏。
    return [mk for mk in expected if _month_end_date(mk) >= first]

def _backfill_income_history(conn: sqlite3.Connection, target: dt.date, stock_ids: List[str], cfg: configparser.ConfigParser) -> None:
    """v16：補最近12個可得季度；官方確實查無的歷史期記錄成 NODATA，避免每天重抓。
    「完整」代表補抓流程無連線/解析失敗，不強迫成立未滿12季的公司憑空擁有12季資料。
    """
    if not stock_ids:
        return
    min_quarters=max(4,cfg.getint("general","fundamental_backfill_quarters",fallback=12))
    attempted=0; added=0; failures=[]; nodata_new=[]; nodata_cached=0
    sid_list=[str(x) for x in stock_ids]
    expected_by_sid={sid:_available_quarters_for_stock(conn,sid,target,min_quarters) for sid in sid_list}
    print(f"[財報歷史] 檢查 {len(sid_list)} 檔、每檔最多 {min_quarters} 季...", flush=True)
    for sidx,sid in enumerate(sid_list,1):
        have={(int(y),int(q)) for y,q in conn.execute("SELECT year,quarter FROM fundamental_income_history WHERE stock_id=?",(sid,)).fetchall()}
        eligible=_eligible_quarters_for_stock(conn,sid,target,expected_by_sid.get(sid,[]))
        missing=[pq for pq in eligible if pq not in have]
        uncached=[]
        for y,q in missing:
            period=f"{y}Q{q}"
            cached=conn.execute("SELECT 1 FROM fundamental_backfill_nodata WHERE dataset='income' AND stock_id=? AND period=?",(sid,period)).fetchone()
            if not cached:
                uncached.append((y,q))
        if sidx==1 or sidx%10==0 or uncached or sidx==len(sid_list):
            print(f"[財報歷史 {sidx}/{len(sid_list)}] {sid}：已有 {len(have)} 季；待抓 {len(uncached)} 季；查無快取 {len(missing)-len(uncached)} 季", flush=True)
        for y,q in missing:
            period=f"{y}Q{q}"
            cached=conn.execute("SELECT 1 FROM fundamental_backfill_nodata WHERE dataset='income' AND stock_id=? AND period=?",(sid,period)).fetchone()
            if cached:
                nodata_cached+=1
                continue
            attempted+=1
            print(f"    抓取 {sid} {period} ...", end="", flush=True)
            try:
                r=_fetch_mops_income_history_one(sid,y,q)
                if r:
                    _upsert_income_history_record(conn,r); added+=1
                    print(" OK", flush=True)
                else:
                    nodata_new.append(f"{sid}:{period}")
                    conn.execute("INSERT OR REPLACE INTO fundamental_backfill_nodata(dataset,stock_id,period,checked_at) VALUES('income',?,?,?)",(sid,period,dt.datetime.now().isoformat(timespec='seconds')))
                    conn.commit()
                    print(" 官方查無", flush=True)
            except Exception as e:
                failures.append(f"{sid}:{period} {str(e)[:80]}")
                print(f" 失敗：{str(e)[:120]}", flush=True)
            time.sleep(0.35)
    coverage=[]
    for sid in stock_ids:
        eligible=_eligible_quarters_for_stock(conn,str(sid),target,expected)
        if eligible:
            pairs=set(eligible)
            rows=conn.execute("SELECT year,quarter FROM fundamental_income_history WHERE stock_id=? AND availability_date<=?",(str(sid),target.isoformat())).fetchall()
            n=sum((int(y),int(q)) in pairs for y,q in rows)
        else:
            n=0
        required=len(eligible)
        coverage.append((str(sid),n,required))
    usable=sum((req==0) or n>=min(4,req) for _,n,req in coverage)
    full=sum((req==0) or n>=req for _,n,req in coverage)
    short=[sid for sid,n,req in coverage if req>0 and n<req]
    # 完整 = 所有應嘗試的官方請求皆完成；官方本身查無歷史期不算抓取失敗。
    complete=(len(failures)==0)
    msg=(f"v16 官方補洞：目標最多 {min_quarters} 季；本次實際請求 {attempted}、新增 {added}；"
         f"{len(stock_ids)} 檔中 可供判斷 {usable}、官方可得季度已補齊 {full}；仍有可得季度缺口 {len(short)} 檔"
         + (f" ({','.join(short[:20])})" if short else "")
         + (f"；本次新增查無 {len(nodata_new)} 筆、既有查無快取 {nodata_cached} 筆" if (nodata_new or nodata_cached) else "")
         + (f"；真正連線/解析失敗 {len(failures)} 筆" if failures else "；補抓流程完成"))
    # 成功代表本次補洞至少有可判定結果；真正連線/解析失敗才影響完整度。
    fetch_success = (attempted == 0) or (added > 0) or (len(nodata_new) > 0) or (nodata_cached > 0)
    record_source(conn,target,"官方財報歷史",target.isoformat(),fetch_success,complete,msg)


def _single_quarter_values(period_map: Dict[Tuple[int,int], dict], year: int, quarter: int) -> Optional[dict]:
    cur = period_map.get((year, quarter))
    if not cur:
        return None
    core=["revenue_ytd","gross_profit_ytd","operating_income_ytd","net_income_ytd"]
    if any(cur.get(k) is None for k in core):
        return None
    out: Dict[str, Optional[float]] = {}
    if quarter == 1:
        for k in core:
            out[k]=float(cur[k])
        out["eps_ytd"] = float(cur["eps_ytd"]) if cur.get("eps_ytd") is not None else None
        return out
    prev = period_map.get((year, quarter-1))
    if not prev or any(prev.get(k) is None for k in core):
        return None
    for k in core:
        out[k]=float(cur[k])-float(prev[k])
    out["eps_ytd"] = (float(cur["eps_ytd"])-float(prev["eps_ytd"])) if cur.get("eps_ytd") is not None and prev.get("eps_ytd") is not None else None
    return out


def _prev_period(year: int, quarter: int) -> Tuple[int,int]:
    return (year, quarter-1) if quarter > 1 else (year-1, 4)


def _three_rates_for_stock(conn: sqlite3.Connection, snap: pd.Series, allow_equal: bool) -> Optional[dict]:
    sid=str(snap["stock_id"]); cy,cq=int(snap["year"]),int(snap["quarter"])
    hist=_history_periods(conn,sid)
    period_map: Dict[Tuple[int,int],dict]={}
    for r in hist.to_dict("records"):
        period_map[(int(r["year"]),int(r["quarter"]))]=r
    # 以 point-in-time 官方 snapshot 覆蓋目前期。
    period_map[(cy,cq)]={k:snap.get(k) for k in ["revenue_ytd","gross_profit_ytd","operating_income_ytd","net_income_ytd","eps_ytd"]}
    py,pq=_prev_period(cy,cq)
    cur=_single_quarter_values(period_map,cy,cq); prev=_single_quarter_values(period_map,py,pq)
    if not cur or not prev or cur["revenue_ytd"] == 0 or prev["revenue_ytd"] == 0:
        return None
    cg=cur["gross_profit_ytd"]/cur["revenue_ytd"]; pg=prev["gross_profit_ytd"]/prev["revenue_ytd"]
    co=cur["operating_income_ytd"]/cur["revenue_ytd"]; po=prev["operating_income_ytd"]/prev["revenue_ytd"]
    cn=cur["net_income_ytd"]/cur["revenue_ytd"]; pn=prev["net_income_ytd"]/prev["revenue_ytd"]
    gd,od,nd=cg-pg,co-po,cn-pn
    cmp=(lambda x:x>=-1e-12) if allow_equal else (lambda x:x>1e-12)
    return {
        "毛利率本季%":cg*100,"毛利率上季%":pg*100,"營業利益率本季%":co*100,"營業利益率上季%":po*100,
        "稅後純益率本季%":cn*100,"稅後純益率上季%":pn*100,
        "毛差":gd,"營差":od,"淨差":nd,"三率三升":bool(cmp(gd) and cmp(od) and cmp(nd)),
        "三率資料期":f"{cy}Q{cq} vs {py}Q{pq}","三率快照日":str(snap["observed_date"]),
        "三率可得日依據":str(snap.get("availability_basis") or "official_snapshot"),
        "三率來源":str(snap.get("source") or "官方綜合損益")+"＋SQLite歷史季報",
    }


def _save_revenue_snapshot(conn: sqlite3.Connection, observed: dt.date, rows: List[dict]) -> None:
    if not rows: return
    vals=[(observed.isoformat(),r["stock_id"],r.get("name"),r.get("market"),int(r["month_key"]),r["revenue_mom_pct"],r["revenue_yoy_pct"],r["revenue_cum_yoy_pct"],r.get("source")) for r in rows]
    conn.executemany("""INSERT OR REPLACE INTO fundamental_revenue_snapshot
        (observed_date,stock_id,name,market,month_key,revenue_mom_pct,revenue_yoy_pct,revenue_cum_yoy_pct,source)
        VALUES (?,?,?,?,?,?,?,?,?)""", vals)
    # live 官方月營收同時更新 history，實際可見日就是 observed。
    now = dt.datetime.now().isoformat(timespec="seconds")
    hist=[]
    for r in rows:
        hist.append((r["stock_id"],int(r["month_key"]),r.get("revenue"),r["revenue_mom_pct"],r["revenue_yoy_pct"],r["revenue_cum_yoy_pct"],
                     observed.isoformat(),"official_live_snapshot",r.get("source"),now))
    conn.executemany("""INSERT INTO fundamental_revenue_history
        (stock_id,month_key,revenue,revenue_mom_pct,revenue_yoy_pct,revenue_cum_yoy_pct,availability_date,availability_basis,source,updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?) ON CONFLICT(stock_id,month_key) DO UPDATE SET
        revenue=COALESCE(excluded.revenue,fundamental_revenue_history.revenue),
        revenue_mom_pct=excluded.revenue_mom_pct,revenue_yoy_pct=excluded.revenue_yoy_pct,revenue_cum_yoy_pct=excluded.revenue_cum_yoy_pct,
        availability_date=CASE
            WHEN fundamental_revenue_history.availability_date IS NULL THEN excluded.availability_date
            WHEN excluded.availability_date < fundamental_revenue_history.availability_date THEN excluded.availability_date
            ELSE fundamental_revenue_history.availability_date END,
        availability_basis=CASE WHEN excluded.availability_date <= COALESCE(fundamental_revenue_history.availability_date,excluded.availability_date)
                                THEN excluded.availability_basis ELSE fundamental_revenue_history.availability_basis END,
        source=excluded.source,updated_at=excluded.updated_at""", hist)
    conn.commit()


def _latest_revenue_snapshot(conn: sqlite3.Connection, target: dt.date) -> pd.DataFrame:
    return pd.read_sql_query("""SELECT s.* FROM fundamental_revenue_snapshot s JOIN (
        SELECT stock_id,MAX(observed_date) obs FROM fundamental_revenue_snapshot WHERE observed_date<=? GROUP BY stock_id
        ) x ON s.stock_id=x.stock_id AND s.observed_date=x.obs WHERE s.observed_date<=?""", conn, params=(target.isoformat(),target.isoformat()))


def _latest_revenue_history_asof(conn: sqlite3.Connection, target: dt.date, stock_ids: Optional[List[str]]=None) -> pd.DataFrame:
    params: List[Any] = [target.isoformat()]
    where = "availability_date IS NOT NULL AND availability_date<=?"
    if stock_ids:
        marks=",".join("?" for _ in stock_ids)
        where += f" AND stock_id IN ({marks})"
        params.extend(stock_ids)
    hist=pd.read_sql_query(f"SELECT * FROM fundamental_revenue_history WHERE {where}",conn,params=params)
    if hist.empty: return hist
    hist=hist.sort_values(["stock_id","month_key","availability_date"]).groupby("stock_id",as_index=False).tail(1).copy()
    hist["observed_date"]=hist["availability_date"]
    hist["name"]=None; hist["market"]=None
    return hist[["observed_date","stock_id","name","market","month_key","revenue_mom_pct","revenue_yoy_pct","revenue_cum_yoy_pct","source","availability_basis"]]



def _backfill_revenue_history(conn: sqlite3.Connection,target:dt.date,stock_ids:List[str],cfg:configparser.ConfigParser)->None:
    """v16：批次下載最近24個可得月份的官方全市場月營收。
    「完整」代表24個月份的官方批次頁皆成功取得；個別公司未滿24個月只列為期數不足，不把來源判成失敗。
    """
    if not stock_ids:
        return
    min_months=max(13,cfg.getint("general","revenue_backfill_months",fallback=24))
    wanted=set(str(x) for x in stock_ids)
    # v30：月份可得日改成逐公司判斷。保險/具保險子公司公司在2026起用次月15日保守期限。
    expected_by_sid={sid:_available_month_keys_for_stock(conn,sid,target,min_months) for sid in wanted}
    eligible_by_sid={sid:set(_eligible_months_for_stock(conn,sid,target,expected_by_sid[sid])) for sid in wanted}
    expected=sorted(set().union(*(set(v) for v in expected_by_sid.values()))) if expected_by_sid else []
    need_months=[]
    for mk in expected:
        required={sid for sid in wanted if mk in eligible_by_sid.get(sid,set())}
        if not required:
            continue
        have={str(x[0]) for x in conn.execute("SELECT stock_id FROM fundamental_revenue_history WHERE month_key=?",(mk,)).fetchall()}
        if not required.issubset(have):
            need_months.append(mk)
    downloaded=0; inserted=0; errors=[]
    total_need=len(need_months)
    if total_need:
        print(f"[營收歷史] 需要補 {total_need} 個月份，開始處理...", flush=True)
    else:
        print(f"[營收歷史] 最近 {min_months} 個月已無需補抓。", flush=True)
    for idx,mk in enumerate(need_months,1):
        print(f"[營收歷史 {idx}/{total_need}] {mk//100}-{mk%100:02d} ...", flush=True)
        try:
            rows,errs=_fetch_mops_month_revenue_history(mk)
            nwrite=_upsert_revenue_history_rows(conn,rows)
            inserted+=nwrite; downloaded+=1
            errors.extend(errs)
            print(f"    完成：解析 {len(rows)} 檔，寫入/更新 {nwrite} 筆" + (f"；子來源警告 {len(errs)}" if errs else ""), flush=True)
        except Exception as e:
            errors.append(f"{mk}: {e}")
            print(f"    [失敗] {e}", flush=True)
        time.sleep(0.12)
    coverage=[]
    for sid in stock_ids:
        eligible=_eligible_months_for_stock(conn,str(sid),target,expected_by_sid.get(str(sid),[]))
        if eligible:
            marks=','.join('?' for _ in eligible)
            n=int(conn.execute(f"SELECT COUNT(*) FROM fundamental_revenue_history WHERE stock_id=? AND month_key IN ({marks})",(str(sid),*eligible)).fetchone()[0])
        else:
            n=0
        required=len(eligible)
        coverage.append((str(sid),n,required))
    usable=sum((req==0) or n>=min(13,req) for _,n,req in coverage)
    full=sum((req==0) or n>=req for _,n,req in coverage)
    short=[sid for sid,n,req in coverage if req>0 and n<req]
    # 下載/解析24個月份都成功即視為來源完整；公司自身沒有24個月資料另外標示。
    complete=(len(errors)==0)
    msg=(f"v30 深度歷史（研究/回測用途）：目標最多 {min_months} 月；本次需下載 {len(need_months)} 個月份、成功 {downloaded}、寫入/更新 {inserted} 筆；"
         f"依各公司上市日/特殊申報期限計算：可供判斷 {usable}/{len(stock_ids)}、深度期數已補齊 {full}/{len(stock_ids)}；仍有深度歷史缺口 {len(short)} 檔"
         + (f" ({','.join(short[:20])})" if short else "")
         + (f"；真正歷史頁連線/解析錯誤 {len(errors)} 筆" if errors else "；深度歷史批次流程完成")
         + "；此列不代表當日營收缺漏，當日可用性請看『營收驗收』")
    # 成功代表官方批次抓取流程有跑通；個股期數不足只放在訊息中，不再把「成功」誤判為否。
    fetch_success = (total_need == 0) or (downloaded > 0)
    record_source(conn,target,"官方月營收歷史",target.isoformat(),fetch_success,complete,msg)


def _fetch_official_revenue_rows(source: str, url: str, market: str) -> List[dict]:
    js = request_json(url, timeout=20, tries=2)
    if not isinstance(js, list):
        raise RuntimeError(f"官方回傳格式非 list: {type(js).__name__}")
    rows: List[dict] = []
    for r in js:
        code = str(_row_value(r,["公司代號","股票代號","證券代號","Code"]) or "").strip()
        name = str(_row_value(r,["公司名稱","名稱","Name"]) or "").strip()
        if not is_common_stock(code, name):
            continue
        mk = _parse_month_key(_row_value(r,["資料年月","年月","RevenueMonth","Date"]))
        mom = clean_num(_row_value(r,["營業收入-上月比較增減(%)","上月比較增減","月增率","MoM"]))
        yoy = clean_num(_row_value(r,["營業收入-去年同月增減(%)","去年同月增減","年增率","YoY"]))
        cum = clean_num(_row_value(r,["累計營業收入-前期比較增減(%)","累計營收-前期比較增減","累計年增率","累計營收增","CumYoY"]))
        rev = clean_num(_row_value(r,["當月營收","營業收入-當月營收","營收","Revenue"]))
        if mk is None:
            continue
        if rev is None and mom is None and yoy is None and cum is None:
            continue
        rows.append({"stock_id":code,"name":name,"market":market,"month_key":mk,"revenue":rev,
                     "revenue_mom_pct":mom,"revenue_yoy_pct":yoy,"revenue_cum_yoy_pct":cum,"source":source})
    latest: Dict[str,dict] = {}
    for r in rows:
        if r["stock_id"] not in latest or r["month_key"] > latest[r["stock_id"]]["month_key"]:
            latest[r["stock_id"]] = r
    return list(latest.values())


def _store_official_revenue_as_history_if_safe(conn: sqlite3.Connection, target: dt.date, rows: List[dict],
                                                stock_ids: Optional[List[str]]=None) -> int:
    """歷史補跑時，官方『目前最新月營收』只有在保守可得日 <= target 才可回填，避免未來穿越。"""
    wanted = set(str(x) for x in stock_ids) if stock_ids else None
    vals = []; now = dt.datetime.now().isoformat(timespec="seconds")
    for r in rows:
        sid = str(r["stock_id"])
        if wanted is not None and sid not in wanted:
            continue
        mk = int(r["month_key"]); y, m = mk//100, mk%100
        avail = _month_revenue_available_for_identity(sid, str(r.get("name") or _stock_name_asof(conn,sid,target)), y, m)
        if avail > target:
            continue
        vals.append((sid, mk, r.get("revenue"), r["revenue_mom_pct"], r["revenue_yoy_pct"], r["revenue_cum_yoy_pct"],
                     avail.isoformat(), "official_latest_safe_asof_per_stock_deadline", r.get("source") or "official-month-revenue", now))
    if vals:
        conn.executemany("""INSERT INTO fundamental_revenue_history
            (stock_id,month_key,revenue,revenue_mom_pct,revenue_yoy_pct,revenue_cum_yoy_pct,availability_date,availability_basis,source,updated_at)
            VALUES (?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(stock_id,month_key) DO UPDATE SET
            revenue=COALESCE(excluded.revenue,fundamental_revenue_history.revenue),
            revenue_mom_pct=excluded.revenue_mom_pct,revenue_yoy_pct=excluded.revenue_yoy_pct,
            revenue_cum_yoy_pct=excluded.revenue_cum_yoy_pct,
            availability_date=CASE WHEN fundamental_revenue_history.availability_date IS NULL OR excluded.availability_date < fundamental_revenue_history.availability_date
                                   THEN excluded.availability_date ELSE fundamental_revenue_history.availability_date END,
            availability_basis=CASE WHEN fundamental_revenue_history.availability_date IS NULL OR excluded.availability_date <= fundamental_revenue_history.availability_date
                                    THEN excluded.availability_basis ELSE fundamental_revenue_history.availability_basis END,
            source=excluded.source,updated_at=excluded.updated_at""", vals)
        conn.commit()
    return len(vals)



def _prev_month_key(month_key:int)->int:
    y,m=month_key//100,month_key%100
    m-=1
    if m==0:
        y-=1; m=12
    return y*100+m


def _recompute_revenue_growth_from_history(conn: sqlite3.Connection, target: dt.date,
                                           stock_ids: Optional[List[str]]=None) -> Tuple[int,int]:
    """v30：逐公司取 target 日依法已可得的最新月份，再用原始月營收補算 MoM/YoY/累計YoY。"""
    wanted=[str(x) for x in (stock_ids or [])]
    if not wanted:
        return 0,0
    updated=0; not_applicable=0
    now=dt.datetime.now().isoformat(timespec="seconds")
    for sid in wanted:
        latest_list=_available_month_keys_for_stock(conn,sid,target,1)
        if not latest_list:
            continue
        latest_mk=latest_list[-1]
        row=conn.execute("""SELECT revenue,revenue_mom_pct,revenue_yoy_pct,revenue_cum_yoy_pct
                            FROM fundamental_revenue_history
                            WHERE stock_id=? AND month_key=? AND availability_date<=?
                            ORDER BY availability_date DESC LIMIT 1""",(sid,latest_mk,target.isoformat())).fetchone()
        if not row:
            continue
        rev=clean_num(row[0]); mom=clean_num(row[1]); yoy=clean_num(row[2]); cyoy=clean_num(row[3])
        if rev is None:
            continue
        pmk=_prev_month_key(latest_mk); y,m=latest_mk//100,latest_mk%100; pyear_mk=(y-1)*100+m
        def get_rev(mk):
            rr=conn.execute("""SELECT revenue FROM fundamental_revenue_history
                               WHERE stock_id=? AND month_key=? AND availability_date<=?
                               ORDER BY availability_date DESC LIMIT 1""",(sid,mk,target.isoformat())).fetchone()
            return clean_num(rr[0]) if rr else None
        if mom is None:
            prev=get_rev(pmk)
            if prev not in (None,0): mom=(rev/prev-1.0)*100.0
        if yoy is None:
            prevy=get_rev(pyear_mk)
            if prevy not in (None,0): yoy=(rev/prevy-1.0)*100.0
        if cyoy is None:
            cur_vals=[get_rev(y*100+mm) for mm in range(1,m+1)]
            prev_vals=[get_rev((y-1)*100+mm) for mm in range(1,m+1)]
            if all(v is not None for v in cur_vals) and all(v is not None for v in prev_vals):
                ps=sum(prev_vals)
                if ps!=0: cyoy=(sum(cur_vals)/ps-1.0)*100.0
        before=(clean_num(row[1]),clean_num(row[2]),clean_num(row[3])); after=(mom,yoy,cyoy)
        if after!=before:
            conn.execute("""UPDATE fundamental_revenue_history SET revenue_mom_pct=?,revenue_yoy_pct=?,revenue_cum_yoy_pct=?,updated_at=?
                            WHERE stock_id=? AND month_key=?""",(mom,yoy,cyoy,now,sid,latest_mk)); updated+=1
        if any(v is None for v in after): not_applicable+=1
    conn.commit()
    return updated,not_applicable

def fetch_official_month_revenue(conn: sqlite3.Connection, target: dt.date, cfg: configparser.ConfigParser, stock_ids: Optional[List[str]]=None) -> pd.DataFrame:
    endpoints=[
        ("TWSE-月營收",TWSE_REVENUE_URL,"TWSE"),
        ("TPEx-月營收",TPEX_REVENUE_URL,"TPEx"),
        # v29：補興櫃月營收。新上市/轉板公司在比較月份可能仍屬興櫃，避免市場身分切換造成空洞。
        ("TPEx-興櫃月營收","https://www.tpex.org.tw/openapi/v1/t187ap05_R","TPEx"),
    ]
    if target != dt.date.today():
        # v25 根因修正：舊版只要有部分 snapshot 就直接回傳，會把後來已補齊的 history 擋掉。
        # 月營收 snapshot 建立時也同步寫入 history，因此歷史重跑直接以 history as-of target 為母體。
        hist = _latest_revenue_history_asof(conn,target,stock_ids)
        wanted = set(str(x) for x in stock_ids) if stock_ids else set()
        latest_required={sid:(_available_month_keys_for_stock(conn,sid,target,1)[-1] if _available_month_keys_for_stock(conn,sid,target,1) else None) for sid in wanted}
        required={sid for sid,mk in latest_required.items() if mk is not None}
        have_month={}
        if not hist.empty:
            have_month={str(r["stock_id"]):int(r["month_key"]) for _,r in hist.iterrows()}
        missing=sorted(sid for sid in required if have_month.get(sid)!=latest_required.get(sid))
        # 若 SQLite 歷史月營收有缺檔，再利用 TWSE/TPEx 官方「目前最新月份」安全回填。
        # 只有該月份的保守可得日（次月10日）<= target 才能使用，因此不會用未來資料污染歷史日。
        official_added = 0; official_errors: List[str] = []
        if missing:
            for source,url,market in endpoints:
                try:
                    rows = _fetch_official_revenue_rows(source,url,market)
                    official_added += _store_official_revenue_as_history_if_safe(conn,target,rows,missing)
                except Exception as e:
                    official_errors.append(f"{source}:{e}")
            hist = _latest_revenue_history_asof(conn,target,stock_ids)
            have_month={str(r["stock_id"]):int(r["month_key"]) for _,r in hist.iterrows()} if not hist.empty else {}
            missing=sorted(sid for sid in required if have_month.get(sid)!=latest_required.get(sid))
        complete = (not required) or (len(missing) == 0)
        msg = f"歷史重跑：使用初始化歷史月營收 as-of {len(hist)}/{len(required)} 檔；官方最新月份安全補回 {official_added} 筆；仍缺 {len(missing)} 檔"
        if official_errors: msg += "；官方備援錯誤=" + " | ".join(official_errors[:3])
        record_source(conn,target,"營收Point-in-time",target.isoformat() if not hist.empty else None,bool(not hist.empty),complete,msg)
        return hist

    # 今日若候選股的最新應有月份都已存在 history，直接沿用；月營收不用每天重抓。
    wanted=set(str(x) for x in (stock_ids or []))
    if wanted:
        hist=_latest_revenue_history_asof(conn,target,stock_ids)
        latest_required={sid:(_available_month_keys_for_stock(conn,sid,target,1)[-1] if _available_month_keys_for_stock(conn,sid,target,1) else None) for sid in wanted}
        required={sid for sid,mk in latest_required.items() if mk is not None}
        have_month={str(r["stock_id"]):int(r["month_key"]) for _,r in hist.iterrows()} if not hist.empty else {}
        missing=[sid for sid in required if have_month.get(sid)!=latest_required.get(sid)]
        if not missing:
            record_source(conn,target,"營收Point-in-time",target.isoformat(),True,True,
                          f"增量更新：最新應有月份已存在，沿用 {len(hist)}/{len(required)} 檔，不重抓相同月份")
            return hist

    for source,url,market in endpoints:
        try:
            lr = _fetch_official_revenue_rows(source,url,market)
            _save_revenue_snapshot(conn,target,lr)
            lm=max([r["month_key"] for r in lr],default=None); label=f"{lm//100}-{lm%100:02d}" if lm else None
            record_source(conn,target,source,label,bool(lr),len(lr)>=100,f"官方月營收可解析 {len(lr)} 檔；已保存 point-in-time 快照")
        except Exception as e:
            record_source(conn,target,source,None,False,False,f"失敗：{e}；改用 target 日以前已存快照/歷史回補")
            log.warning("%s 失敗：%s",source,e)
    live=_latest_revenue_snapshot(conn,target)
    if not live.empty: return live
    return _latest_revenue_history_asof(conn,target,stock_ids)


def build_fundamental_filters(conn: sqlite3.Connection, target: dt.date, current: pd.DataFrame, cfg: configparser.ConfigParser) -> Tuple[pd.DataFrame,pd.DataFrame,pd.DataFrame,pd.DataFrame]:
    """由呼叫端傳入正式鴨嘴＋預備鴨嘴母體；第一次回補歷史營收/季報，之後每日只補缺口。"""
    stock_ids=[str(x) for x in current.get("stock_id",pd.Series(dtype=str)).tolist()] if not current.empty else []
    if stock_ids:
        # v30：先修正舊版資料庫內用一般時限寫入的保守 availability_date，避免歷史穿越/假缺漏。
        _normalize_special_availability_dates(conn,target)
        # v30：每日核心流程不再逐檔撞 MOPS 舊 HTML，且抓完必做逐檔驗收。
        # 先抓 TWSE/TPEx 六類官方 OpenAPI 最新季，再立刻用 Yahoo 補「當期必要季度」缺口。
        # 12季深度歷史是研究用途，不應阻塞每日鴨嘴/三率結果。
        bulk_added, bulk_errors = _seed_bulk_income_for_target(conn,target,stock_ids)
        yf_ok,yf_fail,yf_errors = _yfinance_income_backfill_missing(conn,target,current,cfg)
        if cfg.getboolean("general","deep_financial_history_on_daily",fallback=False):
            _backfill_income_history(conn,target,stock_ids,cfg)
        else:
            record_source(conn,target,"財報深度歷史",target.isoformat(),True,True,
                          "v30 每日驗收模式：不執行 MOPS 舊 HTML 12季深度回補；當期季報改用官方 OpenAPI＋Yahoo，最後逐檔驗收。")
        validation=_record_income_validation(conn,target,stock_ids)
        real_missing=int((validation["status"]=="缺資料").sum()) if not validation.empty else 0
        prior_short=int((validation["status"]=="比較期歷史不足").sum()) if not validation.empty else 0
        record_source(conn,target,"基本面備援",target.isoformat(),True,(real_missing==0),
                      f"官方六類批次補入 {bulk_added} 筆；Yahoo驗收後補齊/可分類 {yf_ok}、備援未完全補齊 {yf_fail}；真正缺最新資料 {real_missing}、比較期歷史不足 {prior_short}" +
                      (f"；官方端點異常 {len(bulk_errors)} 個" if bulk_errors else "") +
                      (f"；Yahoo首批={' | '.join(yf_errors[:5])}" if yf_errors else ""))
        _backfill_revenue_history(conn,target,stock_ids,cfg)
        rev_recalc, rev_na = _recompute_revenue_growth_from_history(conn,target,stock_ids)
        record_source(conn,target,"營收成長率重算",target.isoformat(),True,True,
                      f"由原始月營收補算成長率 {rev_recalc} 檔；仍有分母為0/歷史不足而無法計算 {rev_na} 檔（此類不視為抓取失敗）")
    income_snap=fetch_official_income_snapshot(conn,target,stock_ids)
    revenue=fetch_official_month_revenue(conn,target,cfg,stock_ids)
    # fetch_official_month_revenue 可能剛用上市/上櫃/興櫃最新端點安全補回缺檔；
    # 再補算一次，確保新寫入的原始營收也立即得到可計算的成長率。
    if stock_ids:
        post_recalc, post_na = _recompute_revenue_growth_from_history(conn,target,stock_ids)
        if post_recalc:
            revenue=_latest_revenue_history_asof(conn,target,stock_ids)
    if stock_ids:
        raw_have=set(); metric_have=set(); rate_na=set(); miss_rev=[]
        for sid in stock_ids:
            latest_list=_available_month_keys_for_stock(conn,sid,target,1)
            if not latest_list:
                continue
            lm=latest_list[-1]
            rr=conn.execute("""SELECT revenue,revenue_mom_pct,revenue_yoy_pct,revenue_cum_yoy_pct
                               FROM fundamental_revenue_history
                               WHERE stock_id=? AND month_key=? AND availability_date<=?
                               ORDER BY availability_date DESC LIMIT 1""",(sid,lm,target.isoformat())).fetchone()
            if not rr or (clean_num(rr[0]) is None and not any(clean_num(x) is not None for x in rr[1:])):
                miss_rev.append(sid); continue
            raw_have.add(sid)
            vals=[clean_num(x) for x in rr[1:]]
            if all(v is not None for v in vals): metric_have.add(sid)
            else: rate_na.add(sid)
        complete=len(miss_rev)==0
        record_source(conn,target,"營收驗收",target.isoformat(),True,complete,
                      f"逐檔法定時限驗收 {len(stock_ids)} 檔：完整成長率 {len(metric_have)}、成長率不適用/基期不足 {len(rate_na)}、真正缺依法應有月營收 {len(miss_rev)}" +
                      (f"；真正缺漏代號={','.join(miss_rev[:30])}" if miss_rev else "；當期營收無實質抓取缺漏"))
    if current.empty:
        base=current.copy()
        for c in ["培育中心類別","毛差","營差","淨差","三率資料期","三率快照日","三率可得日依據","三率來源","營收月增%","營收年增%","累計營收增%","營收資料月","營收快照日","營收可得日依據"]: base[c]=pd.Series(dtype="object")
        return base,base.copy(),base.copy(),base.copy()
    allow_equal=cfg.getboolean("general","three_rates_allow_equal",fallback=True); revmin=cfg.getfloat("general","revenue_growth_min_pct",fallback=0.0)
    prof_map={}
    if not income_snap.empty:
        isub=income_snap[income_snap["stock_id"].astype(str).isin(stock_ids)]
        for _,snap in isub.iterrows():
            x=_three_rates_for_stock(conn,snap,allow_equal)
            if x: prof_map[str(snap["stock_id"])]=x
    rev_map={}
    if not revenue.empty:
        rsub=revenue[revenue["stock_id"].astype(str).isin(stock_ids)]
        for _,r in rsub.iterrows():
            mom,yoy,cum=(clean_num(r.get("revenue_mom_pct")),clean_num(r.get("revenue_yoy_pct")),clean_num(r.get("revenue_cum_yoy_pct"))); mk=int(r["month_key"])
            if None in (mom,yoy,cum):
                continue
            rev_map[str(r["stock_id"]) ]={"營收月增%":mom,"營收年增%":yoy,"累計營收增%":cum,"營收三增":bool(mom>revmin and yoy>revmin and cum>revmin),"營收資料月":f"{mk//100}-{mk%100:02d}","營收快照日":str(r["observed_date"]),"營收可得日依據":str(r.get("availability_basis") or "official_snapshot")}
    out=current.copy()
    extra=["毛利率本季%","毛利率上季%","營業利益率本季%","營業利益率上季%","稅後純益率本季%","稅後純益率上季%","毛差","營差","淨差","三率資料期","三率快照日","三率可得日依據","三率來源","基本面資料狀態","基本面缺漏原因","基本面所需季度","營收資料狀態","營收月增%","營收年增%","累計營收增%","營收資料月","營收快照日","營收可得日依據"]
    for c in extra: out[c]=None
    out["三率三升"]=False; out["營收三增"]=False
    vdf=_income_validation_df(conn,target,stock_ids)
    vmap={str(r["stock_id"]):r for _,r in vdf.iterrows()} if not vdf.empty else {}
    for i,r in out.iterrows():
        sid=str(r["stock_id"])
        for k,v in prof_map.get(sid,{}).items(): out.at[i,k]=v
        for k,v in rev_map.get(sid,{}).items(): out.at[i,k]=v
        vv=vmap.get(sid)
        if vv is not None:
            out.at[i,"基本面資料狀態"]=vv.get("status")
            out.at[i,"基本面缺漏原因"]=vv.get("reason")
            out.at[i,"基本面所需季度"]=vv.get("required")
        if sid in rev_map:
            out.at[i,"營收資料狀態"]="完整"
        elif len(_available_month_keys_for_stock(conn,sid,target,1))==0:
            out.at[i,"營收資料狀態"]="上市時間不足/不適用"
        else:
            lm=max(_available_month_keys_for_stock(conn,sid,target,1),default=None)
            rawrow=conn.execute("""SELECT revenue,revenue_mom_pct,revenue_yoy_pct,revenue_cum_yoy_pct
                                   FROM fundamental_revenue_history
                                   WHERE stock_id=? AND month_key=? AND availability_date<=?
                                   LIMIT 1""",(sid,lm,target.isoformat())).fetchone() if lm else None
            if rawrow and (clean_num(rawrow[0]) is not None or any(clean_num(x) is not None for x in rawrow[1:])):
                out.at[i,"營收資料狀態"]="成長率不適用/基期不足"
            else:
                out.at[i,"營收資料狀態"]="缺資料"
    def cat(r):
        a,b=bool(r.get("三率三升",False)),bool(r.get("營收三增",False))
        return "三率三升＋營收三增" if a and b else ("三率三升" if a else ("營收三增" if b else ""))
    out["培育中心類別"]=out.apply(cat,axis=1)
    tr=out[out["三率三升"]].copy().sort_values(["bias20_pct","stock_id"],ascending=[False,True])
    rv=out[out["營收三增"]].copy().sort_values(["bias20_pct","stock_id"],ascending=[False,True])
    both=out[out["三率三升"] & out["營收三增"]].copy().sort_values(["bias20_pct","stock_id"],ascending=[False,True])
    return out,tr,rv,both


# ============================
# EPS／本益比估值
# ============================
INSTITUTION_VALUATION_CSV = APP_DIR / "institution_valuation.csv"


def _period_prev_n(year: int, quarter: int, n: int = 1) -> Tuple[int, int]:
    y, q = year, quarter
    for _ in range(n):
        y, q = _prev_period(y, q)
    return y, q


def _eps_metrics_for_stock(conn: sqlite3.Connection, stock_id: str, target: dt.date) -> dict:
    hist = pd.read_sql_query(
        """SELECT stock_id,year,quarter,revenue_ytd,gross_profit_ytd,operating_income_ytd,net_income_ytd,eps_ytd,
                  availability_date,availability_basis,source
           FROM fundamental_income_history
           WHERE stock_id=? AND availability_date IS NOT NULL AND availability_date<=?
           ORDER BY year,quarter""",
        conn, params=(stock_id, target.isoformat())
    )
    if hist.empty:
        return {}
    pmap: Dict[Tuple[int,int], dict] = {}
    for r in hist.to_dict("records"):
        pmap[(int(r["year"]), int(r["quarter"]))] = r
    latest_key = max(pmap.keys())
    ly, lq = latest_key
    latest_sq = _single_quarter_values(pmap, ly, lq)
    if not latest_sq:
        return {}
    latest_eps = latest_sq.get("eps_ytd")
    py, pq = _prev_period(ly, lq)
    prev_sq = _single_quarter_values(pmap, py, pq)
    prev_eps = prev_sq.get("eps_ytd") if prev_sq else None
    yoy_sq = _single_quarter_values(pmap, ly - 1, lq)
    yoy_eps = yoy_sq.get("eps_ytd") if yoy_sq else None

    # 最近四個連續單季 EPS 加總成 TTM EPS。
    ttm_parts: List[float] = []
    contiguous = True
    for i in range(4):
        y, q = _period_prev_n(ly, lq, i)
        sq = _single_quarter_values(pmap, y, q)
        if not sq or sq.get("eps_ytd") is None:
            contiguous = False
            break
        ttm_parts.append(float(sq["eps_ytd"]))
    ttm_eps = sum(ttm_parts) if contiguous and len(ttm_parts) == 4 else None
    yoy_growth = None
    if latest_eps is not None and yoy_eps is not None and abs(float(yoy_eps)) > 1e-12:
        yoy_growth = (float(latest_eps) / abs(float(yoy_eps)) - (1 if float(yoy_eps) > 0 else -1)) * 100
        # 負基期時百分比容易失真，改只標示數值、不給成長率。
        if float(yoy_eps) <= 0:
            yoy_growth = None
    return {
        "最新單季EPS": latest_eps,
        "前一季EPS": prev_eps,
        "去年同期EPS": yoy_eps,
        "TTM EPS": ttm_eps,
        "EPS年增率%": yoy_growth,
        "EPS資料期": f"{ly}Q{lq}",
        "EPS來源": "MOPS綜合損益",
        "EPS觀測日": None,
    }


# ============================
# Yahoo Finance / yfinance 季度 EPS
# ============================
def _yahoo_symbol(stock_id: str, market: str) -> str:
    suffix = ".TW" if str(market).upper() == "TWSE" else ".TWO"
    return f"{str(stock_id).strip()}{suffix}"


def _normalize_financial_label(v: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(v).lower())


def _find_eps_row(df: pd.DataFrame, candidates: List[str]) -> Optional[Any]:
    if df is None or df.empty:
        return None
    wanted = {_normalize_financial_label(x) for x in candidates}
    for idx in df.index:
        key = _normalize_financial_label(idx)
        if key in wanted:
            return idx
    return None



def _find_fin_row(df: pd.DataFrame, candidates: List[str]) -> Optional[Any]:
    return _find_eps_row(df, candidates)


def _fetch_yfinance_income_one(stock_id: str, market: str, target: dt.date, name: str = "") -> List[dict]:
    """Yahoo Finance 缺漏備援：抓季度損益，轉成 DB 使用的「累季」格式。

    Yahoo 季報欄位為單季值；為相容既有三率邏輯，按曆年由 Q1 起累加成 YTD。
    只保留保守可得日 <= target 的季度，避免把 target 之後的季度塞回歷史。
    """
    if yf is None:
        return []
    symbol = _yahoo_symbol(stock_id, market)
    t = yf.Ticker(symbol)
    frames=[]
    for getter in (
        lambda: t.quarterly_income_stmt,
        lambda: t.quarterly_financials,
        lambda: t.get_income_stmt(freq="quarterly", pretty=False),
        lambda: t.get_income_stmt(freq="quarterly", pretty=True),
        lambda: t.get_financials(freq="quarterly"),
    ):
        try:
            df=getter()
            if isinstance(df,pd.DataFrame) and not df.empty:
                frames.append(df)
        except Exception:
            pass
    best=None
    for df in frames:
        rr=_find_fin_row(df,[
            "Total Revenue","Operating Revenue","Revenue","TotalRevenue","OperatingRevenue",
            "Interest Revenue","Net Interest Income","Total Operating Income As Reported"
        ])
        op=_find_fin_row(df,[
            "Operating Income","OperatingIncome","Operating Income Loss","OperatingIncomeLoss",
            "Total Operating Income As Reported","TotalOperatingIncomeAsReported"
        ])
        net=_find_fin_row(df,[
            "Net Income","NetIncome","Net Income Common Stockholders","NetIncomeCommonStockholders",
            "Net Income Including Noncontrolling Interests","NetIncomeIncludingNoncontrollingInterests"
        ])
        gross=_find_fin_row(df,["Gross Profit","GrossProfit","Gross Profit As Reported","GrossProfitAsReported"])
        basic=_find_fin_row(df,["Basic EPS","BasicEPS","Basic EPS Continuing Operations"])
        diluted=_find_fin_row(df,["Diluted EPS","DilutedEPS","Diluted EPS Continuing Operations"])
        if rr is not None and op is not None and net is not None:
            best=(df,rr,gross,op,net,basic,diluted); break
    if best is None:
        return []
    df,rr,gross_r,op_r,net_r,basic_r,diluted_r=best
    singles=[]
    for col in df.columns:
        pdte=pd.to_datetime(col,errors="coerce")
        if pd.isna(pdte):
            continue
        d=pdte.date(); q=(d.month-1)//3+1; y=d.year
        avail=_income_available_for_identity(str(stock_id), str(name or ""), y, q)
        if avail>target:
            continue
        revenue=clean_num(df.at[rr,col]) if rr is not None else None
        gross=clean_num(df.at[gross_r,col]) if gross_r is not None else None
        op=clean_num(df.at[op_r,col]) if op_r is not None else None
        net=clean_num(df.at[net_r,col]) if net_r is not None else None
        basic=clean_num(df.at[basic_r,col]) if basic_r is not None else None
        dil=clean_num(df.at[diluted_r,col]) if diluted_r is not None else None
        eps=basic if basic is not None else dil
        if revenue is None or op is None or net is None:
            continue
        singles.append({"year":y,"quarter":q,"revenue":revenue,"gross":gross,"op":op,"net":net,"eps":eps})
    # 去重並由舊到新，按每年累加成 YTD。若缺 Q1，該年的 Q2+不寫入，避免錯把單季當累季。
    by={(x["year"],x["quarter"]):x for x in singles}
    out=[]
    years=sorted({y for y,q in by})
    for y in years:
        cum_rev=cum_gross=cum_op=cum_net=cum_eps=0.0
        gross_valid=True; eps_valid=True; started=False
        for q in range(1,5):
            x=by.get((y,q))
            if not x:
                if q==1: started=False
                continue
            if q>1 and not started:
                # 沒有 Q1 無法重建正確累季。
                continue
            if q==1: started=True
            cum_rev += float(x["revenue"]); cum_op += float(x["op"]); cum_net += float(x["net"])
            if x["gross"] is None: gross_valid=False
            elif gross_valid: cum_gross += float(x["gross"])
            if x["eps"] is None: eps_valid=False
            elif eps_valid: cum_eps += float(x["eps"])
            out.append({
                "stock_id":str(stock_id),"year":y,"quarter":q,"revenue_ytd":cum_rev,
                "gross_profit_ytd":cum_gross if gross_valid else None,"operating_income_ytd":cum_op,
                "net_income_ytd":cum_net,"eps_ytd":cum_eps if eps_valid else None,
                "source":"Yahoo Finance quarterly income fallback",
            })
    return out


def _income_validation_one(conn: sqlite3.Connection, stock_id: str, target: dt.date) -> dict:
    """v30逐檔驗收：先判斷策略適用性，再用『該公司自己的法定期限』決定應有季度。

    缺最新依法應有季度 = 真正缺資料；只有較舊比較期不足 = 比較期歷史不足，不再誤報抓取管線失敗。
    """
    sid=str(stock_id)
    if _is_three_rate_na_stock(conn,sid,target):
        return {"stock_id":sid,"status":"三率不適用","reason":"金融/保險/金控/異業報表結構不同，不以一般業毛利率/營益率/淨利率三率判斷","required":""}
    req=_three_rate_required_periods_for_stock(conn,sid,target)
    latest=_available_quarters_for_stock(conn,sid,target,1)
    latest_period=latest[-1] if latest else None
    if not req or latest_period is None:
        return {"stock_id":sid,"status":"上市時間不足/不適用","reason":"target日前沒有足夠可比季度","required":""}
    rows=pd.read_sql_query("""SELECT year,quarter,revenue_ytd,gross_profit_ytd,operating_income_ytd,net_income_ytd,source
                              FROM fundamental_income_history
                              WHERE stock_id=? AND availability_date IS NOT NULL AND availability_date<=?""",
                           conn,params=(sid,target.isoformat()))
    pmap={(int(r.year),int(r.quarter)):r.to_dict() for _,r in rows.iterrows()} if not rows.empty else {}
    mixed_like=any(any(tok in str(rr.get("source") or "").lower() for tok in ["_mim","異業"]) for rr in pmap.values())
    missing_period=[]; core_missing=[]; gross_missing=[]
    for y,q in req:
        rr=pmap.get((y,q))
        if rr is None:
            missing_period.append((y,q)); continue
        for k,label in [("revenue_ytd","營收/收益"),("operating_income_ytd","營業利益"),("net_income_ytd","淨利")]:
            v=rr.get(k)
            if v is None or pd.isna(v): core_missing.append(f"{y}Q{q}:{label}")
        v=rr.get("gross_profit_ytd")
        if v is None or pd.isna(v): gross_missing.append(f"{y}Q{q}:毛利")
    required=",".join(f"{y}Q{q}" for y,q in req)
    if latest_period in missing_period:
        return {"stock_id":sid,"status":"缺資料","reason":"缺依法應有最新季="+f"{latest_period[0]}Q{latest_period[1]}","required":required}
    # 最新季存在，但較舊比較季不足，不把它叫做資料源失敗。
    if missing_period:
        first=_first_trade_date(conn,sid,target)
        if first is not None and all(_quarter_end_date(y,q) < first for y,q in missing_period):
            miss_txt=",".join(f"{y}Q{q}" for y,q in missing_period)
            return {"stock_id":sid,"status":"上市時間不足/不適用","reason":"上市前比較期無可用財報="+miss_txt,"required":required}
        return {"stock_id":sid,"status":"比較期歷史不足","reason":"最新季已取得；較舊比較期缺="+",".join(f"{y}Q{q}" for y,q in missing_period),"required":required}
    # 最新季若核心欄位缺，仍屬真正資料不足；較舊季欄位缺則標比較期不足。
    latest_label=f"{latest_period[0]}Q{latest_period[1]}:"
    all_missing=core_missing+gross_missing
    if all_missing and mixed_like:
        return {"stock_id":sid,"status":"三率不適用","reason":"官方異業報表結構無法完整對應一般業三率欄位","required":required}
    if any(x.startswith(latest_label) for x in all_missing):
        return {"stock_id":sid,"status":"缺資料","reason":"最新季欄位缺漏="+"；".join(all_missing[:8]),"required":required}
    if all_missing:
        return {"stock_id":sid,"status":"比較期歷史不足","reason":"最新季已取得；較舊欄位缺漏="+"；".join(all_missing[:8]),"required":required}
    return {"stock_id":sid,"status":"可計算三率","reason":"","required":required}

def _income_validation_df(conn: sqlite3.Connection, target: dt.date, stock_ids: List[str]) -> pd.DataFrame:
    return pd.DataFrame([_income_validation_one(conn,str(sid),target) for sid in stock_ids]) if stock_ids else pd.DataFrame(columns=["stock_id","status","reason","required"])


def _record_income_validation(conn: sqlite3.Connection, target: dt.date, stock_ids: List[str]) -> pd.DataFrame:
    df=_income_validation_df(conn,target,stock_ids)
    if df.empty:
        return df
    missing=df.loc[df["status"]=="缺資料","stock_id"].astype(str).tolist()
    usable=int((df["status"]=="可計算三率").sum())
    na=int((df["status"].isin(["三率不適用","上市時間不足/不適用"])).sum())
    prior_short=int((df["status"]=="比較期歷史不足").sum())
    complete=len(missing)==0
    msg=(f"逐檔驗收 {len(df)} 檔：可計算三率 {usable}、三率不適用/上市時間不足 {na}、比較期歷史不足 {prior_short}、真正缺最新資料 {len(missing)}"
         + (f"；真正缺漏代號={','.join(missing[:30])}" if missing else "；當期最新資料無實質缺漏"))
    record_source(conn,target,"基本面驗收",target.isoformat(),True,complete,msg)
    return df


def _yfinance_income_backfill_missing(conn: sqlite3.Connection, target: dt.date, current: pd.DataFrame,
                                      cfg: configparser.ConfigParser) -> Tuple[int,int,List[str]]:
    """v27：只對「當期三率真正需要的確切季度」不足者使用 Yahoo 備援。

    舊版只要資料庫裡任意有兩個可算季度就不補，會出現已有舊季、但最新 Q2/Q1 仍缺的情況。
    v27 改成依 target 的最新可得季，逐檔檢查所需期別。
    """
    if yf is None or current.empty:
        return 0,0,[]
    meta={str(r["stock_id"]):(str(r.get("name") or ""),str(r.get("market") or "")) for _,r in current.iterrows()}
    need=[]
    for sid in meta:
        # v29：是否要進 Yahoo 備援，直接以「最終驗收結果」為準。
        # v28 只檢查營收/營益/淨利，會漏掉「只有毛利缺值」的普通產業，造成最後仍缺資料卻沒進 Yahoo。
        check=_income_validation_one(conn,sid,target)
        if check.get("status") in ("缺資料","比較期歷史不足"):
            need.append(sid)
    maxn=max(20,cfg.getint("general","yfinance_income_fallback_max_stocks",fallback=350))
    need=need[:maxn]
    if not need:
        return 0,0,[]
    workers=max(1,min(6,cfg.getint("general","yfinance_income_workers",fallback=3)))
    ok=fail=0; errors=[]
    print(f"    [Yahoo財報備援] 當期必要季度不足 {len(need)} 檔；workers={workers}",flush=True)
    def task(sid):
        return sid,_fetch_yfinance_income_one(sid,meta[sid][1],target,meta[sid][0])
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs={ex.submit(task,sid):sid for sid in need}
        done=0
        for fut in as_completed(futs):
            sid=futs[fut]; done+=1
            try:
                _,rows=fut.result()
                if rows:
                    for r in rows:
                        avail=_income_available_for_identity(sid,meta[sid][0],int(r["year"]),int(r["quarter"]))
                        _upsert_income_record_with_basis(conn,r,avail,"yfinance_conservative_deadline",r.get("source"))
                    conn.commit()
                    check=_income_validation_one(conn,sid,target)
                    if check["status"] in ("可計算三率","三率不適用","上市時間不足/不適用","比較期歷史不足"):
                        ok+=1
                    else:
                        fail+=1; errors.append(f"{sid}:Yahoo有回資料但驗收未通過:{check['reason']}")
                else:
                    fail+=1; errors.append(f"{sid}:Yahoo無可用季報")
            except Exception as e:
                fail+=1; errors.append(f"{sid}:{type(e).__name__}:{str(e)[:80]}")
            if done==1 or done%20==0 or done==len(need):
                print(f"        Yahoo財報 {done}/{len(need)}：成功 {ok}、無資料/失敗 {fail}",flush=True)
    return ok,fail,errors


def _fetch_yfinance_eps_one(stock_id: str, market: str) -> List[dict]:
    """抓 Yahoo Finance 的季度 Basic/Diluted EPS。

    yfinance 官方目前同時提供 quarterly_income_stmt 與 quarterly_financials。
    以 quarterly_income_stmt 優先；若資料框為空，再用 quarterly_financials 備援。
    Basic EPS 優先，缺值才用 Diluted EPS。
    """
    if yf is None:
        raise RuntimeError("未安裝 yfinance；請先 pip install yfinance")
    symbol = _yahoo_symbol(stock_id, market)
    ticker = yf.Ticker(symbol)
    frames: List[pd.DataFrame] = []
    try:
        q = ticker.quarterly_income_stmt
        if isinstance(q, pd.DataFrame) and not q.empty:
            frames.append(q)
    except Exception:
        pass
    try:
        q2 = ticker.quarterly_financials
        if isinstance(q2, pd.DataFrame) and not q2.empty:
            frames.append(q2)
    except Exception:
        pass
    # 某些 yfinance 版本 pretty=True 的欄名較容易找到 Basic EPS。
    try:
        q3 = ticker.get_income_stmt(freq="quarterly", pretty=True)
        if isinstance(q3, pd.DataFrame) and not q3.empty:
            frames.append(q3)
    except Exception:
        pass
    if not frames:
        return []

    chosen = None
    basic_row = diluted_row = None
    for df in frames:
        br = _find_eps_row(df, ["Basic EPS", "BasicEPS", "Basic EPS Continuing Operations"])
        dr = _find_eps_row(df, ["Diluted EPS", "DilutedEPS", "Diluted EPS Continuing Operations"])
        if br is not None or dr is not None:
            chosen, basic_row, diluted_row = df, br, dr
            break
    if chosen is None:
        return []

    rows: List[dict] = []
    for col in chosen.columns:
        try:
            period = pd.to_datetime(col, errors="coerce")
            if pd.isna(period):
                continue
            period_date = period.date().isoformat()
        except Exception:
            continue
        basic = clean_num(chosen.at[basic_row, col]) if basic_row is not None else None
        diluted = clean_num(chosen.at[diluted_row, col]) if diluted_row is not None else None
        used = basic if basic is not None else diluted
        if used is None:
            continue
        rows.append({
            "stock_id": str(stock_id), "market": str(market), "yahoo_symbol": symbol,
            "period_end": period_date, "basic_eps": basic, "diluted_eps": diluted, "used_eps": used,
            "source": "Yahoo Finance via yfinance",
        })
    rows.sort(key=lambda r: r["period_end"], reverse=True)
    # 只保留足夠計算 TTM / 去年同期的近期欄位。
    return rows[:8]


def _save_yfinance_eps_rows(conn: sqlite3.Connection, observed: dt.date, rows: List[dict]) -> None:
    if not rows:
        return
    now = dt.datetime.now().isoformat(timespec="seconds")
    vals = [(observed.isoformat(), r["stock_id"], r.get("market"), r.get("yahoo_symbol"),
             r["period_end"], r.get("basic_eps"), r.get("diluted_eps"), r.get("used_eps"),
             r.get("source"), now) for r in rows]
    conn.executemany(
        """INSERT OR REPLACE INTO yfinance_eps_snapshot
           (observed_date,stock_id,market,yahoo_symbol,period_end,basic_eps,diluted_eps,used_eps,source,fetched_at)
           VALUES (?,?,?,?,?,?,?,?,?,?)""", vals)
    conn.commit()


def _yfinance_eps_metrics_for_stock(conn: sqlite3.Connection, stock_id: str, target: dt.date) -> dict:
    """只使用 target 當日以前『實際觀測到』的 Yahoo EPS 快照，避免歷史重跑穿越。"""
    row = conn.execute(
        "SELECT MAX(observed_date) FROM yfinance_eps_snapshot WHERE stock_id=? AND observed_date<=?",
        (str(stock_id), target.isoformat())).fetchone()
    obs = str(row[0]) if row and row[0] else ""
    if not obs:
        return {}
    df = pd.read_sql_query(
        """SELECT period_end,basic_eps,diluted_eps,used_eps,source FROM yfinance_eps_snapshot
           WHERE stock_id=? AND observed_date=? ORDER BY period_end DESC""",
        conn, params=(str(stock_id), obs))
    if df.empty:
        return {}
    df["used_eps"] = pd.to_numeric(df["used_eps"], errors="coerce")
    df = df.dropna(subset=["used_eps"]).copy()
    if df.empty:
        return {}
    df["period_dt"] = pd.to_datetime(df["period_end"], errors="coerce")
    df = df.dropna(subset=["period_dt"]).sort_values("period_dt", ascending=False).drop_duplicates("period_dt")
    if df.empty:
        return {}

    latest = float(df.iloc[0]["used_eps"])
    prev = float(df.iloc[1]["used_eps"]) if len(df) >= 2 else None
    latest_dt = df.iloc[0]["period_dt"]
    yoy = None
    if pd.notna(latest_dt):
        target_y = int(latest_dt.year) - 1
        target_m = int(latest_dt.month)
        matches = df[(df["period_dt"].dt.year == target_y) & (df["period_dt"].dt.month == target_m)]
        if not matches.empty:
            yoy = float(matches.iloc[0]["used_eps"])
    # Yahoo 的 quarterly income statement 為單季資料；只有最近四個曆季連續時才加總成 TTM。
    ttm = None
    if len(df) >= 4:
        qkeys = []
        for x in df.head(4)["period_dt"]:
            qkeys.append(int(x.year) * 4 + ((int(x.month) - 1) // 3))
        if all(qkeys[i] - qkeys[i + 1] == 1 for i in range(3)):
            ttm = float(df.head(4)["used_eps"].sum())
    growth = None
    if yoy is not None and yoy > 0:
        growth = (latest / yoy - 1.0) * 100.0
    q = ((int(latest_dt.month) - 1) // 3 + 1) if pd.notna(latest_dt) else None
    period_label = f"{int(latest_dt.year)}Q{q} ({latest_dt.date().isoformat()})" if q else str(df.iloc[0]["period_end"])
    return {
        "最新單季EPS": latest, "前一季EPS": prev, "去年同期EPS": yoy,
        "TTM EPS": ttm, "EPS年增率%": growth, "EPS資料期": period_label,
        "EPS來源": "Yahoo Finance/yfinance", "EPS觀測日": obs,
        "Yahoo TTM EPS": ttm, "Yahoo最新單季EPS": latest,
    }


def _refresh_yfinance_eps(conn: sqlite3.Connection, target: dt.date, meta: Dict[str, Tuple[str,str]],
                          priority_ids: List[str], cfg: configparser.ConfigParser) -> None:
    """刷新 Yahoo EPS。預設只抓『估值接近買進區＋鴨嘴』，避免每天對近2000檔逐檔打 Yahoo。"""
    enabled = cfg.getboolean("general", "yfinance_eps_enabled", fallback=True)
    if not enabled:
        record_source(conn, target, "Yahoo Finance EPS", None, True, True, "已在 settings.ini 關閉")
        return
    if yf is None:
        record_source(conn, target, "Yahoo Finance EPS", None, False, False, "未安裝 yfinance；請執行 pip install yfinance")
        return

    today = dt.date.today()
    # 嚴格 PIT：今天才抓到的 Yahoo 財報，不倒灌到過去日期。
    if target < today:
        existing = int(conn.execute(
            "SELECT COUNT(DISTINCT stock_id) FROM yfinance_eps_snapshot WHERE observed_date<=?",
            (target.isoformat(),)).fetchone()[0])
        record_source(conn, target, "Yahoo Finance EPS", target.isoformat(), True, True,
                      f"歷史重跑：不抓今日 Yahoo 快照以避免未來資料穿越；可使用 target 日以前已保存快照 {existing} 檔")
        return

    ids = []
    seen = set()
    for sid in priority_ids:
        sid = str(sid)
        if sid in meta and sid not in seen:
            ids.append(sid); seen.add(sid)
    max_stocks = max(20, cfg.getint("general", "yfinance_eps_max_stocks", fallback=400))
    ids = ids[:max_stocks]
    if not ids:
        record_source(conn, target, "Yahoo Finance EPS", target.isoformat(), True, True, "本次沒有需要刷新 Yahoo EPS 的優先股票")
        return

    # EPS 是季度資料：近期已有成功快照就沿用，不需每天重抓。
    refresh_days=max(1,cfg.getint("general","yfinance_eps_refresh_days",fallback=7))
    cutoff=(today-dt.timedelta(days=refresh_days-1)).isoformat()
    fetch_ids = []
    cached = 0
    for sid in ids:
        got = conn.execute(
            "SELECT 1 FROM yfinance_eps_snapshot WHERE observed_date>=? AND observed_date<=? AND stock_id=? LIMIT 1",
            (cutoff,today.isoformat(), sid)).fetchone()
        if got:
            cached += 1
        else:
            fetch_ids.append(sid)

    try:
        # 新版 yfinance 支援全域 retries；舊版沒有也不影響。
        yf.config.network.retries = 2
    except Exception:
        pass
    workers = max(1, min(8, cfg.getint("general", "yfinance_eps_workers", fallback=4)))
    ok = fail = 0; errors: List[str] = []
    total = len(fetch_ids)
    if total:
        print(f"    [Yahoo EPS] 優先名單 {len(ids)} 檔；已快取 {cached}；待抓 {total}；workers={workers}", flush=True)
        def task(sid: str):
            name, market = meta[sid]
            return sid, _fetch_yfinance_eps_one(sid, market)
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(task, sid): sid for sid in fetch_ids}
            done = 0
            for fut in as_completed(futs):
                sid = futs[fut]; done += 1
                try:
                    _, rows = fut.result()
                    if rows:
                        _save_yfinance_eps_rows(conn, today, rows); ok += 1
                    else:
                        fail += 1; errors.append(f"{sid}:Yahoo 無 EPS")
                except Exception as e:
                    fail += 1; errors.append(f"{sid}:{type(e).__name__}:{e}")
                if done == 1 or done % 20 == 0 or done == total:
                    print(f"        Yahoo EPS {done}/{total}：成功 {ok}、無資料/失敗 {fail}", flush=True)
    success_total = cached + ok
    msg = f"優先名單 {len(ids)} 檔；{refresh_days}日內快取 {cached}、本次成功 {ok}、無資料/失敗 {fail}。Yahoo EPS 為季度資料，不每日重抓；官方 MOPS/交易所隱含EPS仍保留備援"
    if errors:
        msg += "；首批=" + " | ".join(errors[:5])
    # Yahoo 是選配補強；少數個股查無 EPS 時由官方備援，不把整套估值標成不完整。
    yahoo_process_ok = success_total > 0 or len(ids) == 0
    record_source(conn, target, "Yahoo Finance EPS", today.isoformat(), yahoo_process_ok, yahoo_process_ok, msg)


def _month_last_calendar_days(start: dt.date, target: dt.date) -> List[dt.date]:
    """回傳 start~target 每個月份的月底候選日；實際交易日由 API 往前尋找。"""
    out: List[dt.date] = []
    y, m = start.year, start.month
    while (y, m) <= (target.year, target.month):
        if m == 12:
            next_month = dt.date(y + 1, 1, 1)
        else:
            next_month = dt.date(y, m + 1, 1)
        last = next_month - dt.timedelta(days=1)
        if (y, m) == (target.year, target.month):
            last = min(last, target)
        out.append(last)
        if m == 12:
            y, m = y + 1, 1
        else:
            m += 1
    return out


def _twse_per_day(day: dt.date) -> List[dict]:
    """TWSE 官方：上市個股日本益比/殖利率/PBR（依日期）。"""
    urls = [
        ("https://www.twse.com.tw/rwd/zh/afterTrading/BWIBBU",
         {"date": day.strftime("%Y%m%d"), "selectType": "ALL", "response": "json"}),
        ("https://www.twse.com.tw/exchangeReport/BWIBBU_d",
         {"response": "json", "date": day.strftime("%Y%m%d"), "selectType": "ALL"}),
    ]
    errors: List[str] = []
    for url, params in urls:
        try:
            js = request_json(url, params=params, timeout=10, tries=1)
            stat = str(js.get("stat") or js.get("status") or "").upper() if isinstance(js, dict) else ""
            data = js.get("data") if isinstance(js, dict) else None
            fields = js.get("fields") if isinstance(js, dict) else None
            if not isinstance(data, list) or not data:
                raise RuntimeError(f"無資料 stat={stat}")
            rows: List[dict] = []
            for row in data:
                if not isinstance(row, (list, tuple)) or len(row) < 6:
                    continue
                code = str(row[0]).strip()
                name = str(row[1]).strip()
                if not is_common_stock(code, name):
                    continue
                # TWSE 固定欄序：代號、名稱、殖利率、股利年度、本益比、PBR、財報年/季
                pe = clean_num(row[4] if len(row) > 4 else None)
                pbr = clean_num(row[5] if len(row) > 5 else None)
                dy = clean_num(row[2] if len(row) > 2 else None)
                rows.append({"stock_id": code, "name": name, "per": pe, "pbr": pbr,
                             "dividend_yield": dy, "date": day.isoformat(), "market": "TWSE"})
            if rows:
                return rows
            raise RuntimeError("可解析資料列為0")
        except Exception as e:
            errors.append(f"{url}:{e}")
    raise RuntimeError("TWSE PER 失敗：" + " | ".join(errors))


def _tpex_per_day(day: dt.date) -> List[dict]:
    """TPEx 官方：上櫃個股日本益比/殖利率/PBR（依日期）。"""
    roc = f"{day.year-1911}/{day.month:02d}/{day.day:02d}"
    endpoints = [
        ("https://www.tpex.org.tw/web/stock/aftertrading/peratio_analysis/pera_result.php",
         {"l": "zh-tw", "o": "json", "d": roc, "c": "", "s": "0,asc"}),
        ("https://www.tpex.org.tw/www/zh-tw/afterTrading/peratio",
         {"date": day.strftime("%Y/%m/%d"), "id": "", "response": "json"}),
    ]
    errors: List[str] = []
    for url, params in endpoints:
        try:
            js = request_json(url, params=params, timeout=10, tries=1)
            raw = None
            fields = None
            if isinstance(js, dict):
                raw = js.get("aaData") or js.get("data")
                if raw is None and isinstance(js.get("tables"), list) and js["tables"]:
                    tab = js["tables"][0]
                    raw = tab.get("data")
                    fields = tab.get("fields")
            if not isinstance(raw, list) or not raw:
                raise RuntimeError("無資料")
            rows: List[dict] = []
            for rr in raw:
                if isinstance(rr, dict):
                    code = str(_row_value(rr, ["股票代號","證券代號","代號","SecuritiesCompanyCode","Code"]) or "").strip()
                    name = str(_row_value(rr, ["名稱","證券名稱","股票名稱","CompanyName","Name"]) or "").strip()
                    pe = clean_num(_row_value(rr, ["本益比","本益比(倍)","PER","PEratio"]))
                    pbr = clean_num(_row_value(rr, ["股價淨值比","PBR","PBratio"]))
                    dy = clean_num(_row_value(rr, ["殖利率(%)","殖利率","DividendYield"]))
                elif isinstance(rr, (list, tuple)) and len(rr) >= 7:
                    code, name = str(rr[0]).strip(), str(rr[1]).strip()
                    # 舊版 TPEx 欄序：代號、名稱、本益比、每股股利、股利年度、殖利率、PBR
                    pe = clean_num(rr[2]); dy = clean_num(rr[5]); pbr = clean_num(rr[6])
                else:
                    continue
                if not is_common_stock(code, name):
                    continue
                rows.append({"stock_id": code, "name": name, "per": pe, "pbr": pbr,
                             "dividend_yield": dy, "date": day.isoformat(), "market": "TPEx"})
            if rows:
                return rows
            raise RuntimeError("可解析資料列為0")
        except Exception as e:
            errors.append(f"{url}:{e}")
    raise RuntimeError("TPEx PER 失敗：" + " | ".join(errors))


def _find_exchange_per_near(day: dt.date, market: str, max_back: int = 10) -> Tuple[Optional[dt.date], List[dict]]:
    """從候選日向前尋找最近有資料的交易日。"""
    fetcher = _twse_per_day if market == "TWSE" else _tpex_per_day
    last_err: Optional[Exception] = None
    for i in range(max_back + 1):
        d = day - dt.timedelta(days=i)
        if d.weekday() >= 5:
            continue
        try:
            rows = fetcher(d)
            if rows:
                return d, rows
        except Exception as e:
            last_err = e
    if last_err:
        log.warning("%s 官方 PER 月底附近抓取失敗 %s：%s", market, day, last_err)
    return None, []


def _backfill_official_pe_history(conn: sqlite3.Connection, target: dt.date,
                                  meta: Dict[str, Tuple[str, str]],
                                  cfg: configparser.ConfigParser) -> None:
    """TWSE/TPEx 官方 PER 增量快取。

    舊版每天先刪掉近3年 PER 再重抓 36 個月底，耗時最大。
    v2.4 改為：歷史月份有足夠快取就直接沿用；只補新月份/缺漏月份，
    另抓 target 當日（這本來就是新的每日資料）。
    """
    if not meta:
        return
    years=max(1,cfg.getint("general","valuation_history_years",fallback=3))
    start_day=target-dt.timedelta(days=366*years)
    month_candidates=_month_last_calendar_days(start_day,target)
    min_samples=max(12,cfg.getint("general","valuation_min_pe_samples",fallback=24))
    total_written=0; skipped_months=0; errors=[]
    by_market={"TWSE":set(),"TPEx":set()}
    for sid,(_,market) in meta.items():
        by_market["TPEx" if str(market).lower().startswith("tpex") else "TWSE"].add(str(sid))

    for market,ids in by_market.items():
        if not ids: continue
        source=f"{market}-official-daily-PER"
        threshold=max(40,min(len(ids),int(len(ids)*0.45)))
        print(f"[官方PER增量] {market}：{len(ids)} 檔；歷史月份有快取即跳過",flush=True)
        for pidx,cand in enumerate(month_candidates,1):
            ms=dt.date(cand.year,cand.month,1)
            if cand.month==12: me=dt.date(cand.year+1,1,1)-dt.timedelta(days=1)
            else: me=dt.date(cand.year,cand.month+1,1)-dt.timedelta(days=1)
            existing=int(conn.execute("""SELECT COUNT(DISTINCT stock_id) FROM valuation_pe_history
                                       WHERE source=? AND date>=? AND date<=? AND per>0""",
                                      (source,ms.isoformat(),me.isoformat())).fetchone()[0])
            if existing>=threshold:
                skipped_months+=1; continue
            if pidx==1 or pidx%6==0 or pidx==len(month_candidates):
                print(f"    [{market} PER 補月 {pidx}/{len(month_candidates)}] {cand}",flush=True)
            d,rows=_find_exchange_per_near(cand,market,max_back=10)
            if not d: continue
            vals=[]
            for r in rows:
                sid=str(r["stock_id"]); pe=clean_num(r.get("per"))
                if sid not in ids or pe is None or pe<=0: continue
                vals.append((sid,d.isoformat(),pe,clean_num(r.get("pbr")),clean_num(r.get("dividend_yield")),source))
            if vals:
                conn.executemany("""INSERT OR REPLACE INTO valuation_pe_history
                    (stock_id,date,per,pbr,dividend_yield,source) VALUES (?,?,?,?,?,?)""",vals)
                conn.commit(); total_written+=len(vals)
            time.sleep(0.15)

        # target 當日：若已跑過同一日就直接沿用。
        existing_today=int(conn.execute("SELECT COUNT(DISTINCT stock_id) FROM valuation_pe_history WHERE source=? AND date=? AND per>0",
                                        (source,target.isoformat())).fetchone()[0])
        if existing_today<threshold:
            try:
                rows=(_twse_per_day(target) if market=="TWSE" else _tpex_per_day(target))
                vals=[]
                for r in rows:
                    sid=str(r["stock_id"]); pe=clean_num(r.get("per"))
                    if sid not in ids or pe is None or pe<=0: continue
                    vals.append((sid,target.isoformat(),pe,clean_num(r.get("pbr")),clean_num(r.get("dividend_yield")),source))
                if vals:
                    conn.executemany("""INSERT OR REPLACE INTO valuation_pe_history
                        (stock_id,date,per,pbr,dividend_yield,source) VALUES (?,?,?,?,?,?)""",vals)
                    conn.commit(); total_written+=len(vals)
            except Exception as e:
                errors.append(f"{market} target PER:{e}")

    counts={sid:int(conn.execute("SELECT COUNT(*) FROM valuation_pe_history WHERE stock_id=? AND date>=? AND date<=? AND per>0",
                                 (sid,start_day.isoformat(),target.isoformat())).fetchone()[0]) for sid in meta}
    ok=sum(n>=min_samples for n in counts.values()); short=sum(0<n<min_samples for n in counts.values()); missing=sum(n==0 for n in counts.values())
    msg=(f"增量PER：歷史快取月份跳過 {skipped_months} 次，只補新月/缺漏月＋目標日；{len(meta)}檔：足夠 {ok}、短樣本 {short}、無資料 {missing}；本次新增/更新 {total_written} 筆")
    if errors: msg += "；首批錯誤="+" | ".join(errors[:3])
    record_source(conn,target,"官方PER歷史",target.isoformat(),ok+short>0,len(errors)==0,msg)


def _backfill_valuation_price_history(conn: sqlite3.Connection, target: dt.date, meta: Dict[str, Tuple[str, str]],
                                      cfg: configparser.ConfigParser) -> None:
    """v12：估值直接使用 SQLite 內 TWSE/TPEx 官方股價，不再呼叫 FinMind TaiwanStockPrice。"""
    stock_ids=list(meta.keys())
    if not stock_ids:
        return
    years=max(1,cfg.getint("general","valuation_history_years",fallback=3))
    cutoff=(target-dt.timedelta(days=366*years)).isoformat()
    preferred=max(60,cfg.getint("general","valuation_min_price_samples",fallback=60))
    counts={}
    for sid in stock_ids:
        counts[sid]=int(conn.execute("SELECT COUNT(*) FROM prices WHERE stock_id=? AND date>=? AND date<=?",
                                     (sid,cutoff,target.isoformat())).fetchone()[0])
    enough=sum(n>=preferred for n in counts.values())
    short=sum(30<=n<preferred for n in counts.values())
    missing=sum(n<30 for n in counts.values())
    mn=min(counts.values()) if counts else 0; mx=max(counts.values()) if counts else 0
    record_source(conn,target,"估值股價歷史",target.isoformat(),enough+short>0,missing==0,
                  f"零 FinMind：使用本機 TWSE/TPEx 官方股價；{len(stock_ids)} 檔：足夠 {enough}、短樣本可估 {short}、不足30筆 {missing}；樣本 {mn}~{mx} 筆")


def _quarter_eps_and_availability(period_map: Dict[Tuple[int,int], dict], year: int, quarter: int) -> Optional[Tuple[float, dt.date]]:
    """還原單季 EPS，並回傳該單季在歷史上最早可安全使用的日期。"""
    cur = period_map.get((year, quarter))
    if not cur or cur.get("eps_ytd") is None or not cur.get("availability_date"):
        return None
    try:
        cur_av = dt.date.fromisoformat(str(cur["availability_date"])[:10])
    except Exception:
        return None
    if quarter == 1:
        return float(cur["eps_ytd"]), cur_av
    prev = period_map.get((year, quarter - 1))
    if not prev or prev.get("eps_ytd") is None or not prev.get("availability_date"):
        return None
    try:
        prev_av = dt.date.fromisoformat(str(prev["availability_date"])[:10])
    except Exception:
        return None
    return float(cur["eps_ytd"]) - float(prev["eps_ytd"]), max(cur_av, prev_av)


def _ttm_eps_timeline(conn: sqlite3.Connection, stock_id: str, target: dt.date) -> pd.DataFrame:
    """建立 Point-in-time TTM EPS 時間軸；每個歷史交易日只會看到當時已公開的季報。"""
    hist = pd.read_sql_query(
        """SELECT stock_id,year,quarter,eps_ytd,availability_date,availability_basis,source
           FROM fundamental_income_history
           WHERE stock_id=? AND availability_date IS NOT NULL AND availability_date<=?
           ORDER BY year,quarter""", conn, params=(stock_id, target.isoformat()))
    if hist.empty:
        return pd.DataFrame(columns=["availability_date","ttm_eps","eps_period"])
    pmap: Dict[Tuple[int,int], dict] = {(int(r["year"]), int(r["quarter"])): r for r in hist.to_dict("records")}
    rows: List[dict] = []
    for y, q in sorted(pmap.keys()):
        parts: List[float] = []; avails: List[dt.date] = []; ok = True
        for i in range(4):
            py, pq = _period_prev_n(y, q, i)
            item = _quarter_eps_and_availability(pmap, py, pq)
            if item is None:
                ok = False; break
            eps, av = item
            parts.append(float(eps)); avails.append(av)
        if not ok or len(parts) != 4:
            continue
        rows.append({"availability_date": max(avails), "ttm_eps": sum(parts), "eps_period": f"{y}Q{q}"})
    if not rows:
        return pd.DataFrame(columns=["availability_date","ttm_eps","eps_period"])
    out = pd.DataFrame(rows).sort_values(["availability_date","eps_period"]).drop_duplicates("availability_date", keep="last")
    out["availability_date"] = pd.to_datetime(out["availability_date"])
    return out


def _rebuild_self_pe_history(conn: sqlite3.Connection, stock_id: str, target: dt.date,
                             cfg: configparser.ConfigParser) -> dict:
    """歷史 PE = 歷史收盤價 / 當日 Point-in-time TTM EPS。完全不依賴 FinMind PER。"""
    years = max(1, cfg.getint("general", "valuation_history_years", fallback=3))
    start = target - dt.timedelta(days=366 * years)
    prices = pd.read_sql_query(
        "SELECT date,close FROM prices WHERE stock_id=? AND date>=? AND date<=? ORDER BY date",
        conn, params=(stock_id, start.isoformat(), target.isoformat()))
    if prices.empty:
        return {"samples": 0, "reason": "無歷史股價"}
    timeline = _ttm_eps_timeline(conn, stock_id, target)
    if timeline.empty:
        return {"samples": 0, "reason": "無可用TTM EPS時間軸"}
    prices["date"] = pd.to_datetime(prices["date"])
    prices["close"] = pd.to_numeric(prices["close"], errors="coerce")
    merged = pd.merge_asof(prices.sort_values("date"), timeline.sort_values("availability_date"),
                           left_on="date", right_on="availability_date", direction="backward")
    merged = merged[(merged["close"] > 0) & (pd.to_numeric(merged["ttm_eps"], errors="coerce") > 0)].copy()
    if merged.empty:
        return {"samples": 0, "reason": "歷史期間TTM EPS非正或尚未可得"}
    merged["per"] = merged["close"] / merged["ttm_eps"]
    pe_min = cfg.getfloat("general", "valuation_pe_min", fallback=3.0)
    pe_max = cfg.getfloat("general", "valuation_pe_max", fallback=100.0)
    valid = merged[(merged["per"] >= pe_min) & (merged["per"] <= pe_max)].copy()
    vals = [(stock_id, r.date.date().isoformat(), float(r.per), None, None,
             f"SELF:close/point-in-time-TTM-EPS({r.eps_period})") for r in valid.itertuples()]
    if vals:
        conn.executemany("""INSERT OR REPLACE INTO valuation_pe_history
            (stock_id,date,per,pbr,dividend_yield,source) VALUES (?,?,?,?,?,?)""", vals)
        conn.commit()
    return {"samples": len(valid), "raw_price_days": len(prices), "reason": ""}


def _rebuild_self_pe_for_priority(conn: sqlite3.Connection, target: dt.date, meta: Dict[str, Tuple[str,str]],
                                  cfg: configparser.ConfigParser) -> None:
    if not meta:
        return
    _backfill_valuation_price_history(conn, target, meta, cfg)
    ok = fail = 0; samples = 0; errors: List[str] = []
    for sid in meta:
        try:
            res = _rebuild_self_pe_history(conn, sid, target, cfg)
            if int(res.get("samples", 0)) > 0:
                ok += 1; samples += int(res["samples"])
            else:
                fail += 1; errors.append(f"{sid}:{res.get('reason','無樣本')}")
        except Exception as e:
            fail += 1; errors.append(f"{sid}:{e}")
    msg = f"自行估值PE：優先名單 {len(meta)} 檔；成功 {ok}、不足 {fail}；Point-in-time PE 樣本 {samples} 筆；不呼叫 TaiwanStockPER"
    if errors: msg += "；首批不足=" + " | ".join(errors[:5])
    record_source(conn, target, "自行計算PE歷史", target.isoformat(), ok > 0, fail == 0, msg)


def _pe_band_for_stock(conn: sqlite3.Connection, stock_id: str, target: dt.date, cfg: configparser.ConfigParser) -> dict:
    years = max(1, cfg.getint("general", "valuation_history_years", fallback=3))
    pe_min = cfg.getfloat("general", "valuation_pe_min", fallback=3.0)
    pe_max = cfg.getfloat("general", "valuation_pe_max", fallback=100.0)
    start = (target - dt.timedelta(days=366 * years)).isoformat()
    df = pd.read_sql_query(
        "SELECT date,per,pbr,dividend_yield,source FROM valuation_pe_history WHERE stock_id=? AND date>=? AND date<=? ORDER BY date",
        conn, params=(stock_id, start, target.isoformat()))
    if df.empty:
        return {}
    df["per"] = pd.to_numeric(df["per"], errors="coerce")
    valid = df[(df["per"] >= pe_min) & (df["per"] <= pe_max)]["per"].dropna()
    latest = df.dropna(subset=["per"]).tail(1)
    out = {"市場PER": float(latest.iloc[0]["per"]) if not latest.empty else None}
    min_samples = max(12, cfg.getint("general", "valuation_min_pe_samples", fallback=24))
    if len(valid) >= min_samples:
        out.update({
            "歷史PE中位數": float(valid.median()),
            "歷史PE_25%": float(valid.quantile(0.25)),
            "歷史PE_75%": float(valid.quantile(0.75)),
            "歷史PE樣本數": int(len(valid)),
        })
    return out


def _load_institution_valuation(target: dt.date) -> Dict[str, dict]:
    """選配：若使用者另有合法取得的機構 Forward EPS/目標PE，可放 CSV；沒有就自動走歷史PE備援。"""
    if not INSTITUTION_VALUATION_CSV.exists():
        return {}
    try:
        df = pd.read_csv(INSTITUTION_VALUATION_CSV, dtype={"stock_id": str}, encoding="utf-8-sig")
    except Exception:
        return {}
    out: Dict[str, dict] = {}
    for _, r in df.iterrows():
        sid = str(r.get("stock_id") or "").strip()
        fwd = clean_num(r.get("forward_eps")); pe = clean_num(r.get("target_pe"))
        if not sid or fwd is None or pe is None or fwd <= 0 or pe <= 0:
            continue
        asof = str(r.get("as_of_date") or "").strip()
        valid_until = str(r.get("valid_until") or "").strip()
        try:
            if asof and dt.date.fromisoformat(asof) > target: continue
            if valid_until and dt.date.fromisoformat(valid_until) < target: continue
        except Exception:
            continue
        out[sid] = {
            "Forward EPS": float(fwd), "機構目標PE": float(pe),
            "機構估值來源": str(r.get("source") or "user-supplied institution valuation"),
            "機構估值日期": asof or None,
        }
    return out


def _valuation_status(discount_pct: Optional[float], cfg: configparser.ConfigParser) -> str:
    if discount_pct is None or pd.isna(discount_pct):
        return "資料不足"
    deep = cfg.getfloat("general", "valuation_deep_discount_pct", fallback=20.0)
    buy = cfg.getfloat("general", "valuation_buy_discount_pct", fallback=10.0)
    if discount_pct >= deep: return "明顯低估"
    if discount_pct >= buy: return "買進區"
    if discount_pct >= 0: return "合理偏低"
    if discount_pct >= -10: return "合理區"
    return "偏貴"


def _market_snapshot(conn: sqlite3.Connection, target: dt.date) -> pd.DataFrame:
    """目標日全部上市／上櫃普通股，估值母體不受鴨嘴或基本面條件限制。"""
    df = pd.read_sql_query(
        "SELECT date,stock_id,name,market,close FROM prices WHERE date=? ORDER BY market,stock_id",
        conn, params=(target.isoformat(),))
    if df.empty:
        return df
    df["date"] = pd.to_datetime(df["date"])
    df["close"] = pd.to_numeric(df["close"], errors="coerce")
    df = df[df.apply(lambda r: is_common_stock(str(r.get("stock_id") or ""), str(r.get("name") or "")), axis=1)].copy()
    return df


def _enrich_valuation_rows(conn: sqlite3.Connection, target: dt.date, base: pd.DataFrame,
                           pe_map: Dict[str, dict], institution: Dict[str, dict],
                           cfg: configparser.ConfigParser, eps_ids: Optional[set] = None) -> pd.DataFrame:
    """估值 EPS 優先序：機構 Forward EPS > Yahoo TTM EPS > MOPS TTM EPS > 交易所隱含 EPS。"""
    out = base.copy()
    cols = ["最新單季EPS","前一季EPS","去年同期EPS","TTM EPS","EPS年增率%","EPS資料期","EPS來源","EPS觀測日",
            "Yahoo最新單季EPS","Yahoo TTM EPS","MOPS TTM EPS","交易所隱含EPS",
            "市場PER","歷史PE中位數","歷史PE_25%","歷史PE_75%","歷史PE樣本數",
            "Forward EPS","機構目標PE","機構估值來源","機構估值日期",
            "估值採用EPS","估值採用PE","合理價","折價率%","估值方法","估值狀態"]
    for c in cols:
        if c not in out.columns:
            out[c] = None
    eps_ids = eps_ids or set()
    official_cache: Dict[str, dict] = {}
    yahoo_cache: Dict[str, dict] = {}
    for i, r in out.iterrows():
        sid = str(r["stock_id"]); close = clean_num(r.get("close"))
        official = {}
        if sid in eps_ids:
            if sid not in official_cache:
                official_cache[sid] = _eps_metrics_for_stock(conn, sid, target)
            official = official_cache[sid]
        if sid not in yahoo_cache:
            yahoo_cache[sid] = _yfinance_eps_metrics_for_stock(conn, sid, target)
        yahoo = yahoo_cache[sid]

        official_ttm = clean_num(official.get("TTM EPS"))
        yahoo_ttm = clean_num(yahoo.get("TTM EPS"))
        out.at[i, "MOPS TTM EPS"] = official_ttm
        out.at[i, "Yahoo TTM EPS"] = yahoo_ttm
        out.at[i, "Yahoo最新單季EPS"] = clean_num(yahoo.get("最新單季EPS"))

        # 通用 EPS 欄位顯示實際採用的季度來源；Yahoo 有完整四季時優先。
        em = yahoo if yahoo_ttm is not None else official
        for k in ["最新單季EPS","前一季EPS","去年同期EPS","TTM EPS","EPS年增率%","EPS資料期","EPS來源","EPS觀測日"]:
            if k in em:
                out.at[i, k] = em.get(k)

        pm = pe_map.get(sid, {})
        for k, v in pm.items():
            out.at[i, k] = v
        inst = institution.get(sid, {})
        for k, v in inst.items():
            out.at[i, k] = v

        ttm = clean_num(out.at[i, "TTM EPS"])
        market_pe = clean_num(out.at[i, "市場PER"])
        implied_eps = (close / market_pe) if close is not None and market_pe is not None and market_pe > 0 else None
        out.at[i, "交易所隱含EPS"] = implied_eps
        if market_pe is None and close is not None and ttm is not None and ttm > 0:
            market_pe = close / ttm
            out.at[i, "市場PER"] = market_pe

        fair = used_eps = used_pe = None; method = ""
        if inst:
            used_eps = clean_num(inst.get("Forward EPS")); used_pe = clean_num(inst.get("機構目標PE"))
            if used_eps is not None and used_pe is not None and used_eps > 0 and used_pe > 0:
                fair = used_eps * used_pe; method = "機構Forward EPS×目標PE"
        if fair is None:
            med = clean_num(pm.get("歷史PE中位數"))
            if yahoo_ttm is not None and yahoo_ttm > 0:
                base_eps = yahoo_ttm; eps_method = "Yahoo TTM EPS"
            elif official_ttm is not None and official_ttm > 0:
                base_eps = official_ttm; eps_method = "MOPS TTM EPS"
            else:
                base_eps = implied_eps; eps_method = "交易所隱含EPS"
            if base_eps is not None and base_eps > 0 and med is not None and med > 0:
                used_eps, used_pe = base_eps, med
                fair = base_eps * med
                method = f"{eps_method}×官方歷史PE中位數"
        discount = ((fair - close) / fair * 100) if fair and close is not None and fair > 0 else None
        out.at[i, "估值採用EPS"] = used_eps
        out.at[i, "估值採用PE"] = used_pe
        out.at[i, "合理價"] = fair
        out.at[i, "折價率%"] = discount
        out.at[i, "估值方法"] = method or "資料不足"
        out.at[i, "估值狀態"] = _valuation_status(discount, cfg)
    return out

def build_valuation_filters(conn: sqlite3.Connection, target: dt.date, current: pd.DataFrame,
                            both: pd.DataFrame, cfg: configparser.ConfigParser,
                            preduck: Optional[pd.DataFrame] = None) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    v21：估值母體仍為全部上市＋上櫃普通股。
    先用官方 PER / MOPS 建立全市場初估，再對「接近買進區＋鴨嘴＋預備鴨嘴」優先抓 Yahoo 季度 EPS，最後重算合理價。
    """
    market = _market_snapshot(conn, target)
    if market.empty:
        return current.copy(), both.copy(), market.copy(), market.copy()

    all_ids = market["stock_id"].astype(str).tolist()
    meta: Dict[str, Tuple[str,str]] = {
        str(r["stock_id"]):(str(r.get("name") or ""), str(r.get("market") or "TWSE"))
        for _, r in market.iterrows()
    }
    _backfill_official_pe_history(conn, target, meta, cfg)
    institution = _load_institution_valuation(target)
    pe_map = {sid: _pe_band_for_stock(conn, sid, target, cfg) for sid in all_ids}

    eps_ids = {str(r[0]) for r in conn.execute(
        "SELECT DISTINCT stock_id FROM fundamental_income_history WHERE availability_date IS NOT NULL AND availability_date<=?",
        (target.isoformat(),)).fetchall()}

    # 第一輪：不必等 Yahoo，先算官方版，用來挑 Yahoo EPS 優先名單。
    market_pre = _enrich_valuation_rows(conn, target, market, pe_map, institution, cfg, eps_ids)
    buy_min = cfg.getfloat("general", "valuation_buy_discount_pct", fallback=10.0)
    buffer_pct = max(0.0, cfg.getfloat("general", "yfinance_eps_candidate_buffer_pct", fallback=10.0))
    threshold = buy_min - buffer_pct
    pre_discount = pd.to_numeric(market_pre.get("折價率%", pd.Series(index=market_pre.index, dtype=float)), errors="coerce")
    near_ids = market_pre.loc[pre_discount >= threshold, "stock_id"].astype(str).tolist()
    duck_ids = current["stock_id"].astype(str).tolist() if not current.empty and "stock_id" in current.columns else []
    preduck_ids = preduck["stock_id"].astype(str).tolist() if preduck is not None and not preduck.empty and "stock_id" in preduck.columns else []
    # 正式鴨嘴、預備鴨嘴優先，再依官方初估折價率由高到低補候選。
    sorted_pre = market_pre.assign(_d=pre_discount).sort_values("_d", ascending=False, na_position="last")
    near_sorted = [str(x) for x in sorted_pre.loc[sorted_pre["stock_id"].astype(str).isin(set(near_ids)), "stock_id"].tolist()]
    priority_ids = list(dict.fromkeys(duck_ids + preduck_ids + near_sorted))
    _refresh_yfinance_eps(conn, target, meta, priority_ids, cfg)

    # 第二輪：Yahoo 有可用 PIT 快照者改用 Yahoo TTM EPS，否則自動退回 MOPS / 交易所隱含 EPS。
    market_val = _enrich_valuation_rows(conn, target, market, pe_map, institution, cfg, eps_ids)

    market_val["是否鴨嘴"] = False
    market_val["鴨嘴狀態"] = ""
    market_val["月線乖離率%"] = None
    market_val["培育中心類別"] = ""
    if not current.empty:
        tag_cols = [c for c in ["stock_id","狀態","bias20_pct","培育中心類別"] if c in current.columns]
        tags = current[tag_cols].drop_duplicates("stock_id").copy()
        tags["是否鴨嘴"] = True
        tags = tags.rename(columns={"狀態":"鴨嘴狀態","bias20_pct":"月線乖離率%"})
        market_val = market_val.drop(columns=["是否鴨嘴","鴨嘴狀態","月線乖離率%","培育中心類別"], errors="ignore")
        market_val = market_val.merge(tags, on="stock_id", how="left")
        market_val["是否鴨嘴"] = market_val["是否鴨嘴"].fillna(False).astype(bool)
        market_val["鴨嘴狀態"] = market_val["鴨嘴狀態"].fillna("")
        market_val["培育中心類別"] = market_val["培育中心類別"].fillna("")

    market_val = market_val.sort_values(["折價率%","market","stock_id"], ascending=[False,True,True], na_position="last")

    current_out = current.copy()
    valuation_cols = ["stock_id","最新單季EPS","前一季EPS","去年同期EPS","TTM EPS","EPS年增率%","EPS資料期","EPS來源","EPS觀測日",
                      "Yahoo最新單季EPS","Yahoo TTM EPS","MOPS TTM EPS","交易所隱含EPS",
                      "市場PER","歷史PE中位數","歷史PE_25%","歷史PE_75%","歷史PE樣本數",
                      "Forward EPS","機構目標PE","機構估值來源","機構估值日期","估值採用EPS","估值採用PE",
                      "合理價","折價率%","估值方法","估值狀態"]
    vmeta = market_val[[c for c in valuation_cols if c in market_val.columns]].drop_duplicates("stock_id")
    if not current_out.empty:
        current_out = current_out.drop(columns=[c for c in valuation_cols if c != "stock_id" and c in current_out.columns], errors="ignore")
        current_out = current_out.merge(vmeta, on="stock_id", how="left")

    both_enriched = current_out[current_out.get("三率三升", False) & current_out.get("營收三增", False)].copy() if not current_out.empty else both.copy()
    if not both_enriched.empty:
        sort_cols=[c for c in ["折價率%","bias20_pct"] if c in both_enriched.columns]
        if sort_cols:
            both_enriched=both_enriched.sort_values(sort_cols,ascending=[False]*len(sort_cols),na_position="last")

    buy = market_val[(pd.to_numeric(market_val["折價率%"], errors="coerce") >= buy_min) &
                     (pd.to_numeric(market_val["估值採用EPS"], errors="coerce") > 0)].copy()
    buy = buy.sort_values(["折價率%","stock_id"], ascending=[False,True], na_position="last")
    return current_out, both_enriched, market_val, buy

def _flatten_cols(cols) -> List[str]:
    out=[]
    for c in cols:
        if isinstance(c, tuple):
            txt=" ".join(str(x) for x in c if str(x).lower()!='nan')
        else:
            txt=str(c)
        out.append(re.sub(r"\s+", "", txt))
    return out



def _csv_col(cols: Iterable[Any], needles: List[str], exact: bool=False) -> Optional[Any]:
    for c in cols:
        t=re.sub(r"\s+","",str(c))
        for n in needles:
            nn=re.sub(r"\s+","",n)
            if (t==nn) if exact else (nn in t):
                return c
    return None


def _decode_taifex_csv(raw: bytes) -> str:
    for enc in ("ms950","cp950","big5","utf-8-sig","utf-8"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("cp950",errors="replace")


def _parse_taifex_csv(text: str, start: dt.date, end: dt.date) -> List[dict]:
    # TAIFEX CSV 偶爾行尾多一個逗號，index_col=False 可避免第一欄被吃成 index。
    df=pd.read_csv(io.StringIO(text),index_col=False)
    if df.empty:
        return []
    df.columns=[str(c).strip() for c in df.columns]
    date_col=_csv_col(df.columns,["交易日期","日期"])
    contract_col=_csv_col(df.columns,["契約"],exact=True)
    month_col=_csv_col(df.columns,["到期月份(週別)","到期月份"])
    session_col=_csv_col(df.columns,["交易時段"])
    open_col=_csv_col(df.columns,["開盤價"])
    high_col=_csv_col(df.columns,["最高價"])
    low_col=_csv_col(df.columns,["最低價"])
    close_col=_csv_col(df.columns,["最後成交價","收盤價"])
    settle_col=_csv_col(df.columns,["結算價"])
    vol_col=_csv_col(df.columns,["成交量","一般交易時段成交量"])
    oi_col=_csv_col(df.columns,["未沖銷契約數","未沖銷契約量"])
    if not contract_col or not month_col or not close_col:
        raise RuntimeError(f"TAIFEX CSV 欄位無法辨識: {df.columns.tolist()}")
    sub=df[df[contract_col].astype(str).str.strip().eq("TX")].copy()
    if session_col:
        # 每日下載檔的日盤通常標示「一般」；若欄位沒有此值則不強制過濾。
        ss=sub[session_col].astype(str).str.strip()
        if ss.str.contains("一般").any():
            sub=sub[ss.str.contains("一般")]
    # 排除價差契約，只留 YYYYMM 月契約。
    sub["_month_raw"]=sub[month_col].astype(str).str.strip()
    sub=sub[sub["_month_raw"].str.fullmatch(r"\d{6}")]
    if sub.empty:
        return []
    def parse_day(v):
        digits=re.sub(r"\D","",str(v))
        if len(digits)==8:
            try:return dt.date(int(digits[:4]),int(digits[4:6]),int(digits[6:8]))
            except:return None
        if len(digits)==7:
            try:return dt.date(int(digits[:3])+1911,int(digits[3:5]),int(digits[5:7]))
            except:return None
        return None
    if date_col:
        sub["_date"]=sub[date_col].map(parse_day)
    else:
        if start!=end:
            raise RuntimeError("TAIFEX CSV 缺交易日期欄，無法解析多日")
        sub["_date"]=start
    sub=sub[sub["_date"].notna()]
    sub=sub[(sub["_date"]>=start)&(sub["_date"]<=end)]
    out=[]
    for day,g in sub.groupby("_date"):
        # 近月：選當日尚未到期的最小 YYYYMM；若資料異常則退回成交量最大者。
        g=g.copy(); g["_ym"]=pd.to_numeric(g["_month_raw"],errors="coerce")
        current_ym=day.year*100+day.month
        valid=g[g["_ym"]>=current_ym]
        if valid.empty: valid=g
        valid=valid.sort_values("_ym")
        rr=valid.iloc[0]
        def num(col): return clean_num(rr[col]) if col is not None else None
        close=num(close_col); settle=num(settle_col)
        if close is None and settle is None:
            continue
        out.append({"date":day.isoformat(),"contract":str(rr[month_col]).strip(),"open":num(open_col),"high":num(high_col),
                    "low":num(low_col),"close":close,"settlement":settle,"volume":num(vol_col),"open_interest":num(oi_col)})
    return sorted(out,key=lambda x:x["date"])


def _fetch_taifex_csv_range(start: dt.date, end: dt.date) -> List[dict]:
    """TAIFEX 官方『期貨每日交易行情下載』CSV。官方頁面限制單次區間不超過一個月。"""
    if end<start:
        return []
    endpoints=[
        "https://www.taifex.com.tw/cht/3/futDataDown",
        "https://www.bq888.taifex.com.tw/cht/3/futDataDown",
    ]
    form={"down_type":"1","commodity_id":"TX","queryStartDate":start.strftime("%Y/%m/%d"),
          "queryEndDate":end.strftime("%Y/%m/%d"),"MarketCode":"0"}
    errs=[]
    for url in endpoints:
        for attempt in range(3):
            try:
                r=SESSION.post(url,data=form,timeout=30,headers={"User-Agent":UA,"Accept":"text/csv,text/plain,*/*",
                                                                  "Referer":"https://www.taifex.com.tw/cht/3/futDailyMarketView",
                                                                  "Connection":"close"})
                r.raise_for_status()
                text=_decode_taifex_csv(r.content)
                rows=_parse_taifex_csv(text,start,end)
                if rows:
                    return rows
                # 非交易日全區間才可能合理為空；有平日卻空就當錯誤重試。
                if all((start+dt.timedelta(days=i)).weekday()>=5 for i in range((end-start).days+1)):
                    return []
                raise RuntimeError("CSV 回傳但未解析到 TX 一般交易時段")
            except Exception as e:
                errs.append(f"{url.split('/')[2]}/{attempt+1}:{type(e).__name__}:{str(e)[:100]}")
                if attempt<2: time.sleep(1.0*(attempt+1))
    raise RuntimeError("TAIFEX CSV下載失敗："+" | ".join(errs[-6:]))


def _upsert_futures_rows(conn: sqlite3.Connection, rows: List[dict]) -> int:
    if not rows:return 0
    vals=[(r['date'],r['contract'],r.get('open'),r.get('high'),r.get('low'),r.get('close'),r.get('settlement'),r.get('volume'),r.get('open_interest')) for r in rows]
    conn.executemany("""INSERT OR REPLACE INTO futures(date,contract,open,high,low,close,settlement,volume,open_interest)
                        VALUES (?,?,?,?,?,?,?,?,?)""",vals)
    conn.commit(); return len(vals)


def fetch_futures_taifex(conn: sqlite3.Connection, target: dt.date, cfg: configparser.ConfigParser) -> pd.DataFrame:
    """v25：不用 HTML 表單逐日查詢，改走 TAIFEX 官方 CSV 下載，一次抓一段日期。"""
    existing=int(conn.execute("SELECT COUNT(DISTINCT date) FROM futures WHERE date<=?",(target.isoformat(),)).fetchone()[0])
    has_target=bool(conn.execute("SELECT 1 FROM futures WHERE date=? LIMIT 1",(target.isoformat(),)).fetchone())
    ok_ranges=fail_ranges=0; added=0; errs=[]
    # 先保證目標日。
    if not has_target:
        try:
            rows=_fetch_taifex_csv_range(target,target); added+=_upsert_futures_rows(conn,rows); ok_ranges+=1
        except Exception as e:
            fail_ranges+=1; errs.append(f"target:{e}")
    # MA60 不夠時，以 <=28 天區段往前補，遠比逐日 76 次穩定。
    current=int(conn.execute("SELECT COUNT(DISTINCT date) FROM futures WHERE date<=?",(target.isoformat(),)).fetchone()[0])
    if current<60:
        end=target-dt.timedelta(days=1)
        floor=target-dt.timedelta(days=180)
        while end>=floor and current<70:
            start=max(floor,end-dt.timedelta(days=27))
            try:
                rows=_fetch_taifex_csv_range(start,end); added+=_upsert_futures_rows(conn,rows); ok_ranges+=1
            except Exception as e:
                fail_ranges+=1; errs.append(f"{start}~{end}:{e}")
            current=int(conn.execute("SELECT COUNT(DISTINCT date) FROM futures WHERE date<=?",(target.isoformat(),)).fetchone()[0])
            print(f"[TAIFEX CSV] {start}~{end}；目前 {current} 個交易日",flush=True)
            end=start-dt.timedelta(days=1)
            time.sleep(0.25)
    df=pd.read_sql_query("SELECT * FROM futures WHERE date<=? ORDER BY date",conn,params=(target.isoformat(),))
    has=not df.empty and (df['date'].astype(str)==target.isoformat()).any()
    final_days=int(df['date'].nunique()) if not df.empty else 0
    history_ready=final_days>=60
    msg=(f"TAIFEX 官方 CSV下載；原有 {existing} 日；成功區段 {ok_ranges}、失敗區段 {fail_ranges}；新增/更新 {added} 筆；"
         f"目標日={'有' if has else '無'}；目前歷史 {final_days} 日（MA60需60日）")
    if errs: msg += "；首批錯誤="+" | ".join(errs[:2])
    record_source(conn,target,"TAIFEX-TX",target.isoformat() if has else None,has,has and history_ready,msg)
    return df

def futures_analysis(conn: sqlite3.Connection, target: dt.date) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = pd.read_sql_query(
        "SELECT * FROM futures WHERE date<=? ORDER BY date", conn, params=(target.isoformat(),)
    )
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()
    df["date"] = pd.to_datetime(df["date"])
    df["close"] = pd.to_numeric(df["close"], errors="coerce")
    df["MA20"] = df["close"].rolling(20, min_periods=20).mean()
    df["MA60"] = df["close"].rolling(60, min_periods=60).mean()
    df["MA20_change"] = df["MA20"].diff()
    df["MA60_change"] = df["MA60"].diff()
    df["spread"] = df["MA20"] - df["MA60"]
    df["spread_change"] = df["spread"].diff()
    df["duck"] = (
        (df["close"] > df["MA20"]) & (df["close"] > df["MA60"]) &
        (df["MA20"] > df["MA60"]) & (df["MA20_change"] > 0) &
        (df["MA60_change"] > 0) & (df["spread_change"] > 0)
    )
    latest = df[df["date"] == pd.Timestamp(target)].copy()
    if latest.empty:
        latest = df.tail(1).copy()

    # 近 120 個交易日最高/最低與費波反彈位（由低往高的回復架構）
    w = df.tail(120)
    swing_high = float(pd.to_numeric(w["high"], errors="coerce").max())
    swing_low = float(pd.to_numeric(w["low"], errors="coerce").min())
    diff = swing_high - swing_low
    fib = pd.DataFrame({
        "項目": ["近120日低點","23.6%","38.2%","50%","61.8%","78.6%","近120日高點"],
        "點位": [swing_low, swing_low+diff*0.236, swing_low+diff*0.382, swing_low+diff*0.5, swing_low+diff*0.618, swing_low+diff*0.786, swing_high]
    })
    return latest, fib


def humanize(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    x = df.copy()
    rename = {
        "date":"日期","stock_id":"代號","name":"名稱","market":"市場","close":"收盤",
        "MA20":"MA20","MA60":"MA60","MA20_change":"MA20較前日","MA60_change":"MA60較前日",
        "spread":"MA20-MA60","spread_change":"開口較前日","bias20_pct":"月線乖離率%",
    }
    x = x.rename(columns=rename)
    if "日期" in x:
        x["日期"] = pd.to_datetime(x["日期"]).dt.strftime("%Y-%m-%d")
    return x



def build_integrated_master(conn: sqlite3.Connection, target: dt.date, tech_all: pd.DataFrame,
                            candidate_enriched: pd.DataFrame, market_valuation: pd.DataFrame,
                            valuation_buy: pd.DataFrame, cfg: configparser.ConfigParser) -> pd.DataFrame:
    """v23 主工作表：一列一檔股票，把技術面、預備鴨嘴、營收、三率、EPS、估值放在同一列。

    - 技術面：全市場都有。
    - 月營收：MOPS 歷史頁本來就是全市場批次下載，因此直接把全市場最新可得月份合併進來。
    - 三率：為控制第一次回補量，優先對「正式鴨嘴＋預備鴨嘴」補足歷史季報；其他股票若資料庫已有資料也會顯示。
    - 估值：沿用全市場估值母體。
    """
    if tech_all is None or tech_all.empty:
        return market_valuation.copy() if market_valuation is not None else pd.DataFrame()

    tech_cols=[
        "date","stock_id","name","market","close","MA20","MA60","MA20_change","MA60_change","spread","spread_change","bias20_pct",
        "首次過熱日期","狀態","鴨嘴階段","鴨嘴完成條件數","鴨嘴完成度%","距MA20%","MA20距MA60%","預備鴨嘴狀態","鴨嘴未達條件",
        "duck","new","exit","preduck","cond_price","cond_cross","cond_ma20_up","cond_ma60_up","cond_spread_up",
    ]
    m=tech_all[[c for c in tech_cols if c in tech_all.columns]].copy()
    m["正式鴨嘴"] = m.get("duck", False).fillna(False).astype(bool) if isinstance(m.get("duck"), pd.Series) else False
    m["今日新進"] = m.get("new", False).fillna(False).astype(bool) if isinstance(m.get("new"), pd.Series) else False
    m["今日退出"] = m.get("exit", False).fillna(False).astype(bool) if isinstance(m.get("exit"), pd.Series) else False
    m["預備鴨嘴"] = m.get("preduck", False).fillna(False).astype(bool) if isinstance(m.get("preduck"), pd.Series) else False
    cond_rename={
        "cond_price":"條件1_收盤站上MA20MA60","cond_cross":"條件2_MA20大於MA60","cond_ma20_up":"條件3_MA20上揚",
        "cond_ma60_up":"條件4_MA60上揚","cond_spread_up":"條件5_開口擴大",
    }
    m=m.rename(columns=cond_rename)
    m=m.drop(columns=[c for c in ["duck","new","exit","preduck"] if c in m.columns],errors="ignore")

    # 三率資料：候選股優先補洞；只合併三率欄，營收另由全市場歷史月營收統一帶入。
    fund_cols=[
        "stock_id","毛利率本季%","毛利率上季%","營業利益率本季%","營業利益率上季%","稅後純益率本季%","稅後純益率上季%",
        "毛差","營差","淨差","三率三升","三率資料期","三率快照日","三率可得日依據","三率來源","基本面資料狀態","基本面缺漏原因","基本面所需季度","營收資料狀態",
    ]
    if candidate_enriched is not None and not candidate_enriched.empty:
        fm=candidate_enriched[[c for c in fund_cols if c in candidate_enriched.columns]].drop_duplicates("stock_id")
        m=m.merge(fm,on="stock_id",how="left")

    # 月營收歷史是批次全市場資料，主表一律取 target 當日以前最新可得月份。
    rev=_latest_revenue_history_asof(conn,target,None)
    revmin=cfg.getfloat("general","revenue_growth_min_pct",fallback=0.0)
    if not rev.empty:
        rv=rev[[c for c in ["stock_id","month_key","revenue_mom_pct","revenue_yoy_pct","revenue_cum_yoy_pct","observed_date","availability_basis","source"] if c in rev.columns]].copy()
        rv=rv.rename(columns={
            "revenue_mom_pct":"營收月增%","revenue_yoy_pct":"營收年增%","revenue_cum_yoy_pct":"累計營收增%",
            "observed_date":"營收快照日","availability_basis":"營收可得日依據","source":"營收來源",
        })
        if "month_key" in rv.columns:
            rv["營收資料月"]=rv["month_key"].apply(lambda x: f"{int(x)//100}-{int(x)%100:02d}" if pd.notna(x) else None)
            rv=rv.drop(columns=["month_key"])
        rv["營收三增"]=(pd.to_numeric(rv.get("營收月增%"),errors="coerce")>revmin) & (pd.to_numeric(rv.get("營收年增%"),errors="coerce")>revmin) & (pd.to_numeric(rv.get("累計營收增%"),errors="coerce")>revmin)
        m=m.merge(rv.drop_duplicates("stock_id"),on="stock_id",how="left")

    # 全市場估值欄位直接併進主表；避免重複識別/技術標籤欄。
    if market_valuation is not None and not market_valuation.empty:
        skip={"name","market","close","是否鴨嘴","鴨嘴狀態","月線乖離率%","培育中心類別","預備鴨嘴狀態","預備鴨嘴尚缺條件"}
        val_cols=["stock_id"]+[c for c in market_valuation.columns if c not in skip and c!="stock_id"]
        vm=market_valuation[val_cols].drop_duplicates("stock_id")
        m=m.merge(vm,on="stock_id",how="left")

    buy_ids=set(valuation_buy["stock_id"].astype(str).tolist()) if valuation_buy is not None and not valuation_buy.empty and "stock_id" in valuation_buy.columns else set()
    m["估值買進候選"]=m["stock_id"].astype(str).isin(buy_ids)

    # 狀態欄不把「沒有資料」誤判成「不符合」。
    if "三率資料期" in m.columns:
        has3=m["三率資料期"].notna() & (m["三率資料期"].astype(str).str.len()>0)
        trflag=m.get("三率三升",pd.Series(False,index=m.index)).fillna(False).astype(bool)
        base_status=m.get("基本面資料狀態",pd.Series(None,index=m.index)).fillna("").astype(str)
        vals=[]
        for h,f,bs in zip(has3,trflag,base_status):
            if bs=="三率不適用": vals.append("三率不適用")
            elif bs=="上市時間不足/不適用": vals.append("上市時間不足")
            elif bs=="缺資料": vals.append("資料缺漏")
            else: vals.append("三率三升" if h and f else ("未三升" if h else "資料不足/未回補"))
        m["三率狀態"]=vals
    else:
        m["三率狀態"]="資料不足/未回補"
    hasrev=m.get("營收資料月",pd.Series(None,index=m.index)).notna()
    rvflag=m.get("營收三增",pd.Series(False,index=m.index)).fillna(False).astype(bool)
    m["營收狀態"]=["營收三增" if h and f else ("未三增" if h else "資料不足") for h,f in zip(hasrev,rvflag)]

    def label_row(r: pd.Series) -> str:
        parts=[]
        stage=str(r.get("鴨嘴階段") or "")
        if stage and stage!="未符合": parts.append(stage)
        if str(r.get("三率狀態") or "")=="三率三升": parts.append("三率三升")
        if str(r.get("營收狀態") or "")=="營收三增": parts.append("營收三增")
        if bool(r.get("估值買進候選",False)): parts.append("估值買進候選")
        return "｜".join(parts)
    m["整合觀察標籤"]=m.apply(label_row,axis=1)

    # 主表欄位排序：識別→鴨嘴技術→基本面→營收→EPS/估值。
    first=[
        "date","stock_id","name","market","整合觀察標籤","鴨嘴階段","正式鴨嘴","預備鴨嘴","預備鴨嘴狀態","今日新進","今日退出",
        "鴨嘴完成條件數","鴨嘴完成度%","鴨嘴未達條件","close","MA20","MA60","MA20_change","MA60_change","spread","spread_change","bias20_pct","首次過熱日期",
        "條件1_收盤站上MA20MA60","條件2_MA20大於MA60","條件3_MA20上揚","條件4_MA60上揚","條件5_開口擴大","距MA20%","MA20距MA60%",
        "三率狀態","基本面資料狀態","基本面缺漏原因","基本面所需季度","毛利率本季%","毛利率上季%","營業利益率本季%","營業利益率上季%","稅後純益率本季%","稅後純益率上季%","毛差","營差","淨差","三率資料期","三率快照日",
        "營收狀態","營收資料狀態","營收月增%","營收年增%","累計營收增%","營收資料月","營收快照日",
        "最新單季EPS","前一季EPS","去年同期EPS","TTM EPS","EPS年增率%","EPS資料期","EPS來源","Yahoo最新單季EPS","Yahoo TTM EPS","MOPS TTM EPS","交易所隱含EPS",
        "市場PER","歷史PE中位數","歷史PE_25%","歷史PE_75%","歷史PE樣本數","Forward EPS","機構目標PE","估值採用EPS","估值採用PE","合理價","折價率%","估值方法","估值狀態","估值買進候選",
    ]
    rest=[c for c in m.columns if c not in first and c not in ["三率三升","營收三增"]]
    m=m[[c for c in first if c in m.columns]+rest]
    stage_rank={"新進／不建議追價／過熱":0,"新進":0,"A級：臨門一腳（4/5）":1,"B級：接近形成（3/5）":2,"持續符合":3,"今日退出":4,"未符合":5}
    m["__stage_rank"]=m["鴨嘴階段"].map(stage_rank).fillna(m["鴨嘴階段"].astype(str).apply(lambda x: 3 if x.startswith("持續符合") else 5))
    m=m.sort_values(["__stage_rank","鴨嘴完成條件數","折價率%" if "折價率%" in m.columns else "stock_id","stock_id"],ascending=[True,False,False,True],na_position="last")
    return m.drop(columns=["__stage_rank"],errors="ignore")


def export_excel(conn: sqlite3.Connection, target: dt.date, master: pd.DataFrame, current: pd.DataFrame, new: pd.DataFrame, exits: pd.DataFrame, hot: pd.DataFrame, preduck: pd.DataFrame, fut_latest: pd.DataFrame, fib: pd.DataFrame, three_rates: pd.DataFrame, three_revenue: pd.DataFrame, both: pd.DataFrame, market_valuation: pd.DataFrame, valuation_buy: pd.DataFrame) -> Path:
    out = OUT_DIR / f"{target.strftime('%Y%m%d')}_鴨嘴篩選結果.xlsx"
    status = pd.read_sql_query(
        "SELECT source AS 資料來源,data_date AS 資料日期,success AS 成功,complete AS 完整,message AS 訊息,checked_at AS 檢查時間 FROM source_status WHERE run_date=? ORDER BY source",
        conn, params=(target.isoformat(),)
    )
    if not status.empty:
        def _status_label(r):
            if int(r["成功"]) == 0:
                return "需處理"
            if int(r["完整"]) == 1:
                return "正常"
            if str(r["資料來源"]) in ("官方財報歷史","官方月營收歷史"):
                return "歷史深度不足"
            return "部分缺漏"
        status["判讀"] = status.apply(_status_label, axis=1)
        # 把判讀放在前面，避免看到「成功/完整」兩個布林值還要自己猜。
        status = status[["資料來源","資料日期","判讀","成功","完整","訊息","檢查時間"]]
        status["成功"] = status["成功"].map({1:"是",0:"否"})
        status["完整"] = status["完整"].map({1:"是",0:"否"})

    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        book = writer.book
        header_fmt = book.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center"})
        header_id_fmt = book.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "align": "center"})
        header_tech_fmt = book.add_format({"bold": True, "bg_color": "#BDD7EE", "border": 1, "align": "center"})
        header_fund_fmt = book.add_format({"bold": True, "bg_color": "#C6E0B4", "border": 1, "align": "center"})
        header_rev_fmt = book.add_format({"bold": True, "bg_color": "#FCE4D6", "border": 1, "align": "center"})
        header_val_fmt = book.add_format({"bold": True, "bg_color": "#E4DFEC", "border": 1, "align": "center"})
        title_fmt = book.add_format({"bold": True, "font_size": 14})
        warn_fmt = book.add_format({"bg_color": "#FFF2CC"})
        hot_fmt = book.add_format({"bg_color": "#F4CCCC"})
        good_fmt = book.add_format({"bg_color": "#E2F0D9"})
        pct_fmt = book.add_format({"num_format": "0.00"})

        def export_view(d: pd.DataFrame) -> pd.DataFrame:
            x = humanize(d)
            # 內部布林欄位不直接顯示；用「培育中心類別」呈現。
            for c in ["三率三升", "營收三增"]:
                if c in x.columns:
                    x = x.drop(columns=[c])
            return x

        sheets = [
            ("全市場整合篩選", export_view(master)),
            ("即將可能符合", export_view(preduck)),
            ("今日新進", export_view(new)),
            ("今日退出", export_view(exits)),
            ("全部符合", export_view(current)),
            ("過熱", export_view(hot)),
            ("鴨嘴×三率三升", export_view(three_rates)),
            ("鴨嘴×營收三增", export_view(three_revenue)),
            ("鴨嘴×兩者皆有", export_view(both)),
            ("全市場估值", export_view(market_valuation)),
            ("估值買進候選", export_view(valuation_buy)),
        ]
        for sname, data in sheets:
            data.to_excel(writer, sheet_name=sname, index=False, startrow=1)
            ws = writer.sheets[sname]
            ws.write(0, 0, f"{target} {sname}", title_fmt)
            ws.freeze_panes(2, 4 if sname=="全市場整合篩選" else 0)
            ws.autofilter(1, 0, max(1, len(data)+1), max(0, len(data.columns)-1))
            for col, name in enumerate(data.columns):
                hfmt=header_fmt
                if sname=="全市場整合篩選":
                    n=str(name)
                    if n in ["日期","代號","名稱","市場","整合觀察標籤"]:
                        hfmt=header_id_fmt
                    elif n.startswith("三率") or "毛利率" in n or "營業利益率" in n or "稅後純益率" in n or n in ["毛差","營差","淨差"]:
                        hfmt=header_fund_fmt
                    elif n.startswith("營收") or n=="累計營收增%":
                        hfmt=header_rev_fmt
                    elif any(k in n for k in ["EPS","PER","PE","合理價","折價率","估值","機構"]):
                        hfmt=header_val_fmt
                    else:
                        hfmt=header_tech_fmt
                ws.write(1, col, name, hfmt)
                width = min(30 if sname=="全市場整合篩選" else 24, max(10, len(str(name))*2 + 2))
                ws.set_column(col, col, width)
            if sname=="全市場整合篩選":
                for text_col in ["整合觀察標籤","鴨嘴階段","預備鴨嘴狀態","鴨嘴未達條件","三率狀態","基本面資料狀態","基本面缺漏原因","基本面所需季度","營收狀態","營收資料狀態","EPS來源","估值方法","估值狀態"]:
                    if text_col in data.columns:
                        c=data.columns.get_loc(text_col); ws.set_column(c,c,24)
                if "鴨嘴階段" in data.columns:
                    c=data.columns.get_loc("鴨嘴階段"); ws.conditional_format(2,c,max(2,len(data)+1),c,{"type":"text","criteria":"containing","value":"A級","format":good_fmt})
                if "營收狀態" in data.columns:
                    c=data.columns.get_loc("營收狀態"); ws.conditional_format(2,c,max(2,len(data)+1),c,{"type":"text","criteria":"containing","value":"營收三增","format":good_fmt})
            if "月線乖離率%" in data.columns:
                c = data.columns.get_loc("月線乖離率%")
                ws.set_column(c, c, 14, pct_fmt)
                ws.conditional_format(2, c, max(2, len(data)+1), c, {"type":"cell","criteria":">=","value":8,"format":hot_fmt})
            if "折價率%" in data.columns:
                c = data.columns.get_loc("折價率%")
                ws.set_column(c, c, 12, pct_fmt)
                ws.conditional_format(2, c, max(2, len(data)+1), c, {"type":"cell","criteria":">=","value":10,"format":warn_fmt})

        # 台指期
        if fut_latest.empty:
            pd.DataFrame({"狀態":["當日台指期資料未取得"]}).to_excel(writer, sheet_name="台指期", index=False)
        else:
            f = fut_latest.copy()
            f["date"] = pd.to_datetime(f["date"]).dt.strftime("%Y-%m-%d")
            f = f.rename(columns={"date":"日期","contract":"契約","open":"開盤","high":"最高","low":"最低","close":"收盤","settlement":"結算","volume":"成交量","open_interest":"未平倉","MA20_change":"MA20較前日","MA60_change":"MA60較前日","spread_change":"開口較前日","duck":"月季線鴨嘴"})
            f.to_excel(writer, sheet_name="台指期", index=False, startrow=1)
            fib.to_excel(writer, sheet_name="台指期", index=False, startrow=5)
            ws = writer.sheets["台指期"]
            ws.write(0,0,f"{target} 台指期近月（一般交易時段）", title_fmt)
            ws.set_column(0, 20, 14)

        status.to_excel(writer, sheet_name="資料來源狀態", index=False, startrow=1)
        ws = writer.sheets["資料來源狀態"]
        ws.write(0, 0, f"{target} 資料來源狀態", title_fmt)
        ws.set_column(0, 0, 18)
        ws.set_column(1, 4, 14)
        ws.set_column(5, 5, 75)
        ws.set_column(6, 6, 22)

        notes = pd.DataFrame({
            "規則":[

                "v25 TAIFEX：改用期交所官方 futDataDown CSV 批次下載，不再逐日解析 HTML；先抓目標日，再以28日區段補足 MA60。",
                "v25 財報：TWSE/TPEx 六種產業綜合損益 OpenAPI 全部合併；歷史目標只在保守可得日以前寫入，官方兩季不足才用 Yahoo 季報備援。",
                "v25 狀態：『官方12季歷史深度』與『當日能否判斷三率』分開，不再把歷史深度不足誤當今天篩選失敗。",
                "v23 主工作表『全市場整合篩選』：一列一檔股票，技術面／正式鴨嘴／預備鴨嘴／三率／月營收／EPS／估值全部放同一列；其他工作表只保留為快捷檢視。",
                "v23 基本面回補母體改為『正式鴨嘴＋預備鴨嘴』，因此即將可能符合的股票可直接在同一列查看營收三增與三率狀態，不必跨工作表比對。",
                "資料來源狀態新增『判讀』：正常=抓取與可用性正常；部分缺漏=主要資料可用但歷史/個股有缺口；需處理=核心來源未取得。",
                "收盤 > MA20 且收盤 > MA60",
                "MA20 > MA60",
                "MA20、MA60 都較前一交易日上揚",
                "MA20-MA60 開口較前一交易日擴大",
                "v21 預備鴨嘴 A級：正式鴨嘴尚未成立，但5條件已符合4條，而且唯一缺少的條件已在設定的接近門檻內。",
                "v21 預備鴨嘴 B級：已符合3條，另外2條都接近門檻，且至少一個均線/開口動能條件已轉正；用於提早觀察，不代表隔日一定成立。",
                "預備鴨嘴預設接近門檻：股價距需站上的均線<=2%；MA20距MA60<=1%；MA20日變化不低於-0.20%；MA60不低於-0.10%；開口日變化不低於MA60的-0.10%。均可在 settings.ini 調整。",
                "月線正乖離 >= 8%：不建議追價／過熱",
                "曾過熱後月線乖離 <= 5%：過熱解除",
                "普通股篩選：四位數股票代號且 >= 1000；排除 DR/TDR 存託憑證與常見特別股文字",
                "三率三升：比較最新兩個『單季』毛利率、營業利益率、稅後純益率；Q2/Q3/Q4 先由累季損益扣除前期還原單季",
                "三率只採 TWSE／TPEx／MOPS 官方綜合損益 point-in-time 資料；v13 不依賴 FinMind",
                "營收三增：最新月份月增率、年增率、累計營收年增率均 > 0，並保存每日快照避免歷史重跑日期穿越",
                "歷史日期若沒有當時已保存的基本面快照，寧可不分類，也不使用後來公布的資料回填",
                "EPS：v21 延續使用 Yahoo Finance/yfinance 季度 Basic EPS（缺值才用 Diluted EPS），最近四季直接加總為 TTM EPS；Yahoo 無資料時退回 MOPS EPS／交易所隱含 EPS",
                "Yahoo EPS 僅使用 target 日以前已實際保存的快照；補跑過去日期時不會拿今天才看到的 Yahoo 財報倒灌，避免 Point-in-time 日期穿越",
                "估值優先：若 institution_valuation.csv 有有效的 Forward EPS＋機構目標PE，合理價=Forward EPS×目標PE",
                "估值母體：全部上市＋上櫃普通股，不要求先符合鴨嘴／三率三升／營收三增；上述條件只作為額外標籤。",
                "估值備援：抓不到機構目標PE時，合理價依序使用 Yahoo TTM EPS、MOPS TTM EPS、交易所隱含 EPS × TWSE／TPEx 官方近3年歷史PE中位數；Yahoo 只抓優先名單避免全市場逐檔請求造成限流",
                "估值狀態預設：折價>=20% 明顯低估；>=10% 買進區；0~10% 合理偏低；可在 settings.ini 調整。估值只是候選條件，不是無條件下單訊號",
            ]
        })
        notes.to_excel(writer, sheet_name="說明", index=False)
        writer.sheets["說明"].set_column(0,0,70)
    return out


def source_check(conn: sqlite3.Connection, target: dt.date, twse_n: int, tpex_n: int) -> None:
    record_source(conn,target,"TWSE",target.isoformat() if twse_n else None,twse_n>0,twse_n>=500,f"取得普通股 {twse_n} 檔")
    record_source(conn,target,"TPEx",target.isoformat() if tpex_n else None,tpex_n>0,tpex_n>=400,f"取得普通股 {tpex_n} 檔")


def count_market_on_date(conn: sqlite3.Connection, target: dt.date, market: str) -> int:
    rows=conn.execute("SELECT stock_id,name FROM prices WHERE date=? AND market=?",(target.isoformat(),market)).fetchall()
    return sum(1 for code,name in rows if is_common_stock(str(code),str(name or "")))


def self_test(target: dt.date) -> int:
    print("=== 台股鴨嘴系統連線測試 v30 ===",flush=True)
    tests=[]
    base_tests=[("TWSE OpenAPI",fetch_twse_openapi_current),("TPEx OpenAPI",fetch_tpex_openapi_current)]
    for label,fn in base_tests:
        print(f"[{len(tests)+1}/5] {label} 測試中...",flush=True)
        try:
            x=fn(); good=len(x)>0; tests.append(good); print(("[OK] " if good else "[失敗] ")+f"{label}: {len(x)} 檔",flush=True)
        except Exception as e:
            tests.append(False); print(f"[失敗] {label}: {e}",flush=True)
    for label,fn in [("TWSE 官方PER",lambda:_twse_per_day(target)),("TPEx 官方PER",lambda:_tpex_per_day(target))]:
        print(f"[{len(tests)+1}/5] {label} 測試中...",flush=True)
        try:
            x=fn(); good=len(x)>0; tests.append(good); print(("[OK] " if good else "[失敗] ")+f"{label}: {len(x)} 檔",flush=True)
        except Exception as e:
            tests.append(False); print(f"[失敗] {label}: {e}",flush=True)
    print("[5/5] TAIFEX TX 測試中...",flush=True)
    try:
        row=_fetch_taifex_day(target); good=bool(row); tests.append(good); print(("[OK] " if good else "[失敗] ")+f"TAIFEX TX: {row or '無資料'}",flush=True)
    except Exception as e:
        tests.append(False); print(f"[失敗] TAIFEX TX: {e}",flush=True)
    print(f"結果：{sum(tests)}/{len(tests)} 個核心來源成功。v30 不需要 FinMind。",flush=True)
    return 0 if all(tests) else 2


def repair_history_only(target: dt.date) -> int:
    """v18：只修補鴨嘴目前符合名單的 24 個月營收與 12 季財報。
    不跑台指期、不重抓全市場 PER、不輸出 Excel，避免 REPAIR_HISTORY 看似卡死。
    """
    cfg=load_config()
    conn=sqlite3.connect(DB_PATH)
    init_db(conn)
    print("=== v18 歷史基本面專用補洞 ===", flush=True)
    print(f"目標日期：{target}", flush=True)
    print("[1/4] 讀取當日鴨嘴名單...", flush=True)
    current,_,_,_,_,_=compute_signals(conn,target,cfg)
    stock_ids=[str(x) for x in current.get("stock_id",pd.Series(dtype=str)).tolist()] if not current.empty else []
    print(f"    目前符合鴨嘴：{len(stock_ids)} 檔", flush=True)
    if not stock_ids:
        print("[停止] 找不到目標日鴨嘴符合名單；請先確認 stock.db 已有該日股價。", flush=True)
        conn.close(); return 2
    print("[2/4] 補最近 24 個可得月份營收...", flush=True)
    _backfill_revenue_history(conn,target,stock_ids,cfg)
    print("[3/4] 補最近 12 個可得季度財報...", flush=True)
    _backfill_income_history(conn,target,stock_ids,cfg)
    print("[4/4] 完整度檢查...", flush=True)
    min_months=max(13,cfg.getint("general","revenue_backfill_months",fallback=24))
    min_q=max(4,cfg.getint("general","fundamental_backfill_quarters",fallback=12))
    mfull=qfull=0
    for sid in stock_ids:
        expected_m=_available_month_keys_for_stock(conn,sid,target,min_months)
        marks=','.join('?' for _ in expected_m)
        mn=int(conn.execute(f"SELECT COUNT(*) FROM fundamental_revenue_history WHERE stock_id=? AND month_key IN ({marks})",(sid,*expected_m)).fetchone()[0]) if expected_m else 0
        expected_q=_available_quarters_for_stock(conn,sid,target,min_q)
        qn=sum(1 for y,q in expected_q if conn.execute("SELECT 1 FROM fundamental_income_history WHERE stock_id=? AND year=? AND quarter=? AND availability_date<=?",(sid,y,q,target.isoformat())).fetchone())
        mfull += mn>=min_months
        qfull += qn>=min_q
    print(f"    營收滿 {min_months} 月：{mfull}/{len(stock_ids)} 檔", flush=True)
    print(f"    財報滿 {min_q} 季：{qfull}/{len(stock_ids)} 檔", flush=True)
    print("=== 歷史補洞完成 ===", flush=True)
    print("接著執行 DAILY_UPDATE.cmd 重新產生 Excel 即可。", flush=True)
    conn.close(); return 0


def fresh_reset() -> None:
    """v23：從零重跑。只刪除本版本資料夾內的資料庫/輸出，不碰任何舊版資料夾。"""
    print("=== v23 從零重跑：清空本版本資料 ===", flush=True)
    # SQLite WAL/SHM 必須一起移除，避免殘留狀態。
    for path in [DB_PATH, Path(str(DB_PATH)+"-wal"), Path(str(DB_PATH)+"-shm")]:
        try:
            if path.exists():
                path.unlink()
                print(f"[RESET] 已刪除 {path.name}", flush=True)
        except Exception as e:
            raise RuntimeError(f"無法刪除 {path}: {e}")
    # 舊 Excel / log 清掉，避免誤看舊結果。
    for f in OUT_DIR.glob("*.xlsx"):
        try: f.unlink()
        except Exception: pass
    # update.log 由本次執行直接續寫；不在 FileHandler 開啟後刪除，避免 Windows 檔案鎖定。
    print("[RESET] 完成；即將建立全新 stock.db。", flush=True)


def run(target: dt.date, init_mode: bool = False) -> int:
    print(f"[START] 實際分析日期：{target}", flush=True)
    log.info("Duckbill v30 start | target=%s | init_mode=%s", target, init_mode)
    cfg = load_config()
    calendar_days = cfg.getint("general", "history_calendar_days", fallback=180)
    conn = sqlite3.connect(DB_PATH)
    init_db(conn)
    clear_run_source_status(conn,target)

    print("[1/7] 股票歷史 / 當日行情...", flush=True)
    # 股票歷史初始化/增量更新
    backfill_prices(conn, target, calendar_days)
    # 若目標日仍沒有資料，再單獨重抓一次（排程 18:00 時可補剛更新的盤後資料）
    twse_n = count_market_on_date(conn, target, "TWSE")
    tpex_n = count_market_on_date(conn, target, "TPEx")
    if twse_n == 0 or tpex_n == 0:
        try:
            n1, n2, errs = fetch_and_store_day(conn, target, allow_current_openapi=True)
            twse_n = count_market_on_date(conn, target, "TWSE")
            tpex_n = count_market_on_date(conn, target, "TPEx")
            for e in errs: log.warning(e)
        except Exception as e:
            log.error("目標日股票資料重抓失敗：%s", e)

    print("[2/7] TAIFEX 台指期...", flush=True)
    # 期貨：TAIFEX 官方；FinMind 402/額度不再影響。
    try:
        fetch_futures_taifex(conn,target,cfg)
    except Exception as e:
        log.warning("TAIFEX TX 失敗：%s",e)
        record_source(conn,target,"TAIFEX-TX",None,False,False,f"失敗：{e}")

    source_check(conn,target,twse_n,tpex_n)

    print("[3/7] 計算 MA20 / MA60 與鴨嘴訊號...", flush=True)
    current, new, exits, hot, preduck, tech_all = compute_signals(conn, target, cfg)

    print(f"[4/7] 基本面與營收（正式鴨嘴 {len(current)}＋預備鴨嘴 {len(preduck)}）...", flush=True)
    # v23：基本面回補母體擴大為「正式鴨嘴＋預備鴨嘴」。
    # 這樣 A/B 級早期候選可以直接看到三率與營收，不需要等正式鴨嘴後再重跑一次。
    candidate_base = pd.concat([current, preduck], ignore_index=True, sort=False) if (not current.empty or not preduck.empty) else pd.DataFrame()
    if not candidate_base.empty:
        candidate_base = candidate_base.drop_duplicates("stock_id")
    candidate_enriched, _, _, _ = build_fundamental_filters(conn, target, candidate_base, cfg)

    current_ids=set(current["stock_id"].astype(str).tolist()) if not current.empty else set()
    preduck_ids=set(preduck["stock_id"].astype(str).tolist()) if not preduck.empty else set()
    if not candidate_enriched.empty:
        enriched_current=candidate_enriched[candidate_enriched["stock_id"].astype(str).isin(current_ids)].copy()
        enriched_preduck=candidate_enriched[candidate_enriched["stock_id"].astype(str).isin(preduck_ids)].copy()
    else:
        enriched_current=current.copy(); enriched_preduck=preduck.copy()
    both_seed = enriched_current[enriched_current.get("三率三升",False).fillna(False).astype(bool) & enriched_current.get("營收三增",False).fillna(False).astype(bool)].copy() if not enriched_current.empty and isinstance(enriched_current.get("三率三升"),pd.Series) else pd.DataFrame()

    print("[5/7] 全市場估值 / 官方歷史 PER / Yahoo EPS...", flush=True)
    # 估值仍是全市場母體；Yahoo EPS 對正式鴨嘴＋預備鴨嘴＋接近買進區者優先補強。
    current, both, market_valuation, valuation_buy = build_valuation_filters(conn, target, enriched_current, both_seed, cfg, enriched_preduck)

    # 正式鴨嘴的快捷工作表保留；主工作表稍後會整合全市場所有欄位。
    if not current.empty:
        trflag=current.get("三率三升",pd.Series(False,index=current.index)).fillna(False).astype(bool)
        rvflag=current.get("營收三增",pd.Series(False,index=current.index)).fillna(False).astype(bool)
        three_rates = current[trflag].copy().sort_values(["bias20_pct","stock_id"], ascending=[False,True])
        three_revenue = current[rvflag].copy().sort_values(["bias20_pct","stock_id"], ascending=[False,True])
        both = current[trflag & rvflag].copy()
        if not both.empty:
            sort_cols=[c for c in ["折價率%","bias20_pct"] if c in both.columns]
            if sort_cols: both=both.sort_values(sort_cols,ascending=[False]*len(sort_cols),na_position="last")
        enrich_cols = [c for c in current.columns if c not in ["date","name","market","close","MA20","MA60","MA20_change","MA60_change","spread","spread_change","bias20_pct","首次過熱日期","狀態"]]
        enrich_cols = ["stock_id"] + [c for c in enrich_cols if c != "stock_id"]
        meta = current[enrich_cols].drop_duplicates("stock_id") if enrich_cols else pd.DataFrame()
        if not meta.empty:
            if not new.empty: new = new.merge(meta, on="stock_id", how="left")
            if not hot.empty: hot = hot.merge(meta, on="stock_id", how="left")
    else:
        three_rates=pd.DataFrame(); three_revenue=pd.DataFrame(); both=pd.DataFrame()

    # 預備鴨嘴現在同時帶三率、營收與估值。
    preduck = enriched_preduck.copy()
    if not market_valuation.empty:
        market_valuation["預備鴨嘴狀態"] = ""
        market_valuation["預備鴨嘴尚缺條件"] = ""
        if not preduck.empty:
            miss_col="尚缺條件" if "尚缺條件" in preduck.columns else "鴨嘴未達條件"
            tag_cols=["stock_id","預備鴨嘴狀態"] + ([miss_col] if miss_col in preduck.columns else [])
            ptags = preduck[tag_cols].drop_duplicates("stock_id")
            if miss_col in ptags.columns:
                ptags=ptags.rename(columns={miss_col:"預備鴨嘴尚缺條件"})
            market_valuation = market_valuation.drop(columns=["預備鴨嘴狀態","預備鴨嘴尚缺條件"], errors="ignore").merge(ptags, on="stock_id", how="left")
            market_valuation["預備鴨嘴狀態"] = market_valuation["預備鴨嘴狀態"].fillna("")
            market_valuation["預備鴨嘴尚缺條件"] = market_valuation.get("預備鴨嘴尚缺條件","").fillna("") if isinstance(market_valuation.get("預備鴨嘴尚缺條件"),pd.Series) else ""

            val_cols = ["stock_id","最新單季EPS","前一季EPS","去年同期EPS","TTM EPS","EPS年增率%","EPS資料期","EPS來源","EPS觀測日",
                        "Yahoo最新單季EPS","Yahoo TTM EPS","MOPS TTM EPS","交易所隱含EPS",
                        "市場PER","歷史PE中位數","歷史PE_25%","歷史PE_75%","歷史PE樣本數",
                        "Forward EPS","機構目標PE","機構估值來源","機構估值日期","估值採用EPS","估值採用PE",
                        "合理價","折價率%","估值方法","估值狀態"]
            pv = market_valuation[[c for c in val_cols if c in market_valuation.columns]].drop_duplicates("stock_id")
            preduck = preduck.drop(columns=[c for c in val_cols if c != "stock_id" and c in preduck.columns], errors="ignore")
            preduck = preduck.merge(pv, on="stock_id", how="left")

    # v23 主工作表：全市場一張表，直接篩「預備鴨嘴 + 營收三增 + 估值」即可。
    master = build_integrated_master(conn,target,tech_all,candidate_enriched,market_valuation,valuation_buy,cfg)

    print("[6/7] 台指期波浪 / 費波整理...", flush=True)
    fut_latest, fib = futures_analysis(conn, target)
    print("[7/7] 輸出 Excel...", flush=True)
    out = export_excel(conn, target, master, current, new, exits, hot, preduck, fut_latest, fib, three_rates, three_revenue, both, market_valuation, valuation_buy)

    print("\n=== 完成 ===")
    print(f"資料庫：{DB_PATH}")
    print(f"Excel：{out}")
    print(f"TWSE {target}: {twse_n} 檔；TPEx {target}: {tpex_n} 檔")
    print(f"預備鴨嘴：{len(preduck)}；今日新進：{len(new)}；今日退出：{len(exits)}；目前符合：{len(current)}；過熱：{len(hot)}")
    estimable = int(pd.to_numeric(market_valuation.get("合理價", pd.Series(dtype=float)), errors="coerce").notna().sum()) if not market_valuation.empty else 0
    print(f"鴨嘴×三率三升：{len(three_rates)}；鴨嘴×營收三增：{len(three_revenue)}；兩者皆有：{len(both)}")
    print(f"全市場整合篩選：{len(master)} 檔；全市場估值母體：{len(market_valuation)}；可計算合理價：{estimable}；估值買進候選：{len(valuation_buy)}")
    if twse_n < 500 or tpex_n < 400:
        print("注意：其中一個市場的當日資料可能不完整，請查看 Excel『資料來源狀態』與 log/update.log。")
    conn.close()
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="指定分析日期 YYYY-MM-DD；預設今天")
    ap.add_argument("--init", action="store_true", help="第一次初始化")
    ap.add_argument("--update", action="store_true", help="每日更新")
    ap.add_argument("--repair-history", action="store_true", help="只補24個月營收與12季財報，不跑全市場估值")
    ap.add_argument("--self-test", action="store_true", help="測試資料來源連線")
    ap.add_argument("--fresh", action="store_true", help="從零重跑：刪除本版本 stock.db / Excel 後完整重建")
    args = ap.parse_args()
    target = parse_date(args.date)
    if args.fresh:
        fresh_reset()
    if args.self_test:
        return self_test(target)
    if args.repair_history:
        return repair_history_only(target)
    return run(target, init_mode=args.init)


if __name__ == "__main__":
    raise SystemExit(main())
