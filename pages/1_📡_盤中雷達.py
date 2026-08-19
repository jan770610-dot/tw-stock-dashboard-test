# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in __import__('sys').path:
    __import__('sys').path.insert(0, str(ROOT))

from intraday_live import render_intraday_panel

st.set_page_config(page_title="盤中即時市場雷達", page_icon="📡", layout="wide", initial_sidebar_state="collapsed")

RS = ROOT / "rs_latest.xlsx"
DUCK = ROOT / "duck_latest.xlsx"

@st.cache_data(show_spinner=False)
def read_sheet(path_s: str, mtime_ns: int, sheet: str, header: int = 0):
    return pd.read_excel(path_s, sheet_name=sheet, header=header)

st.title("📡 台股盤中即時市場雷達")
st.caption("與正式盤後系統分離：盤中只做試算，不會改寫歷史訊號。")

if not RS.exists() or not DUCK.exists():
    st.error("找不到 rs_latest.xlsx 或 duck_latest.xlsx，請先確認主系統資料檔存在。")
    st.stop()

try:
    daily = read_sheet(str(RS), RS.stat().st_mtime_ns, "每日強勢股數量", 0)
    strong = read_sheet(str(RS), RS.stat().st_mtime_ns, "最新強勢股名單", 0)
    all_ok = read_sheet(str(DUCK), DUCK.stat().st_mtime_ns, "全部符合", 1)
    pre = read_sheet(str(DUCK), DUCK.stat().st_mtime_ns, "即將可能符合", 1)
    if "日期" in daily.columns:
        daily["日期"] = pd.to_datetime(daily["日期"], errors="coerce")
    rs_latest = daily.iloc[-1]
except Exception as e:
    st.error(f"正式資料讀取失敗：{e}")
    st.stop()

render_intraday_panel(daily, strong, all_ok, pre, rs_latest)

st.divider()
st.caption("盤中雷達 v0.1｜正式 RS／鴨嘴訊號仍以盤後資料為準。")
