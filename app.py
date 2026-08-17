# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import datetime as dt
import math

import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
DEFAULT_RS = APP_DIR / "rs_latest.xlsx"
DEFAULT_DUCK = APP_DIR / "duck_latest.xlsx"

st.set_page_config(
    page_title="台股分析中心",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
.block-container {padding-top: 1.3rem; padding-bottom: 3rem; max-width: 1500px;}
[data-testid="stMetric"] {border: 1px solid rgba(120,120,120,.18); border-radius: 12px; padding: 10px 14px;}
.small-note {font-size: .88rem; opacity: .72;}
.status-box {padding: .8rem 1rem; border-radius: 10px; background: rgba(120,120,120,.08); margin-bottom: .8rem;}
@media (max-width: 640px) {
  .block-container {padding-left: .75rem; padding-right: .75rem; padding-top: .8rem;}
  h1 {font-size: 1.8rem !important;}
  h2 {font-size: 1.35rem !important;}
}
</style>
""",
    unsafe_allow_html=True,
)


def _xlsx_source(default_path: Path, uploaded_key: str):
    raw = st.session_state.get(uploaded_key)
    if raw:
        return BytesIO(raw)
    return default_path


@st.cache_data(show_spinner=False)
def _read_excel_path(path_s: str, sheet: str, header=0):
    return pd.read_excel(path_s, sheet_name=sheet, header=header)


def read_sheet(source, sheet: str, header=0):
    if isinstance(source, Path):
        return _read_excel_path(str(source), sheet, header)
    source.seek(0)
    return pd.read_excel(source, sheet_name=sheet, header=header)


def safe_num(v, digits=2, suffix=""):
    try:
        if pd.isna(v):
            return "—"
        n = float(v)
        if not math.isfinite(n):
            return "—"
        if digits == 0:
            return f"{n:,.0f}{suffix}"
        return f"{n:,.{digits}f}{suffix}"
    except Exception:
        return "—" if v is None else str(v)


def text(v, default="—"):
    try:
        if pd.isna(v):
            return default
    except Exception:
        pass
    s = str(v).strip()
    return s if s else default


def normalize_date_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        x = pd.to_numeric(s, errors="coerce")
        med = x.dropna().median() if x.notna().any() else None
        if med is not None and 25000 < med < 80000:
            return pd.to_datetime(x, unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(s, errors="coerce")


def latest_date_from_df(df: pd.DataFrame) -> str:
    if "日期" not in df.columns or df.empty:
        return "—"
    d = normalize_date_series(df["日期"])
    if d.notna().any():
        return d.max().strftime("%Y-%m-%d")
    return text(df.iloc[-1].get("日期"))


def rs_data(source):
    daily = read_sheet(source, "每日強勢股數量")
    strong = read_sheet(source, "最新強勢股名單")
    extreme = read_sheet(source, "極值基準")
    overview = read_sheet(source, "市場廣度總覽", header=None)
    if daily.empty:
        raise ValueError("RS 檔案的『每日強勢股數量』工作表沒有資料")
    daily = daily.copy()
    if "日期" in daily.columns:
        daily["日期"] = normalize_date_series(daily["日期"])
    r = daily.iloc[-1]
    return daily, strong, extreme, overview, r


def duck_sheet(source, name):
    return read_sheet(source, name, header=1)


def duck_data(source):
    all_ok = duck_sheet(source, "全部符合")
    new = duck_sheet(source, "今日新進")
    exits = duck_sheet(source, "今日退出")
    pre = duck_sheet(source, "即將可能符合")
    both = duck_sheet(source, "鴨嘴×兩者皆有")
    three_rate = duck_sheet(source, "鴨嘴×三率三升")
    three_rev = duck_sheet(source, "鴨嘴×營收三增")
    value = duck_sheet(source, "估值買進候選")
    tx = duck_sheet(source, "台指期")
    source_status = duck_sheet(source, "資料來源狀態")
    return all_ok, new, exits, pre, both, three_rate, three_rev, value, tx, source_status


def clean_duck(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    x = df.copy()
    keep = [
        "日期", "代號", "名稱", "市場", "收盤", "MA20", "MA60", "開口較前日",
        "月線乖離率%", "首次過熱日期", "狀態", "預備鴨嘴狀態", "尚缺條件",
        "培育中心類別", "最新單季EPS", "合理價", "折價率%", "估值狀態"
    ]
    cols = [c for c in keep if c in x.columns]
    return x[cols]


def filter_stock_table(df: pd.DataFrame, query: str) -> pd.DataFrame:
    if df.empty or not query.strip():
        return df
    q = query.strip()
    code = df.get("代號", pd.Series("", index=df.index)).astype(str)
    name = df.get("名稱", pd.Series("", index=df.index)).astype(str)
    return df[code.str.contains(q, case=False, na=False) | name.str.contains(q, case=False, na=False)]


rs_source = _xlsx_source(DEFAULT_RS, "rs_uploaded_bytes")
duck_source = _xlsx_source(DEFAULT_DUCK, "duck_uploaded_bytes")

try:
    daily, strong, extreme, overview, rs_latest = rs_data(rs_source)
    all_ok, duck_new, duck_exit, pre, both, three_rate, three_rev, value_candidates, tx, source_status = duck_data(duck_source)
except Exception as e:
    st.error(f"資料載入失敗：{e}")
    st.info("請確認 rs_latest.xlsx 與 duck_latest.xlsx 都放在 app.py 同一層。")
    st.stop()

rs_date = latest_date_from_df(daily)
duck_date = latest_date_from_df(all_ok)
latest_date = max([d for d in [rs_date, duck_date] if d != "—"], default="—")

water = text(rs_latest.get("歷史水位"))
direction = text(rs_latest.get("強勢股方向"))
stage = text(rs_latest.get("市場階段"))
strong_count = rs_latest.get("強勢股檔數")
strong_pct = rs_latest.get("強勢股占比%")
change = rs_latest.get("較前日增減")
full_rank = rs_latest.get("全期回顧百分位%")
rolling_rank = rs_latest.get("近3年滾動百分位%")
spread = rs_latest.get("歷史水位.1") if "歷史水位.1" in daily.columns else None
speed = text(rs_latest.get("強勢股方向.1")) if "強勢股方向.1" in daily.columns else "—"

pre_status = pre.get("預備鴨嘴狀態", pd.Series(dtype=str)).fillna("").astype(str) if not pre.empty else pd.Series(dtype=str)
pre_a = int(pre_status.str.contains("A級").sum()) if not pre_status.empty else 0
pre_b = int(pre_status.str.contains("B級").sum()) if not pre_status.empty else 0

st.title("📈 台股分析中心")
st.caption("RS 強勢股市場廣度 × 鴨嘴型態 × 培育中心｜正式網頁版 v1")

if rs_date != duck_date:
    st.warning(f"目前兩套資料日期不同：RS={rs_date}、鴨嘴={duck_date}。請以各區塊標示日期為準。")
else:
    st.success(f"目前資料日期：{latest_date}｜手機 4G/5G 可直接使用")

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("市場階段", stage)
m2.metric("強勢股方向", direction, f"{safe_num(change,0)} 檔")
m3.metric("RS 強勢股", f"{safe_num(strong_count,0)} 檔", safe_num(strong_pct,2,"%"))
m4.metric("鴨嘴符合", f"{len(all_ok):,} 檔", f"新進 {len(duck_new):,} 檔")
m5.metric("預備 A 級", f"{pre_a:,} 檔", f"B級 {pre_b:,} 檔")

summary_msg = f"目前屬「{water}／{stage}」，強勢股方向為「{direction}」，變化速度「{speed}」。"
if direction == "增加" and "低檔" in stage:
    summary_msg += " 市場內部正在從低檔改善，但尚不等同於大盤已確認見底。"
elif direction == "減少" and "高檔" in stage:
    summary_msg += " 市場內部出現高檔退潮訊號，宜提高風險警覺。"
st.markdown(f'<div class="status-box"><b>一句話判讀：</b>{summary_msg}</div>', unsafe_allow_html=True)

tabs = st.tabs(["🏠 總覽", "📊 RS／市場廣度", "🦆 鴨嘴系統", "🔥 鴨嘴×培育中心", "📐 台指期", "🧪 資料品質", "🔄 更新資料"])

with tabs[0]:
    c1, c2 = st.columns([1.1, 1])
    with c1:
        st.subheader("RS 市場溫度")
        a, b, c, d = st.columns(4)
        a.metric("歷史水位", water)
        b.metric("全期百分位", f"P{safe_num(full_rank,1)}")
        c.metric("近3年百分位", f"P{safe_num(rolling_rank,1)}")
        d.metric("5日－20日差", safe_num(spread,2," 個百分點"))
        st.caption(f"RS 資料日：{rs_date}")

        if "日期" in daily.columns:
            chart_cols = [c for c in ["強勢股占比%", "5日平均占比%", "20日平均占比%", "60日平均占比%"] if c in daily.columns]
            h = daily[["日期"] + chart_cols].dropna(subset=["日期"]).tail(120).set_index("日期")
            st.line_chart(h)

    with c2:
        st.subheader("今日鴨嘴摘要")
        x1, x2, x3 = st.columns(3)
        x1.metric("全部符合", f"{len(all_ok):,}")
        x2.metric("今日新進", f"{len(duck_new):,}")
        x3.metric("今日退出", f"{len(duck_exit):,}")
        y1, y2, y3 = st.columns(3)
        y1.metric("A級預備", f"{pre_a:,}")
        y2.metric("兩培育條件皆有", f"{len(both):,}")
        y3.metric("估值候選", f"{len(value_candidates):,}")
        st.caption(f"鴨嘴資料日：{duck_date}")
        if not duck_new.empty:
            st.write("**今日新進（前 15 檔）**")
            st.dataframe(clean_duck(duck_new).head(15), hide_index=True, use_container_width=True)

with tabs[1]:
    st.subheader("RS 強勢股市場廣度")
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("有效樣本", f"{safe_num(rs_latest.get('有效樣本數'),0)}")
    k2.metric("RS母池", f"{safe_num(rs_latest.get('RS強勢母池檔數'),0)}")
    k3.metric("強勢股", f"{safe_num(strong_count,0)}")
    k4.metric("占比", safe_num(strong_pct,2,"%"))
    k5.metric("較前日", safe_num(change,0," 檔"))

    days = st.slider("歷史圖顯示交易日", 30, min(500, len(daily)), min(180, len(daily)), 10)
    chart_cols = [c for c in ["強勢股占比%", "5日平均占比%", "20日平均占比%", "60日平均占比%"] if c in daily.columns]
    hist = daily[["日期"] + chart_cols].dropna(subset=["日期"]).tail(days).set_index("日期")
    st.line_chart(hist)

    st.write("**最新 RS 強勢股名單**")
    rs_q = st.text_input("搜尋代號／名稱", key="rs_search")
    rs_show = strong.copy()
    if rs_q.strip() and not rs_show.empty:
        code = rs_show.get("代號", pd.Series("", index=rs_show.index)).astype(str)
        name = rs_show.get("名稱", pd.Series("", index=rs_show.index)).astype(str)
        rs_show = rs_show[code.str.contains(rs_q, case=False, na=False) | name.str.contains(rs_q, case=False, na=False)]
    st.dataframe(rs_show, hide_index=True, use_container_width=True, height=520)

    with st.expander("歷史極值基準"):
        st.dataframe(extreme, hide_index=True, use_container_width=True)

with tabs[2]:
    st.subheader("鴨嘴型態篩選")
    mode = st.radio("查看", ["全部符合", "今日新進", "今日退出", "即將可能符合"], horizontal=True)
    source_map = {"全部符合": all_ok, "今日新進": duck_new, "今日退出": duck_exit, "即將可能符合": pre}
    dfx = source_map[mode].copy()
    q = st.text_input("搜尋代號／名稱", key="duck_search")
    dfx = filter_stock_table(dfx, q)

    if mode == "全部符合" and "狀態" in dfx.columns:
        options = sorted([x for x in dfx["狀態"].dropna().astype(str).unique().tolist() if x])
        chosen = st.multiselect("狀態篩選", options)
        if chosen:
            dfx = dfx[dfx["狀態"].astype(str).isin(chosen)]
    if mode == "即將可能符合" and "預備鴨嘴狀態" in dfx.columns:
        grade = st.selectbox("預備等級", ["全部", "A級", "B級"])
        if grade != "全部":
            dfx = dfx[dfx["預備鴨嘴狀態"].fillna("").astype(str).str.contains(grade)]

    st.caption(f"顯示 {len(dfx):,} 檔")
    st.dataframe(clean_duck(dfx), hide_index=True, use_container_width=True, height=620)

with tabs[3]:
    st.subheader("鴨嘴 × 基本面培育中心")
    bmode = st.radio("交集類型", ["兩者皆有", "三率三升", "營收三增", "估值買進候選"], horizontal=True)
    bmap = {"兩者皆有": both, "三率三升": three_rate, "營收三增": three_rev, "估值買進候選": value_candidates}
    bx = bmap[bmode].copy()
    bq = st.text_input("搜尋代號／名稱", key="both_search")
    bx = filter_stock_table(bx, bq)

    if bmode == "估值買進候選":
        cols = [c for c in ["日期","代號","名稱","市場","收盤","合理價","折價率%","估值狀態","鴨嘴狀態","月線乖離率%","培育中心類別","是否鴨嘴"] if c in bx.columns]
        if "折價率%" in bx.columns:
            bx = bx.sort_values("折價率%", ascending=False, na_position="last")
        st.dataframe(bx[cols] if cols else bx, hide_index=True, use_container_width=True, height=620)
    else:
        st.dataframe(clean_duck(bx), hide_index=True, use_container_width=True, height=620)

with tabs[4]:
    st.subheader("台指期月／季線鴨嘴參考")
    if tx.empty:
        st.info("目前沒有台指期資料。")
    else:
        first = tx.iloc[0]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("收盤", safe_num(first.get("收盤"),0))
        c2.metric("MA20", safe_num(first.get("MA20"),2))
        c3.metric("MA60", safe_num(first.get("MA60"),2))
        mouth = first.get("月季線鴨嘴")
        mouth_text = "符合" if str(mouth).lower() in {"true","1","是"} else "未符合"
        c4.metric("月季線鴨嘴", mouth_text)
        st.dataframe(tx, hide_index=True, use_container_width=True)

with tabs[5]:
    st.subheader("資料來源與缺漏狀態")
    if source_status.empty:
        st.info("沒有資料來源狀態表。")
    else:
        bad = source_status.copy()
        if "完整" in bad.columns:
            incomplete = bad[~bad["完整"].fillna("").astype(str).isin(["是", "True", "true", "1"])]
        else:
            incomplete = pd.DataFrame()
        if incomplete.empty:
            st.success("目前資料來源均標記為完整。")
        else:
            st.warning(f"目前有 {len(incomplete)} 個來源標記為部分缺漏／不完整；這不一定代表整套系統失效，請看『訊息』欄。")
        st.dataframe(source_status, hide_index=True, use_container_width=True, height=620)

with tabs[6]:
    st.subheader("更新網頁資料")
    st.info("這一版支援『上傳最新輸出 → 立即在本次網頁工作階段查看』。要永久更新雲端網站，請把 GitHub 裡的 rs_latest.xlsx / duck_latest.xlsx 替換成最新檔案。下一版可再接自動排程。")

    ru = st.file_uploader("上傳最新 RS／市場廣度 xlsx", type=["xlsx"], key="rs_upload_widget")
    du = st.file_uploader("上傳最新 鴨嘴篩選結果 xlsx", type=["xlsx"], key="duck_upload_widget")

    c1, c2 = st.columns(2)
    if c1.button("套用上傳資料", type="primary", use_container_width=True):
        changed = False
        if ru is not None:
            st.session_state["rs_uploaded_bytes"] = ru.getvalue()
            changed = True
        if du is not None:
            st.session_state["duck_uploaded_bytes"] = du.getvalue()
            changed = True
        if changed:
            _read_excel_path.clear()
            st.rerun()
        else:
            st.warning("請先選擇至少一個 xlsx 檔案。")

    if c2.button("恢復 GitHub 內建資料", use_container_width=True):
        st.session_state.pop("rs_uploaded_bytes", None)
        st.session_state.pop("duck_uploaded_bytes", None)
        _read_excel_path.clear()
        st.rerun()

    st.divider()
    st.write("**下載目前 GitHub 內建結果檔**")
    d1, d2 = st.columns(2)
    with open(DEFAULT_RS, "rb") as f:
        d1.download_button("下載 RS 最新結果", f.read(), file_name="rs_latest.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with open(DEFAULT_DUCK, "rb") as f:
        d2.download_button("下載鴨嘴最新結果", f.read(), file_name="duck_latest.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

st.divider()
st.caption(f"正式版 v1｜RS資料：{rs_date}｜鴨嘴資料：{duck_date}｜此頁為量化篩選與市場廣度工具，不構成投資建議。")
