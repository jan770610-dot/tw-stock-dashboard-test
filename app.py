# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import datetime as dt
import html
import json
import math

import altair as alt
import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
DEFAULT_RS = APP_DIR / "rs_latest.xlsx"
DEFAULT_DUCK = APP_DIR / "duck_latest.xlsx"
STATUS_FILE = APP_DIR / "update_status.json"
ACTIONS_URL = "https://github.com/jan770610-dot/tw-stock-dashboard-test/actions/workflows/daily-update.yml"

st.set_page_config(page_title="台股分析中心", page_icon="📈", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 3rem; max-width: 1500px;}
.hero-sub {opacity:.72; margin-top:-.7rem; margin-bottom:1rem;}
.kpi-grid {display:grid; grid-template-columns:repeat(auto-fit,minmax(175px,1fr)); gap:12px; margin:.6rem 0 1rem 0;}
.kpi-card {border:1px solid rgba(120,120,120,.20); border-radius:14px; padding:14px 15px; min-height:105px; background:rgba(127,127,127,.035);}
.kpi-label {font-size:.85rem; opacity:.72; margin-bottom:7px;}
.kpi-value {font-size:1.75rem; font-weight:700; line-height:1.15; overflow-wrap:anywhere;}
.kpi-delta {font-size:.82rem; opacity:.78; margin-top:8px;}
.signal-box {padding:.85rem 1rem; border-radius:12px; background:rgba(120,120,120,.08); margin:.7rem 0 1rem; line-height:1.65;}
.status-strip {padding:.72rem .9rem; border-radius:10px; margin-bottom:.8rem;}
.status-ok {background:#eaf7ee; color:#176b32;}
.status-warn {background:#fff6df; color:#795600;}
.status-bad {background:#fdeaea; color:#8a1c1c;}
.small-note {font-size:.88rem; opacity:.72;}
[data-testid="stDataFrame"] {border:1px solid rgba(120,120,120,.12); border-radius:10px; overflow:hidden;}
@media (max-width: 700px) {
  .block-container {padding-left:.7rem; padding-right:.7rem; padding-top:.65rem;}
  h1 {font-size:1.75rem !important;}
  h2 {font-size:1.28rem !important;}
  .kpi-grid {grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px;}
  .kpi-card {min-height:92px; padding:11px 12px;}
  .kpi-value {font-size:1.35rem;}
}
@media (max-width: 420px) {.kpi-grid {grid-template-columns:1fr 1fr;}}
</style>
""", unsafe_allow_html=True)


def _xlsx_source(default_path: Path, uploaded_key: str):
    raw = st.session_state.get(uploaded_key)
    if raw:
        return BytesIO(raw)
    return default_path


@st.cache_data(show_spinner=False)
def _read_excel_path(path_s: str, mtime_ns: int, sheet: str, header=0):
    return pd.read_excel(path_s, sheet_name=sheet, header=header)


def read_sheet(source, sheet: str, header=0):
    if isinstance(source, Path):
        return _read_excel_path(str(source), source.stat().st_mtime_ns, sheet, header)
    source.seek(0)
    return pd.read_excel(source, sheet_name=sheet, header=header)


def safe_num(v, digits=2, suffix=""):
    try:
        if pd.isna(v): return "—"
        n = float(v)
        if not math.isfinite(n): return "—"
        if digits == 0: return f"{n:,.0f}{suffix}"
        return f"{n:,.{digits}f}{suffix}"
    except Exception:
        return "—" if v is None else str(v)


def text(v, default="—"):
    try:
        if pd.isna(v): return default
    except Exception:
        pass
    s = str(v).strip()
    return s if s else default


def first_value(row, names, default=None):
    for name in names:
        if name in row.index:
            v = row.get(name)
            try:
                if not pd.isna(v): return v
            except Exception:
                if v is not None: return v
    return default


def normalize_date_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        x = pd.to_numeric(s, errors="coerce")
        med = x.dropna().median() if x.notna().any() else None
        if med is not None and 25000 < med < 80000:
            return pd.to_datetime(x, unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(s, errors="coerce")


def latest_date_from_df(df: pd.DataFrame) -> str:
    if "日期" not in df.columns or df.empty: return "—"
    d = normalize_date_series(df["日期"])
    return d.max().strftime("%Y-%m-%d") if d.notna().any() else text(df.iloc[-1].get("日期"))


def rs_data(source):
    daily = read_sheet(source, "每日強勢股數量").copy()
    strong = read_sheet(source, "最新強勢股名單")
    extreme = read_sheet(source, "極值基準")
    if daily.empty: raise ValueError("RS 檔案的『每日強勢股數量』沒有資料")
    if "日期" in daily.columns: daily["日期"] = normalize_date_series(daily["日期"])
    return daily, strong, extreme, daily.iloc[-1]


def duck_sheet(source, name):
    try: return read_sheet(source, name, header=1)
    except Exception: return pd.DataFrame()


def duck_data(source):
    names = ["全部符合","今日新進","今日退出","即將可能符合","鴨嘴×兩者皆有","鴨嘴×三率三升","鴨嘴×營收三增","估值買進候選","台指期","資料來源狀態"]
    return [duck_sheet(source, n) for n in names]


def clean_duck(df):
    if df is None or df.empty: return pd.DataFrame()
    keep = ["日期","代號","名稱","市場","收盤","MA20","MA60","開口較前日","月線乖離率%","首次過熱日期","狀態","預備鴨嘴狀態","尚缺條件","培育中心類別","最新單季EPS","合理價","折價率%","估值狀態"]
    return df[[c for c in keep if c in df.columns]].copy()


def filter_stock_table(df, query):
    if df is None or df.empty or not query.strip(): return df
    q=query.strip(); code=df.get("代號",pd.Series("",index=df.index)).astype(str); name=df.get("名稱",pd.Series("",index=df.index)).astype(str)
    return df[code.str.contains(q,case=False,na=False)|name.str.contains(q,case=False,na=False)]


def load_update_status():
    try: return json.loads(STATUS_FILE.read_text(encoding="utf-8"))
    except Exception: return {}


def card(label, value, delta=""):
    return f'<div class="kpi-card"><div class="kpi-label">{html.escape(str(label))}</div><div class="kpi-value">{html.escape(str(value))}</div><div class="kpi-delta">{html.escape(str(delta))}</div></div>'


def cards(items):
    st.markdown('<div class="kpi-grid">'+''.join(card(*x) for x in items)+'</div>', unsafe_allow_html=True)


def market_signal(stage):
    if stage in {"低檔回升","中段回升"}: return "🟢 結構改善"
    if stage == "高檔續強": return "🟢 多頭擴散"
    if stage in {"低檔續弱","中段走弱","高檔轉弱"}: return "🔴 結構轉弱"
    return "🟡 整理觀察"


def breadth_chart(df, days=120):
    cols=[c for c in ["強勢股占比%","5日平均占比%","20日平均占比%","60日平均占比%"] if c in df.columns]
    if "日期" not in df.columns or not cols: return None
    x=df[["日期"]+cols].dropna(subset=["日期"]).tail(days).copy()
    rename={"強勢股占比%":"當日強勢股占比","5日平均占比%":"5日平均","20日平均占比%":"20日平均","60日平均占比%":"60日平均"}
    x=x.rename(columns=rename).melt("日期",var_name="指標",value_name="占比")
    chart=alt.Chart(x).mark_line().encode(
        x=alt.X("日期:T",title="日期"), y=alt.Y("占比:Q",title="占有效樣本比例 (%)"),
        color=alt.Color("指標:N",title="圖例"),
        tooltip=[alt.Tooltip("日期:T",title="日期"),alt.Tooltip("指標:N"),alt.Tooltip("占比:Q",format=".2f")]
    ).properties(height=330)
    return chart


rs_source=_xlsx_source(DEFAULT_RS,"rs_uploaded_bytes")
duck_source=_xlsx_source(DEFAULT_DUCK,"duck_uploaded_bytes")
try:
    daily,strong,extreme,rs_latest=rs_data(rs_source)
    all_ok,duck_new,duck_exit,pre,both,three_rate,three_rev,value_candidates,tx,source_status=duck_data(duck_source)
except Exception as e:
    st.error(f"資料載入失敗：{e}"); st.stop()

rs_date=latest_date_from_df(daily); duck_date=latest_date_from_df(all_ok)
latest_date=max([d for d in [rs_date,duck_date] if d!="—"], default="—")
water=text(first_value(rs_latest,["歷史水位"]))
direction=text(first_value(rs_latest,["強勢股方向"]))
stage=text(first_value(rs_latest,["市場階段"]))
speed=text(first_value(rs_latest,["變化速度","強勢股方向.1"]))
spread=first_value(rs_latest,["5日-20日廣度差","歷史水位.1"])
spread_change=first_value(rs_latest,["廣度差5日變化","極值訊號.1"])
strong_count=first_value(rs_latest,["強勢股檔數"])
strong_pct=first_value(rs_latest,["強勢股占比%"])
change=first_value(rs_latest,["較前日增減"])
full_rank=first_value(rs_latest,["全期回顧百分位%"])
rolling_rank=first_value(rs_latest,["近3年滾動百分位%"])
pre_status=pre.get("預備鴨嘴狀態",pd.Series(dtype=str)).fillna("").astype(str) if not pre.empty else pd.Series(dtype=str)
pre_a=int(pre_status.str.contains("A級").sum()) if not pre_status.empty else 0
pre_b=int(pre_status.str.contains("B級").sum()) if not pre_status.empty else 0
update_status=load_update_status()

st.title("📈 台股分析中心")
st.markdown('<div class="hero-sub">RS 市場廣度 × 鴨嘴型態 × 培育中心｜正式網頁版 v2</div>',unsafe_allow_html=True)

u_status=update_status.get("status","unknown")
last_run=update_status.get("last_run_taipei","—")
message=update_status.get("message","")
if rs_date==duck_date:
    cls="status-ok" if u_status in {"success","ready","bootstrap"} else "status-warn"
    st.markdown(f'<div class="status-strip {cls}"><b>資料日：</b>{latest_date}　｜　<b>自動更新：</b>週一至週五 18:05 啟動　｜　<b>最近排程：</b>{html.escape(str(last_run))}</div>',unsafe_allow_html=True)
else:
    st.markdown(f'<div class="status-strip status-warn"><b>資料日期不同：</b>RS {rs_date}／鴨嘴 {duck_date}。排程會保留各系統最後成功結果。</div>',unsafe_allow_html=True)
if u_status in {"partial","no_new_data"} and message: st.warning(message)
elif u_status=="error" and message: st.error(message)

cards([
    ("市場訊號",market_signal(stage),f"市場階段：{stage}"),
    ("歷史水位",water,f"全期 P{safe_num(full_rank,1)}／近3年 P{safe_num(rolling_rank,1)}"),
    ("強勢股方向",direction,f"變化速度：{speed}"),
    ("5日－20日廣度差",safe_num(spread,2," 個百分點"),f"5日變化 {safe_num(spread_change,2)}"),
    ("RS 強勢股",safe_num(strong_count,0," 檔"),f"占比 {safe_num(strong_pct,2,'%')}／單日 {safe_num(change,0,' 檔')}"),
    ("鴨嘴符合",f"{len(all_ok):,} 檔",f"新進 {len(duck_new):,}／退出 {len(duck_exit):,}／A級 {pre_a:,}"),
])

summary=f"目前為「{water}／{stage}」，強勢股方向「{direction}」、速度「{speed}」。"
if change is not None:
    try:
        c=float(change)
        if c<0 and direction in {"增加","明顯增加"}: summary += f" 雖然單日減少 {abs(int(c))} 檔，但中短期廣度結構仍維持增加，兩者不矛盾。"
    except Exception: pass
if stage=="低檔回升": summary += " 代表市場內部由低水位開始修復，仍需觀察是否持續擴散。"
elif stage=="高檔轉弱": summary += " 代表高水位開始退潮，風險明顯高於單純高水位。"
st.markdown(f'<div class="signal-box"><b>一句話判讀：</b>{html.escape(summary)}</div>',unsafe_allow_html=True)

tabs=st.tabs(["🏠 總覽","📊 RS／市場廣度","🦆 鴨嘴系統","🔥 鴨嘴×培育中心","📐 台指期","🧪 資料品質","🔄 更新資料"])

with tabs[0]:
    c1,c2=st.columns([1.12,1])
    with c1:
        st.subheader("RS 市場溫度")
        cards([("歷史水位",water,""),("全期百分位",f"P{safe_num(full_rank,1)}",""),("近3年百分位",f"P{safe_num(rolling_rank,1)}",""),("廣度差",safe_num(spread,2),speed)])
        ch=breadth_chart(daily,120)
        if ch is not None: st.altair_chart(ch,use_container_width=True)
        st.caption(f"RS 資料日：{rs_date}")
    with c2:
        st.subheader("今日鴨嘴摘要")
        cards([("全部符合",f"{len(all_ok):,}",""),("今日新進",f"{len(duck_new):,}",""),("今日退出",f"{len(duck_exit):,}",""),("A級預備",f"{pre_a:,}",f"B級 {pre_b:,}"),("兩培育條件皆有",f"{len(both):,}",""),("估值候選",f"{len(value_candidates):,}","")])
        if not duck_new.empty:
            st.write("**今日新進（前 15 檔）**")
            st.dataframe(clean_duck(duck_new).head(15),hide_index=True,use_container_width=True)
        st.caption(f"鴨嘴資料日：{duck_date}")

with tabs[1]:
    st.subheader("RS 強勢股市場廣度")
    cards([("有效樣本",safe_num(first_value(rs_latest,["有效樣本數"]),0),""),("RS 強勢母池",safe_num(first_value(rs_latest,["RS強勢母池檔數"]),0),""),("強勢股",safe_num(strong_count,0),""),("占比",safe_num(strong_pct,2,"%"),""),("較前日",safe_num(change,0," 檔"),"")])
    days=st.slider("歷史圖顯示交易日",30,min(500,len(daily)),min(180,len(daily)),10)
    ch=breadth_chart(daily,days)
    if ch is not None: st.altair_chart(ch,use_container_width=True)
    st.write("**最新 RS 強勢股名單**")
    rs_q=st.text_input("搜尋代號／名稱",key="rs_search")
    rs_show=filter_stock_table(strong.copy(),rs_q)
    st.dataframe(rs_show,hide_index=True,use_container_width=True,height=540)
    with st.expander("歷史極值基準"): st.dataframe(extreme,hide_index=True,use_container_width=True)

with tabs[2]:
    st.subheader("鴨嘴型態篩選")
    mode=st.radio("查看",["全部符合","今日新進","今日退出","即將可能符合"],horizontal=True)
    source_map={"全部符合":all_ok,"今日新進":duck_new,"今日退出":duck_exit,"即將可能符合":pre}
    dfx=source_map[mode].copy(); q=st.text_input("搜尋代號／名稱",key="duck_search"); dfx=filter_stock_table(dfx,q)
    if mode=="全部符合" and "狀態" in dfx.columns:
        options=sorted([x for x in dfx["狀態"].dropna().astype(str).unique().tolist() if x]); chosen=st.multiselect("狀態篩選",options)
        if chosen: dfx=dfx[dfx["狀態"].astype(str).isin(chosen)]
    if mode=="即將可能符合" and "預備鴨嘴狀態" in dfx.columns:
        grade=st.selectbox("預備等級",["全部","A級","B級"])
        if grade!="全部": dfx=dfx[dfx["預備鴨嘴狀態"].fillna("").astype(str).str.contains(grade)]
    st.caption(f"顯示 {len(dfx):,} 檔")
    st.dataframe(clean_duck(dfx),hide_index=True,use_container_width=True,height=620)

with tabs[3]:
    st.subheader("鴨嘴 × 基本面培育中心")
    bmode=st.radio("交集類型",["兩者皆有","三率三升","營收三增","估值買進候選"],horizontal=True)
    bmap={"兩者皆有":both,"三率三升":three_rate,"營收三增":three_rev,"估值買進候選":value_candidates}
    bx=filter_stock_table(bmap[bmode].copy(),st.text_input("搜尋代號／名稱",key="both_search"))
    if bmode=="估值買進候選":
        cols=[c for c in ["日期","代號","名稱","市場","收盤","合理價","折價率%","估值狀態","鴨嘴狀態","月線乖離率%","培育中心類別","是否鴨嘴"] if c in bx.columns]
        if "折價率%" in bx.columns: bx=bx.sort_values("折價率%",ascending=False,na_position="last")
        st.dataframe(bx[cols] if cols else bx,hide_index=True,use_container_width=True,height=620)
    else: st.dataframe(clean_duck(bx),hide_index=True,use_container_width=True,height=620)

with tabs[4]:
    st.subheader("台指期月／季線鴨嘴參考")
    if tx.empty: st.info("目前沒有台指期資料。")
    else:
        first=tx.iloc[0]; mouth=first.get("月季線鴨嘴"); mouth_text="符合" if str(mouth).lower() in {"true","1","是"} else "未符合"
        cards([("收盤",safe_num(first.get("收盤"),0),""),("MA20",safe_num(first.get("MA20"),2),""),("MA60",safe_num(first.get("MA60"),2),""),("月季線鴨嘴",mouth_text,"")])
        st.dataframe(tx,hide_index=True,use_container_width=True)

with tabs[5]:
    st.subheader("資料來源與自動更新狀態")
    st.write("**GitHub Actions 更新狀態**")
    st.json(update_status if update_status else {"狀態":"尚無 update_status.json"})
    st.write("**鴨嘴資料來源驗收**")
    if source_status.empty: st.info("沒有資料來源狀態表。")
    else:
        bad=source_status.copy(); incomplete=pd.DataFrame()
        if "完整" in bad.columns: incomplete=bad[~bad["完整"].fillna("").astype(str).isin(["是","True","true","1"])]
        if incomplete.empty: st.success("目前資料來源均標記為完整。")
        else: st.warning(f"有 {len(incomplete)} 個來源標記為部分缺漏／不完整，請看訊息欄。")
        st.dataframe(source_status,hide_index=True,use_container_width=True,height=620)

with tabs[6]:
    st.subheader("更新資料")
    st.success("v2 已加入 GitHub Actions 自動排程：週一至週五台灣時間 18:05 啟動。若官方資料尚未齊，會自動重試。")
    st.link_button("▶ 手動執行 GitHub Actions 更新",ACTIONS_URL,use_container_width=True)
    st.caption("手動執行：進入 Actions 後按 Run workflow；可留空日期跑當天，也可指定 YYYY-MM-DD。")
    st.divider()
    st.write("**臨時預覽（不會永久寫回 GitHub）**")
    ru=st.file_uploader("上傳最新 RS／市場廣度 xlsx",type=["xlsx"],key="rs_upload_widget")
    du=st.file_uploader("上傳最新 鴨嘴篩選結果 xlsx",type=["xlsx"],key="duck_upload_widget")
    c1,c2=st.columns(2)
    if c1.button("套用上傳資料",type="primary",use_container_width=True):
        changed=False
        if ru is not None: st.session_state["rs_uploaded_bytes"]=ru.getvalue(); changed=True
        if du is not None: st.session_state["duck_uploaded_bytes"]=du.getvalue(); changed=True
        if changed: _read_excel_path.clear(); st.rerun()
        else: st.warning("請先選擇至少一個 xlsx。")
    if c2.button("恢復 GitHub 內建資料",use_container_width=True):
        st.session_state.pop("rs_uploaded_bytes",None); st.session_state.pop("duck_uploaded_bytes",None); _read_excel_path.clear(); st.rerun()
    st.divider(); st.write("**下載目前正式結果**")
    d1,d2=st.columns(2)
    with open(DEFAULT_RS,"rb") as f: d1.download_button("下載 RS 最新結果",f.read(),file_name="rs_latest.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    with open(DEFAULT_DUCK,"rb") as f: d2.download_button("下載鴨嘴最新結果",f.read(),file_name="duck_latest.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

st.divider(); st.caption(f"正式版 v2｜RS：{rs_date}｜鴨嘴：{duck_date}｜量化篩選與市場廣度工具，不構成投資建議。")
