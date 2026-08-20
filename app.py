# -*- coding: utf-8 -*-
from __future__ import annotations

from io import BytesIO
from pathlib import Path
import datetime as dt
from zoneinfo import ZoneInfo
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
INTRADAY_BASELINE_FILE = APP_DIR / "intraday_baseline.pkl.gz"
ACTIONS_URL = "https://github.com/jan770610-dot/tw-stock-dashboard-test/actions/workflows/daily-update.yml"
LIVE_SCHEMA_VERSION = "v357-live-history-1"
INTRADAY_ENGINE_GENERATION = "3.5.6-auditfix-1"

st.set_page_config(page_title="台股分析中心", page_icon="📈", layout="wide", initial_sidebar_state="expanded")

# 部署新版時清除舊版盤中 session 快取。v3.5.4/3.5.5 曾共用同一組 key，
# 瀏覽器 session 可能在程式更新後仍保留舊 radar/override，造成看起來「價格不動」。
if st.session_state.get("_intraday_live_schema") != LIVE_SCHEMA_VERSION:
    for _k in list(st.session_state.keys()):
        if str(_k).startswith(("v357_", "v356_", "v355_", "v354_")) or _k == "v32_single_override":
            st.session_state.pop(_k, None)
    st.session_state["_intraday_live_schema"] = LIVE_SCHEMA_VERSION

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
.stage-flow {display:flex; flex-wrap:wrap; gap:7px; align-items:center; margin:.5rem 0 1rem;}
.stage-step {padding:7px 10px; border-radius:999px; background:rgba(120,120,120,.08); border:1px solid rgba(120,120,120,.16); font-size:.9rem;}
.stage-current {font-weight:700; border:2px solid currentColor; background:rgba(120,120,120,.14);}
.action-box {padding:1rem 1.1rem; border-radius:14px; border:1px solid rgba(120,120,120,.18); background:rgba(120,120,120,.05); margin:.7rem 0 1rem;}
.decision-grid {display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:10px; margin:.7rem 0 1rem;}
.decision-card {border:1px solid rgba(120,120,120,.18); border-radius:14px; padding:13px 14px; background:rgba(127,127,127,.035); min-height:108px;}
.decision-title {font-size:.82rem; opacity:.7; margin-bottom:6px;}
.decision-main {font-size:1.15rem; font-weight:700; line-height:1.35;}
.decision-sub {font-size:.84rem; opacity:.78; margin-top:7px; line-height:1.45;}
.focus-box {padding:.9rem 1rem; border-radius:12px; border-left:4px solid rgba(90,90,90,.55); background:rgba(120,120,120,.06); margin:.7rem 0 1rem; line-height:1.65;}
[data-testid="stDataFrame"] {border:1px solid rgba(120,120,120,.12); border-radius:10px; overflow:hidden;}
@media (max-width: 700px) {
  .block-container {padding-left:.7rem; padding-right:.7rem; padding-top:.65rem;}
  h1 {font-size:1.75rem !important;}
  h2 {font-size:1.28rem !important;}
  .kpi-grid {grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px;}
  .kpi-card {min-height:92px; padding:11px 12px;}
  .kpi-value {font-size:1.35rem;}
  .decision-grid {grid-template-columns:1fr 1fr; gap:8px;}
  .decision-card {padding:11px 12px; min-height:96px;}
  .decision-main {font-size:1rem;}
}
@media (max-width: 420px) {.kpi-grid,.decision-grid {grid-template-columns:1fr 1fr;}}
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
    try: recovery = read_sheet(source, "被遺忘資金候選")
    except Exception: recovery = pd.DataFrame()
    if daily.empty: raise ValueError("RS 檔案的『每日強勢股數量』沒有資料")
    if "日期" in daily.columns: daily["日期"] = normalize_date_series(daily["日期"])
    return daily, strong, extreme, recovery, daily.iloc[-1]


def duck_sheet(source, name):
    try: return read_sheet(source, name, header=1)
    except Exception: return pd.DataFrame()


def duck_data(source):
    names = ["全部符合","今日新進","今日退出","即將可能符合","鴨嘴×兩者皆有","鴨嘴×三率三升","鴨嘴×營收三增","估值買進候選","台指期","資料來源狀態","全市場整合篩選"]
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


@st.cache_data(show_spinner=False)
def _load_all_market_formal_rs(path_s: str, mtime_ns: int) -> pd.DataFrame:
    """從每日更新後建立的 intraday baseline 讀取全市場正式 RS。

    formal_rs 由 automation/build_intraday_baseline.py 直接用最新正式交易日
    的 250 交易日報酬全市場百分位計算，與 RS 主系統的排名定義一致。
    舊版 baseline 尚未有 formal_rs 時回傳空表，由既有強勢/復甦 RS 備援。
    """
    try:
        b = pd.read_pickle(path_s, compression="gzip")
    except Exception:
        return pd.DataFrame(columns=["代號", "RS_全市場正式", "RS正式日期"])
    if b is None or b.empty or "code" not in b.columns or "formal_rs" not in b.columns:
        return pd.DataFrame(columns=["代號", "RS_全市場正式", "RS正式日期"])
    z = b[[c for c in ["code", "formal_rs", "base_date"] if c in b.columns]].copy()
    z["代號"] = z["code"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    z["RS_全市場正式"] = pd.to_numeric(z["formal_rs"], errors="coerce")
    z["RS正式日期"] = z["base_date"].astype(str) if "base_date" in z.columns else ""
    return z[["代號", "RS_全市場正式", "RS正式日期"]].drop_duplicates("代號", keep="last")


def load_all_market_formal_rs() -> pd.DataFrame:
    if not INTRADAY_BASELINE_FILE.exists():
        return pd.DataFrame(columns=["代號", "RS_全市場正式", "RS正式日期"])
    return _load_all_market_formal_rs(str(INTRADAY_BASELINE_FILE), INTRADAY_BASELINE_FILE.stat().st_mtime_ns)


def card(label, value, delta=""):
    return f'<div class="kpi-card"><div class="kpi-label">{html.escape(str(label))}</div><div class="kpi-value">{html.escape(str(value))}</div><div class="kpi-delta">{html.escape(str(delta))}</div></div>'


def cards(items):
    st.markdown('<div class="kpi-grid">'+''.join(card(*x) for x in items)+'</div>', unsafe_allow_html=True)


def market_signal(stage):
    mapping={
        "低檔回升":"結構改善","中段回升":"結構轉強","高檔續強":"多頭擴散",
        "低檔續弱":"結構偏弱","中段走弱":"結構轉弱","高檔轉弱":"高檔退潮",
        "低檔整理":"低檔整理","中段整理":"結構整理","高檔整理":"高檔震盪"}
    return mapping.get(stage,"整理觀察")

def stage_plain(stage):
    return {
        "低檔續弱":"強勢股很少，而且仍在減少", "低檔整理":"強勢股偏少，暫時沒有明顯方向",
        "低檔回升":"強勢股仍在低水位，但正在增加", "中段回升":"市場廣度持續擴散，結構轉強",
        "中段整理":"多空拉鋸，等待方向", "中段走弱":"強勢股開始減少，市場正在降溫",
        "高檔續強":"市場雖熱，但攻擊力仍維持", "高檔整理":"高水位震盪，先觀察是否續強或退潮",
        "高檔轉弱":"市場仍在高位，但內部強勢股開始退潮"}.get(stage,"市場結構暫無明確方向")

def infer_action(stage,wade_score,left_alert,early_signal):
    level=risk_signal_level(left_alert)
    if level=="reduce": return "分批減碼／汰弱留強"
    if level=="watch": return "停止擴大曝險／風險觀察"
    try: score=float(wade_score)
    except Exception: score=None
    if stage in {"中段走弱","低檔續弱"} and (score is None or score<45): return "防守降低曝險"
    if text(early_signal,"—")!="—" or stage=="低檔回升": return "提高關注／分批試單"
    if stage in {"中段回升","高檔續強"} and (score is None or score>=60): return "偏多持有／強股續抱"
    if stage in {"低檔整理","中段整理","高檔整理"}: return "觀望等待／不追高"
    return "觀望等待"

def stage_flow_html(stage):
    steps=[("📉","弱勢",{"低檔續弱","中段走弱"}),("🧊","冰點",set()),("🌱","低檔回升",{"低檔回升"}),("📈","中段回升",{"中段回升"}),("🚀","高檔續強",{"高檔續強"}),("⚠️","高檔轉弱",{"高檔轉弱"})]
    # 低/中/高整理就放在最近的位置，不硬塞成買賣訊號。
    cur_alias={"低檔整理":"冰點","中段整理":"中段回升","高檔整理":"高檔續強"}.get(stage)
    parts=[]
    for icon,label,states in steps:
        current=(stage in states) or (cur_alias==label)
        parts.append(f'<span class="stage-step {"stage-current" if current else ""}">{icon} {label}{" ← 目前" if current else ""}</span>')
        if label!="高檔轉弱": parts.append('<span>→</span>')
    return '<div class="stage-flow">'+''.join(parts)+'</div>'



def action_level(stage, wade_score, market_lr_phase, left_alert, left_score=None, right_score=None, wade_change5=None):
    """把詳細建議濃縮成可直接執行的操作語言；左側與右側分開表達。"""
    risk_level=risk_signal_level(left_alert)
    if risk_level=="reduce" or "減碼" in text(market_lr_phase,""):
        return "⚠️ 分批減碼／汰弱留強"
    if risk_level=="watch" or "風險升高" in text(market_lr_phase,""):
        return "🟡 停止加碼／風險觀察"
    score=_num(wade_score)
    left=_num(left_score,0) or 0
    right=_num(right_score,0) or 0
    w5=_num(wade_change5,0) or 0
    if stage in {"低檔續弱","中段走弱"} and (score is None or score < 45):
        return "🛡️ 防守降低曝險"
    if "右側確認" in text(market_lr_phase,"") or (stage in {"中段回升","高檔續強"} and (score or 0)>=60 and right>=65):
        return "🚀 右側偏多持有"
    if "左轉右" in text(market_lr_phase,""):
        return "↗️ 左轉右逐步布局"
    if stage=="低檔回升" and left>=right+10:
        return "🌱 左側小部位試單｜等待右側確認"
    if "左側" in text(market_lr_phase,""):
        return "🌱 左側小部位試單"
    if "試單" in text(market_lr_phase,"") or right>=58:
        return "📈 右側小部位試單"
    if w5 < -8 and score is not None and score < 55:
        return "⏳ 觀望等待／先看5日結構修復"
    return "⏳ 觀望等待"


def _quality_label(score):
    return "高" if score>=80 else "中高" if score>=68 else "中" if score>=52 else "偏低"


def market_confidence(row, market_lr):
    """把「資料完整度」與「訊號一致度」拆開；不把大幅波動本身誤當成更高信心。"""
    fields=["Wade內部強度分數","上漲比例%","52週新高家數","52週新低家數","上市上漲比例%","上櫃上漲比例%","量價效率分數","權值同步分數","主流延續分數"]
    present=sum(_num(first_value(row,[f])) is not None for f in fields)
    completeness=max(0,min(100,present/len(fields)*100))

    stage=text(first_value(row,["市場階段"]),"")
    direction=text(first_value(row,["強勢股方向"]),"")
    wade=_num(first_value(row,["Wade內部強度分數"]),50) or 50
    w5=_num(first_value(row,["Wade分數5日變化"]),0) or 0
    adv=_num(first_value(row,["上漲比例%"]),50) or 50
    nh=_num(first_value(row,["52週新高家數"]),0) or 0
    nl=_num(first_value(row,["52週新低家數"]),0) or 0
    left=_num(market_lr.get("左側分數"),0) or 0
    right=_num(market_lr.get("右側分數"),0) or 0
    phase=text(market_lr.get("左右側階段"),"")

    consistency=50.0
    # 市場階段與 RS 方向是否同向。
    if ("回升" in stage and "增加" in direction) or ("續強" in stage and direction in {"增加","明顯增加","持平"}) or ("走弱" in stage and "減少" in direction):
        consistency += 15
    elif ("回升" in stage and "減少" in direction) or ("走弱" in stage and "增加" in direction):
        consistency -= 12

    # Wade 與上漲家數是否互相確認。
    if wade>=60 and adv>=55:
        consistency += 12
    elif wade<45 and adv<45:
        consistency += 10
    elif (wade>=60 and adv<48) or (wade<45 and adv>=55):
        consistency -= 10
    else:
        consistency += 2

    # 5日變化要看方向，不能用絕對值加分。
    if stage in {"低檔回升","中段回升","高檔續強"}:
        if w5>=5: consistency += 12
        elif w5<=-5: consistency -= 18
        else: consistency += 2
    elif stage in {"中段走弱","高檔轉弱","低檔續弱"}:
        if w5<=-5: consistency += 10
        elif w5>=5: consistency -= 10

    # 左右側階段與分數差距是否吻合。
    if "左側" in phase and left>=right+10:
        consistency += 8
    elif "右側" in phase and right>=left+10:
        consistency += 8
    elif "左轉右" in phase and abs(left-right)<=20:
        consistency += 6
    elif abs(left-right)<5:
        consistency -= 4

    if nh>nl: consistency += 4
    elif nl>nh*1.5 and nl>=5: consistency -= 5

    consistency=max(0,min(100,consistency))
    overall=max(0,min(100,completeness*0.4+consistency*0.6))
    return {
        "資料完整度": round(completeness,0),
        "資料完整度標籤": _quality_label(completeness),
        "訊號一致度": round(consistency,0),
        "訊號一致度標籤": _quality_label(consistency),
        "判讀品質": round(overall,0),
        "判讀品質標籤": _quality_label(overall),
    }


def metric_day_change(daily, col):
    if daily is None or len(daily)<2 or col not in daily.columns:
        return None
    a=_num(daily.iloc[-1].get(col)); b=_num(daily.iloc[-2].get(col))
    if a is None or b is None: return None
    return a-b


def signed_num(v, digits=1, suffix=""):
    x=_num(v)
    if x is None: return "—"
    return f"{x:+.{digits}f}{suffix}"


def wade_timing_summary(score, day_change, five_change):
    d=_num(day_change); f=_num(five_change)
    if d is None or f is None:
        return "Wade 多時間週期資料不足"
    if d>=5 and f<=-5:
        return "短線明顯反彈，但5日趨勢尚未完全扭轉"
    if d>=0 and f>=0:
        return "今日與5日同步改善，內部結構較一致"
    if d<0 and f>0:
        return "今日回落，但5日結構仍較前期改善"
    if d<=-5 and f<=-5:
        return "短線與5日同步轉弱，先提高防守"
    return "短線與5日訊號仍有分歧，等待進一步確認"


def signal_active(v):
    """事件型訊號是否啟動。空白、破折號與『未觸發』都視為未啟動。"""
    s=text(v,"—")
    return s not in {"—","-","未觸發","None","nan","NaN",""}


def signal_display(v):
    return text(v,"—") if signal_active(v) else "未觸發"


def signal_streak_stats(df, col, predicate=None):
    """統計事件型訊號連續天數。
    predicate 可用來只統計某一級事件，例如只算真正橘色『向上減碼』，不把黃色風險觀察混在一起。
    """
    empty={
        "目前連續":0,"段數":0,"中位數":None,"80百分位":None,"最長":None,
        "開始日期":None,"上次連續":0,"上次開始日期":None,"上次結束日期":None
    }
    if df is None or df.empty or col not in df.columns:
        return empty
    x=df.copy()
    if "日期" in x.columns:
        x=x.sort_values("日期").reset_index(drop=True)
        dates=normalize_date_series(x["日期"])
    else:
        x=x.reset_index(drop=True)
        dates=pd.Series([pd.NaT]*len(x))
    fn=predicate or signal_active
    flags=[]
    for v in x[col].tolist():
        try: flags.append(bool(fn(v)))
        except Exception: flags.append(False)
    runs=[]; run_start=None
    for i,on in enumerate(flags):
        if on and run_start is None:
            run_start=i
        if run_start is not None and ((not on) or i==len(flags)-1):
            end_i=(i if on and i==len(flags)-1 else i-1)
            runs.append((run_start,end_i,end_i-run_start+1))
            run_start=None
    current=0
    for on in reversed(flags):
        if on: current+=1
        else: break
    if not runs: return empty
    lens=pd.Series([r[2] for r in runs],dtype="float64")
    current_start=None
    if current>0:
        idx=len(flags)-current
        if idx < len(dates) and not pd.isna(dates.iloc[idx]):
            current_start=dates.iloc[idx].strftime("%Y-%m-%d")
    prior=None
    if current>0 and len(runs)>=2: prior=runs[-2]
    elif current==0 and runs: prior=runs[-1]
    prev_len=prior[2] if prior else 0
    prev_start=prev_end=None
    if prior:
        a,b,_=prior
        if a<len(dates) and not pd.isna(dates.iloc[a]): prev_start=dates.iloc[a].strftime("%Y-%m-%d")
        if b<len(dates) and not pd.isna(dates.iloc[b]): prev_end=dates.iloc[b].strftime("%Y-%m-%d")
    return {
        "目前連續":int(current),"段數":int(len(runs)),
        "中位數":float(lens.median()),"80百分位":float(lens.quantile(.8)),"最長":int(lens.max()),
        "開始日期":current_start,"上次連續":int(prev_len),"上次開始日期":prev_start,"上次結束日期":prev_end,
    }


def signal_stat_text(stats):
    if not stats or not stats.get("段數"):
        return "歷史樣本不足"
    med=stats.get("中位數"); p80=stats.get("80百分位")
    return f"歷史中位 {med:.0f}天｜80% 約≤{p80:.0f}天｜樣本 {stats['段數']} 段"


def signal_current_text(stats):
    cur=int((stats or {}).get("目前連續",0) or 0)
    if cur>0:
        start=(stats or {}).get("開始日期")
        return f"目前第 {cur} 天" + (f"｜起始 {start}" if start else "")
    prev=int((stats or {}).get("上次連續",0) or 0)
    end=(stats or {}).get("上次結束日期")
    if prev>0:
        return f"目前未觸發｜上次 {prev} 天" + (f"（至 {end}）" if end else "")
    return "目前未觸發"


def risk_signal_level(v):
    """把既有左側減碼警示拆成『黃色風險觀察』與真正『橘色向上減碼』。"""
    s=text(v,"—")
    if not signal_active(s): return "none"
    if "🟠" in s or "左側減碼" in s or "高檔內部轉弱" in s:
        return "reduce"
    return "watch"


def early_signal_phase(stats):
    cur=int((stats or {}).get("目前連續",0) or 0)
    if cur<=0: return "未觸發"
    if cur==1: return "🟢 第1天｜初現"
    if cur==2: return "🟢🟢 第2天｜確認"
    return f"🚀 第{cur}天｜持續轉強"


def risk_signal_phase(value, risk_stats, reduce_stats):
    level=risk_signal_level(value)
    if level=="none": return "未觸發"
    if level=="watch":
        cur=max(1,int((risk_stats or {}).get("目前連續",0) or 0))
        return f"🟡 第{cur}天｜風險升高觀察"
    cur=max(1,int((reduce_stats or {}).get("目前連續",0) or 0))
    if cur==1: return "🟠 第1天｜向上減碼觀察"
    if cur==2: return "🟠 第2天｜風險升高"
    return f"🔴 第{cur}天｜持續退潮"


def signal_event_backtest(df, signal_col, price_col, mode="up", horizons=(5,10,20)):
    """以每一段訊號的第一天為事件日，計算後續指數報酬與區間最大回撤。
    mode='up'：成功定義為期末報酬 > 0；mode='down'：成功定義為期末報酬 < 0。
    """
    cols=["訊號","市場","期間","樣本數","平均報酬%","符合方向%","平均區間最大回撤%"]
    if df is None or df.empty or signal_col not in df.columns or price_col not in df.columns:
        return pd.DataFrame(columns=cols)
    x=df.copy()
    if "日期" in x.columns: x=x.sort_values("日期").reset_index(drop=True)
    prices=pd.to_numeric(x[price_col],errors="coerce")
    flags=x[signal_col].map(signal_active)
    starts=[i for i,on in enumerate(flags) if on and (i==0 or not bool(flags.iloc[i-1]))]
    rows=[]
    for h in horizons:
        rets=[]; dds=[]
        for i in starts:
            if i+h>=len(x): continue
            p0=prices.iloc[i]; p1=prices.iloc[i+h]
            if pd.isna(p0) or pd.isna(p1) or p0<=0: continue
            path=prices.iloc[i:i+h+1].dropna()
            if path.empty: continue
            ret=(float(p1)/float(p0)-1)*100
            dd=(path.astype(float)/float(p0)-1).min()*100
            rets.append(ret); dds.append(float(dd))
        if not rets: continue
        s=pd.Series(rets,dtype="float64")
        success=(s>0).mean()*100 if mode=="up" else (s<0).mean()*100
        rows.append({
            "期間":f"{h}日","樣本數":len(rets),"平均報酬%":round(float(s.mean()),2),
            "符合方向%":round(float(success),1),"平均區間最大回撤%":round(float(pd.Series(dds).mean()),2)
        })
    return pd.DataFrame(rows)


def market_exposure_plan(row, market_lr, quality, early_stats, risk_value, reduce_stats):
    """把市場狀態轉成『股票總曝險參考區間』。
    這是相對風險預算，不是總資產配置，也不是最佳化後的固定答案。
    """
    stage=text(first_value(row,["市場階段"]),"")
    wade=_num(first_value(row,["Wade內部強度分數"]),50) or 50
    w5=_num(first_value(row,["Wade分數5日變化"]),0) or 0
    left=_num(market_lr.get("左側分數"),0) or 0
    right=_num(market_lr.get("右側分數"),0) or 0
    consistency=_num((quality or {}).get("訊號一致度"),50) or 50
    early_days=int((early_stats or {}).get("目前連續",0) or 0)
    reduce_days=int((reduce_stats or {}).get("目前連續",0) or 0)
    risk=risk_signal_level(risk_value)

    base={
        "低檔續弱":(10,25),"低檔整理":(15,30),"低檔回升":(25,40),
        "中段走弱":(20,35),"中段整理":(30,50),"中段回升":(50,70),
        "高檔轉弱":(20,40),"高檔整理":(40,60),"高檔續強":(60,80),
    }.get(stage,(25,45))
    lo,hi=base
    reasons=[]

    if early_days>=2 and stage in {"低檔回升","中段回升"}:
        lo+=5; hi+=10; reasons.append(f"早期轉強已連續{early_days}天")
    elif early_days==1:
        hi+=5; reasons.append("早期轉強初現")
    if right>=70 and wade>=60:
        lo+=5; hi+=5; reasons.append("右側與Wade同步確認")
    elif left>=65 and right<60:
        reasons.append("左側優勢高於右側，仍採分批")
    if w5<=-8:
        lo-=5; hi-=5; reasons.append("Wade 5日結構仍弱")
    if consistency<45:
        lo-=5; hi-=10; reasons.append("訊號一致度偏低")
    elif consistency>=68:
        hi+=5; reasons.append("訊號一致度較高")

    if risk=="watch":
        hi-=10; reasons.append("黃色風險觀察：停止擴大曝險")
    elif risk=="reduce":
        lo=min(lo,30); hi=min(hi,45)
        if reduce_days>=2: lo-=5; hi-=10
        if reduce_days>=3: lo-=5; hi-=5
        reasons.append(f"向上減碼訊號第{max(1,reduce_days)}天")

    lo=max(0,min(90,int(round(lo/5)*5)))
    hi=max(lo+5,min(95,int(round(hi/5)*5)))
    if hi<=25: label="低曝險"
    elif hi<=45: label="中低曝險"
    elif hi<=65: label="中等曝險"
    elif hi<=80: label="中高曝險"
    else: label="高曝險"
    return {
        "下限":lo,"上限":hi,"顯示":f"{lo}–{hi}%","層級":label,
        "說明":"；".join(reasons[:4]) or "依市場階段與左右側分數設定",
        "註記":"以你自行設定的股票最大風險預算為100%，不是總資產配置。"
    }


EXIT_BACKTEST_V12 = {
    "一般鴨嘴": {
        "策略": "過熱分數≥80", "門檻": 80,
        "觸發中位報酬%": 9.375, "觸發勝率%": 87.93,
        "觸發中位持有日": 7, "有效觸發率%": 52.75,
    },
    "可交易強股": {
        "策略": "過熱分數≥70", "門檻": 70,
        "觸發中位報酬%": 10.118, "觸發勝率%": 94.77,
        "觸發中位持有日": 13, "有效觸發率%": 50.33,
    },
    "RS85領頭強股": {
        "策略": "過熱分數≥90", "門檻": 90,
        "觸發中位報酬%": 11.596, "觸發勝率%": 97.86,
        "觸發中位持有日": 10, "有效觸發率%": 59.04,
    },
}


ENTRY_BACKTEST_V10 = {
    "預備80＋RS70＋市場改善": {
        "排名": 1, "樣本": 3756, "20日中位報酬%": 0.000,
        "20日MFE%": 7.854, "20日MAE%": -5.853, "20日達5率%": 63.68,
    },
    "右側65＋市場改善": {
        "排名": 2, "樣本": 7464, "20日中位報酬%": 0.000,
        "20日MFE%": 5.769, "20日MAE%": -4.385, "20日達5率%": 54.62,
    },
    "正式鴨嘴＋市場改善": {
        "排名": 3, "樣本": 6408, "20日中位報酬%": 0.000,
        "20日MFE%": 5.461, "20日MAE%": -4.205, "20日達5率%": 53.00,
    },
}

FAILURE_EXIT_V10 = {
    "規則": "收盤相對成本≤-8%",
    "跨情境排名": 1,
    "跨情境平均排名": 1.67,
    "10分位改善pp": 4.596,
    "最大回撤改善pp": 0.639,
    "停損率%": 37.94,
    "停損後20日達原進場加5率%": 15.18,
    "中位報酬改善pp": -3.289,
}


def current_market_entry_gate(row, market_lr, master=None):
    """把進場回測的『市場廣度改善』轉成目前正式頁面可即時計算的代理。

    歷史回測使用上漲比例、MA20上方比例與廣度改善；目前正式頁面可直接取得
    上漲比例、全市場 MA20 狀態、強勢股方向/Wade 5日變化，因此以同方向代理。
    這是操作閘門，不宣稱與歷史回測欄位逐項完全相同。
    """
    adv=_num(first_value(row,["上漲比例%"]),0) or 0
    direction=text(first_value(row,["強勢股方向"]),"")
    w5=_num(first_value(row,["Wade分數5日變化"]),0) or 0
    stage=text(first_value(row,["市場階段"]),"")
    alert=text(first_value(row,["左側減碼警示"]),"—")
    risk=risk_signal_level(alert)

    ma20_pct=None
    if master is not None and not master.empty and "收盤" in master.columns and "MA20" in master.columns:
        c=pd.to_numeric(master["收盤"],errors="coerce")
        m=pd.to_numeric(master["MA20"],errors="coerce")
        valid=c.notna() & m.notna() & (m>0)
        if valid.any():
            ma20_pct=float((c[valid] > m[valid]).mean()*100.0)

    improving=(direction in {"增加","明顯增加"}) or w5>0 or stage in {"低檔回升","中段回升"}
    breadth_ok=adv>=50 and (ma20_pct is None or ma20_pct>=50) and improving and risk!="reduce"
    ma20_text="MA20廣度資料不足" if ma20_pct is None else f"MA20上方 {ma20_pct:.1f}%"
    label=("✅ 市場廣度改善" if breadth_ok else "⏳ 市場廣度未確認")
    detail=f"上漲 {adv:.1f}%｜{ma20_text}｜強勢股{direction or '—'}｜Wade5日 {w5:+.1f}"
    return {"允許":bool(breadth_ok),"標籤":label,"說明":detail,"MA20上方比例%":ma20_pct}


def stock_entry_plan(complete, rs, right, heat, formal, today_new, pre,
                     market_entry_ok=False, market_entry_label="市場廣度未確認"):
    """把進場 v1.0 研究轉成『試單 → 主要進場 → 正式確認』流程。"""
    comp=_num(complete,0) or 0
    rsn=_num(rs)
    r=_num(right,0) or 0
    h=_num(heat,0) or 0
    pre_s=text(pre,"")

    if h>=70:
        return {
            "進場階段":"⚠️ 過熱不新增", "進場動作":"先不新增部位；等待過熱降溫或新的右側觸發",
            "進場回測參考":"進場回測顯示等待 RS85 才追入沒有優勢；過熱與進場條件分開管理。"
        }
    if not market_entry_ok:
        hold_note="；既有強勢持股可依獲利管理續抱" if formal else ""
        return {
            "進場階段":"⏳ 市場等待", "進場動作":f"個股條件可觀察，但市場廣度尚未確認，暫不把它升級成新買點{hold_note}",
            "進場回測參考":f"進場 v1.0 前三名都含市場改善條件；目前：{market_entry_label}。"
        }

    if comp>=80 and rsn is not None and rsn>=70 and not formal:
        ref=ENTRY_BACKTEST_V10["預備80＋RS70＋市場改善"]
        return {
            "進場階段":"🌱 預備試單", "進場動作":"可先用小部位試單；右側≥65再升級基本部位，不一次重押",
            "進場回測參考":f"進場v1.0第{ref['排名']}名｜樣本{ref['樣本']:,}｜20日MFE中位 {ref['20日MFE%']:.2f}%｜20日達+5% {ref['20日達5率%']:.1f}%"
        }
    if r>=65:
        ref=ENTRY_BACKTEST_V10["右側65＋市場改善"]
        extra="；RS≥85不是必要進場門檻" if rsn is not None and rsn>=85 else ""
        return {
            "進場階段":"📈 主要進場區", "進場動作":f"右側已證明，可建立／補到基本部位；仍採分批，不追滿{extra}",
            "進場回測參考":f"進場v1.0第{ref['排名']}名｜樣本{ref['樣本']:,}｜20日MFE中位 {ref['20日MFE%']:.2f}%｜20日MAE中位 {ref['20日MAE%']:.2f}%"
        }
    if formal or today_new:
        ref=ENTRY_BACKTEST_V10["正式鴨嘴＋市場改善"]
        return {
            "進場階段":"🚀 正式確認", "進場動作":"型態已正式確認；已有前置部位以續抱為主，未持有者避免只因『正式新進』追滿",
            "進場回測參考":f"正式鴨嘴＋市場改善為進場v1.0第{ref['排名']}名；更早的預備80／右側65訊號排名較前。"
        }
    if r>=58 or "A級" in pre_s:
        return {
            "進場階段":"👀 接近觸發", "進場動作":"先列觀察；等RS≥70且完成度80，或右側≥65再升級",
            "進場回測參考":"目前只接近回測較佳進場條件，尚未完整觸發。"
        }
    return {
        "進場階段":"⏳ 等待", "進場動作":"尚未進入回測較佳進場區；等待結構與RS同步改善",
        "進場回測參考":"—"
    }


def stock_failure_risk_plan(close, cost=None, exit_today=False, right=None, ma20=None, ma20chg=None):
    """失敗交易 v1.0：-8% 只用『收盤相對平均成本』確認；結構轉弱只做預警。"""
    c=_num(close)
    k=_num(cost)
    r=_num(right)
    m=_num(ma20)
    mchg=_num(ma20chg)
    structure_warn=bool(exit_today or (r is not None and r<45) or (
        c is not None and m is not None and c<m and mchg is not None and mchg<=0
    ))
    ref=FAILURE_EXIT_V10
    ref_text=(f"失敗退出v1.0第1名｜跨情境平均排名 {ref['跨情境平均排名']:.2f}｜"
              f"尾端10%改善 {ref['10分位改善pp']:+.2f}pp｜停損後20日又達原進場+5% {ref['停損後20日達原進場加5率%']:.1f}%")

    if k is not None and k>0 and c is not None and c>0:
        dd=(c/k-1.0)*100.0
        stop_price=k*0.92
        if dd<=-8:
            return {"失敗風控":"🔴 收盤-8%硬停損","風控動作":"依回測最後防線退出；不要再用MA20/右側拖延","成本報酬%":dd,"硬停損價":stop_price,"風控參考":ref_text}
        if dd<=-6:
            return {"失敗風控":"🟠 接近-8%硬停損","風控動作":"停止加碼；收盤若到成本-8%即退出","成本報酬%":dd,"硬停損價":stop_price,"風控參考":ref_text}
        if structure_warn:
            return {"失敗風控":"🟡 結構轉弱預警","風控動作":"暫停加碼並觀察；MA20/右側轉弱本身不再等同硬性賣出","成本報酬%":dd,"硬停損價":stop_price,"風控參考":ref_text}
        return {"失敗風控":"🟢 尚未觸及失敗線","風控動作":"持續依進場／獲利流程管理；收盤-8%才是最後硬停損","成本報酬%":dd,"硬停損價":stop_price,"風控參考":ref_text}

    if structure_warn:
        stage="🟡 結構轉弱預警"
        action="目前只列預警，不因MA20／右側單一轉弱直接全退；填入成本價後才能判斷-8%硬停損"
    else:
        stage="⚪ 未設定成本價"
        action="填入你的平均成本價後，系統才會啟用『收盤-8%』硬停損判讀"
    return {"失敗風控":stage,"風控動作":action,"成本報酬%":None,"硬停損價":None,"風控參考":ref_text}


def stock_profit_exit_plan(heat, rs, right, formal, today_new, exit_today,
                           has_hot_date=False, ma20chg=None, spreadchg=None):
    """v3.5：過熱負責獲利管理；結構轉弱本身只預警，不再當失敗交易硬停損。"""
    h=_num(heat,0) or 0
    rsn=_num(rs)
    r=_num(right,0) or 0
    m20=_num(ma20chg)
    spr=_num(spreadchg)

    if rsn is not None and rsn >= 85:
        profile="RS85領頭強股"
    elif r >= 58:
        profile="可交易強股"
    else:
        profile="一般鴨嘴"
    ref=EXIT_BACKTEST_V12[profile]

    weak_count=sum([
        m20 is not None and m20 <= 0,
        spr is not None and spr <= 0,
        r < 60,
    ])

    # 已經是贏家、曾進入過熱後才使用二階段退潮；這不是一般失敗交易停損。
    if has_hot_date and h < 70 and weak_count >= 2:
        stage="🔴 過熱後退潮"
        action="提高減碼幅度／退出剩餘趨勢部位；這是贏家退潮，不是一般MA20停損"
    elif profile == "RS85領頭強股":
        if h >= 90:
            stage="🟠 領頭強股積極鎖利"
            action="分批減碼／提高移動停利；不再新增部位"
        elif h >= 85:
            stage="🟡 領頭強股高熱觀察"
            action="停止加碼、續抱觀察；接近回測較佳鎖利區"
        elif h >= 70:
            stage="🟢 領頭強股續抱"
            action="停止追價但不急著退出；RS85領頭股可容忍較高過熱"
        else:
            stage="🟢 趨勢持有"
            action="尚未進入領頭強股獲利保護區；依右側結構續抱"
    elif profile == "可交易強股":
        if h >= 80:
            stage="🟠 積極鎖利"
            action="已超過第一鎖利門檻；分批減碼並提高移動停利"
        elif h >= 70:
            stage="🟡 開始鎖利"
            action="進入回測較佳第一獲利保護區；先分批鎖利，不必一次全退"
        elif h >= 60:
            stage="🟡 接近獲利保護區"
            action="停止追價，準備進入分批鎖利模式"
        else:
            stage="🟢 趨勢持有"
            action="尚未進入獲利保護區；依右側結構續抱"
    else:
        if h >= 80:
            stage="🟠 積極鎖利"
            action="進入一般鴨嘴回測較佳獲利區；分批減碼／提高移動停利"
        elif h >= 70:
            stage="🟡 高熱觀察"
            action="已偏熱但尚未到一般鴨嘴主要鎖利門檻；停止追價並準備減碼"
        elif exit_today or not formal and not today_new:
            stage="🟡 結構轉弱觀察" if exit_today else "⚪ 尚未進入獲利管理"
            action="結構轉弱只做預警；一般失敗交易改以成本收盤-8%作最後硬停損" if exit_today else "先等右側／正式結構確認，再啟動獲利保護門檻"
        else:
            stage="🟢 趨勢持有"
            action="尚未進入一般鴨嘴獲利保護區；依結構續抱"

    hist=(
        f"v1.2｜{profile}：{ref['策略']}；真正觸發中位報酬 {ref['觸發中位報酬%']:.2f}%、"
        f"勝率 {ref['觸發勝率%']:.1f}%、中位第 {ref['觸發中位持有日']} 交易日、"
        f"60日內觸發率 {ref['有效觸發率%']:.1f}%"
    )
    secondary="過熱後若 RS 自持有高點回落 ≥5，再把減碼升級；目前頁面未保存完整持有期 RS 高點，因此先列為下一確認條件。"
    return {
        "獲利階段":stage,
        "獲利動作":action,
        "獲利門檻":f"過熱 {ref['門檻']}",
        "獲利管理類型":profile,
        "回測參考":hist,
        "二階段確認":secondary,
    }

def stock_action_plan(left, right, heat, rs, complete, formal, pre, exit_today, today_new, rec,
                      ma20=None, ma20chg=0, market_entry_ok=False):
    """v3.5：進場、獲利、失敗風控分工，不再把結構轉弱直接等同賣出。"""
    l=_num(left,0) or 0; r=_num(right,0) or 0; h=_num(heat,0) or 0
    rsn=_num(rs); comp=_num(complete,0) or 0; recv=_num(rec,0) or 0
    pre_s=text(pre,"")
    leader=rsn is not None and rsn>=85

    if leader and h>=90 and formal:
        tag="🟠 分批減碼"
    elif leader and h>=85 and formal:
        tag="⚠️ 停止追價"
    elif (not leader) and h>=80 and formal:
        tag="🟠 分批減碼"
    elif (not leader) and h>=70 and formal:
        tag="⚠️ 停止追價"
    elif exit_today:
        tag="🟡 結構警示"
    elif market_entry_ok and comp>=80 and rsn is not None and rsn>=70 and not formal and h<70:
        tag="🌱 預備試單"
    elif market_entry_ok and r>=65 and h<70:
        tag="📈 主要進場"
    elif formal and r>=65 and h<70:
        tag="🚀 偏多持有"
    elif r>=58 and h<70:
        tag="👀 觀察"
    elif recv>=45 or "A級" in pre_s or "B級" in pre_s or l>=60:
        tag="👀 觀察"
    else:
        tag="⏳ 等待"

    need=[]
    if h>=70:
        need=["先不新增","等待過熱降溫／新的右側觸發"]
    elif not market_entry_ok:
        need=["市場廣度改善","右側維持／升到65以上"]
    elif comp<80:
        need=["鴨嘴完成度≥80","RS≥70"]
    elif rsn is None or rsn<70:
        need=["RS≥70","右側≥65"]
    elif r<65:
        need=["右側≥65"]
    elif not formal:
        need=["先小部位試單","右側持續證明再加碼"]
    else:
        need=["回檔守住結構","不在過熱區追價"]
    add_condition="＋".join(need[:3])

    invalid_condition="硬停損：收盤≤平均成本-8%（單檔填成本後判斷）／MA20、右側或鴨嘴退出僅作結構預警"

    if exit_today:
        change="⚠️ 今日降級：鴨嘴退出（結構預警，不自動全退）"
    elif today_new:
        change="⬆️ 今日升級：正式鴨嘴新進"
    elif formal:
        change="➡️ 維持右側結構"
    elif "A級" in pre_s:
        change="⬆️ 接近升級：A級預備"
    elif "B級" in pre_s:
        change="➡️ 培育中：B級預備"
    else:
        change="—"
    return tag, add_condition, invalid_condition, change

def overheat_label(score):
    x=_num(score,0) or 0
    if x>=85: return "極高"
    if x>=70: return "高"
    if x>=55: return "偏高"
    if x>=35: return "中"
    return "低"


def classify_recovery_stage(left, right, overheat, rec_score=None):
    """復甦候選三層：觀察 → 左側試單 → 接近右側確認。"""
    l=_num(left,0) or 0; r=_num(right,0) or 0; oh=_num(overheat,0) or 0; rec=_num(rec_score,0) or 0
    if r>=58 or (r>=52 and l>=55):
        return "↗️ 接近右側確認"
    if l>=60 and rec>=50 and oh<70:
        return "🌱 可左側試單"
    return "👀 觀察名單"


def short_market_trigger(row, market_lr):
    stage=text(first_value(row,["市場階段"]),"")
    if stage=="低檔回升": return "Wade≥60＋上漲≥55%＋RS持續增加"
    if stage=="中段回升": return "右側≥70＋主流延續不降"
    if stage in {"高檔續強","高檔整理"}: return "留意Wade 5日轉弱＋強勢股減少"
    if stage in {"高檔轉弱","中段走弱"}: return "Wade回升＋上漲比例重回50～55%"
    return "至少2項：廣度、Wade、價格同向確認"


def day_change_summary(daily):
    if daily is None or len(daily)<2:
        return "前日比較資料不足"
    a=daily.iloc[-1]; b=daily.iloc[-2]
    parts=[]
    pairs=[("強勢股檔數","強勢股",0,"檔"),("Wade內部強度分數","Wade",1,"分"),("上漲比例%","上漲比例",1,"pct")]
    for col,label,digits,suffix in pairs:
        if col not in daily.columns: continue
        x=_num(a.get(col)); y=_num(b.get(col))
        if x is None or y is None: continue
        d=x-y
        sign="+" if d>0 else ""
        parts.append(f"{label} {sign}{d:.{digits}f}{suffix}")
    return "；".join(parts) if parts else "前日比較資料不足"


def next_market_trigger(row, market_lr):
    stage=text(first_value(row,["市場階段"]),"")
    score=_num(first_value(row,["Wade內部強度分數"]),50) or 50
    adv=_num(first_value(row,["上漲比例%"]),50) or 50
    w5=_num(first_value(row,["Wade分數5日變化"]),0) or 0
    right=_num(market_lr.get("右側分數"),0) or 0
    if stage=="低檔回升":
        return "升級條件：Wade ≥60、上漲比例 ≥55%，且 RS 維持增加 → 觀察是否進入中段回升／右側試單。"
    if stage=="中段回升":
        return "續強條件：右側分數 ≥70 且主流延續不降；若 Wade 5日轉負且廣度收縮，先降一級。"
    if stage in {"高檔續強","高檔整理"}:
        return "風險條件：Wade 5日快速轉弱、強勢股減少或新低增加 → 啟動向上減碼觀察。"
    if stage in {"高檔轉弱","中段走弱"}:
        return "重新轉強條件：Wade 回升、上漲比例重新站回 50～55%，且強勢股方向轉增加。"
    if score<45 or adv<45:
        return "先等內部止跌：Wade 回到 45～50 以上、上漲家數改善，再考慮提高部位。"
    if right>=58:
        return "價格已開始證明；下一步看右側分數能否持續升到 70 以上，而不是單日突破。"
    return "等待市場廣度、Wade 與價格趨勢出現至少兩項同方向確認。"


def decision_cards(items):
    parts=[]
    for title,main,sub in items:
        parts.append(f'<div class="decision-card"><div class="decision-title">{html.escape(str(title))}</div><div class="decision-main">{html.escape(str(main))}</div><div class="decision-sub">{html.escape(str(sub))}</div></div>')
    st.markdown('<div class="decision-grid">'+''.join(parts)+'</div>',unsafe_allow_html=True)


def _num(v, default=None):
    try:
        if pd.isna(v): return default
        x=float(v)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _boolish(v):
    if isinstance(v,bool): return v
    if v is None: return False
    try:
        if pd.isna(v): return False
    except Exception: pass
    return str(v).strip().lower() in {"true","1","是","yes","y"}


def market_lr_assessment(row):
    """把既有市場廣度/Wade 指標轉成左側、右側兩條獨立軸。
    這是本系統量化代理，不是 Wade 官方公式。
    """
    stage=text(first_value(row,["市場階段"]),"")
    direction=text(first_value(row,["強勢股方向"]),"")
    speed=text(first_value(row,["變化速度"]),"")
    full=_num(first_value(row,["全期回顧百分位%"]),50)
    roll=_num(first_value(row,["近3年滾動百分位%"]),full)
    wade=_num(first_value(row,["Wade內部強度分數"]),50)
    w5=_num(first_value(row,["Wade分數5日變化"]),0)
    spread=_num(first_value(row,["5日-20日廣度差"]),0)
    adv=_num(first_value(row,["上漲比例%"]),50)
    nh=_num(first_value(row,["52週新高家數"]),0)
    nl=_num(first_value(row,["52週新低家數"]),0)
    sync=_num(first_value(row,["權值同步分數"]),50)
    lead=_num(first_value(row,["主流延續分數"]),50)
    early=text(first_value(row,["早期轉強訊號"]),"—")
    reduce=text(first_value(row,["左側減碼警示"]),"—")

    # 左側：重點是「位置仍低＋內部開始改善」，而不是單純越跌越買。
    cold=max(0,min(45,(45-min(full,roll))*1.15))
    improve=0
    if direction in {"增加","明顯增加"}: improve += 18
    if "加速" in speed and "增加" in speed: improve += 10
    elif "增加" in speed or "放緩" in speed: improve += 5
    if spread>0: improve += min(12,4+spread*4)
    if w5>0: improve += min(10,w5*1.2)
    if early!="—": improve += 10
    left=max(0,min(100,cold+improve))

    # 右側：市場價格/廣度已經證明趨勢，偏向確認後才增加曝險。
    right=0
    right += max(0,min(35,(wade-30)*0.78))
    right += {"中段回升":20,"高檔續強":25,"低檔回升":10,"中段整理":8,"高檔整理":10}.get(stage,0)
    if direction=="明顯增加": right+=12
    elif direction=="增加": right+=8
    elif direction=="持平": right+=3
    if adv>=55: right+=8
    elif adv>=50: right+=4
    if nh>nl: right+=6
    right += max(0,min(7,(sync-40)*0.18))
    right += max(0,min(7,(lead-40)*0.18))
    right=max(0,min(100,right))

    reduce_level=risk_signal_level(reduce)
    if reduce_level=="reduce" or (stage=="高檔轉弱" and w5<=-5):
        phase="⚠️ 向上減碼觀察"
        action="停止追高；持有強股可續抱，但分批鎖利／汰弱留強"
    elif reduce_level=="watch":
        phase="🟡 風險升高觀察"
        action="先停止擴大曝險與追價；尚未視為正式減碼，等待是否升級成高檔內部轉弱"
    elif stage in {"低檔續弱","中段走弱"} and wade<38 and left<55:
        phase="🛡️ 防守／等待"
        action="降低曝險；等市場內部止跌與轉強再提高部位"
    elif left>=60 and wade<40 and w5<0:
        phase="🔍 左側研究／等待觸發"
        action="位置偏低但短線內部仍弱；先研究與觀察，等廣度重新改善再試單"
    elif left>=65 and right>=60:
        phase="↗️ 左轉右確認"
        action="可由試單逐步建立基本部位；後續仍以價格持續確認為前提"
    elif right>=72:
        phase="🚀 右側確認"
        action="偏多持有／強勢部位續抱；新增部位以回檔不破結構或新觸發點為主"
    elif left>=65:
        phase="🌱 左側機會觀察"
        action="可提高關注或小部位分批試單，但不要把低檔等同已落底"
    elif right>=58:
        phase="📈 右側試單區"
        action="趨勢開始證明，可小部位試單；價格繼續證明再增加曝險"
    else:
        phase="⏳ 中性等待"
        action="目前左右側優勢都不明顯，等待更好的位置或更明確觸發"

    reasons=[]
    if min(full,roll)<=25: reasons.append(f"市場廣度仍在低水位(P{min(full,roll):.0f})")
    if direction in {"增加","明顯增加"}: reasons.append(f"強勢股方向{direction}")
    if spread>0: reasons.append(f"5日廣度高於20日 {spread:.2f}pct")
    if wade>=60: reasons.append(f"Wade內部強度 {wade:.0f}")
    elif wade<40: reasons.append(f"Wade內部強度偏弱 {wade:.0f}")
    if w5>=5: reasons.append(f"5日內部強度改善 +{w5:.1f}")
    elif w5<=-5: reasons.append(f"5日內部強度轉弱 {w5:.1f}")
    if reduce_level=="reduce": reasons.append("高檔退潮／向上減碼條件觸發")
    elif reduce_level=="watch": reasons.append("市場風險升高，但尚未達正式減碼級")
    return {"左側分數":round(left,1),"右側分數":round(right,1),"左右側階段":phase,"左右側建議":action,"判讀依據":"；".join(reasons[:5]) or "目前訊號分散"}



def index_lr_assessment(row, prefix, label):
    close=_num(first_value(row,[f"{label}指數收盤", f"{prefix}_close"]))
    ma20=_num(first_value(row,[f"{label}MA20", f"{prefix}_ma20"]))
    ma60=_num(first_value(row,[f"{label}MA60", f"{prefix}_ma60"]))
    ma20c=_num(first_value(row,[f"{label}MA20較前日", f"{prefix}_ma20_change"]),0) or 0
    ma60c=_num(first_value(row,[f"{label}MA60較前日", f"{prefix}_ma60_change"]),0) or 0
    ret=_num(first_value(row,["加權指數報酬%" if prefix=="twii" else "櫃買指數報酬%"]),0) or 0
    breadth=_num(first_value(row,["上市上漲比例%" if prefix=="twii" else "上櫃上漲比例%"]),50) or 50
    wade=_num(first_value(row,["Wade內部強度分數"]),50) or 50
    direction=text(first_value(row,["強勢股方向"]),"")
    left=0; right=0; reasons=[]
    if close is not None and ma20 is not None:
        dist=(close/ma20-1)*100 if ma20 else 0
        if -8<=dist<=3: left+=20; reasons.append(f"接近20日線({dist:+.1f}%)")
        if close>=ma20: right+=20
        if ma60 is not None:
            if close>=ma20>=ma60: right+=25; reasons.append("站上20/60日線")
            elif close<ma60: left+=10
        if ma20c>0: right+=10
        if ma60c>0: right+=8
    else:
        reasons.append("指數均線待下一次更新補齊")
    if breadth<40: left+=15
    elif breadth>=55: right+=15
    if direction in {"增加","明顯增加"}: left+=10; right+=8
    if wade<40: left+=10
    elif wade>=60: right+=12
    if ret>0: right+=5
    left=max(0,min(100,left)); right=max(0,min(100,right))
    if right>=70: phase="🚀 右側確認"; action="指數趨勢已確認，偏多看待；避免追在過熱擴張段"
    elif left>=55 and right>=50: phase="↗️ 左轉右"; action="低檔修復正轉向確認，可由試單逐步提高曝險"
    elif left>=55: phase="🌱 左側觀察"; action="位置/廣度有修復條件，但趨勢尚未完整確認，採分批與小部位"
    elif right>=50: phase="📈 右側試單"; action="價格開始證明，先小部位，持續站穩再提高曝險"
    else: phase="⏳ 等待"; action="尚無明顯左右側優勢"
    return {"名稱":label,"收盤":close,"左側分數":round(left,1),"右側分數":round(right,1),"階段":phase,"建議":action,"依據":"；".join(reasons)}

def build_stock_lr_table(master, strong, recovery, formal_rs=None, market_entry_ok=False, market_entry_label="市場廣度未確認"):
    """個股左右側量化代理。
    左側偏重估值、基本面與跌深修復；右側偏重價格趨勢、鴨嘴完成度與 RS。
    """
    if master is None or master.empty: return pd.DataFrame()
    x=master.copy()
    if "代號" not in x.columns: return pd.DataFrame()
    x["代號"]=x["代號"].astype(str).str.replace(r"\.0$","",regex=True)

    # RS 優先順序：全市場正式 RS → 強勢名單 RS → 復甦候選 RS。
    # v3.3.1 起全市場正式 RS 由 intraday_baseline 每日盤後一起產生，
    # 因此不再只有「強勢／復甦」股票才看得到 RS。
    if formal_rs is not None and not formal_rs.empty and "代號" in formal_rs.columns:
        a0=formal_rs[[c for c in ["代號","RS_全市場正式","RS正式日期"] if c in formal_rs.columns]].copy()
        a0["代號"]=a0["代號"].astype(str).str.replace(r"\.0$","",regex=True)
        a0=a0.drop_duplicates("代號",keep="last")
        x=x.merge(a0,on="代號",how="left")
    if strong is not None and not strong.empty and "代號" in strong.columns:
        a=strong[[c for c in ["代號","RS"] if c in strong.columns]].copy()
        a["代號"]=a["代號"].astype(str).str.replace(r"\.0$","",regex=True)
        a=a.rename(columns={"RS":"RS_強勢"}).drop_duplicates("代號")
        x=x.merge(a,on="代號",how="left")
    if recovery is not None and not recovery.empty and "代號" in recovery.columns:
        b=recovery[[c for c in ["代號","RS","復甦分數","距52週高點%","5日報酬%","20日報酬%"] if c in recovery.columns]].copy()
        b["代號"]=b["代號"].astype(str).str.replace(r"\.0$","",regex=True)
        b=b.rename(columns={"RS":"RS_復甦"}).drop_duplicates("代號")
        x=x.merge(b,on="代號",how="left")
    if "RS_全市場正式" not in x.columns: x["RS_全市場正式"]=pd.NA
    if "RS正式日期" not in x.columns: x["RS正式日期"]=""
    if "RS_強勢" not in x.columns: x["RS_強勢"]=pd.NA
    if "RS_復甦" not in x.columns: x["RS_復甦"]=pd.NA
    rs_all=pd.to_numeric(x["RS_全市場正式"],errors="coerce")
    rs_strong=pd.to_numeric(x["RS_強勢"],errors="coerce")
    rs_recovery=pd.to_numeric(x["RS_復甦"],errors="coerce")
    x["RS判讀"]=rs_all.combine_first(rs_strong).combine_first(rs_recovery)
    x["RS來源"]=pd.Series("資料不足",index=x.index,dtype="object")
    x.loc[rs_recovery.notna(),"RS來源"]="正式復甦候選"
    x.loc[rs_strong.notna(),"RS來源"]="正式強勢名單"
    x.loc[rs_all.notna(),"RS來源"]="全市場正式"

    rows=[]
    for _,r in x.iterrows():
        discount=_num(r.get("折價率%"))
        rec=_num(r.get("復甦分數"))
        rs=_num(r.get("RS判讀"))
        complete=_num(r.get("鴨嘴完成度%"),0) or 0
        stage=str(r.get("鴨嘴階段") or "")
        pre=str(r.get("預備鴨嘴狀態") or "")
        val=str(r.get("估值狀態") or "")
        tr=str(r.get("三率狀態") or "")
        rev=str(r.get("營收狀態") or "")
        epsg=_num(r.get("EPS年增率%"))
        overheat="過熱" in stage or "不建議追價" in stage
        exit_today=_boolish(r.get("今日退出")) or stage=="今日退出"
        today_new=_boolish(r.get("今日新進")) or stage.startswith("新進")
        formal=_boolish(r.get("正式鴨嘴")) or stage.startswith("持續符合") or stage.startswith("新進")
        ma20chg=_num(r.get("MA20較前日"),0) or 0
        ma60chg=_num(r.get("MA60較前日"),0) or 0
        spreadchg=_num(r.get("開口較前日"),0) or 0
        gap20=_num(r.get("月線乖離率%"),0) or 0
        first_hot=r.get("首次過熱日期")
        try:
            has_hot_date=not pd.isna(first_hot) and str(first_hot).strip() not in {"","—","-"}
        except Exception:
            has_hot_date=bool(first_hot)

        # 過熱程度 0-100：與右側強度分開。右側可很強，同時也可能很熱。
        # 月線乖離是主體；RS/正式型態/既有過熱標記只做加成，不直接壓低右側分數。
        heat=max(0,min(70,max(0,gap20)*5.0))
        if rs is not None and rs>=95: heat+=12
        elif rs is not None and rs>=85: heat+=8
        if formal or complete>=100: heat+=8
        # 若原始鴨嘴狀態已明確標示「過熱／不建議追價」，過熱程度至少進入高檔區，避免文字與分數互相矛盾。
        if overheat: heat=max(70,heat+18)
        if has_hot_date: heat+=8
        heat=max(0,min(100,heat))

        # 左側 0-100：估值 35 + 修復 30 + 基本面 25 + 接近右側 10。
        left=0; lreason=[]
        if discount is not None:
            if discount>=20: left+=35; lreason.append(f"折價{discount:.0f}%")
            elif discount>=10: left+=28; lreason.append(f"進入估值買進區({discount:.0f}%)")
            elif discount>=0: left+=16; lreason.append("估值合理偏低")
            elif discount<=-20: left-=8; lreason.append("估值偏貴")
        elif val in {"明顯低估","買進區","合理偏低"}:
            left+={"明顯低估":35,"買進區":28,"合理偏低":16}[val]; lreason.append(val)
        if rec is not None:
            left+=max(12,min(30,rec*0.32)); lreason.append(f"復甦分數{rec:.0f}")
        if tr=="三率三升": left+=10; lreason.append("三率三升")
        if rev=="營收三增": left+=10; lreason.append("營收三增")
        if epsg is not None and epsg>0: left+=5; lreason.append("EPS年增為正")
        if "A級" in pre: left+=10; lreason.append("預備鴨嘴A級")
        elif "B級" in pre: left+=5; lreason.append("預備鴨嘴B級")
        if overheat: left-=20
        if formal and complete>=100: left-=5  # 已轉成右側，不再把它當純左側。
        left=max(0,min(100,left))

        # 右側 0-100：價格型態/完成度 55 + RS 25 + 觸發/延續 20。
        right=max(0,min(55,complete*0.55)); rreason=[]
        if complete>=80: rreason.append(f"鴨嘴完成{complete:.0f}%")
        if rs is not None:
            if rs>=85: right+=25; rreason.append(f"RS {rs:.0f}")
            elif rs>=70: right+=18; rreason.append(f"RS {rs:.0f}")
            elif rs>=50: right+=10; rreason.append(f"RS {rs:.0f}")
        if stage.startswith("新進"): right+=12; rreason.append("今日右側觸發/新進")
        elif formal: right+=8; rreason.append("正式鴨嘴")
        elif "A級" in pre: right+=6; rreason.append("接近右側觸發")
        if ma20chg>0 and ma60chg>0 and spreadchg>0: right+=6; rreason.append("均線與開口同步改善")
        if overheat: rreason.append("過熱，不宜追價")
        if exit_today: right-=30; rreason.append("今日退出")
        right=max(0,min(100,right))

        # 右側分數算完後，用完整資訊重建獲利／出場階段。
        profit_plan=stock_profit_exit_plan(
            heat,rs,right,formal,today_new,exit_today,has_hot_date=has_hot_date,
            ma20chg=_num(r.get("MA20較前日")),spreadchg=_num(r.get("開口較前日"))
        )
        entry_plan=stock_entry_plan(
            complete,rs,right,heat,formal,today_new,pre,
            market_entry_ok=market_entry_ok,market_entry_label=market_entry_label
        )

        if exit_today:
            phase="🟡 右側結構轉弱"
            action="停止新增部位並提高警覺；不再因單一MA20／右側／鴨嘴退出直接全退，硬停損改由成本-8%確認"
        elif profit_plan["獲利階段"].startswith("🔴"):
            phase="🔴 贏家退潮管理"
            action=profit_plan["獲利動作"]
        elif entry_plan["進場階段"].startswith("🌱"):
            phase="🌱 預備試單"
            action=entry_plan["進場動作"]
        elif entry_plan["進場階段"].startswith("📈"):
            phase="📈 右側主要進場"
            action=entry_plan["進場動作"]
        elif formal and right>=65:
            phase="🚀 右側確認"
            action="已有部位以持有為主；新增部位仍看市場廣度與過熱，不因正式鴨嘴追滿"
        elif left>=65:
            phase="🔍 左側研究名單"
            action="可研究但不把低檔直接當買點；本版新買點優先等預備80+RS70+市場改善或右側65"
        elif right>=58:
            phase="👀 接近右側觸發"
            action="價格開始證明，但尚未到主要進場條件；等右側≥65與市場廣度確認"
        else:
            phase="⏳ 尚未形成優勢"
            action="等待結構、RS與市場環境同步改善"

        if profit_plan["獲利階段"].startswith(("🟡","🟠")):
            action=f"{action}｜獲利管理：{profit_plan['獲利動作']}"

        op_tag, add_cond, invalid_cond, change_tag = stock_action_plan(
            left,right,heat,rs,complete,formal,pre,exit_today,today_new,rec,
            ma20=_num(r.get("MA20")),ma20chg=ma20chg,market_entry_ok=market_entry_ok
        )
        rows.append({
            "左側分數":round(left,1),"右側分數":round(right,1),
            "過熱程度":round(heat,1),"過熱等級":overheat_label(heat),
            "操作標籤":op_tag,"今日升降級":change_tag,
            "進場階段":entry_plan["進場階段"],"進場動作":entry_plan["進場動作"],
            "進場回測參考":entry_plan["進場回測參考"],"市場進場閘門":market_entry_label,
            "獲利階段":profit_plan["獲利階段"],"獲利動作":profit_plan["獲利動作"],
            "獲利門檻":profit_plan["獲利門檻"],"獲利管理類型":profit_plan["獲利管理類型"],
            "回測參考":profit_plan["回測參考"],"二階段確認":profit_plan["二階段確認"],
            "加碼條件":add_cond,"失效條件":invalid_cond,
            "左右側階段":phase,"系統建議":action,
            "左側依據":"；".join(lreason[:5]) or "—","右側依據":"；".join(rreason[:5]) or "—"
        })
    z=pd.concat([x.reset_index(drop=True),pd.DataFrame(rows)],axis=1)
    z["決策排序分數"]=z[["左側分數","右側分數"]].max(axis=1)
    action_rank={"🌱 預備試單":0,"📈 主要進場":1,"🚀 偏多持有":2,"👀 觀察":3,"⚠️ 停止追價":4,"🟠 分批減碼":5,"🟡 結構警示":6,"⏳ 等待":7}
    z["操作排序"]=z["操作標籤"].map(action_rank).fillna(9)
    return z


def _clean_cell(v) -> str:
    try:
        if pd.isna(v): return ""
    except Exception:
        pass
    s=str(v).strip()
    return "" if s in {"", "—", "-", "nan", "None", "<NA>"} else s


def _short_text(v, limit=72) -> str:
    s=_clean_cell(v) or "—"
    return s if len(s)<=limit else s[:max(1,limit-1)]+"…"


def stock_duck_summary(row) -> str:
    stage=_clean_cell(row.get("鴨嘴階段"))
    pre_s=_clean_cell(row.get("預備鴨嘴狀態"))
    complete=_num(row.get("鴨嘴完成度%"))
    if stage:
        if "退出" in stage: return "🛑 今日退出"
        if "新進" in stage: return "🔥 今日新進"
        if "持續符合" in stage or "正式" in stage: return "✅ 正式符合"
        if "過熱" in stage: return "⚠️ 正式／過熱"
        return _short_text(stage,18)
    if "A級" in pre_s: return "🌱 預備 A級"
    if "B級" in pre_s: return "🌱 預備 B級"
    if complete is not None and complete>=80: return f"🌱 完成 {complete:.0f}%"
    if complete is not None and complete>0: return f"進度 {complete:.0f}%"
    return "—"


def stock_cultivation_summary(row) -> str:
    cat=_clean_cell(row.get("培育中心類別"))
    tr=_clean_cell(row.get("三率狀態"))
    rev=_clean_cell(row.get("營收狀態"))
    tr_ok=(tr=="三率三升") or ("三率三升" in tr)
    rev_ok=(rev=="營收三增") or ("營收三增" in rev)
    if tr_ok and rev_ok: return "✅ 三率＋營收"
    if tr_ok: return "✅ 三率三升"
    if rev_ok: return "✅ 營收三增"
    if cat: return _short_text(cat,18)
    return "—"


def stock_valuation_summary(row) -> str:
    discount=_num(row.get("折價率%"))
    val=_clean_cell(row.get("估值狀態"))
    if discount is not None:
        if discount>=10: return f"折價 {discount:.0f}%"
        if discount>=0: return f"折價 {discount:.0f}%"
        return f"溢價 {abs(discount):.0f}%"
    return _short_text(val,16) if val else "—"


def stock_one_line(row) -> str:
    action=_clean_cell(row.get("操作標籤")) or "⏳ 等待"
    phase=_clean_cell(row.get("左右側階段")) or "階段未明"
    duck=stock_duck_summary(row)
    cult=stock_cultivation_summary(row)
    heat=_num(row.get("過熱程度"))
    heat_s=f"{heat:.0f}/100" if heat is not None else "—"
    entry=_clean_cell(row.get("進場階段")) or "⏳ 等待"
    profit=_clean_cell(row.get("獲利階段")) or "⚪ 尚未進入獲利管理"
    return f"{action}｜進場：{entry}｜獲利：{profit}｜{phase}；鴨嘴：{duck}；培育：{cult}；過熱：{heat_s}。"


def stock_decision_compact(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame()
    out=pd.DataFrame(index=df.index)
    for c in ["代號","名稱","市場","收盤"]:
        if c in df.columns: out[c]=df[c]
    out["操作"]=df.apply(lambda r:_clean_cell(r.get("操作標籤")) or "⏳ 等待",axis=1)
    if "進場階段" in df.columns:
        out["進場"]=df.apply(lambda r:_clean_cell(r.get("進場階段")) or "—",axis=1)
    if "獲利階段" in df.columns:
        out["獲利/出場"]=df.apply(lambda r:_clean_cell(r.get("獲利階段")) or "—",axis=1)
    out["左右側"]=df.apply(lambda r:f"{_num(r.get('左側分數'),0):.0f} / {_num(r.get('右側分數'),0):.0f}",axis=1)
    out["過熱"]=df.apply(lambda r:f"{_num(r.get('過熱程度'),0):.0f}",axis=1)
    out["鴨嘴"]=df.apply(stock_duck_summary,axis=1)
    out["培育"]=df.apply(stock_cultivation_summary,axis=1)
    if "RS判讀" in df.columns:
        out["RS"]=df.apply(lambda r: (f"{_num(r.get('RS判讀')):.1f}" if _num(r.get('RS判讀')) is not None else "—"),axis=1)
    out["估值"]=df.apply(stock_valuation_summary,axis=1)
    if "今日升降級" in df.columns: out["今日變化"]=df["今日升降級"].map(lambda v:_clean_cell(v) or "—")
    return out.reset_index(drop=True)


def stock_choice_label(row) -> str:
    return f"{_clean_cell(row.get('代號'))} {_clean_cell(row.get('名稱'))}｜{_clean_cell(row.get('操作標籤')) or '⏳ 等待'}｜{_clean_cell(row.get('左右側階段')) or '—'}"


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


def wade_chart(df, days=120):
    if "日期" not in df.columns or "Wade內部強度分數" not in df.columns: return None
    x=df[["日期","Wade內部強度分數"]].copy()
    x["Wade內部強度分數"]=pd.to_numeric(x["Wade內部強度分數"],errors="coerce")
    x=x.dropna().tail(days)
    if x.empty: return None
    base=alt.Chart(x).encode(x=alt.X("日期:T",title="日期"))
    line=base.mark_line(point=True).encode(y=alt.Y("Wade內部強度分數:Q",title="市場內部強度",scale=alt.Scale(domain=[0,100])),tooltip=[alt.Tooltip("日期:T"),alt.Tooltip("Wade內部強度分數:Q",format=".1f")])
    rules=alt.Chart(pd.DataFrame({"y":[35,50,65,75]})).mark_rule(strokeDash=[5,5],opacity=.35).encode(y="y:Q")
    return (line+rules).properties(height=320)


rs_source=_xlsx_source(DEFAULT_RS,"rs_uploaded_bytes")
duck_source=_xlsx_source(DEFAULT_DUCK,"duck_uploaded_bytes")
try:
    daily,strong,extreme,recovery,rs_latest=rs_data(rs_source)
    all_ok,duck_new,duck_exit,pre,both,three_rate,three_rev,value_candidates,tx,source_status,master=duck_data(duck_source)
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
wade_score=first_value(rs_latest,["Wade內部強度分數"])
wade_change5=first_value(rs_latest,["Wade分數5日變化"])
wade_state=text(first_value(rs_latest,["Wade市場狀態"]))
left_alert=text(first_value(rs_latest,["左側減碼警示"]),"—")
early_signal=text(first_value(rs_latest,["早期轉強訊號"]),"—")
advance_count=first_value(rs_latest,["上漲家數"])
decline_count=first_value(rs_latest,["下跌家數"])
advance_ratio=first_value(rs_latest,["上漲比例%"])
limit_up=first_value(rs_latest,["漲停近似家數"])
limit_down=first_value(rs_latest,["跌停近似家數"])
new_high=first_value(rs_latest,["52週新高家數"])
new_low=first_value(rs_latest,["52週新低家數"])
twse_adv=first_value(rs_latest,["上市上漲比例%"])
tpex_adv=first_value(rs_latest,["上櫃上漲比例%"])
amount_ratio20=first_value(rs_latest,["成交金額20日比%"])
mean_return=first_value(rs_latest,["全市場等權平均報酬%"])
twse_mean_return=first_value(rs_latest,["上市等權平均報酬%"])
tpex_mean_return=first_value(rs_latest,["上櫃等權平均報酬%"])
twii_ret=first_value(rs_latest,["加權指數報酬%"])
twoii_ret=first_value(rs_latest,["櫃買指數報酬%"])
leader_retention=first_value(rs_latest,["主流延續率%"])
new_leaders=first_value(rs_latest,["新強勢領頭股數"])
volume_eff=first_value(rs_latest,["量價效率分數"])
index_sync=first_value(rs_latest,["權值同步分數"])
tpex_rel=first_value(rs_latest,["櫃買相對強弱分數"])
leadership_score=first_value(rs_latest,["主流延續分數"])
market_summary_plain=text(first_value(rs_latest,["市場總結白話"]), market_signal(stage)+"："+stage_plain(stage))
operation_action=text(first_value(rs_latest,["操作建議"]), infer_action(stage,wade_score,left_alert,early_signal))
if risk_signal_level(left_alert)=="watch": operation_action="停止擴大曝險／風險觀察"
elif risk_signal_level(left_alert)=="reduce": operation_action="分批減碼／汰弱留強"
market_lr=market_lr_assessment(rs_latest)
quality=market_confidence(rs_latest,market_lr)
wade_day_change=metric_day_change(daily,"Wade內部強度分數")
wade_timing_text=wade_timing_summary(wade_score,wade_day_change,wade_change5)
operation_level=action_level(
    stage,wade_score,market_lr["左右側階段"],left_alert,
    market_lr.get("左側分數"),market_lr.get("右側分數"),wade_change5
)
day_change_text=day_change_summary(daily)
next_trigger=next_market_trigger(rs_latest,market_lr)
short_trigger=short_market_trigger(rs_latest,market_lr)
twii_lr=index_lr_assessment(rs_latest,"twii","加權")
twoii_lr=index_lr_assessment(rs_latest,"twoii","櫃買")
formal_rs_all=load_all_market_formal_rs()
entry_market_gate=current_market_entry_gate(rs_latest,market_lr,master)
stock_lr=build_stock_lr_table(
    master,strong,recovery,formal_rs_all,
    market_entry_ok=entry_market_gate["允許"],market_entry_label=entry_market_gate["標籤"]
)
pre_status=pre.get("預備鴨嘴狀態",pd.Series(dtype=str)).fillna("").astype(str) if not pre.empty else pd.Series(dtype=str)
pre_a=int(pre_status.str.contains("A級").sum()) if not pre_status.empty else 0
pre_b=int(pre_status.str.contains("B級").sum()) if not pre_status.empty else 0
update_status=load_update_status()
left_signal_stats=signal_streak_stats(daily,"左側減碼警示")
left_reduce_stats=signal_streak_stats(daily,"左側減碼警示",lambda v: risk_signal_level(v)=="reduce")
early_signal_stats=signal_streak_stats(daily,"早期轉強訊號")
left_alert_show=signal_display(left_alert)
early_signal_show=signal_display(early_signal)
left_signal_phase=risk_signal_phase(left_alert,left_signal_stats,left_reduce_stats)
early_signal_phase=early_signal_phase(early_signal_stats)
exposure_plan=market_exposure_plan(rs_latest,market_lr,quality,early_signal_stats,left_alert,left_reduce_stats)

# ===== v3.2 單一儀表板：背景掃描 + 前台只讀快照 =====
TAIPEI_TZ = ZoneInfo("Asia/Taipei")
now_tpe = dt.datetime.now(TAIPEI_TZ)
today_tpe = now_tpe.strftime("%Y-%m-%d")
if st.session_state.get("v356_intraday_hot_day") != today_tpe:
    st.session_state["v356_intraday_hot_day"] = today_tpe
    st.session_state["v356_intraday_hot_codes"] = set()
is_weekday = now_tpe.weekday() < 5
is_market_session = is_weekday and dt.time(9, 0) <= now_tpe.time() <= dt.time(13, 30)
formal_today_ready = latest_date == today_tpe

if formal_today_ready:
    dashboard_mode = "formal"
elif is_market_session:
    dashboard_mode = "intraday"
elif is_weekday and now_tpe.time() > dt.time(13, 30):
    dashboard_mode = "close_trial"
else:
    dashboard_mode = "previous_close"


def _trial_exposure_reference(snap: dict) -> tuple[str, str]:
    stage_now = str(snap.get("stage") or "")
    lo, hi = {
        "低檔續弱": (10, 25), "低檔整理": (15, 30), "低檔回升": (25, 45),
        "中段走弱": (20, 35), "中段整理": (30, 50), "中段回升": (45, 65),
        "高檔轉弱": (20, 40), "高檔整理": (40, 60), "高檔續強": (55, 75),
    }.get(stage_now, (25, 45))
    w = _num(snap.get("wade"), 50) or 50
    risk_now = str(snap.get("risk") or "")
    if w >= 65:
        lo += 5; hi += 5
    elif w < 45:
        lo -= 5; hi -= 5
    if risk_now.startswith("🟡"):
        hi -= 10
    elif risk_now.startswith("🟠"):
        lo = min(lo, 25); hi = min(hi, 40)
    lo = max(0, min(90, int(round(lo / 5) * 5)))
    hi = max(lo + 5, min(95, int(round(hi / 5) * 5)))
    label = "盤中試算" if dashboard_mode == "intraday" else "收盤試算"
    return f"{lo}–{hi}%", label


def _metric_delta(cur, prev, digits=1, suffix=""):
    a = _num(cur); b = _num(prev)
    if a is None or b is None:
        return None
    d = a - b
    return f"較{latest_date}正式 {d:+.{digits}f}{suffix}"


def _fmt_trial(v, digits=1, suffix=""):
    x = _num(v)
    return "—" if x is None else f"{x:,.{digits}f}{suffix}"


def _norm_stock_code(v) -> str:
    return str(v).replace(".0", "").strip()


def _codes_from_df(df: pd.DataFrame) -> set[str]:
    if df is None or df.empty or "代號" not in df.columns:
        return set()
    return set(df["代號"].map(_norm_stock_code))


def _event_table(stocks: pd.DataFrame, codes: set[str], limit=30) -> pd.DataFrame:
    if stocks is None or stocks.empty or not codes:
        return pd.DataFrame()
    z = stocks[stocks["代號"].astype(str).isin(codes)].copy()
    order_cols = [c for c in ["盤中RS", "漲跌幅%"] if c in z.columns]
    if order_cols:
        z = z.sort_values(order_cols, ascending=[False] * len(order_cols), na_position="last")
    show = [c for c in [
        "代號", "名稱", "市場", "盤中價", "漲跌幅%", "盤中RS", "盤中強勢",
        "強勢條件通過數", "強勢尚缺條件", "即將強勢", "盤中鴨嘴價格結構",
        "盤中MA20", "盤中MA60", "月線乖離率%"
    ] if c in z.columns]
    return z[show].head(limit)


def _intraday_trade_radar_table(snap: dict, codes: set[str] | None = None) -> pd.DataFrame:
    """用與盤中單檔查詢完全相同的 v3.5 交易規則，重算全市場盤中操作分類。

    主交易雷達只把 MIS z 真實成交價列為可操作清單；委託簿/前收參考價仍保留
    在資料中，但不會因參考價跨過門檻而製造假的進場或減碼訊號。
    """
    stocks=snap.get("stocks") if snap else None
    if stocks is None or stocks.empty:
        return pd.DataFrame()

    live=stocks.copy()
    live["代號"]=live["代號"].map(_norm_stock_code)
    if codes is not None:
        wanted={_norm_stock_code(c) for c in codes}
        live=live[live["代號"].isin(wanted)].copy()

    formal_map={}
    if stock_lr is not None and not stock_lr.empty and "代號" in stock_lr.columns:
        f=stock_lr.copy()
        f["代號"]=f["代號"].map(_norm_stock_code)
        f=f.drop_duplicates("代號",keep="last")
        formal_map={str(r.get("代號")):r for r in f.to_dict("records")}

    rows=[]
    for lr in live.to_dict("records"):
        code=_norm_stock_code(lr.get("代號"))
        formal_row=formal_map.get(code,{})
        try:
            plan=_intraday_stock_trade_plan(lr,formal_row,snap,cost=None)
        except Exception:
            # 單一股票資料不足不應讓整個雷達消失。
            continue
        rows.append({
            "代號":code,
            "名稱":lr.get("名稱",""),
            "市場":lr.get("市場",""),
            "盤中價":_num(lr.get("盤中價")),
            "漲跌幅%":_num(lr.get("漲跌幅%")),
            "盤中RS":_num(lr.get("盤中RS")),
            "價格來源":_clean_cell(lr.get("價格來源")) or "—",
            "價格為估計":_boolish(lr.get("價格為估計")),
            "MIS報價時間":_clean_cell(lr.get("MIS報價時間")) or "—",
            "快速追蹤時間":_clean_cell(lr.get("快速追蹤時間")) or _clean_cell(lr.get("快速抓取完成時間")) or "",
            "快速追蹤epoch":_num(lr.get("快速追蹤epoch")),
            "快速追蹤批次":_num(lr.get("快速追蹤批次")),
            "行情引擎版本":_clean_cell(lr.get("行情引擎版本")) or _clean_cell(snap.get("engine_version")) or "—",
            "盤中操作":plan.get("盤中操作","⏳ 等待"),
            "盤中動作":plan.get("盤中動作","—"),
            "進場階段":plan.get("進場階段","—"),
            "獲利階段":plan.get("獲利階段","—"),
            "失敗風控":plan.get("失敗風控","—"),
            "右側分數":_num(plan.get("右側分數")),
            "過熱程度":_num(plan.get("過熱程度")),
            "鴨嘴完成度%":_num(plan.get("鴨嘴完成度%")),
            "盤中正式鴨嘴":bool(plan.get("盤中正式鴨嘴",False)),
            "盤中今日新進":bool(plan.get("盤中今日新進",False)),
            "盤中結構退出":bool(plan.get("盤中結構退出",False)),
            "市場閘門":plan.get("市場閘門","—"),
        })

    out=pd.DataFrame(rows)
    if out.empty:
        return out
    out["成交狀態"]=out["價格為估計"].map(lambda v:"⚪ 參考價" if bool(v) else "✅ MIS z成交")
    out["資料層級"]=out["快速追蹤時間"].map(lambda v:"⚡ 快速追蹤" if _clean_cell(v) else "🌐 全市場快照")
    return out


def _decorate_live_freshness(radar: pd.DataFrame, bg_state: dict, snap: dict) -> pd.DataFrame:
    """顯示每一列最後實際抓取時間；價格沒變時也能分辨『沒有新成交』與『系統沒更新』。"""
    if radar is None or radar.empty:
        return radar
    out=radar.copy()
    now=dt.datetime.now(TAIPEI_TZ)
    fast_interval=int(bg_state.get("fast_interval_seconds",10) or 10)
    full_time=_clean_cell(bg_state.get("last_completed")) or _clean_cell(snap.get("snapshot_time"))
    upd=[]; ages=[]; states=[]; levels=[]
    for _,r in out.iterrows():
        fast_t=_clean_cell(r.get("快速追蹤時間"))
        t=fast_t or full_time
        level="⚡ 快速追蹤" if fast_t else "🌐 全市場快照"
        age=None
        if t:
            try:
                tt=pd.Timestamp(t).to_pydatetime()
                if tt.tzinfo is None:
                    tt=tt.replace(tzinfo=TAIPEI_TZ)
                age=max(0.0,(now-tt).total_seconds())
            except Exception:
                age=None
        if age is None:
            state="⚠️ 無更新時間"
        elif level.startswith("⚡") and age <= max(25,fast_interval*3):
            state="✅ 快速行情正常"
        elif level.startswith("🌐") and age <= 75:
            state="🟡 等待首輪快刷"
        else:
            state="⚠️ 行情逾時"
        upd.append(t or "—"); ages.append(round(age,1) if age is not None else None); states.append(state); levels.append(level)
    out["資料更新時間"]=upd
    out["資料年齡秒"]=ages
    out["資料狀態"]=states
    out["資料層級"]=levels
    return out


def _snapshot_with_fast_rows(snap: dict, bg_state: dict) -> dict:
    """Overlay reactive watched-stock rows on the latest full-market snapshot."""
    fast=bg_state.get("fast_rows")
    if fast is None or not isinstance(fast,pd.DataFrame) or fast.empty:
        return snap
    stocks=snap.get("stocks")
    if stocks is None or stocks.empty or "代號" not in stocks.columns:
        return snap
    base=stocks.copy()
    base["代號"]=base["代號"].map(_norm_stock_code)
    f=fast.copy()
    f["代號"]=f["代號"].map(_norm_stock_code)
    base=base.set_index("代號")
    f=f.set_index("代號")
    for col in f.columns:
        if col not in base.columns:
            base[col]=pd.NA
    common=base.index.intersection(f.index)
    if len(common):
        for col in f.columns:
            base.loc[common,col]=f.loc[common,col]
    out=dict(snap)
    out["stocks"]=base.reset_index()
    return out


def _get_intraday_trade_radar(snap: dict, bg_state: dict) -> pd.DataFrame:
    """Full-market discovery on pulse scan; reactive versions recalc every watched row."""
    full_key=(str(latest_date),int(bg_state.get("version",0) or 0))
    cache=st.session_state.get("v356_trade_radar_full_cache")
    if not (isinstance(cache,dict) and cache.get("key")==full_key and isinstance(cache.get("data"),pd.DataFrame)):
        base_data=_intraday_trade_radar_table(snap)
        cache={"key":full_key,"data":base_data}
        st.session_state["v356_trade_radar_full_cache"]=cache
    base_data=cache["data"]

    fast=bg_state.get("fast_rows")
    fast_ver=int(bg_state.get("fast_version",0) or 0)
    if fast is None or not isinstance(fast,pd.DataFrame) or fast.empty:
        return base_data
    compose_key=(full_key,fast_ver)
    cc=st.session_state.get("v356_trade_radar_fast_cache")
    if isinstance(cc,dict) and cc.get("key")==compose_key and isinstance(cc.get("data"),pd.DataFrame):
        return cc["data"]

    effective=_snapshot_with_fast_rows(snap,bg_state)
    returned_codes=set(fast["代號"].astype(str).map(_norm_stock_code)) if "代號" in fast.columns else set()
    requested_codes={_norm_stock_code(c) for c in (bg_state.get("fast_watch_codes") or []) if _norm_stock_code(c)}
    refreshed=_intraday_trade_radar_table(effective,codes=returned_codes)
    if refreshed.empty:
        data=base_data.copy()
    else:
        # 只用實際 fast store 裡的列覆蓋；fast store 在引擎端是逐碼累積，不會因單批少回股票而整列消失。
        data=pd.concat([base_data[~base_data["代號"].astype(str).isin(returned_codes)],refreshed],ignore_index=True)
    if requested_codes:
        missing=requested_codes-returned_codes
        if missing:
            m=data["代號"].astype(str).isin(missing)
            data.loc[m,"資料層級"]="⏳ 快速行情待回傳"
    st.session_state["v356_trade_radar_fast_cache"]={"key":compose_key,"data":data}
    return data


def _signal_bucket(row) -> str:
    profit=str(row.get("獲利階段") or "")
    entry=str(row.get("進場階段") or "")
    op=str(row.get("盤中操作") or "")
    if profit.startswith("🔴"): return "🔴 贏家退潮"
    if profit.startswith("🟠"): return "🟠 積極鎖利"
    if profit.startswith("🟡") and "結構轉弱" not in profit: return "🟡 獲利保護"
    if entry.startswith("📈"): return "📈 主要進場"
    if entry.startswith("🌱"): return "🌱 回測試單"
    if op.startswith("🚀"): return "🚀 趨勢確認持有"
    return ""


def _signal_data_issue(row) -> str:
    """Return a data-quality reason that makes a row unsuitable as a live actionable signal."""
    if _boolish(row.get("價格為估計")):
        return "MIS 暫無 z 真實成交價（目前僅參考價）"
    state=_clean_cell(row.get("資料狀態"))
    if state.startswith("⚠️"):
        return state
    return ""


def _signal_leave_reason(old_bucket: str, row) -> str:
    """Explain why a previously active bucket is no longer active."""
    issue=_signal_data_issue(row)
    if issue:
        return issue

    new_bucket=_signal_bucket(row)
    if new_bucket and new_bucket!=old_bucket:
        return f"分類轉換：{old_bucket} → {new_bucket}"

    right=_num(row.get("右側分數"))
    rs=_num(row.get("盤中RS"))
    heat=_num(row.get("過熱程度"))
    complete=_num(row.get("鴨嘴完成度%"))
    gate=_clean_cell(row.get("市場閘門"))
    live_duck=_boolish(row.get("盤中正式鴨嘴"))
    operation=_clean_cell(row.get("盤中操作"))
    reasons=[]

    if old_bucket.startswith("📈"):
        if right is not None and right < 65:
            reasons.append(f"右側 {right:.1f}<65")
        if "未確認" in gate:
            reasons.append("市場廣度閘門未確認")
        if heat is not None and heat >= 70:
            reasons.append(f"過熱 {heat:.1f}，不再屬進場區")
    elif old_bucket.startswith("🌱"):
        if complete is not None and complete < 80:
            reasons.append(f"鴨嘴完成度 {complete:.0f}%<80%")
        if rs is not None and rs < 70:
            reasons.append(f"RS {rs:.1f}<70")
        if "未確認" in gate:
            reasons.append("市場廣度閘門未確認")
        if heat is not None and heat >= 70:
            reasons.append(f"過熱 {heat:.1f}，不再屬試單區")
    elif old_bucket.startswith("🚀"):
        if not live_duck:
            reasons.append("正式鴨嘴結構暫時不成立")
        if right is not None and right < 65:
            reasons.append(f"右側 {right:.1f}<65")
    elif old_bucket.startswith("🟡"):
        if heat is not None and heat < 70:
            reasons.append(f"過熱降至 {heat:.1f}<70")
    elif old_bucket.startswith("🟠"):
        threshold=90 if rs is not None and rs >= 85 else 80
        if heat is not None and heat < threshold:
            reasons.append(f"過熱降至 {heat:.1f}<{threshold}")
    elif old_bucket.startswith("🔴"):
        reasons.append("贏家退潮條件目前解除")

    if reasons:
        return "；".join(reasons[:3])
    if operation and operation not in {"—",""}:
        return f"目前改判：{operation}"
    return "目前不再符合原分類"


def _start_signal_episode(history: list[dict], now, row, bucket: str, price):
    code=str(row.get("代號"))
    name=_clean_cell(row.get("名稱"))
    rec={
        "日期":now.strftime("%Y-%m-%d"),
        "代號":code,
        "名稱":name,
        "觸發分類":bucket,
        "進入時間":now.strftime("%H:%M:%S"),
        "觸發價":price,
        "離開時間":"",
        "離開價":None,
        "離開原因":"",
        "目前分類":bucket,
        "目前狀態":"✅ 目前有效",
        "最新價":price,
        "觸發後漲跌%":0.0,
    }
    history.append(rec)
    return len(history)-1


def _close_signal_episode(history: list[dict], idx, now, row, reason: str, current_bucket: str=""):
    if idx is None or not (0 <= int(idx) < len(history)):
        return
    rec=history[int(idx)]
    price=_num(row.get("盤中價")) if row is not None else None
    rec["離開時間"]=now.strftime("%H:%M:%S")
    rec["離開價"]=price
    rec["離開原因"]=reason
    rec["目前分類"]=current_bucket or "⏳ 等待"
    rec["目前狀態"]="➡️ 已轉分類" if current_bucket else "⏹️ 已離開原分類"
    rec["最新價"]=price
    p0=_num(rec.get("觸發價"))
    rec["觸發後漲跌%"]=(price/p0-1)*100 if p0 not in (None,0) and price is not None else None


def _decorate_signal_tracking(radar: pd.DataFrame) -> pd.DataFrame:
    """Track current-valid signals and retain today's triggered/left history with reasons."""
    if radar is None or radar.empty:
        return radar

    now=dt.datetime.now(TAIPEI_TZ)
    day=now.strftime("%Y-%m-%d")
    if st.session_state.get("v357_signal_day") != day:
        st.session_state["v357_signal_day"]=day
        st.session_state["v357_signal_tracker"]={}
        st.session_state["v357_signal_change_log"]=[]
        st.session_state["v357_signal_history"]=[]

    tracker=st.session_state.setdefault("v357_signal_tracker",{})
    log=st.session_state.setdefault("v357_signal_change_log",[])
    history=st.session_state.setdefault("v357_signal_history",[])

    out=radar.copy()
    out["追蹤分類"]=out.apply(_signal_bucket,axis=1)
    row_map={str(r.get("代號")):r for _,r in out.iterrows()}

    actionable={}
    for code,r in row_map.items():
        bucket=str(r.get("追蹤分類") or "")
        if bucket and not _signal_data_issue(r):
            actionable[code]=(bucket,r)

    for code in list(tracker):
        old=tracker.get(code) or {}
        old_bucket=str(old.get("分類") or "")
        old_idx=old.get("history_index")
        cur=actionable.get(code)
        row=row_map.get(code)

        if cur is None:
            reason=_signal_leave_reason(old_bucket,row) if row is not None else "目前行情列暫時不存在"
            _close_signal_episode(history,old_idx,now,row,reason,"")
            log.append({
                "時間":now.strftime("%H:%M:%S"),"代號":code,
                "名稱":_clean_cell(row.get("名稱")) if row is not None else (history[int(old_idx)].get("名稱","") if old_idx is not None and int(old_idx)<len(history) else ""),
                "變化":f"離開 {old_bucket}","原因":reason,
                "價格":_num(row.get("盤中價")) if row is not None else None,
            })
            tracker.pop(code,None)
            continue

        new_bucket,row=cur
        if new_bucket != old_bucket:
            reason=f"分類轉換：{old_bucket} → {new_bucket}"
            _close_signal_episode(history,old_idx,now,row,reason,new_bucket)
            log.append({
                "時間":now.strftime("%H:%M:%S"),"代號":code,"名稱":row.get("名稱",""),
                "變化":f"{old_bucket} → {new_bucket}","原因":reason,"價格":_num(row.get("盤中價")),
            })
            tracker.pop(code,None)
            price=_num(row.get("盤中價"))
            new_idx=_start_signal_episode(history,now,row,new_bucket,price)
            tracker[code]={
                "分類":new_bucket,"觸發價":price,"觸發時間":now.isoformat(),
                "history_index":new_idx,"名稱":_clean_cell(row.get("名稱")),
            }

    for code,(bucket,r) in actionable.items():
        price=_num(r.get("盤中價"))
        old=tracker.get(code)
        if old is None:
            idx=_start_signal_episode(history,now,r,bucket,price)
            tracker[code]={
                "分類":bucket,"觸發價":price,"觸發時間":now.isoformat(),
                "history_index":idx,"名稱":_clean_cell(r.get("名稱")),
            }
            log.append({
                "時間":now.strftime("%H:%M:%S"),"代號":code,"名稱":r.get("名稱",""),
                "變化":f"進入 {bucket}","原因":"即時條件成立","價格":price,
            })
        else:
            idx=old.get("history_index")
            if idx is not None and 0 <= int(idx) < len(history):
                rec=history[int(idx)]
                rec["最新價"]=price
                rec["目前分類"]=bucket
                rec["目前狀態"]="✅ 目前有效"
                p0=_num(rec.get("觸發價"))
                rec["觸發後漲跌%"]=(price/p0-1)*100 if p0 not in (None,0) and price is not None else None

    for rec in history:
        code=str(rec.get("代號"))
        r=row_map.get(code)
        if r is None:
            continue
        price=_num(r.get("盤中價"))
        rec["最新價"]=price
        p0=_num(rec.get("觸發價"))
        rec["觸發後漲跌%"]=(price/p0-1)*100 if p0 not in (None,0) and price is not None else None
        if rec.get("離開時間"):
            issue=_signal_data_issue(r)
            current=_signal_bucket(r)
            if issue:
                rec["目前分類"]="⚠️ 行情待刷新"
            else:
                rec["目前分類"]=current or "⏳ 等待"

    del log[:-300]
    del history[:-1200]

    trigger_price=[]; trigger_time=[]; change_pct=[]; duration=[]
    for _,r in out.iterrows():
        code=str(r.get("代號")); t=tracker.get(code)
        if not t:
            trigger_price.append(None); trigger_time.append(""); change_pct.append(None); duration.append("")
            continue
        p0=_num(t.get("觸發價")); p=_num(r.get("盤中價"))
        trigger_price.append(p0)
        try:
            t0=dt.datetime.fromisoformat(t.get("觸發時間"))
            sec=max(0,int((now-t0).total_seconds()))
            t0_text=t0.strftime("%H:%M:%S")
        except Exception:
            sec=0
            t0_text=""
        trigger_time.append(t0_text)
        change_pct.append((p/p0-1)*100 if p0 not in (None,0) and p is not None else None)
        duration.append(f"{sec//60}分{sec%60:02d}秒" if sec>=60 else f"{sec}秒")
    out["觸發價"]=trigger_price
    out["觸發時間"]=trigger_time
    out["觸發後漲跌%"]=change_pct
    out["訊號持續"]=duration
    return out


def _today_signal_history_frames() -> tuple[pd.DataFrame,pd.DataFrame]:
    hist=st.session_state.get("v357_signal_history",[])
    if not hist:
        return pd.DataFrame(),pd.DataFrame()
    h=pd.DataFrame(hist).copy()
    if h.empty:
        return pd.DataFrame(),h
    h["仍在原分類"]=h["離開時間"].fillna("").astype(str).eq("")
    summary=[]
    for bucket,g in h.groupby("觸發分類",dropna=False):
        summary.append({
            "分類":bucket,
            "今日曾觸發檔數":int(g["代號"].astype(str).nunique()),
            "訊號段數":int(len(g)),
            "目前仍在原分類":int(g.loc[g["仍在原分類"],"代號"].astype(str).nunique()),
            "已離開／轉分類":int((~g["仍在原分類"]).sum()),
        })
    s=pd.DataFrame(summary)
    if not s.empty:
        order={"🌱 回測試單":1,"📈 主要進場":2,"🚀 趨勢確認持有":3,"🟡 獲利保護":4,"🟠 積極鎖利":5,"🔴 贏家退潮":6}
        s["_o"]=s["分類"].map(order).fillna(99)
        s=s.sort_values(["_o","分類"]).drop(columns="_o")
    return s,h

def _manual_query_watch_codes(snap: dict) -> set[str]:
    """If the user is looking at a manual intraday stock, keep it in the reactive feed too."""
    q=str(st.session_state.get("v32_manual_last_query") or "").strip()
    stocks=snap.get("stocks") if snap else None
    if not q or stocks is None or stocks.empty or "代號" not in stocks.columns:
        return set()
    z=stocks.copy()
    z["代號"]=z["代號"].astype(str).map(_norm_stock_code)
    exact=z[z["代號"].eq(_norm_stock_code(q))]
    if not exact.empty:
        return {str(exact.iloc[0]["代號"])}
    names=z.get("名稱",pd.Series("",index=z.index)).astype(str)
    hit=z[names.str.contains(q,case=False,na=False,regex=False)]
    return {str(hit.iloc[0]["代號"])} if not hit.empty else set()


def _all_reactive_watch_codes(radar: pd.DataFrame, bg_state: dict, snap: dict) -> list[str]:
    """Track ALL stocks currently participating in any盤中/即時 view.

    This intentionally has no 120-name ranking cut.  It combines the six trading
    buckets, market-event radar names and the current manual lookup.
    """
    codes=set()
    if radar is not None and not radar.empty:
        z=radar.copy()
        if "追蹤分類" not in z.columns:
            z["追蹤分類"]=z.apply(_signal_bucket,axis=1)
        active=z[(z["追蹤分類"].astype(str)!="") & (~z["價格為估計"].fillna(True).astype(bool))]
        codes |= set(active["代號"].astype(str).map(_norm_stock_code))
    events=bg_state.get("events") or {}
    for key in ["new_strong","out_strong","new_duck","near"]:
        codes |= {_norm_stock_code(c) for c in (events.get(key) or set()) if _norm_stock_code(c)}
    codes |= _manual_query_watch_codes(snap)
    for rec in (st.session_state.get("v357_signal_history",[]) or []):
        c=_norm_stock_code(rec.get("代號"))
        if c:
            codes.add(c)
    return sorted(codes)


def _reactive_interval_for_count(n: int) -> int:
    """Keep all active names live while automatically protecting MIS/network load."""
    n=max(0,int(n or 0))
    if n<=250: return 5
    if n<=500: return 8
    if n<=800: return 12
    if n<=1200: return 15
    return 20


def _trade_radar_view(df: pd.DataFrame, limit=80, sort_kind="entry") -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    z=df.copy()
    if sort_kind=="exit":
        sort_cols=[c for c in ["過熱程度","盤中RS","右側分數"] if c in z.columns]
    else:
        sort_cols=[c for c in ["右側分數","盤中RS","鴨嘴完成度%"] if c in z.columns]
    if sort_cols:
        z=z.sort_values(sort_cols,ascending=[False]*len(sort_cols),na_position="last")
    show=[c for c in [
        "代號","名稱","市場","盤中價","漲跌幅%","觸發價","觸發後漲跌%","訊號持續",
        "盤中RS","鴨嘴完成度%","右側分數","過熱程度","盤中操作","進場階段","獲利階段",
        "市場閘門","成交狀態","資料狀態","資料層級","資料更新時間","資料年齡秒","MIS報價時間","價格來源"
    ] if c in z.columns]
    return z[show].head(limit)


def _render_change_radar(snap: dict, bg_state: dict, manager=None):
    # Every table marked盤中/即時 reads the same reactive overlay.
    snap = _snapshot_with_fast_rows(snap,bg_state)
    stocks = snap.get("stocks")
    if stocks is None or stocks.empty:
        return

    # ===== 第一層：最新回測後真正可操作的進出場雷達 =====
    st.write("#### 🎯 盤中回測交易雷達")
    st.caption(
        "這一層與下方『盤中單檔查詢』使用同一套規則："
        "預備80＋RS70＋市場改善＝試單；右側≥65＋市場改善＝主要進場；"
        "一般/可交易股過熱70～80、RS85領頭股90＝獲利管理。"
        "全市場約30秒刷新廣度與候選發現；六大交易分類、事件雷達、手動查詢，以及今日曾觸發後已離開的股票全部進入動態5–20秒批次追蹤。主清單只採 MIS z 真實成交價。真正離開條件的股票不會留在『目前有效』清單，但會保留在『今日曾觸發』紀錄並標示離開原因。"
    )

    try:
        radar=_get_intraday_trade_radar(snap,bg_state)
        radar=_decorate_live_freshness(radar,bg_state,snap)
        radar=_decorate_signal_tracking(radar)
    except Exception as e:
        radar=pd.DataFrame()
        st.warning(f"盤中交易雷達暫時無法重算：{e}；下方市場事件雷達仍可使用。")

    if not radar.empty:
        real=radar[~radar["價格為估計"].fillna(True).astype(bool)].copy()
        est_count=int(radar["價格為估計"].fillna(False).astype(bool).sum())
        stale=real[real.get("資料狀態",pd.Series("",index=real.index)).astype(str).str.startswith("⚠️")].copy()
        real=real[~real.index.isin(stale.index)].copy()
        entry_s=real["進場階段"].fillna("").astype(str)
        profit_s=real["獲利階段"].fillna("").astype(str)
        op_s=real["盤中操作"].fillna("").astype(str)

        trial=real[entry_s.str.startswith("🌱")]
        major=real[entry_s.str.startswith("📈")]
        hold=real[op_s.str.startswith("🚀")]
        protect=real[profit_s.str.startswith("🟡") & ~profit_s.str.contains("結構轉弱",na=False)]
        lock=real[profit_s.str.startswith("🟠")]
        retreat=real[profit_s.str.startswith("🔴")]

        # Keep the discovery universe stable until the next full-market pulse, but do not
        # drop names because of an arbitrary 120-stock cap.  All six trade buckets +
        # event radar + current manual lookup are reactive.
        full_cache=st.session_state.get("v356_trade_radar_full_cache",{})
        watch_source=full_cache.get("data") if isinstance(full_cache,dict) else None
        if not isinstance(watch_source,pd.DataFrame) or watch_source.empty:
            watch_source=radar.copy()
        else:
            watch_source=watch_source.copy()
        watch_source["追蹤分類"]=watch_source.apply(_signal_bucket,axis=1)
        fast_codes=_all_reactive_watch_codes(watch_source,bg_state,snap)
        reactive_interval=_reactive_interval_for_count(len(fast_codes))
        if manager is not None and dashboard_mode=="intraday":
            try:
                manager.set_fast_watch_codes(
                    fast_codes,interval_seconds=reactive_interval,enabled=True,max_codes=2500
                )
            except Exception as e:
                st.caption(f"盤中動態追蹤設定暫時失敗：{e}")

        c1,c2,c3,c4,c5,c6=st.columns(6)
        c1.metric("🌱 回測試單候選",f"{len(trial)} 檔")
        c2.metric("📈 主要進場候選",f"{len(major)} 檔")
        c3.metric("🚀 趨勢確認持有",f"{len(hold)} 檔")
        c4.metric("🟡 獲利保護",f"{len(protect)} 檔")
        c5.metric("🟠 積極鎖利",f"{len(lock)} 檔")
        c6.metric("🔴 贏家退潮",f"{len(retreat)} 檔")

        fast_count=int(bg_state.get("fast_watch_count",0) or 0)
        fast_returned=int(bg_state.get("fast_returned_count",0) or 0)
        fast_cov=_num(bg_state.get("fast_coverage"),0) or 0
        fast_last=bg_state.get("fast_last_completed") or "等待第一輪"
        fast_running=bool(bg_state.get("fast_running"))
        fast_error=bg_state.get("fast_last_error")
        fast_interval=int(bg_state.get("fast_interval_seconds",reactive_interval) or reactive_interval)
        engine=_clean_cell(bg_state.get("engine_version")) or "—"
        st.caption(f"⚡ 盤中動態追蹤：{fast_count} 檔｜本輪回傳 {fast_returned} 檔（{fast_cov:.1f}%）｜約{fast_interval}秒/批｜最近完成：{fast_last}｜{'更新中' if fast_running else '待命'}｜引擎 {engine}。全市場廣度/RS/Wade約30秒更新。")
        if fast_error:
            st.warning(f"快速追蹤上一輪失敗：{fast_error}；目前保留上一筆可用資料並顯示資料年齡。")
        if not stale.empty:
            st.warning(f"有 {len(stale)} 檔真實成交列已超過允許的新鮮度，暫時不計入『即時可操作』數量；展開下方逾時清單可檢查。")
            with st.expander(f"⚠️ 行情逾時／待刷新：{len(stale)} 檔", expanded=False):
                st.dataframe(_trade_radar_view(stale,120,"entry"),hide_index=True,use_container_width=True,height=320)

        gate=_intraday_entry_gate(snap)
        if gate["允許"]:
            st.success(f"目前盤中進場閘門：{gate['標籤']}｜{gate['說明']}")
        else:
            st.warning(f"目前盤中進場閘門：{gate['標籤']}｜{gate['說明']}。即使個股結構到位，也先列觀察，不直接升級新買點。")
        if est_count:
            st.caption(f"另有 {est_count} 檔目前只有委託簿/前收參考價，已排除在上述可操作數量之外。")

        with st.expander(f"🌱 回測試單候選：{len(trial)} 檔｜預備80＋RS70＋市場改善", expanded=False):
            if trial.empty:
                st.caption("目前沒有以真實成交價觸發的試單候選。")
            else:
                st.dataframe(_trade_radar_view(trial,max(80,len(trial)),"entry"),hide_index=True,use_container_width=True,height=420)

        with st.expander(f"📈 主要進場候選：{len(major)} 檔｜右側≥65＋市場改善", expanded=False):
            if major.empty:
                st.caption("目前沒有以真實成交價觸發的主要進場候選。")
            else:
                st.dataframe(_trade_radar_view(major,max(100,len(major)),"entry"),hide_index=True,use_container_width=True,height=460)

        with st.expander(f"🚀 趨勢確認持有：{len(hold)} 檔", expanded=False):
            if hold.empty:
                st.caption("目前沒有盤中右側確認持有名單。")
            else:
                st.dataframe(_trade_radar_view(hold,max(100,len(hold)),"entry"),hide_index=True,use_container_width=True,height=440)

        with st.expander(f"🟡 獲利保護：{len(protect)} 檔｜停止追價／準備鎖利", expanded=False):
            if protect.empty:
                st.caption("目前沒有進入第一層獲利保護區的股票。")
            else:
                st.dataframe(_trade_radar_view(protect,max(100,len(protect)),"exit"),hide_index=True,use_container_width=True,height=440)

        with st.expander(f"🟠 積極鎖利：{len(lock)} 檔｜一般股80／RS85領頭股90附近", expanded=False):
            if lock.empty:
                st.caption("目前沒有進入積極鎖利區的股票。")
            else:
                st.dataframe(_trade_radar_view(lock,max(100,len(lock)),"exit"),hide_index=True,use_container_width=True,height=440)

        with st.expander(f"🔴 贏家退潮：{len(retreat)} 檔｜曾過熱後結構退潮", expanded=False):
            if retreat.empty:
                st.caption("目前沒有觸發贏家退潮管理的股票。")
            else:
                st.dataframe(_trade_radar_view(retreat,max(100,len(retreat)),"exit"),hide_index=True,use_container_width=True,height=440)

        hist_summary,hist_df=_today_signal_history_frames()
        if not hist_df.empty:
            closed=hist_df[hist_df["離開時間"].fillna("").astype(str)!=""].copy()
            closed_unique=int(closed["代號"].astype(str).nunique()) if not closed.empty else 0
            with st.expander(
                f"🕘 今日曾觸發／已離開原分類：{closed_unique} 檔｜今日共 {len(hist_df)} 段訊號",
                expanded=False
            ):
                st.caption(
                    "『目前有效』六區只放此刻仍符合條件的股票；這裡保留今天曾經觸發過的完整軌跡。"
                    "離開原因會區分條件失效、分類轉換、MIS 無真實成交價或行情逾時；離開原分類本身不等於正式賣出指令。"
                )
                if not hist_summary.empty:
                    st.dataframe(hist_summary,hide_index=True,use_container_width=True,height=min(280,80+35*len(hist_summary)))
                show_hist=[c for c in [
                    "代號","名稱","觸發分類","進入時間","觸發價","離開時間","離開價",
                    "離開原因","目前分類","目前狀態","最新價","觸發後漲跌%"
                ] if c in hist_df.columns]
                hz=hist_df.copy()
                hz["_sort"]=pd.to_datetime(hz["進入時間"],format="%H:%M:%S",errors="coerce")
                hz=hz.sort_values("_sort",ascending=False).drop(columns="_sort")
                st.dataframe(hz[show_hist],hide_index=True,use_container_width=True,height=460)

        fast_log=st.session_state.get("v357_signal_change_log",[])
        if fast_log:
            with st.expander(f"⚡ 盤中動態訊號變化（最近 {min(len(fast_log),300)} 筆）",expanded=False):
                st.dataframe(pd.DataFrame(fast_log[::-1]),hide_index=True,use_container_width=True,height=340)

        st.caption(
            "注意：全市場雷達不知道你的個別持股成本，因此不會把『收盤≤成本-8%』列成全市場硬停損。"
            "-8% 最後防線仍在單檔查詢／個股決策卡填入成本後判斷。"
        )
    else:
        st.info("盤中交易雷達尚無可用結果；等待背景第一批完成後會自動出現。")

    # ===== 第二層：保留原本的市場事件掃描，但不再把掉出強勢條件叫做「退出」 =====
    st.divider()
    events = bg_state.get("events") or {}
    new_strong = set(events.get("new_strong") or set())
    out_strong = set(events.get("out_strong") or set())
    new_duck = set(events.get("new_duck") or set())
    near_codes = set(events.get("near") or set())
    compare_label = str(events.get("compare_label") or f"相對 {latest_date} 正式")

    st.write("##### 📡 市場事件觀察（原強勢／鴨嘴雷達）")
    st.caption(
        f"{compare_label}｜這一層是市場結構事件，不等同回測後的買賣指令。"
        "原本『暫時退出』改名為『結構轉弱』，避免與真正出場條件混淆。"
    )
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🔥 本輪新進強勢", f"{len(new_strong)} 檔")
    c2.metric("🌱 即將符合強勢", f"{len(near_codes)} 檔")
    c3.metric("⚠️ 本輪結構轉弱", f"{len(out_strong)} 檔")
    c4.metric("🦆 鴨嘴結構新進", f"{len(new_duck)} 檔")

    if new_strong:
        st.success(f"🔥 最新背景掃描發現 {len(new_strong)} 檔新進強勢；這是事件觀察，不代表一定是最佳買點。")
    with st.expander(f"🔥 本輪新進強勢：{len(new_strong)} 檔", expanded=False):
        if new_strong:
            st.dataframe(_event_table(stocks, new_strong, 40), hide_index=True, use_container_width=True, height=360)
        else:
            st.caption("本輪沒有新進強勢。")
    with st.expander(f"🌱 即將符合強勢：{len(near_codes)} 檔", expanded=False):
        if near_codes:
            st.dataframe(_event_table(stocks, near_codes, 80), hide_index=True, use_container_width=True, height=420)
        else:
            st.caption("目前沒有只差一項且已接近門檻的股票。")
    with st.expander(f"⚠️ 本輪結構轉弱：{len(out_strong)} 檔", expanded=False):
        if out_strong:
            st.dataframe(_event_table(stocks, out_strong, 50), hide_index=True, use_container_width=True, height=360)
        else:
            st.caption("本輪沒有強勢結構轉弱。")
    with st.expander(f"🦆 本輪新進鴨嘴價格結構：{len(new_duck)} 檔", expanded=False):
        if new_duck:
            st.dataframe(_event_table(stocks, new_duck, 50), hide_index=True, use_container_width=True, height=360)
        else:
            st.caption("本輪沒有新進鴨嘴價格結構。")

    log = bg_state.get("event_log") or []
    if log:
        with st.expander(f"🕘 今日盤中事件紀錄（最近 {min(len(log), 300)} 筆）", expanded=False):
            log_df=pd.DataFrame(log[::-1])
            if "類型" in log_df.columns:
                log_df["類型"]=log_df["類型"].replace({"⚠️ 暫時退出":"⚠️ 結構轉弱"})
            st.dataframe(log_df, hide_index=True, use_container_width=True, height=360)

def _render_trial_summary(bg_state: dict, manager=None):
    snap = bg_state.get("snapshot")
    if snap:
        snap = _snapshot_with_fast_rows(snap,bg_state)
    running = bool(bg_state.get("running"))
    last_error = bg_state.get("last_error")
    if not snap:
        if running:
            st.info("🔄 背景正在建立第一批全市場行情；你可以繼續使用下面功能，不需要等在這裡。")
        elif last_error:
            st.error(f"背景盤中行情暫時無法建立：{last_error}")
        else:
            st.info("背景掃描已啟動，第一批完成後會顯示盤中試算。")
        return

    trial_exposure, trial_exposure_label = _trial_exposure_reference(snap)
    mode_title = "📡 今日盤中試算" if dashboard_mode == "intraday" else "🧾 今日收盤試算（等待正式更新）"
    status_text = "🔄 背景正在處理下一批；目前畫面維持上一批完成資料" if running else "🟢 背景待命／等待下一批"
    mode_note = (
        "全市場約30秒刷新廣度/RS/Wade與候選發現；所有盤中交易分類、事件與手動查詢個股再以動態5–20秒批次追蹤。前台每5秒只讀完成結果。"
        if dashboard_mode == "intraday"
        else "13:30 後不再連續背景掃描；保留最後完成結果，等待 18:05 正式資料。"
    )

    st.write(f"### {mode_title}")
    st.caption(
        f"{mode_note}｜{status_text}｜比較基準：{latest_date} 正式收盤｜"
        f"背景完成：{bg_state.get('last_completed') or snap.get('snapshot_time','—')}｜"
        f"可用報價覆蓋：{_fmt_trial(snap.get('coverage'),1,'%')}｜MIS z真實成交覆蓋：{_fmt_trial(snap.get('real_trade_coverage'),1,'%')}"
    )
    if last_error:
        st.warning(f"上一輪背景更新失敗：{last_error}；目前仍顯示上一批成功資料。")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("RS 強勢股", f"{int(snap.get('strong_count',0)):,} 檔", _metric_delta(snap.get("strong_count"), strong_count, 0, " 檔"))
    c2.metric("Wade 試算", _fmt_trial(snap.get("wade"), 1), _metric_delta(snap.get("wade"), wade_score, 1))
    c3.metric("上漲比例", _fmt_trial(snap.get("advance_ratio"), 1, "%"), _metric_delta(snap.get("advance_ratio"), advance_ratio, 1, " pct"))
    c4.metric("市場階段", str(snap.get("stage") or "—"), f"{latest_date}：{stage}")
    c5.metric("市場水位", str(snap.get("water") or "—"), f"{latest_date}：{water}")
    c6.metric("總曝險參考", trial_exposure, trial_exposure_label)

    risk_now = str(snap.get("risk") or "未觸發")
    early_now = str(snap.get("early") or "未觸發")
    action_now = str(snap.get("action") or "觀望等待")
    if risk_now.startswith(("🟠", "🟡")):
        st.warning(f"**今日試算風險：{risk_now}**｜操作：{action_now}")
    elif early_now.startswith("🟢"):
        st.success(f"**今日試算轉強：{early_now}**｜操作：{action_now}")
    else:
        st.info(f"**今日試算：{snap.get('wade_state','—')}**｜風險：{risk_now}｜操作：{action_now}")

    decision_cards([
        ("今天市場在哪裡", f"{snap.get('stage','—')}｜{snap.get('water','—')}", f"昨日正式：{stage}｜{water}"),
        ("現在怎麼做", action_now, f"昨日正式：{operation_level}"),
        ("總曝險參考", trial_exposure, f"昨日正式：{exposure_plan['顯示']}"),
        ("早期轉強", early_now, f"昨日正式：{early_signal_phase}"),
        ("風險／減碼", risk_now, f"昨日正式：{left_signal_phase}"),
        ("盤中強弱", f"Wade {_fmt_trial(snap.get('wade'),1)}", f"上漲 {_fmt_trial(snap.get('advance_ratio'),1,'%')}｜新高/新低 {snap.get('new_high',0)}/{snap.get('new_low',0)}"),
    ])
    _render_change_radar(snap, bg_state, manager)



def _intraday_entry_gate(snap: dict) -> dict:
    """盤中版本的市場進場閘門，盡量貼近進場回測所用的廣度條件。"""
    adv=_num(snap.get("advance_ratio"),0) or 0
    ma20_pct=_num(snap.get("above_ma20_ratio"))
    trend=str(snap.get("trend") or "")
    stage_now=str(snap.get("stage") or "")
    risk_now=str(snap.get("risk") or "")
    improving=trend in {"增加","明顯增加"} or stage_now in {"低檔回升","中段回升"}
    breadth_ok=adv>=50 and (ma20_pct is None or ma20_pct>=50) and improving and not risk_now.startswith("🟠")
    ma20_text="MA20廣度不足" if ma20_pct is None else f"MA20上方 {ma20_pct:.1f}%"
    label="✅ 盤中市場廣度改善" if breadth_ok else "⏳ 盤中市場廣度未確認"
    detail=f"上漲 {adv:.1f}%｜{ma20_text}｜強勢股{trend or '—'}｜{risk_now or '未觸發'}"
    return {"允許":bool(breadth_ok),"標籤":label,"說明":detail}


def _intraday_stock_trade_plan(live_row, formal_row, snap: dict, cost=None) -> dict:
    """把單檔即時行情套入 v3.5 的進場/獲利/失敗風控，產生盤中試算操作。"""
    price=_num(live_row.get("盤中價"))
    rs_live=_num(live_row.get("盤中RS"))
    live_ma20=_num(live_row.get("盤中MA20"))
    live_ma60=_num(live_row.get("盤中MA60"))
    estimated=_boolish(live_row.get("價格為估計"))

    if formal_row is None:
        formal_row=pd.Series(dtype=object)

    prev_ma20=_num(formal_row.get("MA20"))
    prev_ma60=_num(formal_row.get("MA60"))
    ma20chg=(live_ma20-prev_ma20) if live_ma20 is not None and prev_ma20 is not None else _num(formal_row.get("MA20較前日"),0) or 0
    ma60chg=(live_ma60-prev_ma60) if live_ma60 is not None and prev_ma60 is not None else _num(formal_row.get("MA60較前日"),0) or 0
    if live_ma20 is not None and live_ma60 is not None and prev_ma20 is not None and prev_ma60 is not None:
        spreadchg=(live_ma20-live_ma60)-(prev_ma20-prev_ma60)
    else:
        spreadchg=_num(formal_row.get("開口較前日"),0) or 0

    conds=[
        price is not None and live_ma20 is not None and live_ma60 is not None and price>live_ma20 and price>live_ma60,
        live_ma20 is not None and live_ma60 is not None and live_ma20>live_ma60,
        ma20chg>0,
        ma60chg>0,
        spreadchg>0,
    ]
    complete=float(sum(bool(x) for x in conds)*20)

    prev_stage=_clean_cell(formal_row.get("鴨嘴階段"))
    prev_formal=_boolish(formal_row.get("正式鴨嘴")) or prev_stage.startswith("持續符合") or prev_stage.startswith("新進")
    live_formal=all(conds)
    today_new=bool(live_formal and not prev_formal)
    exit_today=bool(prev_formal and not live_formal)
    pre_s=_clean_cell(formal_row.get("預備鴨嘴狀態"))

    gap20=((price/live_ma20-1.0)*100.0) if price is not None and live_ma20 not in (None,0) else 0.0
    heat=max(0,min(70,max(0,gap20)*5.0))
    if rs_live is not None and rs_live>=95: heat+=12
    elif rs_live is not None and rs_live>=85: heat+=8
    if live_formal or complete>=100: heat+=8
    first_hot=formal_row.get("首次過熱日期")
    try:
        has_hot=not pd.isna(first_hot) and str(first_hot).strip() not in {"","—","-"}
    except Exception:
        has_hot=bool(first_hot)
    code_live=_norm_stock_code(live_row.get("代號"))
    hot_codes=st.session_state.setdefault("v356_intraday_hot_codes",set())
    if heat>=70 and code_live:
        hot_codes.add(code_live)
    has_hot=bool(has_hot or (code_live and code_live in hot_codes))
    if has_hot: heat+=8
    heat=max(0,min(100,heat))

    right=max(0,min(55,complete*0.55))
    if rs_live is not None:
        if rs_live>=85: right+=25
        elif rs_live>=70: right+=18
        elif rs_live>=50: right+=10
    if today_new: right+=12
    elif live_formal: right+=8
    elif "A級" in pre_s: right+=6
    if ma20chg>0 and ma60chg>0 and spreadchg>0: right+=6
    if exit_today: right-=30
    right=max(0,min(100,right))

    gate=_intraday_entry_gate(snap)
    entry=stock_entry_plan(
        complete,rs_live,right,heat,live_formal,today_new,pre_s,
        market_entry_ok=gate["允許"],market_entry_label=gate["標籤"]
    )
    profit=stock_profit_exit_plan(
        heat,rs_live,right,live_formal,today_new,exit_today,
        has_hot_date=has_hot,ma20chg=ma20chg,spreadchg=spreadchg
    )

    # -8% 回測規則是「收盤確認」。盤中只顯示是否已進入該區，不把盤中價直接當正式停損成交訊號。
    failure=stock_failure_risk_plan(
        price,cost=cost if cost and cost>0 else None,exit_today=exit_today,
        right=right,ma20=live_ma20,ma20chg=ma20chg
    )
    fail_stage=_clean_cell(failure.get("失敗風控")) or "—"
    profit_stage=_clean_cell(profit.get("獲利階段")) or "—"
    entry_stage=_clean_cell(entry.get("進場階段")) or "—"

    if fail_stage.startswith("🔴"):
        operation="🔴 出場區（收盤確認）"
        action="目前盤中已落入成本-8%區；停止加碼，若收盤仍符合則依回測規則退出。"
    elif profit_stage.startswith("🔴"):
        operation="🔴 贏家退潮／退出"
        action=profit.get("獲利動作") or "提高減碼幅度"
    elif profit_stage.startswith("🟠"):
        operation="🟠 分批減碼"
        action=(profit.get("獲利動作") or "分批鎖利")+"；盤中先列警示，收盤確認。"
    elif profit_stage.startswith("🟡") and ("鎖利" in profit_stage or "高熱" in profit_stage or heat>=70):
        operation="🟡 停止追價／準備鎖利"
        action=(profit.get("獲利動作") or "停止追價")+"；盤中先列警示，收盤確認。"
    elif exit_today:
        operation="🟡 結構轉弱"
        action="停止新增部位並提高警覺；單一盤中結構轉弱不直接等同全數賣出。"
    elif entry_stage.startswith("🌱"):
        operation="🌱 小部位試單"
        action=entry.get("進場動作") or "小部位試單"
    elif entry_stage.startswith("📈"):
        operation="📈 主要進場"
        action=entry.get("進場動作") or "建立／補到基本部位"
    elif live_formal and right>=65:
        operation="🚀 偏多持有"
        action="右側結構盤中仍成立；已有部位續抱，未持有者避免追滿。"
    elif entry_stage.startswith("⚠️"):
        operation="⚠️ 暫不進場"
        action=entry.get("進場動作") or "等待降溫"
    else:
        operation="⏳ 等待"
        action=entry.get("進場動作") or "等待更完整觸發"

    if estimated:
        action=f"{action}｜目前價格是委託簿/前收參考，不是最新成交，判讀可信度降一級。"

    return {
        "盤中操作":operation,
        "盤中動作":action,
        "進場階段":entry_stage,
        "進場動作":entry.get("進場動作") or "—",
        "獲利階段":profit_stage,
        "獲利動作":profit.get("獲利動作") or "—",
        "失敗風控":fail_stage,
        "風控動作":failure.get("風控動作") or "—",
        "右側分數":right,
        "過熱程度":heat,
        "鴨嘴完成度%":complete,
        "盤中正式鴨嘴":live_formal,
        "盤中今日新進":today_new,
        "盤中結構退出":exit_today,
        "市場閘門":gate["標籤"],
        "市場閘門說明":gate["說明"],
        "硬停損價":failure.get("硬停損價"),
        "成本報酬%":failure.get("成本報酬%"),
    }


def _manual_stock_lookup(manager):
    st.write("### 🔎 個股最新狀況（手動查詢）")
    st.caption(
        "查詢會先用背景快照定位股票，再自動只向 MIS 更新這一檔；不會重抓 2,000 檔。"
        "除了價格、RS與鴨嘴外，會直接顯示『盤中要進場、持有、減碼還是等待』。所有盤中操作都是試算，正式仍以收盤確認。"
    )
    with st.form("v32_manual_lookup_form", clear_on_submit=False):
        q = st.text_input("輸入股票代號或名稱", key="v32_manual_lookup_input", placeholder="例如 2330、台積電")
        submitted = st.form_submit_button("🔎 查詢＋更新這一檔", type="primary", use_container_width=True)
    if submitted and q.strip():
        st.session_state["v32_manual_last_query"] = q.strip()
        st.session_state.pop("v32_single_override", None)

    query = st.session_state.get("v32_manual_last_query")
    bg_state = manager.get_state()
    snap = bg_state.get("snapshot")
    if snap:
        snap = _snapshot_with_fast_rows(snap,bg_state)
    if not query:
        return
    if not snap or snap.get("stocks") is None:
        st.info("背景第一批行情尚未完成；完成後再按一次查詢即可。")
        return

    stocks = snap["stocks"].copy()
    stocks["代號"] = stocks["代號"].map(_norm_stock_code)
    qq = str(query).strip()
    hit = stocks[
        stocks["代號"].astype(str).str.contains(qq, case=False, na=False, regex=False)
        | stocks.get("名稱", pd.Series("", index=stocks.index)).astype(str).str.contains(qq, case=False, na=False, regex=False)
    ].copy()
    if hit.empty:
        st.warning(f"找不到「{qq}」的盤中資料。")
        return
    sort_cols = [c for c in ["盤中強勢", "即將強勢", "盤中RS"] if c in hit.columns]
    if sort_cols:
        hit = hit.sort_values(sort_cols, ascending=[False] * len(sort_cols), na_position="last")
    r = hit.iloc[0].copy()

    # v3.5.2：搜尋即單檔刷新；全市場仍保留背景快照架構。
    if submitted and q.strip():
        try:
            from intraday_live import refresh_single_stock
            with st.spinner(f"更新 {r.get('代號')} 最新行情…"):
                fresh = refresh_single_stock(snap, str(r.get("代號")))
            st.session_state["v32_single_override"] = fresh
        except Exception as e:
            st.warning(f"單檔即時更新失敗，暫時顯示背景快照：{e}")

    override = st.session_state.get("v32_single_override")
    use_override=False
    if override and str(override.get("代號")) == str(r.get("代號")):
        try:
            ot=pd.Timestamp(override.get("單檔抓取時間"))
            rt_raw=_clean_cell(r.get("快速追蹤時間")) or _clean_cell(bg_state.get("last_completed")) or _clean_cell(snap.get("snapshot_time"))
            rt=pd.Timestamp(rt_raw) if rt_raw else pd.Timestamp.min
            use_override=ot>=rt
        except Exception:
            use_override=True
    if use_override:
        for k, v in override.items():
            r[k] = v
        source_time = override.get("單檔抓取時間", "—")
        source_label = "單檔即時抓取"
    else:
        if override and str(override.get("代號")) == str(r.get("代號")):
            st.session_state.pop("v32_single_override",None)
        source_time = _clean_cell(r.get("快速追蹤時間")) or bg_state.get("last_completed") or snap.get("snapshot_time", "—")
        source_label = "快速追蹤" if _clean_cell(r.get("快速追蹤時間")) else "背景快照"

    code_s=_norm_stock_code(r.get("代號"))
    formal_hit=stock_lr[stock_lr["代號"].astype(str).map(_norm_stock_code)==code_s] if not stock_lr.empty and "代號" in stock_lr.columns else pd.DataFrame()
    formal_row=formal_hit.iloc[0] if not formal_hit.empty else pd.Series(dtype=object)

    mis_time = _clean_cell(r.get("MIS報價時間")) or "—"
    price_source = _clean_cell(r.get("價格來源")) or source_label
    bid_s = _fmt_trial(r.get("最佳買價"), 2)
    ask_s = _fmt_trial(r.get("最佳賣價"), 2)
    estimated = _boolish(r.get("價格為估計"))
    st.caption(
        f"查詢：{qq}｜資料來源：{source_label}｜程式時間：{source_time}｜MIS報價：{mis_time}｜"
        f"價格來源：{price_source}｜最佳買/賣：{bid_s} / {ask_s}｜背景全市場版本：#{bg_state.get('version',0)}"
    )
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("股票", f"{r.get('代號','—')} {r.get('名稱','')}")
    c2.metric("盤中參考價" if estimated else "最新成交價", _fmt_trial(r.get("盤中價"), 2), _fmt_trial(r.get("漲跌幅%"), 2, "%"))
    c3.metric("盤中 RS", _fmt_trial(r.get("盤中RS"), 1), "強勢" if bool(r.get("盤中強勢", False)) else "未達強勢")
    c4.metric("強勢條件", f"{int(_num(r.get('強勢條件通過數'),0) or 0)}/6", str(r.get("強勢尚缺條件") or "—"))
    c5.metric("鴨嘴價格結構", "✅ 符合" if bool(r.get("盤中鴨嘴價格結構", False)) else "—", f"MA20乖離 {_fmt_trial(r.get('月線乖離率%'),1,'%')}")

    if estimated:
        st.warning(
            f"目前沒有 MIS z 最新成交，顯示的是「{price_source}」；不會再用買賣中間價，所以不會出現 245.25 這種不可成交價位。"
        )

    cost_v=st.number_input(
        "我的平均成本價（選填；盤中會提示是否已進入 -8% 收盤硬停損區）",
        min_value=0.0,value=0.0,step=0.1,format="%.2f",key=f"intraday_cost_{code_s}"
    )

    plan=_intraday_stock_trade_plan(r,formal_row,snap,cost=cost_v if cost_v>0 else None)
    st.write("#### 🎯 盤中操作判讀")
    decision_cards([
        ("現在要做什麼",plan["盤中操作"],_short_text(plan["盤中動作"],110)),
        ("進場判讀",plan["進場階段"],_short_text(plan["進場動作"],95)),
        ("獲利／出場",plan["獲利階段"],_short_text(plan["獲利動作"],95)),
        ("失敗風控",plan["失敗風控"],_short_text(plan["風控動作"],95)),
    ])
    st.info(f"**盤中下一步：** {plan['盤中動作']}")
    st.caption(
        f"盤中試算：右側 {plan['右側分數']:.0f}／過熱 {plan['過熱程度']:.0f}／"
        f"鴨嘴完成 {plan['鴨嘴完成度%']:.0f}%｜{plan['市場閘門']}｜{plan['市場閘門說明']}"
    )
    if cost_v>0 and plan.get("硬停損價") is not None:
        ret_txt="—" if plan.get("成本報酬%") is None else f"{plan['成本報酬%']:+.2f}%"
        st.caption(f"平均成本 {cost_v:.2f}｜-8% 收盤硬停損參考價 {plan['硬停損價']:.2f}｜目前相對成本 {ret_txt}。盤中跌破只預警，收盤仍符合才正式觸發。")

    if bool(r.get("即將強勢", False)):
        st.warning(f"🌱 **即將符合強勢**｜目前尚缺：{r.get('強勢尚缺條件','—')}")
    elif bool(r.get("盤中強勢", False)):
        st.success("🔥 **目前盤中已符合強勢條件**（收盤後才正式確認）。")
    else:
        st.info(f"目前尚未符合盤中強勢條件｜尚缺：{r.get('強勢尚缺條件','—')}")

    if st.button("⚡ 只更新這一檔（不重掃全市場）", key="v32_force_single", use_container_width=True):
        try:
            from intraday_live import refresh_single_stock
            with st.spinner(f"只抓 {r.get('代號')} 最新行情…"):
                fresh = refresh_single_stock(snap, str(r.get("代號")))
            st.session_state["v32_single_override"] = fresh
            st.rerun()
        except Exception as e:
            st.error(f"單檔即時更新失敗：{e}")

    show = [c for c in ["代號", "名稱", "市場", "盤中價", "漲跌幅%", "最佳買價", "最佳賣價", "價格來源", "MIS報價時間", "盤中RS", "盤中MA20", "盤中MA50", "盤中MA60", "盤中MA200", "月線乖離率%", "盤中強勢", "即將強勢", "強勢條件通過數", "強勢尚缺條件", "盤中鴨嘴價格結構"] if c in hit.columns]
    st.dataframe(hit[show].head(20), hide_index=True, use_container_width=True, height=min(500, 70 + 35 * min(len(hit), 12)))


st.title("📈 台股分析中心")
st.markdown(
    '<div class="hero-sub">RS 市場廣度 × 鴨嘴型態 × 培育中心｜正式網頁版 v3.5.6｜進場回測 × 獲利管理 × -8%失敗風控 × 個股決策中心</div>',
    unsafe_allow_html=True
)

u_status = update_status.get("status", "unknown")
last_run = update_status.get("last_run_taipei", "—")
message = update_status.get("message", "")

if dashboard_mode == "intraday":
    st.markdown(
        f'<div class="status-strip status-ok"><b>目前模式：</b>📡 今日盤中試算　｜　'
        f'<b>正式比較基準：</b>{latest_date}　｜　<b>全市場同步：</b>約30秒　｜　'
        f'<b>盤中個股同步：</b>動態5–20秒　｜　<b>最近正式排程：</b>{html.escape(str(last_run))}</div>',
        unsafe_allow_html=True
    )
elif dashboard_mode == "close_trial":
    st.markdown(
        f'<div class="status-strip status-warn"><b>目前模式：</b>🧾 今日收盤試算（等待正式更新）　｜　'
        f'<b>正式比較基準：</b>{latest_date}　｜　<b>盤中背景連續掃描：</b>已停止　｜　'
        f'<b>18:05：</b>正式更新後自動切換</div>',
        unsafe_allow_html=True
    )
else:
    cls = "status-ok" if u_status in {"success", "ready", "bootstrap"} else "status-warn"
    mode_text = "✅ 今日正式收盤" if dashboard_mode == "formal" else "🕘 最新正式收盤"
    st.markdown(
        f'<div class="status-strip {cls}"><b>目前模式：</b>{mode_text}　｜　'
        f'<b>資料日：</b>{latest_date}　｜　<b>自動更新：</b>週一至週五 18:05　｜　'
        f'<b>最近排程：</b>{html.escape(str(last_run))}</div>',
        unsafe_allow_html=True
    )

if u_status in {"partial", "no_new_data"} and message:
    st.warning(message)
elif u_status == "error" and message:
    st.error(message)

if dashboard_mode in {"intraday", "close_trial"}:
    from intraday_live import get_background_manager
    bg_manager = get_background_manager(INTRADAY_ENGINE_GENERATION)
    official_strong_codes = _codes_from_df(strong)
    official_duck_codes = _codes_from_df(all_ok)
    duck_watch_codes = official_duck_codes | _codes_from_df(pre)

    if dashboard_mode == "intraday":
        ctrl1, ctrl2 = st.columns([3, 1])
        auto_scan = ctrl1.toggle(
            "背景全市場同步（約 30 秒）",
            value=True,
            key="v32_background_scan",
            help="背景執行行情抓取與約2,000檔運算；不會讓前台等待。關閉後仍可用手動單檔查詢。"
        )
        bg_manager.configure(
            daily, strong,
            official_strong_codes=official_strong_codes,
            official_duck_codes=official_duck_codes,
            duck_watch_codes=duck_watch_codes,
            formal_date=latest_date,
            interval_seconds=30,
            enabled=auto_scan,
            ensure_once=True,
        )
        if ctrl2.button("🔄 背景立即重掃", key="v32_bg_force", use_container_width=True):
            bg_manager.request_refresh()
            st.toast("已要求背景重掃；畫面不會卡住。")

        if hasattr(st, "fragment"):
            @st.fragment(run_every="5s")
            def _memory_view_fragment():
                _render_trial_summary(bg_manager.get_state(), bg_manager)
                _manual_stock_lookup(bg_manager)
            _memory_view_fragment()
        else:
            st.info("目前 Streamlit 版本不支援局部刷新；背景仍會掃描，但畫面需手動重新整理才會讀到新快照。")
            _render_trial_summary(bg_manager.get_state(), bg_manager)
            _manual_stock_lookup(bg_manager)
    else:
        # 收盤後只確保有一批最後行情，不再連續掃描。
        bg_manager.configure(
            daily, strong,
            official_strong_codes=official_strong_codes,
            official_duck_codes=official_duck_codes,
            duck_watch_codes=duck_watch_codes,
            formal_date=latest_date,
            interval_seconds=30,
            enabled=False,
            ensure_once=True,
        )
        state_now = bg_manager.get_state()
        _render_trial_summary(state_now, bg_manager)
        if not state_now.get("snapshot") or state_now.get("running"):
            st.caption("背景正在建立收盤後最後一批資料；可稍後按下方『讀取背景最新結果』，不用等它跑完。")
        if st.button("📥 讀取背景最新結果", key="v32_read_close", use_container_width=True):
            st.rerun()

    if dashboard_mode != "intraday":
        _manual_stock_lookup(bg_manager)
    st.caption(
        "⬇️ 再往下為『最新正式收盤／歷史詳細資料』。背景掃描與行情計算不在前台執行，"
        "全市場約30秒同步、盤中活躍個股動態5–20秒同步都在背景執行，不會因更新卡住你正在看的表格或搜尋。"
    )
else:
    st.write("### 今日決策摘要")
    decision_cards([
        ("市場在哪裡", f"{market_signal(stage)}｜{stage}", f"{water}；{stage_plain(stage)}"),
        ("總曝險參考", exposure_plan["顯示"], f"{exposure_plan['層級']}｜{exposure_plan['說明']}"),
        ("現在怎麼做", operation_level, market_lr["左右側建議"]),
        ("早期轉強", early_signal_phase, signal_current_text(early_signal_stats)),
        ("風險／減碼", left_signal_phase, signal_current_text(left_reduce_stats if risk_signal_level(left_alert)=="reduce" else left_signal_stats)),
        ("下一個升級／降級條件", short_trigger, next_trigger),
    ])
    st.markdown(
        f'<div class="focus-box"><b>今天只看三件事：</b><br>'
        f'① 部位：總曝險參考 <b>{html.escape(exposure_plan["顯示"])}</b>（{html.escape(exposure_plan["層級"])}）<br>'
        f'② 訊號：早期轉強 {html.escape(early_signal_phase)}；風險/減碼 {html.escape(left_signal_phase)}<br>'
        f'③ 操作：{html.escape(operation_level)}；{html.escape(next_trigger)}<br>'
        f'<span class="small-note">曝險百分比是相對於你自行設定的股票最大風險預算，不是總資產配置。</span></div>',
        unsafe_allow_html=True
    )
    st.markdown(stage_flow_html(stage), unsafe_allow_html=True)

    st.write("### 市場關鍵數據")
    cards([
        ("歷史水位", water, f"全期 P{safe_num(full_rank,1)}／近3年 P{safe_num(rolling_rank,1)}"),
        ("RS 強勢股", safe_num(strong_count,0," 檔"), f"占比 {safe_num(strong_pct,2,'%')}／單日 {safe_num(change,0,' 檔')}"),
        ("Wade市場內部", safe_num(wade_score,1," 分"), f"今日 {signed_num(wade_day_change)}／5日 {signed_num(wade_change5)}"),
        ("大盤左右側", market_lr["左右側階段"], f"左 {market_lr['左側分數']:.0f}／右 {market_lr['右側分數']:.0f}"),
        ("總曝險參考", exposure_plan["顯示"], exposure_plan["層級"]),
        ("資料完整度", f"{quality['資料完整度標籤']}／{quality['資料完整度']:.0f}", "今天應有資料是否齊全"),
        ("訊號一致度", f"{quality['訊號一致度標籤']}／{quality['訊號一致度']:.0f}", "RS、Wade、廣度、多時間週期是否同向"),
        ("鴨嘴符合", f"{len(all_ok):,} 檔", f"新進 {len(duck_new):,}／退出 {len(duck_exit):,}／A級 {pre_a:,}"),
        ("復甦候選", f"{len(recovery):,} 檔", "觀察 → 左側試單 → 接近右側確認"),
    ])

tabs=st.tabs(["🏠 正式基準","🎯 個股決策中心","📊 RS／市場廣度","🧭 Wade大盤強弱","🌱 復甦候選","📐 台指期","🧪 資料品質","🔄 更新資料"])

with tabs[0]:
    st.subheader("正式收盤基準／歷史比較")
    decision_cards([
        ("市場狀態",f"{market_signal(stage)}｜{stage}",stage_plain(stage)),
        ("總曝險參考",exposure_plan["顯示"],f"{exposure_plan['層級']}｜{exposure_plan['說明']}"),
        ("操作",operation_level,market_lr["左右側建議"]),
        ("訊號天數",f"轉強：{early_signal_phase}",f"風險：{left_signal_phase}"),
        ("Wade節奏",wade_timing_text,f"目前 {safe_num(wade_score,1)}｜今日 {signed_num(wade_day_change)}｜5日 {signed_num(wade_change5)}"),
        ("下一個確認點",short_trigger,next_trigger),
    ])
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
    st.subheader("🎯 個股決策中心")
    st.caption("把左右側、鴨嘴、培育中心、估值與 v1.2 出場回測整合在同一頁。主表先看「現在怎麼做＋獲利/出場階段」，單檔卡再看下一動作；詳細文字收進展開區。")

    # 市場環境只保留真正會影響個股決策的摘要，避免一進頁面先被大段文字淹沒。
    cards([
        ("整體市場",market_lr["左右側階段"],f"左 {market_lr['左側分數']:.0f}／右 {market_lr['右側分數']:.0f}"),
        ("加權指數",twii_lr["階段"],f"左 {twii_lr['左側分數']:.0f}／右 {twii_lr['右側分數']:.0f}"),
        ("櫃買指數",twoii_lr["階段"],f"左 {twoii_lr['左側分數']:.0f}／右 {twoii_lr['右側分數']:.0f}"),
        ("總曝險參考",exposure_plan["顯示"],exposure_plan["層級"]),
        ("市場操作",operation_level,market_lr["左右側建議"]),
        ("個股進場閘門",entry_market_gate["標籤"],entry_market_gate["說明"]),
    ])
    with st.expander("查看大盤／指數詳細判讀", expanded=False):
        ic1,ic2=st.columns(2)
        with ic1: st.info(f"**加權：{twii_lr['階段']}**\n\n{twii_lr['建議']}\n\n依據：{twii_lr['依據']}")
        with ic2: st.info(f"**櫃買：{twoii_lr['階段']}**\n\n{twoii_lr['建議']}\n\n依據：{twoii_lr['依據']}")
        st.markdown(f'<div class="action-box"><b>整體市場：</b>{html.escape(market_lr["左右側階段"])}<br><b>建議：</b>{html.escape(market_lr["左右側建議"])}<br><b>依據：</b>{html.escape(market_lr["判讀依據"])}<br><span class="small-note">左側高分不等於最低點已出現；右側高分也不代表可以無條件追高。過熱／退出訊號優先於加碼訊號。</span></div>',unsafe_allow_html=True)

    st.write("### 回測版完整交易流程")
    decision_cards([
        ("① 試單","🌱 預備80＋RS70",f"{entry_market_gate['標籤']}｜只用小部位"),
        ("② 主要進場","📈 右側≥65","市場改善時建立／補到基本部位"),
        ("③ 趨勢持有","🚀 正式鴨嘴／右側續強","正式新進不是追滿訊號；RS85也不是必要買點"),
        ("④ 獲利保護","🟡70～80／領頭90","一般／可交易股先鎖利；RS85領頭股可放寬到90"),
        ("⑤ 失敗退出","🔴 收盤≤成本-8%","MA20／右側轉弱只預警；-8%是最後硬停損"),
    ])
    st.caption(f"市場進場閘門：{entry_market_gate['標籤']}｜{entry_market_gate['說明']}。進場v1.0、出場v1.2、失敗風控v1.0分工使用，避免單一訊號同時負責買賣。")

    st.write("### 今日操作分布")
    if not stock_lr.empty and "操作標籤" in stock_lr.columns:
        ac=stock_lr["操作標籤"].value_counts()
        cards([
            ("🌱 預備試單",f"{int(ac.get('🌱 預備試單',0)):,} 檔","預備80＋RS70＋市場改善"),
            ("📈 主要進場",f"{int(ac.get('📈 主要進場',0)):,} 檔","右側≥65＋市場改善"),
            ("🚀 偏多持有",f"{int(ac.get('🚀 偏多持有',0)):,} 檔","正式／右側結構續強"),
            ("⚠️ 停止追價",f"{int(ac.get('⚠️ 停止追價',0)):,} 檔","進入獲利保護區"),
            ("🟠 分批減碼",f"{int(ac.get('🟠 分批減碼',0)):,} 檔","過熱獲利管理"),
            ("🟡 結構警示",f"{int(ac.get('🟡 結構警示',0)):,} 檔","只預警；硬停損改看成本-8%"),
        ])
        if "獲利階段" in stock_lr.columns:
            ps=stock_lr["獲利階段"].fillna("").astype(str)
            cards([
                ("🟢 趨勢持有",f"{int(ps.str.startswith('🟢').sum()):,} 檔","尚未到主要獲利保護區"),
                ("🟡 獲利保護",f"{int(ps.str.startswith('🟡').sum()):,} 檔","停止追價／開始鎖利"),
                ("🟠 積極鎖利",f"{int(ps.str.startswith('🟠').sum()):,} 檔","分批減碼／提高移動停利"),
                ("🔴 贏家退潮",f"{int(ps.str.startswith('🔴').sum()):,} 檔","曾過熱後轉弱；提高減碼"),
            ])

    if stock_lr.empty:
        st.info("全市場整合資料尚未產生；跑一次正式資料更新後即可使用。")
    else:
        st.write("### ① 一眼比較主表")
        f1,f2,f3=st.columns([1.3,1,1])
        q=f1.text_input("搜尋代號／名稱",key="dc_stock_search",placeholder="例如 2330、台積電")
        actions=["全部"]+sorted(stock_lr["操作標籤"].dropna().astype(str).unique().tolist()) if "操作標籤" in stock_lr.columns else ["全部"]
        action_chosen=f2.selectbox("操作",actions,key="dc_action")
        phase_options=["全部"]+sorted(stock_lr["左右側階段"].dropna().astype(str).unique().tolist()) if "左右側階段" in stock_lr.columns else ["全部"]
        phase_chosen=f3.selectbox("左右側階段",phase_options,key="dc_phase")

        sx=filter_stock_table(stock_lr.copy(),q)
        if action_chosen!="全部": sx=sx[sx["操作標籤"]==action_chosen]
        if phase_chosen!="全部": sx=sx[sx["左右側階段"]==phase_chosen]

        a1,a2,a3=st.columns(3)
        duck_filter=a1.selectbox("鴨嘴",["全部","正式／新進","預備 A級","預備 B級","尚未形成"],key="dc_duck")
        cultivate_filter=a2.selectbox("培育中心",["全部","三率＋營收","三率三升","營收三增","尚無"],key="dc_cultivate")
        max_heat=a3.slider("最高過熱程度",0,100,100,5,key="dc_heat")

        if "過熱程度" in sx.columns:
            sx=sx[pd.to_numeric(sx["過熱程度"],errors="coerce").fillna(0)<=max_heat]
        # pandas 在空 DataFrame 上 apply(axis=1) 可能回傳 DataFrame，
        # 此時再使用 .str 會發生 AttributeError。篩選已無結果時直接略過後續字串篩選。
        if duck_filter!="全部" and not sx.empty:
            ds=sx.apply(stock_duck_summary,axis=1).astype("string")
            if duck_filter=="正式／新進": sx=sx[ds.str.contains("正式|新進",regex=True,na=False)]
            elif duck_filter=="預備 A級": sx=sx[ds.str.contains("A級",regex=False,na=False)]
            elif duck_filter=="預備 B級": sx=sx[ds.str.contains("B級",regex=False,na=False)]
            else: sx=sx[ds.eq("—")]
        if cultivate_filter!="全部" and not sx.empty:
            cs=sx.apply(stock_cultivation_summary,axis=1).astype("string")
            if cultivate_filter=="三率＋營收": sx=sx[cs.str.contains("三率＋營收",regex=False,na=False)]
            elif cultivate_filter=="三率三升": sx=sx[cs.str.contains("三率三升|三率＋營收",regex=True,na=False)]
            elif cultivate_filter=="營收三增": sx=sx[cs.str.contains("營收三增|三率＋營收",regex=True,na=False)]
            else: sx=sx[cs.eq("—")]

        sort_cols=[c for c in ["操作排序","決策排序分數","右側分數","左側分數"] if c in sx.columns]
        if sort_cols:
            asc=[True,False,False,False][:len(sort_cols)]
            sx=sx.sort_values(sort_cols,ascending=asc,na_position="last")
        compact=stock_decision_compact(sx)
        st.caption(f"符合目前篩選：{len(sx):,} 檔。進場：預備80＋RS70＋市場改善先試單、右側65為主要進場；獲利：一般/可交易股70～80、RS85領頭股90；失敗：單檔填成本後以收盤-8%作硬停損。")
        st.dataframe(compact,hide_index=True,use_container_width=True,height=520)

        st.write("### ② 單檔決策卡")
        if sx.empty:
            st.info("目前篩選沒有個股；放寬條件後即可查看單檔決策卡。")
        else:
            choice_map={stock_choice_label(r):i for i,r in sx.iterrows()}
            labels=list(choice_map.keys())
            selected_label=st.selectbox("選一檔看完整判讀",labels,key="dc_single_stock")
            r=sx.loc[choice_map[selected_label]]
            left_v=_num(r.get("左側分數"),0) or 0
            right_v=_num(r.get("右側分數"),0) or 0
            heat_v=_num(r.get("過熱程度"),0) or 0
            duck_s=stock_duck_summary(r)
            cult_s=stock_cultivation_summary(r)
            val_s=stock_valuation_summary(r)
            rs_v=_num(r.get("RS判讀"))
            rec_v=_num(r.get("復甦分數"))
            eps_v=_num(r.get("EPS年增率%"))
            fair_v=_num(r.get("合理價"))

            code_s=_clean_cell(r.get('代號'))
            st.markdown(f"#### {code_s} {_clean_cell(r.get('名稱'))}")
            cost_v=st.number_input(
                "我的平均成本價（選填；用來判斷回測版 -8% 硬停損）",
                min_value=0.0,value=0.0,step=0.1,format="%.2f",key=f"dc_cost_{code_s}"
            )
            failure_plan=stock_failure_risk_plan(
                r.get("收盤"),cost=cost_v if cost_v>0 else None,
                exit_today=_boolish(r.get("今日退出")) or text(r.get("鴨嘴階段"),"")=="今日退出",
                right=right_v,ma20=r.get("MA20"),ma20chg=r.get("MA20較前日")
            )
            if cost_v>0 and failure_plan.get("硬停損價") is not None:
                st.caption(f"你的 -8% 收盤硬停損參考價：約 {failure_plan['硬停損價']:.2f}；目前相對成本 {failure_plan['成本報酬%']:+.2f}%。盤中碰到不等於正式觸發，以收盤確認。")
            else:
                st.caption("尚未填成本價：全市場表只會顯示結構預警，不會把 MA20／右側轉弱誤判成硬性退出。")
            profit_stage=_clean_cell(r.get("獲利階段")) or "⚪ 尚未進入獲利管理"
            profit_action=_clean_cell(r.get("獲利動作")) or "—"
            profit_threshold=_clean_cell(r.get("獲利門檻")) or "—"
            profit_profile=_clean_cell(r.get("獲利管理類型")) or "—"
            entry_stage=_clean_cell(r.get("進場階段")) or "⏳ 等待"
            entry_action=_clean_cell(r.get("進場動作")) or "—"
            failure_stage=_clean_cell(failure_plan.get("失敗風控")) or "—"
            failure_action=_clean_cell(failure_plan.get("風控動作")) or "—"
            if failure_stage.startswith("🔴"):
                next_step=failure_action
            elif profit_stage.startswith(("🟡","🟠","🔴")):
                next_step=profit_action
            else:
                next_step=entry_action

            decision_cards([
                ("現在怎麼做",_clean_cell(r.get("操作標籤")) or "⏳ 等待",f"{_clean_cell(r.get('左右側階段')) or '—'}｜左 {left_v:.0f}／右 {right_v:.0f}／過熱 {heat_v:.0f}"),
                ("進場階段",entry_stage,_short_text(entry_action,88)),
                ("獲利／出場階段",profit_stage,f"{profit_profile}｜主要門檻 {profit_threshold}"),
                ("失敗風控",failure_stage,_short_text(failure_action,88)),
            ])
            st.info(f"**下一個動作：** {_short_text(next_step,150)}")
            decision_cards([
                ("鴨嘴",duck_s,f"完成度 {safe_num(r.get('鴨嘴完成度%'),0,'%')}｜{_clean_cell(r.get('預備鴨嘴狀態')) or '無預備標記'}"),
                ("RS／復甦",f"RS {safe_num(rs_v,1)}",f"{_clean_cell(r.get('RS來源')) or '資料不足'}｜復甦分數 {safe_num(rec_v,1)}"),
                ("培育中心",cult_s,f"EPS年增 {safe_num(eps_v,1,'%')}｜{_clean_cell(r.get('培育中心類別')) or '—'}"),
                ("估值",val_s,f"合理價 {safe_num(fair_v,2)}｜{_clean_cell(r.get('估值狀態')) or '—'}"),
            ])
            st.info(f"**一句話判讀：** {stock_one_line(r)}")
            with st.expander("查看完整判讀／回測依據", expanded=False):
                st.write(f"**系統建議：** {_clean_cell(r.get('系統建議')) or '—'}")
                st.write(f"**左側依據：** {_clean_cell(r.get('左側依據')) or '—'}")
                st.write(f"**右側依據：** {_clean_cell(r.get('右側依據')) or '—'}")
                st.write(f"**加碼條件：** {_clean_cell(r.get('加碼條件')) or '—'}")
                st.write(f"**進場回測：** {_clean_cell(r.get('進場回測參考')) or '—'}")
                st.write(f"**獲利回測：** {_clean_cell(r.get('回測參考')) or '—'}")
                st.write(f"**失敗風控回測：** {_clean_cell(failure_plan.get('風控參考')) or '—'}")
                st.caption("v3.5 將三件事拆開：進場看市場改善＋結構/RS；贏家用過熱管理；失敗交易最後防線用收盤相對平均成本-8%。MA20／右側／鴨嘴退出只作結構預警，不再自動等同全數賣出。")

        st.write("### ③ 多股橫向比較")
        st.caption("可選 2～5 檔，把原本散在後面的大量分析欄位收斂成同一張表。")
        compare_source=sx if not sx.empty else stock_lr
        compare_map={stock_choice_label(r):i for i,r in compare_source.iterrows()}
        selected_multi=st.multiselect("選擇要比較的股票",list(compare_map.keys()),key="dc_multi_stock")
        if len(selected_multi)>5:
            st.warning("為了單頁可讀性，目前只顯示前 5 檔。")
            selected_multi=selected_multi[:5]
        if selected_multi:
            cm=compare_source.loc[[compare_map[x] for x in selected_multi]].copy()
            cmp=stock_decision_compact(cm)
            # 多股比較只再帶加碼／失效兩個文字欄；獲利/出場階段已在主表欄位中。
            cmp["加碼條件"]=cm["加碼條件"].map(lambda v:_short_text(v,42)).values if "加碼條件" in cm.columns else "—"
            cmp["失效條件"]=cm["失效條件"].map(lambda v:_short_text(v,42)).values if "失效條件" in cm.columns else "—"
            st.dataframe(cmp,hide_index=True,use_container_width=True,height=min(360,85+48*len(cmp)))

    st.divider()
    with st.expander("🦆 進階：鴨嘴原始清單／今日新進退出",expanded=False):
        mode=st.radio("查看",["全部符合","今日新進","今日退出","即將可能符合"],horizontal=True,key="dc_duck_raw_mode")
        source_map={"全部符合":all_ok,"今日新進":duck_new,"今日退出":duck_exit,"即將可能符合":pre}
        dfx=source_map[mode].copy()
        dq=st.text_input("搜尋代號／名稱",key="dc_duck_raw_search")
        dfx=filter_stock_table(dfx,dq)
        if mode=="即將可能符合" and "預備鴨嘴狀態" in dfx.columns:
            grade=st.selectbox("預備等級",["全部","A級","B級"],key="dc_duck_grade")
            if grade!="全部": dfx=dfx[dfx["預備鴨嘴狀態"].fillna("").astype(str).str.contains(grade)]
        st.caption(f"顯示 {len(dfx):,} 檔")
        st.dataframe(clean_duck(dfx),hide_index=True,use_container_width=True,height=500)

    with st.expander("🔥 進階：培育中心交集／估值候選",expanded=False):
        bmode=st.radio("交集類型",["兩者皆有","三率三升","營收三增","估值買進候選"],horizontal=True,key="dc_breed_mode")
        bmap={"兩者皆有":both,"三率三升":three_rate,"營收三增":three_rev,"估值買進候選":value_candidates}
        bx=filter_stock_table(bmap[bmode].copy(),st.text_input("搜尋代號／名稱",key="dc_breed_search"))
        if bmode=="估值買進候選":
            cols=[c for c in ["日期","代號","名稱","市場","收盤","合理價","折價率%","估值狀態","鴨嘴狀態","月線乖離率%","培育中心類別","是否鴨嘴"] if c in bx.columns]
            if "折價率%" in bx.columns: bx=bx.sort_values("折價率%",ascending=False,na_position="last")
            st.dataframe(bx[cols] if cols else bx,hide_index=True,use_container_width=True,height=500)
        else:
            st.dataframe(clean_duck(bx),hide_index=True,use_container_width=True,height=500)

with tabs[2]:
    st.subheader("RS 強勢股市場廣度")
    st.markdown(stage_flow_html(stage),unsafe_allow_html=True)
    st.info(f"目前 RS 階段：{stage}｜{stage_plain(stage)}｜建議：{operation_action}")
    cards([("有效樣本",safe_num(first_value(rs_latest,["有效樣本數"]),0),""),("RS 強勢母池",safe_num(first_value(rs_latest,["RS強勢母池檔數"]),0),""),("強勢股",safe_num(strong_count,0),""),("占比",safe_num(strong_pct,2,"%"),""),("較前日",safe_num(change,0," 檔"),"")])
    days=st.slider("歷史圖顯示交易日",30,min(500,len(daily)),min(180,len(daily)),10)
    ch=breadth_chart(daily,days)
    if ch is not None: st.altair_chart(ch,use_container_width=True)
    st.write("**最新 RS 強勢股名單**")
    rs_q=st.text_input("搜尋代號／名稱",key="rs_search")
    rs_show=filter_stock_table(strong.copy(),rs_q)
    st.dataframe(rs_show,hide_index=True,use_container_width=True,height=540)
    with st.expander("歷史極值基準"): st.dataframe(extreme,hide_index=True,use_container_width=True)

with tabs[3]:
    st.subheader("Wade 式大盤內部強弱")
    st.caption("v0.2：把 Wade 公開常看的市場內部概念量化；不是 Wade 官方公式。已納入量價效率、加權/等權同步、櫃買相對強弱與強勢股主流延續性；指數資料抓不到時會自動退回市場內部代理。")
    cards([
        ("市場內部強度",safe_num(wade_score,1," / 100"),wade_state),
        ("風險／向上減碼",left_signal_phase,f"原始訊號：{left_alert_show}｜{signal_current_text(left_reduce_stats if risk_signal_level(left_alert)=="reduce" else left_signal_stats)}"),
        ("早期轉強",early_signal_phase,f"原始訊號：{early_signal_show}｜{signal_current_text(early_signal_stats)}"),
        ("上漲／下跌",f"{safe_num(advance_count,0)}／{safe_num(decline_count,0)}",f"上漲比例 {safe_num(advance_ratio,1,'%')}"),
        ("漲停／跌停*",f"{safe_num(limit_up,0)}／{safe_num(limit_down,0)}","*目前以單日報酬 ±9.5% 近似"),
        ("52週新高／新低",f"{safe_num(new_high,0)}／{safe_num(new_low,0)}","收盤價 250 交易日新高低"),
        ("上市上漲比例",safe_num(twse_adv,1,"%"),"TWSE 普通股"),
        ("上櫃上漲比例",safe_num(tpex_adv,1,"%"),"TPEx 普通股"),
        ("成交金額／20日",safe_num(amount_ratio20,1,"%"),f"量價效率 {safe_num(volume_eff,1)}"),
        ("加權／櫃買日報酬",f"{safe_num(twii_ret,2,'%')}／{safe_num(twoii_ret,2,'%')}",f"權值同步 {safe_num(index_sync,1)}"),
        ("上市／上櫃等權",f"{safe_num(twse_mean_return,2,'%')}／{safe_num(tpex_mean_return,2,'%')}",f"櫃買相對 {safe_num(tpex_rel,1)}"),
        ("主流延續率",safe_num(leader_retention,1,"%"),f"新領頭股 {safe_num(new_leaders,0)}／延續分數 {safe_num(leadership_score,1)}"),
    ])
    if risk_signal_level(left_alert)=="reduce":
        st.warning(f"目前：{left_signal_phase}｜{left_alert}。屬向上減碼觀察：停止追高、汰弱留強、分批降低風險，不等於一次清倉。")
    elif risk_signal_level(left_alert)=="watch":
        st.warning(f"目前：{left_signal_phase}｜{left_alert}。這一級先視為風險升高觀察，優先停止擴大曝險，尚未等同正式減碼訊號。")
    elif signal_active(early_signal):
        st.success(f"目前：{early_signal_phase}｜{early_signal}。第1天視為初現，第2天視為確認，第3天以上視為持續轉強。")
    else:
        st.info(f"目前狀態：{wade_state}。尚未觸發明確向上減碼或早期轉強訊號。")

    st.write("**訊號持續天數／歷史常態**")
    cards([
        ("風險／向上減碼",left_signal_phase,signal_stat_text(left_reduce_stats if risk_signal_level(left_alert)=="reduce" else left_signal_stats)),
        ("早期轉強",early_signal_phase,signal_stat_text(early_signal_stats)),
    ])
    with st.expander("訊號有效性回測（每段只算第一次觸發）"):
        bt=[]
        for sig_col,sig_name,mode in [("左側減碼警示","向上減碼","down"),("早期轉強訊號","早期轉強","up")]:
            for price_col,market_name in [("加權指數收盤","加權"),("櫃買指數收盤","櫃買")]:
                b=signal_event_backtest(daily,sig_col,price_col,mode)
                if not b.empty:
                    b.insert(0,"市場",market_name); b.insert(0,"訊號",sig_name); bt.append(b)
        if bt:
            bt_show=pd.concat(bt,ignore_index=True)
            st.dataframe(bt_show,hide_index=True,use_container_width=True)
            st.caption("早期轉強的『符合方向』＝期末上漲；左側風險欄位的回測仍包含黃色與橘色事件，操作上已分級顯示。樣本少時只供觀察，不應把比例當成固定勝率。")
        else:
            st.info("目前可用的指數歷史或訊號樣本不足；之後每天增量累積後會自動出現 5／10／20 日統計。")

    wc=wade_chart(daily,180)
    if wc is not None: st.altair_chart(wc,use_container_width=True)
    st.markdown("**目前 v0.2 分數組成**：上漲廣度 20%＋RS 水位 20%＋新高/低 10%＋漲跌停 7%＋RS方向 12%＋量價效率 10%＋加權/等權同步 8%＋櫃買相對強弱 6%＋主流延續 7%。")
    cols=[c for c in ["日期","Wade內部強度分數","Wade分數5日變化","Wade市場狀態","左側減碼警示","早期轉強訊號","上漲家數","下跌家數","上漲比例%","漲停近似家數","跌停近似家數","52週新高家數","52週新低家數","上市上漲比例%","上櫃上漲比例%","成交金額20日比%","加權指數報酬%","櫃買指數報酬%","上市等權平均報酬%","上櫃等權平均報酬%","主流延續率%","新強勢領頭股數","量價效率分數","權值同步分數","櫃買相對強弱分數","主流延續分數","市場總結白話","操作建議","市場階段","強勢股方向"] if c in daily.columns]
    if cols:
        hist_base=daily[cols].copy()

        # 歷史資料區間：預設最近 3 個月，可一路切到全部歷史。
        if "日期" in hist_base.columns:
            hist_base["日期"]=normalize_date_series(hist_base["日期"])
            hist_base=hist_base.dropna(subset=["日期"]).sort_values("日期")

        hist_range=st.selectbox(
            "歷史顯示區間",
            ["最近1個月","最近3個月","最近6個月","最近1年","全部"],
            index=1,
            key="wade_history_range",
        )

        days_map={
            "最近1個月":31,
            "最近3個月":93,
            "最近6個月":186,
            "最近1年":366,
        }

        if hist_range=="全部" or "日期" not in hist_base.columns or hist_base.empty:
            hist_show=hist_base.copy()
        else:
            latest_hist_date=hist_base["日期"].max()
            cutoff=latest_hist_date-pd.Timedelta(days=days_map[hist_range])
            hist_show=hist_base[hist_base["日期"]>=cutoff].copy()

        if "日期" in hist_show.columns:
            hist_show=hist_show.sort_values("日期",ascending=False)
            hist_show["日期"]=hist_show["日期"].dt.strftime("%Y-%m-%d")

        hist_show=hist_show.rename(columns={"左側減碼警示":"向上減碼觀察"})
        st.caption(f"目前顯示 {len(hist_show):,} 個交易日；可切換到『全部』一路往前查看既有歷史資料。")
        st.dataframe(hist_show,hide_index=True,use_container_width=True,height=600)


with tabs[4]:
    st.subheader("被遺忘資金／災後復甦候選")
    st.caption("v0.2：找跌深後開始修復的股票，再依左側／右側／過熱程度分成『👀觀察名單 → 🌱可左側試單 → ↗️接近右側確認』。復甦本身不是買點，但條件完整時可升級操作層級。")
    if recovery.empty:
        st.info("目前沒有符合條件的復甦候選；第一次新版更新後才會產生此工作表。")
    else:
        rx=recovery.copy()
        if "代號" in rx.columns:
            rx["代號"]=rx["代號"].astype(str).str.replace(r"\.0$","",regex=True)
        # 與鴨嘴/培育中心交叉標記，讓復甦候選能直接接現有系統。
        if not all_ok.empty and "代號" in all_ok.columns and "代號" in rx.columns:
            tag=all_ok.copy(); tag["代號"]=tag["代號"].astype(str).str.replace(r"\.0$","",regex=True)
            tcols=[c for c in ["代號","狀態","培育中心類別"] if c in tag.columns]
            tag=tag[tcols].drop_duplicates("代號")
            rx=rx.merge(tag,on="代號",how="left")
            if "狀態" in rx.columns: rx=rx.rename(columns={"狀態":"鴨嘴狀態"})
        # 接入左右側/過熱分數，直接把復甦候選分成三個可理解層級。
        if not stock_lr.empty and "代號" in stock_lr.columns and "代號" in rx.columns:
            lrcols=[c for c in ["代號","操作標籤","今日升降級","加碼條件","失效條件","左側分數","右側分數","過熱程度","過熱等級","左右側階段","系統建議","估值狀態","折價率%","合理價"] if c in stock_lr.columns]
            lrtag=stock_lr[lrcols].drop_duplicates("代號")
            rx=rx.merge(lrtag,on="代號",how="left",suffixes=("","_左右側"))
        if "左側分數" not in rx.columns: rx["左側分數"]=0
        if "右側分數" not in rx.columns: rx["右側分數"]=0
        if "過熱程度" not in rx.columns: rx["過熱程度"]=0
        rx["復甦階段"]=[classify_recovery_stage(l,r,h,rec) for l,r,h,rec in zip(
            rx["左側分數"],rx["右側分數"],rx["過熱程度"],rx.get("復甦分數",pd.Series(0,index=rx.index))
        )]
        order={"↗️ 接近右側確認":0,"🌱 可左側試單":1,"👀 觀察名單":2}
        cnt=rx["復甦階段"].value_counts()
        cards([
            ("👀 觀察名單",f"{int(cnt.get('👀 觀察名單',0)):,} 檔","剛開始修復，先追蹤"),
            ("🌱 可左側試單",f"{int(cnt.get('🌱 可左側試單',0)):,} 檔","修復＋左側條件較完整"),
            ("↗️ 接近右側確認",f"{int(cnt.get('↗️ 接近右側確認',0)):,} 檔","價格／RS／型態開始證明"),
        ])
        c1,c2=st.columns([1,1])
        rq=c1.text_input("搜尋代號／名稱",key="recovery_search")
        rstage=c2.selectbox("復甦階段",["全部","👀 觀察名單","🌱 可左側試單","↗️ 接近右側確認"],key="recovery_stage")
        rx=filter_stock_table(rx,rq)
        if rstage!="全部": rx=rx[rx["復甦階段"]==rstage]
        rx["_stage_order"]=rx["復甦階段"].map(order).fillna(9)
        sort_cols=["_stage_order"] + (["復甦分數"] if "復甦分數" in rx.columns else [])
        asc=[True] + ([False] if "復甦分數" in rx.columns else [])
        rx=rx.sort_values(sort_cols,ascending=asc,na_position="last").drop(columns=["_stage_order"])
        front=[c for c in ["代號","名稱","市場","收盤","復甦階段","操作標籤","今日升降級","加碼條件","失效條件","復甦分數","左側分數","右側分數","過熱程度","過熱等級","左右側階段","系統建議"] if c in rx.columns]
        rest=[c for c in rx.columns if c not in front]
        st.dataframe(rx[front+rest],hide_index=True,use_container_width=True,height=620)

with tabs[5]:
    st.subheader("台指期月／季線鴨嘴參考")
    if tx.empty: st.info("目前沒有台指期資料。")
    else:
        first=tx.iloc[0]; mouth=first.get("月季線鴨嘴"); mouth_text="符合" if str(mouth).lower() in {"true","1","是"} else "未符合"
        cards([("收盤",safe_num(first.get("收盤"),0),""),("MA20",safe_num(first.get("MA20"),2),""),("MA60",safe_num(first.get("MA60"),2),""),("月季線鴨嘴",mouth_text,"")])
        st.dataframe(tx,hide_index=True,use_container_width=True)

with tabs[6]:
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

with tabs[7]:
    st.subheader("更新資料")
    st.success("資料更新核心維持既有增量邏輯：股價只補新交易日；RS 舊歷史沿用；財報/營收已有當期資料就不重抓；PER 歷史只補新月份/缺漏月；Yahoo EPS 使用近期快取。排程仍為週一至週五台灣時間 18:05。")
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

st.divider(); st.caption(f"正式版 v3.5.7｜進場 v1.0 × 獲利出場 v1.2 × 失敗風控 v1.0 × 個股決策中心｜RS：{rs_date}｜鴨嘴：{duck_date}｜量化篩選與市場廣度工具，不構成投資建議。")
