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
    if text(left_alert,"—")!="—": return "分批減碼／汰弱留強"
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
    if text(left_alert,"—") != "—" or "減碼" in text(market_lr_phase,""):
        return "⚠️ 分批減碼／汰弱留強"
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

    if reduce!="—" or (stage=="高檔轉弱" and w5<=-5):
        phase="⚠️ 向上減碼觀察"
        action="停止追高；持有強股可續抱，但分批鎖利／汰弱留強"
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
    if reduce!="—": reasons.append("高檔退潮／向上減碼條件觸發")
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

def build_stock_lr_table(master, strong, recovery):
    """個股左右側量化代理。
    左側偏重估值、基本面與跌深修復；右側偏重價格趨勢、鴨嘴完成度與 RS。
    """
    if master is None or master.empty: return pd.DataFrame()
    x=master.copy()
    if "代號" not in x.columns: return pd.DataFrame()
    x["代號"]=x["代號"].astype(str).str.replace(r"\.0$","",regex=True)

    # 合併 RS。強勢名單與復甦名單只有符合者才有值；其餘不硬造 RS。
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
    if "RS_強勢" not in x.columns: x["RS_強勢"]=pd.NA
    if "RS_復甦" not in x.columns: x["RS_復甦"]=pd.NA
    x["RS判讀"]=pd.to_numeric(x["RS_強勢"],errors="coerce").combine_first(pd.to_numeric(x["RS_復甦"],errors="coerce"))

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
        formal=_boolish(r.get("正式鴨嘴")) or stage.startswith("持續符合") or stage.startswith("新進")
        ma20chg=_num(r.get("MA20較前日"),0) or 0
        ma60chg=_num(r.get("MA60較前日"),0) or 0
        spreadchg=_num(r.get("開口較前日"),0) or 0

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
        if overheat: right-=8; rreason.append("過熱，不宜追價")
        if exit_today: right-=30; rreason.append("今日退出")
        right=max(0,min(100,right))

        if exit_today:
            phase="🛑 右側失敗／退出觀察"
            action="停止新增部位；依原停損/退場規則降低部位，等下一次觸發"
        elif overheat and right>=60:
            phase="⚠️ 右側過熱／向上減碼"
            action="不追高；持有者可續抱強勢核心，但分批鎖利並提高移動停利"
        elif left>=65 and right>=60:
            phase="↗️ 左轉右確認"
            action="原左側試單可轉為基本部位；後續只有價格繼續證明才加碼"
        elif right>=75:
            phase="🚀 右側確認"
            action="趨勢已較明確；偏多持有，新增部位宜分批並設定結構失效點"
        elif left>=65:
            phase="🌱 左側布局候選"
            action="可列入低接/分批試單觀察，但需接受尚未確認底部，避免一次重押"
        elif right>=58:
            phase="📈 右側試單"
            action="價格開始證明，可小部位試單；獲利且結構續強再建立基本部位"
        elif left>=48 and right>=45:
            phase="🔄 左右側交界觀察"
            action="價值與價格都出現部分條件，但優勢不夠集中；等觸發更清楚"
        elif left>=48:
            phase="🔍 左側研究名單"
            action="具部分價值/修復條件，先研究與等待，不急著買"
        else:
            phase="⏳ 尚未形成優勢"
            action="目前不屬高品質左側或右側機會，等待條件改善"

        rows.append({
            "左側分數":round(left,1),"右側分數":round(right,1),"左右側階段":phase,"系統建議":action,
            "左側依據":"；".join(lreason[:5]) or "—","右側依據":"；".join(rreason[:5]) or "—"
        })
    z=pd.concat([x.reset_index(drop=True),pd.DataFrame(rows)],axis=1)
    z["決策排序分數"]=z[["左側分數","右側分數"]].max(axis=1)
    return z

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
stock_lr=build_stock_lr_table(master,strong,recovery)
pre_status=pre.get("預備鴨嘴狀態",pd.Series(dtype=str)).fillna("").astype(str) if not pre.empty else pd.Series(dtype=str)
pre_a=int(pre_status.str.contains("A級").sum()) if not pre_status.empty else 0
pre_b=int(pre_status.str.contains("B級").sum()) if not pre_status.empty else 0
update_status=load_update_status()

st.title("📈 台股分析中心")
st.markdown('<div class="hero-sub">RS 市場廣度 × 鴨嘴型態 × 培育中心｜正式網頁版 v2.7｜決策校準版</div>',unsafe_allow_html=True)

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

st.write("### 今日決策摘要")
decision_cards([
    ("市場在哪裡",f"{market_signal(stage)}｜{stage}",f"{water}；{stage_plain(stage)}"),
    ("現在怎麼做",operation_level,market_lr["左右側建議"]),
    ("今天變化",wade_timing_text,f"Wade 今日 {signed_num(wade_day_change)}｜5日 {signed_num(wade_change5)}"),
    ("下一個升級／降級條件",short_trigger,next_trigger),
])
st.markdown(
    f'<div class="focus-box"><b>今天只看三件事：</b><br>'
    f'① 今日變化：{html.escape(day_change_text)}<br>'
    f'② Wade節奏：{html.escape(wade_timing_text)}（今日 {html.escape(signed_num(wade_day_change))}／5日 {html.escape(signed_num(wade_change5))}）<br>'
    f'③ 操作：{html.escape(operation_level)}；{html.escape(next_trigger)}<br>'
    f'<span class="small-note">資料完整度與訊號一致度分開評估；分數不是未來上漲機率。左右側評分為本系統量化代理，不是 Wade 官方公式。</span></div>',
    unsafe_allow_html=True
)
st.markdown(stage_flow_html(stage),unsafe_allow_html=True)

st.write("### 市場關鍵數據")
cards([
    ("歷史水位",water,f"全期 P{safe_num(full_rank,1)}／近3年 P{safe_num(rolling_rank,1)}"),
    ("RS 強勢股",safe_num(strong_count,0," 檔"),f"占比 {safe_num(strong_pct,2,'%')}／單日 {safe_num(change,0,' 檔')}"),
    ("Wade市場內部",safe_num(wade_score,1," 分"),f"今日 {signed_num(wade_day_change)}／5日 {signed_num(wade_change5)}"),
    ("大盤左右側",market_lr["左右側階段"],f"左 {market_lr['左側分數']:.0f}／右 {market_lr['右側分數']:.0f}"),
    ("資料完整度",f"{quality['資料完整度標籤']}／{quality['資料完整度']:.0f}","今天應有資料是否齊全"),
    ("訊號一致度",f"{quality['訊號一致度標籤']}／{quality['訊號一致度']:.0f}","RS、Wade、廣度、多時間週期是否同向"),
    ("鴨嘴符合",f"{len(all_ok):,} 檔",f"新進 {len(duck_new):,}／退出 {len(duck_exit):,}／A級 {pre_a:,}"),
    ("復甦候選",f"{len(recovery):,} 檔","跌深後開始修復；非買進訊號"),
])

tabs=st.tabs(["🏠 總覽","↔️ 左右側決策","📊 RS／市場廣度","🧭 Wade大盤強弱","🌱 復甦候選","🦆 鴨嘴系統","🔥 鴨嘴×培育中心","📐 台指期","🧪 資料品質","🔄 更新資料"])

with tabs[0]:
    st.subheader("今日一眼看懂")
    decision_cards([
        ("市場狀態",f"{market_signal(stage)}｜{stage}",stage_plain(stage)),
        ("操作",operation_level,market_lr["左右側建議"]),
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
    st.subheader("左右側決策中心")
    st.caption("左側＝價格尚未完全確認前，從價值/低位/修復找機會；右側＝等價格與市場結構證明後再增加部位。以下為本系統量化代理，不是 Wade 官方公式。")
    st.write("### 大盤／指數判讀")
    cards([
        ("整體市場",market_lr["左右側階段"],f"左 {market_lr['左側分數']:.0f}／右 {market_lr['右側分數']:.0f}"),
        ("加權指數",twii_lr["階段"],f"左 {twii_lr['左側分數']:.0f}／右 {twii_lr['右側分數']:.0f}"),
        ("櫃買指數",twoii_lr["階段"],f"左 {twoii_lr['左側分數']:.0f}／右 {twoii_lr['右側分數']:.0f}"),
        ("系統建議",market_lr["左右側建議"],"分數是決策輔助，不是自動下單"),
    ])
    ic1,ic2=st.columns(2)
    with ic1: st.info(f"**加權指數：{twii_lr['階段']}**\n\n{twii_lr['建議']}\n\n依據：{twii_lr['依據']}")
    with ic2: st.info(f"**櫃買指數：{twoii_lr['階段']}**\n\n{twoii_lr['建議']}\n\n依據：{twoii_lr['依據']}")
    st.markdown(f'<div class="action-box"><b>大盤判斷：</b>{html.escape(market_lr["左右側階段"])}<br><b>建議：</b>{html.escape(market_lr["左右側建議"])}<br><b>依據：</b>{html.escape(market_lr["判讀依據"])}<br><span class="small-note">左側高分不等於最低點已出現；右側高分也不代表可以無條件追高。過熱時會優先切換為「向上減碼」。</span></div>',unsafe_allow_html=True)

    st.divider()
    st.write("### 個股左右側判讀")
    if stock_lr.empty:
        st.info("全市場整合資料尚未產生；跑一次正式資料更新後即可使用。")
    else:
        q=st.text_input("搜尋個股代號／名稱",key="lr_stock_search")
        sx=filter_stock_table(stock_lr.copy(),q)
        stages=["全部"]+sorted(sx["左右側階段"].dropna().astype(str).unique().tolist()) if "左右側階段" in sx.columns else ["全部"]
        c1,c2,c3=st.columns([1,1,1])
        chosen=c1.selectbox("階段",stages,key="lr_stage")
        min_left=c2.slider("最低左側分數",0,100,0,5,key="lr_left")
        min_right=c3.slider("最低右側分數",0,100,0,5,key="lr_right")
        if chosen!="全部": sx=sx[sx["左右側階段"]==chosen]
        sx=sx[(pd.to_numeric(sx["左側分數"],errors="coerce").fillna(0)>=min_left)&(pd.to_numeric(sx["右側分數"],errors="coerce").fillna(0)>=min_right)]
        sx=sx.sort_values(["決策排序分數","右側分數","左側分數"],ascending=False,na_position="last")
        show=[c for c in ["代號","名稱","市場","收盤","左右側階段","左側分數","右側分數","系統建議","左側依據","右側依據","RS判讀","復甦分數","鴨嘴階段","預備鴨嘴狀態","鴨嘴完成度%","三率狀態","營收狀態","EPS年增率%","估值狀態","折價率%","合理價"] if c in sx.columns]
        st.caption(f"符合目前篩選：{len(sx):,} 檔。左側偏估值/基本面/修復；右側偏價格/RS/鴨嘴觸發。")
        st.dataframe(sx[show],hide_index=True,use_container_width=True,height=650)
        with st.expander("怎麼解讀各階段"):
            st.markdown("""
- **🌱 左側布局候選**：有估值/基本面/跌深修復優勢，但趨勢尚未充分確認；適合小部位、分批、等待市場證明。
- **↗️ 左轉右確認**：原本左側理由仍在，同時價格開始觸發；可從試單逐步升級基本部位。
- **📈 右側試單**：價格已出現初步觸發，但確認度還沒到高檔；價格繼續證明才增加。
- **🚀 右側確認**：趨勢、型態與相對強度較完整；偏向持有強股，不代表無腦追高。
- **⚠️ 右側過熱／向上減碼**：趨勢仍可能強，但價格過熱；停止追價並考慮分批鎖利。
- **🛑 右側失敗／退出觀察**：型態退出或結構失效，先降低部位，等待下一次觸發。
""")

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
        ("向上減碼觀察",left_alert,f"5日分數變化 {safe_num(wade_change5,1)}"),
        ("早期轉強",early_signal,"市場內部先於指數改善時提示"),
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
    if left_alert!="—": st.warning(f"目前警示：{left_alert}。這代表高檔內部結構開始退潮，屬分批降低風險的觀察訊號，不等於一次清倉。")
    elif early_signal!="—": st.success(f"目前訊號：{early_signal}。市場內部改善速度領先價格表現，可觀察是否持續擴散。")
    else: st.info(f"目前狀態：{wade_state}。尚未觸發明確向上減碼或早期轉強訊號。")
    wc=wade_chart(daily,180)
    if wc is not None: st.altair_chart(wc,use_container_width=True)
    st.markdown("**目前 v0.2 分數組成**：上漲廣度 20%＋RS 水位 20%＋新高/低 10%＋漲跌停 7%＋RS方向 12%＋量價效率 10%＋加權/等權同步 8%＋櫃買相對強弱 6%＋主流延續 7%。")
    cols=[c for c in ["日期","Wade內部強度分數","Wade分數5日變化","Wade市場狀態","左側減碼警示","早期轉強訊號","上漲家數","下跌家數","上漲比例%","漲停近似家數","跌停近似家數","52週新高家數","52週新低家數","上市上漲比例%","上櫃上漲比例%","成交金額20日比%","加權指數報酬%","櫃買指數報酬%","上市等權平均報酬%","上櫃等權平均報酬%","主流延續率%","新強勢領頭股數","量價效率分數","權值同步分數","櫃買相對強弱分數","主流延續分數","市場總結白話","操作建議","市場階段","強勢股方向"] if c in daily.columns]
    if cols:
        hist_show=daily[cols].tail(30).sort_values("日期",ascending=False).copy()
        hist_show=hist_show.rename(columns={"左側減碼警示":"向上減碼觀察"})
        st.dataframe(hist_show,hide_index=True,use_container_width=True,height=520)


with tabs[4]:
    st.subheader("被遺忘資金／災後復甦候選")
    st.caption("v0.1 量化代理：找『距52週高點仍有一段跌幅，但已站回月線、月線上彎、5/20日動能改善，RS尚未過熱』的股票。這是觀察名單，不是買進訊號。")
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
        rq=st.text_input("搜尋代號／名稱",key="recovery_search")
        rx=filter_stock_table(rx,rq)
        if "復甦分數" in rx.columns: rx=rx.sort_values("復甦分數",ascending=False,na_position="last")
        st.dataframe(rx,hide_index=True,use_container_width=True,height=620)

with tabs[5]:
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

with tabs[6]:
    st.subheader("鴨嘴 × 基本面培育中心")
    bmode=st.radio("交集類型",["兩者皆有","三率三升","營收三增","估值買進候選"],horizontal=True)
    bmap={"兩者皆有":both,"三率三升":three_rate,"營收三增":three_rev,"估值買進候選":value_candidates}
    bx=filter_stock_table(bmap[bmode].copy(),st.text_input("搜尋代號／名稱",key="both_search"))
    if bmode=="估值買進候選":
        cols=[c for c in ["日期","代號","名稱","市場","收盤","合理價","折價率%","估值狀態","鴨嘴狀態","月線乖離率%","培育中心類別","是否鴨嘴"] if c in bx.columns]
        if "折價率%" in bx.columns: bx=bx.sort_values("折價率%",ascending=False,na_position="last")
        st.dataframe(bx[cols] if cols else bx,hide_index=True,use_container_width=True,height=620)
    else: st.dataframe(clean_duck(bx),hide_index=True,use_container_width=True,height=620)

with tabs[7]:
    st.subheader("台指期月／季線鴨嘴參考")
    if tx.empty: st.info("目前沒有台指期資料。")
    else:
        first=tx.iloc[0]; mouth=first.get("月季線鴨嘴"); mouth_text="符合" if str(mouth).lower() in {"true","1","是"} else "未符合"
        cards([("收盤",safe_num(first.get("收盤"),0),""),("MA20",safe_num(first.get("MA20"),2),""),("MA60",safe_num(first.get("MA60"),2),""),("月季線鴨嘴",mouth_text,"")])
        st.dataframe(tx,hide_index=True,use_container_width=True)

with tabs[8]:
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

with tabs[9]:
    st.subheader("更新資料")
    st.success("v2.5 延續增量更新：股價只補新交易日；RS 舊歷史沿用；財報/營收已有當期資料就不重抓；PER 歷史只補新月份/缺漏月；Yahoo EPS 使用近期快取。排程仍為週一至週五台灣時間 18:05。")
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

st.divider(); st.caption(f"正式版 v2.6｜決策摘要優化｜RS：{rs_date}｜鴨嘴：{duck_date}｜量化篩選與市場廣度工具，不構成投資建議。")
