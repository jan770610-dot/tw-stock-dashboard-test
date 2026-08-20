# -*- coding: utf-8 -*-
"""失敗交易／停損條件回測 v1.0

目的：只研究「進場後沒有成功發動時，何時退出最好」。
所有訊號均在 T 日收盤後確認，T+1 開盤執行，避免未來資訊。
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import sys
from pathlib import Path
from typing import List, Tuple

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
AUTO_DIR = ROOT / "automation"
if str(AUTO_DIR) not in sys.path:
    sys.path.insert(0, str(AUTO_DIR))
import backtest_entry_strategies as entry_bt

OUT_XLSX = ROOT / "failure_exit_latest.xlsx"
OUT_SUMMARY = ROOT / "failure_exit_summary.csv"
OUT_ROLLUP = ROOT / "failure_exit_rule_rollup.csv"
OUT_ANNUAL = ROOT / "failure_exit_annual.csv"
OUT_STATUS = ROOT / "failure_exit_status.json"

ENTRY_CONTEXTS = [
    ("預備80＋RS70＋市場改善", "pre80_rs70_market"),
    ("右側65＋市場改善", "right65_market"),
    ("正式鴨嘴新進＋市場改善", "duck_new_market"),
]
PROFIT_MODES = [
    ("一般70／領頭90", "profit70_90", 70.0),
    ("一般80／領頭90", "profit80_90", 80.0),
]
STOP_RULES = [
    ("無失敗停損", "none"),
    ("收盤跌幅達-3%", "hard3"),
    ("收盤跌幅達-5%", "hard5"),
    ("收盤跌幅達-7%", "hard7"),
    ("收盤跌幅達-8%", "hard8"),
    ("收盤跌破訊號日低點", "signal_low"),
    ("連續2日跌破MA20", "ma20_2"),
    ("連續3日跌破MA20", "ma20_3"),
    ("右側自高點回落10分", "right_drop10"),
    ("右側自高點回落15分", "right_drop15"),
    ("右側自高點回落20分", "right_drop20"),
    ("寬限5日→連2日跌破MA20", "grace5_ma20_2"),
    ("寬限10日→連2日跌破MA20", "grace10_ma20_2"),
    ("寬限10日→右側回落15分", "grace10_right15"),
    ("寬限10日→MA20或右側確認", "grace10_combo"),
    ("第10日仍未曾+3%", "no_launch10_3"),
    ("第15日仍未曾+5%", "no_launch15_5"),
    ("緊急-8%＋第10日後MA20/右側", "hybrid_emergency8_grace10"),
]


def num(v, default=None):
    try:
        if pd.isna(v):
            return default
        x = float(v)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def cooldown_positions(pos: np.ndarray, cooldown: int) -> np.ndarray:
    if len(pos) <= 1 or cooldown <= 0:
        return pos
    out = [int(pos[0])]
    last = int(pos[0])
    for p in pos[1:]:
        p = int(p)
        if p - last > cooldown:
            out.append(p)
            last = p
    return np.asarray(out, dtype=int)


def first_true(mask: np.ndarray):
    z = np.flatnonzero(mask)
    return int(z[0]) if len(z) else None


def consecutive(mask: np.ndarray, n: int) -> np.ndarray:
    out = np.zeros(len(mask), dtype=bool)
    run = 0
    for i, flag in enumerate(mask.astype(bool)):
        run = run + 1 if flag else 0
        if run >= n:
            out[i] = True
    return out


def profit_mask(heat: np.ndarray, rs: np.ndarray, base: float) -> np.ndarray:
    leader = np.maximum.accumulate(np.nan_to_num(rs, nan=-999.0) >= 85.0)
    threshold = np.where(leader, 90.0, base)
    return np.nan_to_num(heat, nan=-999.0) >= threshold


def build_stop_masks(entry_price, signal_low, closes, highs, ma20, right) -> dict:
    n = len(closes)
    day = np.arange(1, n + 1)
    ret = (closes / entry_price - 1.0) * 100.0
    below = np.isfinite(ma20) & np.isfinite(closes) & (closes < ma20)
    ma2 = consecutive(below, 2)
    ma3 = consecutive(below, 3)
    rc = np.nan_to_num(right, nan=-999.0)
    rpeak = np.maximum.accumulate(rc)
    rdrop = rpeak - rc
    masks = {
        "none": np.zeros(n, bool),
        "hard3": ret <= -3,
        "hard5": ret <= -5,
        "hard7": ret <= -7,
        "hard8": ret <= -8,
        "signal_low": np.isfinite(signal_low) & np.isfinite(closes) & (closes < signal_low),
        "ma20_2": ma2,
        "ma20_3": ma3,
        "right_drop10": rdrop >= 10,
        "right_drop15": rdrop >= 15,
        "right_drop20": rdrop >= 20,
        "grace5_ma20_2": (day >= 6) & ma2,
        "grace10_ma20_2": (day >= 11) & ma2,
        "grace10_right15": (day >= 11) & (rdrop >= 15),
        "grace10_combo": (day >= 11) & (ma2 | (rdrop >= 15)),
        "hybrid_emergency8_grace10": (ret <= -8) | ((day >= 11) & (ma2 | (rdrop >= 15))),
    }
    m10 = np.zeros(n, bool)
    if n >= 10:
        h = np.nanmax(highs[:10]) if np.isfinite(highs[:10]).any() else np.nan
        m10[9] = np.isfinite(h) and h < entry_price * 1.03
    masks["no_launch10_3"] = m10
    m15 = np.zeros(n, bool)
    if n >= 15:
        h = np.nanmax(highs[:15]) if np.isfinite(highs[:15]).any() else np.nan
        m15[14] = np.isfinite(h) and h < entry_price * 1.05
    masks["no_launch15_5"] = m15
    return masks


def net_return(entry_price, exit_price, commission, sell_tax):
    return ((exit_price / entry_price) * (1 - commission - sell_tax) - (1 + commission)) * 100.0


def event_paths(g: pd.DataFrame, sig_idx: int, max_hold: int):
    e = sig_idx + 1
    force = e + max_hold
    if e >= len(g) or force >= len(g):
        return None
    ep = num(g.iloc[e].get("open")) or num(g.iloc[e].get("close"))
    if ep is None or ep <= 0:
        return None
    z = g.iloc[e:force].copy()  # closes that can trigger an exit before force-open
    return {
        "entry_idx": e,
        "force_idx": force,
        "entry_price": float(ep),
        "closes": pd.to_numeric(z["close"], errors="coerce").to_numpy(float),
        "highs": pd.to_numeric(z["high"], errors="coerce").to_numpy(float),
        "lows": pd.to_numeric(z["low"], errors="coerce").to_numpy(float),
        "ma20": pd.to_numeric(z["ma20"], errors="coerce").to_numpy(float),
        "right": pd.to_numeric(z["right"], errors="coerce").to_numpy(float),
        "rs": pd.to_numeric(z["rs"], errors="coerce").to_numpy(float),
        "heat": pd.to_numeric(z["heat"], errors="coerce").to_numpy(float),
        "signal_low": float(num(g.iloc[sig_idx].get("low"), np.nan)),
    }


def simulate_one(g, sig_idx, entry_name, entry_key, profit_name, profit_key, profit_base,
                 stop_name, stop_key, path, commission, sell_tax):
    e, force, ep = path["entry_idx"], path["force_idx"], path["entry_price"]
    p_mask = profit_mask(path["heat"], path["rs"], profit_base)
    s_masks = build_stop_masks(ep, path["signal_low"], path["closes"], path["highs"], path["ma20"], path["right"])
    s_mask = s_masks[stop_key]
    ph, sh = first_true(p_mask), first_true(s_mask)
    reason = "固定60日到期"
    hit = None
    if ph is not None and (sh is None or ph <= sh):
        hit, reason = ph, "獲利成功退出"
    elif sh is not None:
        hit, reason = sh, "失敗停損"
    exit_idx = force if hit is None else min(force, e + hit + 1)
    xp = num(g.iloc[exit_idx].get("open")) or num(g.iloc[exit_idx].get("close")) or ep
    xp = float(xp)
    held = max(1, exit_idx - e)
    hz = g.iloc[e:exit_idx]
    hh = pd.to_numeric(hz["high"], errors="coerce")
    ll = pd.to_numeric(hz["low"], errors="coerce")
    cc = pd.to_numeric(hz["close"], errors="coerce")
    mfe = (float(hh.max()) / ep - 1) * 100 if hh.notna().any() else np.nan
    mae = (float(ll.min()) / ep - 1) * 100 if ll.notna().any() else np.nan
    eq = (cc / ep).dropna().to_numpy(float)
    maxdd = np.nan
    if len(eq):
        peaks = np.maximum.accumulate(eq)
        maxdd = float(np.nanmin((eq / peaks - 1) * 100))

    # Diagnostics after a failure stop. These do not affect the signal.
    later_profit = bool(reason == "失敗停損" and ph is not None and sh is not None and sh < ph)
    back_entry = plus5_entry = plus5_exit = False
    post_mfe = np.nan
    if reason == "失敗停損":
        post = g.iloc[exit_idx:min(len(g), exit_idx + 20)]
        phigh = pd.to_numeric(post["high"], errors="coerce")
        if phigh.notna().any():
            mx = float(phigh.max())
            back_entry = mx >= ep
            plus5_entry = mx >= ep * 1.05
            plus5_exit = mx >= xp * 1.05
            post_mfe = (mx / xp - 1) * 100

    return {
        "進場情境": entry_name, "進場代碼": entry_key,
        "獲利模式": profit_name, "獲利代碼": profit_key,
        "失敗規則": stop_name, "失敗代碼": stop_key,
        "代號": str(g.iloc[sig_idx].get("code")), "名稱": str(g.iloc[sig_idx].get("name") or ""),
        "訊號日": pd.Timestamp(g.iloc[sig_idx]["date"]).date().isoformat(),
        "進場日": pd.Timestamp(g.iloc[e]["date"]).date().isoformat(),
        "退出日": pd.Timestamp(g.iloc[exit_idx]["date"]).date().isoformat(),
        "退出原因": reason, "持有日": int(held),
        "進場價": round(ep, 4), "退出價": round(xp, 4),
        "毛報酬%": round((xp / ep - 1) * 100, 4),
        "淨報酬%": round(net_return(ep, xp, commission, sell_tax), 4),
        "MFE%": round(float(mfe), 4) if np.isfinite(mfe) else np.nan,
        "MAE%": round(float(mae), 4) if np.isfinite(mae) else np.nan,
        "最大回撤%": round(float(maxdd), 4) if np.isfinite(maxdd) else np.nan,
        "停損早於後續獲利訊號": later_profit,
        "停損後20日回到進場價": back_entry,
        "停損後20日達原進場+5%": plus5_entry,
        "停損後20日自退出價反彈+5%": plus5_exit,
        "停損後20日最大反彈%": round(float(post_mfe), 4) if np.isfinite(post_mfe) else np.nan,
    }


def simulate_all(signals, start, end, max_hold, cooldown, commission, sell_tax):
    rows: List[dict] = []
    events = skipped = 0
    groups = {str(code): g.sort_values("date").reset_index(drop=True)
              for code, g in signals.groupby("code", observed=True)}
    for gi, (code, g) in enumerate(groups.items(), 1):
        masks = entry_bt.build_strategy_masks(g)
        dates = pd.to_datetime(g["date"], errors="coerce")
        for entry_name, entry_key in ENTRY_CONTEXTS:
            mask = masks.get(entry_key)
            if mask is None:
                continue
            pos = cooldown_positions(np.flatnonzero(mask.to_numpy(bool)), cooldown)
            for sig_idx in pos:
                sig_idx = int(sig_idx)
                d = dates.iloc[sig_idx].date()
                if d < start or d > end:
                    continue
                events += 1
                path = event_paths(g, sig_idx, max_hold)
                if path is None:
                    skipped += 1
                    continue
                for profit_name, profit_key, profit_base in PROFIT_MODES:
                    # Build the common stop masks once per profit mode/event through simulate_one simplicity.
                    for stop_name, stop_key in STOP_RULES:
                        rows.append(simulate_one(g, sig_idx, entry_name, entry_key,
                            profit_name, profit_key, profit_base, stop_name, stop_key,
                            path, commission, sell_tax))
        if gi == 1 or gi % 250 == 0 or gi == len(groups):
            print(f"[failure] {gi}/{len(groups)} stocks | rows={len(rows)}", flush=True)
    return pd.DataFrame(rows), {
        "entry_contexts": len(ENTRY_CONTEXTS), "profit_modes": len(PROFIT_MODES),
        "stop_rules": len(STOP_RULES),
        "context_combinations": len(ENTRY_CONTEXTS) * len(PROFIT_MODES) * len(STOP_RULES),
        "entry_events_after_cooldown": int(events), "skipped_recent_no_full_horizon": int(skipped),
        "trade_rows": int(len(rows)),
    }


def annual_summary(trades):
    x = trades.copy()
    x["年度"] = pd.to_datetime(x["進場日"]).dt.year
    rows = []
    keys = ["進場情境", "進場代碼", "獲利模式", "獲利代碼", "失敗規則", "失敗代碼", "年度"]
    for vals, g in x.groupby(keys, dropna=False):
        en, ek, pn, pk, sn, sk, year = vals
        r = pd.to_numeric(g["淨報酬%"], errors="coerce")
        dd = pd.to_numeric(g["最大回撤%"], errors="coerce")
        rows.append({"進場情境": en, "進場代碼": ek, "獲利模式": pn, "獲利代碼": pk,
            "失敗規則": sn, "失敗代碼": sk, "年度": int(year), "樣本數": len(g),
            "平均淨報酬%": round(float(r.mean()),3), "中位淨報酬%": round(float(r.median()),3),
            "勝率%": round(float((r>0).mean()*100),2), "10分位淨報酬%": round(float(r.quantile(.10)),3),
            "中位最大回撤%": round(float(dd.median()),3)})
    return pd.DataFrame(rows)


def summarize(trades, annual):
    rows = []
    keys = ["進場情境", "進場代碼", "獲利模式", "獲利代碼", "失敗規則", "失敗代碼"]
    for vals, g in trades.groupby(keys, dropna=False):
        en, ek, pn, pk, sn, sk = vals
        r = pd.to_numeric(g["淨報酬%"], errors="coerce")
        dd = pd.to_numeric(g["最大回撤%"], errors="coerce")
        mae = pd.to_numeric(g["MAE%"], errors="coerce")
        hold = pd.to_numeric(g["持有日"], errors="coerce")
        stopped = g[g["退出原因"] == "失敗停損"]
        a = annual[(annual["進場代碼"]==ek)&(annual["獲利代碼"]==pk)&(annual["失敗代碼"]==sk)]
        a = a[a["樣本數"]>=30]
        av = pd.to_numeric(a["中位淨報酬%"], errors="coerce")
        rows.append({"進場情境":en,"進場代碼":ek,"獲利模式":pn,"獲利代碼":pk,"失敗規則":sn,"失敗代碼":sk,
            "樣本數":len(g),"平均淨報酬%":round(float(r.mean()),3),"中位淨報酬%":round(float(r.median()),3),
            "勝率%":round(float((r>0).mean()*100),2),"10分位淨報酬%":round(float(r.quantile(.10)),3),
            "5分位淨報酬%":round(float(r.quantile(.05)),3),"中位MAE%":round(float(mae.median()),3),
            "中位最大回撤%":round(float(dd.median()),3),"中位持有日":round(float(hold.median()),1),
            "失敗停損率%":round(float((g["退出原因"]=="失敗停損").mean()*100),2),
            "獲利成功退出率%":round(float((g["退出原因"]=="獲利成功退出").mean()*100),2),
            "固定到期率%":round(float((g["退出原因"]=="固定60日到期").mean()*100),2),
            "停損後回到進場價率%":round(float(stopped["停損後20日回到進場價"].astype(bool).mean()*100),2) if len(stopped) else np.nan,
            "停損後達原進場+5率%":round(float(stopped["停損後20日達原進場+5%"].astype(bool).mean()*100),2) if len(stopped) else np.nan,
            "停損後自退出價反彈5率%":round(float(stopped["停損後20日自退出價反彈+5%"].astype(bool).mean()*100),2) if len(stopped) else np.nan,
            "停損早於後續獲利訊號率%":round(float(stopped["停損早於後續獲利訊號"].astype(bool).mean()*100),2) if len(stopped) else np.nan,
            "年度有效年數":len(a),"年度正中位報酬率%":round(float((av>0).mean()*100),2) if len(av) else np.nan,
            "最差年度中位報酬%":round(float(av.min()),3) if len(av) else np.nan})
    out = pd.DataFrame(rows)
    base = out[out["失敗代碼"]=="none"][["進場代碼","獲利代碼","中位淨報酬%","平均淨報酬%","勝率%","10分位淨報酬%","中位最大回撤%"]].rename(columns={
        "中位淨報酬%":"基準中位淨報酬%","平均淨報酬%":"基準平均淨報酬%","勝率%":"基準勝率%",
        "10分位淨報酬%":"基準10分位淨報酬%","中位最大回撤%":"基準中位最大回撤%"})
    out = out.merge(base,on=["進場代碼","獲利代碼"],how="left")
    out["中位報酬改善pp"]=(out["中位淨報酬%"]-out["基準中位淨報酬%"]).round(3)
    out["平均報酬改善pp"]=(out["平均淨報酬%"]-out["基準平均淨報酬%"]).round(3)
    out["勝率改善pp"]=(out["勝率%"]-out["基準勝率%"]).round(2)
    out["10分位改善pp"]=(out["10分位淨報酬%"]-out["基準10分位淨報酬%"]).round(3)
    out["最大回撤改善pp"]=(out["中位最大回撤%"]-out["基準中位最大回撤%"]).round(3)
    def pct(s,higher=True):
        x=pd.to_numeric(s,errors="coerce")
        if not higher: x=-x
        return x.rank(pct=True,method="average").fillna(0)*100
    score=(.20*pct(out["中位報酬改善pp"])+.15*pct(out["平均報酬改善pp"])+.10*pct(out["勝率改善pp"])
           +.20*pct(out["10分位改善pp"])+.15*pct(out["最大回撤改善pp"])
           +.10*pct(out["停損後達原進場+5率%"],False)+.10*pct(out["年度正中位報酬率%"]))
    out["情境綜合分數"]=score.round(2)
    out["情境排名"]=out.groupby(["進場代碼","獲利代碼"])["情境綜合分數"].rank(ascending=False,method="min").astype(int)
    return out.sort_values(["進場情境","獲利模式","情境排名"]).reset_index(drop=True)


def rule_rollup(summary):
    rows=[]
    for (sn,sk),g in summary.groupby(["失敗規則","失敗代碼"]):
        ranks=pd.to_numeric(g["情境排名"],errors="coerce")
        rows.append({"失敗規則":sn,"失敗代碼":sk,"跨情境數":len(g),"平均情境排名":round(float(ranks.mean()),2),
            "排名前3比例%":round(float((ranks<=3).mean()*100),2),
            "中位報酬改善pp_中位":round(float(pd.to_numeric(g["中位報酬改善pp"],errors="coerce").median()),3),
            "平均報酬改善pp_中位":round(float(pd.to_numeric(g["平均報酬改善pp"],errors="coerce").median()),3),
            "10分位改善pp_中位":round(float(pd.to_numeric(g["10分位改善pp"],errors="coerce").median()),3),
            "最大回撤改善pp_中位":round(float(pd.to_numeric(g["最大回撤改善pp"],errors="coerce").median()),3),
            "失敗停損率%_中位":round(float(pd.to_numeric(g["失敗停損率%"],errors="coerce").median()),2),
            "停損後達原進場+5率%_中位":round(float(pd.to_numeric(g["停損後達原進場+5率%"],errors="coerce").median()),2),
            "停損早於後續獲利訊號率%_中位":round(float(pd.to_numeric(g["停損早於後續獲利訊號率%"],errors="coerce").median()),2),
            "年度正中位報酬率%_中位":round(float(pd.to_numeric(g["年度正中位報酬率%"],errors="coerce").median()),2),
            "平均情境綜合分數":round(float(pd.to_numeric(g["情境綜合分數"],errors="coerce").mean()),2)})
    out=pd.DataFrame(rows)
    elig=out[out["失敗代碼"]!="none"].copy(); elig["總排名"]=elig["平均情境綜合分數"].rank(ascending=False,method="min").astype(int)
    base=out[out["失敗代碼"]=="none"].copy(); base["總排名"]=999
    return pd.concat([elig,base],ignore_index=True).sort_values(["總排名","平均情境排名"]).reset_index(drop=True)


def write_outputs(summary,rollup,annual,trades,status):
    summary.to_csv(OUT_SUMMARY,index=False,encoding="utf-8-sig")
    rollup.to_csv(OUT_ROLLUP,index=False,encoding="utf-8-sig")
    annual.to_csv(OUT_ANNUAL,index=False,encoding="utf-8-sig")
    OUT_STATUS.write_text(json.dumps(status,ensure_ascii=False,indent=2),encoding="utf-8")
    top=rollup[rollup["失敗代碼"]!="none"].head(5)["失敗代碼"].tolist()
    detail=trades[trades["失敗代碼"].isin(top)].head(300000).copy()
    notes=pd.DataFrame({"說明":[
        "本回測只研究失敗交易怎麼退；獲利成功端固定為一般70/80、RS85領頭股90。",
        "所有訊號T日收盤確認、T+1開盤執行。",
        "硬停損採收盤報酬觸發，不假設盤中碰價即可成交。",
        "寬限10日：前10日不因MA20/右側正常震盪退出；混合規則仍保留-8%緊急風控。",
        "第10日仍未曾+3%、第15日仍未曾+5%用來測試『沒發動就離場』。",
        "停損後20日又達原進場+5%代表可能把後續贏家洗掉，比例越低越理想。",
        "排名以同一進場/獲利情境的無停損基準作比較，重視報酬改善、尾端損失、回撤、假停損與年度穩定度。",
        "這是事件交易研究，不是資金上限固定的投資組合回測。"]})
    with pd.ExcelWriter(OUT_XLSX,engine="xlsxwriter") as w:
        rollup.to_excel(w,sheet_name="失敗規則總排行",index=False); summary.to_excel(w,sheet_name="各情境比較",index=False)
        annual.to_excel(w,sheet_name="年度驗證",index=False); detail.to_excel(w,sheet_name="前5規則交易明細",index=False); notes.to_excel(w,sheet_name="說明",index=False)
        head=w.book.add_format({"bold":True,"bg_color":"#D9EAF7","border":1})
        for sh,df in [("失敗規則總排行",rollup),("各情境比較",summary),("年度驗證",annual),("前5規則交易明細",detail),("說明",notes)]:
            ws=w.sheets[sh]; ws.freeze_panes(1,0)
            if len(df.columns): ws.autofilter(0,0,max(0,len(df)),len(df.columns)-1)
            for j,c in enumerate(df.columns):
                ws.write(0,j,c,head); ws.set_column(j,j,110 if sh=="說明" else min(42,max(11,len(str(c))+2)))


def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--start",default="2023-01-03"); ap.add_argument("--end",default="")
    ap.add_argument("--max-hold",type=int,default=60); ap.add_argument("--cooldown",type=int,default=10); ap.add_argument("--workers",type=int,default=2)
    ap.add_argument("--fetch-missing",action="store_true"); ap.add_argument("--commission-rate",type=float,default=.001425); ap.add_argument("--sell-tax-rate",type=float,default=.003)
    a=ap.parse_args(); start=entry_bt._date(a.start); requested=entry_bt._date(a.end) if a.end else entry_bt._latest_formal_date()
    warmup=start-dt.timedelta(days=entry_bt.WARMUP_CALENDAR_DAYS); max_hold=max(20,min(80,a.max_hold)); cooldown=max(0,min(30,a.cooldown))
    print("=== 失敗交易／停損條件回測 v1.0 ===",flush=True)
    cache=entry_bt.ensure_history_cache(warmup,requested,a.fetch_missing,a.workers); raw,meta=entry_bt.load_market_cache(warmup,requested)
    effective=min(requested,pd.to_datetime(meta["last_date"]).date()); signals,_=entry_bt.compute_signals(raw)
    trades,sim=simulate_all(signals,start,effective,max_hold,cooldown,max(0,a.commission_rate),max(0,a.sell_tax_rate))
    if trades.empty: raise RuntimeError("沒有可用停損研究交易")
    annual=annual_summary(trades); summary=summarize(trades,annual); rollup=rule_rollup(summary)
    leaders=[]
    for _,r in rollup[rollup["失敗代碼"]!="none"].head(8).iterrows():
        leaders.append({"排名":int(r["總排名"]),"失敗規則":str(r["失敗規則"]),"平均情境綜合分數":num(r["平均情境綜合分數"]),
            "平均情境排名":num(r["平均情境排名"]),"中位報酬改善pp":num(r["中位報酬改善pp_中位"]),"10分位改善pp":num(r["10分位改善pp_中位"]),
            "最大回撤改善pp":num(r["最大回撤改善pp_中位"]),"停損後達原進場+5率%":num(r["停損後達原進場+5率%_中位"])})
    status={"status":"success","generated_at_taipei":dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).strftime("%Y-%m-%d %H:%M:%S"),
        "research_start":start.isoformat(),"research_end_requested":requested.isoformat(),"research_end_effective":effective.isoformat(),"warmup_start":warmup.isoformat(),
        "max_hold_sessions":max_hold,"cooldown_sessions":cooldown,"commission_rate_per_side":a.commission_rate,"sell_tax_rate":a.sell_tax_rate,
        "cache":cache,"data":meta,"simulation":sim,"summary_rows":len(summary),"rollup_rows":len(rollup),"annual_rows":len(annual),
        "method":"T close signal -> T+1 open execution; failure-rule deltas versus same-context no-stop baseline","current_leaders":leaders,
        "outputs":[OUT_XLSX.name,OUT_SUMMARY.name,OUT_ROLLUP.name,OUT_ANNUAL.name,OUT_STATUS.name]}
    write_outputs(summary,rollup,annual,trades,status)
    cols=["總排名","失敗規則","平均情境排名","排名前3比例%","中位報酬改善pp_中位","10分位改善pp_中位","最大回撤改善pp_中位","停損後達原進場+5率%_中位","平均情境綜合分數"]
    print(rollup[rollup["失敗代碼"]!="none"][cols].head(12).to_string(index=False),flush=True)

if __name__=="__main__": main()
