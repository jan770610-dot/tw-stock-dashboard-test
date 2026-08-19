# -*- coding: utf-8 -*-
"""Combined entry × exit backtest for the Taiwan-stock decision system.

Uses the already-validated point-in-time signal reconstruction from
automation/backtest_entry_strategies.py.

Research design
---------------
Entry candidates:
1) 預備80＋RS>=70＋市場廣度改善（一次買）
2) 右側首次>=65＋市場廣度改善（一次買）
3) 鴨嘴新進＋市場廣度改善（一次買）
4) 預備80先50% → 右側>=65再50%
5) 預備80先33% → 右側>=65再33% → 正式鴨嘴再34%

Exit candidates:
- 過熱70
- 過熱80
- 動態70/90（成為RS85領頭股則放寬至90）
- 動態80/90
- 過熱70後RS自高點回落5
- 分層獲利管理（部分減碼）
- 只做結構風險退出
- 固定60日基準

All signals are known after T close and executed on T+1 open.
The backtest includes configurable commission and stock sell tax.

This is an event-trade research backtest. It is NOT a capital-constrained
portfolio simulation; trades across different stocks may overlap.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import sys
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
AUTO_DIR = ROOT / "automation"
if str(AUTO_DIR) not in sys.path:
    sys.path.insert(0, str(AUTO_DIR))

import backtest_entry_strategies as entry_bt  # noqa: E402

OUT_XLSX = ROOT / "entry_exit_combo_latest.xlsx"
OUT_SUMMARY = ROOT / "entry_exit_combo_summary.csv"
OUT_ANNUAL = ROOT / "entry_exit_combo_annual.csv"
OUT_STATUS = ROOT / "entry_exit_combo_status.json"

ENTRY_PLANS = [
    {
        "name": "預備80＋RS70＋市場改善｜一次買",
        "key": "single_pre80",
        "start_key": "pre80_rs70_market",
        "kind": "一次買",
        "initial_weight": 1.00,
        "add_b": 0.00,
        "add_c": 0.00,
    },
    {
        "name": "右側65＋市場改善｜一次買",
        "key": "single_right65",
        "start_key": "right65_market",
        "kind": "一次買",
        "initial_weight": 1.00,
        "add_b": 0.00,
        "add_c": 0.00,
    },
    {
        "name": "正式鴨嘴新進＋市場改善｜一次買",
        "key": "single_duck",
        "start_key": "duck_new_market",
        "kind": "一次買",
        "initial_weight": 1.00,
        "add_b": 0.00,
        "add_c": 0.00,
    },
    {
        "name": "預備80試單50%→右側65加滿",
        "key": "stage_50_50",
        "start_key": "pre80_rs70_market",
        "kind": "分批進場",
        "initial_weight": 0.50,
        "add_b": 0.50,
        "add_c": 0.00,
    },
    {
        "name": "預備80試33%→右側65加33%→正式鴨嘴加34%",
        "key": "stage_33_33_34",
        "start_key": "pre80_rs70_market",
        "kind": "分批進場",
        "initial_weight": 0.33,
        "add_b": 0.33,
        "add_c": 0.34,
    },
]

EXIT_PLANS = [
    {"name": "過熱70＋結構風險", "key": "heat70", "kind": "全出"},
    {"name": "過熱80＋結構風險", "key": "heat80", "kind": "全出"},
    {"name": "動態70/90＋結構風險", "key": "dynamic70_90", "kind": "全出"},
    {"name": "動態80/90＋結構風險", "key": "dynamic80_90", "kind": "全出"},
    {"name": "過熱70後RS回落5＋結構風險", "key": "hot70_rs5", "kind": "全出"},
    {"name": "分層獲利管理＋結構風險", "key": "tiered", "kind": "分批出場"},
    {"name": "只做結構風險退出", "key": "protective", "kind": "風險基準"},
    {"name": "固定60日（不做結構風險）", "key": "hold60", "kind": "持有基準"},
]


def _num(v, default=None):
    try:
        if pd.isna(v):
            return default
        x = float(v)
        return x if math.isfinite(x) else default
    except Exception:
        return default


def _cooldown_positions(pos: np.ndarray, cooldown: int) -> np.ndarray:
    if len(pos) <= 1 or cooldown <= 0:
        return pos
    keep = [int(pos[0])]
    last = int(pos[0])
    for p in pos[1:]:
        p = int(p)
        if p - last > cooldown:
            keep.append(p)
            last = p
    return np.asarray(keep, dtype=int)


def _price_at_open(g: pd.DataFrame, idx: int) -> float:
    p = _num(g.iloc[idx].get("open"))
    if p is None or p <= 0:
        p = _num(g.iloc[idx].get("close"))
    return float(p) if p is not None and p > 0 else np.nan


def _entry_add_b(row: pd.Series) -> bool:
    right = _num(row.get("right"), 0.0) or 0.0
    heat = _num(row.get("heat"), 999.0)
    market = bool(row.get("市場廣度改善", False))
    return right >= 65 and heat is not None and heat < 70 and market


def _entry_add_c(row: pd.Series) -> bool:
    heat = _num(row.get("heat"), 999.0)
    return bool(row.get("duck", False)) and heat is not None and heat < 70


def _protective_fail(row: pd.Series) -> bool:
    close = _num(row.get("close"))
    ma20 = _num(row.get("ma20"))
    ma20chg = _num(row.get("ma20_change"))
    right = _num(row.get("right"), 100.0) or 100.0

    ma20_fail = (
        close is not None
        and ma20 is not None
        and ma20chg is not None
        and close < ma20
        and ma20chg <= 0
    )
    right_fail = right < 45
    return bool(ma20_fail or right_fail)


def _new_lot(weight: float, buy_price: float, commission: float) -> dict:
    return {
        "weight": float(weight),
        "remaining": float(weight),
        "buy": float(buy_price),
        "buy_commission": float(weight) * float(commission),
    }


def _position_weight(lots: List[dict]) -> float:
    return float(sum(max(0.0, x["remaining"]) for x in lots))


def _sell_fraction(
    lots: List[dict],
    fraction: float,
    sell_price: float,
    commission: float,
    sell_tax: float,
) -> Tuple[float, float, float]:
    """Sell fraction of every open lot; return gross pnl, net pnl, sold weight."""
    fraction = max(0.0, min(1.0, float(fraction)))
    gross_pnl = 0.0
    net_pnl = 0.0
    sold_weight = 0.0
    for lot in lots:
        rem = float(lot["remaining"])
        if rem <= 1e-12:
            continue
        sw = rem * fraction
        if sw <= 0:
            continue
        ratio = float(sell_price) / float(lot["buy"])
        gross_pnl += sw * (ratio - 1.0)
        # Each allocated weight represents pre-fee buy notional.
        # Net pnl = sale value after sell costs - original buy notional - buy commission.
        buy_fee_share = sw * commission
        net_pnl += sw * (ratio * (1.0 - commission - sell_tax) - 1.0) - buy_fee_share
        lot["remaining"] -= sw
        sold_weight += sw
    return gross_pnl, net_pnl, sold_weight


def _mark_equity(
    lots: List[dict],
    realized_gross: float,
    realized_net: float,
    close_price: float,
    commission: float,
) -> Tuple[float, float]:
    gross = 1.0 + realized_gross
    net = 1.0 + realized_net
    for lot in lots:
        rem = float(lot["remaining"])
        if rem <= 1e-12:
            continue
        ratio = float(close_price) / float(lot["buy"])
        gross += rem * (ratio - 1.0)
        # Mark open lots after buy commission, without assuming a future sell cost yet.
        net += rem * (ratio - 1.0) - rem * commission
    return gross, net


def _exit_signal(
    key: str,
    row: pd.Series,
    state: dict,
) -> Tuple[str, float]:
    """Return action: none/full/partial1/partial2 and fraction for next open."""
    heat = _num(row.get("heat"), 0.0) or 0.0
    rs = _num(row.get("rs"))

    if rs is not None:
        if state.get("peak_rs") is None:
            state["peak_rs"] = rs
        else:
            state["peak_rs"] = max(float(state["peak_rs"]), rs)
        if rs >= 85:
            state["leader"] = True

    if heat >= 70:
        state["hot70"] = True

    peak_rs = _num(state.get("peak_rs"))
    rs_drop5 = (
        rs is not None
        and peak_rs is not None
        and rs <= peak_rs - 5.0
    )

    if key == "heat70" and heat >= 70:
        return "full", 1.0
    if key == "heat80" and heat >= 80:
        return "full", 1.0
    if key == "dynamic70_90":
        threshold = 90 if state.get("leader") else 70
        if heat >= threshold:
            return "full", 1.0
    if key == "dynamic80_90":
        threshold = 90 if state.get("leader") else 80
        if heat >= threshold:
            return "full", 1.0
    if key == "hot70_rs5" and state.get("hot70") and rs_drop5:
        return "full", 1.0

    if key == "tiered":
        leader = bool(state.get("leader"))
        first_threshold = 85 if leader else 70
        second_threshold = 90 if leader else 80

        if not state.get("tier1") and heat >= first_threshold:
            state["tier1"] = True
            state["profit_started"] = True
            return "partial1", 1.0 / 3.0

        if state.get("tier1") and not state.get("tier2") and heat >= second_threshold:
            state["tier2"] = True
            state["profit_started"] = True
            return "partial2", 0.5  # sell half of remaining ~= second third

        if state.get("hot70") and rs_drop5:
            return "full", 1.0

    return "none", 0.0


def simulate_one_trade(
    g: pd.DataFrame,
    signal_idx: int,
    entry_plan: dict,
    exit_plan: dict,
    max_hold: int,
    commission: float,
    sell_tax: float,
) -> dict | None:
    entry_idx = signal_idx + 1
    force_idx = entry_idx + max_hold
    if entry_idx >= len(g) or force_idx >= len(g):
        return None

    first_price = _price_at_open(g, entry_idx)
    if not np.isfinite(first_price) or first_price <= 0:
        return None

    lots: List[dict] = []
    initial_w = float(entry_plan["initial_weight"])
    lots.append(_new_lot(initial_w, first_price, commission))
    buy_tx = 1
    sell_tx = 0
    max_invested = initial_w
    stage_b_done = entry_plan["add_b"] <= 0
    stage_c_done = entry_plan["add_c"] <= 0

    # If the start-day close already satisfies the add-B state, both tranches
    # are legitimately known before the same T+1 open.
    start_row = g.iloc[signal_idx]
    if not stage_b_done and _entry_add_b(start_row):
        w = float(entry_plan["add_b"])
        lots.append(_new_lot(w, first_price, commission))
        stage_b_done = True
        buy_tx += 1
        max_invested += w
    if not stage_c_done and _entry_add_c(start_row):
        w = float(entry_plan["add_c"])
        lots.append(_new_lot(w, first_price, commission))
        stage_c_done = True
        buy_tx += 1
        max_invested += w

    state = {
        "peak_rs": _num(start_row.get("rs")),
        "leader": bool((_num(start_row.get("rs"), -999) or -999) >= 85),
        "hot70": bool((_num(start_row.get("heat"), 0) or 0) >= 70),
        "tier1": False,
        "tier2": False,
        "profit_started": False,
    }

    realized_gross = 0.0
    realized_net = 0.0
    equity_net_path = []
    equity_gross_path = []
    exit_reason = "固定持有到期"
    exit_date = None
    exit_price = np.nan
    first_profit_signal_day = None

    # Pending action from prior close; tuple(kind, fraction, reason)
    pending_exit = None
    pending_b = False
    pending_c = False

    for i in range(entry_idx, force_idx + 1):
        open_price = _price_at_open(g, i)
        if not np.isfinite(open_price) or open_price <= 0:
            continue

        # Execute actions known after previous close.
        if pending_exit is not None:
            kind, frac, reason = pending_exit
            gp, npnl, sw = _sell_fraction(lots, frac, open_price, commission, sell_tax)
            realized_gross += gp
            realized_net += npnl
            if sw > 1e-12:
                sell_tx += 1
            pending_exit = None
            if _position_weight(lots) <= 1e-10:
                exit_reason = reason
                exit_date = pd.Timestamp(g.iloc[i]["date"]).date().isoformat()
                exit_price = float(open_price)
                break

        if not state.get("profit_started"):
            if pending_b and not stage_b_done:
                w = float(entry_plan["add_b"])
                lots.append(_new_lot(w, open_price, commission))
                stage_b_done = True
                buy_tx += 1
                max_invested += w
            if pending_c and not stage_c_done:
                w = float(entry_plan["add_c"])
                lots.append(_new_lot(w, open_price, commission))
                stage_c_done = True
                buy_tx += 1
                max_invested += w
        pending_b = False
        pending_c = False

        close_price = _num(g.iloc[i].get("close"))
        if close_price is not None and close_price > 0:
            gross_eq, net_eq = _mark_equity(
                lots, realized_gross, realized_net, close_price, commission
            )
            equity_gross_path.append(gross_eq)
            equity_net_path.append(net_eq)

        if i >= force_idx:
            # At the force-day open, liquidate any remaining position.
            gp, npnl, sw = _sell_fraction(lots, 1.0, open_price, commission, sell_tax)
            realized_gross += gp
            realized_net += npnl
            if sw > 1e-12:
                sell_tx += 1
            exit_reason = "固定持有到期"
            exit_date = pd.Timestamp(g.iloc[i]["date"]).date().isoformat()
            exit_price = float(open_price)
            break

        row = g.iloc[i]

        # hold60 intentionally ignores protective exits as a pure baseline.
        if exit_plan["key"] != "hold60" and _protective_fail(row):
            pending_exit = ("full", 1.0, "結構風險退出")
            continue

        action, frac = _exit_signal(exit_plan["key"], row, state)
        if action != "none":
            if first_profit_signal_day is None:
                first_profit_signal_day = pd.Timestamp(row["date"]).date().isoformat()
            if action == "full":
                pending_exit = ("full", 1.0, "獲利/退潮退出")
                state["profit_started"] = True
            else:
                pending_exit = (action, frac, "分層獲利")
            continue

        if exit_plan["key"] == "protective":
            # no profit-taking rule, only protective + time exit
            pass

        if not state.get("profit_started"):
            if not stage_b_done and _entry_add_b(row):
                pending_b = True
            if not stage_c_done and _entry_add_c(row):
                pending_c = True

    if exit_date is None:
        # Safety fallback.
        idx = min(force_idx, len(g) - 1)
        p = _price_at_open(g, idx)
        if not np.isfinite(p) or p <= 0:
            p = _num(g.iloc[idx].get("close"), first_price) or first_price
        gp, npnl, sw = _sell_fraction(lots, 1.0, p, commission, sell_tax)
        realized_gross += gp
        realized_net += npnl
        if sw > 1e-12:
            sell_tx += 1
        exit_reason = "固定持有到期"
        exit_date = pd.Timestamp(g.iloc[idx]["date"]).date().isoformat()
        exit_price = float(p)

    entry_date = pd.Timestamp(g.iloc[entry_idx]["date"]).date()
    exit_dt = dt.datetime.strptime(exit_date, "%Y-%m-%d").date()

    net_arr = np.asarray(equity_net_path, dtype=float)
    gross_arr = np.asarray(equity_gross_path, dtype=float)
    if net_arr.size:
        net_rets = (net_arr - 1.0) * 100.0
        mfe = float(np.nanmax(net_rets))
        mae = float(np.nanmin(net_rets))
        peaks = np.maximum.accumulate(net_arr)
        dd = np.where(peaks > 0, (net_arr / peaks - 1.0) * 100.0, 0.0)
        max_dd = float(np.nanmin(dd))
    else:
        mfe = mae = max_dd = np.nan

    deployed = max(1e-9, float(max_invested))
    target_net_ret = realized_net * 100.0
    gross_ret = realized_gross * 100.0
    deployed_net_ret = target_net_ret / deployed

    return {
        "進場方案": entry_plan["name"],
        "進場代碼": entry_plan["key"],
        "進場類型": entry_plan["kind"],
        "出場方案": exit_plan["name"],
        "出場代碼": exit_plan["key"],
        "出場類型": exit_plan["kind"],
        "代號": str(g.iloc[signal_idx].get("code")),
        "名稱": str(g.iloc[signal_idx].get("name") or ""),
        "市場": str(g.iloc[signal_idx].get("market") or ""),
        "訊號日": pd.Timestamp(g.iloc[signal_idx]["date"]).date().isoformat(),
        "進場日": entry_date.isoformat(),
        "出場日": exit_date,
        "持有交易日": int(max(1, (pd.to_datetime(g["date"]).dt.date.tolist().index(exit_dt) - entry_idx + 1))),
        "首次進場價": round(float(first_price), 4),
        "最後出場價": round(float(exit_price), 4),
        "進場RS": round(_num(g.iloc[signal_idx].get("rs"), np.nan), 2),
        "進場右側": round(_num(g.iloc[signal_idx].get("right"), np.nan), 2),
        "進場過熱": round(_num(g.iloc[signal_idx].get("heat"), np.nan), 2),
        "進場完成度": round(_num(g.iloc[signal_idx].get("complete"), np.nan), 1),
        "進場市場上漲比例": round(_num(g.iloc[signal_idx].get("上漲比例"), np.nan), 2),
        "最大投入比例%": round(max_invested * 100.0, 1),
        "是否加滿": bool(max_invested >= 0.99),
        "買進交易次數": int(buy_tx),
        "賣出交易次數": int(sell_tx),
        "毛報酬%": round(gross_ret, 4),
        "淨報酬_目標資金%": round(target_net_ret, 4),
        "淨報酬_已投入資金%": round(deployed_net_ret, 4),
        "交易中MFE_淨值%": round(mfe, 4) if np.isfinite(mfe) else np.nan,
        "交易中MAE_淨值%": round(mae, 4) if np.isfinite(mae) else np.nan,
        "交易中最大回撤%": round(max_dd, 4) if np.isfinite(max_dd) else np.nan,
        "退出原因": exit_reason,
        "首次獲利訊號日": first_profit_signal_day or "",
    }


def simulate_all(
    signals: pd.DataFrame,
    start: dt.date,
    end: dt.date,
    max_hold: int,
    cooldown: int,
    commission: float,
    sell_tax: float,
) -> Tuple[pd.DataFrame, dict]:
    rows = []
    skipped = 0
    events = 0

    groups = {
        str(code): g.sort_values("date").reset_index(drop=True)
        for code, g in signals.groupby("code", observed=True)
    }

    for gi, (code, g) in enumerate(groups.items(), 1):
        masks = entry_bt.build_strategy_masks(g)
        dates = pd.to_datetime(g["date"], errors="coerce")

        for ep in ENTRY_PLANS:
            mask = masks.get(ep["start_key"])
            if mask is None:
                continue
            pos = np.flatnonzero(mask.to_numpy(dtype=bool))
            pos = _cooldown_positions(pos, cooldown)

            for sig_idx in pos:
                sig_day = dates.iloc[sig_idx].date()
                if sig_day < start or sig_day > end:
                    continue
                events += 1
                if sig_idx + 1 + max_hold >= len(g):
                    skipped += 1
                    continue
                for xp in EXIT_PLANS:
                    r = simulate_one_trade(
                        g, int(sig_idx), ep, xp, max_hold, commission, sell_tax
                    )
                    if r is not None:
                        rows.append(r)

        if gi == 1 or gi % 250 == 0 or gi == len(groups):
            print(f"[combo] {gi}/{len(groups)} stocks | rows={len(rows)}", flush=True)

    return pd.DataFrame(rows), {
        "entry_events_after_cooldown": int(events),
        "skipped_recent_no_full_horizon": int(skipped),
        "entry_plan_count": len(ENTRY_PLANS),
        "exit_plan_count": len(EXIT_PLANS),
        "combo_count": len(ENTRY_PLANS) * len(EXIT_PLANS),
        "trade_rows": int(len(rows)),
    }


def annual_summary(trades: pd.DataFrame) -> pd.DataFrame:
    if trades.empty:
        return pd.DataFrame()
    x = trades.copy()
    x["年度"] = pd.to_datetime(x["進場日"], errors="coerce").dt.year
    rows = []
    keys = ["進場方案", "進場代碼", "出場方案", "出場代碼", "年度"]
    for vals, g in x.groupby(keys, dropna=False):
        entry_name, entry_key, exit_name, exit_key, year = vals
        ret = pd.to_numeric(g["淨報酬_目標資金%"], errors="coerce")
        dd = pd.to_numeric(g["交易中最大回撤%"], errors="coerce")
        rows.append({
            "進場方案": entry_name,
            "進場代碼": entry_key,
            "出場方案": exit_name,
            "出場代碼": exit_key,
            "年度": int(year),
            "樣本數": int(len(g)),
            "平均淨報酬%": round(float(ret.mean()), 3),
            "中位淨報酬%": round(float(ret.median()), 3),
            "勝率%": round(float((ret > 0).mean() * 100.0), 2),
            "中位最大回撤%": round(float(dd.median()), 3),
        })
    return pd.DataFrame(rows)


def summarize(trades: pd.DataFrame, annual: pd.DataFrame) -> pd.DataFrame:
    if trades.empty:
        return pd.DataFrame()
    rows = []
    keys = ["進場方案", "進場代碼", "進場類型", "出場方案", "出場代碼", "出場類型"]
    for vals, g in trades.groupby(keys, dropna=False):
        entry_name, entry_key, entry_kind, exit_name, exit_key, exit_kind = vals
        ret = pd.to_numeric(g["淨報酬_目標資金%"], errors="coerce")
        dep = pd.to_numeric(g["淨報酬_已投入資金%"], errors="coerce")
        mfe = pd.to_numeric(g["交易中MFE_淨值%"], errors="coerce")
        mae = pd.to_numeric(g["交易中MAE_淨值%"], errors="coerce")
        dd = pd.to_numeric(g["交易中最大回撤%"], errors="coerce")
        hold = pd.to_numeric(g["持有交易日"], errors="coerce")
        winners = ret[ret > 0]
        losers = ret[ret <= 0]
        payoff = (
            float(winners.mean()) / abs(float(losers.mean()))
            if len(winners) and len(losers) and abs(float(losers.mean())) > 1e-9
            else np.nan
        )

        a = annual[
            (annual["進場代碼"] == entry_key)
            & (annual["出場代碼"] == exit_key)
        ].copy()
        a = a[a["樣本數"] >= 30] if not a.empty else a
        if a.empty:
            annual_years = 0
            annual_positive = np.nan
            annual_median = np.nan
            annual_worst = np.nan
        else:
            vals2 = pd.to_numeric(a["中位淨報酬%"], errors="coerce")
            annual_years = int(len(a))
            annual_positive = float((vals2 > 0).mean() * 100.0)
            annual_median = float(vals2.median())
            annual_worst = float(vals2.min())

        rows.append({
            "進場方案": entry_name,
            "進場代碼": entry_key,
            "進場類型": entry_kind,
            "出場方案": exit_name,
            "出場代碼": exit_key,
            "出場類型": exit_kind,
            "樣本數": int(len(g)),
            "平均淨報酬%": round(float(ret.mean()), 3),
            "中位淨報酬%": round(float(ret.median()), 3),
            "勝率%": round(float((ret > 0).mean() * 100.0), 2),
            "平均已投入資金報酬%": round(float(dep.mean()), 3),
            "中位已投入資金報酬%": round(float(dep.median()), 3),
            "平均持有日": round(float(hold.mean()), 2),
            "中位持有日": round(float(hold.median()), 1),
            "中位MFE%": round(float(mfe.median()), 3),
            "中位MAE%": round(float(mae.median()), 3),
            "中位最大回撤%": round(float(dd.median()), 3),
            "獲利虧損比": round(payoff, 3) if np.isfinite(payoff) else np.nan,
            "加滿比例%": round(float(g["是否加滿"].astype(bool).mean() * 100.0), 2),
            "平均最大投入比例%": round(float(pd.to_numeric(g["最大投入比例%"], errors="coerce").mean()), 2),
            "平均買進次數": round(float(pd.to_numeric(g["買進交易次數"], errors="coerce").mean()), 2),
            "平均賣出次數": round(float(pd.to_numeric(g["賣出交易次數"], errors="coerce").mean()), 2),
            "結構風險退出率%": round(float((g["退出原因"] == "結構風險退出").mean() * 100.0), 2),
            "固定到期率%": round(float((g["退出原因"] == "固定持有到期").mean() * 100.0), 2),
            "年度有效年數": annual_years,
            "年度正中位報酬率%": round(annual_positive, 2) if np.isfinite(annual_positive) else np.nan,
            "跨年中位報酬%": round(annual_median, 3) if np.isfinite(annual_median) else np.nan,
            "最差年度中位報酬%": round(annual_worst, 3) if np.isfinite(annual_worst) else np.nan,
        })

    out = pd.DataFrame(rows)

    def pct(s, higher=True):
        x = pd.to_numeric(s, errors="coerce")
        if not higher:
            x = -x
        return x.rank(pct=True, method="average").fillna(0.0) * 100.0

    score = (
        0.25 * pct(out["中位淨報酬%"], True)
        + 0.15 * pct(out["平均淨報酬%"], True)
        + 0.15 * pct(out["勝率%"], True)
        + 0.10 * pct(out["獲利虧損比"], True)
        + 0.10 * pct(out["中位MAE%"], True)
        + 0.10 * pct(out["中位最大回撤%"], True)
        + 0.15 * pct(out["年度正中位報酬率%"], True)
    )
    n = pd.to_numeric(out["樣本數"], errors="coerce").fillna(0)
    reliability = np.where(n >= 3000, 1.0, np.where(n >= 1000, 0.98, np.where(n >= 300, 0.94, 0.88)))
    out["綜合分數"] = (score * reliability).round(2)
    out["總排名"] = out["綜合分數"].rank(ascending=False, method="min").astype(int)
    out["可靠度"] = np.select(
        [n >= 3000, n >= 1000, n >= 300],
        ["高", "中高", "中"],
        default="需擴充樣本",
    )
    return out.sort_values(["總排名", "樣本數"], ascending=[True, False]).reset_index(drop=True)


def strategy_rollups(summary: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    if summary.empty:
        return pd.DataFrame(), pd.DataFrame()

    entry_rows = []
    for (name, key, kind), g in summary.groupby(["進場方案", "進場代碼", "進場類型"]):
        entry_rows.append({
            "進場方案": name,
            "進場代碼": key,
            "進場類型": kind,
            "組合數": int(len(g)),
            "最佳總排名": int(g["總排名"].min()),
            "最佳綜合分數": round(float(g["綜合分數"].max()), 2),
            "各出場中位淨報酬中位數%": round(float(pd.to_numeric(g["中位淨報酬%"], errors="coerce").median()), 3),
            "各出場勝率中位數%": round(float(pd.to_numeric(g["勝率%"], errors="coerce").median()), 2),
        })
    entry_roll = pd.DataFrame(entry_rows).sort_values(["最佳總排名", "最佳綜合分數"], ascending=[True, False])

    exit_rows = []
    for (name, key, kind), g in summary.groupby(["出場方案", "出場代碼", "出場類型"]):
        exit_rows.append({
            "出場方案": name,
            "出場代碼": key,
            "出場類型": kind,
            "組合數": int(len(g)),
            "最佳總排名": int(g["總排名"].min()),
            "平均綜合分數": round(float(g["綜合分數"].mean()), 2),
            "各進場中位淨報酬中位數%": round(float(pd.to_numeric(g["中位淨報酬%"], errors="coerce").median()), 3),
            "各進場勝率中位數%": round(float(pd.to_numeric(g["勝率%"], errors="coerce").median()), 2),
        })
    exit_roll = pd.DataFrame(exit_rows).sort_values(["平均綜合分數", "最佳總排名"], ascending=[False, True])
    return entry_roll, exit_roll


def write_outputs(
    summary: pd.DataFrame,
    annual: pd.DataFrame,
    trades: pd.DataFrame,
    entry_roll: pd.DataFrame,
    exit_roll: pd.DataFrame,
    status: dict,
):
    summary.to_csv(OUT_SUMMARY, index=False, encoding="utf-8-sig")
    annual.to_csv(OUT_ANNUAL, index=False, encoding="utf-8-sig")
    OUT_STATUS.write_text(json.dumps(status, ensure_ascii=False, indent=2), encoding="utf-8")

    top_pairs = summary.head(10)[["進場代碼", "出場代碼"]].drop_duplicates()
    detail = trades.merge(top_pairs, on=["進場代碼", "出場代碼"], how="inner")
    if len(detail) > 300000:
        detail = detail.head(300000).copy()

    notes = pd.DataFrame({
        "說明": [
            "所有進場/加碼/減碼/退出訊號都在 T 日收盤後確認，統一於 T+1 開盤執行。",
            "同一股票同一進場方案採 cooldown，避免短期重複事件灌水。",
            "分批進場以『目標資金=100%』衡量；若只有50%試單而未加滿，目標資金報酬會保留未投入現金的影響。",
            "分層獲利管理：一般股過熱70先減1/3、80再減半剩餘；若RS曾達85，門檻放寬為85與90；過熱後RS自高點回落5時退出剩餘。",
            "結構風險退出：收盤跌破MA20且MA20轉下，或右側分數跌破45；這是保護失敗交易，不代表最佳獲利賣點。",
            "交易成本預設可由 workflow 調整。股票賣出另計證券交易稅；手續費每次買/賣皆計。",
            "本研究是事件交易統計，不是資金受限的完整投資組合回測；不同股票的交易可以同時重疊。",
            "歷史基本面/估值沒有逐日 point-in-time 快照，仍未納入，以避免未來資訊偏誤。",
        ]
    })

    with pd.ExcelWriter(OUT_XLSX, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="完整組合排行榜", index=False)
        entry_roll.to_excel(writer, sheet_name="進場方案總覽", index=False)
        exit_roll.to_excel(writer, sheet_name="出場方案總覽", index=False)
        annual.to_excel(writer, sheet_name="年度驗證", index=False)
        detail.to_excel(writer, sheet_name="前10名交易明細", index=False)
        notes.to_excel(writer, sheet_name="說明", index=False)

        wb = writer.book
        head = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        for sheet, df in [
            ("完整組合排行榜", summary),
            ("進場方案總覽", entry_roll),
            ("出場方案總覽", exit_roll),
            ("年度驗證", annual),
            ("前10名交易明細", detail),
            ("說明", notes),
        ]:
            ws = writer.sheets[sheet]
            ws.freeze_panes(1, 0)
            if len(df.columns):
                ws.autofilter(0, 0, max(0, len(df)), len(df.columns) - 1)
            for j, col in enumerate(df.columns):
                ws.write(0, j, col, head)
                width = min(38, max(11, len(str(col)) + 2))
                if "方案" in str(col) or col in {"名稱", "退出原因"}:
                    width = max(width, 24)
                if sheet == "說明":
                    width = 100
                ws.set_column(j, j, width)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", default="2023-01-03")
    ap.add_argument("--end", default="")
    ap.add_argument("--max-hold", type=int, default=60)
    ap.add_argument("--cooldown", type=int, default=10)
    ap.add_argument("--workers", type=int, default=2)
    ap.add_argument("--fetch-missing", action="store_true")
    ap.add_argument("--commission-rate", type=float, default=0.001425)
    ap.add_argument("--sell-tax-rate", type=float, default=0.003)
    args = ap.parse_args()

    start = entry_bt._date(args.start)
    requested_end = entry_bt._date(args.end) if args.end else entry_bt._latest_formal_date()
    warmup = start - dt.timedelta(days=entry_bt.WARMUP_CALENDAR_DAYS)
    max_hold = max(20, min(80, int(args.max_hold)))
    cooldown = max(0, min(30, int(args.cooldown)))
    commission = max(0.0, float(args.commission_rate))
    sell_tax = max(0.0, float(args.sell_tax_rate))

    print("=== 進場 × 出場完整組合回測 ===", flush=True)
    print(f"requested: {start} ~ {requested_end}", flush=True)
    print(f"warmup: {warmup}", flush=True)
    print(f"max_hold: {max_hold} sessions | cooldown: {cooldown}", flush=True)
    print(f"commission per side: {commission:.6f} | sell tax: {sell_tax:.6f}", flush=True)

    cache_report = entry_bt.ensure_history_cache(
        warmup, requested_end, args.fetch_missing, args.workers
    )
    raw, meta = entry_bt.load_market_cache(warmup, requested_end)
    effective_end = min(requested_end, pd.to_datetime(meta["last_date"]).date())
    signals, market = entry_bt.compute_signals(raw)

    trades, sim_meta = simulate_all(
        signals, start, effective_end, max_hold, cooldown, commission, sell_tax
    )
    if trades.empty:
        raise RuntimeError("沒有可用的進出場組合交易")

    annual = annual_summary(trades)
    summary = summarize(trades, annual)
    entry_roll, exit_roll = strategy_rollups(summary)

    leaders = []
    for _, r in summary.head(10).iterrows():
        leaders.append({
            "排名": int(r["總排名"]),
            "進場方案": str(r["進場方案"]),
            "出場方案": str(r["出場方案"]),
            "綜合分數": float(r["綜合分數"]),
            "樣本數": int(r["樣本數"]),
            "中位淨報酬%": _num(r.get("中位淨報酬%")),
            "平均淨報酬%": _num(r.get("平均淨報酬%")),
            "勝率%": _num(r.get("勝率%")),
            "中位持有日": _num(r.get("中位持有日")),
            "中位最大回撤%": _num(r.get("中位最大回撤%")),
            "年度正中位報酬率%": _num(r.get("年度正中位報酬率%")),
            "可靠度": str(r.get("可靠度")),
        })

    status = {
        "status": "success",
        "generated_at_taipei": dt.datetime.now(
            dt.timezone(dt.timedelta(hours=8))
        ).strftime("%Y-%m-%d %H:%M:%S"),
        "research_start": start.isoformat(),
        "research_end_requested": requested_end.isoformat(),
        "research_end_effective": effective_end.isoformat(),
        "warmup_start": warmup.isoformat(),
        "max_hold_sessions": max_hold,
        "cooldown_sessions": cooldown,
        "commission_rate_per_side": commission,
        "sell_tax_rate": sell_tax,
        "cache": cache_report,
        "data": meta,
        "simulation": sim_meta,
        "summary_rows": int(len(summary)),
        "annual_rows": int(len(annual)),
        "method": "T close signal -> T+1 open execution; point-in-time technical/RS/breadth; configurable trading costs",
        "portfolio_note": "Event-trade statistics; different-stock trades can overlap; not a capital-constrained portfolio backtest.",
        "fundamental_history_used": False,
        "current_leaders": leaders,
        "outputs": [
            OUT_XLSX.name,
            OUT_SUMMARY.name,
            OUT_ANNUAL.name,
            OUT_STATUS.name,
        ],
    }
    write_outputs(summary, annual, trades, entry_roll, exit_roll, status)

    cols = [
        "總排名", "進場方案", "出場方案", "樣本數", "中位淨報酬%",
        "平均淨報酬%", "勝率%", "中位持有日", "中位最大回撤%",
        "年度正中位報酬率%", "綜合分數", "可靠度"
    ]
    print("[done] top combos:", flush=True)
    print(summary[cols].head(15).to_string(index=False), flush=True)


if __name__ == "__main__":
    main()
