# -*- coding: utf-8 -*-
from __future__ import annotations
import argparse, datetime as dt, json, os, shutil, subprocess, sys, time
from pathlib import Path
from zoneinfo import ZoneInfo
import pandas as pd

ROOT=Path(__file__).resolve().parents[1]
RS_DIR=ROOT/'automation'/'rs'; DUCK_DIR=ROOT/'automation'/'duckbill'
STATUS=ROOT/'update_status.json'

def run(cmd,cwd,env=None):
    print('[RUN]', ' '.join(map(str,cmd)), flush=True)
    p=subprocess.run(cmd,cwd=cwd,env=env)
    print('[EXIT]',p.returncode,flush=True)
    return p.returncode

def norm_date(series):
    if pd.api.types.is_numeric_dtype(series):
        x=pd.to_numeric(series,errors='coerce'); med=x.dropna().median() if x.notna().any() else None
        if med is not None and 25000<med<80000: return pd.to_datetime(x,unit='D',origin='1899-12-30',errors='coerce')
    return pd.to_datetime(series,errors='coerce')

def excel_latest(path,sheet,header=0):
    try:
        df=pd.read_excel(path,sheet_name=sheet,header=header)
        if '日期' not in df.columns or df.empty: return None
        d=norm_date(df['日期']); return d.max().date() if d.notna().any() else None
    except Exception as e:
        print('[WARN] read date failed',path,e,flush=True); return None

def latest_file(folder,pattern):
    fs=sorted(folder.glob(pattern),key=lambda p:p.stat().st_mtime,reverse=True)
    return fs[0] if fs else None

def write_status(status,target,rs_date,duck_date,message,attempts):
    now=dt.datetime.now(ZoneInfo('Asia/Taipei')).strftime('%Y-%m-%d %H:%M:%S')
    data={'status':status,'target_date':target.isoformat(),'rs_date':rs_date.isoformat() if rs_date else None,'duck_date':duck_date.isoformat() if duck_date else None,'last_run_taipei':now,'message':message,'attempts':attempts,'schedule':'週一至週五 18:05（Asia/Taipei）'}
    STATUS.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--date'); ap.add_argument('--attempts',type=int,default=3); ap.add_argument('--retry-minutes',type=int,default=8); args=ap.parse_args()
    if args.date: target=dt.datetime.strptime(args.date,'%Y-%m-%d').date()
    else: target=dt.datetime.now(ZoneInfo('Asia/Taipei')).date()
    while target.weekday()>=5: target-=dt.timedelta(days=1)
    existing_rs=excel_latest(ROOT/'rs_latest.xlsx','每日強勢股數量',0)
    existing_duck=excel_latest(ROOT/'duck_latest.xlsx','全部符合',1)
    rs_date=existing_rs; duck_date=existing_duck; rs_out=None; duck_out=None; errors=[]
    used_attempts=0
    for attempt in range(1,max(1,args.attempts)+1):
        used_attempts=attempt; print(f'=== Attempt {attempt}/{args.attempts} target={target} ===',flush=True)
        if rs_date!=target:
            env=os.environ.copy(); env['RS_TARGET_DATE']=target.isoformat(); env['TZ']='Asia/Taipei'
            rc=run([sys.executable,'-u','rs_breadth.py','--date',target.isoformat()],RS_DIR,env)
            if rc!=0: errors.append(f'RS exit={rc}')
            rs_out=latest_file(RS_DIR/'output','*_RS強勢股市場廣度_極值趨勢.xlsx')
            if rs_out: rs_date=excel_latest(rs_out,'每日強勢股數量',0)
        if duck_date!=target:
            env=os.environ.copy(); env['TZ']='Asia/Taipei'
            rc=run([sys.executable,'-u','stock_duckbill.py','--update','--date',target.isoformat()],DUCK_DIR,env)
            if rc!=0: errors.append(f'Duck exit={rc}')
            duck_out=DUCK_DIR/'output'/f'{target.strftime("%Y%m%d")}_鴨嘴篩選結果.xlsx'
            if duck_out.exists(): duck_date=excel_latest(duck_out,'全部符合',1)
        if rs_date==target and duck_date==target: break
        if attempt<args.attempts:
            print(f'[WAIT] New trading data not complete. Sleep {args.retry_minutes} minutes.',flush=True)
            time.sleep(max(0,args.retry_minutes)*60)
    # copy only results that are not older than current published files
    if rs_out and rs_out.exists() and rs_date and (not existing_rs or rs_date>=existing_rs): shutil.copy2(rs_out,ROOT/'rs_latest.xlsx')
    if duck_out and duck_out.exists() and duck_date and (not existing_duck or duck_date>=existing_duck): shutil.copy2(duck_out,ROOT/'duck_latest.xlsx')
    if rs_date==target and duck_date==target:
        status='success'; msg=f'{target} 兩套系統皆已更新完成。'
    elif (rs_date and rs_date>=target) or (duck_date and duck_date>=target):
        status='partial'; msg=f'部分更新完成：RS={rs_date}、鴨嘴={duck_date}。保留各系統最後成功結果。'
    elif errors:
        status='error'; msg=f'{target} 重試後仍未取得完整新資料；可能為官方資料延遲、休市或來源錯誤。'+'；'.join(errors[-4:])
    else:
        status='no_new_data'; msg=f'{target} 重試後沒有新交易日資料，可能休市或官方盤後資料尚未發布；目前保留 RS={rs_date}、鴨嘴={duck_date}。'
    write_status(status,target,rs_date,duck_date,msg,used_attempts)
    print('[STATUS]',msg,flush=True)
    return 0

if __name__=='__main__': raise SystemExit(main())
