# -*- coding: utf-8 -*-
from pathlib import Path
import hashlib, json, zipfile
ROOT=Path(__file__).resolve().parents[2]
HERE=Path(__file__).resolve().parent
manifest=json.loads((HERE/'manifest.json').read_text(encoding='utf-8'))
parts=[HERE/x for x in manifest['parts']]
for p in parts:
    if not p.exists(): raise SystemExit(f'缺少 seed 分割檔：{p.name}')
tmp=HERE/'_bootstrap_seed.tmp.zip'
h=hashlib.sha256()
with tmp.open('wb') as out:
    for p in parts:
        data=p.read_bytes(); out.write(data); h.update(data)
if h.hexdigest()!=manifest['sha256']: tmp.unlink(missing_ok=True); raise SystemExit('Seed SHA256 驗證失敗，請重新上傳 seed parts。')
with zipfile.ZipFile(tmp,'r') as z: z.extractall(ROOT)
tmp.unlink(missing_ok=True)
print('[OK] 初始歷史資料已還原。')
