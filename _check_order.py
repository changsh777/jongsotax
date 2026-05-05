# -*- coding: utf-8 -*-
import sys, json, urllib.request, urllib.parse, time
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

PAT_FILE = Path(__file__).parent / '.credentials' / 'airtable_pat.txt'
PAT = PAT_FILE.read_text(encoding='utf-8').strip()

BASE_ID  = 'appSvDTDOmYfBeIFs'
TABLE_ID = 'tbl2f2h6GfSnLCQpt'

def get(path):
    url = f'https://api.airtable.com/v0/{path}'
    req = urllib.request.Request(url, headers={'Authorization': f'Bearer {PAT}'})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode())

# 뷰 ID
meta = get(f'meta/bases/{BASE_ID}/tables')
view_id = None
for t in meta.get('tables', []):
    if t['id'] == TABLE_ID:
        views = t.get('views', [])
        if views:
            view_id = views[0]['id']
            print(f'뷰: {views[0]["name"]} / {view_id}')
        break

records, offset = [], None
while True:
    params = [f'view={urllib.parse.quote(view_id)}'] if view_id else []
    if offset:
        params.append(f'offset={urllib.parse.quote(offset)}')
    qs = ('?' + '&'.join(params)) if params else ''
    d = get(f'{BASE_ID}/{TABLE_ID}{qs}')
    records.extend(d.get('records', []))
    offset = d.get('offset')
    if not offset:
        break
    time.sleep(0.2)

print(f'에어테이블 총 {len(records)}건')
for i, r in enumerate(records, 1):
    print(f'{i}:{r["fields"].get("성명", "")}')
