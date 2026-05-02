"""
airtable_sync_mac.py - 에어테이블 → 구글시트 단방향 자동 동기화 (Mac Mini 전용)

- 1분 크론으로 실행
- 2026-06-01부터 자동 종료 (종소세 시즌 끝)
- 에어테이블 = 마스터 / 구글시트 = 작업 사본 (덮어쓰기)
"""
import sys, json, time, urllib.request, pickle
from datetime import datetime, date
from pathlib import Path

import gspread
from google.auth.transport.requests import Request

# ── 설정 ─────────────────────────────────────────────
SEASON_END     = date(2026, 6, 1)          # 이 날부터 자동 중단
CRED_DIR       = Path.home() / "종소세2026/.credentials"
TOKEN_FILE     = CRED_DIR / "token.pickle"
AIRTABLE_PAT   = open(Path.home() / "종소세2026/.credentials/airtable_pat.txt").read().strip()
BASE_ID        = "appSvDTDOmYfBeIFs"
TABLE_ID       = "tbl2f2h6GfSnLCQpt"
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME     = "접수명단"
SKIP_TYPES     = {"multipleRecordLinks", "multipleAttachments", "multipleLookupValues"}


# ── 인증 ─────────────────────────────────────────────
def get_creds():
    with open(TOKEN_FILE, "rb") as f:
        creds = pickle.load(f)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return creds


# ── 에어테이블 ────────────────────────────────────────
def airtable_get(path):
    url = f"https://api.airtable.com/v0/{path}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {AIRTABLE_PAT}"})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode())

def fetch_fields():
    data = airtable_get(f"meta/bases/{BASE_ID}/tables")
    for t in data.get("tables", []):
        if t["id"] == TABLE_ID:
            return [(f["name"], f["type"]) for f in t.get("fields", [])]
    return []

def fetch_all_records():
    records, offset = [], None
    while True:
        path = f"{BASE_ID}/{TABLE_ID}" + (f"?offset={offset}" if offset else "")
        data = airtable_get(path)
        records.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset:
            break
        time.sleep(0.2)
    return records

def cell_value(val, ftype):
    if val is None: return ""
    if ftype in SKIP_TYPES: return f"[{len(val)}건]" if isinstance(val, list) else str(val)
    if isinstance(val, bool): return "O" if val else ""
    if isinstance(val, (int, float)): return val
    if isinstance(val, dict): return val.get("name") or val.get("email") or str(val)
    return str(val)


# ── 메인 ─────────────────────────────────────────────
def main():
    # 시즌 종료 체크
    if date.today() >= SEASON_END:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 종소세 시즌 종료({SEASON_END}) - 싱크 중단")
        sys.exit(0)

    fields  = fetch_fields()
    records = fetch_all_records()

    normal = [(n, t) for n, t in fields if t not in SKIP_TYPES]
    linked = [(n, t) for n, t in fields if t in SKIP_TYPES]
    header = [n for n, _ in normal] + ["[링크]" + n for n, _ in linked] + ["_sync_at"]

    now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [header] + [
        [cell_value(r["fields"].get(n), t) for n, t in normal + linked] + [now]
        for r in records
    ]

    ws = gspread.authorize(get_creds()).open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    ws.clear()
    ws.update(range_name="A1", values=rows, value_input_option="USER_ENTERED")
    print(f"[{now}] 동기화 완료: {len(records)}건")


if __name__ == "__main__":
    main()
