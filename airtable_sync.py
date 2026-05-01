"""
airtable_sync.py - 에어테이블 → 구글시트 단방향 동기화

원칙:
- 에어테이블 = 마스터 DB (읽기만)
- 구글시트 '접수명단' = 작업 사본 (봇이 자유롭게 읽기/쓰기)
- 이 스크립트는 에어테이블 → 구글시트 방향만 (역방향 없음)
- 실행할 때마다 전체 덮어쓰기 (최신 상태 유지)

실행:
  python airtable_sync.py          # 전체 동기화
  python airtable_sync.py --dry    # 에어테이블 조회만 (시트 안 씀)
"""
import sys
import argparse
import urllib.request
import json
import time
from datetime import datetime

sys.path.insert(0, r"F:\종소세2026")
from gsheet_writer import get_credentials
import gspread

# ===== 설정 =====
AIRTABLE_PAT  = os.environ.get("AIRTABLE_PAT", "")
BASE_ID       = "appSvDTDOmYfBeIFs"
TABLE_ID      = "tbl2f2h6GfSnLCQpt"   # 종소세2026
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME    = "접수명단"

# 복잡한 타입 → 단순 텍스트 변환 대상 (링크·첨부는 건너뜀)
SKIP_TYPES = {"multipleRecordLinks", "multipleAttachments", "multipleLookupValues"}

# ===== 에어테이블 조회 =====
def airtable_get(path):
    url = f"https://api.airtable.com/v0/{path}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {AIRTABLE_PAT}"})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode("utf-8"))


def fetch_all_records():
    """페이지네이션 처리해서 전체 레코드 가져오기"""
    records = []
    offset = None
    page = 1
    while True:
        url = f"meta/bases/{BASE_ID}/tables"   # 먼저 필드 순서 확인용
        path = f"{BASE_ID}/{TABLE_ID}"
        if offset:
            path += f"?offset={offset}"
        data = airtable_get(path)
        batch = data.get("records", [])
        records.extend(batch)
        print(f"  페이지 {page}: {len(batch)}건 (누적 {len(records)}건)")
        offset = data.get("offset")
        if not offset:
            break
        page += 1
        time.sleep(0.2)   # API rate limit 여유
    return records


def fetch_field_order():
    """테이블 스키마에서 필드 이름·타입 순서대로 가져오기"""
    data = airtable_get(f"meta/bases/{BASE_ID}/tables")
    for t in data.get("tables", []):
        if t["id"] == TABLE_ID:
            return [(f["name"], f["type"]) for f in t.get("fields", [])]
    return []


# ===== 값 변환 =====
def cell_value(val, ftype):
    """에어테이블 필드값 → 구글시트 셀값"""
    if val is None:
        return ""
    if ftype in SKIP_TYPES:
        # 링크/첨부: 건수만 표시
        if isinstance(val, list):
            return f"[{len(val)}건]"
        return str(val)
    if isinstance(val, bool):
        return "O" if val else ""
    if isinstance(val, (int, float)):
        return val
    if isinstance(val, dict):
        # createdBy 등
        return val.get("name") or val.get("email") or str(val)
    return str(val)


# ===== 구글시트 쓰기 =====
def get_or_create_sheet(gc, sh):
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=500, cols=60)
        print(f"  '{SHEET_NAME}' 시트 새로 생성")
    return ws


def write_to_sheet(fields, records):
    """전체 덮어쓰기"""
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = get_or_create_sheet(gc, sh)

    # 헤더: 에어테이블 필드명 그대로 + 동기화일시 추가
    header = [name for name, ftype in fields if ftype not in SKIP_TYPES]
    header += ["[링크필드]" + name for name, ftype in fields if ftype in SKIP_TYPES]
    header.append("_sync_at")

    # 데이터 행
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [header]
    # 스킵 필드 / 일반 필드 순서 분리
    normal_fields = [(n, t) for n, t in fields if t not in SKIP_TYPES]
    link_fields   = [(n, t) for n, t in fields if t in SKIP_TYPES]
    ordered_fields = normal_fields + link_fields

    for rec in records:
        f = rec.get("fields", {})
        row = [cell_value(f.get(name), ftype) for name, ftype in ordered_fields]
        row.append(now)
        rows.append(row)

    # 기존 내용 전체 클리어 후 덮어쓰기
    ws.clear()
    ws.update(range_name="A1", values=rows, value_input_option="USER_ENTERED")
    print(f"  구글시트 '{SHEET_NAME}' 업데이트: {len(records)}행 × {len(header)}열")

    # 헤더 볼드 처리
    try:
        ws.format("1:1", {"textFormat": {"bold": True}})
    except Exception:
        pass


# ===== 메인 =====
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry", action="store_true", help="에어테이블만 조회 (시트 안 씀)")
    args = parser.parse_args()

    print(f"\n[에어테이블 → 구글시트 동기화]")
    print(f"  베이스: {BASE_ID} / 테이블: {TABLE_ID}")
    print(f"  대상 시트: {SHEET_NAME}\n")

    print("[1] 필드 구조 조회")
    fields = fetch_field_order()
    print(f"  총 {len(fields)}개 필드")

    print("[2] 레코드 전체 조회")
    records = fetch_all_records()
    print(f"  총 {len(records)}건")

    if args.dry:
        print("\n[DRY] 시트 쓰기 생략")
        print("샘플 (첫 2건):")
        for rec in records[:2]:
            print(" ", {k: v for k, v in list(rec.get("fields", {}).items())[:5]})
        return

    print("[3] 구글시트 쓰기")
    write_to_sheet(fields, records)

    print("\n[완료]")
    print(f"  https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")


if __name__ == "__main__":
    main()
