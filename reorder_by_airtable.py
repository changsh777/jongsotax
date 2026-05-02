"""
에어테이블 레코드 순서대로 구글시트 접수명단 행 재정렬
- 파싱값(수입/할인가/수수료 등) 컬럼 포함 전체 행 순서만 바꿈
- 에어테이블에 없는 행은 맨 아래로
"""
import sys, io, os, json, urllib.request, urllib.parse, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"F:\종소세2026")

from config_secret import AIRTABLE_PAT
from gsheet_writer import get_credentials
import gspread

BASE_ID        = "appSvDTDOmYfBeIFs"
TABLE_ID       = "tbl2f2h6GfSnLCQpt"
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"


def airtable_get(path):
    url = f"https://api.airtable.com/v0/{path}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {AIRTABLE_PAT}"})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode("utf-8"))


def fetch_records():
    """에어테이블 레코드 순서대로 (성명, fields dict) 리스트 반환"""
    records = []
    offset = None
    page = 1
    while True:
        params = {"view": "viwEZG7bdAmIC32cu"}  # 총괄화면
        if offset:
            params["offset"] = offset
        qs = urllib.parse.urlencode(params)
        path = f"{BASE_ID}/{TABLE_ID}?{qs}"
        data = airtable_get(path)
        for rec in data.get("records", []):
            fields = rec.get("fields", {})
            name = fields.get("성명", "").strip()
            if name:
                records.append((name, fields))
        print(f"  페이지 {page}: {len(data.get('records', []))}건")
        offset = data.get("offset")
        if not offset:
            break
        page += 1
        time.sleep(0.2)
    return records


def main():
    # 1) 에어테이블 전체 레코드
    print("[에어테이블] 레코드 가져오는 중...")
    at_records = fetch_records()
    at_order = [name for name, _ in at_records]
    print(f"  총 {len(at_records)}건: {at_order[:5]}...")

    # 2) 구글시트 현재 데이터 전체 읽기
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet("접수명단")

    all_vals = ws.get_all_values()
    header = all_vals[0]
    data_rows = all_vals[1:]
    print(f"[구글시트] {len(data_rows)}행 읽기 완료")

    name_col = header.index("성명")
    jumin_col = header.index("주민번호") if "주민번호" in header else None
    phone_col = header.index("핸드폰번호") if "핸드폰번호" in header else None

    # 성명 → 행 데이터 dict
    row_by_name = {}
    for row in data_rows:
        if len(row) > name_col:
            name = row[name_col].strip()
            if name:
                row_by_name[name] = row

    # Airtable → 구글시트 컬럼 직접 매핑 (신규 행에 채워넣을 필드)
    AT_COL_MAP = {
        "기존신규":   "기존신규",
        "주민번호":   "주민번호",
        "핸드폰번호": "핸드폰번호",
        "자동회신":   "자동회신",
        "입금체크":   "입금체크",
    }

    # 3) 에어테이블 순서대로 재배열 + 신규 행 추가
    sorted_rows = []
    matched = []
    new_added = []

    for name, fields in at_records:
        if name in row_by_name:
            # 기존 행: header 길이에 맞게 패딩 (끝 빈칸 복원)
            row = row_by_name[name]
            if len(row) < len(header):
                row = row + [""] * (len(header) - len(row))
            sorted_rows.append(row)
            matched.append(name)
        else:
            # 신규: 에어테이블 기본 정보로 빈 행 생성
            new_row = [""] * len(header)
            new_row[name_col] = name
            for sheet_col, at_field in AT_COL_MAP.items():
                if sheet_col in header:
                    val = fields.get(at_field, "")
                    if isinstance(val, bool):
                        val = "O" if val else ""
                    new_row[header.index(sheet_col)] = str(val) if val else ""
            sorted_rows.append(new_row)
            new_added.append(name)

    at_names = set(at_order)
    removed = [row[name_col].strip() for row in data_rows
               if row[name_col].strip() and row[name_col].strip() not in at_names]

    if new_added:
        print(f"  신규 추가: {new_added}")
    if removed:
        print(f"  제외(삭제됨): {removed}")
    print(f"  매칭 {len(matched)}건 / 신규 {len(new_added)}건 / 제외 {len(removed)}건")

    # 4) 구글시트 덮어쓰기 (헤더 제외)
    from gspread.utils import rowcol_to_a1
    end_col = rowcol_to_a1(1, len(header)).rstrip("0123456789")
    end_row = len(sorted_rows) + 1

    ws.batch_clear([f"A2:{end_col}{end_row + 10}"])
    if sorted_rows:
        ws.update(f"A2:{end_col}{end_row}", sorted_rows, value_input_option="USER_ENTERED")

    print(f"[완료] 접수명단 {len(sorted_rows)}행 에어테이블 순서로 재정렬")
    print(f"  앞 5명: {[r[name_col] for r in sorted_rows[:5]]}")


if __name__ == "__main__":
    main()
