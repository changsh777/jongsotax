"""
접수명단 시트 컬럼 재배치
- 맨 끝에 있는 이자/배당/근로(단일)/근로(복수)/연금/기타 삭제
- 타소득여부 바로 다음(J열)에 6개 컬럼 삽입 후 데이터 재기록
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

import openpyxl
from gspread.utils import rowcol_to_a1
from config import PARSE_RESULT_XLSX
from gsheet_writer import get_credentials
import gspread

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
타소득_COLS = ["이자", "배당", "근로(단일)", "근로(복수)", "연금", "기타"]

def col_letter(idx_1based):
    return rowcol_to_a1(1, idx_1based).rstrip("0123456789")

def main():
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet("접수명단")
    sheet_id = ws.id

    headers = ws.row_values(1)
    print(f"현재 헤더 수: {len(headers)}")

    # 1) 맨 끝에 있는 이자~기타 컬럼 위치 확인
    end_cols = {}
    for col in 타소득_COLS:
        if col in headers:
            end_cols[col] = headers.index(col)  # 0-indexed

    if end_cols:
        print(f"삭제 대상: {list(end_cols.items())}")
        # 뒤에서부터 삭제 (인덱스 밀리지 않게)
        for col in reversed(타소득_COLS):
            if col not in end_cols:
                continue
            idx0 = end_cols[col]
            sh.batch_update({"requests": [{
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": idx0,
                        "endIndex": idx0 + 1
                    }
                }
            }]})
            print(f"  삭제: '{col}' (열 {idx0+1})")
        print("기존 컬럼 삭제 완료")
    else:
        print("삭제할 컬럼 없음 (처음 실행)")

    # 2) 헤더 재조회
    headers = ws.row_values(1)

    # 3) 타소득여부 위치 확인
    if "타소득여부" not in headers:
        print("타소득여부 컬럼 없음 - 종료")
        return
    타소득_idx0 = headers.index("타소득여부")  # 0-indexed
    insert_at = 타소득_idx0 + 1               # 타소득여부 다음에 삽입 (0-indexed)
    print(f"타소득여부: 열 {타소득_idx0+1}, 삽입 위치: 열 {insert_at+1}")

    # 4) 6개 컬럼 한번에 삽입
    sh.batch_update({"requests": [{
        "insertDimension": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "COLUMNS",
                "startIndex": insert_at,
                "endIndex": insert_at + 6
            },
            "inheritFromBefore": False
        }
    }]})
    print(f"6개 컬럼 삽입 완료 (열 {insert_at+1}~{insert_at+6})")

    # 5) 헤더 쓰기
    for i, col_name in enumerate(타소득_COLS):
        ws.update_cell(1, insert_at + 1 + i, col_name)
    print(f"헤더 기록: {타소득_COLS}")

    # 6) 파싱결과.xlsx에서 데이터 읽기
    wb = openpyxl.load_workbook(PARSE_RESULT_XLSX, data_only=True)
    ws_xlsx = wb.active
    parse_headers = [c.value for c in ws_xlsx[1]]
    parsed = {}
    for row in ws_xlsx.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        name = str(row[0]).strip()
        entry = {}
        for i, h in enumerate(parse_headers):
            if h in 타소득_COLS:
                entry[h] = row[i] if row[i] is not None else ""
        parsed[name] = entry
    print(f"파싱결과.xlsx {len(parsed)}건 로드")

    # 7) 접수명단 성명→행번호
    all_vals = ws.get_all_values()
    headers2 = ws.row_values(1)
    name_col = headers2.index("성명") + 1  # 1-indexed
    name_to_row = {}
    for i, row in enumerate(all_vals[1:], start=2):
        if len(row) >= name_col:
            n = row[name_col - 1].strip()
            if n:
                name_to_row[n] = i

    # 8) 배치 업데이트
    updates = []
    for name, entry in parsed.items():
        row_idx = name_to_row.get(name)
        if not row_idx:
            continue
        for i, col_name in enumerate(타소득_COLS):
            val = entry.get(col_name, "")
            cl = col_letter(insert_at + 1 + i)
            updates.append({"range": f"{cl}{row_idx}", "values": [[val]]})

    if updates:
        ws.batch_update(updates)
        print(f"데이터 업데이트 {len(parsed)}건 완료")
    else:
        print("업데이트 없음")

if __name__ == "__main__":
    main()
