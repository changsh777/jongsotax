"""
접수명단 장부유형(16열) 바로 다음에 추계시적용경비율 컬럼 삽입 + 파싱값 기록
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
COL_NAME = "추계시적용경비율"

def main():
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet("접수명단")
    sheet_id = ws.id

    headers = ws.row_values(1)

    # 이미 있으면 위치만 확인
    if COL_NAME in headers:
        col_idx = headers.index(COL_NAME) + 1
        print(f"이미 존재: '{COL_NAME}' (열 {col_idx})")
    else:
        # 장부유형 다음에 삽입
        insert_at = headers.index("장부유형") + 1  # 0-indexed (장부유형 다음)
        sh.batch_update({"requests": [{
            "insertDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": insert_at,
                    "endIndex": insert_at + 1
                },
                "inheritFromBefore": False
            }
        }]})
        col_idx = insert_at + 1  # 1-indexed
        ws.update_cell(1, col_idx, COL_NAME)
        print(f"컬럼 삽입: '{COL_NAME}' (열 {col_idx})")

    # 파싱결과.xlsx에서 값 읽기
    wb = openpyxl.load_workbook(PARSE_RESULT_XLSX, data_only=True)
    ws_xlsx = wb.active
    parse_headers = [c.value for c in ws_xlsx[1]]
    parsed = {}
    for row in ws_xlsx.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]: continue
        name = str(row[0]).strip()
        val = row[parse_headers.index(COL_NAME)] if COL_NAME in parse_headers else ""
        parsed[name] = val if val is not None else ""
    print(f"파싱결과.xlsx {len(parsed)}건 로드")

    # 성명 → 행번호
    all_vals = ws.get_all_values()
    headers2 = ws.row_values(1)
    name_col = headers2.index("성명") + 1
    name_to_row = {}
    for i, row in enumerate(all_vals[1:], start=2):
        if len(row) >= name_col:
            n = row[name_col - 1].strip()
            if n:
                name_to_row[n] = i

    # 배치 업데이트
    cl = rowcol_to_a1(1, col_idx).rstrip("0123456789")
    updates = []
    for name, val in parsed.items():
        row_idx = name_to_row.get(name)
        if row_idx:
            updates.append({"range": f"{cl}{row_idx}", "values": [[val]]})

    if updates:
        ws.batch_update(updates)
        print(f"{len(updates)}건 업데이트 완료")

if __name__ == "__main__":
    main()
