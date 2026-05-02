"""
파싱결과.xlsx → 접수명단 시트 upsert
- 기존 컬럼: 수입, 할인가, 수수료
- 추가 컬럼: 장부유형(기장의무), 이자, 배당, 근로(단일), 근로(복수), 연금, 기타
- 없는 컬럼은 시트에 자동 추가
- 나머지 접수명단 컬럼은 절대 건드리지 않음
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

# 접수명단 컬럼명 → 파싱결과.xlsx 컬럼명
COL_MAP = {
    "수입":           "수입금액총계",
    "장부유형":       "기장의무",
    "추계시적용경비율": "추계시적용경비율",
    "할인가":         "사전접수할인가",
    "수수료":         "일반접수가",
    "이자":           "이자",
    "배당":           "배당",
    "근로(단일)":     "근로(단일)",
    "근로(복수)":     "근로(복수)",
    "연금":           "연금",
    "기타":           "기타",
}

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"


def ensure_columns(ws, needed_cols):
    """없는 컬럼을 시트 오른쪽에 추가하고, 컬럼명→인덱스(1-based) 반환"""
    headers = ws.row_values(1)
    col_idx = {}

    for col in needed_cols:
        if col in headers:
            col_idx[col] = headers.index(col) + 1
        else:
            # 컬럼 추가
            new_idx = len(headers) + 1
            if ws.col_count < new_idx:
                ws.resize(rows=ws.row_count, cols=new_idx)
            ws.update_cell(1, new_idx, col)
            headers.append(col)
            col_idx[col] = new_idx
            print(f"  + 컬럼 추가: '{col}' (열 {new_idx})")

    return col_idx


def main():
    # 1) 파싱결과.xlsx 로드
    if not PARSE_RESULT_XLSX.exists():
        print(f"파싱결과.xlsx 없음: {PARSE_RESULT_XLSX}")
        return

    wb = openpyxl.load_workbook(PARSE_RESULT_XLSX, data_only=True)
    ws_xlsx = wb.active
    parse_headers = [c.value for c in ws_xlsx[1]]

    parsed_rows = []
    for row in ws_xlsx.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        entry = {}
        for i, h in enumerate(parse_headers):
            if h:
                entry[h] = row[i] if row[i] is not None else ""
        parsed_rows.append(entry)

    print(f"[파싱결과.xlsx] {len(parsed_rows)}건 로드")

    # 2) 접수명단 시트 열기
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet("접수명단")

    # 3) 없는 컬럼 추가
    needed = list(COL_MAP.keys())
    col_idx = ensure_columns(ws, needed)

    # 4) 성명 → 행번호 인덱스
    all_vals = ws.get_all_values()
    headers = ws.row_values(1)
    name_col = headers.index("성명") + 1

    name_to_row = {}
    for i, row in enumerate(all_vals[1:], start=2):
        if len(row) >= name_col:
            n = row[name_col - 1].strip()
            if n:
                name_to_row[n] = i

    # 5) 배치 업데이트
    updates = []
    skipped = []

    for data in parsed_rows:
        name = str(data.get("성명", "")).strip()
        row_idx = name_to_row.get(name)
        if not row_idx:
            skipped.append(name)
            continue

        for sheet_col, parsed_key in COL_MAP.items():
            cidx = col_idx[sheet_col]
            val = data.get(parsed_key, "")
            col_letter = rowcol_to_a1(1, cidx).rstrip("0123456789")
            updates.append({"range": f"{col_letter}{row_idx}", "values": [[val]]})

    if updates:
        ws.batch_update(updates)
        print(f"[접수명단] {len(parsed_rows) - len(skipped)}건 업데이트 완료")
    else:
        print("[접수명단] 업데이트 없음")

    if skipped:
        print(f"  매칭 안된 이름 {len(skipped)}명: {skipped[:10]}")


if __name__ == "__main__":
    main()
