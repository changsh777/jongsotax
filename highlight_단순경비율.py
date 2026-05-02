"""
접수명단 추계시적용경비율 컬럼에서 단순경비율 셀 빨간색 표시
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"F:\종소세2026")

from gsheet_writer import get_credentials
import gspread
from gspread.utils import rowcol_to_a1

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"

RED_BG   = {"red": 1.0, "green": 0.4, "blue": 0.4}   # 빨간
CLEAR_BG = {"red": 1.0, "green": 1.0, "blue": 1.0}   # 흰색 (초기화)

def main():
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet("접수명단")
    sheet_id = ws.id

    headers = ws.row_values(1)
    col_idx = headers.index("추계시적용경비율") + 1  # 1-indexed
    col_letter = rowcol_to_a1(1, col_idx).rstrip("0123456789")

    vals = ws.col_values(col_idx)  # 헤더 포함

    requests = []
    red_cnt = 0
    for i, val in enumerate(vals[1:], start=2):  # 2행부터
        is_red = val.strip() == "단순경비율"
        bg = RED_BG if is_red else CLEAR_BG
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": i - 1,
                    "endRowIndex": i,
                    "startColumnIndex": col_idx - 1,
                    "endColumnIndex": col_idx
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": bg
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        })
        if is_red:
            red_cnt += 1

    if requests:
        sh.batch_update({"requests": requests})
        print(f"단순경비율 {red_cnt}명 빨간색 표시 완료")

if __name__ == "__main__":
    main()
