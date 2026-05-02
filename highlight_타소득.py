"""
접수명단 셀 색상 표시:
- 타소득(이자/배당/근로/연금/기타) O → 빨간색
- 장부유형 복식부기의무자 → 노란색
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"F:\종소세2026")

from gsheet_writer import get_credentials
import gspread

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
타소득_COLS = ["이자", "배당", "근로(단일)", "근로(복수)", "연금", "기타"]

RED_BG    = {"red": 1.0, "green": 0.4, "blue": 0.4}
YELLOW_BG = {"red": 1.0, "green": 0.95, "blue": 0.2}
CLEAR_BG  = {"red": 1.0, "green": 1.0, "blue": 1.0}

def main():
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet("접수명단")
    sheet_id = ws.id

    headers = ws.row_values(1)
    all_vals = ws.get_all_values()

    # 컬럼 인덱스
    타소득_idx = {c: headers.index(c) for c in 타소득_COLS if c in headers}
    장부_idx   = headers.index("장부유형") if "장부유형" in headers else None

    requests = []
    red_cnt = yellow_cnt = 0

    for row_i, row in enumerate(all_vals[1:], start=1):  # 0-indexed
        # 타소득 O → 빨간색
        for col_name, col_i in 타소득_idx.items():
            val = row[col_i].strip() if len(row) > col_i else ""
            is_red = (val == "O")
            requests.append({
                "repeatCell": {
                    "range": {"sheetId": sheet_id,
                               "startRowIndex": row_i, "endRowIndex": row_i + 1,
                               "startColumnIndex": col_i, "endColumnIndex": col_i + 1},
                    "cell": {"userEnteredFormat": {"backgroundColor": RED_BG if is_red else CLEAR_BG}},
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })
            if is_red: red_cnt += 1

        # 장부유형 복식부기의무자 → 노란색
        if 장부_idx is not None:
            val = row[장부_idx].strip() if len(row) > 장부_idx else ""
            is_yellow = "복식" in val
            requests.append({
                "repeatCell": {
                    "range": {"sheetId": sheet_id,
                               "startRowIndex": row_i, "endRowIndex": row_i + 1,
                               "startColumnIndex": 장부_idx, "endColumnIndex": 장부_idx + 1},
                    "cell": {"userEnteredFormat": {"backgroundColor": YELLOW_BG if is_yellow else CLEAR_BG}},
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })
            if is_yellow: yellow_cnt += 1

    # 배치 전송 (1000개씩)
    for i in range(0, len(requests), 1000):
        sh.batch_update({"requests": requests[i:i+1000]})
        print(f"  {min(i+1000, len(requests))}/{len(requests)} 처리 중...")

    print(f"완료: 타소득O {red_cnt}건 빨간색 / 복식부기 {yellow_cnt}건 노란색")

if __name__ == "__main__":
    main()
