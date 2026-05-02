"""
안내문파싱 시트에 에러사유 컬럼 추가 및 기재
- 처리상태=완료 → 에러사유 공백
- 처리상태=에러 → 사유 분류 기재
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from gsheet_writer import get_credentials
import gspread

# 2차인증 실패 확인된 명단
TWO_FA_FAIL = {'마금현', '지성호', '김진곤', '정도민', '이윤경'}
JUMIN_SHORT  = {'김혜린'}   # 주민번호 6자리만 (rilriria 계정)

def main():
    creds = get_credentials()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key('1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI')

    # ── 접수명단에서 수임동의여부·홈택스ID 읽기 ──
    ws_order   = sh.worksheet('접수명단')
    order_data = ws_order.get_all_records()
    order_map  = {}
    for r in order_data:
        name = str(r.get('성명', '') or '').strip()
        order_map[name] = {
            'suim' : str(r.get('수임동의완료여부', '') or '').strip(),
            'ht_id': str(r.get('홈택스아이디', '') or '').strip(),
        }
    print(f"접수명단 로드: {len(order_map)}명")

    # ── 안내문파싱 시트 ──
    ws      = sh.worksheet('안내문파싱')
    headers = ws.row_values(1)

    # 에러사유 컬럼 확인 / 추가
    if '에러사유' not in headers:
        col_idx = len(headers) + 1
        # 컬럼 수 부족하면 시트 확장
        if ws.col_count < col_idx:
            ws.resize(rows=ws.row_count, cols=col_idx)
        ws.update_cell(1, col_idx, '에러사유')
        headers.append('에러사유')
        print(f"에러사유 컬럼 추가 (열 {col_idx})")
    else:
        col_idx = headers.index('에러사유') + 1
        print(f"에러사유 컬럼 기존 (열 {col_idx})")

    status_col = headers.index('처리상태') + 1
    name_col   = headers.index('성명') + 1

    all_rows = ws.get_all_values()
    updates  = []

    for i, row in enumerate(all_rows[1:], start=2):
        name   = row[name_col - 1].strip()   if len(row) >= name_col   else ''
        status = row[status_col - 1].strip() if len(row) >= status_col else ''
        if not name:
            continue

        info  = order_map.get(name, {})
        suim  = info.get('suim', '')
        ht_id = info.get('ht_id', '')

        # PDF경로로 Track 판별
        pdf_path_col = headers.index('PDF경로') + 1 if 'PDF경로' in headers else None
        pdf_path = row[pdf_path_col - 1].strip() if pdf_path_col and len(row) >= pdf_path_col else ''
        is_track_b = 'TrackB' in pdf_path

        if status == '완료':
            reason = ''
        elif status == '에러':
            if name in TWO_FA_FAIL:
                reason = '2차인증실패(주민번호불일치)'
            elif name in JUMIN_SHORT:
                reason = '주민번호부족(6자리)'
            elif ht_id:
                reason = 'TrackB로그인실패'
            else:
                reason = 'ID/PW없음-수동처리필요'
        else:
            # 처리상태 없음 = PDF없음
            if not status:
                if ht_id:
                    reason = 'TrackB-PDF없음'
                else:
                    reason = 'TrackA-PDF없음'
            else:
                reason = ''

        # 현재값과 다를 때만 업데이트
        cur = row[col_idx - 1].strip() if len(row) >= col_idx else ''
        if cur != reason:
            col_letter = ''
            n = col_idx
            while n:
                n, r = divmod(n - 1, 26)
                col_letter = chr(65 + r) + col_letter
            updates.append({'range': f'{col_letter}{i}', 'values': [[reason]]})

    if updates:
        ws.batch_update(updates)
        print(f"에러사유 업데이트: {len(updates)}건")
    else:
        print("변경 없음")

    # 에러 요약 출력
    err_summary = {}
    for u in updates:
        r = u['values'][0][0]
        if r:
            err_summary[r] = err_summary.get(r, 0) + 1

    print("\n[에러 사유 분류]")
    for reason, cnt in sorted(err_summary.items(), key=lambda x: -x[1]):
        print(f"  {reason}: {cnt}명")

if __name__ == "__main__":
    main()
