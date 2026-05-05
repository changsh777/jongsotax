"""손준, 공진주, 임효열 신규 처리 (홈택스 PDF 다운 + 파싱)"""
import sys, os, subprocess
os.environ['SEOTAX_ENV'] = 'nas'
sys.path.insert(0, r'F:\종소세2026')

from gsheet_writer import get_credentials
import gspread

gc = gspread.authorize(get_credentials())
ws = gc.open_by_key('1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI').worksheet('접수명단')
rows = ws.get_all_records()

TARGETS = ['손준', '공진주', '임효열']

for r in rows:
    name = str(r.get('성명', '')).strip()
    if name not in TARGETS:
        continue
    hid   = str(r.get('홈택스아이디', '') or '').strip()
    pw    = str(r.get('홈택스비번', '') or '').strip()
    jumin = str(r.get('주민번호', '') or '').replace('-', '').strip()
    if jumin.isdigit() and len(jumin) < 13:
        jumin = jumin.zfill(13)

    if not hid or not jumin:
        print(f'[{name}] 홈택스ID 또는 주민번호 없음 — 스킵')
        continue

    print(f'\n[{name}] 처리 시작...', flush=True)
    result = subprocess.run(
        [sys.executable, r'F:\종소세2026\_run_one.py', name, hid, pw, jumin],
        cwd=r'F:\종소세2026'
    )
    print(f'[{name}] 완료 (코드: {result.returncode})', flush=True)
