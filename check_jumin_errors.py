import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r'F:\종소세2026')
from gsheet_writer import get_credentials
import gspread

creds = get_credentials()
gc = gspread.authorize(creds)
sh = gc.open_by_key('1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI')
ws = sh.worksheet('접수명단')
rows = ws.get_all_records()

ERROR_NAMES = ['마금현', '지성호', '김혜린', '김진곤', '정도민', '이윤경']

print(f"{'이름':10} {'주민번호':20} {'normalize7':10} {'앞6':8} {'7번째':6}")
print("-" * 60)
for r in rows:
    name = str(r.get('성명', '') or '').strip()
    if name not in ERROR_NAMES:
        continue
    jumin_raw = str(r.get('주민번호', '') or '').strip()
    jumin_clean = jumin_raw.replace('-', '').replace(' ', '')
    jumin7 = jumin_clean[:7]
    front6 = jumin7[:6]
    seventh = jumin7[6:] if len(jumin7) >= 7 else '(없음)'
    print(f"{name:10} {jumin_raw:20} {jumin7:10} {front6:8} {seventh:6}")
