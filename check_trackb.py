import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ['SEOTAX_ENV'] = 'nas'
sys.path.insert(0, r'F:\종소세2026')
from gsheet_writer import get_credentials
import gspread
from pathlib import Path

creds = get_credentials()
gc = gspread.authorize(creds)
sh = gc.open_by_key('1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI')
ws = sh.worksheet('접수명단')
rows = ws.get_all_records()

# PDF 없는 사람 목록 (NAS 기준)
nas_customer = Path(r'Z:\종소세2026\고객')
no_pdf = set()
for folder in nas_customer.iterdir():
    if not folder.is_dir(): continue
    pdfs = list(folder.glob('종소세안내문_*.pdf')) + list(folder.glob('자료/종소세안내문_*.pdf'))
    if not pdfs:
        no_pdf.add(folder.name.split('_')[0])

print(f'PDF 없는 고객: {len(no_pdf)}명')
print()

# 홈택스 ID/PW 있는 사람
has_cred = []
for r in rows:
    name = str(r.get('성명', '') or '').strip()
    ht_id = str(r.get('홈택스아이디', '') or '').strip()
    ht_pw = str(r.get('홈택스비번', '') or '').strip()
    if ht_id and ht_pw:
        has_cred.append(name)

print(f'홈택스 ID/PW 입력된 사람: {len(has_cred)}명')
print()

# 교집합: PDF 없고 + ID/PW 있는 사람
target = [n for n in has_cred if n in no_pdf]
print(f'Track B 대상 (PDF없음 + ID/PW있음): {len(target)}명')
for n in target:
    print(f'  {n}')

# PDF 없는데 ID/PW도 없는 사람
no_cred_no_pdf = [n for n in no_pdf if n not in has_cred]
print(f'\nPDF없음 + ID/PW없음 (수동처리 필요): {len(no_cred_no_pdf)}명')
for n in sorted(no_cred_no_pdf):
    print(f'  {n}')
