import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")
from gsheet_writer import get_credentials
import gspread

creds = get_credentials()
gc = gspread.authorize(creds)
sh = gc.open_by_key('1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI')
ws = sh.worksheet('접수명단')
rows = ws.get_all_records()

vals = set()
for r in rows:
    vals.add(repr(str(r.get('수임동의완료여부','')).strip()))
print('수임동의완료여부 값 종류:', vals)

target = {'마금현','지성호','김진곤','정도민','이윤경','김혜린'}
for r in rows:
    name = str(r.get('성명','')).strip()
    if name in target:
        print(f"  {name}: 수임동의={r.get('수임동의완료여부','')!r} / ID={r.get('홈택스아이디','')!r}")
