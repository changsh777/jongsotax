import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"F:\종소세2026")
from gsheet_writer import get_credentials
from config import CUSTOMER_DIR
import gspread

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
creds = get_credentials()
gc = gspread.authorize(creds)
ws = gc.open_by_key(SPREADSHEET_ID).worksheet("접수명단")
rows = ws.get_all_records()

no_pdf = []
has_pdf = []
for r in rows:
    name = str(r.get("성명","") or "").strip()
    구분  = str(r.get("고객구분","") or "").strip()
    jumin = str(r.get("주민번호","") or "").strip()
    if not name or 구분 != "기존":
        continue
    folders = list(CUSTOMER_DIR.glob(f"{name}_*")) + [CUSTOMER_DIR / name]
    folder = next((f for f in folders if f.is_dir()), None)
    pdf = folder and bool(list(folder.glob("종소세안내문_*.pdf")))
    if pdf:
        has_pdf.append(name)
    else:
        no_pdf.append((name, jumin))

print(f"[PDF 있음 - {len(has_pdf)}명]")
for n in has_pdf:
    print(f"  {n}")

print(f"\n[PDF 없음 - {len(no_pdf)}명] ← 처리 필요")
for n, j in no_pdf:
    print(f"  {n}  ({j})")
