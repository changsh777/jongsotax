import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"F:\종소세2026")
from gsheet_writer import get_credentials
from config import CUSTOMER_DIR
import gspread

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
creds = get_credentials()
gc = gspread.authorize(creds)
sh = gc.open_by_key(SPREADSHEET_ID)
ws = sh.worksheet("접수명단")
all_vals = ws.get_all_values()
headers = all_vals[0]
idx = {h: i for i, h in enumerate(headers)}

신규, 기존 = [], []
for row in all_vals[1:]:
    def g(k): return row[idx[k]] if k in idx and idx[k] < len(row) else ""
    구분 = g("고객구분"); name = g("성명"); 동의 = g("수임동의완료여부")
    아이디 = g("홈택스아이디"); 수입 = g("수입")
    if not name.strip(): continue
    # PDF 존재 여부
    folders = list(CUSTOMER_DIR.glob(f"{name.strip()}_*")) + [CUSTOMER_DIR / name.strip()]
    pdf = any(list(f.glob("종소세안내문_*.pdf")) for f in folders if f.is_dir())
    entry = (name.strip(), 동의, 아이디, 수입, "✓PDF" if pdf else "")
    if 구분 == "신규": 신규.append(entry)
    elif 구분 == "기존": 기존.append(entry)

print(f"[신규 {len(신규)}명]")
for n,d,a,s,p in 신규:
    print(f"  {n:<8} 동의={d:<6} 아이디={('O' if a else 'X')} 수입={s:<12} {p}")
print(f"\n[기존 {len(기존)}명]")
for n,d,a,s,p in 기존:
    print(f"  {n:<8} 동의={d:<6} 아이디={('O' if a else 'X')} 수입={s:<12} {p}")
