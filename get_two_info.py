import sys, os
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from gsheet_writer import get_credentials
import gspread

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
TARGETS = ["탁설환", "유영주"]

creds = get_credentials()
gc = gspread.authorize(creds)
sh = gc.open_by_key(SPREADSHEET_ID)
ws = sh.worksheet("접수명단")

all_vals = ws.get_all_values()
header = all_vals[0]
data_rows = all_vals[1:]

c = {h: i for i, h in enumerate(header)}
name_col = c.get("성명", 0)
jumin_col = c.get("주민번호")
phone_col = c.get("핸드폰번호")

for row in data_rows:
    if len(row) > name_col:
        name = row[name_col].strip()
        if name in TARGETS:
            jumin = row[jumin_col].strip() if jumin_col is not None and len(row) > jumin_col else ""
            phone = row[phone_col].strip() if phone_col is not None and len(row) > phone_col else ""
            print(f"{name}: 주민={jumin} / 핸드폰={phone}")
