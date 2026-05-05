import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"F:\종소세2026")
from gsheet_writer import get_credentials
import gspread

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
creds = get_credentials()
gc = gspread.authorize(creds)
sh = gc.open_by_key(SPREADSHEET_ID)
ws = sh.worksheet("접수명단")
all_vals = ws.get_all_values()
headers = all_vals[0]
print("헤더(앞20):", headers[:20])
for row in all_vals[1:]:
    name = row[0] if row else ""
    if "한효성" in name or "효성" in name:
        d = dict(zip(headers, row))
        for k, v in d.items():
            if v:
                print(f"  {k}: {v}")
        break
