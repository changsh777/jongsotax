"""이준엽 정보 구글시트에서 읽어서 _run_one.py 실행"""
import os, sys, subprocess
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from gsheet_writer import get_credentials
import gspread

NAME = sys.argv[1] if len(sys.argv) > 1 else "이준엽"

creds = get_credentials()
gc = gspread.authorize(creds)
ws = gc.open_by_key("1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI").worksheet("접수명단")
rows = ws.get_all_records()

for r in rows:
    if str(r.get("성명", "")).strip() == NAME:
        hometax_id = str(r.get("홈택스아이디", "")).strip()
        hometax_pw = str(r.get("홈택스비번", "")).strip()
        jumin = str(r.get("주민번호", "")).replace("-", "").strip()
        if jumin.isdigit() and len(jumin) < 13:
            jumin = jumin.zfill(13)
        print(f"[{NAME}] ID:{hometax_id} / 주민번호:{jumin[:6]}******")
        if not hometax_id or not hometax_pw:
            print(f"[오류] 홈택스 아이디/비번 없음 → 구글시트 확인 필요")
            sys.exit(1)
        subprocess.run([
            sys.executable, r"F:\종소세2026\_run_one.py",
            NAME, hometax_id, hometax_pw, jumin
        ])
        sys.exit(0)

print(f"[오류] '{NAME}' 구글시트에서 찾을 수 없음")
sys.exit(1)
