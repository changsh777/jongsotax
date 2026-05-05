"""
test_jipgum.py - download_jipgum_pdf 실제 호출 테스트 v14

사용법:
  python test_jipgum.py 박일준
"""
import sys, io, time, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import (
    _find_main_page,
    fill_jumin_and_search, normalize_jumin,
    REPORT_HELP_URL, download_jipgum_pdf,
)
from config import customer_folder
from gsheet_writer import get_credentials
import gspread

GSHEET_ID  = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME = "접수명단"
COL_NAME  = 2
COL_JUMIN = 4

def get_jumin_from_sheet(target_name):
    creds = get_credentials()
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(GSHEET_ID).worksheet(SHEET_NAME)
    rows = ws.get_all_values()
    for row in rows[1:]:
        name  = row[COL_NAME].strip()  if len(row) > COL_NAME  else ""
        jumin = row[COL_JUMIN].strip() if len(row) > COL_JUMIN else ""
        if name == target_name and len(jumin.replace("-","")) >= 13:
            return jumin
    return None


def main():
    name = sys.argv[1] if len(sys.argv) > 1 else "김도영"

    print(f"[구글시트] {name} 주민번호 조회 중...")
    jumin = get_jumin_from_sheet(name)
    if not jumin:
        print(f"  주민번호 없음 — 종료"); sys.exit(1)
    print(f"  {name} / {jumin[:6]}-{jumin[6:7]}******")

    jumin_plain = jumin.replace("-", "").replace(" ", "")
    folder = customer_folder(name, jumin_plain)
    print(f"  폴더: {folder}")

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = _find_main_page(ctx)
        print(f"\n[메인 페이지] {page.url[:80]}")
        page.bring_to_front()
        page.on("dialog", lambda d: (print(f"  [DIALOG] {d.message[:80]}"), d.dismiss()))

        # 홈택스 조회
        front, back = normalize_jumin(jumin_plain)
        page.goto(REPORT_HELP_URL, wait_until="domcontentloaded", timeout=20000)
        time.sleep(2)
        fill_jumin_and_search(page, front, back)
        time.sleep(4)

        # download_jipgum_pdf 호출
        print(f"\n[download_jipgum_pdf 호출]")
        result = download_jipgum_pdf(ctx, page, folder, name, jumin_plain)
        print(f"\n결과: {result}")

        # 저장된 파일 확인
        jipgum_dir = folder / "지급명세서"
        if jipgum_dir.exists():
            files = list(jipgum_dir.glob("*.pdf"))
            print(f"지급명세서 폴더: {[f.name for f in files]}")
        else:
            print(f"지급명세서 폴더 없음")


if __name__ == "__main__":
    main()
