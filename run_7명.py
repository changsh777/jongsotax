"""
run_7명.py - 특정 고객 지급명세서 + 간이용역소득 타겟 처리
이혜주/이완호/배성섭/박형우/김지혁/이재윤/두봉수

전제조건: Edge 디버그 모드 + 홈택스 세무사 로그인
실행: python F:\종소세2026\run_7명.py
"""
import sys, io, os, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import (
    _find_main_page, download_jipgum_pdf, download_ganiyiyong_xlsx,
    fill_jumin_and_search, normalize_jumin, REPORT_HELP_URL,
)
from config import CUSTOMER_DIR, customer_folder
from gsheet_writer import get_credentials
import gspread

TARGET_NAMES = ["이혜주", "이완호", "배성섭", "박형우", "김지혁", "이재윤", "두봉수"]

GSHEET_ID  = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME = "접수명단"
COL_NAME  = 2
COL_JUMIN = 4


def load_targets():
    creds = get_credentials()
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(GSHEET_ID).worksheet(SHEET_NAME)
    rows = ws.get_all_values()
    result = []
    for row in rows[1:]:
        name  = str(row[COL_NAME]).strip()  if len(row) > COL_NAME  else ""
        jumin = str(row[COL_JUMIN]).strip() if len(row) > COL_JUMIN else ""
        if name in TARGET_NAMES:
            result.append({"name": name, "jumin_raw": jumin})
    found = [r["name"] for r in result]
    missing = [n for n in TARGET_NAMES if n not in found]
    if missing:
        print(f"  [경고] 접수명단에 없음: {missing}")
    return result


def main():
    print(f"=== 타겟 {len(TARGET_NAMES)}명 처리 시작 ===")
    customers = load_targets()
    print(f"  접수명단 매칭: {len(customers)}명\n")

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = _find_main_page(ctx)
        print(f"  [메인 페이지] {page.url[:80]}\n")
        page.bring_to_front()

        for c in customers:
            name      = c["name"]
            jumin_raw = c["jumin_raw"]
            front, back = normalize_jumin(jumin_raw)
            jumin6    = front
            folder    = customer_folder(name, jumin6)

            if not folder or not folder.is_dir():
                print(f"[{name}] 폴더 없음 - 스킵")
                continue

            print(f"\n{'─'*40}")
            print(f"[{name}] {front}-{back}")

            # ① 지급명세서
            jip_dir = folder / "지급명세서"
            jip_dir.mkdir(exist_ok=True)
            jip_pdf = jip_dir / f"{name}_{jumin6}.pdf"
            if jip_pdf.exists():
                print(f"  지급명세서: 이미 있음 - 스킵")
            else:
                print(f"  지급명세서: 다운로드 중...")
                try:
                    page.goto(REPORT_HELP_URL, wait_until="networkidle", timeout=30000)
                    time.sleep(1)
                    fill_jumin_and_search(page, front, back)
                    time.sleep(2)
                    download_jipgum_pdf(ctx, page, folder, name, jumin6)
                except Exception as e:
                    print(f"  지급명세서 오류: {e}")

            # ② 간이용역소득
            ganyi_dir = folder / "간이용역소득"
            ganyi_dir.mkdir(exist_ok=True)
            files = list(ganyi_dir.iterdir())
            if len(files) >= 2:
                print(f"  간이용역: 이미 있음 - 스킵")
            else:
                print(f"  간이용역: 다운로드 중...")
                try:
                    page.goto(REPORT_HELP_URL, wait_until="networkidle", timeout=30000)
                    time.sleep(1)
                    fill_jumin_and_search(page, front, back)
                    time.sleep(2)
                    download_ganiyiyong_xlsx(ctx, page, folder, name, jumin6)
                except Exception as e:
                    print(f"  간이용역 오류: {e}")

            time.sleep(1)

    print(f"\n=== 완료 ===")


if __name__ == "__main__":
    main()
