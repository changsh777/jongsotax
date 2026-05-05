"""지급명세서 + 간이용역 다운로드 단독 테스트"""
import sys, os, time
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import (
    REPORT_HELP_URL, normalize_jumin, fill_jumin_and_search,
    wait_preview_button, download_jipgum_pdf, download_ganiyiyong_xlsx,
    customer_folder
)
from gsheet_writer import get_credentials
import gspread

NAME = sys.argv[1] if len(sys.argv) > 1 else "박명훈"
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"

# 구글시트에서 주민번호 읽기
creds = get_credentials()
gc = gspread.authorize(creds)
ws = gc.open_by_key(SPREADSHEET_ID).worksheet("접수명단")
rows = ws.get_all_records()

customer = None
for r in rows:
    if str(r.get("성명", "")).strip() == NAME:
        customer = {
            "name": NAME,
            "jumin_raw": str(r.get("주민번호", "")).strip(),
        }
        break

if not customer:
    print(f"[오류] '{NAME}' 구글시트에서 찾을 수 없음")
    sys.exit(1)

print(f"[테스트] {NAME} 지급명세서+간이용역 테스트")
print(f"  주민번호: {str(customer['jumin_raw'])[:6]}******")

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    ctx  = browser.contexts[0]
    page = ctx.pages[0]
    page.bring_to_front()
    page.on("dialog", lambda d: d.dismiss())

    front, back = normalize_jumin(customer["jumin_raw"])
    folder = customer_folder(NAME, customer["jumin_raw"])

    # 신고도움서비스 이동 + 조회
    page.goto(REPORT_HELP_URL, wait_until="domcontentloaded")
    time.sleep(2)
    fill_jumin_and_search(page, front, back)
    time.sleep(3)

    btn = wait_preview_button(page, timeout_ms=10000)
    print(f"  미리보기 버튼 존재: {btn is not None}")

    # 지급명세서 PDF
    print(f"\n[1] 지급명세서 PDF 다운로드...")
    ok1 = download_jipgum_pdf(ctx, page, folder, NAME, customer["jumin_raw"])
    print(f"  결과: {'성공' if ok1 else '실패/자료없음'}")

    # 간이용역 엑셀
    print(f"\n[2] 간이용역 엑셀 다운로드...")
    ok2 = download_ganiyiyong_xlsx(page, folder, NAME, customer["jumin_raw"])
    print(f"  결과: {'성공' if ok2 else '실패/자료없음'}")

print(f"\n[완료] 지급명세서:{ok1}, 간이용역:{ok2}")
