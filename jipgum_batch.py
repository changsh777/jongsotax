"""
jipgum_batch.py - 구글시트 접수명단 기준 지급명세서 일괄 PDF 다운로드

처리: 지급명세서 일괄출력 → PDF 저장
스킵: PDF 이미 있으면 스킵
재개: python jipgum_batch.py [시작번호]  (예: 10 → 10번부터)

전제조건:
  1. python launch_edge.py (Edge 디버그 창)
  2. 홈택스 세무사 계정 로그인
"""
import sys, io, os, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import (
    _find_main_page, download_jipgum_pdf,
    fill_jumin_and_search, normalize_jumin,
    REPORT_HELP_URL,
)
from config import CUSTOMER_DIR, customer_folder
from gsheet_writer import get_credentials
import gspread

GSHEET_ID  = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME = "접수명단"

COL_NAME  = 2   # 성명 (0-based)
COL_JUMIN = 4   # 주민번호


def load_customers_from_gsheet():
    creds = get_credentials()
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(GSHEET_ID).worksheet(SHEET_NAME)
    rows = ws.get_all_values()
    customers = []
    for row in rows[1:]:   # 헤더 제외
        name  = str(row[COL_NAME]).strip()  if len(row) > COL_NAME  else ""
        jumin = str(row[COL_JUMIN]).strip() if len(row) > COL_JUMIN else ""
        if not name:
            continue
        customers.append({"name": name, "jumin_raw": jumin})
    return customers


def already_done(folder, name, jumin_raw):
    """PDF 이미 있으면 True"""
    jumin6 = str(jumin_raw).replace("-", "").replace(" ", "")[:6]
    pdf = folder / "지급명세서" / f"{name}_{jumin6}.pdf"
    return pdf.exists()


def main():
    start_idx = int(sys.argv[1]) - 1 if len(sys.argv) > 1 else 0

    print("[지급명세서 배치] 구글시트 접수명단 로드 중...", flush=True)
    customers = load_customers_from_gsheet()
    total = len(customers)
    print(f"  총 {total}명, {start_idx + 1}번부터 시작\n", flush=True)

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = _find_main_page(ctx)
        print(f"  [메인 페이지] {page.url[:80]}", flush=True)
        page.bring_to_front()
        page.on("dialog", lambda d: d.dismiss())

        for i, c in enumerate(customers[start_idx:], start_idx + 1):
            name  = c["name"]
            jumin = str(c.get("jumin_raw", "")).replace("-", "").replace(" ", "").strip()

            print(f"[{i}/{total}] {name}", flush=True)

            # 주민번호 불완전하면 스킵
            if len(jumin) < 13:
                print(f"    [스킵] 주민번호 없음/불완전 ({jumin})", flush=True)
                continue

            folder = customer_folder(name, jumin)

            # 이미 완료된 경우 스킵
            if already_done(folder, name, jumin):
                print(f"    [스킵] 기존 파일 존재", flush=True)
                continue

            # 홈택스 조회
            try:
                front, back = normalize_jumin(jumin)
            except ValueError as e:
                print(f"    [스킵] 주민번호 오류: {e}", flush=True)
                continue

            try:
                page.goto(REPORT_HELP_URL, wait_until="domcontentloaded", timeout=20000)
                time.sleep(2)
                fill_jumin_and_search(page, front, back)
                time.sleep(3)
            except Exception as e:
                print(f"    [에러] 조회 실패: {e}", flush=True)
                continue

            # 지급명세서 다운로드
            try:
                result = download_jipgum_pdf(ctx, page, folder, name, jumin)
                status = "성공" if result else "자료없음/실패"
                print(f"    → {status}\n", flush=True)
            except Exception as e:
                print(f"    [에러] 지급명세서: {e}\n", flush=True)

    print(f"[완료] {total}명 처리")


if __name__ == "__main__":
    main()
