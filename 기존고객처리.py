"""
기존고객처리.py - 수임동의 완료 고객 종합소득세 안내문 조회 (건바이건)

전제조건:
  1. launch_edge.py 실행 → Edge 디버그 창 열기
  2. 홈택스에 세무사 계정으로 직접 로그인
  3. python 기존고객처리.py

처리방식: 세무사 계정 세션 유지 → 고객 주민번호 입력 → 안내문 다운
완료 후: PDF 파싱 + 구글시트 자동 동기화
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from datetime import datetime
from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import process_one, ensure_output_workbook

# ============================================================
# ★ 여기만 수정 ★  (수임동의 완료 고객만)
# ============================================================
CUSTOMERS = [
    {"name": "신정숙", "jumin_raw": "6906302917114", "phone_raw": "1039406940"},
    # {"name": "홍길동", "jumin_raw": "XXXXXXXXXXXXX", "phone_raw": "10XXXXXXXX"},
]
# ============================================================

def main():
    total = len(CUSTOMERS)
    print(f"[기존고객처리] {total}명 처리 시작")
    print(f"  ※ Edge 디버그 창 + 세무사 계정 홈택스 로그인 확인 후 실행\n")

    wb, ws = ensure_output_workbook()

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx  = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()
        page.on("dialog", lambda d: d.dismiss())

        for i, c in enumerate(CUSTOMERS, 1):
            print(f"[{i}/{total}] {c['name']}")
            r = process_one(ctx, page, c)
            ws.append([
                c["name"], str(c["jumin_raw"]), r["status"],
                r["anneam_pdf"], r["prev_income_xlsx"],
                r["biznos"], r["vat_xlsx_count"],
                r["error_msg"],
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ])
            wb.save(r"F:\종소세2026\output\결과.xlsx")
            print(f"    → {r['status']} {r['error_msg'] or ''}\n")

    print(f"[완료] {total}명 처리")

    # 파싱 + 구글시트 동기화
    names = [c["name"] for c in CUSTOMERS]
    print(f"\n[파싱 시작] {names}")
    try:
        import parse_and_sync_신규 as pm
        pm.NEW_NAMES = names
        pm.main()
    except Exception as e:
        print(f"[파싱 실패] {e}")
        print("  수동: 안내문파싱_신규동기화.py 에서 NEW_NAMES 수정 후 실행")


if __name__ == "__main__":
    main()
