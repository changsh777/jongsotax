"""
test_ganiyiyong.py - 이성민으로 간이용역 단독 테스트
사전 조건: Edge CDP 열림 + 홈택스 세무사 로그인 + 이성민 조회 완료 상태
"""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import os
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import (
    download_ganiyiyong_xlsx, _find_main_page, _find_ganiyiyong_popup
)
from config import customer_folder

# 테스트 대상 (주민번호는 실제값으로 변경 필요)
TEST_NAME = "이성민"
TEST_JUMIN = "650501-1234567"  # ← 실제 주민번호로 변경

def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]

        print(f"열린 페이지 수: {len(ctx.pages)}")
        for i, pg in enumerate(ctx.pages):
            print(f"  [{i}] {pg.url[:80]}")

        main_page = _find_main_page(ctx)
        print(f"\n[메인 페이지] {main_page.url[:80]}")

        existing_popup = _find_ganiyiyong_popup(ctx)
        print(f"[기존 팝업] {existing_popup.url[:60] if existing_popup else '없음'}")

        folder = customer_folder(TEST_NAME, TEST_JUMIN)
        print(f"[저장 폴더] {folder}")

        print(f"\n[테스트 시작] {TEST_NAME}")
        result = download_ganiyiyong_xlsx(
            ctx, main_page, folder, TEST_NAME, TEST_JUMIN
        )
        print(f"\n[결과] {'성공' if result else '실패/자료없음'}")

if __name__ == "__main__":
    main()
