"""
step2_one_customer.py
- 오상연 1명 PoC: 주민번호 입력 → 조회 → 미리보기 → PDF 저장
- 사전 조건: launch_edge.bat 실행 후 홈택스 로그인 + 신고도움서비스 진입 완료 상태
"""
from playwright.sync_api import sync_playwright
from pathlib import Path
import time

PDF_DIR = Path(r"F:\종소세2026\output\PDF")
PDF_DIR.mkdir(parents=True, exist_ok=True)

REPORT_HELP_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=06&tm2lIdx=0601000000&tm3lIdx=0601200000"
)

# 테스트 대상 1명
NAME = "오상연"
JUMIN_FRONT = "841212"
JUMIN_BACK = "1056818"


def main():
    with sync_playwright() as p:
        print(f"[1] 엣지에 CDP 붙기")
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()

        if "tm3lIdx=0601200000" not in page.url:
            print(f"[1-1] 신고도움서비스 URL이 아니네요. 이동합니다.")
            page.goto(REPORT_HELP_URL, wait_until="domcontentloaded")
            time.sleep(2)

        print(f"[2] 주민번호 입력: {NAME} ({JUMIN_FRONT}-{JUMIN_BACK})")
        # 주민등록번호 행 td 안의 모든 input 중 화면에 보이는 것만 사용
        # (숨겨진 '주민등록번호 외의 번호' 필드 제외)
        all_inputs = page.locator(
            "xpath=//th[contains(normalize-space(.),'주민등록번호')]"
            "/following-sibling::td//input"
        ).all()
        visible_inputs = [el for el in all_inputs if el.is_visible()]
        print(f"    전체 input {len(all_inputs)}개 중 보이는 것 {len(visible_inputs)}개")

        if len(visible_inputs) < 2:
            print(f"    [에러] 주민번호 입력 필드 2개가 안 보입니다. 중단.")
            return

        visible_inputs[0].fill(JUMIN_FRONT)
        visible_inputs[1].fill(JUMIN_BACK)

        print(f"[3] 조회하기 클릭 (주민번호 행)")
        # 엘리먼트 타입 불문: button/input/a/span 중 텍스트가 '조회하기'인 것
        # 주민번호 행(td) 안으로 스코프 한정
        search_btn_candidates = page.locator(
            "xpath=//th[contains(normalize-space(.),'주민등록번호')]"
            "/following-sibling::td//*[normalize-space(text())='조회하기']"
        ).all()
        visible_btns = [b for b in search_btn_candidates if b.is_visible()]
        print(f"    조회하기 후보 {len(search_btn_candidates)}개 중 보이는 것 {len(visible_btns)}개")

        if not visible_btns:
            # fallback: input[value=조회하기]
            fallback = page.locator(
                "xpath=//th[contains(.,'주민등록번호')]/following-sibling::td//input[@value='조회하기']"
            ).all()
            visible_btns = [b for b in fallback if b.is_visible()]
            print(f"    fallback input[value=조회하기]: {len(visible_btns)}개")

        if not visible_btns:
            print(f"    [에러] 조회하기 버튼을 찾지 못함. 중단.")
            return

        visible_btns[0].click()

        print(f"[4] 조회 결과 로드 대기 (3초)")
        time.sleep(3)

        print(f"[5] 미리보기 버튼 대기 (최대 15초)")
        preview_btn = page.get_by_text("미리보기", exact=False).first
        try:
            preview_btn.wait_for(timeout=15000, state="visible")
            print(f"    → 미리보기 나타남")
        except Exception as e:
            print(f"    → 미리보기 버튼 없음. 조회 실패/데이터 없음으로 판단")
            print(f"    에러: {e}")
            return

        print(f"[6] 미리보기 클릭 → 팝업 대기")
        with ctx.expect_page(timeout=15000) as popup_info:
            preview_btn.click()
        popup = popup_info.value

        print(f"[7] 팝업 로드 대기")
        popup.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(2)  # 리포트 렌더링 여유
        print(f"    팝업 URL: {popup.url}")

        print(f"[8] PDF로 저장")
        pdf_path = PDF_DIR / f"종소세안내문_{NAME}.pdf"
        popup.pdf(
            path=str(pdf_path),
            format="A4",
            print_background=True,
            margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"},
        )
        print(f"    저장 완료: {pdf_path}")

        popup.close()
        print(f"\n[완료] {pdf_path}")


if __name__ == "__main__":
    main()
