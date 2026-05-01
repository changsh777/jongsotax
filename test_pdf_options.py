"""
A안 테스트: PDF 옵션 변경으로 5페이지 다 잡히는지 확인
- 박수경 1명만
- 다양한 옵션 조합으로 PDF 저장 후 페이지 수 비교
"""
from playwright.sync_api import sync_playwright
from pathlib import Path
import pdfplumber
import time
import warnings
import logging

warnings.filterwarnings("ignore")
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfplumber").setLevel(logging.ERROR)

OUT_DIR = Path(r"F:\종소세2026\output\PDF\테스트")
OUT_DIR.mkdir(parents=True, exist_ok=True)

REPORT_HELP_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=06&tm2lIdx=0601000000&tm3lIdx=0601200000"
)

JUMIN_FRONT = "800222"
JUMIN_BACK = "2047531"


def count_pages(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            return len(pdf.pages)
    except Exception as e:
        return f"열기실패: {e}"


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()

        print("[1] 신고도움서비스 이동")
        page.goto(REPORT_HELP_URL, wait_until="domcontentloaded")
        time.sleep(2)

        print("[2] 박수경 주민번호 입력")
        inputs = [el for el in page.locator(
            "xpath=//th[contains(.,'주민등록번호')]/following-sibling::td//input"
        ).all() if el.is_visible()]
        inputs[0].fill(JUMIN_FRONT)
        inputs[1].fill(JUMIN_BACK)

        btns = [b for b in page.locator(
            "xpath=//th[contains(.,'주민등록번호')]/following-sibling::td//input[@value='조회하기']"
        ).all() if b.is_visible()]
        btns[0].click()
        time.sleep(3)

        print("[3] 미리보기 클릭 → 팝업")
        preview = page.get_by_text("미리보기", exact=False).first
        preview.wait_for(timeout=10000, state="visible")

        with ctx.expect_page(timeout=15000) as popup_info:
            preview.click()
        popup = popup_info.value
        popup.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(3)
        print(f"    팝업 URL: {popup.url}")

        # 옵션 1: format 제거 + prefer_css_page_size
        opt1 = OUT_DIR / "test_opt1_css.pdf"
        try:
            popup.pdf(path=str(opt1), print_background=True, prefer_css_page_size=True)
            print(f"\n[옵션1: prefer_css_page_size=True]")
            print(f"    파일: {opt1.name} ({opt1.stat().st_size/1024:.0f}KB)")
            print(f"    페이지수: {count_pages(opt1)}")
        except Exception as e:
            print(f"    옵션1 실패: {e}")

        # 옵션 2: 큰 height 강제
        opt2 = OUT_DIR / "test_opt2_largeheight.pdf"
        try:
            popup.pdf(
                path=str(opt2),
                print_background=True,
                width="210mm",
                height="1500mm",  # A4 5장 분량
            )
            print(f"\n[옵션2: 큰 height(1500mm)]")
            print(f"    파일: {opt2.name} ({opt2.stat().st_size/1024:.0f}KB)")
            print(f"    페이지수: {count_pages(opt2)}")
        except Exception as e:
            print(f"    옵션2 실패: {e}")

        # 옵션 3: 화면 미디어 에뮬레이션
        opt3 = OUT_DIR / "test_opt3_screenmedia.pdf"
        try:
            popup.emulate_media(media="screen")
            popup.pdf(
                path=str(opt3),
                print_background=True,
                prefer_css_page_size=True,
            )
            print(f"\n[옵션3: screen media + css page size]")
            print(f"    파일: {opt3.name} ({opt3.stat().st_size/1024:.0f}KB)")
            print(f"    페이지수: {count_pages(opt3)}")
        except Exception as e:
            print(f"    옵션3 실패: {e}")
        finally:
            popup.emulate_media(media="print")  # 원복

        # 옵션 4: A3 가로 (페이지 크게)
        opt4 = OUT_DIR / "test_opt4_a3.pdf"
        try:
            popup.pdf(path=str(opt4), format="A3", print_background=True, landscape=True)
            print(f"\n[옵션4: A3 landscape]")
            print(f"    파일: {opt4.name} ({opt4.stat().st_size/1024:.0f}KB)")
            print(f"    페이지수: {count_pages(opt4)}")
        except Exception as e:
            print(f"    옵션4 실패: {e}")

        print("\n[팝업 그대로 두기 - 확인 후 수동으로 닫으세요]")
        input("ENTER로 종료...")


if __name__ == "__main__":
    main()
