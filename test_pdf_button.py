"""
ClipReport4 뷰어의 PDF 버튼 직접 클릭 테스트
- pdf.svg가 로드된 점에서 PDF 버튼 존재 확인됨
- 클래스명 패턴: report_menu_pdf_button 추정
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


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()

        print("[1] 박수경 조회")
        page.goto(REPORT_HELP_URL, wait_until="domcontentloaded")
        time.sleep(2)
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

        print("[2] 미리보기 → 팝업")
        preview = page.get_by_text("미리보기", exact=False).first
        preview.wait_for(timeout=10000, state="visible")
        with ctx.expect_page(timeout=15000) as popup_info:
            preview.click()
        popup = popup_info.value
        popup.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(3)

        print("[3] 뷰어 메뉴 버튼 전체 탐색")
        menu_buttons = popup.evaluate("""
            () => {
                const buttons = document.querySelectorAll('.report_menu_button, [class*="report_menu"]');
                const out = [];
                for (const b of buttons) {
                    out.push({
                        tag: b.tagName,
                        cls: (typeof b.className === 'string') ? b.className : '',
                        title: b.title || '',
                        txt: (b.innerText || '').trim().slice(0, 20),
                        visible: b.offsetParent !== null,
                    });
                }
                return out;
            }
        """)
        for i, m in enumerate(menu_buttons):
            print(f"    [{i}] {m}")

        # PDF 버튼 시도
        print("\n[4] PDF 버튼 클릭 시도")
        pdf_path = OUT_DIR / "test_pdf_button.pdf"
        try:
            with popup.expect_download(timeout=15000) as dl_info:
                # 클래스명에 'pdf' 포함된 버튼 클릭
                clicked = popup.evaluate("""
                    () => {
                        const all = document.querySelectorAll('button, a, [class*="report_menu"]');
                        for (const el of all) {
                            const cls = (typeof el.className === 'string') ? el.className.toLowerCase() : '';
                            const title = (el.title || '').toLowerCase();
                            if ((cls.includes('pdf') || title.includes('pdf')) && el.offsetParent !== null) {
                                el.click();
                                return {tag: el.tagName, cls: el.className, title: el.title};
                            }
                        }
                        return null;
                    }
                """)
                print(f"    클릭한 element: {clicked}")
            download = dl_info.value
            download.save_as(str(pdf_path))
            print(f"\n[5] PDF 다운로드 성공!")
            print(f"    파일: {pdf_path.name} ({pdf_path.stat().st_size/1024:.0f}KB)")
            with pdfplumber.open(pdf_path) as pdf:
                print(f"    페이지 수: {len(pdf.pages)}")
                if pdf.pages:
                    text = pdf.pages[0].extract_text() or ""
                    print(f"    1페이지 텍스트 미리보기 (200자):")
                    print(f"    {text[:200]}")
        except Exception as e:
            print(f"    PDF 다운로드 실패: {type(e).__name__}: {e}")

        print("\n팝업은 그대로 둡니다.")


if __name__ == "__main__":
    main()
