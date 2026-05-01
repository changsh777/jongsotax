"""
PDF 버튼이 disabled로 막혀있음 - 강제 활성화 + JS API 직접 호출 시도
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

        preview = page.get_by_text("미리보기", exact=False).first
        preview.wait_for(timeout=10000, state="visible")
        with ctx.expect_page(timeout=15000) as popup_info:
            preview.click()
        popup = popup_info.value
        popup.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(3)

        # 시도 1: ClipReport JS API 함수 탐색
        print("\n[시도1] window 객체에서 PDF/save 관련 함수 탐색")
        funcs = popup.evaluate("""
            () => {
                const out = [];
                for (const k of Object.keys(window)) {
                    const v = window[k];
                    if (typeof v === 'function' && /pdf|save|export|print/i.test(k)) {
                        out.push(k);
                    }
                    if (typeof v === 'object' && v !== null) {
                        try {
                            for (const k2 of Object.keys(v)) {
                                if (/pdf|save|export/i.test(k2)) {
                                    out.push(`${k}.${k2}`);
                                }
                            }
                        } catch (e) {}
                    }
                }
                return out.slice(0, 30);
            }
        """)
        print(f"    발견된 함수/객체: {funcs}")

        # 시도 2: PDF 버튼 강제 활성화 + 클릭
        print("\n[시도2] PDF 버튼 disabled 클래스 제거 + 클릭")
        pdf_path1 = OUT_DIR / "test_force1.pdf"
        try:
            with popup.expect_download(timeout=10000) as dl_info:
                clicked = popup.evaluate("""
                    () => {
                        const btn = document.querySelector('.report_menu_pdf_button');
                        if (!btn) return 'no button';
                        // disabled 클래스 제거
                        btn.classList.remove('report_menu_pdf_button_svg_dis');
                        btn.classList.add('report_menu_pdf_button_svg');
                        btn.disabled = false;
                        btn.style.display = 'inline-block';
                        btn.click();
                        return 'clicked';
                    }
                """)
                print(f"    {clicked}")
            download = dl_info.value
            download.save_as(str(pdf_path1))
            print(f"    성공! {pdf_path1.stat().st_size/1024:.0f}KB")
            with pdfplumber.open(pdf_path1) as pdf:
                print(f"    페이지 수: {len(pdf.pages)}")
        except Exception as e:
            print(f"    실패: {type(e).__name__}: {str(e)[:150]}")

        # 시도 3: onclick 핸들러 직접 호출
        print("\n[시도3] PDF 버튼 onclick 핸들러 직접 호출")
        pdf_path2 = OUT_DIR / "test_force2.pdf"
        try:
            with popup.expect_download(timeout=10000) as dl_info:
                result = popup.evaluate("""
                    () => {
                        const btn = document.querySelector('.report_menu_pdf_button');
                        if (!btn) return 'no button';
                        if (btn.onclick) {
                            btn.onclick();
                            return 'onclick called';
                        }
                        // jQuery 이벤트 시도
                        if (window.$ && $(btn).trigger) {
                            $(btn).trigger('click');
                            return 'jquery triggered';
                        }
                        return 'no handler';
                    }
                """)
                print(f"    {result}")
            download = dl_info.value
            download.save_as(str(pdf_path2))
            print(f"    성공! {pdf_path2.stat().st_size/1024:.0f}KB")
            with pdfplumber.open(pdf_path2) as pdf:
                print(f"    페이지 수: {len(pdf.pages)}")
        except Exception as e:
            print(f"    실패: {type(e).__name__}: {str(e)[:150]}")

        # 시도 4: 인쇄 버튼 클릭 후 DOM 변화 확인 + popup.pdf
        print("\n[시도4] 인쇄 버튼 클릭 → window.print 후킹 → popup.pdf")
        pdf_path3 = OUT_DIR / "test_force3.pdf"
        try:
            # window.print 노출시키지 않게 미리 무력화
            popup.evaluate("window.print = () => { window.__printed = true; };")

            print_btn = popup.locator(".report_menu_print_button").first
            print_btn.click(timeout=3000)
            time.sleep(2)

            printed = popup.evaluate("() => window.__printed")
            print(f"    인쇄 함수 호출됨: {printed}")

            # 인쇄 후킹 후 popup.pdf로 캡처
            popup.pdf(
                path=str(pdf_path3),
                print_background=True,
                prefer_css_page_size=True,
            )
            print(f"    popup.pdf 결과: {pdf_path3.stat().st_size/1024:.0f}KB")
            with pdfplumber.open(pdf_path3) as pdf:
                print(f"    페이지 수: {len(pdf.pages)}")
        except Exception as e:
            print(f"    실패: {type(e).__name__}: {str(e)[:150]}")

        print("\n팝업 그대로 둡니다.")


if __name__ == "__main__":
    main()
