"""
C안 테스트: 페이지별 스크린샷 + PDF 병합
- 박수경 1명만
- 다양한 페이지 네비게이션 전략 시도
- 결과 PDF의 페이지 수 검증
"""
from playwright.sync_api import sync_playwright
from pathlib import Path
from PIL import Image
import io
import time
import re
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


def get_total_pages(popup, default=5):
    """팝업에서 'N / M' 패턴 찾아서 M 반환"""
    try:
        result = popup.evaluate("""
            () => {
                const all = document.querySelectorAll('*');
                for (const el of all) {
                    const txt = (el.innerText || el.textContent || '').trim();
                    const m = txt.match(/^(\\d+)\\s*\\/\\s*(\\d+)$/);
                    if (m) return {current: parseInt(m[1]), total: parseInt(m[2])};
                }
                return null;
            }
        """)
        if result:
            print(f"    페이지 정보 발견: {result['current']}/{result['total']}")
            return result["total"]
        print(f"    페이지 정보 못 찾음, 기본값 {default} 사용")
        return default
    except Exception as e:
        print(f"    페이지 정보 파싱 에러: {e}")
        return default


def try_navigate_next(popup):
    """다음 페이지로 이동 - 여러 전략 순차 시도"""
    strategies = [
        ("keyboard PageDown", lambda: popup.keyboard.press("PageDown")),
        ("keyboard ArrowRight", lambda: popup.keyboard.press("ArrowRight")),
        ("click '>' or 'next'", lambda: click_next_button(popup)),
    ]
    before = popup.evaluate("""
        () => {
            const all = document.querySelectorAll('*');
            for (const el of all) {
                const txt = (el.innerText || '').trim();
                const m = txt.match(/^(\\d+)\\s*\\/\\s*\\d+$/);
                if (m) return parseInt(m[1]);
            }
            return null;
        }
    """)

    for name, action in strategies:
        try:
            action()
            time.sleep(1)
            after = popup.evaluate("""
                () => {
                    const all = document.querySelectorAll('*');
                    for (const el of all) {
                        const txt = (el.innerText || '').trim();
                        const m = txt.match(/^(\\d+)\\s*\\/\\s*\\d+$/);
                        if (m) return parseInt(m[1]);
                    }
                    return null;
                }
            """)
            if before is not None and after is not None and after > before:
                print(f"    네비 성공: {name} ({before}→{after})")
                return True
        except Exception as e:
            print(f"    {name} 실패: {e}")
    return False


def click_next_button(popup):
    """다음 페이지 버튼으로 추정되는 element 클릭"""
    popup.evaluate("""
        () => {
            const all = document.querySelectorAll('button, a, span, div, img, input');
            for (const el of all) {
                const txt = (el.innerText || el.value || el.title || el.alt || '').trim();
                const cls = (typeof el.className === 'string') ? el.className : '';
                if ((txt === '>' || txt === '다음' || cls.includes('next') || cls.includes('Next'))
                    && el.offsetParent !== null) {
                    el.click();
                    return true;
                }
            }
            return false;
        }
    """)


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()

        print("[1] 신고도움서비스 + 박수경 조회")
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

        print("[2] 미리보기 클릭 → 팝업")
        preview = page.get_by_text("미리보기", exact=False).first
        preview.wait_for(timeout=10000, state="visible")
        with ctx.expect_page(timeout=15000) as popup_info:
            preview.click()
        popup = popup_info.value
        popup.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(3)

        print(f"[3] 팝업 페이지 수 확인")
        total = get_total_pages(popup, default=5)

        print(f"[4] 페이지별 스크린샷 시작 (총 {total}장)")
        images = []
        for i in range(total):
            print(f"    [{i+1}/{total}] 캡처 중...")
            time.sleep(0.8)  # 페이지 렌더링 여유
            png = popup.screenshot(full_page=False)
            img = Image.open(io.BytesIO(png)).convert("RGB")
            images.append(img)

            if i < total - 1:
                ok = try_navigate_next(popup)
                if not ok:
                    print(f"    ❌ 다음 페이지 이동 실패. 캡처 중단")
                    break

        out_path = OUT_DIR / "test_multipage.pdf"
        if images:
            images[0].save(
                str(out_path),
                save_all=True,
                append_images=images[1:],
                format="PDF",
            )
            print(f"\n[결과] {out_path.name}")
            print(f"    크기: {out_path.stat().st_size/1024:.0f}KB")
            print(f"    저장된 페이지: {len(images)}장")

        print("\n팝업은 그대로 둡니다. 결과 확인 후 직접 닫아주세요.")


if __name__ == "__main__":
    main()
