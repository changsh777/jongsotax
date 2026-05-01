"""
팝업 네트워크 트래픽 모니터링
- ClipReport 팝업이 서버에서 어떤 데이터를 받는지 추적
- PDF 응답이 있으면 직접 캡처 가능 (가장 깔끔한 해결)
"""
from playwright.sync_api import sync_playwright
from pathlib import Path
import time
import json

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
    captured_responses = []

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

        print("[2] 미리보기 → 팝업 (네트워크 모니터링 시작)")
        preview = page.get_by_text("미리보기", exact=False).first
        preview.wait_for(timeout=10000, state="visible")

        # 새 팝업 페이지 캡처 + 응답 모니터링
        def on_response(response):
            url = response.url
            ct = response.headers.get("content-type", "")
            try:
                length = int(response.headers.get("content-length", "0"))
            except Exception:
                length = 0
            captured_responses.append({"url": url, "type": ct, "size": length})

        ctx.on("response", on_response)

        with ctx.expect_page(timeout=15000) as popup_info:
            preview.click()
        popup = popup_info.value
        popup.wait_for_load_state("networkidle", timeout=30000)
        time.sleep(5)  # 추가 트래픽 대기

        print(f"\n[3] 캡처된 응답 {len(captured_responses)}개 분석")
        # PDF, octet-stream, 큰 사이즈 위주로 보기
        interesting = []
        for r in captured_responses:
            if (
                "pdf" in r["type"].lower()
                or "octet-stream" in r["type"].lower()
                or "msexcel" in r["type"].lower()
                or "spreadsheet" in r["type"].lower()
                or r["size"] > 100000  # 100KB 이상
            ):
                interesting.append(r)

        print(f"\n[관심 응답 {len(interesting)}개]")
        for i, r in enumerate(interesting, 1):
            print(f"  {i}. {r['type']} | {r['size']/1024:.0f}KB")
            print(f"     {r['url'][:120]}")

        # sesw.hometax 도메인 모든 응답
        print(f"\n[sesw.hometax 도메인 응답]")
        for r in captured_responses:
            if "sesw.hometax" in r["url"]:
                print(f"  {r['type']} | {r['size']/1024:.0f}KB")
                print(f"  {r['url'][:120]}")

        # 인쇄 버튼 위치/존재 확인
        print(f"\n[4] 팝업 안 인쇄 버튼 탐색")
        result = popup.evaluate("""
            () => {
                const all = document.querySelectorAll('*');
                const candidates = [];
                for (const el of all) {
                    const txt = (el.innerText || el.title || el.alt || '').trim();
                    const cls = (typeof el.className === 'string') ? el.className : '';
                    if (
                        (txt === '인쇄' || txt === 'Print' || txt === '출력')
                        || cls.toLowerCase().includes('print')
                        || (el.tagName === 'IMG' && (el.src || '').toLowerCase().includes('print'))
                    ) {
                        if (el.offsetParent !== null) {
                            candidates.push({
                                tag: el.tagName,
                                txt: txt.slice(0, 30),
                                cls: cls.slice(0, 50),
                                title: (el.title || '').slice(0, 30),
                                src: (el.src || '').slice(0, 80),
                            });
                        }
                    }
                }
                return candidates;
            }
        """)
        print(f"    인쇄 후보 {len(result)}개:")
        for r in result:
            print(f"      {r}")

        print("\n팝업 그대로 둡니다. 브라우저 직접 닫으세요.")


if __name__ == "__main__":
    main()
