"""
_inspect_popup.py — 신고내역조회 팝업 HTML 검사 + 버튼 ID 확인
실행 전 홈택스 Edge에서 신고내역조회 팝업이 열려 있어야 함
"""
import sys, io, time
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

from playwright.sync_api import sync_playwright

RESULT_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=04&tm2lIdx=0405000000&tm3lIdx=0405040000"
)
CDP_PORT = 9222

def inspect():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
        ctx = browser.contexts[0]
        pages = ctx.pages
        print(f"열린 탭 수: {len(pages)}")
        for i, pg in enumerate(pages):
            try:
                url = pg.url
                print(f"  [{i}] {url[:100]}")
            except:
                print(f"  [{i}] (error)")

        # 홈택스 메인 페이지 찾기
        main_page = None
        for pg in pages:
            try:
                if "hometax.go.kr" in pg.url:
                    main_page = pg
                    break
            except:
                pass

        if not main_page:
            print("홈택스 탭 없음 — 먼저 홈택스 로그인 필요")
            return

        print(f"\n메인 페이지: {main_page.url[:80]}")

        # 신고내역조회 팝업 열기
        print("\n신고내역 페이지 이동...")
        main_page.goto(RESULT_URL, wait_until="networkidle", timeout=60000)
        time.sleep(3)

        # 신고내역조회 버튼 클릭
        print("신고내역조회 팝업 열기...")
        main_page.click("#mf_txppWframe_btnRtnInqr", timeout=10000)
        time.sleep(3)

        # 팝업 내 모든 버튼/입력 요소 검사
        print("\n=== 팝업 내 input[type=button] ===")
        btns = main_page.evaluate("""() => {
            return Array.from(document.querySelectorAll("input[type=button]")).map(el => ({
                id: el.id,
                value: el.value,
                class: el.className,
                style: el.getAttribute('style') || '',
                visible: el.offsetWidth > 0
            }));
        }""")
        for b in btns:
            print(f"  id={b['id'][:60]} value={b['value']} class={b['class'][:40]}")

        print("\n=== 팝업 내 button ===")
        btns2 = main_page.evaluate("""() => {
            return Array.from(document.querySelectorAll("button")).map(el => ({
                id: el.id,
                text: el.innerText.strip()[:40],
                class: el.className[:60],
            }));
        }""")
        for b in btns2:
            print(f"  id={b['id'][:60]} text={b['text']}")

        print("\n=== select (건수 선택) ===")
        selects = main_page.evaluate("""() => {
            return Array.from(document.querySelectorAll("select")).map(el => ({
                id: el.id,
                options: Array.from(el.options).map(o => o.value + ':' + o.text)
            }));
        }""")
        for s in selects:
            print(f"  id={s['id'][:50]} options={s['options'][:5]}")

        # 빨강 스타일 버튼 찾기
        print("\n=== 빨강 배경 버튼 ===")
        red_btns = main_page.evaluate("""() => {
            const all = Array.from(document.querySelectorAll("input[type=button], button, a"));
            return all.filter(el => {
                const st = window.getComputedStyle(el);
                const bg = st.backgroundColor;
                // 빨강 계열: rgb(2xx, 0~100, 0~100) 또는 class에 red/danger/orange
                const isRed = bg.includes('rgb(') && (() => {
                    const m = bg.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/);
                    if (!m) return false;
                    return parseInt(m[1]) > 150 && parseInt(m[2]) < 100 && parseInt(m[3]) < 100;
                })();
                const cls = (el.className || '').toLowerCase();
                return isRed || cls.includes('red') || cls.includes('danger') || cls.includes('btn_orange') || cls.includes('btn_red');
            }).map(el => ({
                tag: el.tagName,
                id: el.id,
                value: el.value || el.innerText.trim().slice(0,50),
                class: el.className
            }));
        }""")
        for b in red_btns:
            print(f"  {b['tag']} id={b['id']} value={b['value'][:40]} class={b['class'][:40]}")

        print("\n검사 완료")

if __name__ == "__main__":
    inspect()
