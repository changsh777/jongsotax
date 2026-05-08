"""빠른 DOM 검사 — Playwright async로 타임아웃 단축"""
import sys, asyncio, json, time

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

async def main():
    from playwright.async_api import async_playwright

    async with async_playwright() as p:
        print("CDP 연결 시도...")
        try:
            browser = await asyncio.wait_for(
                p.chromium.connect_over_cdp("http://localhost:9222"),
                timeout=60
            )
        except asyncio.TimeoutError:
            print("60초 타임아웃!")
            return

        print("연결 성공!")
        ctx = browser.contexts[0]
        pages = ctx.pages
        print(f"탭 수: {len(pages)}")
        for pg in pages:
            print(f"  {pg.url[:80]}")

        # 홈택스 페이지
        main_page = None
        for pg in pages:
            if "hometax.go.kr" in pg.url:
                main_page = pg
                break

        if not main_page:
            print("홈택스 탭 없음")
            return

        # 신고내역조회 팝업 열기
        RESULT_URL = (
            "https://hometax.go.kr/websquare/websquare.html"
            "?w2xPath=/ui/pp/index_pp.xml"
            "&tmIdx=04&tm2lIdx=0405000000&tm3lIdx=0405040000"
        )
        print(f"\n신고내역 페이지로 이동...")
        await main_page.goto(RESULT_URL, wait_until="networkidle", timeout=60000)
        await asyncio.sleep(3)

        print("신고내역조회 팝업 클릭...")
        await main_page.click("#mf_txppWframe_btnRtnInqr", timeout=10000)
        await asyncio.sleep(3)

        # 모든 버튼 검사
        print("\n=== input[type=button] ===")
        btns = await main_page.evaluate("""() => {
            return Array.from(document.querySelectorAll('input[type=button]')).map(el => ({
                id: el.id || '',
                value: el.value || '',
                cls: (el.className || '').slice(0,60),
                bg: window.getComputedStyle(el).backgroundColor,
                vis: el.offsetParent !== null
            }));
        }""")
        for b in btns:
            flag = " <<< VISIBLE" if b['vis'] else ""
            print(f"  id={b['id']:50s} val={b['value']:20s} bg={b['bg']}{flag}")

        print("\n=== select 건수 ===")
        selects = await main_page.evaluate("""() => {
            return Array.from(document.querySelectorAll('select')).map(el => ({
                id: el.id || '',
                opts: Array.from(el.options).map(o => o.value + ':' + o.text).join(', ')
            }));
        }""")
        for s in selects:
            print(f"  id={s['id']:50s} opts={s['opts']}")

        print("\n=== 빨강/오렌지 계열 버튼 ===")
        red = await main_page.evaluate("""() => {
            const all = Array.from(document.querySelectorAll('input, button, a'));
            return all.filter(el => {
                const bg = window.getComputedStyle(el).backgroundColor;
                const m = bg.match(/rgb\\((\\d+),\\s*(\\d+),\\s*(\\d+)\\)/);
                if (!m) return false;
                const r=+m[1],g=+m[2],b=+m[3];
                return r > 150 && g < 100 && b < 100;
            }).map(el => ({
                tag: el.tagName,
                id: el.id || '',
                value: (el.value || el.innerText || '').trim().slice(0,50),
                cls: (el.className||'').slice(0,50)
            }));
        }""")
        if red:
            for b in red:
                print(f"  {b['tag']} id={b['id']} value={b['value']} cls={b['cls']}")
        else:
            print("  없음 — 오렌지 계열도 확인:")
            orange = await main_page.evaluate("""() => {
                const all = Array.from(document.querySelectorAll('input, button, a'));
                return all.filter(el => {
                    const bg = window.getComputedStyle(el).backgroundColor;
                    const m = bg.match(/rgb\\((\\d+),\\s*(\\d+),\\s*(\\d+)\\)/);
                    if (!m) return false;
                    const r=+m[1],g=+m[2],b=+m[3];
                    return r > 180 && g > 80 && g < 160 && b < 80;
                }).map(el => ({
                    tag: el.tagName, id: el.id||'',
                    value: (el.value||el.innerText||'').trim().slice(0,50),
                    bg: window.getComputedStyle(el).backgroundColor,
                    cls: (el.className||'').slice(0,50)
                }));
            }""")
            for b in orange:
                print(f"  {b['tag']} id={b['id']} val={b['value']} bg={b['bg']}")

asyncio.run(main())
