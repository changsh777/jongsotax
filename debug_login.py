import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"F:\종소세2026")
from playwright.sync_api import sync_playwright

HOMETAX_MAIN_URL = (
    "https://www.hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
)

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    ctx = browser.contexts[0]

    print("전체 탭:")
    for i, pg in enumerate(ctx.pages):
        print(f"  탭{i}: {pg.url[:120]}")

    page = ctx.pages[0]
    page.bring_to_front()

    # 현재 상태
    page.screenshot(path=r"F:\종소세2026\output\debug_before.png")
    print(f"\n[현재] URL: {page.url[:120]}")

    # 홈택스 메인으로 이동
    print("\n[GOTO 메인]")
    page.goto(HOMETAX_MAIN_URL, wait_until="domcontentloaded")
    time.sleep(3)
    page.screenshot(path=r"F:\종소세2026\output\debug_main.png")
    print(f"메인URL: {page.url[:120]}")

    # 로그인 관련 요소 확인
    login_els = page.evaluate("""
    () => {
        const items = [];
        ['#mf_txppWframe_loginboxFrame_trigger2',
         '#mf_txppWframe_iptUserId',
         '#mf_txppWframe_iptUserPw',
         '#mf_wfHeader_group1503'].forEach(sel => {
            const el = document.querySelector(sel);
            items.push({
                sel: sel,
                exists: !!el,
                visible: el ? el.offsetParent !== null : false,
                txt: el ? (el.innerText || el.value || '').trim().slice(0,30) : ''
            });
        });
        return items;
    }
    """)
    print("\n로그인 요소 상태:")
    for e in login_els:
        print(f"  {e['sel']}: exists={e['exists']}, visible={e['visible']}, txt={e['txt']}")

    # 로그인박스 trigger2 클릭
    print("\n[아이디로그인 버튼 클릭 시도]")
    try:
        btn1 = page.locator("#mf_txppWframe_loginboxFrame_trigger2")
        vis = btn1.is_visible(timeout=3000)
        print(f"  trigger2 visible: {vis}")
        if vis:
            btn1.click()
            time.sleep(2)
            page.screenshot(path=r"F:\종소세2026\output\debug_after_click1.png")
            print("  클릭 완료")
    except Exception as e:
        print(f"  오류: {e}")

    # 아이디 로그인 탭 클릭
    print("\n[아이디 로그인 탭 시도]")
    try:
        tab = page.get_by_text("아이디 로그인", exact=True).first
        vis = tab.is_visible(timeout=3000)
        print(f"  탭 visible: {vis}")
        if vis:
            tab.click()
            time.sleep(1)
    except Exception as e:
        print(f"  오류: {e}")

    # ID 입력란 상태
    print("\n[ID 입력란 확인]")
    try:
        id_el = page.locator("#mf_txppWframe_iptUserId")
        vis = id_el.is_visible(timeout=5000)
        print(f"  iptUserId visible: {vis}")
    except Exception as e:
        print(f"  오류: {e}")

    page.screenshot(path=r"F:\종소세2026\output\debug_final.png")
    print("\n디버그 스크린샷 저장 완료")
