import sys, os, time
sys.path.insert(0, r"F:\종소세2026")
from playwright.sync_api import sync_playwright
from 신규고객처리 import login_hometax_id, logout_hometax

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    ctx = browser.contexts[0]
    page = ctx.pages[0]
    page.bring_to_front()

    logout_hometax(page)
    ok = login_hometax_id(page, "tlswjdtnr69", "ch5470015!")
    print("로그인:", ok)
    if not ok:
        exit(1)
    time.sleep(2)

    # ── 1. 메인페이지에서 신고도움 서비스 클릭해서 URL 확인 ──
    print("\n[1] 메인페이지 LNB에서 신고도움 서비스 클릭 시도")
    page.goto("https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&menuCd=index3",
              wait_until="domcontentloaded")
    time.sleep(3)

    # 신고/납부 상단메뉴 클릭 후 종합소득세 > 신고도움 서비스 찾기
    # JS로 메뉴트리 ID 추출
    menu_data = page.evaluate("""
    () => {
        const result = [];
        // 모든 메뉴 관련 요소에서 ID와 텍스트 추출
        document.querySelectorAll('[id*="menu"], [id*="Menu"], [id*="lnb"]').forEach(el => {
            const txt = (el.innerText || el.textContent || '').trim().slice(0, 60);
            if (txt && (txt.includes('종합소득') || txt.includes('신고도움') || txt.includes('신고안내'))) {
                result.push({id: el.id, tag: el.tagName, txt: txt.slice(0,60)});
            }
        });
        // data-* 속성으로 메뉴ID 있는 것도
        document.querySelectorAll('[data-menu-cd], [data-menucd]').forEach(el => {
            const txt = (el.innerText || el.textContent || '').trim().slice(0, 60);
            if (txt) {
                result.push({menucd: el.dataset.menuCd || el.dataset.menucd, txt: txt.slice(0,30)});
            }
        });
        return result;
    }
    """)
    print(f"  메뉴 데이터: {menu_data[:15]}")

    # ── 2. 직접 신고/납부 > 종합소득세 메뉴 navigate ──
    print("\n[2] 신고도움서비스 직접 URL 탐색")
    # 홈택스 최신 개인납세자 신고도움서비스 URL 후보
    candidates_v2 = [
        # 2024/2025년 종합소득세 신고 임시 서비스
        ("임시서비스", "https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&tmIdx=09&tm2lIdx=0904000000&tm3lIdx=0904010000"),
        # 신고도움서비스 (신규, nts.go.kr로 리다이렉트?)
        ("ntshome세금신고", "https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&tmIdx=06&tm2lIdx=0601000000"),
        # 홈택스 신버전 종합소득세
        ("신고안내w2x", "https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/ra/aa/UTRRABAA001M.xml"),
        # 국세청 종합소득세 신고안내 직접
        ("ntshome직접", "https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/ra/ab/UTRRABAB001M.xml"),
    ]
    for label, url in candidates_v2:
        page.goto(url, wait_until="domcontentloaded")
        time.sleep(2)
        print(f"\n  [{label}]")
        print(f"  URL: {page.url[:120]}")
        txt = page.evaluate("() => document.body.innerText")
        if "로그인 정보가 없습니다" in txt:
            print("  → 접근불가")
            continue
        lines = [l.strip() for l in txt.split('\n') if l.strip() and len(l.strip()) > 2]
        print(f"  텍스트(앞20줄): {lines[:20]}")
        page.screenshot(path=rf"F:\종소세2026\output\v2_{label}.png")

    # ── 3. 로그인 후 실제로 신고/납부 탭 클릭해서 URL 추출 ──
    print("\n[3] 신고/납부 탭 클릭 → 종합소득세 메뉴 추적")
    page.goto("https://hometax.go.kr/websquare/websquare.html?w2xPath=/ui/pp/index_pp.xml&menuCd=index3",
              wait_until="domcontentloaded")
    time.sleep(3)

    # 상단 GNB에서 "신고/납부" 클릭
    try:
        gnb = page.get_by_text("신고/납부", exact=False).first
        if gnb.is_visible(timeout=3000):
            gnb.hover()
            time.sleep(1)
            page.screenshot(path=r"F:\종소세2026\output\gnb_hover.png")
            print("  신고/납부 hover 성공 → 스크린샷 저장")
    except Exception as e:
        print(f"  신고/납부 hover 실패: {e}")

    # LNB에서 종합소득세 > 신고도움 서비스 링크 href 추출
    links = page.evaluate("""
    () => {
        const result = [];
        document.querySelectorAll('a').forEach(el => {
            const txt = (el.innerText || '').trim();
            const href = el.href || '';
            const onclick = el.getAttribute('onclick') || '';
            if (txt && (txt.includes('신고도움') || txt.includes('종합소득세') || txt.includes('신고안내'))) {
                result.push({text: txt.slice(0,40), href: href.slice(0,100), onclick: onclick.slice(0,80)});
            }
        });
        return result;
    }
    """)
    for l in links[:20]:
        print(f"  링크: [{l['text']}] href={l['href']} onclick={l['onclick']}")
