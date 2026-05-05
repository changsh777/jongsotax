import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, r"F:\종소세2026")
from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    ctx = browser.contexts[0]

    # 탭1 = 신고도움서비스
    page1 = ctx.pages[1]
    print("탭1 URL:", page1.url[:150])

    # 모든 버튼/링크 중 미리보기 관련 찾기
    els = page1.evaluate("""
    () => {
        const items = [];
        document.querySelectorAll('*').forEach(el => {
            const txt = (el.innerText || el.value || el.textContent || '').trim();
            const onclick = el.getAttribute('onclick') || '';
            if (txt.includes('미리보기') || onclick.includes('미리보기') || onclick.includes('preview') || onclick.includes('clipreport')) {
                items.push({
                    tag: el.tagName,
                    id: el.id.slice(0,70),
                    txt: txt.slice(0,40),
                    onclick: onclick.slice(0,100)
                });
            }
        });
        return items;
    }
    """)
    print("\n미리보기 버튼 요소:")
    for e in els[:20]:
        print(f"  [{e['txt']}] id={e['id']} onclick={e['onclick']}")

    # 스크린샷
    page1.bring_to_front()
    page1.screenshot(path=r"F:\종소세2026\output\tab1_shinbub.png")
    print("\n탭1 스크린샷 저장")

    # 탭2 = clipreport
    page2 = ctx.pages[2]
    print("\n탭2 URL:", page2.url[:150])
    page2.bring_to_front()
    time.sleep(1)
    page2.screenshot(path=r"F:\종소세2026\output\tab2_clipreport.png")
    print("탭2 스크린샷 저장")
