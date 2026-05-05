"""지급명세서 팝업 단계별 디버그"""
import sys, os, time
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import (
    REPORT_HELP_URL, normalize_jumin, fill_jumin_and_search,
    wait_preview_button, _click_anneam_tab,
    JIPGUM_BTN_ID, JIPGUM_POPUP_ID, ANNEAM_TAB_ID,
    customer_folder
)
from gsheet_writer import get_credentials
import gspread

NAME = sys.argv[1] if len(sys.argv) > 1 else "배성희"
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"

creds = get_credentials()
gc = gspread.authorize(creds)
ws = gc.open_by_key(SPREADSHEET_ID).worksheet("접수명단")
rows = ws.get_all_records()

customer = None
for r in rows:
    if str(r.get("성명", "")).strip() == NAME:
        customer = {"name": NAME, "jumin_raw": str(r.get("주민번호", "")).strip()}
        break

if not customer:
    print(f"[오류] '{NAME}' 없음")
    sys.exit(1)

front, back = normalize_jumin(customer["jumin_raw"])

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    ctx  = browser.contexts[0]
    page = ctx.pages[0]
    page.bring_to_front()
    page.on("dialog", lambda d: d.dismiss())

    # 신고도움서비스 이동
    page.goto(REPORT_HELP_URL, wait_until="domcontentloaded")
    time.sleep(2)
    fill_jumin_and_search(page, front, back)
    time.sleep(3)

    # 신고 안내자료 탭 클릭
    _click_anneam_tab(page)

    # 지급명세서 버튼 확인
    btn_exists = page.evaluate(f"() => !!document.getElementById('{JIPGUM_BTN_ID}')")
    btn_visible = page.locator(f"#{JIPGUM_BTN_ID}").is_visible(timeout=2000) if btn_exists else False
    print(f"[디버그] 지급명세서 버튼({JIPGUM_BTN_ID}): exists={btn_exists}, visible={btn_visible}")

    if not btn_visible:
        # 버튼 못 찾은 경우 - 신고 안내자료 탭 ID 직접 확인
        tab_exists = page.evaluate(f"() => !!document.getElementById('{ANNEAM_TAB_ID}')")
        print(f"[디버그] 신고안내자료 탭({ANNEAM_TAB_ID}): {tab_exists}")
        triggers = page.evaluate("""
            () => {
                const els = document.querySelectorAll('[id*="trigger"]');
                return Array.from(els).filter(e => e.offsetParent !== null)
                    .map(e => ({id: e.id, text: (e.innerText||e.value||'').slice(0,20)}));
            }
        """)
        print(f"[디버그] 보이는 trigger 버튼들: {triggers}")
        sys.exit(1)

    # 팝업 열기 - locator.click() 사용 (user gesture → 팝업 화면 안에 제대로 렌더)
    print(f"[단계1] 지급명세서 버튼 locator.click()")
    btn_locator = page.locator(f"#{JIPGUM_BTN_ID}")
    btn_locator.click()
    time.sleep(3)

    popup_visible = page.evaluate(f"""
        () => {{ const p = document.getElementById('{JIPGUM_POPUP_ID}'); return !!(p && p.offsetParent); }}
    """)
    print(f"[단계1] 팝업({JIPGUM_POPUP_ID}) visible: {popup_visible}")

    if not popup_visible:
        _chk = page.evaluate(f"() => !!document.getElementById('{JIPGUM_POPUP_ID}')")
        print(f"[디버그] UTERNAAT71 있는지: {_chk}")
        # 팝업 id 검색
        found = page.evaluate("""
            () => {
                const els = document.querySelectorAll('[id*="UTERN"]');
                return Array.from(els).map(e => ({id: e.id, visible: !!e.offsetParent}));
            }
        """)
        print(f"[디버그] UTERN* 요소들: {found}")
        sys.exit(1)

    # 일괄출력 버튼 확인
    batch_btn_id = f"{JIPGUM_POPUP_ID}_wframe_trigger193"
    batch_exists = page.evaluate(f"() => !!document.getElementById('{batch_btn_id}')")
    print(f"[단계2] 일괄출력 버튼({batch_btn_id}): exists={batch_exists}")

    if not batch_exists:
        # trigger 번호 검색
        found = page.evaluate(f"""
            () => {{
                const pop = document.getElementById('{JIPGUM_POPUP_ID}');
                if (!pop) return [];
                return Array.from(pop.querySelectorAll('[id*="trigger"]'))
                    .map(e => ({{id: e.id, text: (e.innerText||e.value||'').slice(0,20), visible: !!e.offsetParent}}));
            }}
        """)
        print(f"[디버그] 팝업 내 trigger 버튼들: {found}")
        sys.exit(1)

    # 체크박스 전체 선택
    print(f"[단계3] 체크박스 전체 선택")
    cb_info = page.evaluate(f"""
        () => {{
            const pop = document.getElementById('{JIPGUM_POPUP_ID}');
            const cbs = pop.querySelectorAll('input[type=checkbox]');
            cbs.forEach(c => {{ c.checked = true; c.dispatchEvent(new MouseEvent('click', {{bubbles: true}})); }});
            return {{count: cbs.length, popText: (pop.innerText||'').slice(0,300)}};
        }}
    """)
    print(f"  체크박스 수: {cb_info.get('count')}")
    print(f"  팝업 텍스트 앞200자: {cb_info.get('popText','')[:200]}")
    time.sleep(1)

    # 일괄출력 버튼 위치 확인
    bbox = page.evaluate(f"""
        () => {{
            const btn = document.getElementById('{batch_btn_id}');
            if (!btn) return null;
            const r = btn.getBoundingClientRect();
            return {{x: r.x, y: r.y, w: r.width, h: r.height, visible: !!btn.offsetParent}};
        }}
    """)
    print(f"[단계4] 일괄출력 버튼 위치: {bbox}")

    before_pages = [pg.url[:60] for pg in ctx.pages]
    print(f"  클릭 전 pages: {before_pages}")

    # viewport 크기 확인
    vp = page.evaluate("() => ({w: window.innerWidth, h: window.innerHeight})")
    print(f"[디버그] viewport: {vp}")

    # 체크박스 실제 상태 확인
    cb_state = page.evaluate(f"""
        () => {{
            const pop = document.getElementById('{JIPGUM_POPUP_ID}');
            const cbs = pop.querySelectorAll('input[type=checkbox]');
            return Array.from(cbs).map(c => {{
                const r = c.getBoundingClientRect();
                return {{id: c.id, checked: c.checked, x: r.x, y: r.y, w: r.width, h: r.height}};
            }});
        }}
    """)
    print(f"[디버그] 체크박스 상태: {cb_state}")

    # 체크박스를 mouse.click으로 실제 클릭 (WebSquare 데이터 바인딩용)
    for cb in cb_state:
        if cb.get('w', 0) > 0:
            cx = cb['x'] + cb['w'] / 2
            cy = cb['y'] + cb['h'] / 2
            print(f"  체크박스 mouse.click({cx:.0f}, {cy:.0f})")
            page.mouse.click(cx, cy)
            time.sleep(0.5)
    time.sleep(1)

    # 체크박스 선택 후 상태 재확인
    cb_state2 = page.evaluate(f"""
        () => {{
            const pop = document.getElementById('{JIPGUM_POPUP_ID}');
            const cbs = pop.querySelectorAll('input[type=checkbox]');
            return Array.from(cbs).map(c => {{return {{id: c.id, checked: c.checked}}}});
        }}
    """)
    print(f"[디버그] 체크박스 클릭 후: {cb_state2}")

    batch_locator = page.locator(f"#{batch_btn_id}")
    page.evaluate(f"""
        () => {{
            const btn = document.getElementById('{batch_btn_id}');
            if (btn) btn.scrollIntoView({{block: 'center', behavior: 'instant'}});
        }}
    """)
    time.sleep(0.5)
    bbox2 = page.evaluate(f"""
        () => {{
            const btn = document.getElementById('{batch_btn_id}');
            if (!btn) return null;
            const r = btn.getBoundingClientRect();
            return {{x: r.x, y: r.y, w: r.width, h: r.height}};
        }}
    """)
    print(f"  스크롤 후 버튼 위치: {bbox2}")

    before_pages = [pg.url[:60] for pg in ctx.pages]
    print(f"  클릭 전 pages: {before_pages}")

    # mouse.click으로 클릭
    if bbox2 and bbox2.get('w', 0) > 0:
        cx = bbox2['x'] + bbox2['w'] / 2
        cy = bbox2['y'] + bbox2['h'] / 2
        print(f"[단계4] 일괄출력 mouse.click({cx:.0f}, {cy:.0f})")
        page.mouse.click(cx, cy)
        time.sleep(7)

    # 모든 ctx 확인
    print(f"  클릭 후 ctx 수: {len(browser.contexts)}")
    for ci, c2 in enumerate(browser.contexts):
        for pj, pg2 in enumerate(c2.pages):
            print(f"  ctx[{ci}] page[{pj}]: {pg2.url[:80]}")

    pdf_popup = None
    for c2 in browser.contexts:
        for pg2 in c2.pages:
            if 'clipreport' in pg2.url.lower():
                pdf_popup = pg2
                break
    print(f"  ClipReport 팝업: {pdf_popup.url[:80] if pdf_popup else 'None'}")
