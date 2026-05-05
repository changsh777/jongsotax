"""
test_jipgum_debug.py - 오상연 ClipReport 로딩 디버그

사용법:
  python test_jipgum_debug.py 오상연
"""
import sys, io, time, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import (
    _find_main_page,
    fill_jumin_and_search, normalize_jumin,
    REPORT_HELP_URL, JIPGUM_BTN_ID, ANNEAM_TAB_ID,
    JIPGUM_POPUP_ID,
)
from gsheet_writer import get_credentials
import gspread

GSHEET_ID  = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME = "접수명단"
COL_NAME  = 2
COL_JUMIN = 4

GRID_ID   = f"{JIPGUM_POPUP_ID}_wframe_grdList"
SCWIN_KEY = f"{JIPGUM_POPUP_ID}_wframe_scwin"

def get_jumin_from_sheet(target_name):
    creds = get_credentials()
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(GSHEET_ID).worksheet(SHEET_NAME)
    rows = ws.get_all_values()
    for row in rows[1:]:
        name  = row[COL_NAME].strip()  if len(row) > COL_NAME  else ""
        jumin = row[COL_JUMIN].strip() if len(row) > COL_JUMIN else ""
        if name == target_name and len(jumin.replace("-","")) >= 13:
            return jumin
    return None


def main():
    name = sys.argv[1] if len(sys.argv) > 1 else "오상연"

    print(f"[구글시트] {name} 주민번호 조회 중...")
    jumin = get_jumin_from_sheet(name)
    if not jumin:
        print(f"  주민번호 없음 — 종료"); sys.exit(1)
    print(f"  {name} / {jumin[:6]}-{jumin[6:7]}******")

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = _find_main_page(ctx)
        page.bring_to_front()

        dialogs = []
        def on_dialog(d):
            dialogs.append(d.message)
            print(f"  [DIALOG] {d.message[:100]}")
            d.dismiss()
        page.on("dialog", on_dialog)

        # 조회
        front, back = normalize_jumin(jumin)
        page.goto(REPORT_HELP_URL, wait_until="domcontentloaded", timeout=20000)
        time.sleep(2)
        fill_jumin_and_search(page, front, back)
        time.sleep(4)
        page.evaluate(f"document.getElementById('{ANNEAM_TAB_ID}').click()")
        time.sleep(2)

        # trigger8 → 팝업
        page.locator(f"#{JIPGUM_BTN_ID}").click()
        time.sleep(3)

        popup_ok = page.evaluate(f"""
            () => {{
                const p = document.getElementById('{JIPGUM_POPUP_ID}');
                return p && window.getComputedStyle(p).display !== 'none';
            }}
        """)
        print(f"  팝업 visible: {popup_ok}")

        row_count = page.evaluate(f"() => window['{GRID_ID}']?.getRowCount() || 0")
        print(f"  행 수: {row_count}")

        # 전체선택
        page.evaluate(f"""
            () => {{
                const g = window['{GRID_ID}'];
                const cnt = g?.getRowCount() || 0;
                for (let i = 0; i < cnt; i++) {{
                    try {{ g.setCellChecked(i, 'chk', true); }} catch(e) {{}}
                }}
            }}
        """)
        time.sleep(0.3)

        # 개인정보 공개
        page.evaluate(f"""
            () => {{
                const msk = window['{JIPGUM_POPUP_ID}_wframe_mskApplcYn'];
                if (msk && msk.setValue) msk.setValue('1');
            }}
        """)
        time.sleep(0.3)

        # ── trigger193 locator.click(force=True) + 탭 모니터링 ──
        print(f"\n[trigger193 force click + 탭 모니터링]")
        pages_before = set(id(pg) for pg in ctx.pages)

        try:
            page.locator(f"#{JIPGUM_POPUP_ID}_wframe_trigger193").click(force=True, timeout=3000)
            print(f"  force click OK")
        except Exception as e:
            print(f"  force click 실패: {e}")
            # JS 방식
            page.evaluate(f"""
                () => {{
                    const sc = window['{SCWIN_KEY}'];
                    if (sc && sc.trigger193_onclick_ev) sc.trigger193_onclick_ev();
                    else document.getElementById('{JIPGUM_POPUP_ID}_wframe_trigger193')?.click();
                }}
            """)
            print(f"  JS 클릭 실행")

        # 30초 동안 탭 변화 모니터링
        print(f"\n[탭 변화 모니터링 (30초)]")
        for i in range(60):
            time.sleep(0.5)
            current_pages = ctx.pages
            new_page_ids = set(id(pg) for pg in current_pages) - pages_before

            if new_page_ids or i % 10 == 0:
                for pg in current_pages:
                    url = pg.url
                    print(f"  [{i*0.5:.1f}s] {url[:80]}")

                # ClipReport 탭 찾기
                clip = next((pg for pg in current_pages if 'clipreport' in pg.url.lower()), None)
                if clip:
                    print(f"\n  ★ ClipReport 발견! {clip.url}")
                    break

        print(f"\n[최종 탭 목록]")
        for i, pg in enumerate(ctx.pages):
            print(f"  [{i}] {pg.url[:80]}")


if __name__ == "__main__":
    main()
