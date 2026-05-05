"""
debug_ganiyiyong2.py - 팝업 열기 + 내부 select/버튼 완전 진단
- 현재 메인 페이지에서 trigger2322 클릭 → 팝업 열기
- 팝업 내 select 옵션값, 버튼 ID, scwin 컴포넌트 목록 출력
"""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"F:\종소세2026")
from playwright.sync_api import sync_playwright

GANIYIYONG_BTN_ID = "mf_txppWframe_trigger2322"
ANNEAM_TAB_ID     = "mf_txppWframe_tabControl1_tab_tabs3"


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0]

        print(f"[메인 페이지] {page.url[:100]}")

        # 신고 안내자료 탭 클릭
        try:
            page.evaluate(f"document.getElementById('{ANNEAM_TAB_ID}').click()")
            time.sleep(2)
            print("[탭 클릭] 신고 안내자료 탭")
        except Exception as e:
            print(f"[탭 클릭 실패] {e}")

        # trigger2322 버튼 찾기
        btn = page.locator(f"#{GANIYIYONG_BTN_ID}")
        if not btn.is_visible(timeout=3000):
            print(f"[오류] {GANIYIYONG_BTN_ID} 버튼이 안 보임")
            # 다른 버튼 ID 시도
            btn2 = page.locator("#mf_txppWframe_trigger23221")
            if btn2.is_visible(timeout=2000):
                print("  → trigger23221 사용")
                btn = btn2
            else:
                return

        # 팝업 열기
        print("[팝업 열기 시도...]")
        try:
            with ctx.expect_page(timeout=10000) as page_info:
                btn.click()
            gp = page_info.value
            gp.wait_for_load_state("domcontentloaded", timeout=15000)
            time.sleep(3)
            print(f"[팝업 열림] {gp.url[:120]}")
        except Exception as e:
            print(f"[expect_page 실패] {e}")
            # 이미 열린 팝업 찾기
            pages_now = ctx.pages
            print(f"  현재 페이지 수: {len(pages_now)}")
            for i, pg in enumerate(pages_now):
                print(f"  [{i}] {pg.url[:100]}")
            gp = next((pg for pg in pages_now if pg != page), None)
            if not gp:
                print("  팝업 없음 - 종료")
                return

        print("\n=== [팝업 내부 진단] ===")

        # 1. select 옵션
        print("\n▶ mf_mateKndCd select 옵션:")
        try:
            opts = gp.evaluate("""
                () => {
                    const sel = document.getElementById('mf_mateKndCd');
                    if (!sel) return '없음';
                    const cur = sel.value;
                    const list = Array.from(sel.options).map(
                        o => `  value='${o.value}' | text='${o.text}' | selected=${o.selected}`
                    ).join('\\n');
                    return `현재값: '${cur}'\\n` + list;
                }
            """)
            print(opts)
        except Exception as e:
            print(f"  실패: {e}")

        # 2. 모든 select 목록
        print("\n▶ 페이지 내 모든 select:")
        try:
            sels = gp.evaluate("""
                () => Array.from(document.querySelectorAll('select'))
                    .map(s => `id='${s.id}' value='${s.value}' options=${s.options.length}`)
                    .join('\\n') || '없음'
            """)
            print(sels)
        except Exception as e:
            print(f"  실패: {e}")

        # 3. 버튼 목록
        print("\n▶ 버튼/input[button] 목록:")
        try:
            btns = gp.evaluate("""
                () => {
                    const items = Array.from(document.querySelectorAll(
                        'button, input[type=button], a[id*=btn], span[id*=btn]'
                    ));
                    return items.slice(0, 30).map(b =>
                        `id='${b.id}' tag=${b.tagName} text='${(b.value||b.textContent||'').trim().slice(0,25)}'`
                    ).join('\\n') || '없음';
                }
            """)
            print(btns)
        except Exception as e:
            print(f"  실패: {e}")

        # 4. scwin 컴포넌트
        print("\n▶ window.scwin 상태:")
        try:
            scwin = gp.evaluate("""
                () => {
                    if (!window.scwin) return 'scwin 없음';
                    const hasW = typeof window.scwin.$w === 'function';
                    return 'scwin 있음, $w=' + hasW;
                }
            """)
            print(scwin)
        except Exception as e:
            print(f"  실패: {e}")

        # 5. scwin으로 셀렉트 조작 시도
        print("\n▶ scwin.$w('mf_mateKndCd') 시도:")
        try:
            comp_info = gp.evaluate("""
                () => {
                    if (!window.scwin || typeof window.scwin.$w !== 'function') return 'scwin.$w 없음';
                    const comp = window.scwin.$w('mf_mateKndCd');
                    if (!comp) return 'comp=null';
                    const methods = Object.getOwnPropertyNames(Object.getPrototypeOf(comp))
                        .filter(m => m.includes('et') || m.includes('alue')).join(', ');
                    const val = typeof comp.getValue === 'function' ? comp.getValue() : '?';
                    return `OK: val=${val} methods=${methods}`;
                }
            """)
            print(comp_info)
        except Exception as e:
            print(f"  실패: {e}")

        # 6. 실제 옵션값으로 테스트
        print("\n▶ 첫 번째 사업소득 옵션값으로 setValue + 조회 테스트:")
        first_opt = gp.evaluate("""
            () => {
                const sel = document.getElementById('mf_mateKndCd');
                if (!sel || !sel.options.length) return null;
                // '사업소득' 포함 옵션 찾기
                for (const o of sel.options) {
                    if (o.text.includes('사업소득')) return {val: o.value, text: o.text};
                }
                return {val: sel.options[0].value, text: sel.options[0].text};
            }
        """)
        print(f"  대상 옵션: {first_opt}")

        if first_opt:
            opt_val = first_opt['val'] if isinstance(first_opt, dict) else None
            if opt_val:
                # scwin setValue
                set_result = gp.evaluate(f"""
                    () => {{
                        // 1) scwin
                        try {{
                            if (window.scwin && typeof window.scwin.$w === 'function') {{
                                const comp = window.scwin.$w('mf_mateKndCd');
                                if (comp && typeof comp.setValue === 'function') {{
                                    comp.setValue('{opt_val}');
                                    const after = comp.getValue ? comp.getValue() : '?';
                                    return 'scwin setValue OK, after=' + after;
                                }}
                            }}
                        }} catch(e) {{ return 'scwin 오류: ' + e.message; }}
                        // 2) DOM
                        const sel = document.getElementById('mf_mateKndCd');
                        if (!sel) return 'select 없음';
                        sel.value = '{opt_val}';
                        sel.dispatchEvent(new Event('change', {{bubbles:true}}));
                        return 'DOM 설정 후 val=' + sel.value;
                    }}
                """)
                print(f"  setValue 결과: {set_result}")

                # 조회 클릭
                time.sleep(0.5)
                try:
                    # 조회 버튼 찾기
                    inqr_ids = ["mf_btnInqr", "btnInqr", "mf_wfMenu_trigger9"]
                    inqr_btn = None
                    for bid in inqr_ids:
                        el = gp.locator(f"#{bid}")
                        if el.count() > 0:
                            try:
                                if el.is_visible(timeout=1000):
                                    inqr_btn = el
                                    print(f"  조회 버튼: #{bid}")
                                    break
                            except Exception:
                                pass
                    if not inqr_btn:
                        # 텍스트로 찾기
                        inqr_btn = gp.locator("text=조회하기, text=조회").first
                        print(f"  조회 버튼: 텍스트 검색")

                    if inqr_btn:
                        inqr_btn.click()
                        print("  조회 버튼 클릭")
                        time.sleep(4)

                        # 테이블 행수
                        tbl = gp.evaluate("""
                            () => {
                                const rows = document.querySelectorAll('table tr');
                                const tbls = document.querySelectorAll('table');
                                return `테이블수=${tbls.length} 총행수=${rows.length}`;
                            }
                        """)
                        print(f"  조회 결과: {tbl}")

                        # 다운로드 버튼 상태
                        for dwld_id in ["mf_btnDwld1", "mf_btnDwld", "btnDwld1", "btnDwld"]:
                            try:
                                info = gp.evaluate(f"""
                                    () => {{
                                        const b = document.getElementById('{dwld_id}');
                                        if (!b) return null;
                                        return '{dwld_id}: disabled=' + b.disabled +
                                               ' display=' + b.style.display +
                                               ' class=' + b.className.slice(0,40);
                                    }}
                                """)
                                if info:
                                    print(f"  {info}")
                            except Exception:
                                pass
                except Exception as e:
                    print(f"  조회 실패: {e}")

        print("\n=== 진단 완료 ===")
        input("팝업 확인 후 Enter 키 (팝업을 닫지 마세요)")


if __name__ == "__main__":
    main()
