"""
debug_ganiyiyong3.py - 팝업 열고 select 옵션 확인 후 log 저장
"""
import sys, io, time
from pathlib import Path

log_path = Path(r"F:\종소세2026\debug_popup.log")
log_lines = []

def log(msg):
    print(msg, flush=True)
    log_lines.append(str(msg))

sys.path.insert(0, r"F:\종소세2026")
from playwright.sync_api import sync_playwright

GANIYIYONG_BTN_ID = "mf_txppWframe_trigger2322"
ANNEAM_TAB_ID     = "mf_txppWframe_tabControl1_tab_tabs3"

def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0]

        log(f"[메인 페이지] {page.url[:100]}")
        log(f"열린 페이지 수: {len(ctx.pages)}")

        # 신고 안내자료 탭 클릭
        try:
            page.evaluate(f"document.getElementById('{ANNEAM_TAB_ID}').click()")
            time.sleep(2)
            log("[탭] 신고 안내자료 탭 클릭")
        except Exception as e:
            log(f"[탭 실패] {e}")

        # 팝업 열기
        log("[팝업] trigger2322 클릭 시도...")
        gp = None
        try:
            with ctx.expect_page(timeout=10000) as pi:
                page.evaluate(f"document.getElementById('{GANIYIYONG_BTN_ID}').click()")
            gp = pi.value
            gp.wait_for_load_state("domcontentloaded", timeout=15000)
            time.sleep(3)
            log(f"[팝업 열림] {gp.url[:120]}")
        except Exception as e:
            log(f"[expect_page 실패] {e}")
            # 이미 열려있나 확인
            for pg in ctx.pages:
                if pg != page:
                    gp = pg
                    log(f"[기존 팝업] {gp.url[:100]}")
                    break

        if not gp:
            log("[오류] 팝업 없음")
            log_path.write_text('\n'.join(log_lines), encoding='utf-8')
            return

        # select 옵션 확인
        log("\n=== select 옵션 ===")
        opts = gp.evaluate("""
            () => {
                const sel = document.getElementById('mf_mateKndCd');
                if (!sel) return '### mf_mateKndCd 없음';
                return Array.from(sel.options)
                    .map((o,i) => i + ': value=[' + o.value + '] text=[' + o.text + ']')
                    .join('\\n');
            }
        """)
        log(opts)

        # 모든 select 목록
        log("\n=== 페이지 내 모든 select ===")
        all_sels = gp.evaluate("""
            () => Array.from(document.querySelectorAll('select'))
                .map(s => 'id=[' + s.id + '] value=[' + s.value + '] opts=' + s.options.length)
                .join('\\n') || '없음'
        """)
        log(all_sels)

        # scwin $w 목록
        log("\n=== scwin 컴포넌트 목록 (select 관련) ===")
        comp_list = gp.evaluate("""
            () => {
                if (!window.scwin || typeof window.scwin.$w !== 'function') return 'scwin.$w 없음';
                // 일반적인 컴포넌트 ID 목록 탐색
                const ids = ['mf_mateKndCd', 'mateKndCd', 'scr_mateKndCd'];
                return ids.map(id => {
                    try {
                        const comp = window.scwin.$w(id);
                        if (!comp) return id + ': null';
                        const val = typeof comp.getValue === 'function' ? comp.getValue() : '?';
                        return id + ': OK val=' + val;
                    } catch(e) { return id + ': err=' + e.message; }
                }).join('\\n');
            }
        """)
        log(comp_list)

        # 버튼 목록
        log("\n=== 버튼 목록 ===")
        btn_list = gp.evaluate("""
            () => {
                const items = Array.from(document.querySelectorAll(
                    'button, input[type=button], [id*=btn], [id*=Btn]'
                ));
                return items.slice(0, 30).map(b =>
                    'id=[' + b.id + '] tag=' + b.tagName +
                    ' text=[' + (b.value||b.textContent||'').trim().slice(0,25) + ']'
                ).join('\\n') || '없음';
            }
        """)
        log(btn_list)

        # 실제 선택 테스트 - 첫 번째 옵션으로
        log("\n=== 첫 번째 옵션 선택 → 조회 테스트 ===")
        first_opt = gp.evaluate("""
            () => {
                const sel = document.getElementById('mf_mateKndCd');
                if (!sel || !sel.options.length) return null;
                for (const o of sel.options) {
                    if (o.text.includes('사업소득')) return {val: o.value, text: o.text};
                }
                return {val: sel.options[0].value, text: sel.options[0].text};
            }
        """)
        log(f"테스트 옵션: {first_opt}")

        if first_opt and isinstance(first_opt, dict):
            opt_val = first_opt.get('val', '')
            # setValue 시도
            set_r = gp.evaluate(f"""
                () => {{
                    // 1) scwin
                    try {{
                        if (window.scwin && typeof window.scwin.$w === 'function') {{
                            const comp = window.scwin.$w('mf_mateKndCd');
                            if (comp && typeof comp.setValue === 'function') {{
                                comp.setValue('{opt_val}');
                                const after = typeof comp.getValue === 'function' ? comp.getValue() : '?';
                                return 'scwin OK after=' + after;
                            }}
                        }}
                    }} catch(e) {{}}
                    // 2) DOM
                    const sel = document.getElementById('mf_mateKndCd');
                    if (!sel) return '없음';
                    sel.value = '{opt_val}';
                    sel.dispatchEvent(new Event('change', {{bubbles:true}}));
                    return 'DOM after=' + sel.value;
                }}
            """)
            log(f"setValue: {set_r}")
            time.sleep(0.5)

            # 조회 버튼 찾기
            inqr_found = gp.evaluate("""
                () => {
                    const ids = ['mf_btnInqr', 'btnInqr'];
                    for (const id of ids) {
                        const b = document.getElementById(id);
                        if (b) return 'found:' + id;
                    }
                    // 텍스트로
                    const btns = Array.from(document.querySelectorAll('button, input[type=button]'));
                    const inqr = btns.find(b => (b.value||b.textContent||'').includes('조회'));
                    return inqr ? 'text-found:' + inqr.id : '없음';
                }
            """)
            log(f"조회 버튼: {inqr_found}")

            if '없음' not in str(inqr_found):
                btn_id = inqr_found.split(':', 1)[1]
                try:
                    gp.locator(f"#{btn_id}").click()
                    log(f"조회 클릭: #{btn_id}")
                    time.sleep(4)
                except Exception as e:
                    log(f"조회 클릭 실패: {e}")

                # 결과 행수
                row_info = gp.evaluate("""
                    () => {
                        const rows = document.querySelectorAll('table tr');
                        const tbls = document.querySelectorAll('table');
                        // 데이터 행 (헤더 제외)
                        const dataTrs = document.querySelectorAll('tbody tr');
                        return 'tables=' + tbls.length + ' rows=' + rows.length + ' tbody_tr=' + dataTrs.length;
                    }
                """)
                log(f"조회 결과: {row_info}")

                # 다운로드 버튼
                dwld_info = gp.evaluate("""
                    () => {
                        const ids = ['mf_btnDwld1', 'mf_btnDwld', 'btnDwld1', 'btnDwld'];
                        for (const id of ids) {
                            const b = document.getElementById(id);
                            if (b) return id + ': disabled=' + b.disabled + ' display=' + b.style.display;
                        }
                        return '다운로드 버튼 없음';
                    }
                """)
                log(f"다운로드 버튼: {dwld_info}")

    log_path.write_text('\n'.join(log_lines), encoding='utf-8')
    log(f"\n[로그 저장] {log_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log_lines.append(f"[예외] {e}")
        log_path.write_text('\n'.join(log_lines), encoding='utf-8')
        raise
