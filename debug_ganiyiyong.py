"""
debug_ganiyiyong.py - 간이용역 팝업 상태 진단
- CDP 연결 → 팝업 페이지 찾기
- select 옵션값 확인
- 버튼 상태 확인
- 셀렉트 값 설정 후 조회 결과 확인
"""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"F:\종소세2026")
from playwright.sync_api import sync_playwright


def inspect_page(gp):
    print("\n=== 팝업 URL ===")
    print(gp.url)

    print("\n=== mf_mateKndCd 셀렉트 옵션 ===")
    try:
        opts = gp.evaluate("""
            () => {
                const sel = document.getElementById('mf_mateKndCd');
                if (!sel) return '셀렉트 없음';
                return Array.from(sel.options).map(o => `value='${o.value}' text='${o.text}'`).join('\\n');
            }
        """)
        print(opts)
    except Exception as e:
        print(f"  셀렉트 읽기 실패: {e}")

    print("\n=== 현재 선택값 ===")
    try:
        v = gp.evaluate("() => { const s = document.getElementById('mf_mateKndCd'); return s ? s.value : '없음'; }")
        print(v)
    except Exception as e:
        print(f"  실패: {e}")

    print("\n=== 버튼 상태 ===")
    for btn_id in ["mf_btnInqr", "mf_btnDwld1", "mf_btnDwld"]:
        try:
            info = gp.evaluate(f"""
                () => {{
                    const b = document.getElementById('{btn_id}');
                    if (!b) return '{btn_id}: 없음';
                    return '{btn_id}: disabled=' + b.disabled + ' visible=' + (b.style.display !== 'none') + ' class=' + b.className;
                }}
            """)
            print(info)
        except Exception as e:
            print(f"  {btn_id}: {e}")

    print("\n=== scwin API 확인 ===")
    try:
        api = gp.evaluate("""
            () => {
                if (!window.scwin) return 'scwin 없음';
                const comp = window.scwin.$w ? window.scwin.$w('mf_mateKndCd') : null;
                if (!comp) return 'mf_mateKndCd 컴포넌트 없음';
                const methods = Object.getOwnPropertyNames(Object.getPrototypeOf(comp)).join(', ');
                return 'getValue=' + (typeof comp.getValue) + ' setValue=' + (typeof comp.setValue);
            }
        """)
        print(api)
    except Exception as e:
        print(f"  {e}")

    print("\n=== 그리드 행수 ===")
    for grid_id in ["mf_grid1", "mf_Grid1", "grid1", "mf_grdList"]:
        try:
            cnt = gp.evaluate(f"""
                () => {{
                    const g = document.getElementById('{grid_id}');
                    if (!g) return '{grid_id}: 없음';
                    const rows = g.querySelectorAll('tr');
                    return '{grid_id}: rows=' + rows.length;
                }}
            """)
            print(cnt)
        except Exception:
            pass

    print("\n=== 모든 버튼/input 목록 ===")
    try:
        elems = gp.evaluate("""
            () => {
                const btns = Array.from(document.querySelectorAll('button, input[type=button]'));
                return btns.slice(0,20).map(b => `id=${b.id} text=${b.value||b.textContent.trim().slice(0,20)}`).join('\\n');
            }
        """)
        print(elems)
    except Exception as e:
        print(f"  {e}")


def set_select_and_query(gp, type_val):
    print(f"\n=== '{type_val}' 선택 후 조회 테스트 ===")

    # scwin 방식
    result = gp.evaluate(f"""
        () => {{
            try {{
                if (window.scwin && window.scwin.$w) {{
                    const comp = window.scwin.$w('mf_mateKndCd');
                    if (comp && comp.setValue) {{
                        comp.setValue('{type_val}');
                        return 'scwin.setValue 성공';
                    }}
                }}
            }} catch(e) {{ return 'scwin 오류: ' + e.message; }}
            // fallback
            const sel = document.getElementById('mf_mateKndCd');
            if (!sel) return 'select 없음';
            sel.value = '{type_val}';
            sel.dispatchEvent(new Event('change', {{bubbles:true}}));
            return 'DOM fallback: value=' + sel.value;
        }}
    """)
    print(f"  setValue 결과: {result}")

    time.sleep(0.5)
    after = gp.evaluate("() => { const s = document.getElementById('mf_mateKndCd'); return s ? s.value : '없음'; }")
    print(f"  설정 후 현재값: {after}")

    # 조회 버튼 클릭
    try:
        gp.locator("#mf_btnInqr").click()
        print("  조회 버튼 클릭 완료")
        time.sleep(3)
    except Exception as e:
        print(f"  조회 버튼 오류: {e}")

    # 결과 확인
    try:
        rows = gp.evaluate("""
            () => {
                // 다양한 방식으로 행 수 탐지
                const tables = document.querySelectorAll('table');
                let info = [];
                tables.forEach((t, i) => {
                    const rows = t.querySelectorAll('tr');
                    if (rows.length > 0) info.push(`table[${i}] rows=${rows.length}`);
                });
                return info.join(' | ') || '테이블 없음';
            }
        """)
        print(f"  조회 후 테이블: {rows}")
    except Exception as e:
        print(f"  테이블 확인 실패: {e}")

    # 다운로드 버튼 상태
    try:
        dwld_state = gp.evaluate("""
            () => {
                const b = document.getElementById('mf_btnDwld1');
                if (!b) return '없음';
                return 'disabled=' + b.disabled + ' display=' + b.style.display;
            }
        """)
        print(f"  다운로드 버튼 상태: {dwld_state}")
    except Exception as e:
        print(f"  버튼 상태 실패: {e}")


def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]

        print(f"열린 페이지 수: {len(ctx.pages)}")
        for i, pg in enumerate(ctx.pages):
            print(f"  [{i}] {pg.url[:100]}")

        # 간이용역 팝업 찾기
        gp = None
        for pg in ctx.pages:
            url = pg.url
            if 'UTERNATB35' in url or 'ganiyiyong' in url.lower() or ('popup' in url and 'TEWEA' in url):
                gp = pg
                break
        if not gp:
            # URL로 특정 못하면 마지막 팝업
            for pg in reversed(ctx.pages):
                if pg.url not in ('about:blank', '') and 'hometax' in pg.url:
                    gp = pg
                    break

        if not gp:
            print("\n[오류] 간이용역 팝업을 찾지 못했습니다.")
            print("팝업이 열려있는지 확인하세요.")
            return

        print(f"\n[사용 페이지] {gp.url[:120]}")
        inspect_page(gp)

        # 실제 옵션값으로 테스트 (첫 번째 옵션)
        first_val = gp.evaluate("""
            () => {
                const sel = document.getElementById('mf_mateKndCd');
                if (!sel || !sel.options.length) return '';
                return sel.options[0].value;
            }
        """)
        if first_val:
            print(f"\n첫 번째 옵션값 '{first_val}'으로 테스트:")
            set_select_and_query(gp, first_val)


if __name__ == "__main__":
    main()
