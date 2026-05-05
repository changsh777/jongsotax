"""pages[0] (UWEICAAD32 팝업) 내부 구조 완전 덤프"""
import sys, io, time
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, r"F:\종소세2026")
from playwright.sync_api import sync_playwright

log_path = Path(r"F:\종소세2026\debug_popup2.log")
out = []

def L(msg): print(msg, flush=True); out.append(str(msg))

def dump_page(pg, label):
    L(f"\n{'='*60}")
    L(f"[{label}] URL: {pg.url}")
    L(f"{'='*60}")

    # 모든 select
    L("▶ 모든 SELECT:")
    try:
        sels = pg.evaluate("""
            () => Array.from(document.querySelectorAll('select'))
                .map(s => {
                    const opts = Array.from(s.options)
                        .map(o => `    val='${o.value}' txt='${o.text}'`).join('\\n');
                    return 'id=['+s.id+'] curVal=['+s.value+'] opts='+s.options.length+'\\n'+opts;
                }).join('\\n---\\n') || '없음'
        """)
        L(sels)
    except Exception as e:
        L(f"  실패: {e}")

    # 모든 input/button
    L("\n▶ 버튼 (최대 30개):")
    try:
        btns = pg.evaluate("""
            () => Array.from(document.querySelectorAll(
                'button, input[type=button], input[type=submit]'
            )).slice(0,30)
            .map(b => 'id=['+b.id+'] val=['+b.value+'] dis='+b.disabled)
            .join('\\n') || '없음'
        """)
        L(btns)
    except Exception as e:
        L(f"  실패: {e}")

    # 페이지 title
    L("\n▶ 페이지 title:")
    try:
        title = pg.title()
        L(f"  {title}")
    except Exception as e:
        L(f"  {e}")

    # h1/h2/h3 등 헤딩
    L("\n▶ 헤딩 텍스트:")
    try:
        heads = pg.evaluate("""
            () => Array.from(document.querySelectorAll('h1,h2,h3,h4,.title,.tit'))
                .slice(0,10).map(h => h.tagName+': '+h.textContent.trim().slice(0,50))
                .join('\\n') || '없음'
        """)
        L(heads)
    except Exception as e:
        L(f"  {e}")

    # scwin 상태
    L("\n▶ scwin 상태:")
    try:
        sw = pg.evaluate("""
            () => {
                if (!window.scwin) return 'scwin 없음';
                const hasW = typeof window.scwin.$w;
                return 'scwin 있음, $w type=' + hasW;
            }
        """)
        L(sw)
    except Exception as e:
        L(f"  {e}")

def main():
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]

        L(f"총 페이지 수: {len(ctx.pages)}")
        for i, pg in enumerate(ctx.pages):
            L(f"  pages[{i}]: {pg.url[:100]}")

        # 모든 페이지 덤프
        for i, pg in enumerate(ctx.pages):
            dump_page(pg, f"pages[{i}]")

    log_path.write_text('\n'.join(out), encoding='utf-8')
    L(f"\n[저장] {log_path}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        out.append(f"[예외] {e}")
        log_path.write_text('\n'.join(out), encoding='utf-8')
        raise
