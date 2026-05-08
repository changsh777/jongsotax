"""신고내역조회 팝업 테이블 상세 구조 검사"""
import sys, json, asyncio, requests
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import websockets

async def main():
    tabs = requests.get("http://localhost:9222/json").json()
    ht_tab = next((t for t in tabs if "hometax.go.kr" in t.get("url","")), None)
    if not ht_tab:
        print("홈택스 탭 없음!"); return

    ws_url = ht_tab["webSocketDebuggerUrl"]
    print(f"연결: {ws_url[:60]}")

    async with websockets.connect(ws_url) as ws:
        print("연결 성공!")

        async def eval_js(code, cmd_id=1):
            await ws.send(json.dumps({
                "id": cmd_id, "method": "Runtime.evaluate",
                "params": {"expression": code, "returnByValue": True}
            }))
            while True:
                resp = json.loads(await asyncio.wait_for(ws.recv(), timeout=20))
                if resp.get("id") == cmd_id:
                    r = resp.get("result", {})
                    val = r.get("result", {}).get("value")
                    return json.loads(val) if isinstance(val, str) else val

        # 1. 신고내역 조회 테이블 행 구조
        rows_info = await eval_js("""(function() {
    var table = document.querySelector('table');
    if (!table) return {error: 'no table'};
    var rows = Array.from(table.querySelectorAll('tbody tr'));
    var result = [];
    rows.forEach(function(tr, rIdx) {
        var tds = Array.from(tr.querySelectorAll('td'));
        if (tds.length < 3) return;
        var cells = tds.map(function(td, cIdx) {
            var btn = td.querySelector('input[type=button], button, a');
            return {
                text: td.innerText.trim().slice(0,20),
                btn: btn ? {
                    tag: btn.tagName,
                    id: btn.id || '',
                    val: (btn.value || btn.innerText || '').trim().slice(0,20),
                    cls: (btn.className || '').slice(0,40)
                } : null
            };
        });
        result.push({row: rIdx, cellCount: tds.length, cells: cells});
    });
    return result;
})()""", cmd_id=1)

        if rows_info and isinstance(rows_info, list):
            for r in rows_info[:3]:  # 처음 3행만
                print(f"\n--- 행 {r['row']} ({r['cellCount']}컬럼) ---")
                for ci, c in enumerate(r.get('cells', [])):
                    if c.get('btn') or c.get('text'):
                        btn_info = f" [BTN: tag={c['btn']['tag']} id={c['btn']['id'][:40]} val={c['btn']['val'][:20]} cls={c['btn']['cls'][:40]}]" if c['btn'] else ""
                        print(f"  col[{ci:2d}] text={c['text']:20s}{btn_info}")
        else:
            print(f"rows_info: {rows_info}")

        # 2. UTERNAAZ0Z31 내부 테이블 검사 (팝업 컨테이너)
        print("\n\n=== 신고내역 팝업 내부 테이블 ===")
        popup_rows = await eval_js("""(function() {
    // UTERNAAZ0Z31 관련 요소 찾기
    var container = document.querySelector('[id*="UTERNAAZ0Z31"]');
    if (!container) return {error: 'no UTERNAAZ0Z31 container'};

    var tables = Array.from(container.querySelectorAll('table'));
    if (!tables.length) return {error: 'no table in popup'};

    // 데이터 행이 가장 많은 테이블 찾기
    var maxTable = tables.reduce(function(a, b) {
        return a.querySelectorAll('tbody tr').length > b.querySelectorAll('tbody tr').length ? a : b;
    });

    var rows = Array.from(maxTable.querySelectorAll('tbody tr'));
    var result = [];
    rows.slice(0,2).forEach(function(tr, rIdx) {
        var tds = Array.from(tr.querySelectorAll('td'));
        var cells = tds.map(function(td, cIdx) {
            var btn = td.querySelector('input[type=button], button, a');
            return {
                col: cIdx,
                text: td.innerText.trim().slice(0,25),
                btn: btn ? {
                    tag: btn.tagName,
                    id: btn.id || '(no-id)',
                    val: (btn.value || btn.innerText || '').trim().slice(0,20),
                    onclick: (btn.getAttribute('onclick') || '').slice(0,60)
                } : null
            };
        });
        result.push({row: rIdx, cols: tds.length, cells: cells});
    });
    return result;
})()""", cmd_id=2)

        if popup_rows and isinstance(popup_rows, list):
            for r in popup_rows:
                print(f"\n--- 팝업 행 {r['row']} ({r['cols']}컬럼) ---")
                for c in r.get('cells', []):
                    if c.get('btn') or c.get('text'):
                        btn_str = ""
                        if c['btn']:
                            b = c['btn']
                            btn_str = f" [BTN val={b['val']:15s} id={b['id'][:30]:30s} onclick={b['onclick'][:40]}]"
                        print(f"  col[{c['col']:2d}] {c['text']:25s}{btn_str}")
        else:
            print(f"결과: {popup_rows}")

asyncio.run(main())
