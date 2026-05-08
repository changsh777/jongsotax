"""
접수증 버튼 클릭 테스트 — 페이지레벨 CDP
결과: 클릭 후 새 탭 생기는지 HTTP 폴링으로 확인
"""
import sys, json, asyncio, requests, time
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import websockets

async def main():
    tabs_before = {t["id"]: t for t in requests.get("http://localhost:9222/json").json()}
    print(f"클릭 전 탭 수: {len(tabs_before)}")

    ht_tab = next((t for t in tabs_before.values() if "hometax.go.kr" in t.get("url","")), None)
    if not ht_tab:
        print("홈택스 탭 없음!"); return

    async with websockets.connect(ht_tab["webSocketDebuggerUrl"]) as ws:
        print("CDP 연결 성공!")

        async def eval_js(code, cmd_id=1):
            await ws.send(json.dumps({
                "id": cmd_id, "method": "Runtime.evaluate",
                "params": {"expression": code, "returnByValue": True}
            }))
            while True:
                resp = json.loads(await asyncio.wait_for(ws.recv(), timeout=30))
                if resp.get("id") == cmd_id:
                    return resp.get("result", {}).get("result", {}).get("value")

        # 현재 팝업 내 첫 번째 데이터 행의 col[12] 보기 버튼 찾기
        btn_info = await eval_js("""(function() {
    // UTERNAAZ0Z31 컨테이너 내 테이블 tbody tr (데이터 행)
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'));
    // col[12] 버튼 찾기
    var dataRows = rows.filter(function(tr) { return tr.querySelectorAll('td').length >= 13; });
    if (!dataRows.length) return 'no data rows';
    var firstRow = dataRows[0];
    var tds = Array.from(firstRow.querySelectorAll('td'));
    var btn = tds[12] ? tds[12].querySelector('input[type=button]') : null;
    if (!btn) return 'no button at col 12';
    return {
        found: true,
        id: btn.id || '(no-id)',
        val: btn.value || '',
        cls: btn.className || '',
        rect: JSON.stringify(btn.getBoundingClientRect()),
        rowName: tds[6] ? tds[6].innerText.trim() : '?'
    };
})()""", cmd_id=1)

        print(f"\n첫 행 col[12] 버튼: {btn_info}")

        if not btn_info or btn_info in ('no container', 'no data rows', 'no button at col 12'):
            print("버튼을 찾을 수 없음! 팝업이 열려있는지 확인하세요.")
            return

        # 클릭 전 상태 저장
        print("\n클릭 전 탭 목록:", list(tabs_before.keys()))

        # 버튼 클릭 (JS로)
        click_result = await eval_js("""(function() {
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'));
    var dataRows = rows.filter(function(tr) { return tr.querySelectorAll('td').length >= 13; });
    if (!dataRows.length) return 'no data rows';
    var firstRow = dataRows[0];
    var tds = Array.from(firstRow.querySelectorAll('td'));
    var btn = tds[12] ? tds[12].querySelector('input[type=button]') : null;
    if (!btn) return 'no button';
    btn.click();
    return 'clicked: ' + (btn.value || '?');
})()""", cmd_id=2)
        print(f"클릭 결과: {click_result}")

        # 새 탭 폴링 (10초 동안)
        print("\n새 탭 대기 중...")
        for i in range(20):
            await asyncio.sleep(0.5)
            tabs_after = {t["id"]: t for t in requests.get("http://localhost:9222/json").json()}
            new_tabs = {k: v for k, v in tabs_after.items() if k not in tabs_before}
            if new_tabs:
                print(f"  새 탭 발견! ({i*0.5:.1f}초)")
                for tid, t in new_tabs.items():
                    print(f"    id={tid} url={t.get('url','')[:80]}")
                break
            if i % 4 == 0:
                print(f"  {i*0.5:.1f}초 대기 중...")
        else:
            print("  10초 후에도 새 탭 없음")

asyncio.run(main())
