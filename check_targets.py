"""CDP Target.getTargets로 모든 타깃 확인 + 팝업 직접 테스트"""
import requests, json, asyncio, websockets

CDP = "http://localhost:9222"

async def _eval(ws, code, cmd_id=1):
    await ws.send(json.dumps({"id": cmd_id, "method": "Runtime.evaluate",
                              "params": {"expression": code, "returnByValue": True}}))
    while True:
        r = json.loads(await asyncio.wait_for(ws.recv(), timeout=15))
        if r.get("id") == cmd_id:
            return r.get("result", {}).get("result", {}).get("value")

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    main = next((t for t in tabs if "hometax.go.kr" in t.get("url","")
                 and "websquare.html" in t.get("url","")
                 and "sesw." not in t.get("url","")), None)
    if not main:
        print("홈택스 탭 없음"); return

    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 1. Target.getTargets — 모든 타깃 확인
        await ws.send(json.dumps({"id": 50, "method": "Target.getTargets", "params": {}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 50:
                targets = r.get("result", {}).get("targetInfos", [])
                print(f"\n현재 타깃 {len(targets)}개:")
                for t in targets:
                    print(f"  [{t.get('type')}] {t.get('url','')[:80]}")
                break

        # 2. window.open 직접 테스트 — popup blocking 실제 확인
        r2 = await _eval(ws, """(function(){
    var w = window.open('about:blank', '_popup_test', 'width=400,height=300');
    if (!w) return 'BLOCKED';
    w.close();
    return 'ALLOWED';
})()""", cmd_id=2)
        print(f"\nwindow.open 팝업차단 테스트: {r2}")

        # 3. btn.click() 후 Target.getTargets
        known = set()
        await ws.send(json.dumps({"id": 51, "method": "Target.getTargets", "params": {}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 51:
                known = {t["targetId"] for t in r.get("result", {}).get("targetInfos", [])}
                break

        r3 = await _eval(ws, """(function(){
    var rows = Array.from(document.querySelectorAll('[id*="UTERNAAZ0Z31_wframe"] table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    var cell = rows[0] && rows[0].querySelectorAll('td')[10];
    var btn = cell && cell.querySelector('a, input[type=button], button');
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked';
})()""", cmd_id=3)
        print(f"버튼 클릭: {r3}")

        await asyncio.sleep(3)

        await ws.send(json.dumps({"id": 52, "method": "Target.getTargets", "params": {}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 52:
                new_targets = [t for t in r.get("result", {}).get("targetInfos", [])
                               if t["targetId"] not in known]
                print(f"\n새 타깃 {len(new_targets)}개:")
                for t in new_targets:
                    print(f"  [{t.get('type')}] {t.get('url','')[:100]}")
                if not new_targets:
                    print("  (없음)")
                break

asyncio.run(run())
