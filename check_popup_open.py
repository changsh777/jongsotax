"""팝업 열리는지 테스트 — popup blocking 해제 후"""
import requests, json, asyncio, websockets, time

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
    main = next((t for t in tabs
                 if "hometax.go.kr" in t.get("url","")
                 and "websquare.html" in t.get("url","")
                 and "sesw." not in t.get("url","")), None)
    if not main:
        print("홈택스 탭 없음"); return

    known_ids = {t["id"] for t in tabs}
    print(f"기존 탭 수: {len(known_ids)}")

    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 단순 btn.click() — popup blocking 해제됐으면 팝업 열림
        r = await _eval(ws, """(function(){
    var rows = Array.from(document.querySelectorAll('[id*="UTERNAAZ0Z31_wframe"] table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    if (!rows.length) return 'no_rows';
    var cell = rows[0].querySelectorAll('td')[10];
    var btn = cell && cell.querySelector('a, input[type=button], button');
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked:' + (btn.innerText||btn.value||'').trim().slice(0,30);
})()""", cmd_id=1)
        print(f"클릭: {r}")

    # 새 탭 감지 (최대 10초)
    new_tab = None
    for i in range(20):
        await asyncio.sleep(0.5)
        all_tabs = requests.get(f"{CDP}/json").json()
        for t in all_tabs:
            if t["id"] not in known_ids and "devtools" not in t.get("url",""):
                new_tab = t
                break
        if new_tab:
            print(f"\n새 탭 열림 ({i*0.5:.1f}초): {new_tab['url'][:120]}")
            break

    if not new_tab:
        print("새 탭 없음 — 팝업 미열림")
    else:
        # 탭 닫기
        requests.get(f"{CDP}/json/close/{new_tab['id']}")
        print("탭 닫음 — 성공!")

asyncio.run(run())
