"""UTERNAAZ34 신고서 팝업 구조 확인 — Input.dispatchMouseEvent 방식"""
import requests, json, asyncio, websockets

CDP = "http://localhost:9222"

async def _eval(ws, code, cmd_id=1):
    await ws.send(json.dumps({
        "id": cmd_id, "method": "Runtime.evaluate",
        "params": {"expression": code, "returnByValue": True}
    }))
    while True:
        r = json.loads(await asyncio.wait_for(ws.recv(), timeout=15))
        if r.get("id") == cmd_id:
            v = r.get("result", {}).get("result", {})
            return v.get("value")

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    main = next((t for t in tabs
                 if "hometax.go.kr" in t.get("url", "")
                 and "websquare.html" in t.get("url", "")
                 and "sesw." not in t.get("url", "")), None)
    if not main:
        print("홈택스 탭 없음"); return

    known_ids = {t["id"] for t in tabs}
    print(f"메인 탭: {main['url'][:80]}")

    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 1. 버튼 좌표 가져오기
        coords = await _eval(ws, """(function(){
    var rows = Array.from(document.querySelectorAll('[id*="UTERNAAZ0Z31_wframe"] table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    if (!rows.length) return JSON.stringify({err:'no_rows'});
    var cell = rows[0].querySelectorAll('td')[10];
    var btn = cell && cell.querySelector('a, input[type=button], button');
    if (!btn) return JSON.stringify({err:'no_btn'});
    var r = btn.getBoundingClientRect();
    return JSON.stringify({x: r.left + r.width/2, y: r.top + r.height/2,
        text:(btn.innerText||btn.value||'').trim().slice(0,30)});
})()""", cmd_id=2)

        import json as _j
        data = _j.loads(coords) if isinstance(coords, str) else coords
        if not data or 'err' in data:
            print(f"버튼 좌표 오류: {data}"); return

        x, y = data['x'], data['y']
        print(f"버튼 좌표: ({x:.0f}, {y:.0f}) 텍스트: {data.get('text','')}")

        # 2. Input.dispatchMouseEvent — 진짜 마우스 클릭
        for ev in ("mouseMoved", "mousePressed", "mouseReleased"):
            await ws.send(json.dumps({
                "id": 10, "method": "Input.dispatchMouseEvent",
                "params": {"type": ev, "x": x, "y": y,
                           "button": "left", "clickCount": 1,
                           "modifiers": 0}
            }))
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
        print("마우스 클릭 전송 완료")

        # 3. 새 탭 대기 (WS 유지하면서)
        new_tab = None
        print("새 탭 대기 중 (최대 20초)...")
        for i in range(40):
            await asyncio.sleep(0.5)
            all_tabs = requests.get(f"{CDP}/json").json()
            for t in all_tabs:
                url = t.get("url", "")
                if t["id"] not in known_ids and "devtools" not in url:
                    new_tab = t
                    break
            if new_tab:
                print(f"새 탭 발견 ({i*0.5:.1f}초): {new_tab['url'][:100]}")
                break

        if not new_tab:
            print("새 탭 안 열림 — WebSquare 팝업이 in-page 방식일 수 있음")
            # in-page 팝업 확인
            overlay = await _eval(ws, """(function(){
    var el = document.querySelector('[id*="UTERNAAZ34"]');
    if (!el) return 'not_found';
    var r = el.getBoundingClientRect();
    var btns = Array.from(el.querySelectorAll('input[type=button],button,a'))
        .map(function(b){ return (b.value||b.innerText||'').trim(); })
        .filter(function(s){ return s.length > 0; });
    return JSON.stringify({w:r.width, h:r.height, visible:!!el.offsetParent, btns:btns});
})()""", cmd_id=20)
            print(f"UTERNAAZ34 in-page: {overlay}")
            return

        # 4. 새 탭 분석
        await asyncio.sleep(3)

    async with websockets.connect(new_tab["webSocketDebuggerUrl"]) as ws2:
        rs = await _eval(ws2, "document.readyState", cmd_id=1)
        title = await _eval(ws2, "document.title", cmd_id=2)
        url2 = await _eval(ws2, "window.location.href", cmd_id=3)
        print(f"\nURL: {url2}")
        print(f"readyState: {rs}, title: {title}")

        # 버튼 목록 (iframe 포함)
        btns = await _eval(ws2, """(function(){
    function getBtns(doc, prefix) {
        var r = Array.from(doc.querySelectorAll('input[type=button],button,a'))
            .filter(function(b){ return (b.value||b.innerText||'').trim().length>0; })
            .map(function(b){ return prefix+':'+(b.value||b.innerText||'').trim().slice(0,30); });
        Array.from(doc.querySelectorAll('iframe')).forEach(function(f,i){
            try{ r = r.concat(getBtns(f.contentDocument, 'iframe'+i)); }catch(e){}
        });
        return r;
    }
    return JSON.stringify(getBtns(document,'main'));
})()""", cmd_id=4)
        import json as _j
        bl = _j.loads(btns) if isinstance(btns, str) else (btns or [])
        print(f"\n버튼 목록 ({len(bl)}개):")
        for b in bl:
            print(f"  {b}")

    requests.get(f"{CDP}/json/close/{new_tab['id']}")
    print("\n팝업 탭 닫음")

asyncio.run(run())
