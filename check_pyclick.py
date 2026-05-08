"""pyautogui 실제 OS 클릭으로 팝업 열기 테스트"""
import requests, json, asyncio, websockets, pyautogui, time

CDP = "http://localhost:9222"
pyautogui.PAUSE = 0.1

async def _send(ws, msg):
    await ws.send(json.dumps(msg))
    while True:
        r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
        if r.get("id") == msg["id"]:
            return r

async def _eval(ws, code, cmd_id=1):
    r = await _send(ws, {"id": cmd_id, "method": "Runtime.evaluate",
                         "params": {"expression": code, "returnByValue": True}})
    return r.get("result", {}).get("result", {}).get("value")

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    main = next((t for t in tabs if "hometax.go.kr" in t.get("url","")
                 and "websquare.html" in t.get("url","")
                 and "sesw." not in t.get("url","")), None)
    if not main:
        print("홈택스 탭 없음"); return

    known_ids = {t["id"] for t in tabs}

    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 1. Edge 창 위치 가져오기
        r_win = await _send(ws, {"id": 10, "method": "Browser.getWindowForTarget",
                                 "params": {"targetId": main["id"]}})
        win_id = r_win.get("result", {}).get("windowId")
        r_bounds = await _send(ws, {"id": 11, "method": "Browser.getWindowBounds",
                                    "params": {"windowId": win_id}})
        bounds = r_bounds.get("result", {}).get("bounds", {})
        win_left = bounds.get("left", 0)
        win_top  = bounds.get("top", 0)
        win_w    = bounds.get("width", 1920)
        win_h    = bounds.get("height", 1080)
        print(f"Edge 창: left={win_left}, top={win_top}, w={win_w}, h={win_h}")

        # 2. 뷰포트 크기 (= 창 크기 - 툴바)
        vp = await _eval(ws, "JSON.stringify({w:window.innerWidth, h:window.innerHeight})", cmd_id=12)
        import json as _j; vp = _j.loads(vp) if vp else {}
        vp_w = vp.get("w", win_w)
        vp_h = vp.get("h", win_h)
        toolbar_h = win_h - vp_h  # 툴바 높이
        print(f"뷰포트: {vp_w}x{vp_h}, 툴바 높이: {toolbar_h}px")

        # 3. 버튼 뷰포트 좌표
        coords = await _eval(ws, """(function(){
    var rows = Array.from(document.querySelectorAll('[id*="UTERNAAZ0Z31_wframe"] table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    if (!rows.length) return null;
    var cell = rows[0].querySelectorAll('td')[10];
    var btn = cell && cell.querySelector('a, input[type=button], button');
    if (!btn) return null;
    var r = btn.getBoundingClientRect();
    return JSON.stringify({x: r.left + r.width/2, y: r.top + r.height/2,
        text: (btn.innerText||btn.value||'').trim().slice(0,30)});
})()""", cmd_id=13)
        if not coords:
            print("버튼 없음"); return
        import json as _j2; c = _j2.loads(coords)
        vx, vy = c["x"], c["y"]
        print(f"버튼 뷰포트 좌표: ({vx:.0f}, {vy:.0f}) — {c.get('text')}")

        # 4. 스크린 절대 좌표 계산
        sx = win_left + (win_w - vp_w) // 2 + vx   # 좌우 스크롤바 보정
        sy = win_top + toolbar_h + vy
        print(f"스크린 클릭 좌표: ({sx:.0f}, {sy:.0f})")

    # 5. pyautogui 실제 OS 클릭
    print("Edge 창 포커스 후 클릭...")
    pyautogui.click(int(sx), int(sy))
    time.sleep(0.3)
    pyautogui.click(int(sx), int(sy))  # 더블클릭 아닌 두 번 클릭 (확실히)

    # 6. 새 탭 감지 (최대 10초)
    print("팝업 대기 중...")
    for i in range(20):
        time.sleep(0.5)
        all_tabs = requests.get(f"{CDP}/json").json()
        for t in all_tabs:
            if t["id"] not in known_ids and "devtools" not in t.get("url",""):
                print(f"\n팝업 열림 ({i*0.5:.1f}초): {t['url'][:120]}")
                requests.get(f"{CDP}/json/close/{t['id']}")
                print("테스트 완료 — 팝업 닫음")
                return
    print("팝업 미열림")

asyncio.run(run())
