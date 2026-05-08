"""팝업 스크린샷 저장 + 모든 iframe 포함 버튼 위치 출력"""
import requests, json, asyncio, websockets, base64

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
    # UTERNAAZ34 팝업 탭 찾기
    popup = next((t for t in tabs if "UTERNAAZ34" in t.get("url","")), None)
    if not popup:
        print("UTERNAAZ34 팝업 없음 — 먼저 접수번호 클릭해서 팝업 열어주세요")
        return
    print(f"팝업 탭: {popup['url'][:80]}")

    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 스크린샷
        await ws.send(json.dumps({"id": 1, "method": "Page.captureScreenshot",
                                  "params": {"format": "png", "quality": 80}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=15))
            if r.get("id") == 1:
                data = r.get("result", {}).get("data", "")
                path = r"C:\Users\pc\종소세2026\popup_screenshot.png"
                with open(path, "wb") as f:
                    f.write(base64.b64decode(data))
                print(f"스크린샷 저장: {path}")
                break

        # 모든 input[type=button], button, a 위치 출력
        info = await _eval(ws, """JSON.stringify(
    Array.from(document.querySelectorAll('input[type=button],button,a'))
        .map(function(el){
            var t = (el.value||el.innerText||'').trim().slice(0,20);
            var r = el.getBoundingClientRect();
            var vis = window.getComputedStyle(el).display !== 'none'
                   && window.getComputedStyle(el).visibility !== 'hidden';
            return {text:t, x:Math.round(r.left), y:Math.round(r.top),
                    w:Math.round(r.width), h:Math.round(r.height), vis:vis};
        })
        .filter(function(o){ return o.text.length > 0; })
)""", cmd_id=2)
        import json as _j; items = _j.loads(info) if info else []
        print(f"\n버튼 {len(items)}개 위치:")
        for o in items:
            print(f"  [{o['text']:15}] x={o['x']:4}, y={o['y']:4}, w={o['w']:3}, h={o['h']:3}, visible={o['vis']}")

asyncio.run(run())
