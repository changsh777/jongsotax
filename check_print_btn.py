"""인쇄 버튼 위치 확인 (페이지 수 변경 후 상태)"""
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
    popup = next((t for t in tabs if "UTERNAAZ34" in t.get("url","")), None)
    if not popup:
        print("팝업 없음"); return

    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 스크린샷
        await ws.send(json.dumps({"id": 1, "method": "Page.captureScreenshot",
                                  "params": {"format": "png"}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=15))
            if r.get("id") == 1:
                import base64
                data = r.get("result", {}).get("data", "")
                with open(r"C:\Users\pc\종소세2026\after_confirm.png", "wb") as f:
                    f.write(base64.b64decode(data))
                print("스크린샷: C:\\Users\\pc\\종소세2026\\after_confirm.png")
                break

        # 모든 버튼 위치
        info = await _eval(ws, """JSON.stringify(
    Array.from(document.querySelectorAll('input,button,a,img'))
        .map(function(el){
            var t = (el.value||el.innerText||el.title||el.alt||el.className||'').trim().slice(0,30);
            var r = el.getBoundingClientRect();
            return {text:t, tag:el.tagName, x:Math.round(r.left), y:Math.round(r.top),
                    w:Math.round(r.width), h:Math.round(r.height)};
        })
        .filter(function(o){ return o.w>0 && o.h>0; })
)""", cmd_id=2)
        import json as _j; items = _j.loads(info) if info else []
        print(f"\n보이는 요소 {len(items)}개:")
        for o in items:
            print(f"  [{o['tag']:6}] ({o['x']:4},{o['y']:4}) {o['w']}x{o['h']} — {o['text']}")

asyncio.run(run())
