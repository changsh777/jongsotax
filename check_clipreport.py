"""ClipReport4 iframe 컨텍스트에서 PDF 버튼 찾기 + 클릭"""
import requests, json, asyncio, websockets, pyautogui, time

CDP = "http://localhost:9222"

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    popup = next((t for t in tabs if "UTERNAAZ34" in t.get("url","")), None)
    if not popup:
        print("팝업 없음"); return

    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 1. popup 창 bounds + viewport
        await ws.send(json.dumps({"id": 900, "method": "Browser.getWindowForTarget",
                                  "params": {"targetId": popup["id"]}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 900:
                win_id = r["result"]["windowId"]; break
        await ws.send(json.dumps({"id": 901, "method": "Browser.getWindowBounds",
                                  "params": {"windowId": win_id}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 901:
                bounds = r["result"]["bounds"]; break
        await ws.send(json.dumps({"id": 902, "method": "Runtime.evaluate",
                                  "params": {"expression": "JSON.stringify({w:window.innerWidth,h:window.innerHeight})",
                                             "returnByValue": True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 902:
                vp = json.loads(r["result"]["result"]["value"])
                vp_w, vp_h = vp["w"], vp["h"]; break

        toolbar_h = bounds["height"] - vp_h
        win_left  = bounds["left"]
        win_top   = bounds["top"]
        win_w     = bounds["width"]
        print(f"팝업 창: left={win_left}, top={win_top}, {win_w}x{bounds['height']}, toolbar={toolbar_h}")
        print(f"뷰포트: {vp_w}x{vp_h}")

        # 2. iframe 위치 (mf_iframe2_UTERNAAZ34)
        await ws.send(json.dumps({"id": 10, "method": "Runtime.evaluate",
            "params": {"expression": """(function(){
    var f = document.getElementById('mf_iframe2_UTERNAAZ34');
    if (!f) return JSON.stringify({err:'iframe not found'});
    var r = f.getBoundingClientRect();
    return JSON.stringify({x:r.left, y:r.top, w:r.width, h:r.height,
        src:f.src||f.getAttribute('src')||''});
})()""", "returnByValue": True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 10:
                iframe_info = json.loads(r["result"]["result"]["value"])
                print(f"\niframe 위치: {iframe_info}")
                break

        # 3. Runtime.enable → executionContextCreated 이벤트 수집
        await ws.send(json.dumps({"id": 11, "method": "Runtime.enable"}))
        contexts = []
        deadline = asyncio.get_event_loop().time() + 2.5
        while asyncio.get_event_loop().time() < deadline:
            try:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.3))
                if r.get("method") == "Runtime.executionContextCreated":
                    ctx = r["params"]["context"]
                    origin = ctx.get("origin","")
                    name   = ctx.get("name","")
                    print(f"  context id={ctx['id']}  origin={origin}  name={name}")
                    contexts.append(ctx)
            except asyncio.TimeoutError:
                break
        print(f"\n컨텍스트 {len(contexts)}개 수집")

        # 4. sesw.hometax.go.kr 컨텍스트 찾기
        sesw_ctx = next((c for c in contexts if "sesw.hometax.go.kr" in c.get("origin","")), None)
        if not sesw_ctx:
            print("sesw 컨텍스트 없음 — 모든 컨텍스트에서 버튼 검색")
            target_ctxs = contexts
        else:
            print(f"sesw 컨텍스트 찾음: id={sesw_ctx['id']}")
            target_ctxs = [sesw_ctx]

        for ctx in target_ctxs:
            ctx_id = ctx["id"]
            cmd_id = 200 + ctx_id % 100
            await ws.send(json.dumps({
                "id": cmd_id, "method": "Runtime.evaluate",
                "params": {
                    "expression": """JSON.stringify(
    Array.from(document.querySelectorAll('input,button,a,img,td,div'))
        .filter(function(el){
            var t=(el.value||el.title||el.alt||el.id||el.className||'').toLowerCase();
            var r=el.getBoundingClientRect();
            return r.width>0 && r.height>0 && (
                t.indexOf('pdf')>=0 || t.indexOf('print')>=0 || t.indexOf('저장')>=0 ||
                t.indexOf('출력')>=0 || t.indexOf('save')>=0
            );
        })
        .map(function(el){
            var r=el.getBoundingClientRect();
            return {tag:el.tagName,id:el.id,cls:el.className.slice(0,30),
                    title:el.title,alt:el.alt,val:el.value,
                    x:Math.round(r.left),y:Math.round(r.top),
                    w:Math.round(r.width),h:Math.round(r.height)};
        })
)""",
                    "contextId": ctx_id,
                    "returnByValue": True
                }
            }))
            deadline2 = asyncio.get_event_loop().time() + 3
            while asyncio.get_event_loop().time() < deadline2:
                try:
                    r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.5))
                    if r.get("id") == cmd_id:
                        val = r.get("result",{}).get("result",{}).get("value")
                        if val:
                            items = json.loads(val)
                            print(f"\n  ctx[{ctx_id}] 버튼 {len(items)}개:")
                            for it in items:
                                print(f"    [{it['tag']:5}] id={it['id']:20} x={it['x']:4},y={it['y']:4} "
                                      f"w={it['w']:3},h={it['h']:3} "
                                      f"title={it.get('title','')} cls={it.get('cls','')}")
                        break
                except asyncio.TimeoutError:
                    break

asyncio.run(run())
