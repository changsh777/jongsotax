"""ClipReport4 — sesw 컨텍스트 전체 버튼 덤프"""
import requests, json, asyncio, websockets

CDP = "http://localhost:9222"

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    popup = next((t for t in tabs if "UTERNAAZ34" in t.get("url","")), None)
    if not popup:
        print("팝업 없음"); return

    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # Runtime.enable → 컨텍스트 수집
        await ws.send(json.dumps({"id": 11, "method": "Runtime.enable"}))
        contexts = []
        deadline = asyncio.get_event_loop().time() + 2.5
        while asyncio.get_event_loop().time() < deadline:
            try:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.3))
                if r.get("method") == "Runtime.executionContextCreated":
                    ctx = r["params"]["context"]
                    contexts.append(ctx)
            except asyncio.TimeoutError:
                break

        sesw_ctxs = [c for c in contexts if "sesw" in c.get("origin","")]
        print(f"sesw 컨텍스트 {len(sesw_ctxs)}개: {[c['id'] for c in sesw_ctxs]}")

        # 모든 sesw 컨텍스트에서 visible 요소 덤프
        for ctx in sesw_ctxs:
            ctx_id = ctx["id"]
            cmd_id = 300 + (ctx_id % 100)
            await ws.send(json.dumps({
                "id": cmd_id, "method": "Runtime.evaluate",
                "params": {
                    "expression": """(function(){
    var els = Array.from(document.querySelectorAll('*'));
    var vis = els.filter(function(el){
        var r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0 && r.height < 60;
    });
    return JSON.stringify({
        url: location.href,
        total: els.length,
        visible: vis.length,
        btns: vis.slice(0, 30).map(function(el){
            var r = el.getBoundingClientRect();
            return {tag:el.tagName, id:el.id.slice(0,20),
                    cls:el.className.toString().slice(0,25),
                    title:(el.title||el.alt||el.value||el.innerText||'').slice(0,20),
                    x:Math.round(r.left), y:Math.round(r.top),
                    w:Math.round(r.width), h:Math.round(r.height)};
        })
    });
})()""",
                    "contextId": ctx_id,
                    "returnByValue": True
                }
            }))

        # 응답 수집
        pending = {300 + (c["id"] % 100) for c in sesw_ctxs}
        deadline2 = asyncio.get_event_loop().time() + 5
        while pending and asyncio.get_event_loop().time() < deadline2:
            try:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.5))
                if r.get("id") in pending:
                    pending.discard(r["id"])
                    val = r.get("result",{}).get("result",{}).get("value")
                    if val:
                        d = json.loads(val)
                        print(f"\n  cmd={r['id']} url={d.get('url','')[:60]}")
                        print(f"  총요소={d['total']} 보이는요소={d['visible']}")
                        for b in d.get("btns", []):
                            print(f"    [{b['tag']:5}] id={b['id']:20} ({b['x']:4},{b['y']:4}) "
                                  f"{b['w']:3}x{b['h']:2} title={b['title']}")
                    else:
                        print(f"  cmd={r['id']} val=None  error={r.get('result',{}).get('exceptionDetails','')}")
            except asyncio.TimeoutError:
                break

asyncio.run(run())
