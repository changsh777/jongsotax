"""ClipReport4 ctx=385 — 툴바 전체 + PDF 관련 JS API 검색"""
import requests, json, asyncio, websockets

CDP = "http://localhost:9222"

async def eval_ctx(ws, code, ctx_id, cmd_id):
    await ws.send(json.dumps({"id": cmd_id, "method": "Runtime.evaluate",
        "params": {"expression": code, "contextId": ctx_id, "returnByValue": True}}))
    deadline = asyncio.get_event_loop().time() + 5
    while asyncio.get_event_loop().time() < deadline:
        try:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.5))
            if r.get("id") == cmd_id:
                return r.get("result",{}).get("result",{}).get("value")
        except asyncio.TimeoutError:
            break
    return None

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    popup = next((t for t in tabs if "UTERNAAZ34" in t.get("url","")), None)
    if not popup:
        print("팝업 없음"); return

    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # Runtime.enable + 컨텍스트 수집
        await ws.send(json.dumps({"id": 11, "method": "Runtime.enable"}))
        contexts = []
        deadline = asyncio.get_event_loop().time() + 2.5
        while asyncio.get_event_loop().time() < deadline:
            try:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.3))
                if r.get("method") == "Runtime.executionContextCreated":
                    contexts.append(r["params"]["context"])
            except asyncio.TimeoutError:
                break

        ctx_id = 385  # sesw clipreport.do

        # 1. 툴바 영역(y < 70) 모든 버튼
        btns = await eval_ctx(ws, """JSON.stringify(
    Array.from(document.querySelectorAll('button,input[type=button],a,img'))
        .map(function(el){
            var r = el.getBoundingClientRect();
            var src = el.src || el.getAttribute('src') || '';
            return {tag:el.tagName, id:el.id, title:el.title,
                    alt:el.alt, src:src.split('/').pop(),
                    onclick:(el.onclick||'').toString().slice(0,50),
                    x:Math.round(r.left), y:Math.round(r.top),
                    w:Math.round(r.width), h:Math.round(r.height)};
        })
        .filter(function(o){ return o.w>0 && o.h>0 && o.y < 70; })
)""", ctx_id, 100)
        print("=== 툴바 버튼 (y<70) ===")
        for b in (json.loads(btns) if btns else []):
            print(f"  [{b['tag']:6}] id={b['id']:30} ({b['x']:4},{b['y']:2}) src={b['src']:20} title={b['title']}")

        # 2. PDF 관련 img 태그
        imgs = await eval_ctx(ws, """JSON.stringify(
    Array.from(document.querySelectorAll('img,button'))
        .filter(function(el){
            var s=(el.src||el.getAttribute('src')||el.id||el.className||el.title||'').toLowerCase();
            return s.indexOf('pdf')>=0 || s.indexOf('hwp')>=0 || s.indexOf('excel')>=0 || s.indexOf('save')>=0;
        })
        .map(function(el){
            var r = el.getBoundingClientRect();
            return {tag:el.tagName, id:el.id, title:el.title,
                    src:(el.src||'').split('/').pop(),
                    cls:el.className.toString(),
                    x:Math.round(r.left), y:Math.round(r.top),
                    w:Math.round(r.width), h:Math.round(r.height)};
        })
)""", ctx_id, 101)
        print("\n=== PDF/HWP/Excel/Save 관련 요소 ===")
        for b in (json.loads(imgs) if imgs else []):
            print(f"  [{b['tag']:6}] ({b['x']:4},{b['y']:3}) {b['w']}x{b['h']} id={b['id']} title={b['title']} src={b['src']}")

        # 3. ClipReport JS API 검색
        api = await eval_ctx(ws, """(function(){
    var keys = [];
    try { keys = Object.keys(window).filter(function(k){
        return k.toLowerCase().indexOf('clip')>=0 ||
               k.toLowerCase().indexOf('report')>=0 ||
               k.toLowerCase().indexOf('cr_')>=0 ||
               k.toLowerCase().indexOf('viewer')>=0 ||
               k.toLowerCase().indexOf('pdf')>=0;
    }); } catch(e){}
    return JSON.stringify(keys.slice(0,30));
})()""", ctx_id, 102)
        print(f"\n=== ClipReport 관련 전역 변수 ===\n  {api}")

        # 4. 전체 페이지 수 읽기
        total = await eval_ctx(ws, """(function(){
    var el = document.querySelector('[id*="totalCount"]');
    return el ? (el.value || el.innerText || el.textContent) : null;
})()""", ctx_id, 103)
        print(f"\n전체 페이지 수: {total}")

asyncio.run(run())
