"""배치 버튼 ID 확인"""
import sys, json, asyncio, requests
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
import websockets

async def main():
    tabs = requests.get("http://localhost:9222/json").json()
    ht = next((t for t in tabs if "hometax.go.kr" in t.get("url","")), None)
    if not ht:
        print("홈택스 탭 없음"); return

    async with websockets.connect(ht["webSocketDebuggerUrl"]) as ws:
        async def ev(code, cid=1):
            await ws.send(json.dumps({"id":cid,"method":"Runtime.evaluate",
                "params":{"expression":code,"returnByValue":True}}))
            while True:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=20))
                if r.get("id")==cid:
                    v = r.get("result",{}).get("result",{}).get("value")
                    return json.loads(v) if isinstance(v,str) else v

        # 모든 visible input[type=button]
        btns = await ev("""(function(){
    return JSON.stringify(Array.from(document.querySelectorAll('input[type=button]'))
        .filter(function(el){ return el.offsetParent !== null; })
        .map(function(el){ return {id:el.id||'',val:el.value||''}; }));
})()""", 1)
        print("VISIBLE input[type=button]:")
        if btns:
            for b in btns:
                print(f"  id={b['id']:60s} val={b['val']}")

        # 접수증/납부서 관련 버튼
        btns2 = await ev("""(function(){
    var all = Array.from(document.querySelectorAll('input[type=button], button'));
    return JSON.stringify(all.filter(function(el){
        var v = (el.value||el.innerText||'').trim();
        return v.includes('접수') || v.includes('납부') || v.includes('일괄') || v.includes('인쇄');
    }).map(function(el){
        return {tag:el.tagName, id:el.id||'', val:(el.value||el.innerText||'').trim().slice(0,40)};
    }));
})()""", 2)
        print("\n접수/납부/일괄/인쇄 버튼:")
        if btns2:
            for b in btns2:
                print(f"  {b['tag']:6s} id={b['id']:60s} val={b['val']}")

asyncio.run(main())
