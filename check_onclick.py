"""홈택스 접수번호 onclick 속성 확인 (Edge CDP 9222)"""
import requests, json, asyncio, websockets

async def run():
    tabs = requests.get("http://localhost:9222/json").json()
    tab = next((t for t in tabs if "hometax" in t.get("url","") and "websquare.html" in t.get("url","")), None)
    if not tab:
        print("홈택스 탭 없음. Edge --remote-debugging-port=9222 로 열려있어야 합니다.")
        return
    print("탭:", tab["url"][:80])
    async with websockets.connect(tab["webSocketDebuggerUrl"]) as ws:
        code = """(function(){
    var rows = Array.from(document.querySelectorAll('[id*="UTERNAAZ0Z31_wframe"] table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 11; });
    var result = [];
    rows.slice(0,3).forEach(function(tr,i){
        var cell = tr.querySelectorAll('td')[10];
        var btn = cell && cell.querySelector('input[type=button],button,a');
        if(btn) result.push({
            row:i, tag:btn.tagName,
            onclick: btn.getAttribute('onclick'),
            id: btn.id,
            text:(btn.innerText||btn.value||'').trim().slice(0,30)
        });
    });
    return JSON.stringify(result);
})()"""
        await ws.send(json.dumps({"id":1,"method":"Runtime.evaluate","params":{"expression":code,"returnByValue":True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id")==1:
                val = r.get("result",{}).get("result",{}).get("value","")
                data = json.loads(val) if val else []
                for d in data:
                    print(f"\n[행{d['row']}] tag={d['tag']} text={d['text']}")
                    print(f"  onclick: {d['onclick']}")
                    print(f"  id:      {d['id']}")
                break

asyncio.run(run())
