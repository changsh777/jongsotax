"""고객 목록 열 구조 확인"""
import requests, json, asyncio, websockets

CDP = "http://localhost:9222"

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    main = next((t for t in tabs if "hometax.go.kr" in t.get("url","")
                 and "websquare.html" in t.get("url","")
                 and "sesw." not in t.get("url","")), None)
    if not main: print("홈택스 탭 없음"); return

    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        await ws.send(json.dumps({"id": 1, "method": "Runtime.evaluate",
            "params": {"expression": """(function(){
    var rows = Array.from(document.querySelectorAll('[id*="UTERNAAZ0Z31_wframe"] table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    if (!rows.length) return '행 없음';
    // 첫 3행 × 전체 열
    return JSON.stringify(rows.slice(0,3).map(function(tr, ri){
        return Array.from(tr.querySelectorAll('td')).map(function(td, ci){
            return '['+ci+']='+td.innerText.trim().replace(/\\s+/g,' ').slice(0,15);
        });
    }));
})()""", "returnByValue": True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 1:
                val = r["result"]["result"]["value"]
                rows = json.loads(val) if val and val != '행 없음' else []
                for i, row in enumerate(rows):
                    print(f"행{i}: {row}")
                break

asyncio.run(run())
