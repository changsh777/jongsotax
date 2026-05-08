import requests, json, asyncio, websockets

async def check():
    tabs = requests.get('http://localhost:9222/json').json()
    main = next((t for t in tabs
                 if 'hometax.go.kr' in t.get('url','')
                 and 'websquare.html' in t.get('url','')), None)
    if not main:
        print('홈택스 탭 없음'); return
    async with websockets.connect(main['webSocketDebuggerUrl']) as ws:
        await ws.send(json.dumps({'id':1,'method':'Runtime.evaluate','params':{
            'expression': "document.getElementById('mf_txppWframe_UTERNAAZ0Z31_wframe_edtGrdRowNum') ? '팝업열림' : '팝업닫힘'",
            'returnByValue': True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
            if r.get('id') == 1:
                print(r['result']['result'].get('value','?'))
                break

asyncio.run(check())
