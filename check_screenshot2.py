"""현재 열린 모든 팝업 탭 스크린샷"""
import requests, json, asyncio, websockets, base64

CDP = "http://localhost:9222"

async def screenshot_tab(tab, filename):
    async with websockets.connect(tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        await ws.send(json.dumps({"id": 1, "method": "Page.captureScreenshot",
                                  "params": {"format": "png"}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=15))
            if r.get("id") == 1:
                data = r.get("result", {}).get("data", "")
                with open(filename, "wb") as f:
                    f.write(base64.b64decode(data))
                print(f"저장: {filename}")
                break

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    print(f"전체 탭 {len(tabs)}개:")
    for i, t in enumerate(tabs):
        print(f"  [{i}] {t['url'][:80]}")

    for i, t in enumerate(tabs):
        url = t.get("url", "")
        if "hometax.go.kr" in url and "devtools" not in url:
            fn = f"C:\\Users\\pc\\종소세2026\\tab_{i}.png"
            await screenshot_tab(t, fn)

asyncio.run(run())
