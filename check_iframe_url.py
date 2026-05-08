"""팝업 iframe src URL + CDP 서브타겟 확인"""
import requests, json, asyncio, websockets

CDP = "http://localhost:9222"

async def _eval(ws, code, cmd_id=1):
    await ws.send(json.dumps({"id": cmd_id, "method": "Runtime.evaluate",
                              "params": {"expression": code, "returnByValue": True}}))
    while True:
        r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
        if r.get("id") == cmd_id:
            return r.get("result", {}).get("result", {}).get("value")

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    print(f"=== 전체 탭/타겟 ({len(tabs)}개) ===")
    for t in tabs:
        print(f"  [{t.get('type','?'):10}] {t.get('url','')[:100]}")

    popup = next((t for t in tabs if "UTERNAAZ34" in t.get("url","")), None)
    if not popup:
        print("\n팝업 없음 — 먼저 접수번호 클릭해서 일괄출력까지 진행하세요"); return

    print(f"\n=== 팝업: {popup['url'][:80]} ===")

    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 1. 모든 iframe src 목록
        iframes = await _eval(ws, """JSON.stringify(
    Array.from(document.querySelectorAll('iframe, frame, embed, object'))
        .map(function(el){
            return {tag:el.tagName,
                    src:(el.src||el.data||el.getAttribute('src')||''),
                    id:(el.id||''), name:(el.name||''), w:el.offsetWidth, h:el.offsetHeight};
        })
)""", cmd_id=1)
        import json as _j
        items = _j.loads(iframes) if iframes else []
        print(f"\niframe/embed 목록 ({len(items)}개):")
        for it in items:
            print(f"  [{it['tag']}] {it['src'][:120]}  ({it['w']}x{it['h']}) id={it['id']}")

        # 2. 페이지 프레임 트리
        await ws.send(json.dumps({"id": 2, "method": "Page.getFrameTree"}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 2:
                def print_frame(f, depth=0):
                    fr = f.get("frame", {})
                    print(f"  {'  '*depth}frame: {fr.get('url','')[:100]}")
                    for child in f.get("childFrames", []):
                        print_frame(child, depth+1)
                print("\nFrame tree:")
                print_frame(r.get("result", {}).get("frameTree", {}))
                break

        # 3. 네트워크 리소스 목록 (로드된 URL들)
        await ws.send(json.dumps({"id": 3, "method": "Page.getResourceTree"}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
            if r.get("id") == 3:
                def print_resources(node, depth=0):
                    fr = node.get("frame", {})
                    print(f"\n  {'  '*depth}[frame] {fr.get('url','')[:80]}")
                    for res in node.get("resources", []):
                        rtype = res.get("type","?")
                        rurl  = res.get("url","")[:100]
                        print(f"  {'  '*depth}  [{rtype:12}] {rurl}")
                    for child in node.get("childFrames", []):
                        print_resources(child, depth+1)
                print("\nResource tree:")
                print_resources(r.get("result", {}).get("frameTree", {}))
                break

asyncio.run(run())
