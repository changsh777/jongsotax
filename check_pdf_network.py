"""PDF 버튼 클릭 후 네트워크 요청 + 새 탭 감지"""
import requests, json, asyncio, websockets, time

CDP  = "http://localhost:9222"

async def eval_ctx(ws, code, ctx_id, cmd_id, timeout=5):
    await ws.send(json.dumps({"id": cmd_id, "method": "Runtime.evaluate",
        "params": {"expression": code, "contextId": ctx_id, "returnByValue": True}}))
    deadline = asyncio.get_event_loop().time() + timeout
    while asyncio.get_event_loop().time() < deadline:
        try:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.3))
            if r.get("id") == cmd_id:
                return r.get("result",{}).get("result",{}).get("value")
        except asyncio.TimeoutError:
            pass
    return None

async def run():
    tabs_before = requests.get(f"{CDP}/json").json()
    known_ids = {t["id"] for t in tabs_before}

    popup = next((t for t in tabs_before if "UTERNAAZ34" in t.get("url","")), None)
    if not popup:
        print("팝업 없음"); return

    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws:

        # Runtime.enable → sesw ctx 찾기
        await ws.send(json.dumps({"id": 2, "method": "Runtime.enable"}))
        contexts = []
        deadline = asyncio.get_event_loop().time() + 2.5
        while asyncio.get_event_loop().time() < deadline:
            try:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.3))
                if r.get("method") == "Runtime.executionContextCreated":
                    c = r["params"]["context"]
                    if "sesw.hometax.go.kr" in c.get("origin",""):
                        contexts.append(c)
            except asyncio.TimeoutError:
                break

        ctx_id = None
        for c in contexts:
            test = await eval_ctx(ws, "document.querySelectorAll('button').length", c["id"], 50+c["id"]%50)
            if test and int(test) > 3:
                ctx_id = c["id"]
                break
        print(f"ctx_id: {ctx_id}")

        # onclick 핸들러 확인
        onclick_info = await eval_ctx(ws, """(function(){
    var btn = Array.from(document.querySelectorAll('button'))
        .find(function(b){ return b.id.indexOf('re_pdf') === 0; });
    if (!btn) return 'not found';
    return JSON.stringify({
        id: btn.id,
        onclick: btn.getAttribute('onclick') || '',
        outerHTML: btn.outerHTML.slice(0, 200)
    });
})()""", ctx_id, 60)
        print(f"PDF 버튼 info:\n  {onclick_info}\n")

        # Network.enable
        await ws.send(json.dumps({"id": 3, "method": "Network.enable"}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
            if r.get("id") == 3:
                print("Network.enable OK"); break

        # window.open 가로채기
        await eval_ctx(ws, """
window._openedUrls = [];
var _orig = window.open;
window.open = function(url, name, specs) {
    window._openedUrls.push(url || '');
    console.log('window.open intercepted: ' + url);
    return _orig.apply(this, arguments);
};
""", ctx_id, 61)
        print("window.open 인터셉트 설정")

        # PDF 버튼 클릭
        pdf_id = await eval_ctx(ws, """(function(){
    var btn = Array.from(document.querySelectorAll('button'))
        .find(function(b){ return b.id.indexOf('re_pdf') === 0; });
    return btn ? btn.id : null;
})()""", ctx_id, 62)

        print(f"PDF 버튼 클릭: {pdf_id}")
        await eval_ctx(ws, f"document.getElementById('{pdf_id}').click()", ctx_id, 63)

        # 15초 동안 네트워크 이벤트 + 새 탭 감지
        print("15초 모니터링 중...\n")
        deadline = asyncio.get_event_loop().time() + 15
        net_events = []
        while asyncio.get_event_loop().time() < deadline:
            try:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.3))
                method = r.get("method","")

                if method == "Network.requestWillBeSent":
                    p = r["params"]
                    url = p.get("request",{}).get("url","")
                    req_method = p.get("request",{}).get("method","")
                    if "sesw" in url or "pdf" in url.lower() or "serp" in url:
                        print(f"  REQ  {req_method} {url[:100]}")
                        net_events.append(("req", url))

                elif method == "Network.responseReceived":
                    p = r["params"]
                    url = p.get("response",{}).get("url","")
                    mime = p.get("response",{}).get("mimeType","")
                    status = p.get("response",{}).get("status",0)
                    if "sesw" in url or "pdf" in mime.lower() or "serp" in url:
                        print(f"  RESP {status} {mime} {url[:100]}")
                        net_events.append(("resp", url, mime, p.get("requestId")))

                elif method == "Browser.downloadWillBegin":
                    p = r.get("params",{})
                    print(f"\n  ✓ DOWNLOAD: {p.get('suggestedFilename','')} — {p.get('url','')[:80]}")

                elif method == "Target.targetCreated":
                    p = r.get("params",{}).get("targetInfo",{})
                    if p.get("targetId") not in known_ids:
                        print(f"\n  ✓ 새 탭: {p.get('url','')[:100]}")

            except asyncio.TimeoutError:
                pass

        # window.open 결과 확인
        opened = await eval_ctx(ws, "JSON.stringify(window._openedUrls)", ctx_id, 90)
        print(f"\nwindow.open 호출 URL: {opened}")

        # PDF 응답 본문 추출 (있으면)
        for ev in net_events:
            if ev[0] == "resp" and ("pdf" in ev[2].lower() or "octet" in ev[2].lower()):
                print(f"\nPDF 응답 발견 — 본문 추출 시도: {ev[1][:80]}")
                req_id = ev[3]
                await ws.send(json.dumps({"id": 999, "method": "Network.getResponseBody",
                                          "params": {"requestId": req_id}}))
                deadline2 = asyncio.get_event_loop().time() + 10
                while asyncio.get_event_loop().time() < deadline2:
                    try:
                        r2 = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.5))
                        if r2.get("id") == 999:
                            body = r2.get("result",{}).get("body","")
                            b64 = r2.get("result",{}).get("base64Encoded", False)
                            print(f"  body len={len(body)}, base64={b64}")
                            if body:
                                import base64 as _b64
                                data = _b64.b64decode(body) if b64 else body.encode()
                                path = r"C:\Users\pc\종소세2026\clipreport_pdf.pdf"
                                with open(path, "wb") as f:
                                    f.write(data)
                                print(f"  ✓ 저장: {path} ({len(data)//1024}KB)")
                            break
                    except asyncio.TimeoutError:
                        pass

asyncio.run(run())
