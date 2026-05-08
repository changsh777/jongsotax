"""pdfDownLoad() 직접 호출 + 전체 네트워크 캡처"""
import requests, json, asyncio, websockets, base64, os

CDP  = "http://localhost:9222"
SAVE = r"C:\Users\pc\종소세2026"
REPORT_KEY = "cad244b29911a4cd1be50985fdb884d08rpt1"  # onclick에서 추출

async def eval_ctx(ws, code, ctx_id, cmd_id, timeout=8):
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
    tabs = requests.get(f"{CDP}/json").json()
    popup = next((t for t in tabs if "UTERNAAZ34" in t.get("url","")), None)
    if not popup:
        print("팝업 없음"); return

    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # Network.enable (전체)
        await ws.send(json.dumps({"id": 1, "method": "Network.enable"}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
            if r.get("id") == 1:
                print("Network.enable OK"); break

        # Browser.setDownloadBehavior
        await ws.send(json.dumps({"id": 2, "method": "Browser.setDownloadBehavior",
            "params": {"behavior": "allow", "downloadPath": SAVE, "eventsEnabled": True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
            if r.get("id") == 2:
                print(f"다운로드 경로: {SAVE}"); break

        # Runtime.enable → ctx 수집
        await ws.send(json.dumps({"id": 3, "method": "Runtime.enable"}))
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
            test = await eval_ctx(ws, "document.querySelectorAll('button').length", c["id"], 50)
            if test and int(test) > 3:
                ctx_id = c["id"]; break
        print(f"ctx_id: {ctx_id}")

        # m_reportHashMap 객체 확인
        methods = await eval_ctx(ws, f"""(function(){{
    var obj = m_reportHashMap['{REPORT_KEY}'];
    if (!obj) return 'obj not found';
    return JSON.stringify(Object.getOwnPropertyNames(obj)
        .filter(function(k){{ return typeof obj[k] === 'function'; }})
        .slice(0, 30));
}})()""", ctx_id, 60)
        print(f"pdfDownLoad 객체 메서드: {methods}")

        # Blob URL 인터셉트
        await eval_ctx(ws, """
window._blobUrls = [];
var _origCreateObj = URL.createObjectURL;
URL.createObjectURL = function(blob) {
    var url = _origCreateObj.call(URL, blob);
    window._blobUrls.push({url: url, size: blob.size, type: blob.type});
    console.log('[BLOB] ' + blob.type + ' size=' + blob.size + ' url=' + url);
    return url;
};
""", ctx_id, 61)
        print("Blob 인터셉트 설정")

        # pdfDownLoad() 직접 호출
        result = await eval_ctx(ws, f"""(function(){{
    try {{
        m_reportHashMap['{REPORT_KEY}'].pdfDownLoad();
        return 'called';
    }} catch(e) {{
        return 'error: ' + e.message;
    }}
}})()""", ctx_id, 62)
        print(f"pdfDownLoad() 결과: {result}")

        # 20초 전체 네트워크 + 이벤트 감시
        print("\n20초 전체 모니터링...\n")
        responses = {}  # requestId → {url, mime}
        deadline = asyncio.get_event_loop().time() + 20

        while asyncio.get_event_loop().time() < deadline:
            try:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.3))
                method = r.get("method","")

                if method == "Network.requestWillBeSent":
                    p = r["params"]
                    url = p.get("request",{}).get("url","")
                    meth = p.get("request",{}).get("method","")
                    rid  = p.get("requestId","")
                    print(f"  REQ  {meth:4} {url[:100]}")

                elif method == "Network.responseReceived":
                    p = r["params"]
                    url  = p.get("response",{}).get("url","")
                    mime = p.get("response",{}).get("mimeType","")
                    status = p.get("response",{}).get("status",0)
                    rid  = p.get("requestId","")
                    responses[rid] = {"url": url, "mime": mime}
                    print(f"  RESP {status} [{mime:30}] {url[:80]}")

                elif method == "Network.loadingFinished":
                    rid = r["params"].get("requestId","")
                    if rid in responses:
                        info = responses[rid]
                        if "pdf" in info["mime"].lower() or "octet" in info["mime"].lower():
                            print(f"\n  ★ PDF 응답 완료: {info['url'][:80]}")
                            # 응답 본문 추출
                            await ws.send(json.dumps({"id": 999, "method": "Network.getResponseBody",
                                "params": {"requestId": rid}}))

                elif r.get("id") == 999:
                    body = r.get("result",{}).get("body","")
                    b64  = r.get("result",{}).get("base64Encoded", False)
                    if body:
                        data = base64.b64decode(body) if b64 else body.encode()
                        path = os.path.join(SAVE, "clipreport_dl.pdf")
                        with open(path, "wb") as f:
                            f.write(data)
                        print(f"  ✓ PDF 저장: {path} ({len(data)//1024}KB)")
                    else:
                        print("  본문 비어있음")

                elif method == "Browser.downloadWillBegin":
                    p = r.get("params",{})
                    print(f"\n  ★ DOWNLOAD 시작: {p.get('suggestedFilename','')} — {p.get('url','')[:80]}")

                elif method == "Browser.downloadProgress":
                    p = r.get("params",{})
                    if p.get("state") == "completed":
                        print(f"  ★ DOWNLOAD 완료!")

                elif method == "Runtime.consoleAPICalled":
                    args = r.get("params",{}).get("args",[])
                    msg = " ".join(str(a.get("value","")) for a in args)
                    if msg.strip():
                        print(f"  [console] {msg[:100]}")

            except asyncio.TimeoutError:
                pass

        # Blob URL 확인
        blobs = await eval_ctx(ws, "JSON.stringify(window._blobUrls)", ctx_id, 90)
        print(f"\nBlob URLs: {blobs}")

asyncio.run(run())
