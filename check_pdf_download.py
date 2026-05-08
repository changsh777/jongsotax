"""ClipReport4 hidden PDF 버튼 CDP 클릭 → 다운로드 가로채기"""
import requests, json, asyncio, websockets, os

CDP  = "http://localhost:9222"
SAVE = r"C:\Users\pc\종소세2026"

async def eval_ctx(ws, code, ctx_id, cmd_id, timeout=10):
    await ws.send(json.dumps({"id": cmd_id, "method": "Runtime.evaluate",
        "params": {"expression": code, "contextId": ctx_id, "returnByValue": True}}))
    deadline = asyncio.get_event_loop().time() + timeout
    while asyncio.get_event_loop().time() < deadline:
        try:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.5))
            if r.get("id") == cmd_id:
                return r.get("result",{}).get("result",{}).get("value")
        except asyncio.TimeoutError:
            pass
    return None

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    popup = next((t for t in tabs if "UTERNAAZ34" in t.get("url","")), None)
    if not popup:
        print("팝업 없음 — 접수번호 클릭 후 일괄출력까지 완료해주세요"); return

    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws:

        # ── 1. 다운로드 경로 설정 ────────────────────────────────────────────
        await ws.send(json.dumps({"id": 1, "method": "Browser.setDownloadBehavior",
            "params": {"behavior": "allow", "downloadPath": SAVE,
                       "eventsEnabled": True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
            if r.get("id") == 1:
                print(f"다운로드 경로 설정: {SAVE}"); break

        # ── 2. Runtime.enable → sesw 컨텍스트 수집 ──────────────────────────
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

        # sesw 컨텍스트 중 clipreport.do 가진 것 찾기 (총요소 많은 것)
        ctx_id = None
        for c in contexts:
            # 간단 확인: document.querySelectorAll('button').length
            test = await eval_ctx(ws, "document.querySelectorAll('button').length", c["id"], 50+c["id"]%50)
            if test and int(test) > 3:
                ctx_id = c["id"]
                print(f"clipreport 컨텍스트: id={ctx_id} (버튼 {test}개)")
                break

        if not ctx_id:
            print("clipreport 컨텍스트 못 찾음"); return

        # ── 3. PDF 버튼 ID 찾기 ──────────────────────────────────────────────
        pdf_id = await eval_ctx(ws, """(function(){
    var btn = Array.from(document.querySelectorAll('button'))
        .find(function(b){ return b.id.indexOf('re_pdf') === 0; });
    return btn ? btn.id : null;
})()""", ctx_id, 60)
        print(f"PDF 버튼 id: {pdf_id}")

        if not pdf_id:
            print("PDF 버튼 없음 — 저장 버튼 먼저 클릭 시도")
            # 저장 버튼 클릭해서 드롭다운 열기
            save_id = await eval_ctx(ws, """(function(){
    var btn = Array.from(document.querySelectorAll('button'))
        .find(function(b){ return b.id.indexOf('re_save') === 0; });
    return btn ? btn.id : null;
})()""", ctx_id, 61)
            print(f"저장 버튼 id: {save_id}")
            if save_id:
                await eval_ctx(ws, f"document.getElementById('{save_id}').click()", ctx_id, 62)
                await asyncio.sleep(0.5)
                pdf_id = await eval_ctx(ws, """(function(){
    var btn = Array.from(document.querySelectorAll('button'))
        .find(function(b){ return b.id.indexOf('re_pdf') === 0; });
    return btn ? btn.id : null;
})()""", ctx_id, 63)
            if not pdf_id:
                print("PDF 버튼 여전히 없음"); return

        # ── 4. PDF 버튼 CDP 클릭 ─────────────────────────────────────────────
        print(f"PDF 버튼 클릭: {pdf_id}")
        result = await eval_ctx(ws, f"""(function(){{
    var btn = document.getElementById('{pdf_id}');
    if (!btn) return 'btn not found';
    btn.click();
    return 'clicked';
}})()""", ctx_id, 70)
        print(f"클릭 결과: {result}")

        # ── 5. 다운로드 이벤트 대기 (최대 30초) ─────────────────────────────
        print("다운로드 대기 중...")
        deadline = asyncio.get_event_loop().time() + 30
        while asyncio.get_event_loop().time() < deadline:
            try:
                r = json.loads(await asyncio.wait_for(ws.recv(), timeout=1))
                method = r.get("method","")
                if method == "Browser.downloadWillBegin":
                    p = r.get("params",{})
                    print(f"\n  다운로드 시작: {p.get('url','')[:80]}")
                    print(f"  저장명: {p.get('suggestedFilename','')}")
                elif method == "Browser.downloadProgress":
                    p = r.get("params",{})
                    state = p.get("state","")
                    if state == "completed":
                        print(f"\n✓ 다운로드 완료! state={state}")
                        print(f"  저장 폴더: {SAVE}")
                        # 최근 파일 확인
                        files = sorted([f for f in os.listdir(SAVE) if f.endswith('.pdf')],
                                      key=lambda f: os.path.getmtime(os.path.join(SAVE,f)), reverse=True)
                        if files:
                            print(f"  최근 PDF: {files[0]}")
                        return
                    elif state == "canceled":
                        print(f"  다운로드 취소됨"); return
                    else:
                        recv = p.get("receivedBytes",0)
                        total = p.get("totalBytes",0)
                        if total:
                            print(f"  진행: {recv}/{total} ({100*recv//total}%)", end="\r")
            except asyncio.TimeoutError:
                pass
        print("30초 타임아웃 — 다운로드 미감지")

asyncio.run(run())
