"""
접수번호 클릭 → UTERNAAZ34 팝업 → 일괄출력 → ClipReport pdfDownLoad → 저장
Downloads 폴더에 저장된 파일을 target 폴더로 이동
"""
import requests, json, asyncio, websockets, pyautogui, time, os, shutil, sys

CDP      = "http://localhost:9222"
DOWNLOADS = os.path.join(os.environ.get("USERPROFILE", r"C:\Users\pc"), "Downloads")
SAVE_DIR  = r"C:\Users\pc\종소세2026"   # 임시 저장 폴더 (나중에 NAS 고객폴더로 교체)
pyautogui.PAUSE = 0.15

# ── 헬퍼 ─────────────────────────────────────────────────────────────────────

def get_screen_xy(bounds, vp_w, vp_h, vx, vy):
    win_left  = bounds.get("left", 0)
    win_top   = bounds.get("top",  0)
    win_w     = bounds.get("width", 1920)
    win_h     = bounds.get("height", 1080)
    toolbar_h = win_h - vp_h
    sx = win_left + (win_w - vp_w) // 2 + vx
    sy = win_top  + toolbar_h + vy
    return int(sx), int(sy)

async def _send(ws, msg, timeout=15):
    await ws.send(json.dumps(msg))
    while True:
        r = json.loads(await asyncio.wait_for(ws.recv(), timeout=timeout))
        if r.get("id") == msg["id"]:
            return r

async def _eval(ws, code, cmd_id=1, ctx_id=None):
    params = {"expression": code, "returnByValue": True}
    if ctx_id:
        params["contextId"] = ctx_id
    r = await _send(ws, {"id": cmd_id, "method": "Runtime.evaluate", "params": params})
    return r.get("result", {}).get("result", {}).get("value")

async def get_win_info(ws, target_id):
    r  = await _send(ws, {"id": 900, "method": "Browser.getWindowForTarget",
                          "params": {"targetId": target_id}})
    wid = r["result"]["windowId"]
    r2 = await _send(ws, {"id": 901, "method": "Browser.getWindowBounds",
                          "params": {"windowId": wid}})
    bounds = r2["result"]["bounds"]
    vp = await _eval(ws, "JSON.stringify({w:window.innerWidth,h:window.innerHeight})", cmd_id=902)
    vp = json.loads(vp) if vp else {}
    return bounds, vp.get("w", 1200), vp.get("h", 800)

async def pyclick_btn(ws, target_id, selector_js, label="버튼", cmd_id=50):
    bounds, vp_w, vp_h = await get_win_info(ws, target_id)
    coords = await _eval(ws, f"""(function(){{
    var btn = ({selector_js});
    if (!btn) return null;
    var r = btn.getBoundingClientRect();
    if (r.width===0 && r.height===0) return null;
    return JSON.stringify({{x:r.left+r.width/2, y:r.top+r.height/2,
        text:(btn.value||btn.innerText||'').trim().slice(0,30)}});
}})()""", cmd_id=cmd_id)
    if not coords:
        print(f"  {label} 좌표 없음"); return False
    c = json.loads(coords)
    sx, sy = get_screen_xy(bounds, vp_w, vp_h, c["x"], c["y"])
    print(f"  {label} 클릭: 뷰포트({c['x']:.0f},{c['y']:.0f}) → 스크린({sx},{sy})")
    pyautogui.click(sx, sy)
    return True

async def collect_contexts(ws, seconds=2.5):
    """Runtime.enable 후 executionContextCreated 이벤트 수집"""
    await ws.send(json.dumps({"id": 11, "method": "Runtime.enable"}))
    contexts = []
    deadline = asyncio.get_event_loop().time() + seconds
    while asyncio.get_event_loop().time() < deadline:
        try:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.3))
            if r.get("method") == "Runtime.executionContextCreated":
                contexts.append(r["params"]["context"])
        except asyncio.TimeoutError:
            break
    return contexts

async def find_clipreport_ctx(ws, contexts):
    """sesw 컨텍스트 중 버튼 많은 것 = clipreport.do"""
    sesw = [c for c in contexts if "sesw.hometax.go.kr" in c.get("origin","")]
    for c in sesw:
        cnt = await _eval(ws, "document.querySelectorAll('button').length",
                          cmd_id=200+c["id"]%100, ctx_id=c["id"])
        if cnt and int(cnt) > 3:
            return c["id"]
    return None

# ── 메인 ─────────────────────────────────────────────────────────────────────

async def run():
    tabs = requests.get(f"{CDP}/json").json()
    main = next((t for t in tabs
                 if "hometax.go.kr" in t.get("url","")
                 and "websquare.html" in t.get("url","")
                 and "sesw." not in t.get("url","")), None)
    if not main:
        print("홈택스 탭 없음"); return

    known_ids = {t["id"] for t in tabs}

    # ── STEP 1: 고객명 + 접수번호 클릭 ──────────────────────────────────────
    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 고객명 (td[3] 또는 td[4] — 실제 열 인덱스는 화면에 맞게 조정)
        name_raw = await _eval(ws, """(function(){
    var rows = Array.from(document.querySelectorAll('[id*="UTERNAAZ0Z31_wframe"] table tbody tr'))
        .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
    if (!rows.length) return null;
    var tds = rows[0].querySelectorAll('td');
    // td[3]=납세자명 추정 — 내용 보고 조정
    return JSON.stringify(Array.from(tds).slice(0,6).map(function(td){
        return td.innerText.trim().replace(/\\s+/g,' ');
    }));
})()""", cmd_id=10)
        print(f"첫 행 앞 6열: {name_raw}")

        ok = await pyclick_btn(ws, main["id"], """
    (function(){
        var rows = Array.from(document.querySelectorAll('[id*="UTERNAAZ0Z31_wframe"] table tbody tr'))
            .filter(function(tr){ return tr.querySelectorAll('td').length >= 13; });
        if (!rows.length) return null;
        var cell = rows[0].querySelectorAll('td')[10];
        return cell && cell.querySelector('a, input[type=button], button');
    })()
""", label="접수번호", cmd_id=20)
        if not ok:
            print("접수번호 버튼 없음"); return

    # ── STEP 2: UTERNAAZ34 팝업 대기 ────────────────────────────────────────
    print("UTERNAAZ34 팝업 대기...")
    popup_tab = None
    for i in range(30):
        time.sleep(0.5)
        all_tabs = requests.get(f"{CDP}/json").json()
        for t in all_tabs:
            if t["id"] not in known_ids and "UTERNAAZ34" in t.get("url",""):
                popup_tab = t
                print(f"  팝업 감지 ({i*0.5:.1f}초)")
                break
        if popup_tab:
            break
    if not popup_tab:
        print("팝업 미감지"); return
    time.sleep(2)

    # ── STEP 3: 일괄출력 클릭 ────────────────────────────────────────────────
    async with websockets.connect(popup_tab["webSocketDebuggerUrl"], ping_interval=None) as ws2:
        for _ in range(20):
            rs = await _eval(ws2, "document.readyState", cmd_id=1)
            if rs == "complete": break
            await asyncio.sleep(0.5)

        await _eval(ws2, "window.confirm = function(){ return true; };", cmd_id=29)
        print("  window.confirm 자동수락")

        ok2 = await pyclick_btn(ws2, popup_tab["id"], """
    (function(){
        var btns = Array.from(document.querySelectorAll('input[type=button],button,a'))
            .filter(function(el){
                return (el.value || el.innerText || '').trim() === '일괄출력';
            });
        return btns.find(function(el){
            var r = el.getBoundingClientRect();
            return r.width > 0 && r.height > 0;
        }) || null;
    })()
""", label="일괄출력", cmd_id=30)
        if not ok2:
            print("일괄출력 버튼 없음"); return
        print("일괄출력 클릭 완료")

    # ── STEP 4: ClipReport 완료 감지 (페이지 수 > 1) ────────────────────────
    print("  완료 감지 대기 (최대 2분)...")
    done = False
    for attempt in range(24):
        await asyncio.sleep(5)
        try:
            tabs_now = requests.get(f"{CDP}/json").json()
            pt2 = next((t for t in tabs_now if "UTERNAAZ34" in t.get("url","")), None)
            if not pt2: break
            async with websockets.connect(pt2["webSocketDebuggerUrl"],
                                          ping_interval=None, open_timeout=3) as ws_chk:
                await ws_chk.send(json.dumps({"id": 10, "method": "Runtime.enable"}))
                contexts = []
                deadline = asyncio.get_event_loop().time() + 2.0
                while asyncio.get_event_loop().time() < deadline:
                    try:
                        r = json.loads(await asyncio.wait_for(ws_chk.recv(), timeout=0.3))
                        if r.get("method") == "Runtime.executionContextCreated":
                            ctx = r["params"]["context"]
                            if ctx.get("id"):
                                contexts.append(ctx)
                    except asyncio.TimeoutError:
                        break
                if not contexts:
                    print(f"  {(attempt+1)*5}초 — 컨텍스트 없음"); continue
                for idx, ctx in enumerate(contexts):
                    await ws_chk.send(json.dumps({
                        "id": 100+idx, "method": "Runtime.evaluate",
                        "params": {"expression": """(function(){
    try {
        var el = document.querySelector('[id*="totalCountNumber"]');
        if (el) return el.value || el.innerText || null;
        var m = document.body.innerText.match(/(\\d+)\\s*\\/\\s*(\\d+)/);
        return m ? m[1]+'/'+m[2] : null;
    } catch(e){ return null; }
})()""", "contextId": ctx["id"], "returnByValue": True}}))
                pending = set(range(100, 100+len(contexts)))
                deadline2 = asyncio.get_event_loop().time() + 3.0
                while pending and asyncio.get_event_loop().time() < deadline2:
                    try:
                        r2 = json.loads(await asyncio.wait_for(ws_chk.recv(), timeout=0.5))
                        if r2.get("id") in pending:
                            pending.discard(r2["id"])
                            val = r2.get("result",{}).get("result",{}).get("value")
                            if val:
                                try:
                                    total = int(str(val).split("/")[-1].strip())
                                    if total > 1:
                                        print(f"  ✓ 완료! 페이지: {val} ({(attempt+1)*5}초)")
                                        done = True; break
                                except Exception:
                                    pass
                    except asyncio.TimeoutError:
                        break
            if done: break
            print(f"  {(attempt+1)*5}초 — 처리 중...")
        except Exception as e:
            pass
    if not done:
        print("  감지 실패 — 그대로 진행")

    # ── STEP 5: pdfDownLoad() 호출 + 다운로드 대기 ───────────────────────────
    tabs_now = requests.get(f"{CDP}/json").json()
    pt = next((t for t in tabs_now if "UTERNAAZ34" in t.get("url","")), None)
    if not pt:
        print("팝업 탭 없음"); return

    async with websockets.connect(pt["webSocketDebuggerUrl"], ping_interval=None) as ws_dl:
        # Network.enable
        await ws_dl.send(json.dumps({"id": 1, "method": "Network.enable"}))
        while True:
            r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=5))
            if r.get("id") == 1: break

        # Browser.setDownloadBehavior (시도)
        await ws_dl.send(json.dumps({"id": 2, "method": "Browser.setDownloadBehavior",
            "params": {"behavior": "allow", "downloadPath": SAVE_DIR, "eventsEnabled": True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=5))
            if r.get("id") == 2: break

        # Runtime.enable → sesw ctx 찾기
        contexts = await collect_contexts(ws_dl)
        ctx_id = await find_clipreport_ctx(ws_dl, contexts)
        if not ctx_id:
            print("ClipReport 컨텍스트 없음"); return
        print(f"  ClipReport ctx: {ctx_id}")

        # REPORT_KEY 추출
        report_key = await _eval(ws_dl, """(function(){
    var btn = Array.from(document.querySelectorAll('button'))
        .find(function(b){ return b.id.indexOf('re_pdf') === 0; });
    if (!btn) return null;
    var m = btn.id.match(/re_pdf(.+)/);
    return m ? m[1] : null;
})()""", cmd_id=60, ctx_id=ctx_id)
        print(f"  report_key: {report_key}")
        if not report_key:
            print("report_key 없음"); return

        # Downloads 폴더 현재 파일 목록 스냅샷
        before_files = set(os.listdir(DOWNLOADS))

        # pdfDownLoad() 호출
        result = await _eval(ws_dl, f"""(function(){{
    try {{
        m_reportHashMap['{report_key}'].pdfDownLoad();
        return 'called';
    }} catch(e) {{
        return 'error:' + e.message;
    }}
}})()""", cmd_id=70, ctx_id=ctx_id)
        print(f"  pdfDownLoad(): {result}")

        # 다운로드 완료 대기 (최대 60초)
        print("  다운로드 대기...")
        suggested_name = None
        dl_done = False
        deadline = asyncio.get_event_loop().time() + 60
        while asyncio.get_event_loop().time() < deadline:
            try:
                r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=0.5))
                method = r.get("method","")
                if method == "Browser.downloadWillBegin":
                    suggested_name = r["params"].get("suggestedFilename","종합소득세.pdf")
                    print(f"  다운로드 시작: {suggested_name}")
                elif method == "Browser.downloadProgress":
                    if r["params"].get("state") == "completed":
                        print("  다운로드 완료!")
                        dl_done = True; break
            except asyncio.TimeoutError:
                pass

        if not dl_done:
            print("  다운로드 이벤트 미감지 — 파일 직접 탐색")

        # Downloads 폴더에서 새 파일 찾기
        await asyncio.sleep(1)
        after_files = set(os.listdir(DOWNLOADS))
        new_files = [f for f in (after_files - before_files) if f.endswith(".pdf")]
        print(f"  새 파일: {new_files}")

        if not new_files:
            # suggestedFilename으로 직접 찾기
            fname = suggested_name or "종합소득세.pdf"
            if fname in after_files:
                new_files = [fname]

        if new_files:
            src = os.path.join(DOWNLOADS, new_files[0])
            dst = os.path.join(SAVE_DIR, new_files[0])
            shutil.move(src, dst)
            size_kb = os.path.getsize(dst) // 1024
            print(f"\n✓ PDF 저장 완료: {dst} ({size_kb}KB)")
        else:
            print(f"\n파일 없음 — Downloads 폴더 확인: {DOWNLOADS}")
            print(f"  현재 파일: {list(after_files)[:10]}")

asyncio.run(run())
