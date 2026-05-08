# -*- coding: utf-8 -*-
"""
납부서 테스트 - col[13] 보기 클릭
  환급 고객: 보기 -> 확인/확인/취소 다이얼로그 자동 처리 -> 스킵
  납부 고객: 보기 -> 출력 버튼 클릭 -> sesw 팝업 -> pdfDownLoad
"""
import requests, json, asyncio, websockets, pyautogui, time, os, shutil

CDP      = "http://localhost:9222"
NAME_COL = 6
NABU_COL = 13
SAVE_DIR = r"C:\Users\pc\종소세2026"
ROW_SEL  = '[id*="UTERNAAZ0Z31_wframe"] table tbody tr'

# 테스트할 행 인덱스 (0=첫번째 고객)
# 납부서 있는 고객이 다른 행이라면 아래 숫자 변경
TEST_ROW = 0

pyautogui.PAUSE = 0.15


def get_screen_xy(bounds, vp_w, vp_h, vx, vy):
    toolbar_h = bounds["height"] - vp_h
    return int(bounds["left"] + (bounds["width"] - vp_w) // 2 + vx), \
           int(bounds["top"]  + toolbar_h + vy)


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
    r2 = await _send(ws, {"id": 901, "method": "Browser.getWindowBounds",
                          "params": {"windowId": r["result"]["windowId"]}})
    bounds = r2["result"]["bounds"]
    vp = json.loads(await _eval(ws, "JSON.stringify({w:innerWidth,h:innerHeight})", cmd_id=902))
    return bounds, vp["w"], vp["h"]


async def dismiss_modal_sequence(ws, label=""):
    """
    WebSquare 모달 순서대로 처리:
    확인->확인->취소  (환급 고객 납부서 보기 후 나타나는 패턴)
    반환: 처리한 버튼 수
    """
    clicked = 0
    for i in range(5):
        await asyncio.sleep(0.5)
        # 확인 버튼 먼저
        r = await _eval(ws, """(function(){
    var b = Array.from(document.querySelectorAll('button,input[type=button]'))
        .find(function(b){
            var t = (b.innerText||b.value||'').trim();
            var r = b.getBoundingClientRect();
            return t === '확인' && r.width > 0;
        });
    if (b){ b.click(); return '확인'; }
    // 취소 버튼
    var c = Array.from(document.querySelectorAll('button,input[type=button]'))
        .find(function(b){
            var t = (b.innerText||b.value||'').trim();
            var r = b.getBoundingClientRect();
            return t === '취소' && r.width > 0;
        });
    if (c){ c.click(); return '취소'; }
    return null;
})()""", cmd_id=41 + i)
        if r:
            print(f"  [{label}] 모달 처리: {r} ({i+1}번째)")
            clicked += 1
            if r == '취소':
                break
        else:
            break
    return clicked


async def run():
    os.makedirs(SAVE_DIR, exist_ok=True)
    tabs = requests.get(f"{CDP}/json").json()
    main = next((t for t in tabs
                 if "hometax.go.kr" in t.get("url", "")
                 and "websquare.html" in t.get("url", "")
                 and "sesw." not in t.get("url", "")), None)
    if not main:
        print("홈택스 탭 없음"); return

    known_ids = set(t["id"] for t in tabs)

    # ── 0. col[13] 전체 행 스캔 ──
    print("=== col[13] 스캔 ===")
    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        scan = await _eval(ws, f"""JSON.stringify(
    Array.from(document.querySelectorAll('{ROW_SEL}'))
        .filter(function(tr){{ return tr.querySelectorAll('td').length >= 13; }})
        .map(function(tr, i){{
            var tds = tr.querySelectorAll('td');
            var name = tds[{NAME_COL}] ? tds[{NAME_COL}].innerText.trim() : '';
            var cell = tds[{NABU_COL}];
            var txt  = cell ? cell.innerText.trim().replace(/\\s+/g,' ').slice(0,20) : '';
            var btn  = cell && cell.querySelector('a,button,input[type=button]') ? 'Y' : 'N';
            return {{i:i, name:name, col13:txt, btn:btn}};
        }})
)""", cmd_id=1)
        rows = json.loads(scan) if scan else []
        for row in rows:
            print(f"  row{row['i']}: {row['name']} | col13={row['col13']} | 버튼={row['btn']}")

    # TEST_ROW 행의 고객명 확인
    target = next((r for r in rows if r["i"] == TEST_ROW), None)
    if not target:
        print(f"row {TEST_ROW} 없음"); return
    name = target["name"]
    has_btn = target["btn"] == "Y"
    print(f"\n테스트 대상: row{TEST_ROW} [{name}] 버튼={has_btn}")

    if not has_btn:
        print("버튼 없음 - TEST_ROW 를 납부서 있는 행으로 변경하세요"); return

    # ── 1. 납부서 보기 클릭 ──
    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        bounds, vp_w, vp_h = await get_win_info(ws, main["id"])
        coords = await _eval(ws, f"""(function(){{
    var rows = Array.from(document.querySelectorAll('{ROW_SEL}'))
        .filter(function(tr){{ return tr.querySelectorAll('td').length >= 13; }});
    var cell = rows[{TEST_ROW}] && rows[{TEST_ROW}].querySelectorAll('td')[{NABU_COL}];
    if (!cell) return null;
    var btn = cell.querySelector('a,button,input[type=button]');
    if (!btn) return null;
    btn.scrollIntoView({{block:'center',behavior:'instant'}});
    var r = btn.getBoundingClientRect();
    if (!r.width || !r.height) return null;
    return JSON.stringify({{x:r.left+r.width/2, y:r.top+r.height/2}});
}})()""", cmd_id=20)
        if not coords:
            print("버튼 좌표 없음"); return
        c = json.loads(coords)
        sx, sy = get_screen_xy(bounds, vp_w, vp_h, c["x"], c["y"])
        print(f"\n납부서 보기 클릭 ({sx},{sy})")
        pyautogui.click(sx, sy)

    # ── 2. 1초 대기 후 상태 확인 ──
    await asyncio.sleep(1.0)

    # 새 탭 확인
    all_tabs = requests.get(f"{CDP}/json").json()
    new_tabs = [t for t in all_tabs if t["id"] not in known_ids]
    print(f"\n새 탭: {len(new_tabs)}개")
    for t in new_tabs:
        print(f"  {t.get('url','')[:80]}")

    # 메인 탭 모달 확인
    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        modal_info = await _eval(ws, """(function(){
    var btns = Array.from(document.querySelectorAll('button,input[type=button]'))
        .filter(function(b){
            var r = b.getBoundingClientRect();
            return r.width > 0;
        })
        .map(function(b){ return (b.innerText||b.value||'').trim(); })
        .filter(function(t){ return t.length > 0; })
        .slice(0, 10);
    return JSON.stringify(btns);
})()""", cmd_id=30)
        print(f"메인 탭 가시 버튼: {modal_info}")

    # ── 3. 환급 케이스: 모달 처리 ──
    #    보기 직후 메인 탭에 확인/취소 버튼이 보이면 환급 고객
    btns = json.loads(modal_info) if modal_info else []
    if '확인' in btns or '취소' in btns:
        print("\n[환급 케이스] 모달 처리 중...")
        async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
            n = await dismiss_modal_sequence(ws, name)
        print(f"  -> {n}개 처리 완료, 납부서 없음 (환급)")
        return

    # ── 4. 새 탭이 없는 경우: 출력 버튼 탐색 ──
    if not new_tabs:
        print("\n새 탭 없음 - 출력 버튼 탐색...")
        async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
            bounds, vp_w, vp_h = await get_win_info(ws, main["id"])
            out_coords = await _eval(ws, """(function(){
    var btns = Array.from(document.querySelectorAll('input[type=button],button,a'));
    var btn = btns.find(function(b){
        var t = (b.value||b.innerText||'').trim();
        var r = b.getBoundingClientRect();
        return r.width>0 && r.height>0 && (t==='출력'||t==='인쇄'||t==='PDF'||t==='저장');
    });
    if (!btn) return null;
    btn.scrollIntoView({block:'center',behavior:'instant'});
    var r = btn.getBoundingClientRect();
    return JSON.stringify({x:r.left+r.width/2, y:r.top+r.height/2,
        text:(btn.value||btn.innerText||'').trim()});
})()""", cmd_id=35)
            if out_coords:
                oc = json.loads(out_coords)
                osx, osy = get_screen_xy(bounds, vp_w, vp_h, oc["x"], oc["y"])
                print(f"  출력 버튼 [{oc['text']}] 클릭 ({osx},{osy})")
                pyautogui.click(osx, osy)
                await asyncio.sleep(1)
                all_tabs = requests.get(f"{CDP}/json").json()
                new_tabs = [t for t in all_tabs if t["id"] not in known_ids]
                print(f"  출력 후 새 탭: {len(new_tabs)}개")
                for t in new_tabs:
                    print(f"    {t.get('url','')[:80]}")
            else:
                print("  출력 버튼 없음")
                return

    if not new_tabs:
        print("팝업 없음 - 흐름 확인 필요"); return

    # ── 5. 팝업에서 pdfDownLoad ──
    popup = new_tabs[0]
    known_ids.add(popup["id"])
    time.sleep(2)
    print(f"\n팝업 연결: {popup.get('url','')[:60]}")

    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws_dl:
        for _ in range(20):
            rs = await _eval(ws_dl, "document.readyState", cmd_id=1)
            if rs == "complete":
                break
            await asyncio.sleep(0.5)

        await ws_dl.send(json.dumps({"id": 2, "method": "Network.enable"}))
        while True:
            r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=5))
            if r.get("id") == 2:
                break
        await ws_dl.send(json.dumps({"id": 3, "method": "Browser.setDownloadBehavior",
            "params": {"behavior": "allow", "downloadPath": SAVE_DIR, "eventsEnabled": True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=5))
            if r.get("id") == 3:
                break

        rkey = await _eval(ws_dl, """(function(){
    var btn = Array.from(document.querySelectorAll('button'))
        .find(function(b){ return b.id.indexOf('re_pdf') === 0; });
    if (!btn) return null;
    var m = btn.id.match(/re_pdf(.+)/);
    return m ? m[1] : null;
})()""", cmd_id=60)

        if not rkey:
            await ws_dl.send(json.dumps({"id": 10, "method": "Runtime.enable"}))
            contexts = []
            deadline = asyncio.get_event_loop().time() + 2.5
            while asyncio.get_event_loop().time() < deadline:
                try:
                    r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=0.3))
                    if r.get("method") == "Runtime.executionContextCreated":
                        contexts.append(r["params"]["context"])
                except asyncio.TimeoutError:
                    break
            for c2 in contexts:
                cid = c2["id"]
                rkey2 = await _eval(ws_dl, """(function(){
    var btn = Array.from(document.querySelectorAll('button'))
        .find(function(b){ return b.id.indexOf('re_pdf') === 0; });
    if(!btn) return null;
    var m = btn.id.match(/re_pdf(.+)/);
    return m?m[1]:null;
})()""", cmd_id=60 + cid % 50, ctx_id=cid)
                if rkey2:
                    rkey = rkey2; break

        if not rkey:
            # 팝업 버튼 목록 출력
            btns2 = await _eval(ws_dl, """JSON.stringify(
    Array.from(document.querySelectorAll('button,input[type=button],a'))
        .filter(function(b){ var r=b.getBoundingClientRect(); return r.width>0; })
        .slice(0,10)
        .map(function(b){ return {id:b.id, text:(b.innerText||b.value||'').trim().slice(0,20)}; })
)""", cmd_id=99)
            print(f"report_key 없음. 버튼 목록: {btns2}")
            return

        print(f"report_key: {rkey[:25]}...")
        before = os.path.getmtime(os.path.join(SAVE_DIR, "종합소득세.pdf")) \
                 if os.path.exists(os.path.join(SAVE_DIR, "종합소득세.pdf")) else 0

        res = await _eval(ws_dl, f"""(function(){{
    try {{ m_reportHashMap['{rkey}'].pdfDownLoad(); return 'ok'; }}
    catch(e) {{ return 'err:' + e.message; }}
}})()""", cmd_id=70)
        print(f"pdfDownLoad: {res}")

        print("다운로드 대기 (최대 30초)...")
        suggested = "종합소득세.pdf"
        deadline = asyncio.get_event_loop().time() + 30
        while asyncio.get_event_loop().time() < deadline:
            try:
                r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=0.5))
                if r.get("method") == "Browser.downloadWillBegin":
                    suggested = r["params"].get("suggestedFilename", suggested)
                    print(f"파일명: {suggested}")
                elif r.get("method") == "Browser.downloadProgress":
                    if r["params"].get("state") == "completed":
                        print("완료!"); break
            except asyncio.TimeoutError:
                pass

        await asyncio.sleep(1)
        for fn in [suggested, "종합소득세.pdf"]:
            fp = os.path.join(SAVE_DIR, fn)
            if os.path.exists(fp) and os.path.getmtime(fp) > before + 0.5:
                dst = os.path.join(SAVE_DIR, f"{name}_납부서_test.pdf")
                shutil.move(fp, dst)
                print(f"\n저장: {dst} ({os.path.getsize(dst)//1024}KB)")
                break
        else:
            print("\n파일 없음")

    try:
        requests.get(f"{CDP}/json/close/{popup['id']}")
    except:
        pass


asyncio.run(run())
