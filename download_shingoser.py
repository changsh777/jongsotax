"""
홈택스 종합소득세 신고서 일괄 PDF 다운로드
- 목록 전체 행 순서대로 처리
- 고객명_종합소득세.pdf 저장
- 이미 있으면 스킵
"""
import requests, json, asyncio, websockets, pyautogui, time, os, shutil, sys

# ═══════════════════════ 설정 ════════════════════════════════════════════════
CDP        = "http://localhost:9222"
NAME_COL   = 6    # 고객명 열 인덱스 (0부터)
ACPT_COL   = 10   # 접수번호 열 인덱스 (클릭 대상)
NAS_BASE   = r"Z:\종소세2026\고객"
LOCAL_BASE = r"C:\Users\pc\종소세2026"
ROW_SEL    = '[id*="UTERNAAZ0Z31_wframe"] table tbody tr'
# ════════════════════════════════════════════════════════════════════════════


def get_save_dir(name):
    if os.path.isdir(NAS_BASE):
        d = os.path.join(NAS_BASE, name)
    else:
        d = LOCAL_BASE
    os.makedirs(d, exist_ok=True)
    return d

pyautogui.PAUSE = 0.15

# ── 헬퍼 ─────────────────────────────────────────────────────────────────────

def get_screen_xy(bounds, vp_w, vp_h, vx, vy):
    toolbar_h = bounds["height"] - vp_h
    sx = bounds["left"] + (bounds["width"] - vp_w) // 2 + vx
    sy = bounds["top"]  + toolbar_h + vy
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
    r2 = await _send(ws, {"id": 901, "method": "Browser.getWindowBounds",
                          "params": {"windowId": r["result"]["windowId"]}})
    bounds = r2["result"]["bounds"]
    vp = json.loads(await _eval(ws, "JSON.stringify({w:innerWidth,h:innerHeight})", cmd_id=902))
    return bounds, vp["w"], vp["h"]

async def pyclick(ws, target_id, selector_js, label, cmd_id, scroll=False):
    bounds, vp_w, vp_h = await get_win_info(ws, target_id)
    scroll_js = "el.scrollIntoView({block:'center',behavior:'instant'});" if scroll else ""
    coords = await _eval(ws, f"""(function(){{
    var el = ({selector_js});
    if (!el) return null;
    {scroll_js}
    var r = el.getBoundingClientRect();
    if (!r.width || !r.height) return null;
    return JSON.stringify({{x:r.left+r.width/2, y:r.top+r.height/2}});
}})()""", cmd_id=cmd_id)
    if not coords:
        print(f"  [{label}] 좌표 없음"); return False
    c = json.loads(coords)
    sx, sy = get_screen_xy(bounds, vp_w, vp_h, c["x"], c["y"])
    print(f"  [{label}] 클릭 → 스크린({sx},{sy})")
    pyautogui.click(sx, sy)
    return True

async def collect_contexts(ws, secs=2.5):
    await ws.send(json.dumps({"id": 11, "method": "Runtime.enable"}))
    ctxs = []
    t = asyncio.get_event_loop().time() + secs
    while asyncio.get_event_loop().time() < t:
        try:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=0.3))
            if r.get("method") == "Runtime.executionContextCreated":
                ctxs.append(r["params"]["context"])
        except asyncio.TimeoutError:
            break
    return ctxs

async def find_cr_ctx(ws, ctxs):
    for c in ctxs:
        if "sesw.hometax.go.kr" not in c.get("origin",""): continue
        n = await _eval(ws, "document.querySelectorAll('button').length",
                        cmd_id=200+c["id"]%100, ctx_id=c["id"])
        if n and int(n) > 3:
            return c["id"]
    return None

# ── 행 목록 읽기 ──────────────────────────────────────────────────────────────

async def get_rows(ws):
    raw = await _eval(ws, f"""JSON.stringify(
    Array.from(document.querySelectorAll('{ROW_SEL}'))
        .filter(function(tr){{ return tr.querySelectorAll('td').length >= 13; }})
        .map(function(tr){{
            var tds = tr.querySelectorAll('td');
            return {{
                name: (tds[{NAME_COL}]||{{innerText:''}}).innerText.trim(),
                acpt: (tds[{ACPT_COL}]||{{innerText:''}}).innerText.trim()
            }};
        }})
)""", cmd_id=5)
    return json.loads(raw) if raw else []

# ── 1건 처리 ─────────────────────────────────────────────────────────────────

async def process_row(main_tab, row_idx, name, acpt, known_ids):
    sdir = get_save_dir(name)
    dst  = os.path.join(sdir, f"{name}_종합소득세.pdf")
    if os.path.exists(dst):
        print(f"  [{name}] 이미 있음 — 스킵")
        return True

    print(f"\n{'='*60}")
    print(f"  처리: {name} / 접수번호: {acpt} (행 {row_idx})")

    # STEP 1: 접수번호 클릭
    async with websockets.connect(main_tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        ok = await pyclick(ws, main_tab["id"], f"""(function(){{
    var rows = Array.from(document.querySelectorAll('{ROW_SEL}'))
        .filter(function(tr){{ return tr.querySelectorAll('td').length >= 13; }});
    var cell = rows[{row_idx}] && rows[{row_idx}].querySelectorAll('td')[{ACPT_COL}];
    return cell && cell.querySelector('a,input[type=button],button');
}})()""", label="접수번호", cmd_id=20)
        if not ok:
            print("  접수번호 버튼 없음"); return False

    # STEP 2: 팝업 대기
    print("  팝업 대기...")
    popup = None
    for _ in range(30):
        time.sleep(0.5)
        for t in requests.get(f"{CDP}/json").json():
            if t["id"] not in known_ids and "UTERNAAZ34" in t.get("url",""):
                popup = t; break
        if popup: break
    if not popup:
        print("  팝업 미감지"); return False
    time.sleep(2)

    # 접수증 등 다른 팝업 닫기 + UTERNAAZ34 앞으로
    for t in requests.get(f"{CDP}/json").json():
        if t["id"] in known_ids: continue
        if "UTERNAAZ34" in t.get("url",""): continue
        if "hometax.go.kr" in t.get("url",""):
            print(f"  부가 팝업 닫음: {t['url'][:60]}")
            try: requests.get(f"{CDP}/json/close/{t['id']}")
            except: pass
    # UTERNAAZ34 활성화 (앞으로 가져오기)
    try:
        requests.post(f"{CDP}/json/activate/{popup['id']}")
    except: pass
    time.sleep(0.5)

    # STEP 3: 일괄출력 클릭
    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws2:
        for _ in range(20):
            if await _eval(ws2, "document.readyState", cmd_id=1) == "complete": break
            await asyncio.sleep(0.5)
        await _eval(ws2, "window.confirm=function(){return true;};", cmd_id=29)
        ok2 = await pyclick(ws2, popup["id"], """(function(){
    return Array.from(document.querySelectorAll('input[type=button],button,a'))
        .filter(function(el){return (el.value||el.innerText||'').trim()==='일괄출력';})
        .find(function(el){var r=el.getBoundingClientRect();return r.width>0&&r.height>0;});
})()""", label="일괄출력", cmd_id=30)
        if not ok2:
            print("  일괄출력 버튼 없음"); return False

    # STEP 4: 완료 감지 (페이지 수 > 1)
    print("  완료 감지 대기...")
    done = False
    for attempt in range(24):
        await asyncio.sleep(5)
        try:
            pt2 = next((t for t in requests.get(f"{CDP}/json").json()
                        if "UTERNAAZ34" in t.get("url","")), None)
            if not pt2: break
            async with websockets.connect(pt2["webSocketDebuggerUrl"],
                                          ping_interval=None, open_timeout=3) as wc:
                await wc.send(json.dumps({"id": 10, "method": "Runtime.enable"}))
                ctxs = []
                t0 = asyncio.get_event_loop().time() + 2.0
                while asyncio.get_event_loop().time() < t0:
                    try:
                        r = json.loads(await asyncio.wait_for(wc.recv(), timeout=0.3))
                        if r.get("method") == "Runtime.executionContextCreated":
                            ctxs.append(r["params"]["context"])
                    except asyncio.TimeoutError:
                        break
                for i, ctx in enumerate(ctxs):
                    await wc.send(json.dumps({"id": 100+i, "method": "Runtime.evaluate",
                        "params": {"expression": """(function(){
    try{
        var el=document.querySelector('[id*="totalCountNumber"]');
        if(el) return el.value||el.innerText||null;
        var m=document.body.innerText.match(/(\\d+)\\s*\\/\\s*(\\d+)/);
        return m?m[1]+'/'+m[2]:null;
    }catch(e){return null;}
})()""", "contextId": ctx["id"], "returnByValue": True}}))
                pending = set(range(100, 100+len(ctxs)))
                t1 = asyncio.get_event_loop().time() + 3.0
                while pending and asyncio.get_event_loop().time() < t1:
                    try:
                        r = json.loads(await asyncio.wait_for(wc.recv(), timeout=0.5))
                        if r.get("id") in pending:
                            pending.discard(r["id"])
                            val = r.get("result",{}).get("result",{}).get("value")
                            if val:
                                try:
                                    total = int(str(val).split("/")[-1].strip())
                                    if total > 1:
                                        print(f"  ✓ 완료! 페이지: {val} ({(attempt+1)*5}초)")
                                        done = True; break
                                except: pass
                    except asyncio.TimeoutError:
                        break
            if done: break
            if (attempt+1) % 6 == 0:
                print(f"  {(attempt+1)*5}초 대기 중...")
        except Exception:
            pass
    if not done:
        print("  감지 실패 — 그대로 진행")

    # STEP 5: pdfDownLoad() + 다운로드 대기
    pt = next((t for t in requests.get(f"{CDP}/json").json()
               if "UTERNAAZ34" in t.get("url","")), None)
    if not pt:
        print("  팝업 사라짐"); return False

    async with websockets.connect(pt["webSocketDebuggerUrl"], ping_interval=None) as ws_dl:
        # Network + 다운로드 설정
        await ws_dl.send(json.dumps({"id": 1, "method": "Network.enable"}))
        while True:
            r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=5))
            if r.get("id") == 1: break
        await ws_dl.send(json.dumps({"id": 2, "method": "Browser.setDownloadBehavior",
            "params": {"behavior": "allow", "downloadPath": sdir, "eventsEnabled": True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=5))
            if r.get("id") == 2: break

        # ClipReport 컨텍스트
        ctxs = await collect_contexts(ws_dl)
        cr_ctx = await find_cr_ctx(ws_dl, ctxs)
        if not cr_ctx:
            print("  ClipReport 컨텍스트 없음"); return False

        # report_key 추출
        rkey = await _eval(ws_dl, """(function(){
    var btn=Array.from(document.querySelectorAll('button'))
        .find(function(b){return b.id.indexOf('re_pdf')===0;});
    if(!btn) return null;
    var m=btn.id.match(/re_pdf(.+)/);
    return m?m[1]:null;
})()""", cmd_id=60, ctx_id=cr_ctx)
        if not rkey:
            print("  report_key 없음"); return False

        # 기존 파일 mtime 스냅샷
        tmp_pdf = os.path.join(sdir, "종합소득세.pdf")
        before_mtime = os.path.getmtime(tmp_pdf) if os.path.exists(tmp_pdf) else 0

        # pdfDownLoad 호출
        res = await _eval(ws_dl, f"""(function(){{
    try{{m_reportHashMap['{rkey}'].pdfDownLoad();return 'ok';}}
    catch(e){{return 'err:'+e.message;}}
}})()""", cmd_id=70, ctx_id=cr_ctx)
        print(f"  pdfDownLoad: {res}")

        # 다운로드 완료 대기 (최대 60초)
        print("  다운로드 대기...")
        dl_done = False
        t_dl = asyncio.get_event_loop().time() + 60
        while asyncio.get_event_loop().time() < t_dl:
            try:
                r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=0.5))
                if r.get("method") == "Browser.downloadProgress":
                    if r["params"].get("state") == "completed":
                        print("  다운로드 완료!")
                        dl_done = True; break
            except asyncio.TimeoutError:
                pass

        # 파일 확인 + 이름 변경 — 반드시 종합소득세.pdf만 체크 (다른 파일 잘못 rename 방지)
        await asyncio.sleep(1)
        new_src = None
        # 1순위: sdir/종합소득세.pdf (NAS 고객 폴더)
        if os.path.exists(tmp_pdf) and os.path.getmtime(tmp_pdf) > before_mtime + 0.5:
            new_src = tmp_pdf
        # 2순위: Downloads/종합소득세.pdf
        if not new_src:
            dld_pdf = os.path.join(os.environ.get("USERPROFILE",""), "Downloads", "종합소득세.pdf")
            if os.path.exists(dld_pdf) and os.path.getmtime(dld_pdf) > time.time() - 30:
                new_src = dld_pdf

        if new_src:
            if new_src != dst:
                shutil.move(new_src, dst)
            size_kb = os.path.getsize(dst) // 1024
            print(f"  저장 OK: {dst} ({size_kb}KB)")
        else:
            print(f"  파일 없음 (before_mtime={before_mtime:.1f})")
            return False

    # STEP 6: 팝업 닫기
    try:
        requests.get(f"{CDP}/json/close/{pt['id']}")
        print("  팝업 닫음")
    except Exception:
        pass
    await asyncio.sleep(1)
    return True

# ── 다음 페이지 클릭 ──────────────────────────────────────────────────────────

async def dismiss_alert(ws, target_id, timeout=3.0):
    """WebSquare '조회가 완료되었습니다' 알림의 확인만 pyautogui 클릭
    - 팝업 헤더 확인(조회 실행) 버튼과 혼동 방지: 메시지 텍스트 근처 버튼만 찾음"""
    deadline = asyncio.get_event_loop().time() + timeout
    while asyncio.get_event_loop().time() < deadline:
        coords = await _eval(ws, """(function(){
    var nodes = Array.from(document.querySelectorAll('*')).filter(function(el){
        return el.children.length === 0 &&
               el.textContent.indexOf('조회가 완료되었습니다') >= 0;
    });
    for(var ni=0; ni<nodes.length; ni++){
        var p = nodes[ni].parentElement;
        for(var d=0; d<10 && p && p!==document.body; d++){
            var btns = p.querySelectorAll('button,input[type=button]');
            for(var bi=0; bi<btns.length; bi++){
                var b = btns[bi];
                var t=(b.innerText||b.value||'').trim();
                var r=b.getBoundingClientRect();
                if(t==='확인' && r.width>0 && r.height>0){
                    return JSON.stringify({x:r.left+r.width/2, y:r.top+r.height/2});
                }
            }
            p = p.parentElement;
        }
    }
    return null;
})()""", cmd_id=41)
        if coords:
            c = json.loads(coords)
            bounds, vp_w, vp_h = await get_win_info(ws, target_id)
            sx, sy = get_screen_xy(bounds, vp_w, vp_h, c["x"], c["y"])
            print(f"  alert 확인 pyautogui 클릭 ({sx},{sy})")
            pyautogui.click(sx, sy)
            return True
        await asyncio.sleep(0.3)
    return False

async def click_next_page(ws, main_tab):
    """다음 페이지 버튼 pyautogui 클릭. 없으면 False"""
    bounds, vp_w, vp_h = await get_win_info(ws, main_tab["id"])
    coords = await _eval(ws, """(function(){
    var all = Array.from(document.querySelectorAll('a,input[type=button],button,td'));
    for (var el of all) {
        var t = (el.innerText||el.value||el.title||'').trim();
        var r = el.getBoundingClientRect();
        if (r.width>0 && r.height>0 && (t==='다음' || t==='>') &&
            r.top > window.innerHeight * 0.5) {
            el.scrollIntoView({block:'center',behavior:'instant'});
            r = el.getBoundingClientRect();
            return JSON.stringify({x:r.left+r.width/2, y:r.top+r.height/2, text:t});
        }
    }
    return null;
})()""", cmd_id=40)
    if not coords:
        return False
    c = json.loads(coords)
    sx, sy = get_screen_xy(bounds, vp_w, vp_h, c["x"], c["y"])
    print(f"  다음 페이지 클릭({c.get('text')}) → ({sx},{sy})")
    pyautogui.click(sx, sy)
    # 페이지 전환 후 WebSquare 알림 자동 닫기
    await asyncio.sleep(1)
    await dismiss_alert(ws, main_tab["id"], timeout=5)
    return True

# ── 메인 루프 ─────────────────────────────────────────────────────────────────

async def run():
    nas_ok = os.path.isdir(NAS_BASE)
    base   = NAS_BASE if nas_ok else LOCAL_BASE
    print(f"저장 경로: {'NAS ' + NAS_BASE if nas_ok else '로컬 ' + LOCAL_BASE}")

    # 로컬 C:\Users\pc\종소세2026 에 이전 파일이 있으면 NAS 고객 폴더로 이동
    if nas_ok and os.path.isdir(LOCAL_BASE):
        moved = 0
        for fn in os.listdir(LOCAL_BASE):
            if fn.endswith("_종합소득세.pdf"):
                name = fn.replace("_종합소득세.pdf", "")
                sdir = get_save_dir(name)
                src  = os.path.join(LOCAL_BASE, fn)
                dst  = os.path.join(sdir, fn)
                if not os.path.exists(dst):
                    shutil.move(src, dst)
                    moved += 1
                    print(f"  이동: {fn} → {dst}")
        if moved:
            print(f"  로컬→NAS 이동 완료: {moved}건")

    tabs = requests.get(f"{CDP}/json").json()
    main = next((t for t in tabs if "hometax.go.kr" in t.get("url","")
                 and "websquare.html" in t.get("url","")
                 and "sesw." not in t.get("url","")), None)
    if not main:
        print("홈택스 탭 없음"); return

    known_ids = {t["id"] for t in tabs}
    page_num      = 1
    total_ok      = 0
    total_fail    = 0
    seen_first    = set()   # 각 페이지 첫 행 이름 — 루프 감지용

    # 스크립트 시작 시 이미 조회된 상태일 경우 초기 alert 처리
    print("초기 alert 확인...")
    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        await dismiss_alert(ws, main["id"], timeout=3.0)
    await asyncio.sleep(0.5)

    while True:
        # 현재 페이지 행 읽기
        async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
            rows = await get_rows(ws)

        if not rows:
            print("행 없음 — 종료"); break

        # ── 루프 감지: 첫 행 이름이 이미 처리한 페이지와 같으면 종료 ──────
        first_name = rows[0]["name"] if rows else ""
        if first_name and first_name in seen_first:
            print(f"\n페이지 루프 감지 ({first_name} 재등장) — 전체 완료")
            break
        if first_name:
            seen_first.add(first_name)

        print(f"\n{'='*60}")
        print(f"  페이지 {page_num} — {len(rows)}건")

        page_new = 0
        for idx, row in enumerate(rows):
            name = row["name"]
            acpt = row["acpt"]
            if not name:
                print(f"  행{idx}: 이름 없음 — 스킵"); continue
            sdir = get_save_dir(name)
            dst  = os.path.join(sdir, f"{name}_종합소득세.pdf")
            if os.path.exists(dst):
                print(f"  [{name}] 이미 있음 — 스킵"); continue
            page_new += 1
            try:
                ok = await process_row(main, idx, name, acpt, known_ids)
                if ok: total_ok += 1
                else:  total_fail += 1
            except Exception as e:
                print(f"  [{name}] 오류: {e}")
                total_fail += 1
            await asyncio.sleep(2)

        if page_new == 0:
            print("  이번 페이지 신규 없음 - 완료")
            break

        # 모든 행이 스킵이면 다음 페이지 의미 없어도 진행 (마지막 페이지 체크)
        # 다음 페이지로 이동
        async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
            has_next = await click_next_page(ws, main)

        if not has_next:
            print(f"\n마지막 페이지 ({page_num}페이지) — 다음 버튼 없음")
            break

        page_num += 1
        await asyncio.sleep(3)  # 페이지 로딩 대기

    print(f"\n{'='*60}")
    print(f"전체 완료! 성공 {total_ok}건 / 실패 {total_fail}건")
    pdfs = []
    for root, _, files in os.walk(base):
        for f in files:
            if "_종합소득세" in f and f.endswith(".pdf"):
                pdfs.append(os.path.join(root, f))
    print(f"신고서 PDF {len(pdfs)}개")
    for p in sorted(pdfs):
        print(f"  {p}")

asyncio.run(run())
