# -*- coding: utf-8 -*-
# 홈택스 접수증 일괄 PDF 다운로드
# col[12] 보기 -> {name}_접수증.pdf
# 저장: Z:\종소세2026\고객\{name}_{birthdate}\발송용\  (root=최신, _archive=이전)
import requests, json, asyncio, websockets, pyautogui, time, os, shutil, traceback, unicodedata
from datetime import datetime

CDP        = "http://localhost:9222"
NAME_COL   = 6
JEUP_COL   = 12
NAS_BASE   = r"Z:\종소세2026\고객"
LOCAL_BASE = r"C:\Users\pc\종소세2026"
ROW_SEL    = '[id*="UTERNAAZ0Z31_wframe"] table tbody tr'

pyautogui.PAUSE = 0.15


def _nfc(s):
    return unicodedata.normalize("NFC", str(s))


def find_balsong_dir(name):
    """NAS에서 {name}_XXXXXX 폴더 찾아 발송용 서브폴더 반환. 없으면 LOCAL_BASE."""
    name_nfc = _nfc(name)
    if os.path.isdir(NAS_BASE):
        hits = [d for d in os.listdir(NAS_BASE)
                if _nfc(d).split("_")[0] == name_nfc
                and len(_nfc(d).split("_")) >= 2
                and os.path.isdir(os.path.join(NAS_BASE, d))]
        if len(hits) == 1:
            balsong = os.path.join(NAS_BASE, hits[0], "발송용")
            os.makedirs(balsong, exist_ok=True)
            return balsong
        elif len(hits) > 1:
            print(f"  !! 동명이인 {name}: {hits} — 로컬 저장")
    os.makedirs(LOCAL_BASE, exist_ok=True)
    return LOCAL_BASE


def archive_if_exists(dst):
    """dst 파일이 있으면 _archive 폴더로 타임스탬프 붙여 이동 (root=최신 원칙)"""
    if not os.path.exists(dst):
        return
    archive_dir = os.path.join(os.path.dirname(dst), "_archive")
    os.makedirs(archive_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(dst))
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    shutil.move(dst, os.path.join(archive_dir, f"{stem}_{ts}{ext}"))
    print(f"  아카이브: {stem}_{ts}{ext}")


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


async def get_rows(ws):
    raw = await _eval(ws, f"""JSON.stringify(
    Array.from(document.querySelectorAll('{ROW_SEL}'))
        .filter(function(tr){{ return tr.querySelectorAll('td').length >= 13; }})
        .map(function(tr){{
            var tds = tr.querySelectorAll('td');
            return {{name:(tds[{NAME_COL}]||{{}}).innerText&&tds[{NAME_COL}].innerText.trim()}};
        }})
)""", cmd_id=5)
    return json.loads(raw) if raw else []


async def click_confirm_once(ws, target_id, timeout=5.0):
    """WebSquare '조회가 완료되었습니다' 알림의 확인만 pyautogui 클릭
    - 팝업 헤더의 확인(조회 실행) 버튼과 혼동 방지: 메시지 텍스트 근처 버튼만 찾음"""
    deadline = asyncio.get_event_loop().time() + timeout
    while asyncio.get_event_loop().time() < deadline:
        coords = await _eval(ws, """(function(){
    // "조회가 완료되었습니다" 텍스트를 포함하는 리프 요소 탐색
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
            print(f"  [alert] 확인 pyautogui 클릭 ({sx},{sy})")
            pyautogui.click(sx, sy)
            return True
        await asyncio.sleep(0.3)
    return False


async def click_next_page(main_tab):
    """다음 버튼을 JS .click()으로 클릭 — 팝업 오버레이 영향 없음"""
    async with websockets.connect(main_tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        r = await _eval(ws, """(function(){
    var all = Array.from(document.querySelectorAll('a,input[type=button],button,td'));
    for (var el of all){
        var t = (el.innerText||el.value||'').trim();
        var r = el.getBoundingClientRect();
        if (r.width>0 && r.height>0 && (t==='다음'||t==='>') && r.top > window.innerHeight*0.5){
            el.click();
            return '클릭:' + t;
        }
    }
    return null;
})()""", cmd_id=40)
        if not r:
            return False
        print(f"  다음 페이지: {r}")

    # 페이지 로딩 대기 (alert 처리 없음 — 버튼 클릭은 팝업 무관하게 동작)
    await asyncio.sleep(3)
    return True


# ── 1건 처리 ─────────────────────────────────────────────────────────────────

async def process_jeupsujeung(main_tab, row_idx, name, known_ids):
    sdir = find_balsong_dir(name)
    dst  = os.path.join(sdir, f"{name}_접수증.pdf")

    print(f"\n{'='*50} [{name}] 접수증")

    # 이미 파일 있으면 스킵
    if os.path.exists(dst):
        print(f"  [{name}] 접수증 이미 존재 — 스킵")
        return True

    # STEP 1: 접수증 보기 버튼 클릭 (pyautogui - 새 팝업 창 오픈 필요)
    async with websockets.connect(main_tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 잔여 alert 안전망 (1.5초만 — 메인 alert는 run()에서 처리)
        await click_confirm_once(ws, main_tab["id"], timeout=1.5)
        bounds, vp_w, vp_h = await get_win_info(ws, main_tab["id"])
        coords = await _eval(ws, f"""(function(){{
    var rows = Array.from(document.querySelectorAll('{ROW_SEL}'))
        .filter(function(tr){{ return tr.querySelectorAll('td').length >= 13; }});
    var cell = rows[{row_idx}] && rows[{row_idx}].querySelectorAll('td')[{JEUP_COL}];
    if (!cell) return null;
    var btn = cell.querySelector('a,input[type=button],button');
    if (!btn) return null;
    btn.scrollIntoView({{block:'center',behavior:'instant'}});
    var r = btn.getBoundingClientRect();
    if (!r.width || !r.height) return null;
    return JSON.stringify({{x:r.left+r.width/2, y:r.top+r.height/2}});
}})()""", cmd_id=20)
        if not coords:
            print("  접수증 버튼 없음"); return False
        c = json.loads(coords)
        sx, sy = get_screen_xy(bounds, vp_w, vp_h, c["x"], c["y"])
        print(f"  클릭 ({sx},{sy})")
        pyautogui.click(sx, sy)

    # STEP 2: sesw 팝업 대기 (최대 15초)
    print("  팝업 대기...")
    popup = None
    for _ in range(30):
        time.sleep(0.5)
        for t in requests.get(f"{CDP}/json").json():
            if t["id"] not in known_ids and "sesw.hometax.go.kr" in t.get("url", ""):
                popup = t; break
        if popup: break
    if not popup:
        print("  팝업 미감지"); return False
    known_ids.add(popup["id"])
    time.sleep(2)

    # STEP 3: pdfDownLoad
    async with websockets.connect(popup["webSocketDebuggerUrl"], ping_interval=None) as ws_dl:
        for _ in range(20):
            rs = await _eval(ws_dl, "document.readyState", cmd_id=1)
            if rs == "complete": break
            await asyncio.sleep(0.5)

        await ws_dl.send(json.dumps({"id": 2, "method": "Network.enable"}))
        while True:
            r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=5))
            if r.get("id") == 2: break

        await ws_dl.send(json.dumps({"id": 3, "method": "Browser.setDownloadBehavior",
            "params": {"behavior": "allow", "downloadPath": sdir, "eventsEnabled": True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=5))
            if r.get("id") == 3: break

        # report_key
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
    if (!btn) return null;
    var m = btn.id.match(/re_pdf(.+)/);
    return m ? m[1] : null;
})()""", cmd_id=60 + cid % 50, ctx_id=cid)
                if rkey2:
                    rkey = rkey2; break

        if not rkey:
            print("  report_key 없음"); return False
        print(f"  report_key: {rkey[:20]}...")

        tmp_pdf = os.path.join(sdir, "종합소득세.pdf")
        before_mtime = os.path.getmtime(tmp_pdf) if os.path.exists(tmp_pdf) else 0

        res = await _eval(ws_dl, f"""(function(){{
    try {{ m_reportHashMap['{rkey}'].pdfDownLoad(); return 'ok'; }}
    catch(e) {{ return 'err:' + e.message; }}
}})()""", cmd_id=70)
        print(f"  pdfDownLoad: {res}")

        print("  다운로드 대기...")
        dl_done = False
        suggested = "종합소득세.pdf"
        t_dl = asyncio.get_event_loop().time() + 60
        while asyncio.get_event_loop().time() < t_dl:
            try:
                r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=0.5))
                if r.get("method") == "Browser.downloadWillBegin":
                    suggested = r["params"].get("suggestedFilename", suggested)
                    print(f"  파일명: {suggested}")
                elif r.get("method") == "Browser.downloadProgress":
                    if r["params"].get("state") == "completed":
                        print("  완료!"); dl_done = True; break
            except asyncio.TimeoutError:
                pass

        await asyncio.sleep(1)
        new_src = None
        for fn in [suggested, "종합소득세.pdf"]:
            fp = os.path.join(sdir, fn)
            if os.path.exists(fp) and os.path.getmtime(fp) > before_mtime + 0.5:
                new_src = fp; break
        if not new_src:
            dld = os.path.join(os.environ.get("USERPROFILE", ""), "Downloads")
            for fn in [suggested, "종합소득세.pdf"]:
                fp = os.path.join(dld, fn)
                if os.path.exists(fp) and os.path.getmtime(fp) > time.time() - 30:
                    new_src = fp; break

        if new_src:
            archive_if_exists(dst)   # 기존 파일 _archive로
            if new_src != dst:
                shutil.move(new_src, dst)
            sz = os.path.getsize(dst) // 1024
            print(f"  저장 OK: {os.path.basename(dst)} ({sz}KB)")
        else:
            print("  파일 없음"); return False

    try:
        requests.get(f"{CDP}/json/close/{popup['id']}")
    except:
        pass
    await asyncio.sleep(1)
    return True


# ── 메인 ─────────────────────────────────────────────────────────────────────

async def run():
    nas_ok = os.path.isdir(NAS_BASE)
    print(f"저장 경로: {'NAS ' + NAS_BASE if nas_ok else '로컬 ' + LOCAL_BASE}")

    tabs = requests.get(f"{CDP}/json").json()
    main = next((t for t in tabs
                 if "hometax.go.kr" in t.get("url", "")
                 and "websquare.html" in t.get("url", "")
                 and "sesw." not in t.get("url", "")), None)
    if not main:
        print("홈택스 탭 없음"); return

    known_ids  = set(t["id"] for t in tabs)
    page_num   = 1
    seen_first = set()
    total_ok = total_fail = 0

    # 스크립트 시작 시 이미 조회된 상태일 경우 초기 alert 처리
    print("초기 alert 확인...")
    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        await click_confirm_once(ws, main["id"], timeout=3.0)
    await asyncio.sleep(0.5)

    while True:
        async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
            rows = await get_rows(ws)
        if not rows:
            print("행 없음"); break

        first = rows[0]["name"] if rows else ""
        if first and first in seen_first:
            print(f"루프 감지 ({first}) - 완료"); break
        if first:
            seen_first.add(first)

        print(f"\n{'='*50} 페이지 {page_num} - {len(rows)}건")

        page_new = 0  # 이번 페이지에서 새로 다운로드한 건수
        for idx, row in enumerate(rows):
            name = row.get("name", "")
            if not name:
                continue
            page_new += 1
            try:
                ok = await process_jeupsujeung(main, idx, name, known_ids)
                if ok:   total_ok += 1
                else:    total_fail += 1
            except Exception as e:
                print(f"  [{name}] 오류: {e}")
                traceback.print_exc()
                total_fail += 1
            await asyncio.sleep(2)

        if page_new == 0:
            print("  이번 페이지 신규 없음 - 완료")
            break

        ok_next = await click_next_page(main)
        if not ok_next:
            print("마지막 페이지"); break

        # 페이지 전환 후 조회완료 alert 한 번만 처리 (pyautogui — JS click은 재조회 유발)
        print("  페이지 전환 alert 대기...")
        await asyncio.sleep(1.5)
        async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
            ok_alert = await click_confirm_once(ws, main["id"], timeout=8.0)
            if ok_alert:
                print("  조회완료 alert 처리")
        await asyncio.sleep(1)
        page_num += 1

    print(f"\n완료! 성공 {total_ok} / 실패 {total_fail}")

    base = NAS_BASE if nas_ok else LOCAL_BASE
    pdfs = []
    for root, _, files in os.walk(base):
        for f in files:
            if "_접수증" in f and f.endswith(".pdf"):
                pdfs.append(f)
    print(f"접수증 PDF {len(pdfs)}개")


asyncio.run(run())
