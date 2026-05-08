# -*- coding: utf-8 -*-
"""
홈택스 접수증 + 납부서 일괄 PDF 다운로드
  col[12] 보기 -> {name}_접수증.pdf
  col[13] 보기 -> {name}_납부서.pdf  (납부서 없는 고객은 즉시 스킵)
  NAS: Z:\종소세2026\고객\{name}\  (NAS 없으면 로컬 C:\Users\pc\종소세2026\{name}\)
"""
import requests, json, asyncio, websockets, pyautogui, time, os, shutil, traceback

CDP        = "http://localhost:9222"
NAME_COL   = 6
JEUP_COL   = 12   # 접수증 보기
NABU_COL   = 13   # 납부서 보기
NAS_BASE   = r"Z:\종소세2026\고객"
LOCAL_BASE = r"C:\Users\pc\종소세2026"
ROW_SEL    = '[id*="UTERNAAZ0Z31_wframe"] table tbody tr'

pyautogui.PAUSE = 0.15


def get_save_dir(name):
    """NAS 마운트 여부 확인, 없으면 로컬 폴더"""
    base = NAS_BASE if os.path.isdir(NAS_BASE) else LOCAL_BASE
    d = os.path.join(base, name)
    os.makedirs(d, exist_ok=True)
    return d


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


async def dismiss_alert(ws, timeout=4.0):
    """WebSquare '확인' 버튼 자동 클릭. 반환: True=처리됨"""
    deadline = asyncio.get_event_loop().time() + timeout
    while asyncio.get_event_loop().time() < deadline:
        r = await _eval(ws, """(function(){
    var b = Array.from(document.querySelectorAll('button,input[type=button]'))
        .find(function(b){
            var t = (b.innerText||b.value||'').trim();
            var r = b.getBoundingClientRect();
            return t === '확인' && r.width > 0;
        });
    if (b){ b.click(); return 'ok'; } return null;
})()""", cmd_id=41)
        if r == 'ok':
            return True
        await asyncio.sleep(0.3)
    return False


async def clear_alerts(main_tab, label=""):
    """메인 탭 WebSocket 열어서 alert 모두 처리"""
    async with websockets.connect(main_tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        r = await dismiss_alert(ws, timeout=2.0)
        if r:
            print(f"  [alert 처리{' - ' + label if label else ''}]")


async def click_next_page(main_tab):
    async with websockets.connect(main_tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        bounds, vp_w, vp_h = await get_win_info(ws, main_tab["id"])
        coords = await _eval(ws, """(function(){
    var all = Array.from(document.querySelectorAll('a,input[type=button],button,td'));
    for (var el of all){
        var t = (el.innerText||el.value||'').trim();
        var r = el.getBoundingClientRect();
        if (r.width>0 && r.height>0 && (t==='다음'||t==='>') && r.top > window.innerHeight*0.5){
            el.scrollIntoView({block:'center',behavior:'instant'});
            r = el.getBoundingClientRect();
            return JSON.stringify({x:r.left+r.width/2, y:r.top+r.height/2});
        }
    }
    return null;
})()""", cmd_id=40)
        if not coords:
            return False
        c = json.loads(coords)
        sx, sy = get_screen_xy(bounds, vp_w, vp_h, c["x"], c["y"])
        print(f"  다음 페이지 클릭 ({sx},{sy})")
        pyautogui.click(sx, sy)
        # 1차 alert (즉시)
        await asyncio.sleep(1.5)
        r1 = await dismiss_alert(ws, timeout=5)
        # 2차 alert (데이터 로딩 후)
        await asyncio.sleep(2.0)
        r2 = await dismiss_alert(ws, timeout=4)
        print(f"  alert 1차={'ok' if r1 else '-'} 2차={'ok' if r2 else '-'}")
    # 3차: 혹시 남은 alert
    await asyncio.sleep(1.5)
    await clear_alerts(main_tab, "3차")
    await asyncio.sleep(1)
    return True


# ── 팝업 PDF 공통 다운로드 ──────────────────────────────────────────────────────

async def download_popup_pdf(main_tab, row_idx, col_idx, dst_path, known_ids, label):
    """
    col_idx 버튼 클릭 -> sesw 팝업 -> report_key -> pdfDownLoad() -> dst_path 저장
    반환: True=성공 or 스킵, False=실패
    """
    save_dir = os.path.dirname(dst_path)

    # pyautogui 클릭 전 혹시 남은 alert 처리
    await clear_alerts(main_tab)

    # STEP 1: 버튼 좌표 확인 + 클릭
    async with websockets.connect(main_tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        bounds, vp_w, vp_h = await get_win_info(ws, main_tab["id"])
        coords = await _eval(ws, f"""(function(){{
    var rows = Array.from(document.querySelectorAll('{ROW_SEL}'))
        .filter(function(tr){{ return tr.querySelectorAll('td').length >= 13; }});
    var cell = rows[{row_idx}] && rows[{row_idx}].querySelectorAll('td')[{col_idx}];
    if (!cell) return null;
    var btn = cell.querySelector('a,input[type=button],button');
    if (!btn) return null;
    btn.scrollIntoView({{block:'center',behavior:'instant'}});
    var r = btn.getBoundingClientRect();
    if (!r.width || !r.height) return null;
    return JSON.stringify({{x:r.left+r.width/2, y:r.top+r.height/2}});
}})()""", cmd_id=20)

        if not coords:
            print(f"  [{label}] 버튼 없음 - 스킵")
            return True   # 납부서 없는 고객 → True(정상 스킵)

        c = json.loads(coords)
        sx, sy = get_screen_xy(bounds, vp_w, vp_h, c["x"], c["y"])
        print(f"  [{label}] 클릭 ({sx},{sy})")
        pyautogui.click(sx, sy)

    # STEP 2: sesw 팝업 대기 (최대 15초)
    print(f"  [{label}] 팝업 대기...")
    popup = None
    for _ in range(30):
        time.sleep(0.5)
        for t in requests.get(f"{CDP}/json").json():
            if t["id"] not in known_ids and "sesw.hometax.go.kr" in t.get("url", ""):
                popup = t
                break
        if popup:
            break

    if not popup:
        # fallback: 아무 새 hometax 탭
        for t in requests.get(f"{CDP}/json").json():
            if t["id"] not in known_ids and "hometax.go.kr" in t.get("url", ""):
                popup = t
                break

    if not popup:
        print(f"  [{label}] 팝업 미감지")
        return False

    known_ids.add(popup["id"])
    time.sleep(2)

    # STEP 3: pdfDownLoad() 호출
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
            "params": {"behavior": "allow", "downloadPath": save_dir, "eventsEnabled": True}}))
        while True:
            r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=5))
            if r.get("id") == 3:
                break

        # report_key: 메인 컨텍스트
        rkey = await _eval(ws_dl, """(function(){
    var btn = Array.from(document.querySelectorAll('button'))
        .find(function(b){ return b.id.indexOf('re_pdf') === 0; });
    if (!btn) return null;
    var m = btn.id.match(/re_pdf(.+)/);
    return m ? m[1] : null;
})()""", cmd_id=60)

        # 없으면 iframe 컨텍스트 탐색
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
                    rkey = rkey2
                    break

        if not rkey:
            print(f"  [{label}] report_key 없음")
            return False
        print(f"  [{label}] report_key: {rkey[:20]}...")

        tmp_pdf = os.path.join(save_dir, "종합소득세.pdf")
        before_mtime = os.path.getmtime(tmp_pdf) if os.path.exists(tmp_pdf) else 0

        res = await _eval(ws_dl, f"""(function(){{
    try {{ m_reportHashMap['{rkey}'].pdfDownLoad(); return 'ok'; }}
    catch(e) {{ return 'err:' + e.message; }}
}})()""", cmd_id=70)
        print(f"  [{label}] pdfDownLoad: {res}")

        print(f"  [{label}] 다운로드 대기...")
        dl_done = False
        suggested = "종합소득세.pdf"
        t_dl = asyncio.get_event_loop().time() + 60
        while asyncio.get_event_loop().time() < t_dl:
            try:
                r = json.loads(await asyncio.wait_for(ws_dl.recv(), timeout=0.5))
                if r.get("method") == "Browser.downloadWillBegin":
                    suggested = r["params"].get("suggestedFilename", suggested)
                    print(f"  [{label}] 파일명: {suggested}")
                elif r.get("method") == "Browser.downloadProgress":
                    if r["params"].get("state") == "completed":
                        print(f"  [{label}] 완료!")
                        dl_done = True
                        break
            except asyncio.TimeoutError:
                pass

        await asyncio.sleep(1)
        new_src = None
        for fn in [suggested, "종합소득세.pdf"]:
            fp = os.path.join(save_dir, fn)
            if os.path.exists(fp) and os.path.getmtime(fp) > before_mtime + 0.5:
                new_src = fp
                break
        if not new_src:
            dld = os.path.join(os.environ.get("USERPROFILE", ""), "Downloads")
            for fn in [suggested, "종합소득세.pdf"]:
                fp = os.path.join(dld, fn)
                if os.path.exists(fp) and os.path.getmtime(fp) > time.time() - 30:
                    new_src = fp
                    break

        if new_src:
            if new_src != dst_path:
                shutil.move(new_src, dst_path)
            sz = os.path.getsize(dst_path) // 1024
            print(f"  [{label}] 저장 완료: {os.path.basename(dst_path)} ({sz}KB)")
        else:
            print(f"  [{label}] 파일 없음")
            return False

    try:
        requests.get(f"{CDP}/json/close/{popup['id']}")
    except:
        pass
    await asyncio.sleep(1)
    return True


# ── 1건 처리 ─────────────────────────────────────────────────────────────────

async def process_row(main_tab, row_idx, name, known_ids):
    sdir = get_save_dir(name)
    jeup_dst = os.path.join(sdir, f"{name}_접수증.pdf")
    nabu_dst = os.path.join(sdir, f"{name}_납부서.pdf")

    # 접수증
    if os.path.exists(jeup_dst):
        print(f"  [{name}] 접수증 이미 있음 - 스킵")
        jeup_ok = True
    else:
        jeup_ok = await download_popup_pdf(
            main_tab, row_idx, JEUP_COL, jeup_dst, known_ids, "접수증")

    await asyncio.sleep(1)

    # 납부서 (없는 고객은 버튼 자체가 없어서 즉시 스킵)
    if os.path.exists(nabu_dst):
        print(f"  [{name}] 납부서 이미 있음 - 스킵")
        nabu_ok = True
    else:
        nabu_ok = await download_popup_pdf(
            main_tab, row_idx, NABU_COL, nabu_dst, known_ids, "납부서")

    return jeup_ok and nabu_ok


# ── 메인 ─────────────────────────────────────────────────────────────────────

async def run():
    nas_ok = os.path.isdir(NAS_BASE)
    if nas_ok:
        print(f"[NAS] {NAS_BASE}")
    else:
        print(f"[로컬] {LOCAL_BASE} (NAS 미연결)")

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

    while True:
        # 혹시 남은 alert 처리 후 행 읽기
        await clear_alerts(main, "페이지 시작")
        await asyncio.sleep(0.5)

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

        for idx, row in enumerate(rows):
            name = row.get("name", "")
            if not name:
                continue
            print(f"\n[{name}] row={idx}")
            try:
                ok = await process_row(main, idx, name, known_ids)
                if ok:
                    total_ok += 1
                else:
                    total_fail += 1
            except Exception as e:
                print(f"  [{name}] 오류: {e}")
                traceback.print_exc()
                total_fail += 1
            await asyncio.sleep(2)

        # 다음 페이지
        ok_next = await click_next_page(main)
        if not ok_next:
            print("마지막 페이지"); break
        page_num += 1

    print(f"\n완료! 성공 {total_ok} / 실패 {total_fail}")

    # 결과 요약
    for base in ([NAS_BASE] if nas_ok else []) + [LOCAL_BASE]:
        if os.path.isdir(base):
            pdfs = []
            for root, _, files in os.walk(base):
                for f in files:
                    if f.endswith(".pdf"):
                        pdfs.append(f)
            print(f"저장된 PDF: {len(pdfs)}개 in {base}")
            break


asyncio.run(run())
