# -*- coding: utf-8 -*-
# ╔══════════════════════════════════════════════════════════════════╗
# ║  실행순서 [3/3]  위택스 지방소득세 접수증                            ║
# ║  시스템: 위택스 (wetax.go.kr) — CDP localhost:9222               ║
# ║  순서:   신고서(1) → 접수증(2) → 지방세(3)                          ║
# ║  ⚠ 반드시 단독 실행! 다른 download_*.py 와 동시 실행 절대 금지 ⚠      ║
# ╚══════════════════════════════════════════════════════════════════╝
# 위택스 지방소득세 접수증 일괄 PDF 다운로드
# pure CDP (pyautogui 없음) — Page.printToPDF 방식 → 마우스 자유 사용 가능
#
# 전제: 위택스 지방소득세 신고내역 화면이 열린 상태에서 실행
#   URL: https://www.wetax.go.kr/etr/lit/b0702/B070202M01.do
#
# 테이블 구조 (DOM 확인 결과):
#   - tblMain 내부 각 행: <tr> 10td (td[1]=납세자명, td[8]=출력버튼 ui-id-N)
#   - 출력버튼 클릭 → tooltipster 드롭다운 → 접수증(title="접수증") 클릭 → OZReport 새창
#   - 페이지네이션: a[href="#n"] .active = 현재페이지, 다음 숫자 클릭
#
# 저장: Z:\종소세2026\고객\{name}_{jumin6}\발송용\{name}_지방세접수증.pdf
# 완료 후: verify_folder_integrity.py --fix 자동 실행
import sys, os, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)

import requests, json, asyncio, websockets, shutil, traceback, unicodedata, base64, subprocess
from datetime import datetime

CDP        = "http://localhost:9222"
NAS_BASE   = r"Z:\종소세2026\고객"
LOCAL_BASE = r"C:\Users\pc\종소세2026"


# ── 유틸 ─────────────────────────────────────────────────────────────────────

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
    """기존 파일 → _archive 폴더로 타임스탬프 붙여 이동 (root=최신 원칙)"""
    if not os.path.exists(dst):
        return
    archive_dir = os.path.join(os.path.dirname(dst), "_archive")
    os.makedirs(archive_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(dst))
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    shutil.move(dst, os.path.join(archive_dir, f"{stem}_{ts}{ext}"))
    print(f"  아카이브: {stem}_{ts}{ext}")


# ── CDP 헬퍼 ─────────────────────────────────────────────────────────────────

async def _send(ws, msg, timeout=20):
    await ws.send(json.dumps(msg))
    while True:
        r = json.loads(await asyncio.wait_for(ws.recv(), timeout=timeout))
        if r.get("id") == msg["id"]:
            return r


async def _eval(ws, code, cmd_id=1):
    r = await _send(ws, {"id": cmd_id, "method": "Runtime.evaluate",
                         "params": {"expression": code, "returnByValue": True}})
    return r.get("result", {}).get("result", {}).get("value")


# ── 행 목록 파싱 ──────────────────────────────────────────────────────────────
# DOM 구조: tblMain 내 각 <tr>에 10개 td
#   td[0]=납부여부(-)  td[1]=납세자명  td[2]=신고일자  td[3]=관할자치단체
#   td[4]=금액  td[5]=진행상태  td[6]=납부기한  td[7]=empty
#   td[8]=출력물보기(a[id="ui-id-N"])  td[9]=신고구분

async def get_rows(ws):
    """현재 페이지 행 목록 반환 [{name, btn_id}]"""
    raw = await _eval(ws, """JSON.stringify((function(){
    var buttons = Array.from(document.querySelectorAll('a[id^="ui-id-"]'));
    return buttons.map(function(btn){
        var tr  = btn.closest("tr");
        var tds = tr ? tr.querySelectorAll("td") : [];
        var name = (tds[1] || {}).innerText || "";
        return { name: name.trim(), btn_id: btn.id };
    }).filter(function(r){ return r.name.length > 0; });
})())""", cmd_id=5)
    return json.loads(raw) if raw else []


# ── 1건 처리 ─────────────────────────────────────────────────────────────────

async def close_tooltip(ws):
    """열린 tooltipster 직접 숨기기 — tooltip-close 클릭 금지 (href="" 페이지 재로드 방지)"""
    await _eval(ws, """(function(){
    Array.from(document.querySelectorAll('.tooltipster-base, .tooltipster-sidetip'))
        .forEach(function(el){ el.style.display = 'none'; });
})()""", cmd_id=9)


async def process_row(main_tab, row_idx, name, known_ids):
    sdir = find_balsong_dir(name)
    dst  = os.path.join(sdir, f"{name}_지방세접수증.pdf")

    print(f"\n{'='*60}")
    print(f"  처리: {name} (행 {row_idx})")

    # 중복 스킵
    if os.path.exists(dst):
        print(f"  [{name}] 파일 이미 존재 — 스킵")
        return True

    async with websockets.connect(main_tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        # 잔여 툴팁 숨기기
        await close_tooltip(ws)
        await asyncio.sleep(0.2)

        # STEP 1: 출력버튼 클릭 + tooltip-content 선택자 반환
        tip_sel = await _eval(ws, f"""(function(){{
    var buttons = Array.from(document.querySelectorAll('a[id^="ui-id-"]'));
    var btn = buttons[{row_idx}];
    if (!btn) return 'no_btn';
    btn.scrollIntoView({{block: 'center', behavior: 'instant'}});
    btn.click();
    return btn.getAttribute('data-tooltip-content') || 'no_tipsel';
}})()""", cmd_id=20)
        print(f"  출력버튼: {tip_sel}")
        if not tip_sel or tip_sel in ('no_btn', 'no_tipsel'):
            print("  버튼 없음"); return False

        # tooltip 컨텐츠 로드 대기
        await asyncio.sleep(1.0)

        # STEP 2: 해당 tooltip div 내에서 접수증 클릭 (visible 여부 무관)
        result2 = await _eval(ws, f"""(function(){{
    var tipDiv = document.querySelector('{tip_sel}');
    if (!tipDiv) return 'no_tipdiv';
    var link = Array.from(tipDiv.querySelectorAll('a'))
        .find(function(a){{
            return a.getAttribute('title') === '접수증' || (a.innerText||'').trim() === '접수증';
        }});
    if (!link) return JSON.stringify({{err:'no_link', html:tipDiv.innerHTML.slice(0,200)}});
    link.click();
    return 'clicked:' + (link.getAttribute('onclick')||'').slice(0, 60);
}})()""", cmd_id=21)
        print(f"  접수증 클릭: {result2}")
        if not result2 or 'no_' in str(result2) or 'err' in str(result2):
            print("  접수증 메뉴 없음"); return False

    # STEP 3: OZReport 새창 대기 (최대 20초)
    print("  OZReport 창 대기...")
    oz_tab = None
    for _ in range(40):
        await asyncio.sleep(0.5)
        try:
            tabs = requests.get(f"{CDP}/json", timeout=5).json()
        except Exception:
            continue
        for t in tabs:
            if t["id"] not in known_ids and t.get("type") == "page":
                url = t.get("url", "")
                if "wetax.go.kr" in url and ("rpt" in url or "ozhJson" in url or "viewer" in url):
                    oz_tab = t
                    break
        if oz_tab:
            break

    if not oz_tab:
        print("  OZReport 창 미감지"); return False
    known_ids.add(oz_tab["id"])
    print(f"  OZReport: {oz_tab['url'][:70]}")

    # STEP 4-A: readyState 완료 대기 (연결1 — 렌더링 전 닫힐 수 있어 분리)
    try:
        async with websockets.connect(oz_tab["webSocketDebuggerUrl"], ping_interval=None) as ws_oz:
            for _ in range(40):
                rs = await _eval(ws_oz, "document.readyState", cmd_id=1)
                if rs == "complete":
                    break
                await asyncio.sleep(0.5)
    except Exception:
        pass  # readyState 실패해도 계속 진행

    # OZ 렌더링 대기 (연결 밖에서 — 연결 유지 불필요)
    print("  OZ 렌더링 대기 (5초)...")
    await asyncio.sleep(5)

    # STEP 4-B: Page.printToPDF — 새 연결로 (기존 연결이 5초 대기 중 닫힐 수 있음)
    pdf_bytes = None
    print("  Page.printToPDF 실행...")
    try:
        fresh_tabs = requests.get(f"{CDP}/json").json()
        oz_fresh = next((t for t in fresh_tabs if t["id"] == oz_tab["id"]), oz_tab)
        async with websockets.connect(oz_fresh["webSocketDebuggerUrl"], ping_interval=None) as ws_pdf:
            r = await _send(ws_pdf, {"id": 50, "method": "Page.printToPDF", "params": {
                "printBackground": True,
                "paperWidth":  8.27,    # A4
                "paperHeight": 11.69,
                "marginTop":    0.4,
                "marginBottom": 0.4,
                "marginLeft":   0.4,
                "marginRight":  0.4,
                "scale": 0.9,
                "landscape": False,
            }}, timeout=30)
            data_b64 = r.get("result", {}).get("data", "")
            if data_b64:
                pdf_bytes = base64.b64decode(data_b64)
                print(f"  PDF {len(pdf_bytes)//1024}KB 생성")
            else:
                print(f"  printToPDF 응답 없음: {r.get('result', {})}")
    except Exception as e:
        print(f"  PDF 생성 오류: {e}")
        traceback.print_exc()

    # OZReport 창 닫기
    try:
        requests.get(f"{CDP}/json/close/{oz_tab['id']}", timeout=5)
    except Exception:
        pass
    await asyncio.sleep(0.5)

    if not pdf_bytes:
        print("  PDF bytes 없음"); return False

    # STEP 5: 저장
    archive_if_exists(dst)
    with open(dst, "wb") as f:
        f.write(pdf_bytes)
    sz = len(pdf_bytes) // 1024
    print(f"  저장 OK: {dst} ({sz}KB)")
    return True


# ── 다음 페이지 ───────────────────────────────────────────────────────────────

async def click_next_page(main_tab):
    """현재 active 페이지 다음 번호 클릭. 없으면 None 반환."""
    async with websockets.connect(main_tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        result = await _eval(ws, """(function(){
    // a[href="#n"] 목록 — 페이지 번호 링크
    var pages = Array.from(document.querySelectorAll('a[href="#n"]'))
        .filter(function(a){
            var r = a.getBoundingClientRect();
            return r.width > 0 && r.height > 0;
        });
    var activeIdx = -1;
    for (var i = 0; i < pages.length; i++){
        if (pages[i].classList.contains('active')){ activeIdx = i; break; }
    }
    if (activeIdx < 0 || activeIdx >= pages.length - 1) return null;
    pages[activeIdx + 1].click();
    return 'page_' + (activeIdx + 2);
})()""", cmd_id=40)
    return result


# ── 메인 ─────────────────────────────────────────────────────────────────────

async def run():
    nas_ok = os.path.isdir(NAS_BASE)
    print(f"저장 경로: {'NAS ' + NAS_BASE if nas_ok else '로컬 ' + LOCAL_BASE}")

    tabs = requests.get(f"{CDP}/json").json()
    main = next((t for t in tabs
                 if "wetax.go.kr" in t.get("url", "")
                 and t.get("type") == "page"), None)
    if not main:
        print("위택스 탭 없음!")
        print("  → https://www.wetax.go.kr 에서 지방소득세 신고내역 조회 후 재실행")
        return

    print(f"탭 발견: {main['url'][:90]}")

    known_ids  = set(t["id"] for t in tabs)
    page_num   = 1
    seen_first = set()
    total_ok = total_fail = 0

    while True:
        async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
            rows = await get_rows(ws)

        if not rows:
            print("행 없음 — 종료"); break

        first = rows[0]["name"]
        if first and first in seen_first:
            print(f"루프 감지 ({first}) — 종료"); break
        if first:
            seen_first.add(first)

        print(f"\n{'='*60}")
        print(f"  페이지 {page_num} — {len(rows)}건")
        print(f"{'='*60}")

        for idx, row in enumerate(rows):
            name = row.get("name", "")
            if not name:
                continue
            try:
                ok = await process_row(main, idx, name, known_ids)
                if ok:   total_ok += 1
                else:    total_fail += 1
            except Exception as e:
                print(f"  [{name}] 오류: {e}")
                traceback.print_exc()
                total_fail += 1
            await asyncio.sleep(1.5)

        # 다음 페이지
        nxt = await click_next_page(main)
        if not nxt:
            print("다음 페이지 없음 — 완료"); break
        print(f"  다음 페이지: {nxt}")
        await asyncio.sleep(3)
        page_num += 1

    print(f"\n{'='*60}")
    print(f"완료!  성공 {total_ok} / 실패 {total_fail}")
    print(f"{'='*60}")

    # ── 혼입검증 자동 실행 ───────────────────────────────────────────────────
    integrity_script = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "verify_folder_integrity.py"
    )
    if os.path.exists(integrity_script):
        print(f"\n혼입검증 자동 실행: {integrity_script}")
        sys.stdout.flush()  # 다운로드 출력 먼저 터미널에 표시
        subprocess.run([sys.executable, integrity_script, "--fix"], check=False)
    else:
        print(f"\n혼입검증 스크립트 없음 (수동 실행): {integrity_script}")


asyncio.run(run())
