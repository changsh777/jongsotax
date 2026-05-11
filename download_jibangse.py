# -*- coding: utf-8 -*-
# 위택스 지방소득세 접수증 일괄 PDF 다운로드
# pure CDP (pyautogui 없음) — Page.printToPDF 방식
# 마우스 자유 사용 가능
#
# 전제: 위택스 신고내역 화면이 열린 상태에서 실행
#   URL: https://www.wetax.go.kr/etr/lit/b0702/B070202M01.do
#
# 저장: Z:\종소세2026\고객\{name}_{jumin6}\발송용\{name}_지방세접수증.pdf
# 완료 후: verify_folder_integrity.py --fix 자동 실행
import sys, os, io
# stdout UTF-8 강제 (Windows cp949 환경 대비)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import requests, json, asyncio, websockets, time, shutil, traceback, unicodedata, base64, subprocess
from datetime import datetime

CDP        = "http://localhost:9222"
NAS_BASE   = r"Z:\종소세2026\고객"
LOCAL_BASE = r"C:\Users\pc\종소세2026"

# 위택스 신고내역 테이블 컬럼 (0-based)
# -, 납세자명, 신고일자, 관할자치단체, 금액, 진행상태, 납부기한, 납부여부, 출력(📋), 신고구분
NAME_COL  = 1
PRINT_COL = 8
MIN_COLS  = 9   # 데이터 행 최소 td 수


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


# ── 테이블 행 파싱 ────────────────────────────────────────────────────────────

async def get_rows(ws):
    """현재 페이지 데이터 행 목록 반환"""
    raw = await _eval(ws, f"""JSON.stringify(
    (function(){{
        var allRows = Array.from(document.querySelectorAll('table tbody tr'));
        var dataRows = allRows.filter(function(tr){{
            return tr.querySelectorAll('td').length >= {MIN_COLS};
        }});
        return dataRows.map(function(tr){{
            var tds = tr.querySelectorAll('td');
            var name = (tds[{NAME_COL}] || {{}}).innerText || '';
            return {{ name: name.trim() }};
        }}).filter(function(r){{ return r.name.length > 0; }});
    }})()
)""", cmd_id=5)
    return json.loads(raw) if raw else []


async def diagnose_table(ws):
    """시작 시 테이블 구조 진단 — 컬럼 인덱스 확인용"""
    raw = await _eval(ws, f"""JSON.stringify(
    (function(){{
        var allRows = Array.from(document.querySelectorAll('table tbody tr'));
        var row = allRows.find(function(tr){{
            return tr.querySelectorAll('td').length >= {MIN_COLS};
        }});
        if (!row) return null;
        var tds = row.querySelectorAll('td');
        var cols = [];
        for (var i = 0; i < tds.length; i++) {{
            cols.push(i + ':' + (tds[i].innerText || '').trim().slice(0, 10));
        }}
        return cols;
    }})()
)""", cmd_id=6)
    if raw:
        cols = json.loads(raw)
        print(f"  테이블 컬럼: {' | '.join(cols)}")


# ── 1건 처리 ─────────────────────────────────────────────────────────────────

async def process_row(main_tab, row_idx, name, known_ids):
    sdir = find_balsong_dir(name)
    dst  = os.path.join(sdir, f"{name}_지방세접수증.pdf")

    print(f"\n{'='*60}")
    print(f"  처리: {name} (행 {row_idx})")

    # 중복 스킵
    if os.path.exists(dst):
        print(f"  [{name}] 파일 이미 존재 — 스킵")
        return True

    # ── STEP 1: 출력(📋) 버튼 클릭 → 드롭다운 오픈 ──────────────────────────
    async with websockets.connect(main_tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        result = await _eval(ws, f"""(function(){{
    var allRows = Array.from(document.querySelectorAll('table tbody tr'))
        .filter(function(tr){{ return tr.querySelectorAll('td').length >= {MIN_COLS}; }});
    var row = allRows[{row_idx}];
    if (!row) return 'no_row';
    row.scrollIntoView({{block:'center', behavior:'instant'}});
    var tds = row.querySelectorAll('td');
    var td = tds[{PRINT_COL}];
    if (!td) return 'no_td';
    // 버튼/링크/이미지/span(onclick) 순서로 탐색
    var btn = td.querySelector('button, a, input[type=button], img[onclick], span[onclick]');
    if (!btn) {{
        // onclick 없는 img도 시도
        btn = td.querySelector('img');
    }}
    if (!btn) {{
        // td 자체 클릭
        td.click();
        return 'td_click';
    }}
    btn.click();
    return 'clicked:' + btn.tagName;
}})()""", cmd_id=20)
        print(f"  출력버튼: {result}")
        if result in ('no_row', 'no_td', None):
            print("  출력 버튼 위치 확인 필요"); return False

        # 드롭다운 애니메이션 대기
        await asyncio.sleep(0.8)

        # ── STEP 2: 드롭다운에서 '접수증' 클릭 ──────────────────────────────
        result2 = await _eval(ws, """(function(){
    // 리프 노드부터 검색 (innerText가 정확히 '접수증')
    var all = Array.from(document.querySelectorAll('*'));
    for (var i = 0; i < all.length; i++) {
        var el = all[i];
        var t  = (el.innerText || el.textContent || '').trim();
        if (t !== '접수증') continue;
        var r = el.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) {
            el.click();
            return 'clicked:' + el.tagName + '#' + (el.id||'') + '.' + (el.className||'');
        }
    }
    return 'not_found';
})()""", cmd_id=21)
        print(f"  접수증 클릭: {result2}")
        if result2 == 'not_found':
            print("  접수증 메뉴 없음 (드롭다운 미표시?)"); return False

    # ── STEP 3: OZReport 팝업 대기 (최대 20초) ───────────────────────────────
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
                # ozhJsonviewer.oz OR wetax 내 rpt/viewer 경로
                if "ozhJsonviewer" in url or (
                        "wetax.go.kr" in url and ("rpt" in url or "viewer" in url)):
                    oz_tab = t
                    break
        if oz_tab:
            break

    if not oz_tab:
        print("  OZReport 창 미감지 — 수동으로 확인 후 재시도")
        return False
    known_ids.add(oz_tab["id"])
    print(f"  OZReport: {oz_tab['url'][:70]}")

    # ── STEP 4: readyState 대기 + OZ 렌더링 대기 + Page.printToPDF ──────────
    pdf_bytes = None
    try:
        async with websockets.connect(oz_tab["webSocketDebuggerUrl"], ping_interval=None) as ws_oz:
            # readyState complete 대기 (최대 20초)
            for _ in range(40):
                rs = await _eval(ws_oz, "document.readyState", cmd_id=1)
                if rs == "complete":
                    break
                await asyncio.sleep(0.5)

            # OZ 리포트 렌더링 추가 대기
            print("  OZ 렌더링 대기 (5초)...")
            await asyncio.sleep(5)

            # Page.printToPDF
            print("  Page.printToPDF 실행...")
            r = await _send(ws_oz, {"id": 50, "method": "Page.printToPDF", "params": {
                "printBackground": True,
                "paperWidth":  8.27,    # A4 세로
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
                print(f"  Page.printToPDF 응답 없음: {r.get('result',{})}")
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
        print("  PDF bytes 없음 — 스킵"); return False

    # ── STEP 5: 저장 ─────────────────────────────────────────────────────────
    archive_if_exists(dst)
    with open(dst, "wb") as f:
        f.write(pdf_bytes)
    sz = len(pdf_bytes) // 1024
    print(f"  저장 OK: {dst} ({sz}KB)")
    return True


# ── 다음 페이지 ───────────────────────────────────────────────────────────────

async def click_next_page(main_tab):
    """위택스 다음 페이지 버튼 JS 클릭"""
    async with websockets.connect(main_tab["webSocketDebuggerUrl"], ping_interval=None) as ws:
        result = await _eval(ws, """(function(){
    var all = Array.from(document.querySelectorAll('a, button, input, span, img'));
    for (var i = 0; i < all.length; i++) {
        var el = all[i];
        var t = (el.innerText || el.value || el.textContent || '').trim();
        var lbl = (el.getAttribute('aria-label') || el.getAttribute('title') || '').trim();
        var r = el.getBoundingClientRect();
        // 하단 절반에 있는 '다음' 계열 버튼
        if (r.width > 0 && r.height > 0 && r.top > window.innerHeight * 0.4) {
            if (t === '다음' || t === '>' || t === '다음페이지' ||
                    lbl.indexOf('다음') >= 0) {
                el.click();
                return 'clicked:' + (t || lbl);
            }
        }
    }
    return null;
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
        print("  → 위택스(https://www.wetax.go.kr) 열고 지방소득세 신고내역 조회 후 재실행")
        return

    print(f"탭 발견: {main['url'][:90]}")

    # 테이블 구조 진단 (컬럼 인덱스 확인)
    print("테이블 구조 진단...")
    async with websockets.connect(main["webSocketDebuggerUrl"], ping_interval=None) as ws:
        await diagnose_table(ws)

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
        subprocess.run([sys.executable, integrity_script, "--fix"], check=False)
    else:
        print(f"\n혼입검증 스크립트 없음 (수동 실행): {integrity_script}")


asyncio.run(run())
