"""
hometax_result_scraper.py — 홈택스 신고결과(접수증·신고서·납부서) 스크래핑
세무회계창연 | 2026

사전조건:
    Edge --remote-debugging-port=9222 으로 홈택스 로그인 완료 상태

실행:
    python hometax_result_scraper.py

파일명 규칙:
    종합소득세 접수증 {이름}.pdf
    종합소득세 신고서 {이름}.pdf
    종합소득세 납부서 {이름}.pdf   (납부액 있을 때만)
    지방소득세 납부서 {이름}.pdf   (지방세 납부액 있을 때만)

구현 방식:
    Playwright connect_over_cdp 대신 순수 websockets CDP 사용
    (Edge 147 + Playwright 1.59 Browser-level 연결 타임아웃 우회)

컬럼 인덱스 (40컬럼 기준, 2026-05-07 실측):
    [0]=체크 [1]=요약 [2]=과세연월 [3]=신고서종류 [4]=신고구분
    [5]=신고유형 [6]=성명 [7]=주민번호 [8]=접수방법 [9]=접수일시
    [10]=접수번호(링크) [11]=접수서류 [12]=접수증보기(빨강버튼)
    [13]=두번째보기 [14..36]=히든데이터 [37]=납부서이동 [38]=- [39]=지방세이동
"""

import sys, io, os, time, json, asyncio, unicodedata, logging, requests
from pathlib import Path
from datetime import date

# NAS 경로 사용 (Z:\종소세2026\고객 폴더)
os.environ.setdefault("SEOTAX_ENV", "nas")

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
elif hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)

sys.path.insert(0, str(Path(__file__).parent))
from config import CUSTOMER_DIR
# safe_save: CDP 다운로드 방식에서는 불필요 (Page.setDownloadBehavior 사용)

import websockets

# ── 상수 ─────────────────────────────────────────────────────────────────
RESULT_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=04&tm2lIdx=0405000000&tm3lIdx=0405040000"
)
START_DATE  = "20260501"
CDP_PORT    = 9222

# 컬럼 인덱스 (40컬럼 실측)
COL_NAME    = 6     # 성명
COL_JUMIN   = 7     # 주민번호
COL_APPNO   = 10    # 접수번호(신고서 뷰어 링크)
COL_RECEIPT = 12    # 접수증 보기 (빨강버튼)
COL_TAX     = 37    # 납부서 이동
COL_LOCAL   = 39    # 지방세 납부서 이동

# 신고내역 팝업 건수 select ID
SELECT_ROWNUM = "mf_txppWframe_UTERNAAZ0Z31_wframe_edtGrdRowNum"
# 조회 버튼 ID
BTN_SEARCH    = "mf_txppWframe_UTERNAAZ0Z31_wframe_trigger70_UTERNAAZ31"
# 신고내역 팝업 열기 버튼
BTN_RTN_POPUP = "mf_txppWframe_btnRtnInqr"

LOG_FILE = Path.home() / "종소세2026" / "hometax_result_scraper.log"
LOG_FILE.parent.mkdir(exist_ok=True)
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO,
    handlers=[
        logging.FileHandler(str(LOG_FILE), encoding="utf-8"),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)


# ── 유틸 ─────────────────────────────────────────────────────────────────

def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def find_folder(name: str, jumin6: str = "") -> Path | None:
    nfc_name = _nfc(name)
    if not CUSTOMER_DIR.exists():
        logger.error("CUSTOMER_DIR 없음: %s", CUSTOMER_DIR)
        return None
    if jumin6:
        target = f"{nfc_name}_{jumin6}"
        candidates = [p for p in CUSTOMER_DIR.iterdir()
                      if p.is_dir() and _nfc(p.name).startswith(target)]
        if not candidates:
            logger.warning("[%s_%s] 고객 폴더 없음", name, jumin6)
            return None
        if len(candidates) > 1:
            logger.warning("[%s_%s] 복수 폴더 — 첫 번째 사용", name, jumin6)
        return sorted(candidates)[0]
    else:
        candidates = [p for p in CUSTOMER_DIR.iterdir()
                      if p.is_dir() and _nfc(p.name).startswith(f"{nfc_name}_")]
        if len(candidates) > 1:
            logger.error("[%s] 동명이인 %d명 — 특정 불가", name, len(candidates))
            return None
        return sorted(candidates)[0] if candidates else None


# ── CDP 헬퍼 ─────────────────────────────────────────────────────────────

def _get_hometax_tab():
    """메인 홈택스 SPA 탭 반환 (sesw.hometax.go.kr ClipReport 탭 제외)"""
    tabs = requests.get(f"http://localhost:{CDP_PORT}/json").json()
    # websquare.html 포함하는 탭 우선 (메인 SPA)
    main = next((t for t in tabs
                 if "hometax.go.kr" in t.get("url","")
                 and "websquare" in t.get("url","")
                 and "sesw." not in t.get("url","")), None)
    if main:
        return main
    # fallback: sesw. 제외한 hometax.go.kr 탭
    return next((t for t in tabs
                 if "hometax.go.kr" in t.get("url","")
                 and "sesw." not in t.get("url","")), None)


def _get_all_tab_ids():
    return {t["id"] for t in requests.get(f"http://localhost:{CDP_PORT}/json").json()}


def _get_new_tab(known_ids: set, timeout_s: int = 30):
    """known_ids 이후 생성된 새 탭 반환 (폴링)"""
    for _ in range(timeout_s * 2):
        time.sleep(0.5)
        tabs = requests.get(f"http://localhost:{CDP_PORT}/json").json()
        for t in tabs:
            if t["id"] not in known_ids and "devtools" not in t.get("url", ""):
                return t
    return None


async def _eval(ws, code: str, cmd_id: int = 99):
    await ws.send(json.dumps({
        "id": cmd_id, "method": "Runtime.evaluate",
        "params": {"expression": code, "returnByValue": True, "awaitPromise": True}
    }))
    while True:
        resp = json.loads(await asyncio.wait_for(ws.recv(), timeout=30))
        if resp.get("id") == cmd_id:
            r = resp.get("result", {}).get("result", {})
            if r.get("type") == "string":
                return r.get("value")
            val = r.get("value")
            if isinstance(val, str):
                try:
                    return json.loads(val)
                except Exception:
                    return val
            return val


# ── ClipReport PDF 다운로드 ─────────────────────────────────────────────

async def _download_from_clipreport(tab: dict, dest: Path, name: str, label: str) -> bool:
    """ClipReport 탭에서 PDF 다운로드 버튼 클릭 후 파일 저장"""
    ws_url = tab["webSocketDebuggerUrl"]
    # 임시 다운로드 폴더 (고객 폴더와 분리해서 충돌 방지)
    import tempfile
    tmp_dir = Path(tempfile.gettempdir()) / "hometax_pdfs"
    tmp_dir.mkdir(exist_ok=True)

    try:
        async with websockets.connect(ws_url) as ws:
            # 임시 폴더에 다운로드
            await ws.send(json.dumps({
                "id": 1, "method": "Page.setDownloadBehavior",
                "params": {"behavior": "allow", "downloadPath": str(tmp_dir)}
            }))
            while True:
                resp = json.loads(await asyncio.wait_for(ws.recv(), timeout=10))
                if resp.get("id") == 1:
                    break

            # 페이지 로딩 대기
            await asyncio.sleep(3)

            # PDF 버튼 클릭
            clicked = await _eval(ws, """(function() {
    var btn = document.querySelector('.report_menu_pdf_button');
    if (!btn) return 'no_btn';
    btn.classList.remove('report_menu_pdf_button_svg_dis');
    btn.classList.add('report_menu_pdf_button_svg');
    btn.disabled = false;
    btn.click();
    return 'clicked';
})()""", cmd_id=3)

            if clicked != "clicked":
                logger.warning("[%s] %s PDF 버튼 없음 (%s)", name, label, clicked)
                return False

            logger.info("[%s] %s PDF 버튼 클릭 — 다운로드 대기...", name, label)

            # Page.downloadProgress 이벤트 수신
            await ws.send(json.dumps({"id": 4, "method": "Page.enable", "params": {}}))
            while True:
                resp = json.loads(await asyncio.wait_for(ws.recv(), timeout=5))
                if resp.get("id") == 4:
                    break

            # 다운로드 완료 대기 (최대 90초)
            suggested_name = ""
            download_done = False
            for _ in range(180):
                await asyncio.sleep(0.5)
                try:
                    resp_raw = await asyncio.wait_for(ws.recv(), timeout=0.1)
                    resp2 = json.loads(resp_raw)
                    method = resp2.get("method", "")
                    if method == "Page.downloadWillBegin":
                        suggested_name = resp2.get("params", {}).get("suggestedFilename", "")
                        logger.info("[%s] %s 다운로드 시작: %s", name, label, suggested_name)
                    elif method == "Page.downloadProgress":
                        p = resp2.get("params", {})
                        if p.get("state") == "completed":
                            logger.info("[%s] %s 다운로드 완료", name, label)
                            download_done = True
                            break
                        elif p.get("state") == "canceled":
                            logger.warning("[%s] %s 다운로드 취소됨", name, label)
                            break
                except asyncio.TimeoutError:
                    pass

            if not download_done:
                logger.warning("[%s] %s 다운로드 이벤트 없음 — 파일 탐색 시도", name, label)

            # 임시 폴더에서 최신 PDF 찾기 → 고객 폴더로 이동
            await asyncio.sleep(1)
            pdf_files = sorted(tmp_dir.glob("*.pdf"), key=lambda f: f.stat().st_mtime, reverse=True)
            if pdf_files:
                src = pdf_files[0]
                dest.parent.mkdir(parents=True, exist_ok=True)
                import shutil
                shutil.move(str(src), str(dest))
                logger.info("[%s] %s 저장: %s → %s", name, label, src.name, dest)
                return True
            else:
                logger.warning("[%s] %s 임시 폴더에 PDF 없음 (tmp=%s)", name, label, tmp_dir)
                return False

    except Exception as e:
        logger.error("[%s] %s ClipReport 오류: %s", name, label, e)
        return False


# ── 신고서: 접수번호 링크 → 일괄출력 → ClipReport ──────────────────────

async def _download_shingoser(
    ws_main,
    appno_cell_idx: int,
    row_idx: int,
    dest: Path,
    name: str,
) -> bool:
    """접수번호(신고보기) 링크 클릭 → 뷰어 팝업 → 일괄출력 → clipreport → PDF"""
    known_tabs = _get_all_tab_ids()

    # 링크 클릭 (JS)
    clicked = await _eval(ws_main, f"""(function() {{
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no_container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(tr) {{ return tr.querySelectorAll('td').length >= 13; }});
    var row = rows[{row_idx}];
    if (!row) return 'no_row';
    var tds = Array.from(row.querySelectorAll('td'));
    var cell = tds[{COL_APPNO}];
    if (!cell) return 'no_cell';
    var a = cell.querySelector('a');
    if (!a) return 'no_link';
    a.click();
    return 'clicked';
}})()""", cmd_id=10 + row_idx)

    if clicked != "clicked":
        logger.warning("[%s] 신고서 링크 없음: %s", name, clicked)
        return False

    # 뷰어 팝업 대기
    viewer_tab = _get_new_tab(known_tabs, timeout_s=20)
    if not viewer_tab:
        logger.warning("[%s] 신고서 뷰어 팝업 못 찾음", name)
        return False

    logger.info("[%s] 신고서 뷰어 탭: %s", name, viewer_tab.get("url","")[:60])

    # 뷰어 로딩 대기
    await asyncio.sleep(4)

    # 일괄출력 클릭
    known_tabs2 = _get_all_tab_ids()
    try:
        async with websockets.connect(viewer_tab["webSocketDebuggerUrl"]) as ws_viewer:
            await asyncio.sleep(2)
            result = await _eval(ws_viewer, """(function() {
    var btn = Array.from(document.querySelectorAll("input[value='일괄출력'], button"))
        .find(function(el) { return (el.value || el.innerText || '').trim() === '일괄출력'; });
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked';
})()""", cmd_id=20)
            logger.info("[%s] 신고서 일괄출력 클릭: %s", name, result)
    except Exception as e:
        logger.warning("[%s] 신고서 뷰어 연결 오류: %s", name, e)
        return False

    # ClipReport 탭 대기
    cr_tab = _get_new_tab(known_tabs2, timeout_s=30)
    if not cr_tab:
        # fallback: url에 clipreport 포함 탭 탐색
        for _ in range(40):
            time.sleep(1)
            tabs = requests.get(f"http://localhost:{CDP_PORT}/json").json()
            cr_tab = next((t for t in tabs if "clipreport" in t.get("url","").lower()), None)
            if cr_tab:
                break

    if not cr_tab:
        logger.warning("[%s] 신고서 clipreport 탭 못 찾음", name)
        return False

    # PDF 저장
    ok = await _download_from_clipreport(cr_tab, dest, name, "신고서")

    # 탭 닫기
    for tab in [cr_tab, viewer_tab]:
        try:
            requests.get(f"http://localhost:{CDP_PORT}/json/close/{tab['id']}")
        except Exception:
            pass

    return ok


# ── 접수증 다운로드 ─────────────────────────────────────────────────────

async def _download_receipt(
    ws_main,
    row_idx: int,
    dest: Path,
    name: str,
) -> bool:
    """접수증 보기 버튼(col[12], 빨강버튼) 클릭 → ClipReport → PDF"""
    known_tabs = _get_all_tab_ids()

    clicked = await _eval(ws_main, f"""(function() {{
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no_container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(tr) {{ return tr.querySelectorAll('td').length >= 13; }});
    var row = rows[{row_idx}];
    if (!row) return 'no_row';
    var tds = Array.from(row.querySelectorAll('td'));
    var cell = tds[{COL_RECEIPT}];
    if (!cell) return 'no_cell';
    var btn = cell.querySelector('input[type=button], button');
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked:' + (btn.value || '?');
}})()""", cmd_id=30 + row_idx)

    if not clicked or "clicked" not in str(clicked):
        logger.warning("[%s] 접수증 버튼 없음: %s", name, clicked)
        return False

    logger.info("[%s] 접수증 버튼 클릭: %s", name, clicked)

    # ClipReport 탭 대기 (새 탭 폴링)
    cr_tab = _get_new_tab(known_tabs, timeout_s=30)
    if not cr_tab:
        # URL에 clipreport 포함 탭도 탐색
        for _ in range(30):
            time.sleep(1)
            tabs = requests.get(f"http://localhost:{CDP_PORT}/json").json()
            cr_tab = next((t for t in tabs
                          if t["id"] not in known_tabs
                          and "clipreport" in t.get("url","").lower()), None)
            if cr_tab:
                break

    if not cr_tab:
        logger.warning("[%s] 접수증 clipreport 탭 못 찾음", name)
        return False

    ok = await _download_from_clipreport(cr_tab, dest, name, "접수증")
    try:
        requests.get(f"http://localhost:{CDP_PORT}/json/close/{cr_tab['id']}")
    except Exception:
        pass
    return ok


# ── 납부서 다운로드 ─────────────────────────────────────────────────────

async def _download_taxbill(
    ws_main,
    col_idx: int,
    row_idx: int,
    dest: Path,
    name: str,
    label: str,
) -> bool:
    """납부서/지방세 이동 버튼 클릭 → ClipReport → PDF"""
    known_tabs = _get_all_tab_ids()

    clicked = await _eval(ws_main, f"""(function() {{
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no_container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(tr) {{ return tr.querySelectorAll('td').length >= 38; }});
    var row = rows[{row_idx}];
    if (!row) return 'no_row';
    var tds = Array.from(row.querySelectorAll('td'));
    var cell = tds[{col_idx}];
    if (!cell) return 'no_cell';
    var cellText = cell.innerText.trim();
    if (cellText === '-' || cellText === '') return 'empty';
    var btn = cell.querySelector('input[type=button], button');
    if (!btn) return 'no_btn_in_cell:' + cellText;
    btn.click();
    return 'clicked:' + (btn.value || '?');
}})()""", cmd_id=50 + row_idx)

    if not clicked or "clicked" not in str(clicked):
        logger.info("[%s] %s 버튼 없음 (환급 또는 납부없음): %s", name, label, clicked)
        return False

    logger.info("[%s] %s 버튼 클릭: %s", name, label, clicked)

    # ClipReport 탭 대기
    cr_tab = _get_new_tab(known_tabs, timeout_s=30)
    if not cr_tab:
        for _ in range(30):
            time.sleep(1)
            tabs = requests.get(f"http://localhost:{CDP_PORT}/json").json()
            cr_tab = next((t for t in tabs
                          if t["id"] not in known_tabs
                          and "clipreport" in t.get("url","").lower()), None)
            if cr_tab:
                break

    if not cr_tab:
        logger.warning("[%s] %s clipreport 탭 못 찾음", name, label)
        return False

    ok = await _download_from_clipreport(cr_tab, dest, name, label)
    try:
        requests.get(f"http://localhost:{CDP_PORT}/json/close/{cr_tab['id']}")
    except Exception:
        pass
    return ok


# ── 데이터 행 정보 추출 ─────────────────────────────────────────────────

async def _get_data_rows_info(ws) -> list:
    """신고내역 팝업 데이터 행들 → [(row_idx, name, jumin6, row_el_exists)]"""
    result = await _eval(ws, """(function() {
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return JSON.stringify({error: 'no container'});
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(tr) { return tr.querySelectorAll('td').length >= 13; });
    var data = rows.map(function(tr, idx) {
        var tds = Array.from(tr.querySelectorAll('td'));
        var name = tds[6] ? tds[6].innerText.trim() : '';
        var jumin = tds[7] ? tds[7].innerText.trim().replace(/[^0-9\\*]/g,'').slice(0,6) : '';
        return {idx: idx, name: name, jumin6: jumin};
    }).filter(function(r) { return r.name.length > 0; });
    return JSON.stringify(data);
})()""", cmd_id=90)

    if isinstance(result, list):
        return result
    if isinstance(result, str):
        try:
            parsed = json.loads(result)
            if isinstance(parsed, list):
                return parsed
        except Exception:
            pass
    return []


# ── 100건 보기 설정 ─────────────────────────────────────────────────────

async def _set_page_100(ws):
    """건수 select를 100건으로 변경 후 조회 버튼 클릭"""
    result = await _eval(ws, f"""(function() {{
    var sel = document.getElementById('{SELECT_ROWNUM}');
    if (!sel) return 'no_select';
    // 100건 옵션 찾기
    var opt = Array.from(sel.options).find(function(o) {{
        return o.text.includes('100') || o.value === '100';
    }});
    if (!opt) return 'no_100_option: ' + Array.from(sel.options).map(function(o){{return o.value;}}).join(',');
    sel.value = opt.value;
    // change 이벤트 발화
    var evt = new Event('change', {{bubbles: true}});
    sel.dispatchEvent(evt);
    return 'set_100:' + opt.value;
}})()""", cmd_id=91)
    logger.info("100건 설정: %s", result)
    await asyncio.sleep(1)

    # 조회 버튼 클릭
    r2 = await _eval(ws, f"""(function() {{
    var btn = document.getElementById('{BTN_SEARCH}');
    if (!btn) {{
        // fallback: value='조회' 버튼
        btn = Array.from(document.querySelectorAll("input[type=button]"))
            .filter(function(el) {{ return el.value === '조회'; }})
            .slice(-1)[0];
    }}
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked';
}})()""", cmd_id=92)
    logger.info("조회 클릭: %s", r2)
    await asyncio.sleep(3)


# ── 메인 ─────────────────────────────────────────────────────────────────

async def run_async():
    today = date.today().strftime("%Y-%m-%d")
    logger.info("=== 홈택스 신고결과 스크래핑 시작: %s ~ %s ===", START_DATE, today)

    ht_tab = _get_hometax_tab()
    if not ht_tab:
        logger.error("홈택스 탭 없음! Edge에서 홈택스 로그인 후 다시 실행하세요.")
        return

    ws_url = ht_tab["webSocketDebuggerUrl"]
    logger.info("홈택스 탭 연결: %s", ws_url[:60])

    async with websockets.connect(ws_url) as ws:
        logger.info("CDP 연결 성공!")

        # 1. 신고내역 페이지 이동
        logger.info("신고내역 페이지 이동...")
        await ws.send(json.dumps({
            "id": 1, "method": "Page.navigate",
            "params": {"url": RESULT_URL}
        }))
        while True:
            resp = json.loads(await asyncio.wait_for(ws.recv(), timeout=60))
            if resp.get("id") == 1:
                break
        await asyncio.sleep(4)

        # 2. 신고내역조회 팝업 열기
        logger.info("신고내역조회 팝업 열기...")
        open_result = await _eval(ws, f"""(function() {{
    var btn = document.getElementById('{BTN_RTN_POPUP}');
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked';
}})()""", cmd_id=2)
        logger.info("팝업 열기: %s", open_result)
        await asyncio.sleep(3)

        # 3. 날짜는 기본 1개월 그대로 사용 (변경 시 WebSquare 검증 오류 발생)
        # 1개월 버튼 클릭으로 초기화 (혹시 다른 범위로 되어있을 경우 대비)
        await _eval(ws, f"""(function() {{
    var btn = document.getElementById('mf_txppWframe_UTERNAAZ0Z31_wframe_btnSch1Month_UTERNAAZ31');
    if (btn) {{ btn.click(); return 'clicked_1month'; }}
    return 'not_found';
}})()""", cmd_id=3)
        await asyncio.sleep(1)

        # 4. 100건 보기 + 조회
        await _set_page_100(ws)

        # 5. 조회완료 알림 닫기 (JS alert 처리 + Enter 키)
        await asyncio.sleep(1)
        # JS confirm/alert 자동 닫기 (WebSquare 알림 대비)
        await _eval(ws, """(function() {
    // WebSquare 알림 버튼 (확인/닫기) 자동 클릭
    var confirms = Array.from(document.querySelectorAll("input[value='확인'], button"))
        .filter(function(el) {
            var txt = (el.value || el.innerText || '').trim();
            return txt === '확인' || txt === '닫기';
        });
    // 팝업/알림이 있는 경우만
    confirms.forEach(function(btn) {
        var r = btn.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) btn.click();
    });
    return 'ok';
})()""", cmd_id=5)
        await asyncio.sleep(5)  # 조회 결과 로딩 대기 충분히

        # 6. 데이터 행 수집
        logger.info("데이터 행 수집...")
        rows_info = await _get_data_rows_info(ws)
        if not rows_info:
            logger.warning("데이터 없음. 팝업이 열렸는지 확인하세요.")
            return

        logger.info("총 %d건 처리 시작", len(rows_info))

        processed = 0
        for r in rows_info:
            row_idx = r["idx"]
            name    = r["name"]
            jumin6  = r.get("jumin6", "")[:6]
            logger.info("── [%d/%d] %s (%s) ──", row_idx+1, len(rows_info), name, jumin6)

            folder = find_folder(name, jumin6)
            if not folder:
                logger.warning("[%s] 고객 폴더 없음 — 스킵", name)
                continue

            # ① 접수증
            receipt = folder / f"종합소득세 접수증 {name}.pdf"
            if receipt.exists():
                logger.info("[%s] 접수증 이미 있음 — 스킵", name)
            else:
                await _download_receipt(ws, row_idx, receipt, name)
                await asyncio.sleep(1)

            # ② 신고서
            shingoser = folder / f"종합소득세 신고서 {name}.pdf"
            if shingoser.exists():
                logger.info("[%s] 신고서 이미 있음 — 스킵", name)
            else:
                await _download_shingoser(ws, COL_APPNO, row_idx, shingoser, name)
                await asyncio.sleep(1)

            # ③ 납부서 (종합소득세)
            taxbill = folder / f"종합소득세 납부서 {name}.pdf"
            if taxbill.exists():
                logger.info("[%s] 종소세 납부서 이미 있음 — 스킵", name)
            else:
                await _download_taxbill(ws, COL_TAX, row_idx, taxbill, name, "종소세납부서")
                await asyncio.sleep(1)

            # ④ 지방세 납부서
            local = folder / f"지방소득세 납부서 {name}.pdf"
            if local.exists():
                logger.info("[%s] 지방세 납부서 이미 있음 — 스킵", name)
            else:
                await _download_taxbill(ws, COL_LOCAL, row_idx, local, name, "지방세납부서")
                await asyncio.sleep(1)

            processed += 1
            await asyncio.sleep(1)

    logger.info("=== 스크래핑 완료: %d건 처리 ===", processed)


def run():
    asyncio.run(run_async())


if __name__ == "__main__":
    run()
