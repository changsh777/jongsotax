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

구현 방식:
    Playwright connect_over_cdp 대신 순수 websockets CDP 사용
    (Edge 147 + Playwright 1.59 Browser-level 연결 타임아웃 우회)

컬럼 인덱스 (40컬럼 기준, 2026-05-08 실측 + 사용자 검수):
    [0]=체크 [1]=요약 [2]=과세연월 [3]=신고서종류 [4]=신고구분
    [5]=신고유형 [6]=성명 [7]=주민번호 [8]=접수방법 [9]=접수일시
    [10]=접수번호(신고서보기 A링크) [11]=접수서류
    [12]=접수증보기(빨강버튼) [13]=납부서보기(두번째버튼)
    [14..36]=히든데이터 [37]=지방소득세이동(위택스,절대클릭금지) [38-39]=기타
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
COL_NAME      = 6   # 성명
COL_JUMIN     = 7   # 주민번호
COL_SHINGOSER = 10  # 신고서 보기 — 접수번호란 링크(A태그) 클릭
COL_RECEIPT   = 12  # 접수증 보기 버튼 (빨강/보기)
COL_TAX       = 13  # 납부서 보기 버튼 (두번째 보기)
# col[37]      = 지방소득세 이동 → 위택스 이동, 절대 클릭 금지
# COL_LOCAL   = 37  # 지방세 이동 → 위택스 별도 모듈 필요, 미구현

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


async def _get_new_tab(known_ids: set, timeout_s: int = 30):
    """known_ids 이후 생성된 새 탭 반환 (async 폴링 — 이벤트루프 블로킹 방지)"""
    for _ in range(timeout_s * 2):
        await asyncio.sleep(0.5)
        tabs = requests.get(f"http://localhost:{CDP_PORT}/json").json()
        for t in tabs:
            if t["id"] not in known_ids and "devtools" not in t.get("url", ""):
                return t
    return None


async def _dismiss_dialog_if_any(ws):
    """열려있는 JS 다이얼로그(alert/confirm)를 CDP로 즉시 닫기.
    Runtime.evaluate는 dialog 중 차단되므로 직접 WS 명령 사용.
    """
    # Page.enable 먼저 (이미 활성화돼 있어도 무해)
    await ws.send(json.dumps({"id": 990, "method": "Page.enable", "params": {}}))
    try:
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=3))
            if r.get("id") == 990:
                break
    except asyncio.TimeoutError:
        pass

    # dialog 닫기 (없으면 error 반환 — 무시)
    await ws.send(json.dumps({"id": 991, "method": "Page.handleJavaScriptDialog",
                               "params": {"accept": True}}))
    try:
        while True:
            r = json.loads(await asyncio.wait_for(ws.recv(), timeout=3))
            if r.get("id") == 991:
                if "error" not in r:
                    logger.info("JS 다이얼로그 처리 완료 (accept)")
                else:
                    logger.debug("JS 다이얼로그 없음 (정상)")
                break
    except asyncio.TimeoutError:
        pass
    await asyncio.sleep(0.3)


async def _close_popup_with_retry(ws, name: str, label: str, max_attempts: int = 14):
    """버튼 클릭 후 나타나는 개인지방세 안내 팝업 반복 닫기
    - 0.5초 간격, 최대 14회(7초) 시도
    - '확인' 버튼 우선, 없으면 '닫기'
    """
    for attempt in range(max_attempts):
        await asyncio.sleep(0.5)
        result = await _eval(ws, """(function(){
    var candidates = Array.from(document.querySelectorAll('input[type=button], button'))
        .filter(function(b){
            if (!b.offsetParent) return false;
            var r = b.getBoundingClientRect();
            if (r.width < 5 || r.height < 5) return false;
            var txt = (b.value || b.innerText || b.textContent || '').trim();
            return txt === '확인' || txt === '닫기';
        });
    if (!candidates.length) return 'none';
    // '닫기' 우선 — '확인'은 위택스 이동 위험이 있으므로 최후 수단
    var btn = candidates.find(function(b){
        return (b.value || b.innerText || '').trim() === '닫기';
    }) || candidates.find(function(b){
        return (b.value || b.innerText || '').trim() === '취소';
    }) || candidates[0];  // 마지막에 확인 (확인만 있는 단순 알림 팝업용)
    btn.click();
    return 'closed:' + (btn.value || btn.innerText || '?').trim();
})()""", cmd_id=200 + attempt)
        if result and result != 'none':
            logger.info("[%s] %s 팝업 닫기: %s (%d회 시도)", name, label, result, attempt + 1)
            await asyncio.sleep(0.5)
            return True
    logger.info("[%s] %s 팝업 미감지 (없거나 이미 닫힘)", name, label)
    return False


async def _eval(ws, code: str, cmd_id: int = 99):
    try:
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
    except asyncio.TimeoutError:
        logger.warning("_eval timeout (cmd_id=%d) — WS 유휴 또는 응답 없음", cmd_id)
        return None
    except Exception as e:
        logger.warning("_eval 오류: %s (cmd_id=%d)", e, cmd_id)
        return None


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

            # 저장/PDF 버튼 클릭 (ClipReport 종류에 따라 클래스명 다름)
            # 납부서·신고서: report_menu_save_button / 접수증: report_menu_pdf_button
            clicked = await _eval(ws, """(function() {
    var btn = document.querySelector('.report_menu_save_button')
           || document.querySelector('.report_menu_pdf_button');
    if (!btn) return 'no_btn';
    btn.disabled = false;
    btn.click();
    return 'clicked:' + btn.className.split(' ').find(function(c){ return c.includes('save') || c.includes('pdf'); });
})()""", cmd_id=3)

            if not clicked or "clicked" not in str(clicked):
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


# ── Page.printToPDF 직접 추출 ─────────────────────────────────────────

async def _print_tab_to_pdf(tab: dict, dest: Path, name: str, label: str) -> bool:
    """ClipReport 탭 내용을 CDP Page.printToPDF 로 직접 PDF 저장.
    저장 버튼 클릭 없이 CDP가 현재 탭을 통째로 PDF화 → 다운로드 이벤트 불필요.
    """
    ws_url = tab["webSocketDebuggerUrl"]
    try:
        async with websockets.connect(ws_url) as ws:
            # ClipReport 내용 로딩 대기
            # readyState 완료 + 뷰어 버튼 출현까지 최대 20초 폴링
            for _ in range(40):
                await asyncio.sleep(0.5)
                ready = await _eval(ws, """(function(){
    if (document.readyState !== 'complete') return 'loading';
    // ClipReport 뷰어 버튼 존재 여부 확인
    var btn = document.querySelector('.report_menu_save_button, .report_menu_pdf_button, #reportDiv, iframe');
    return btn ? 'ready' : 'waiting';
})()""", cmd_id=10)
                if ready == "ready":
                    logger.info("[%s] %s ClipReport 로딩 완료", name, label)
                    break
            else:
                logger.warning("[%s] %s ClipReport 로딩 20초 초과 — 강제 진행", name, label)

            # 추가 안정 대기 (뷰어 렌더링)
            await asyncio.sleep(2)

            # Page.printToPDF
            await ws.send(json.dumps({
                "id": 20, "method": "Page.printToPDF",
                "params": {
                    "printBackground": True,
                    "preferCSSPageSize": True,
                    "marginTop": 0.4,
                    "marginBottom": 0.4,
                    "marginLeft": 0.4,
                    "marginRight": 0.4,
                }
            }))
            resp = None
            for _ in range(60):  # 최대 30초
                raw = json.loads(await asyncio.wait_for(ws.recv(), timeout=35))
                if raw.get("id") == 20:
                    resp = raw
                    break

            if not resp:
                logger.warning("[%s] %s printToPDF 응답 없음", name, label)
                return False

            import base64
            data_b64 = resp.get("result", {}).get("data", "")
            if not data_b64:
                err = resp.get("error", {})
                logger.warning("[%s] %s printToPDF 데이터 없음: %s", name, label, err)
                return False

            pdf_bytes = base64.b64decode(data_b64)
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(pdf_bytes)
            logger.info("[%s] %s printToPDF 저장: %s (%d bytes)", name, label, dest.name, len(pdf_bytes))
            return True

    except Exception as e:
        logger.error("[%s] %s printToPDF 오류: %s", name, label, e)
        return False


# ── UTERNAAZ34 신고서 보기 팝업 → 일괄출력 → printToPDF ──────────────

async def _handle_uternaaz34_to_pdf(tab: dict, dest: Path, name: str) -> bool:
    """UTERNAAZ34 신고서 보기 팝업:
    1) 일괄출력 버튼 클릭 → 전체 서식 미리보기 로딩
    2) Page.printToPDF 로 직접 PDF 저장 (프린트 대화상자 우회)
    """
    ws_url = tab["webSocketDebuggerUrl"]
    try:
        async with websockets.connect(ws_url) as ws:
            # 팝업 로딩 완료 대기
            for _ in range(30):
                await asyncio.sleep(0.5)
                r = await _eval(ws, "document.readyState", cmd_id=10)
                if r == "complete":
                    break
            await asyncio.sleep(1)

            # "일괄출력" 버튼 클릭
            r = await _eval(ws, """(function(){
    var btn = Array.from(document.querySelectorAll('input[type=button],button,a'))
        .find(function(b){
            return (b.value||b.innerText||b.textContent||'').trim() === '일괄출력';
        });
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked:일괄출력';
})()""", cmd_id=11)
            logger.info("[%s] 신고서 일괄출력: %s", name, r)

            # 전체 서식 미리보기 로딩 대기 (페이지 수 1 → N 변화, 보통 10~20초 소요)
            await asyncio.sleep(20)

            # Page.printToPDF
            await ws.send(json.dumps({
                "id": 20, "method": "Page.printToPDF",
                "params": {
                    "printBackground": True,
                    "preferCSSPageSize": True,
                    "marginTop": 0.4,
                    "marginBottom": 0.4,
                    "marginLeft": 0.4,
                    "marginRight": 0.4,
                }
            }))
            resp = None
            for _ in range(120):  # 최대 60초
                raw = json.loads(await asyncio.wait_for(ws.recv(), timeout=65))
                if raw.get("id") == 20:
                    resp = raw
                    break

            if not resp:
                logger.warning("[%s] 신고서 printToPDF 응답 없음", name)
                return False

            import base64
            data_b64 = resp.get("result", {}).get("data", "")
            if not data_b64:
                logger.warning("[%s] 신고서 printToPDF 데이터 없음: %s",
                               name, resp.get("error", {}))
                return False

            pdf_bytes = base64.b64decode(data_b64)
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(pdf_bytes)
            logger.info("[%s] 신고서 저장: %s (%d bytes)", name, dest.name, len(pdf_bytes))
            return True

    except Exception as e:
        logger.error("[%s] 신고서 UTERNAAZ34 처리 오류: %s", name, e)
        return False


# ── 신고서: col[10] 접수번호 링크 → UTERNAAZ 팝업 처리 → PDF ─────────────

async def _download_shingoser(
    ws_main,
    row_idx: int,
    dest: Path,
    name: str,
) -> bool:
    """col[10] 접수번호 A링크 클릭 → 팝업 분기 처리 → PDF

    흐름:
      클릭
       ├─ UTERNAAZ39(개인정보 설정) → 적용 클릭 → 닫기
       ├─ ClipReport(접수증 탭)    → 즉시 닫기
       └─ UTERNAAZ34(신고서 보기) → 일괄출력 → printToPDF

    주의: 메인 WS recv() timeout 루프 절대 금지.
    """
    known_tabs = _get_all_tab_ids()

    clicked = await _eval(ws_main, f"""(function() {{
    var container = document.querySelector('[id*="UTERNAAZ0Z31_wframe"]');
    if (!container) return 'no_container';
    var rows = Array.from(container.querySelectorAll('table tbody tr'))
        .filter(function(tr) {{ return tr.querySelectorAll('td').length >= 13; }});
    var row = rows[{row_idx}];
    if (!row) return 'no_row';
    var tds = Array.from(row.querySelectorAll('td'));
    var cell = tds[{COL_SHINGOSER}];
    if (!cell) return 'no_cell';
    var btn = cell.querySelector('input[type=button], button, a');
    if (!btn) return 'no_btn:' + cell.innerText.trim().slice(0,20);
    btn.click();
    return 'clicked:' + (btn.value || btn.innerText || btn.textContent || '?').trim().slice(0,20);
}})()""", cmd_id=10 + row_idx)

    if not clicked or "clicked" not in str(clicked):
        logger.warning("[%s] 신고서 버튼 없음: %s", name, clicked)
        return False
    logger.info("[%s] 신고서 버튼 클릭: %s", name, clicked)

    # ── 새 탭 수집 (최대 15초, UTERNAAZ34 감지 시 조기 종료) ──────────────
    popup34_tab = None
    popup39_tab = None
    clipreport_tab = None
    seen_ids = set(known_tabs)

    for _ in range(30):
        await asyncio.sleep(0.5)
        tabs_now = requests.get(f"http://localhost:{CDP_PORT}/json").json()
        for t in tabs_now:
            if t["id"] in seen_ids:
                continue
            seen_ids.add(t["id"])
            url = t.get("url", "")
            if "UTERNAAZ34" in url:
                popup34_tab = t
                logger.info("[%s] UTERNAAZ34 팝업 감지: %s", name, url[:80])
            elif "UTERNAAZ39" in url:
                popup39_tab = t
                logger.info("[%s] UTERNAAZ39 개인정보 팝업 감지", name)
            elif "sesw.hometax" in url or "clipreport" in url.lower():
                clipreport_tab = t
                logger.info("[%s] 접수증 ClipReport 탭 감지 → 닫기 예약", name)
        if popup34_tab:
            break

    # UTERNAAZ39 (개인정보 공개여부) → 적용 클릭
    if popup39_tab:
        try:
            async with websockets.connect(popup39_tab["webSocketDebuggerUrl"]) as ws39:
                await asyncio.sleep(1.5)
                r39 = await _eval(ws39, """(function(){
    var btn = Array.from(document.querySelectorAll('input[type=button],button'))
        .find(function(b){
            return (b.value||b.innerText||'').trim() === '적용';
        });
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked:적용';
})()""", cmd_id=5)
                logger.info("[%s] UTERNAAZ39 적용: %s", name, r39)
        except Exception as e:
            logger.warning("[%s] UTERNAAZ39 처리 오류: %s", name, e)
        # 적용 후 UTERNAAZ34 추가 대기
        if not popup34_tab:
            for _ in range(20):
                await asyncio.sleep(0.5)
                tabs_now = requests.get(f"http://localhost:{CDP_PORT}/json").json()
                for t in tabs_now:
                    if t["id"] in seen_ids:
                        continue
                    seen_ids.add(t["id"])
                    url = t.get("url", "")
                    if "UTERNAAZ34" in url:
                        popup34_tab = t
                        logger.info("[%s] UTERNAAZ39→34 전환 감지", name)
                    elif "sesw.hometax" in url or "clipreport" in url.lower():
                        clipreport_tab = t
                if popup34_tab:
                    break

    # 접수증 ClipReport 탭 → 접수증은 _download_receipt(col[12])로 별도 처리, 여기선 탭만 닫기
    if clipreport_tab:
        logger.info("[%s] 접수번호 클릭으로 열린 ClipReport 탭 닫기", name)
        try:
            requests.get(f"http://localhost:{CDP_PORT}/json/close/{clipreport_tab['id']}")
        except Exception:
            pass

    # UTERNAAZ34 없으면 실패
    if not popup34_tab:
        logger.warning("[%s] UTERNAAZ34 신고서 보기 팝업 미감지 — 스킵", name)
        return False

    # UTERNAAZ34 → 일괄출력 → printToPDF
    ok = await _handle_uternaaz34_to_pdf(popup34_tab, dest, name)
    try:
        requests.get(f"http://localhost:{CDP_PORT}/json/close/{popup34_tab['id']}")
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

    # ClipReport 탭 대기 (async 폴링)
    cr_tab = await _get_new_tab(known_tabs, timeout_s=30)
    if not cr_tab:
        for _ in range(30):
            await asyncio.sleep(1)
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

    # 납부서 팝업(TERNAAZ68) 처리: '출력' 버튼 클릭해야 ClipReport 열림
    # ('전자납부 바로가기'·'개인지방소득세 신고 이동'은 위택스 이동 — 절대 클릭 금지)
    popup_handled = False
    for attempt in range(10):  # 최대 5초 대기
        await asyncio.sleep(0.5)
        r = await _eval(ws_main, """(function(){
    var containers = Array.from(document.querySelectorAll('[id*="TERNAAZ68"]'));
    for (var c of containers) {
        if (!c.offsetParent) continue;
        var r = c.getBoundingClientRect();
        if (r.width < 10 || r.height < 10) continue;
        var btns = Array.from(c.querySelectorAll('input[type=button],button'));
        var printBtn = btns.find(function(b){ return (b.value||b.innerText||'').trim() === '출력'; });
        if (printBtn) { printBtn.click(); return 'clicked:출력'; }
    }
    return 'none';
})()""", cmd_id=300 + attempt)
        if r and r != 'none':
            logger.info("[%s] %s 납부서 팝업 출력 클릭 (%d회)", name, label, attempt + 1)
            popup_handled = True
            await asyncio.sleep(0.5)
            break

    if not popup_handled:
        # 세액 없는 고객 → TERNAAZ68 팝업 없음 = 위택스 탭이 열렸을 수 있음
        logger.info("[%s] %s 납부서 팝업 없음 — 세액 없는 고객으로 판단, 스킵", name, label)
        try:
            tabs_now = requests.get(f"http://localhost:{CDP_PORT}/json").json()
            for t in tabs_now:
                url_lower = t.get("url", "").lower()
                if t["id"] not in known_tabs and ("wetax" in url_lower or "witax" in url_lower or "etax" in url_lower):
                    requests.get(f"http://localhost:{CDP_PORT}/json/close/{t['id']}")
                    logger.info("[%s] 위택스 탭 닫음: %s", name, t.get("url", "")[:80])
        except Exception as e:
            logger.warning("[%s] 위택스 탭 닫기 오류: %s", name, e)
        return False

    # ClipReport 탭 대기 (30초 폴링)
    cr_tab = await _get_new_tab(known_tabs, timeout_s=30)
    if not cr_tab:
        for _ in range(30):
            await asyncio.sleep(1)
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

    # 기존 홈택스 탭 사용 (새 탭 생성은 Edge에서 불안정 — WS 즉시 닫힘)
    ht_tab = _get_hometax_tab()
    if not ht_tab:
        logger.error("홈택스 탭 없음!")
        return
    ws_url = ht_tab["webSocketDebuggerUrl"]
    logger.info("홈택스 탭 연결: %s", ws_url[:60])

    async with websockets.connect(ws_url, ping_interval=None) as ws:
        logger.info("CDP 연결 성공!")

        # 혹시 이전 실행 후 남아있는 JS 다이얼로그(aliasDataMap 경고 등) 먼저 처리
        await _dismiss_dialog_if_any(ws)

        # 1. 페이지 로딩 대기
        await asyncio.sleep(3)
        cur_url = await _eval(ws, "window.location.href", cmd_id=99)
        logger.info("현재 URL: %s", (cur_url or "")[:80])

        if cur_url and "tmIdx=04" in cur_url and "tm2lIdx=0405000000" in cur_url:
            logger.info("이미 신고내역 페이지 — navigate 생략")
            await asyncio.sleep(2)
        else:
            logger.info("신고내역 페이지 이동...")
            await ws.send(json.dumps({
                "id": 1, "method": "Page.navigate",
                "params": {"url": RESULT_URL}
            }))
            while True:
                resp = json.loads(await asyncio.wait_for(ws.recv(), timeout=60))
                if resp.get("id") == 1:
                    break
            await asyncio.sleep(8)  # WebSquare 완전 초기화 대기

        # window.alert/confirm 무력화 — WebSquare aliasDataMap 경고 등이 native 다이얼로그로
        # 떠서 Runtime.evaluate 를 차단하는 것 방지
        await _eval(ws, """
window.alert   = function(m){ console.log('[alert]',  m); };
window.confirm = function(m){ console.log('[confirm]',m); return true; };
window.prompt  = function(m){ console.log('[prompt]', m); return ''; };
""", cmd_id=498)
        logger.info("window.alert 무력화 완료")

        # 2. 신고내역조회 팝업 열기 (이미 열려있으면 스킵 — 중복 열기 시 aliasDataMap 경고 발생)
        _already_open = await _eval(ws, f"""
document.getElementById('{SELECT_ROWNUM}') ? 'open' : 'closed'
""", cmd_id=497)
        logger.info("팝업 상태: %s", _already_open)
        if _already_open == 'open':
            logger.info("팝업 이미 열려있음 — 팝업 열기 스킵")
        else:
            logger.info("신고내역조회 팝업 열기...")
            open_result = await _eval(ws, f"""(function() {{
    var btn = document.getElementById('{BTN_RTN_POPUP}');
    if (!btn) return 'no_btn';
    btn.click();
    return 'clicked';
}})()""", cmd_id=2)
            logger.info("팝업 열기: %s", open_result)

        # 팝업 내 SELECT 가 나타날 때까지 최대 15초 폴링 (서버 느릴 때 대비)
        for _wait in range(30):
            await asyncio.sleep(0.5)
            _sel_check = await _eval(ws, f"""(function(){{
    var s = document.getElementById('{SELECT_ROWNUM}');
    return s ? 'found' : 'waiting';
}})()""", cmd_id=50)
            if _sel_check == "found":
                logger.info("팝업 로딩 완료 (%ss)", (_wait+1)*0.5)
                break
        else:
            logger.warning("팝업 SELECT 15초 이상 미감지 — 계속 진행")

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

        # ── 행별 처리: 메인 WS 유지 (팝업 계속 열려있음) ──────────────────
        for r in rows_info:
            row_idx = r["idx"]
            name    = r["name"]
            jumin6  = r.get("jumin6", "")[:6]
            logger.info("── [%d/%d] %s (%s) ──", row_idx+1, len(rows_info), name, jumin6)

            folder = find_folder(name, jumin6)
            if not folder:
                logger.warning("[%s] 고객 폴더 없음 — 스킵", name)
                continue

            # ① 접수증 (col[12] 빨강버튼 → ClipReport → PDF 버튼)
            receipt = folder / f"종합소득세 접수증 {name}.pdf"
            if receipt.exists():
                logger.info("[%s] 접수증 이미 있음 — 스킵", name)
            else:
                await _download_receipt(ws, row_idx, receipt, name)
                await asyncio.sleep(2)

            # ② 신고서 (col[10] 접수번호 클릭 → UTERNAAZ39 적용 → UTERNAAZ34 일괄출력 → printToPDF)
            shingoser = folder / f"종합소득세 신고서 {name}.pdf"
            if shingoser.exists():
                logger.info("[%s] 신고서 이미 있음 — 스킵", name)
            else:
                await _download_shingoser(ws, row_idx, shingoser, name)
                await asyncio.sleep(1)

            processed += 1
            await asyncio.sleep(1)

        logger.info("=== 스크래핑 완료: %d건 처리 ===", processed)


def run():
    asyncio.run(run_async())


if __name__ == "__main__":
    run()
