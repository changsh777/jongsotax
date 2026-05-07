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
"""

import time
import unicodedata
import logging
from datetime import date
from pathlib import Path
from playwright.sync_api import sync_playwright, BrowserContext, Page

import sys
sys.path.insert(0, str(Path(__file__).parent))
from config import CUSTOMER_DIR
from safe_save import safe_download

# ── 상수 ─────────────────────────────────────────────────────────────────
RESULT_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=04&tm2lIdx=0405000000&tm3lIdx=0405040000"
)
START_DATE = "20260501"   # 고정 시작일 (매년 5월 1일)
CDP_PORT   = 9222

# 테이블 컬럼 인덱스 (0부터, 체크박스 포함)
COL_NAME    = 5    # 상호(성명)
COL_JUMIN   = 6    # 사업자(주민)등록번호 → 주민앞6자리 추출용
COL_APPNO   = 9    # 접수번호(신고보기) ← 신고서 링크
COL_RECEIPT = 11   # 접수증 보기
COL_TAX     = 12   # 납부서 보기

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
    """
    성명 + 주민앞6자리로 고객 폴더 정확 매칭 (동명이인 방지)
    jumin6 있으면: 이름_주민앞6 정확 매칭
    jumin6 없으면: 이름_ 접두사 매칭 (후보 여러 개면 경고)
    """
    nfc_name = _nfc(name)
    if not CUSTOMER_DIR.exists():
        logger.error("CUSTOMER_DIR 없음: %s", CUSTOMER_DIR)
        return None

    if jumin6:
        # 정확 매칭: 이명회_920916
        target = f"{nfc_name}_{jumin6}"
        candidates = [
            p for p in CUSTOMER_DIR.iterdir()
            if p.is_dir() and _nfc(p.name).startswith(target)
        ]
        if not candidates:
            logger.warning("[%s_%s] 고객 폴더 없음", name, jumin6)
            return None
        if len(candidates) > 1:
            logger.warning("[%s_%s] 동명이인 여전히 복수 (%s) — 첫 번째 사용",
                           name, jumin6, [p.name for p in candidates])
        return sorted(candidates)[0]
    else:
        # jumin6 없을 때: 이름_ 접두사 매칭
        candidates = [
            p for p in CUSTOMER_DIR.iterdir()
            if p.is_dir() and _nfc(p.name).startswith(f"{nfc_name}_")
        ]
        if len(candidates) > 1:
            logger.error("[%s] 동명이인 %d명 — 주민번호 없어서 특정 불가: %s",
                         name, len(candidates), [p.name for p in candidates])
            return None
        return sorted(candidates)[0] if candidates else None


# ── ClipReport4 PDF 저장 (접수증 / 납부서) ───────────────────────────────

def _clipreport_pdf(
    ctx: BrowserContext,
    click_fn,
    dest: Path,
    name: str,
    label: str,
) -> bool:
    """
    버튼 클릭 → ClipReport 팝업 대기 → .report_menu_pdf_button → 다운로드
    """
    popup = None
    try:
        with ctx.expect_page(timeout=15000) as p_info:
            click_fn()
        popup = p_info.value
    except Exception as e:
        logger.warning("[%s] %s 팝업 안 열림: %s", name, label, e)
        return False

    try:
        # about:blank → sesw.hometax.go.kr/serp/clipreport.do 리다이렉트 대기
        for _ in range(30):
            time.sleep(1)
            if "clipreport" in popup.url.lower():
                break
        else:
            logger.warning("[%s] %s clipreport URL 미도달 (현재: %s)", name, label, popup.url[:60])
            popup.close()
            return False

        try:
            popup.wait_for_load_state("networkidle", timeout=20000)
        except Exception:
            pass
        time.sleep(2)

        with popup.expect_download(timeout=60000) as dl_info:
            clicked = popup.evaluate("""
                () => {
                    const btn = document.querySelector('.report_menu_pdf_button');
                    if (!btn) return false;
                    btn.classList.remove('report_menu_pdf_button_svg_dis');
                    btn.classList.add('report_menu_pdf_button_svg');
                    btn.disabled = false;
                    btn.click();
                    return true;
                }
            """)
            if not clicked:
                logger.warning("[%s] %s PDF 버튼 없음 (clipreport 로딩 미완?)", name, label)
                popup.close()
                return False

        dl = dl_info.value
        status, _ = safe_download(dl, dest.parent, dest.name)
        logger.info("[%s] %s 저장(%s): %s", name, label, status, dest.name)
        popup.close()
        return True

    except Exception as e:
        logger.error("[%s] %s 오류: %s", name, label, e)
        try:
            popup.close()
        except Exception:
            pass
        return False


# ── 신고서: 접수번호 링크 → 일괄출력 → ClipReport PDF ────────────────────

def _download_shingoser(
    ctx: BrowserContext,
    appno_cell,
    dest: Path,
    name: str,
) -> bool:
    """
    접수번호(신고보기) 링크 클릭 → WebSquare 신고서 뷰어 팝업
    → 일괄출력 클릭 → clipreport.do → PDF 저장
    """
    link = appno_cell.locator("a").first
    if not link.count() or not link.is_visible():
        logger.warning("[%s] 신고서 링크 없음", name)
        return False

    # 1단계: 신고서 뷰어 팝업
    viewer = None
    try:
        with ctx.expect_page(timeout=15000) as p_info:
            link.click()
        viewer = p_info.value
        try:
            viewer.wait_for_load_state("networkidle", timeout=30000)
        except Exception:
            pass
        time.sleep(3)
    except Exception as e:
        logger.warning("[%s] 신고서 뷰어 팝업 실패: %s", name, e)
        return False

    # 2단계: 일괄출력 클릭 → clipreport.do 팝업
    clipreport = None
    try:
        with ctx.expect_page(timeout=30000) as p2_info:
            viewer.click(
                "input[value='일괄출력'], button:has-text('일괄출력')",
                timeout=5000,
            )
        clipreport = p2_info.value
    except Exception:
        # fallback: ctx.pages 폴링 (일괄출력이 새 탭을 느리게 열 경우)
        logger.info("[%s] 신고서 일괄출력 expect_page fallback 폴링...", name)
        for _ in range(40):
            time.sleep(1)
            for pg in ctx.pages:
                try:
                    if "clipreport" in pg.url.lower() and pg != viewer:
                        clipreport = pg
                        break
                except Exception:
                    pass
            if clipreport:
                break

    if not clipreport:
        logger.warning("[%s] 신고서 clipreport 팝업 못 찾음", name)
        try:
            viewer.close()
        except Exception:
            pass
        return False

    # 3단계: clipreport.do → PDF 저장 버튼
    try:
        try:
            clipreport.wait_for_load_state("networkidle", timeout=20000)
        except Exception:
            pass
        time.sleep(2)

        with clipreport.expect_download(timeout=60000) as dl_info:
            clicked = clipreport.evaluate("""
                () => {
                    const btn = document.querySelector('.report_menu_pdf_button');
                    if (!btn) return false;
                    btn.classList.remove('report_menu_pdf_button_svg_dis');
                    btn.classList.add('report_menu_pdf_button_svg');
                    btn.disabled = false;
                    btn.click();
                    return true;
                }
            """)
            if not clicked:
                logger.warning("[%s] 신고서 PDF 버튼 없음", name)
                clipreport.close()
                try:
                    viewer.close()
                except Exception:
                    pass
                return False

        dl = dl_info.value
        status, _ = safe_download(dl, dest.parent, dest.name)
        logger.info("[%s] 신고서 저장(%s): %s", name, status, dest.name)
        clipreport.close()
        try:
            viewer.close()
        except Exception:
            pass
        return True

    except Exception as e:
        logger.error("[%s] 신고서 PDF 오류: %s", name, e)
        for pg in [clipreport, viewer]:
            try:
                if pg:
                    pg.close()
            except Exception:
                pass
        return False


# ── 행 처리 ──────────────────────────────────────────────────────────────

def process_row(ctx: BrowserContext, page: Page, row_idx: int):
    rows = page.locator("table tbody tr").all()
    if row_idx >= len(rows):
        return

    row  = rows[row_idx]
    cells = row.locator("td").all()
    if len(cells) <= COL_TAX:
        logger.debug("컬럼 수 부족 (%d) — 헤더행 스킵", len(cells))
        return

    name = cells[COL_NAME].inner_text().strip()
    if not name:
        return

    # 주민(사업자)번호 앞 6자리 추출: "920916-*****" → "920916"
    jumin_raw = cells[COL_JUMIN].inner_text().strip().replace("-", "").replace(" ", "")
    jumin6 = jumin_raw[:6] if len(jumin_raw) >= 6 else ""

    folder = find_folder(name, jumin6)
    if not folder:
        logger.warning("[%s] 고객 폴더 없음 — 스킵", name)
        return

    logger.info("── %s (%s) ──", name, folder.name)

    # ① 접수증
    receipt = folder / f"종합소득세 접수증 {name}.pdf"
    if receipt.exists():
        logger.info("[%s] 접수증 이미 있음 — 스킵", name)
    else:
        btn = cells[COL_RECEIPT].locator("a, button, input[type='button']").first
        if btn.count() and btn.is_visible():
            _clipreport_pdf(ctx, btn.click, receipt, name, "접수증")
        else:
            logger.warning("[%s] 접수증 버튼 없음", name)

    # ② 신고서
    shingoser = folder / f"종합소득세 신고서 {name}.pdf"
    if shingoser.exists():
        logger.info("[%s] 신고서 이미 있음 — 스킵", name)
    else:
        _download_shingoser(ctx, cells[COL_APPNO], shingoser, name)

    # ③ 납부서 (납부액 있을 때만 버튼 있음)
    taxbill = folder / f"종합소득세 납부서 {name}.pdf"
    if taxbill.exists():
        logger.info("[%s] 납부서 이미 있음 — 스킵", name)
    else:
        tax_btn = cells[COL_TAX].locator("a, button, input[type='button']").first
        if tax_btn.count() and tax_btn.is_visible():
            # 납부서 없는 경우 alert dialog 처리
            fired = []
            def _on_dialog(d):
                fired.append(d.message)
                d.dismiss()
            page.on("dialog", _on_dialog)
            try:
                _clipreport_pdf(ctx, tax_btn.click, taxbill, name, "납부서")
            finally:
                page.remove_listener("dialog", _on_dialog)
            if fired:
                logger.info("[%s] 납부서 alert: %s", name, fired[0][:80])
        else:
            logger.info("[%s] 납부서 버튼 없음 — 환급 케이스", name)


# ── 메인 ─────────────────────────────────────────────────────────────────

def run():
    today = date.today().strftime("%Y-%m-%d")
    logger.info("=== 홈택스 신고결과 스크래핑 시작: %s ~ %s ===", START_DATE, today)

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
        ctx     = browser.contexts[0]
        page    = ctx.new_page()
        page.set_default_timeout(30000)

        try:
            logger.info("신고결과 조회 페이지 이동")
            page.goto(RESULT_URL, wait_until="networkidle", timeout=60000)
            time.sleep(3)

            # 시작일 입력
            start_field = page.locator(
                "input[id*='frmDate'], input[id*='startDate'], input[id*='sDt']"
            ).first
            start_field.triple_click()
            start_field.type(START_DATE, delay=50)
            page.keyboard.press("Tab")
            time.sleep(0.5)

            # 조회 클릭
            page.click("input[value='조회'], button:has-text('조회')")
            time.sleep(3)

            # 페이지네이션
            page_no = 1
            while True:
                rows = page.locator("table tbody tr").all()
                logger.info("페이지 %d — %d건", page_no, len(rows))

                for i in range(len(rows)):
                    try:
                        process_row(ctx, page, i)
                    except Exception as e:
                        logger.error("행 %d 처리 오류: %s", i, e)
                    time.sleep(1)

                # 다음 페이지 확인
                try:
                    nxt = page.locator("a:has-text('다음 >'), a[title='다음']").first
                    if nxt.is_visible() and nxt.is_enabled():
                        nxt.click()
                        time.sleep(2)
                        page_no += 1
                    else:
                        break
                except Exception:
                    break

        finally:
            page.close()
            logger.info("=== 홈택스 스크래핑 완료 ===")


if __name__ == "__main__":
    run()
