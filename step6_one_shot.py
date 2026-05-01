"""
step6_one_shot.py - 신규 1명 즉시 처리 (Track B 전용)

용도:
  - 5월 운영: 새 고객 에어테이블 등록 즉시 안내문 받기
  - 수임동의 받기 전이라도 ID/PW로 먼저 작업 시작
  - 나중에 step4(Track A)로 수임동의 완료 고객 일괄 갱신

특징:
  - 파일명 통일 (종소세안내문_{성명}.pdf) → 나중 Track A가 자동 덮어씀
  - PDF 저장 후 파싱결과.xlsx 자동 갱신
"""
from playwright.sync_api import sync_playwright
from pathlib import Path
import time
import warnings
import logging
import sys
import argparse

sys.path.insert(0, r"F:\종소세2026")
from parse_to_xlsx import parse_anneam, write_xlsx, collect_pdfs
from safe_save import safe_download

warnings.filterwarnings("ignore")
logging.getLogger("pdfminer").setLevel(logging.ERROR)

PDF_BASE = Path(r"F:\종소세2026\output\PDF")
OUTPUT_XLSX = Path(r"F:\종소세2026\output\파싱결과.xlsx")

HOMETAX_MAIN = "https://hometax.go.kr"
LOGIN_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml&menuCd=index3"
)
TRACK_B_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml&tmIdx=41&tm2lIdx=4103000000&tm3lIdx=4103080000"
)

# 테스트 디폴트 (CLI 인자 없을 때)
DEFAULT_TEST = {
    "name": "서인미",
    "jumin": "7712252156619",
    "ht_id": "junki0613",
    "ht_pw": "miin0921*",
}


def click_visible_text(page, text):
    return page.evaluate(f"""
        () => {{
            for (const el of document.querySelectorAll('a, button, span, div, input, li')) {{
                const txt = (el.innerText || el.value || '').trim();
                if (txt === '{text}' && el.offsetParent !== null) {{
                    el.click();
                    return true;
                }}
            }}
            return false;
        }}
    """)


def is_text_visible(page, text):
    return page.evaluate(f"""
        () => {{
            for (const el of document.querySelectorAll('a, button, span, div')) {{
                const txt = (el.innerText || '').trim();
                if (txt === '{text}' && el.offsetParent !== null) return true;
            }}
            return false;
        }}
    """)


def dismiss_popups(page, rounds=3):
    for _ in range(rounds):
        time.sleep(1)
        page.evaluate("""
            () => {
                for (const el of document.querySelectorAll('button, input, a, span, div')) {
                    const txt = (el.innerText || el.value || '').replace(/\\s+/g, '');
                    if ((txt === '닫기' || txt === '오늘하루열지않음') && el.offsetParent !== null) {
                        el.click();
                    }
                }
            }
        """)


def normalize_jumin7(raw):
    s = str(raw).replace("-", "").replace(" ", "").strip()
    return s[:7]


def save_anneam_pdf(ctx, page, save_path):
    preview = page.get_by_text("미리보기", exact=False).first
    preview.wait_for(timeout=15000, state="visible")
    with ctx.expect_page(timeout=15000) as popup_info:
        preview.click()
    popup = popup_info.value
    popup.wait_for_load_state("networkidle", timeout=30000)
    time.sleep(3)
    with popup.expect_download(timeout=30000) as dl_info:
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
            popup.close()
            raise RuntimeError("PDF 저장 버튼 못 찾음")
    download = dl_info.value
    sp = Path(save_path)
    status, _ = safe_download(download, sp.parent, sp.name)
    print(f"    [저장:{status}] {sp}", flush=True)
    popup.close()


def login_and_download(ctx, page, customer, save_path):
    """본인 ID/PW 로그인 + 신고도움 PDF 저장"""
    print(f"[1] 로그아웃 + 쿠키 클리어")
    page.goto(HOMETAX_MAIN, wait_until="domcontentloaded")
    time.sleep(2)
    if is_text_visible(page, "로그아웃"):
        click_visible_text(page, "로그아웃")
        time.sleep(2)
    try:
        ctx.clear_cookies()
    except Exception:
        pass
    time.sleep(1)

    print(f"[2] 로그인 페이지 + 아이디 로그인 카드")
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    time.sleep(3)
    page.evaluate("""
        () => {
            for (const el of document.querySelectorAll('span, div, label, p, button, a')) {
                const txt = (el.innerText || '').trim();
                if (txt === '아이디 로그인' && el.offsetParent !== null) {
                    el.click();
                    return;
                }
            }
        }
    """)
    time.sleep(3)

    print(f"[3] ID/PW 입력")
    id_input = None
    for sel in ["input[placeholder*='아이디']", "input[title*='아이디']"]:
        c = [el for el in page.locator(sel).all() if el.is_visible()]
        if c:
            id_input = c[0]
            break
    pw_inputs = [el for el in page.locator("input[type='password']").all() if el.is_visible()]
    if not id_input or not pw_inputs:
        raise RuntimeError("로그인 폼 못 찾음")
    id_input.fill(customer["ht_id"])
    pw_inputs[0].fill(customer["ht_pw"])
    time.sleep(1)

    print(f"[4] 로그인 클릭")
    login_btns = [b for b in page.locator(
        "xpath=//button[normalize-space()='로그인'] | //input[@type='button' and @value='로그인']"
    ).all() if b.is_visible()]
    if not login_btns:
        raise RuntimeError("로그인 버튼 못 찾음")
    login_btns[-1].click()
    time.sleep(2)

    print(f"[5] 2차 인증 처리")
    has_2fa = False
    for _ in range(24):
        time.sleep(0.5)
        state = page.evaluate("""
            () => {
                const all = document.body.innerText || '';
                return {
                    twofa: all.includes('아이디 로그인 2차 인증'),
                    failure: all.includes('로그인 정보가 없습니다') || all.includes('일치하지 않'),
                };
            }
        """)
        if state["twofa"]:
            has_2fa = True
            break
        if state["failure"]:
            raise RuntimeError("로그인 실패 - ID/PW 불일치")

    if has_2fa:
        jumin7 = normalize_jumin7(customer["jumin"])
        jumin_inputs = page.locator(
            "xpath=//*[contains(text(),'주민등록번호')]"
            "/following::input[not(@type='button')][not(@type='checkbox')][position()<=4]"
        ).all()
        visible = [el for el in jumin_inputs if el.is_visible()]
        empty = []
        for el in visible:
            ph = el.get_attribute("placeholder") or ""
            if "검색" in ph or "아이디" in ph or "비밀번호" in ph:
                continue
            try:
                if el.input_value() == "":
                    empty.append(el)
            except Exception:
                pass
        if len(empty) < 2:
            raise RuntimeError("2차 인증 입력 필드 부족")
        empty[0].fill(jumin7[:6])
        empty[1].fill(jumin7[6:])
        time.sleep(0.5)
        click_visible_text(page, "확인")
        time.sleep(4)

    dismiss_popups(page, rounds=3)

    print(f"[6] 신고도움 페이지로 이동")
    page.goto(TRACK_B_URL, wait_until="domcontentloaded")
    time.sleep(4)
    dismiss_popups(page, rounds=2)

    print(f"[7] 조회 + 미리보기 + PDF 저장")
    click_visible_text(page, "조회")
    time.sleep(3)
    save_anneam_pdf(ctx, page, save_path)
    print(f"    저장: {save_path}")


def update_xlsx():
    """parse_to_xlsx.main() 호출 - 로컬 + 구글시트 동시 갱신"""
    from parse_to_xlsx import main as parse_main
    parse_main(sync_gsheet=True)


def parse_args():
    p = argparse.ArgumentParser(description="신규 고객 1명 즉시 처리 (Track B)")
    p.add_argument("--name", help="성명")
    p.add_argument("--jumin", help="주민번호 13자리")
    p.add_argument("--id", dest="ht_id", help="홈택스 아이디")
    p.add_argument("--pw", dest="ht_pw", help="홈택스 비밀번호")
    return p.parse_args()


def main():
    args = parse_args()
    if args.name and args.jumin and args.ht_id and args.ht_pw:
        customer = {
            "name": args.name,
            "jumin": args.jumin,
            "ht_id": args.ht_id,
            "ht_pw": args.ht_pw,
        }
    else:
        print("[테스트 모드] 인자 없음 - 디폴트(서인미) 사용")
        customer = DEFAULT_TEST

    folder = PDF_BASE / customer["name"]
    folder.mkdir(parents=True, exist_ok=True)
    # 파일명 통일 (Track A 갱신 시 자동 덮어씀)
    save_path = folder / f"종소세안내문_{customer['name']}.pdf"

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()
        page.on("dialog", lambda d: d.dismiss())

        try:
            login_and_download(ctx, page, customer, save_path)
            print(f"\n{'='*50}")
            print(f"[완료] {customer['name']}")
            update_xlsx()
        except Exception as e:
            print(f"\n[실패] {customer['name']}: {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
