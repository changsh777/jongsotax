"""
Track B: 본인 ID/PW 로그인 → 안내문 PDF 다운로드 (batch)

흐름:
  1. 입력 엑셀에서 ID/PW 있는 고객만 추출
  2. 각 고객마다:
     a. 로그아웃 (이전 세션 정리)
     b. ID/PW + 2차 인증으로 로그인
     c. 신고도움 서비스 페이지 직접 이동 (딥링크)
     d. 조회 → 미리보기 → ClipReport PDF 저장
  3. 결과 로그
"""
from playwright.sync_api import sync_playwright
from pathlib import Path
import openpyxl
import time
import warnings
import logging
import sys
sys.path.insert(0, r"F:\종소세2026")
from safe_save import safe_download

warnings.filterwarnings("ignore")
logging.getLogger("pdfminer").setLevel(logging.ERROR)

import os
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")
from config import CUSTOMER_DIR, customer_folder as _customer_folder

INPUT_XLSX = Path(r"F:\종소세2026\input\종소세신고도움서비스테스트.xlsx")
PDF_BASE = CUSTOMER_DIR  # NAS: Z:\종소세2026\고객

HOMETAX_MAIN = "https://hometax.go.kr"
LOGIN_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml&menuCd=index3"
)
REPORT_HELP_URL_PERSONAL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml&tmIdx=41&tm2lIdx=4103000000&tm3lIdx=4103080000"
)


# -------------------- 유틸 --------------------

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
    """주민번호 → 앞 7자리 (숫자만)"""
    s = str(raw).replace("-", "").replace(" ", "").strip()
    return s[:7]


# -------------------- 로그인/로그아웃 --------------------

def force_logout(ctx, page):
    """쿠키 모두 삭제 + UI 로그아웃 = 강제 세션 종료"""
    # 1. UI 로그아웃 시도 (정상적인 경우)
    try:
        page.goto(HOMETAX_MAIN, wait_until="domcontentloaded")
        time.sleep(2)
        if is_text_visible(page, "로그아웃"):
            click_visible_text(page, "로그아웃")
            time.sleep(2)
    except Exception:
        pass
    # 2. 쿠키 전부 삭제 (확실한 세션 종료)
    try:
        ctx.clear_cookies()
    except Exception as e:
        print(f"    쿠키 삭제 예외: {e}")
    time.sleep(1)


def login_via_id_pw(page, ht_id, ht_pw, jumin7):
    """ID/PW + 2차 인증 로그인. 성공 True / 실패 False"""
    print(f"    [Login] {ht_id}")

    # 로그인 페이지 직접 이동
    page.goto(LOGIN_URL, wait_until="domcontentloaded")
    time.sleep(4)
    print(f"    [Login] URL 진입 후: {page.url}")

    # '아이디 로그인' 카드 클릭 - 부모 클릭 가능 element까지 trampoline
    print(f"    [Login] '아이디 로그인' 카드 클릭 시도")
    click_result = page.evaluate("""
        () => {
            // 정확한 텍스트 매칭 + 클릭 가능한 컨테이너 찾기
            const all = document.querySelectorAll('span, div, label, p, button, a');
            for (const el of all) {
                const txt = (el.innerText || '').trim();
                if (txt === '아이디 로그인' && el.offsetParent !== null) {
                    // 1) 자기 클릭
                    el.click();
                    // 2) 부모도 클릭 (안전망)
                    let p = el.parentElement;
                    let depth = 0;
                    while (p && depth < 5) {
                        if (p.onclick || p.classList.contains('card') || p.tagName === 'BUTTON' || p.tagName === 'A') {
                            p.click();
                            return {clicked: true, parent: p.tagName, cls: String(p.className).slice(0, 50)};
                        }
                        p = p.parentElement;
                        depth++;
                    }
                    return {clicked: true, parent: null};
                }
            }
            return {clicked: false};
        }
    """)
    print(f"    [Login] 카드 클릭 결과: {click_result}")
    time.sleep(3)

    # 클릭 후 보이는 input 진단
    visible_inputs_info = page.evaluate("""
        () => {
            const inputs = document.querySelectorAll('input');
            const visible = [];
            for (const el of inputs) {
                if (el.offsetParent !== null) {
                    visible.push({
                        type: el.type,
                        ph: el.placeholder || '',
                        name: el.name || '',
                    });
                }
            }
            return visible;
        }
    """)
    print(f"    [Login] 보이는 input들: {visible_inputs_info}")

    # ID 입력 - placeholder '아이디' 가진 input 우선
    id_input = None
    for sel in [
        "input[placeholder*='아이디']",
        "input[placeholder*='ID']",
        "input[title*='아이디']",
    ]:
        candidates = [el for el in page.locator(sel).all() if el.is_visible()]
        if candidates:
            id_input = candidates[0]
            print(f"    [Login] ID input 셀렉터: {sel}")
            break

    if not id_input:
        # fallback: 검색창 제외하고 ID 추정
        all_text = [el for el in page.locator("input[type='text']").all() if el.is_visible()]
        # placeholder가 '검색어'인 것 제외
        non_search = []
        for el in all_text:
            ph = (el.get_attribute("placeholder") or "").lower()
            if "검색" not in ph and "search" not in ph:
                non_search.append(el)
        if non_search:
            id_input = non_search[0]
            print(f"    [Login] ID input (검색 제외 첫번째)")
        else:
            print(f"    [Login] ID input 못 찾음")
            return False

    pw_inputs = [el for el in page.locator("input[type='password']").all() if el.is_visible()]
    if not pw_inputs:
        print(f"    [Login] PW input 못 찾음")
        return False

    # 입력 - fill 후 실제 값 검증 + keyboard.type fallback
    id_input.fill(ht_id)
    pw_inputs[0].fill(ht_pw)
    time.sleep(1)

    filled_id = id_input.input_value()
    filled_pw = pw_inputs[0].input_value()
    print(f"    [Login] 채워진 ID: '{filled_id}' (길이 {len(filled_id)})")
    print(f"    [Login] 채워진 PW: '{'*' * len(filled_pw)}' (길이 {len(filled_pw)})")

    # fill이 안 먹혔으면 keyboard.type으로 재시도
    if filled_id != ht_id or len(filled_pw) != len(ht_pw):
        print(f"    [Login] fill 실패 감지 → keyboard.type 재시도")
        id_input.click()
        page.keyboard.press("Control+a")
        page.keyboard.type(ht_id, delay=50)
        time.sleep(0.5)
        pw_inputs[0].click()
        page.keyboard.press("Control+a")
        page.keyboard.type(ht_pw, delay=50)
        time.sleep(0.5)
        filled_id2 = id_input.input_value()
        filled_pw2 = pw_inputs[0].input_value()
        print(f"    [Login] 재시도 후 ID: '{filled_id2}', PW 길이: {len(filled_pw2)}")

    # 로그인 버튼
    login_btns = [b for b in page.locator(
        "xpath=//button[normalize-space()='로그인'] | //input[@type='button' and @value='로그인'] | //a[normalize-space()='로그인']"
    ).all() if b.is_visible()]
    print(f"    [Login] 로그인 버튼 후보 {len(login_btns)}")
    if not login_btns:
        return False
    login_btns[-1].click()
    time.sleep(2)

    # 로그인 후 결과 polling: 2차 인증 모달 / 실패 알림 / 성공 (최대 12초)
    has_2fa = False
    has_failure = False
    for i in range(24):
        time.sleep(0.5)
        state = page.evaluate("""
            () => {
                const all = document.body.innerText || '';
                return {
                    twofa: all.includes('아이디 로그인 2차 인증') || all.includes('주민번호 앞 7자리'),
                    failure: all.includes('로그인 정보가 없습니다') || all.includes('일치하지 않'),
                };
            }
        """)
        if state["twofa"]:
            has_2fa = True
            print(f"    [Login] 2차 인증 모달 감지됨 ({(i+1)*0.5}초)")
            break
        if state["failure"]:
            has_failure = True
            print(f"    [Login] 로그인 실패 알림 감지됨")
            break

    if has_failure:
        return False

    if has_2fa:
        # 2차 인증 - "주민등록번호" 라벨 다음에 나오는 input 2개 직접 찾기
        modal_inputs_info = page.evaluate("""
            () => {
                const inputs = document.querySelectorAll('input');
                const out = [];
                for (const el of inputs) {
                    if (el.offsetParent !== null) {
                        out.push({
                            type: el.type,
                            ph: el.placeholder || '',
                            name: el.name || '',
                            value: el.value || '',
                            id: el.id || '',
                        });
                    }
                }
                return out;
            }
        """)
        print(f"    [Login] 모달 떴을 때 보이는 input들:")
        for inp in modal_inputs_info:
            print(f"        {inp}")

        # name 속성으로 직접 주민번호 input 찾기
        j1 = page.locator("input[name='iptUserJuminNo1']")
        j2 = page.locator("input[name='iptUserJuminNo2']")

        j1_visible = j1.is_visible()
        j2_visible = j2.is_visible()
        print(f"    [Login] iptUserJuminNo1 visible={j1_visible}, iptUserJuminNo2 visible={j2_visible}")

        if j1_visible and j2_visible:
            front6  = jumin7[:6]
            seventh = jumin7[6:] if len(jumin7) >= 7 else ""
            print(f"    [Login] 주민번호 입력 시도: front6={front6!r}, seventh={seventh!r}")

            # click(3번) 선택 후 press_sequentially 입력 (WebSquare 호환)
            j1.click(click_count=3)
            j1.press_sequentially(front6, delay=50)
            time.sleep(0.3)

            j2.click(click_count=3)
            j2.press_sequentially(seventh, delay=50)
            time.sleep(0.3)

            # 실제 들어갔는지 확인
            v1 = j1.input_value()
            v2 = j2.input_value()
            print(f"    [Login] 입력 확인: j1={v1!r}, j2={v2!r}")

            if v1 != front6 or v2 != seventh:
                # JS fallback
                print(f"    [Login] fill 실패 → JS fallback 시도")
                page.evaluate(f"""
                    () => {{
                        const j1 = document.querySelector("input[name='iptUserJuminNo1']");
                        const j2 = document.querySelector("input[name='iptUserJuminNo2']");
                        if (j1) {{ j1.value = {front6!r}; j1.dispatchEvent(new Event('input', {{bubbles:true}})); j1.dispatchEvent(new Event('change', {{bubbles:true}})); }}
                        if (j2) {{ j2.value = {seventh!r}; j2.dispatchEvent(new Event('input', {{bubbles:true}})); j2.dispatchEvent(new Event('change', {{bubbles:true}})); }}
                    }}
                """)
                time.sleep(0.3)
                v1 = j1.input_value()
                v2 = j2.input_value()
                print(f"    [Login] JS 후 확인: j1={v1!r}, j2={v2!r}")

            time.sleep(0.3)
            click_visible_text(page, "확인")
            time.sleep(4)
        else:
            print(f"    [Login] 2차인증 입력 필드 없음 (visible 아님)")
            return False

    print(f"    [Login] 최종 URL: {page.url}")
    dismiss_popups(page, rounds=3)
    return True


# -------------------- 안내문 PDF --------------------

def save_anneam_pdf(ctx, page, save_path):
    """미리보기 → ClipReport PDF 저장 (Track A 함수와 동일)"""
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


# -------------------- 한 고객 처리 --------------------

def process_customer(ctx, page, customer):
    name = customer["name"]
    print(f"\n=== [{name}] ===")

    # NAS 고객 폴더 (성명_주민앞6 패턴)
    folder_candidates = list(PDF_BASE.glob(f"{name}_*"))
    if folder_candidates:
        folder = folder_candidates[0]
    else:
        folder = PDF_BASE / name
    folder.mkdir(parents=True, exist_ok=True)

    # 1. 강제 로그아웃 (쿠키 클리어 포함)
    force_logout(ctx, page)

    # 2. 로그인 (검증은 신고도움 페이지에서)
    if not login_via_id_pw(page, customer["id"], customer["pw"], customer["jumin7"]):
        return {"name": name, "status": "에러", "msg": "로그인 단계 실패"}

    # 3. 신고도움 서비스 페이지로 (딥링크)
    page.goto(REPORT_HELP_URL_PERSONAL, wait_until="domcontentloaded")
    time.sleep(5)
    dismiss_popups(page, rounds=2)

    # 4. 페이지 진단 - 로그인 페이지로 리다이렉트되었으면 실패
    page_url = page.url
    is_login_redirect = "index_login" in page_url or "login.do" in page_url
    if is_login_redirect:
        return {"name": name, "status": "에러", "msg": f"로그인 페이지 리다이렉트: {page_url}"}

    # 5. 조회 클릭
    click_visible_text(page, "조회")
    time.sleep(3)

    # 6. PDF 저장
    pdf_path = folder / f"종소세안내문_{name}_TrackB.pdf"
    try:
        save_anneam_pdf(ctx, page, pdf_path)
        return {"name": name, "status": "완료", "msg": str(pdf_path)}
    except Exception as e:
        return {"name": name, "status": "에러", "msg": f"PDF 저장 실패: {str(e)[:120]}"}


# -------------------- 메인 --------------------

def read_customers_with_credentials(xlsx_path=None):
    """구글시트 접수명단에서 홈택스 ID/PW 가진 고객만 추출"""
    import sys
    sys.path.insert(0, r"F:\종소세2026")
    from gsheet_writer import get_credentials
    import gspread
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key("1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI")
    ws = sh.worksheet("접수명단")
    all_rows = ws.get_all_records()

    out = []
    for row in all_rows:
        name  = str(row.get("성명", "") or "").strip()
        jumin = str(row.get("주민번호", "") or "").strip()
        ht_id = str(row.get("홈택스아이디", "") or "").strip()
        ht_pw = str(row.get("홈택스비번", "") or "").strip()
        if not (name and jumin and ht_id and ht_pw):
            continue
        out.append({
            "name": name,
            "jumin7": normalize_jumin7(jumin),
            "id": ht_id,
            "pw": ht_pw,
        })
    print(f"[Track B 대상] {len(out)}건 로드")
    return out


def get_no_pdf_names():
    """NAS에서 PDF 없는 고객 이름 목록"""
    no_pdf = set()
    for folder in PDF_BASE.iterdir():
        if not folder.is_dir():
            continue
        pdfs = list(folder.glob("종소세안내문_*.pdf")) + list(folder.glob("자료/종소세안내문_*.pdf"))
        if not pdfs:
            no_pdf.add(folder.name.split("_")[0])
    return no_pdf


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", type=int, default=1, help="시작 번호 (1-based)")
    parser.add_argument("--all", action="store_true", help="PDF없는 사람 필터 없이 전체 ID/PW보유자 처리")
    args = parser.parse_args()

    all_customers = read_customers_with_credentials(INPUT_XLSX)

    if not args.all:
        no_pdf = get_no_pdf_names()
        customers = [c for c in all_customers if c["name"] in no_pdf]
        print(f"[Track B] PDF없음 필터 후: {len(customers)}명 (전체 {len(all_customers)}명)")
    else:
        customers = all_customers

    customers = customers[args.start - 1:]
    print(f"[Track B 대상] {args.start}번부터 {len(customers)}명: {[c['name'] for c in customers]}")

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()
        page.on("dialog", lambda d: d.dismiss())

        results = []
        for customer in customers:
            try:
                r = process_customer(ctx, page, customer)
            except Exception as e:
                r = {"name": customer["name"], "status": "예외", "msg": str(e)[:200]}
            results.append(r)
            print(f"    → {r['status']} | {r['msg']}")

    print("\n" + "=" * 60)
    print("[Track B 결과 요약]")
    for r in results:
        print(f"  {r['status']:6s} {r['name']:6s} {r['msg'][:80]}")


if __name__ == "__main__":
    main()
