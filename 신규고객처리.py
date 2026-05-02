"""
신규고객처리.py - 수임동의 미완료(신규) 고객 종합소득세 안내문 조회 (건바이건)

전제조건:
  1. launch_edge.py 실행 → Edge 디버그 창 열기
  2. 홈택스 로그아웃 상태 (스크립트가 고객 아이디/비번으로 직접 로그인)
  3. python 신규고객처리.py

처리대상: 구글시트 접수명단 중 고객구분=신규 + PDF 없는 고객 자동 감지
처리방식: 고객 아이디/비번 자동 로그인 → 안내문 다운 → 자동 로그아웃 반복
주의: 비밀번호 오류 시 계정잠금 위험 → 구글시트 홈택스아이디/비번 확인 필수
완료 후: PDF 파싱 + 구글시트 자동 동기화
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

import time
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from 종합소득세안내문조회 import (
    ensure_output_workbook,
    save_anneam_pdf, download_prev_income_tax,
    download_vat, extract_biznos,
)
from gsheet_writer import get_credentials
from config import CUSTOMER_DIR, customer_folder
import gspread

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
HOMETAX_MAIN_URL = (
    "https://www.hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
)
# 개인납세자(신규)용 종합소득세 신고도움서비스 URL
# 아이디/비번 로그인 후 접근 가능 — 로그인한 본인 데이터 자동 표시
INDIVIDUAL_REPORT_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=41&tm2lIdx=4103000000&tm3lIdx=4103080000"
)
# 개인납세자 신고도움서비스 미리보기 버튼 ID
INDIVIDUAL_PREVIEW_BTN_ID = "#mf_txppWframe_trigger1"


def load_신규_customers():
    """구글시트 접수명단에서 고객구분=신규 + PDF 없는 고객만 반환"""
    creds = get_credentials()
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet("접수명단")
    rows = ws.get_all_records()

    targets = []
    for r in rows:
        name       = str(r.get("성명", "") or "").strip()
        구분        = str(r.get("고객구분", "") or "").strip()
        jumin      = str(r.get("주민번호", "") or "").strip()
        phone      = str(r.get("핸드폰번호", "") or "").strip()
        hometax_id = str(r.get("홈택스아이디", "") or "").strip()
        hometax_pw = str(r.get("홈택스비번", "") or "").strip()

        if not name or 구분 != "신규":
            continue

        # PDF 존재 여부 확인
        folder_candidates = list(CUSTOMER_DIR.glob(f"{name}_*")) + [CUSTOMER_DIR / name]
        folder = next((f for f in folder_candidates if f.is_dir()), None)
        has_pdf = folder and bool(list(folder.glob("종소세안내문_*.pdf")))

        if not has_pdf:
            if not hometax_id or not hometax_pw:
                print(f"  [스킵] {name}: 홈택스아이디/비번 없음 → 구글시트 입력 필요")
                continue
            targets.append({
                "name":       name,
                "jumin_raw":  jumin,
                "phone_raw":  phone,
                "hometax_id": hometax_id,
                "hometax_pw": hometax_pw,
            })

    return targets


# -------------------- 로그인/로그아웃 --------------------

def logout_hometax(page):
    """홈택스 로그아웃 — 로그아웃 버튼 → HTML 모달 확인 클릭 → 완료"""
    # 메인 페이지로 이동 (이미 메인이면 reload)
    try:
        page.goto(HOMETAX_MAIN_URL, wait_until="domcontentloaded")
        time.sleep(2)
    except Exception:
        pass

    # 로그인 상태인지 확인
    try:
        txt = page.evaluate("() => document.body.innerText")
        if "로그아웃" not in txt:
            return  # 이미 로그아웃 상태
    except Exception:
        return

    # 로그아웃 버튼 클릭
    try:
        logout_btn = page.locator("#mf_wfHeader_group1503")
        if logout_btn.is_visible(timeout=3000):
            logout_btn.click()
            time.sleep(1)
    except Exception:
        pass

    # HTML 모달 확인 버튼 클릭 (로그아웃 확인 팝업)
    try:
        page.locator("[id*=_btn_confirm]").click(timeout=3000)
        time.sleep(3)  # 로그아웃 후 리다이렉트 대기
        print("    [로그아웃] 완료", flush=True)
    except Exception:
        time.sleep(2)


def login_hometax_id(page, hometax_id: str, hometax_pw: str, jumin_raw: str = "") -> bool:
    """홈택스 아이디/비번 로그인 (2차 인증 포함). 반환: True(성공) / False(실패)

    흐름:
      1. 홈택스 메인 이동 → '아이디 로그인' 아이콘 클릭
      2. ID/PW 입력 → 로그인 클릭
      3. 2차 인증 팝업(주민번호) 처리 → 확인
      4. 로그인 성공 확인
    """
    # 로그아웃 후 리다이렉트가 끝날 때까지 대기 후 goto
    time.sleep(2)
    try:
        page.goto(HOMETAX_MAIN_URL, wait_until="domcontentloaded")
    except Exception:
        time.sleep(3)
        page.goto(HOMETAX_MAIN_URL, wait_until="domcontentloaded")
    time.sleep(2)

    # '아이디 로그인' 아이콘 클릭
    try:
        tab = page.get_by_text("아이디 로그인", exact=True).first
        if tab.is_visible(timeout=3000):
            tab.click()
            time.sleep(2)
    except Exception:
        pass

    # ID 입력
    id_el = page.locator("#mf_txppWframe_loginboxFrame_iptUserId")
    if not id_el.is_visible(timeout=5000):
        print("    [로그인] ID 입력란 못 찾음", flush=True)
        return False
    id_el.fill(hometax_id)

    # PW 입력
    pw_el = page.locator("#mf_txppWframe_loginboxFrame_iptUserPw")
    if not pw_el.is_visible(timeout=3000):
        print("    [로그인] PW 입력란 못 찾음", flush=True)
        return False
    pw_el.fill(hometax_pw)
    time.sleep(0.3)

    # 로그인 버튼 클릭
    login_btn = page.locator("#mf_txppWframe_loginboxFrame_wq_uuid_923")
    if not login_btn.is_visible(timeout=3000):
        login_btn = page.get_by_text("로그인", exact=True).first
    if not login_btn.is_visible(timeout=3000):
        print("    [로그인] 로그인 버튼 못 찾음", flush=True)
        return False
    login_btn.click()
    time.sleep(3)

    # 2차 인증 팝업 처리 (아이디 로그인 시 주민번호 확인 요구)
    jumin_field = page.locator("[id*=UTXPPABC12][id*=iptUserJuminNo1]")
    if jumin_field.is_visible(timeout=3000):
        print("    [2차인증] 주민번호 입력 중...", flush=True)
        try:
            s = str(jumin_raw).replace("-", "").replace(" ", "").strip()
            if len(s) == 13:
                front, back = s[:6], s[6:]
            else:
                raise ValueError(f"주민번호 형식 이상: {jumin_raw}")
            jumin_field.fill(front)
            page.locator("[id*=UTXPPABC12][id*=iptUserJuminNo2]").fill(back)
            time.sleep(0.3)
            # 확인 버튼 클릭
            page.locator("[id*=UTXPPABC12][id*=trigger46]").click()
            time.sleep(4)
            print("    [2차인증] 완료", flush=True)
        except Exception as e:
            print(f"    [2차인증 실패] {e}", flush=True)
            return False

    # 혹시 뜨는 정보 알림 팝업 닫기
    try:
        info_btn = page.locator("[id*=_btn_confirm]")
        if info_btn.is_visible(timeout=2000):
            info_btn.click()
            time.sleep(1)
    except Exception:
        pass

    # 성공 여부 확인 — 사용자 이름 또는 메뉴 항목 존재로 확인
    try:
        # 신고도움서비스 메뉴(menuAtag_4103080000)가 있고 로그인 박스가 없으면 성공
        menu_exists = page.locator("#menuAtag_4103080000").count() > 0
        login_box_visible = page.locator("#mf_txppWframe_loginboxFrame_iptUserId").is_visible(timeout=500)
        if menu_exists and not login_box_visible:
            print(f"    [로그인 성공] {hometax_id}", flush=True)
            return True
    except Exception:
        pass

    # 에러 메시지 출력
    try:
        err = page.evaluate("""
            () => {
                const selectors = '.w2modal_body, .w2alert, .err_msg, [class*=alert]';
                for (const el of document.querySelectorAll(selectors)) {
                    const t = (el.innerText||'').trim();
                    if (t && t.length > 3) return t;
                }
                return '';
            }
        """)
        if err:
            print(f"    [로그인 실패] {err[:120]}", flush=True)
    except Exception:
        pass

    return False


# -------------------- 신규 고객 1명 처리 --------------------

def process_one_신규(ctx, page, customer: dict) -> dict:
    name       = customer["name"]
    hometax_id = customer["hometax_id"]
    hometax_pw = customer["hometax_pw"]
    folder     = customer_folder(name, customer.get("jumin_raw", ""))

    result = {
        "status": "에러", "error_msg": "",
        "anneam_pdf": "", "prev_income_xlsx": "",
        "vat_xlsx_count": 0, "biznos": "",
    }

    logout_hometax(page)

    if not login_hometax_id(page, hometax_id, hometax_pw, jumin_raw=customer.get("jumin_raw", "")):
        result["error_msg"] = f"로그인 실패 (ID:{hometax_id}) - 구글시트 아이디/비번/주민번호 확인"
        logout_hometax(page)
        return result

    try:
        # 개인납세자 신고도움서비스 — SPA 내부 메뉴 이동 (goto 직접 접근 시 세션 오류 발생)
        print(f"    [신고도움서비스] 메뉴 이동 중...", flush=True)
        page.evaluate("document.getElementById('menuAtag_4103080000').onclick()")
        time.sleep(4)

        # 접근 불가 체크 (로그인 정보 없음 팝업)
        body_txt = page.evaluate("() => document.body.innerText")
        if "로그인 정보가 없습니다" in body_txt:
            result["error_msg"] = "신고도움서비스 접근불가 (세션 오류 - 로그인 재확인 필요)"
            return result

        # 미리보기 버튼 (개인납세자 페이지 고정 ID)
        preview_btn = page.locator(INDIVIDUAL_PREVIEW_BTN_ID)
        if not preview_btn.is_visible(timeout=8000):
            # 버튼이 안 보이면 귀속년도 2025 선택 후 재시도
            print(f"    [미리보기] 버튼 대기 중...", flush=True)
            time.sleep(3)
            if not preview_btn.is_visible(timeout=5000):
                result["error_msg"] = "안내문 없음 (미리보기 버튼 미표시) - 신고안내 미생성 고객"
                return result

        anneam_path = folder / f"종소세안내문_{name}.pdf"
        if not anneam_path.exists():
            print(f"    [PDF 다운로드] {name} 안내문 저장 중...", flush=True)
            save_anneam_pdf(ctx, page, preview_btn, anneam_path)
        else:
            print(f"    [PDF 스킵] 기존 파일 존재", flush=True)
        result["anneam_pdf"] = str(anneam_path)
        time.sleep(1)

        try:
            prev_path = folder / "전년도종소세신고내역.xlsx"
            ok = download_prev_income_tax(page, prev_path)
            result["prev_income_xlsx"] = str(prev_path) if ok else "자료없음"
        except Exception as e:
            result["error_msg"] += f" [전년도종소세:{e}]"

        if anneam_path.exists():
            biznos = extract_biznos(anneam_path)
            result["biznos"] = ",".join(biznos)
            cnt = 0
            for biz in biznos:
                try:
                    ok = download_vat(page, biz, folder / f"부가세신고내역_{biz}.xlsx")
                    if ok: cnt += 1
                except Exception as e:
                    result["error_msg"] += f" [부가세{biz}:{e}]"
            result["vat_xlsx_count"] = cnt

        result["status"] = "완료" if not result["error_msg"] else "부분완료"

    except Exception as e:
        result["error_msg"] = f"{type(e).__name__}: {str(e)[:200]}"

    finally:
        logout_hometax(page)

    return result


# -------------------- 메인 --------------------

def main():
    print("[신규고객처리] 구글시트에서 처리 대상 조회 중...")
    customers = load_신규_customers()

    if not customers:
        print("  → 처리할 고객 없음 (신규 고객 PDF 모두 완료 또는 아이디/비번 미입력)")
        return

    print(f"  → {len(customers)}명 처리 대상: {[c['name'] for c in customers]}")
    print(f"\n  ※ Edge 디버그 창 열려있는지 확인\n")

    wb, ws_out = ensure_output_workbook()

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx  = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()
        page.on("dialog", lambda d: d.dismiss())

        for i, c in enumerate(customers, 1):
            print(f"[{i}/{len(customers)}] {c['name']} (ID: {c['hometax_id']})")
            r = process_one_신규(ctx, page, c)
            ws_out.append([
                c["name"], str(c["jumin_raw"]), r["status"],
                r["anneam_pdf"], r["prev_income_xlsx"],
                r["biznos"], r["vat_xlsx_count"],
                r["error_msg"],
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ])
            wb.save(r"F:\종소세2026\output\결과.xlsx")
            print(f"    → {r['status']} {r['error_msg'] or ''}\n")

    print(f"[완료] {len(customers)}명 처리")

    names = [c["name"] for c in customers]
    print(f"\n[파싱 시작] {names}")
    try:
        import parse_and_sync_신규 as pm
        pm.NEW_NAMES = names
        pm.main()
    except Exception as e:
        print(f"[파싱 실패] {e}")
        print("  수동: 안내문파싱_신규동기화.py 에서 NEW_NAMES 수정 후 실행")


if __name__ == "__main__":
    main()
