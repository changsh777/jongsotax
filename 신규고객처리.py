"""
신규고객처리.py - 동의자 + 신규 통합 처리 (건바이건)

사용법:
  1. 아래 NEW_CUSTOMERS 리스트만 수정
  2. Edge 열고 세무사 계정으로 홈택스 로그인
  3. python 신규고객처리.py

동의자: 세무사 계정 세션에서 주민번호 조회 (기존 방식)
신규  : 고객 아이디/비번으로 자동 로그인 → 안내문 다운 → 로그아웃 → 세무사 재로그인 안내
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

import time
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from 종합소득세안내문조회 import (
    process_one, ensure_output_workbook,
    fill_jumin_and_search, wait_preview_button,
    save_anneam_pdf, download_prev_income_tax,
    download_vat, extract_biznos, normalize_jumin,
    REPORT_HELP_URL,
)
from config import customer_folder

# ============================================================
# ★ 여기만 수정 ★
# login: "동의" = 세무사 계정으로 주민번호 조회
#        "신규" = 고객 아이디/비번으로 직접 로그인
# ============================================================
NEW_CUSTOMERS = [
    {"name": "신정숙",  "jumin_raw": "6906302917114", "phone_raw": "1039406940", "login": "동의"},
    # {"name": "한효성",  "jumin_raw": "XXXXXXXXXXXXX", "phone_raw": "10XXXXXXXX",  "login": "신규",
    #  "hometax_id": "아이디", "hometax_pw": "비밀번호"},
]
# ============================================================

HOMETAX_LOGIN_URL = (
    "https://www.hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
)
HOMETAX_LOGOUT_URL = (
    "https://www.hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml&tmIdx=00"
)


# -------------------- 로그인/로그아웃 --------------------

def logout_hometax(page):
    """홈택스 로그아웃"""
    try:
        # 로그아웃 버튼 클릭 시도
        logout_btn = page.locator("xpath=//*[contains(text(),'로그아웃') and @onclick]").first
        if logout_btn.is_visible(timeout=3000):
            logout_btn.click()
            time.sleep(2)
            print("    [로그아웃] 버튼 클릭 완료", flush=True)
            return
    except Exception:
        pass
    # fallback: 로그아웃 URL 직접 이동
    try:
        page.evaluate("if(typeof jf_logout === 'function') jf_logout(); else if(typeof logout === 'function') logout();")
        time.sleep(2)
    except Exception:
        pass
    page.goto(HOMETAX_LOGIN_URL, wait_until="domcontentloaded")
    time.sleep(2)
    print("    [로그아웃] 페이지 이동 완료", flush=True)


def login_hometax_id(page, hometax_id: str, hometax_pw: str) -> bool:
    """홈택스 아이디/비번 로그인
    반환: True(성공) / False(실패)
    """
    page.goto(HOMETAX_LOGIN_URL, wait_until="domcontentloaded")
    time.sleep(2)

    # 아이디 로그인 탭 클릭
    try:
        id_tab = page.locator("xpath=//*[contains(text(),'아이디') and (contains(@class,'tab') or contains(@onclick,''))]").first
        if id_tab.is_visible(timeout=3000):
            id_tab.click()
            time.sleep(1)
    except Exception:
        pass

    # ID 입력
    id_inputs = page.locator("xpath=//input[@id='txbUserId' or @name='userId' or @placeholder='아이디']").all()
    id_inputs = [el for el in id_inputs if el.is_visible()]
    if not id_inputs:
        print("    [로그인] ID 입력란 못 찾음", flush=True)
        return False
    id_inputs[0].fill(hometax_id)
    time.sleep(0.3)

    # PW 입력
    pw_inputs = page.locator("xpath=//input[@id='txbUserPw' or @name='userPw' or @type='password']").all()
    pw_inputs = [el for el in pw_inputs if el.is_visible()]
    if not pw_inputs:
        print("    [로그인] PW 입력란 못 찾음", flush=True)
        return False
    pw_inputs[0].fill(hometax_pw)
    time.sleep(0.3)

    # 로그인 버튼
    login_btns = page.locator(
        "xpath=//input[@value='로그인' or @value='확인']"
        " | //button[contains(text(),'로그인')]"
        " | //a[contains(text(),'로그인')]"
    ).all()
    login_btns = [el for el in login_btns if el.is_visible()]
    if not login_btns:
        print("    [로그인] 로그인 버튼 못 찾음", flush=True)
        return False
    login_btns[0].click()
    time.sleep(3)

    # 로그인 성공 여부 확인 (로그아웃 버튼 or 사용자명 표시)
    try:
        success = page.evaluate("""
            () => {
                const t = document.body.innerText || '';
                return t.includes('로그아웃') || t.includes('마이페이지');
            }
        """)
        if success:
            print(f"    [로그인 성공] {hometax_id}", flush=True)
            return True
    except Exception:
        pass

    # 에러 메시지 확인
    try:
        err_msg = page.evaluate("""
            () => {
                for (const el of document.querySelectorAll('.err, .error, [class*=alert]')) {
                    const t = (el.innerText || '').trim();
                    if (t) return t;
                }
                return '';
            }
        """)
        if err_msg:
            print(f"    [로그인 실패] {err_msg[:100]}", flush=True)
    except Exception:
        pass

    print(f"    [로그인] 결과 불명 - 페이지 수동 확인 필요", flush=True)
    return False


# -------------------- 신규 고객 1명 처리 --------------------

def process_one_신규(ctx, page, customer):
    """신규 고객: 고객 아이디/비번 로그인 → 안내문 다운 → 로그아웃"""
    name       = customer["name"]
    hometax_id = customer.get("hometax_id", "")
    hometax_pw = customer.get("hometax_pw", "")
    folder     = customer_folder(name, customer.get("jumin_raw", ""))

    result = {
        "status":           "에러",
        "error_msg":        "",
        "anneam_pdf":       "",
        "prev_income_xlsx": "",
        "vat_xlsx_count":   0,
        "biznos":           "",
    }

    if not hometax_id or not hometax_pw:
        result["error_msg"] = "hometax_id / hometax_pw 없음"
        return result

    # 1) 기존 세션 로그아웃
    print(f"    [신규] 로그아웃 중...", flush=True)
    logout_hometax(page)

    # 2) 고객 아이디/비번 로그인
    print(f"    [신규] {hometax_id} 로그인 시도...", flush=True)
    ok = login_hometax_id(page, hometax_id, hometax_pw)
    if not ok:
        result["error_msg"] = f"홈택스 로그인 실패 (ID:{hometax_id})"
        # 로그아웃 시도 후 반환 (세무사 재로그인 위해)
        logout_hometax(page)
        return result

    try:
        # 3) 신고도움서비스 이동
        page.goto(REPORT_HELP_URL, wait_until="domcontentloaded")
        time.sleep(3)

        # 신규 고객으로 로그인 시 주민번호 조회 필요 여부 확인
        # → 미리보기 버튼이 바로 보이면 자기 것 직접 표시
        # → 주민번호 입력 필드가 있으면 입력 후 조회
        preview_btn = wait_preview_button(page, timeout_ms=5000)

        if preview_btn is None:
            # 주민번호 입력 시도
            try:
                front, back = normalize_jumin(customer["jumin_raw"])
                fill_jumin_and_search(page, front, back)
                time.sleep(3)
                preview_btn = wait_preview_button(page, timeout_ms=10000)
            except Exception as e:
                result["error_msg"] = f"주민번호 조회 실패: {e}"
                logout_hometax(page)
                return result

        if preview_btn is None:
            result["error_msg"] = "미리보기 버튼 없음 (안내문 없음)"
            logout_hometax(page)
            return result

        # 4) 안내문 PDF 다운
        anneam_path = folder / f"종소세안내문_{name}.pdf"
        if anneam_path.exists():
            print(f"    [PDF 스킵] 기존 파일 존재", flush=True)
        else:
            try:
                save_anneam_pdf(ctx, page, preview_btn, anneam_path)
            except Exception as e:
                result["error_msg"] += f" [PDF실패:{e}]"
        result["anneam_pdf"] = str(anneam_path)
        time.sleep(1)

        # 5) 전년도 종소세
        try:
            prev_path = folder / "전년도종소세신고내역.xlsx"
            ok2 = download_prev_income_tax(page, prev_path)
            result["prev_income_xlsx"] = str(prev_path) if ok2 else "자료없음"
        except Exception as e:
            result["error_msg"] += f" [전년도종소세실패:{e}]"

        # 6) 사업자번호 추출 + 부가세
        if anneam_path.exists():
            biznos = extract_biznos(anneam_path)
            result["biznos"] = ",".join(biznos)
            vat_count = 0
            for bizno in biznos:
                try:
                    vat_path = folder / f"부가세신고내역_{bizno}.xlsx"
                    ok3 = download_vat(page, bizno, vat_path)
                    if ok3:
                        vat_count += 1
                except Exception as e:
                    result["error_msg"] += f" [부가세{bizno}실패:{e}]"
            result["vat_xlsx_count"] = vat_count

        result["status"] = "완료" if not result["error_msg"] else "부분완료"

    except Exception as e:
        result["error_msg"] = f"{type(e).__name__}: {str(e)[:200]}"

    finally:
        # 반드시 로그아웃 (다음 고객/세무사 재로그인 위해)
        print(f"    [신규] 로그아웃 처리 중...", flush=True)
        logout_hometax(page)
        print(f"    [신규] ★ 세무사 계정으로 다시 로그인 해주세요 (동의자 남아있는 경우)", flush=True)

    return result


# -------------------- 메인 --------------------

def main():
    동의자 = [c for c in NEW_CUSTOMERS if c.get("login") == "동의"]
    신규   = [c for c in NEW_CUSTOMERS if c.get("login") == "신규"]
    total  = len(NEW_CUSTOMERS)

    print(f"[신규고객처리] 총 {total}명 (동의자 {len(동의자)}명 / 신규 {len(신규)}명)")
    print(f"  처리 순서: 동의자 먼저 → 신규 (로그인/로그아웃 포함)\n")

    wb, ws = ensure_output_workbook()

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx     = browser.contexts[0]
        page    = ctx.pages[0]
        page.bring_to_front()
        page.on("dialog", lambda d: d.dismiss())

        # ── 동의자 처리 (세무사 계정 세션) ──
        for i, c in enumerate(동의자, 1):
            print(f"[동의자 {i}/{len(동의자)}] {c['name']}")
            r = process_one(ctx, page, c)
            ws.append([
                c["name"], str(c["jumin_raw"]), r["status"],
                r["anneam_pdf"], r["prev_income_xlsx"],
                r["biznos"], r["vat_xlsx_count"],
                r["error_msg"],
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ])
            wb.save(r"F:\종소세2026\output\결과.xlsx")
            print(f"    → {r['status']} {r['error_msg'] or ''}\n")

        # ── 신규 처리 (고객 아이디/비번 로그인) ──
        if 신규:
            print(f"\n{'='*50}")
            print(f"[신규 고객 처리 시작] {len(신규)}명")
            print(f"  각 고객마다 로그인/로그아웃 자동 수행")
            print(f"{'='*50}\n")

        for i, c in enumerate(신규, 1):
            print(f"[신규 {i}/{len(신규)}] {c['name']} (ID: {c.get('hometax_id','?')})")
            r = process_one_신규(ctx, page, c)
            ws.append([
                c["name"], str(c["jumin_raw"]), r["status"],
                r["anneam_pdf"], r["prev_income_xlsx"],
                r["biznos"], r["vat_xlsx_count"],
                r["error_msg"],
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ])
            wb.save(r"F:\종소세2026\output\결과.xlsx")
            print(f"    → {r['status']} {r['error_msg'] or ''}\n")

    print(f"\n[전체 완료] {total}명 처리")

    # 파싱 + 구글시트 동기화 (신규만)
    신규_names = [c["name"] for c in NEW_CUSTOMERS]
    print(f"\n[파싱 시작] {신규_names}")
    try:
        from parse_and_sync_신규 import main as parse_main
        import parse_and_sync_신규 as pm
        pm.NEW_NAMES = 신규_names
        parse_main()
    except Exception as pe:
        print(f"[파싱 실패] {pe}")
        print("  수동: python 안내문파싱_신규동기화.py")


if __name__ == "__main__":
    main()
