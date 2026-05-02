"""
신규고객처리.py - 수임동의 미완료(신규) 고객 종합소득세 안내문 조회 (건바이건)

전제조건:
  1. launch_edge.py 실행 → Edge 디버그 창 열기
  2. 홈택스 로그아웃 상태 (스크립트가 고객 아이디/비번으로 직접 로그인)
  3. python 신규고객처리.py

처리방식: 고객 아이디/비번 자동 로그인 → 안내문 다운 → 자동 로그아웃 반복
주의: 비밀번호 오류 시 계정잠금 위험 → 사전에 아이디/비번 확인 필수
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
    download_vat, extract_biznos, normalize_jumin,
    wait_preview_button, fill_jumin_and_search,
    REPORT_HELP_URL,
)
from config import customer_folder

# ============================================================
# ★ 여기만 수정 ★  (수임동의 미완료 신규 고객만)
# ============================================================
CUSTOMERS = [
    # {"name": "한효성", "jumin_raw": "XXXXXXXXXXXXX", "phone_raw": "10XXXXXXXX",
    #  "hometax_id": "아이디", "hometax_pw": "비밀번호"},
]
# ============================================================

HOMETAX_MAIN_URL = (
    "https://www.hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
)


def logout_hometax(page):
    """홈택스 로그아웃"""
    try:
        page.evaluate("if(typeof jf_logout==='function') jf_logout();")
        time.sleep(2)
        return
    except Exception:
        pass
    page.goto(HOMETAX_MAIN_URL, wait_until="domcontentloaded")
    time.sleep(2)


def login_hometax_id(page, hometax_id: str, hometax_pw: str) -> bool:
    """홈택스 아이디/비번 로그인. 반환: True(성공) / False(실패)"""
    page.goto(HOMETAX_MAIN_URL, wait_until="domcontentloaded")
    time.sleep(2)

    # 아이디 탭 클릭
    try:
        tab = page.locator("xpath=//*[contains(text(),'아이디') and contains(@class,'tab')]").first
        if tab.is_visible(timeout=2000):
            tab.click()
            time.sleep(0.5)
    except Exception:
        pass

    # ID 입력
    id_el = page.locator("xpath=//input[@id='txbUserId' or @name='userId' or @placeholder='아이디']").first
    if not id_el.is_visible(timeout=3000):
        print("    [로그인] ID 입력란 못 찾음", flush=True)
        return False
    id_el.fill(hometax_id)

    # PW 입력
    pw_el = page.locator("xpath=//input[@id='txbUserPw' or @name='userPw' or @type='password']").first
    if not pw_el.is_visible(timeout=3000):
        print("    [로그인] PW 입력란 못 찾음", flush=True)
        return False
    pw_el.fill(hometax_pw)

    # 로그인 버튼
    btn = page.locator(
        "xpath=//input[@value='로그인'] | //button[contains(text(),'로그인')]"
    ).first
    if not btn.is_visible(timeout=3000):
        print("    [로그인] 로그인 버튼 못 찾음", flush=True)
        return False
    btn.click()
    time.sleep(4)

    # 성공 여부 확인
    try:
        ok = page.evaluate("() => document.body.innerText.includes('로그아웃')")
        if ok:
            print(f"    [로그인 성공] {hometax_id}", flush=True)
            return True
    except Exception:
        pass

    # 에러 메시지 출력
    try:
        err = page.evaluate("""
            () => {
                for (const el of document.querySelectorAll('.err,.error,[class*=alert]')) {
                    const t = (el.innerText||'').trim();
                    if (t) return t;
                }
                return '';
            }
        """)
        if err:
            print(f"    [로그인 실패] {err[:120]}", flush=True)
    except Exception:
        pass

    return False


def process_one_신규(ctx, page, customer: dict) -> dict:
    """신규(미동의) 고객 1명: 로그인 → 안내문 다운 → 로그아웃"""
    name       = customer["name"]
    hometax_id = customer.get("hometax_id", "")
    hometax_pw = customer.get("hometax_pw", "")
    folder     = customer_folder(name, customer.get("jumin_raw", ""))

    result = {
        "status": "에러", "error_msg": "",
        "anneam_pdf": "", "prev_income_xlsx": "",
        "vat_xlsx_count": 0, "biznos": "",
    }

    if not hometax_id or not hometax_pw:
        result["error_msg"] = "hometax_id / hometax_pw 없음 → CUSTOMERS 리스트 확인"
        return result

    # 기존 세션 로그아웃
    logout_hometax(page)

    # 고객 계정 로그인
    if not login_hometax_id(page, hometax_id, hometax_pw):
        result["error_msg"] = f"로그인 실패 (ID:{hometax_id}) - 아이디/비번 확인 필요"
        logout_hometax(page)
        return result

    try:
        page.goto(REPORT_HELP_URL, wait_until="domcontentloaded")
        time.sleep(3)

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
                return result

        if preview_btn is None:
            result["error_msg"] = "안내문 없음 (미리보기 버튼 미표시)"
            return result

        # 안내문 PDF
        anneam_path = folder / f"종소세안내문_{name}.pdf"
        if not anneam_path.exists():
            save_anneam_pdf(ctx, page, preview_btn, anneam_path)
        else:
            print(f"    [PDF 스킵] 기존 파일 존재", flush=True)
        result["anneam_pdf"] = str(anneam_path)
        time.sleep(1)

        # 전년도 종소세
        try:
            prev_path = folder / "전년도종소세신고내역.xlsx"
            ok = download_prev_income_tax(page, prev_path)
            result["prev_income_xlsx"] = str(prev_path) if ok else "자료없음"
        except Exception as e:
            result["error_msg"] += f" [전년도종소세:{e}]"

        # 부가세
        if anneam_path.exists():
            biznos = extract_biznos(anneam_path)
            result["biznos"] = ",".join(biznos)
            cnt = 0
            for biz in biznos:
                try:
                    ok = download_vat(page, biz, folder / f"부가세신고내역_{biz}.xlsx")
                    if ok:
                        cnt += 1
                except Exception as e:
                    result["error_msg"] += f" [부가세{biz}:{e}]"
            result["vat_xlsx_count"] = cnt

        result["status"] = "완료" if not result["error_msg"] else "부분완료"

    except Exception as e:
        result["error_msg"] = f"{type(e).__name__}: {str(e)[:200]}"

    finally:
        logout_hometax(page)

    return result


def main():
    total = len(CUSTOMERS)
    if total == 0:
        print("[신규고객처리] CUSTOMERS 리스트가 비어있어요. 고객 정보를 입력해주세요.")
        return

    print(f"[신규고객처리] {total}명 처리 시작")
    print(f"  ※ 각 고객마다 자동 로그인/로그아웃 수행\n")

    wb, ws = ensure_output_workbook()

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx  = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()
        page.on("dialog", lambda d: d.dismiss())

        for i, c in enumerate(CUSTOMERS, 1):
            print(f"[{i}/{total}] {c['name']} (ID: {c.get('hometax_id','?')})")
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

    print(f"[완료] {total}명 처리")

    # 파싱 + 구글시트 동기화
    names = [c["name"] for c in CUSTOMERS]
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
