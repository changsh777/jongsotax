"""
Track B 에러 8명 재처리 - 2차인증 상세 로그 포함
"""
import os, sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from step5_individual import (
    login_via_id_pw, force_logout, dismiss_popups,
    save_anneam_pdf, click_visible_text, REPORT_HELP_URL_PERSONAL,
    PDF_BASE, normalize_jumin7
)
from playwright.sync_api import sync_playwright
from gsheet_writer import get_credentials
import gspread

ERROR_NAMES = {'마금현', '지성호', '김혜린', '김진곤', '정도민', '이윤경'}

def read_error_customers():
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key('1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI')
    ws = sh.worksheet('접수명단')
    rows = ws.get_all_records()

    out = []
    for r in rows:
        name  = str(r.get('성명','') or '').strip()
        jumin = str(r.get('주민번호','') or '').strip()
        ht_id = str(r.get('홈택스아이디','') or '').strip()
        ht_pw = str(r.get('홈택스비번','') or '').strip()
        if name not in ERROR_NAMES or not ht_id:
            continue
        jumin7 = normalize_jumin7(jumin)
        no7 = jumin7[6:] if len(jumin7) >= 7 else "(none)"
        print(f"  {name}: raw={jumin!r} j7={jumin7!r} front6={jumin7[:6]!r} 7th={no7!r}")
        out.append({'name': name, 'jumin_raw': jumin, 'jumin7': jumin7, 'id': ht_id, 'pw': ht_pw})
    return out


def process_one(ctx, page, c):
    name   = c['name']
    jumin7 = c['jumin7']
    print(f"\n{'='*50}")
    print(f"[{name}] ID={c['id']} | jumin7={jumin7!r}")
    front6 = jumin7[:6]
    seventh = jumin7[6:] if len(jumin7) >= 7 else "(MISSING - only 6 digits!)"
    print(f"  front6={front6!r}  seventh={seventh!r}")

    force_logout(ctx, page)

    ok = login_via_id_pw(page, c['id'], c['pw'], jumin7)
    if not ok:
        print(f"  ❌ 로그인 실패")
        return

    # 로그인 성공 확인
    page.wait_for_timeout(2000)
    logged_in = page.evaluate("""
        () => document.body.innerText.includes('로그아웃')
    """)
    print(f"  {'✅ 로그인 확인됨' if logged_in else '⚠️ 로그아웃 버튼 안 보임 - 로그인 불확실'}")

    # 신고도움서비스 이동
    page.goto(REPORT_HELP_URL_PERSONAL, wait_until='domcontentloaded')
    time.sleep(5)
    dismiss_popups(page, rounds=2)

    # 페이지 상태 진단
    page_text = page.evaluate("() => document.body.innerText")
    has_login_page = '아이디 로그인' in page_text
    has_josa = '종합소득세' in page_text or '신고도움' in page_text
    has_mibori = '미리보기' in page_text
    has_josa_btn = '조회' in page_text

    print(f"  페이지진단: 로그인페이지={has_login_page} | 신고도움={has_josa} | 조회버튼={has_josa_btn} | 미리보기={has_mibori}")

    if has_login_page:
        print(f"  ❌ 로그인 페이지로 리다이렉트 → 2차인증 실패한 것")
        return

    if not has_josa_btn:
        print(f"  ❌ 조회 버튼 없음")
        return

    click_visible_text(page, '조회')
    time.sleep(4)

    page_text2 = page.evaluate("() => document.body.innerText")
    has_mibori2 = '미리보기' in page_text2
    has_result  = '총수입금액' in page_text2 or '수입금액' in page_text2
    print(f"  조회후: 미리보기={has_mibori2} | 수입금액보임={has_result}")

    if not has_mibori2:
        print(f"  ❌ 미리보기 없음 → 사업소득 없는 분 (신고도움 데이터 없음)")
        return

    # PDF 저장
    folder_candidates = list(PDF_BASE.glob(f"{name}_*"))
    folder = folder_candidates[0] if folder_candidates else PDF_BASE / name
    folder.mkdir(parents=True, exist_ok=True)
    pdf_path = folder / f"종소세안내문_{name}_TrackB.pdf"
    try:
        save_anneam_pdf(ctx, page, pdf_path)
        print(f"  ✅ PDF 저장: {pdf_path}")
    except Exception as e:
        print(f"  ❌ PDF 저장 실패: {e}")


def main():
    print("[에러 8명 재처리]\n")
    customers = read_error_customers()
    print(f"\n총 {len(customers)}명 처리 시작\n")

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        ctx = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()
        page.on("dialog", lambda d: d.dismiss())

        for c in customers:
            try:
                process_one(ctx, page, c)
            except Exception as e:
                print(f"  예외: {e}")


if __name__ == "__main__":
    main()
