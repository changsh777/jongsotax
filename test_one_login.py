import sys, time
sys.path.insert(0, r"F:\종소세2026")
from playwright.sync_api import sync_playwright
from 신규고객처리 import logout_hometax, login_hometax_id, INDIVIDUAL_PREVIEW_BTN_ID
from 종합소득세안내문조회 import save_anneam_pdf
from config import customer_folder

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    ctx = browser.contexts[0]
    page = ctx.pages[0]
    page.bring_to_front()
    page.on("dialog", lambda d: d.dismiss())

    name       = "신정숙"
    hometax_id = "tlswjdtnr69"
    hometax_pw = "ch5470015!"
    jumin_raw  = "690630-2059723"  # 구글시트 주민번호

    print("[1] 로그아웃", flush=True)
    logout_hometax(page)

    print("[2] 로그인 (2차인증 포함)", flush=True)
    ok = login_hometax_id(page, hometax_id, hometax_pw, jumin_raw=jumin_raw)
    print("결과:", ok, flush=True)
    if not ok:
        page.screenshot(path=r"F:\종소세2026\output\FAIL_login.png")
        exit(1)
    time.sleep(1)

    print("[3] 신고도움서비스 이동", flush=True)
    page.evaluate("document.getElementById('menuAtag_4103080000').onclick()")
    time.sleep(4)
    page.screenshot(path=r"F:\종소세2026\output\SUCCESS_report.png")
    print("URL:", page.url[:120], flush=True)

    txt = page.evaluate("() => document.body.innerText")
    if "로그인 정보가 없습니다" in txt:
        print("접근 실패!", flush=True)
        exit(1)

    print("[4] 미리보기 → PDF 저장", flush=True)
    preview_btn = page.locator(INDIVIDUAL_PREVIEW_BTN_ID)
    print("미리보기 visible:", preview_btn.is_visible(timeout=5000), flush=True)
    folder = customer_folder(name, jumin_raw)
    anneam_path = folder / f"종소세안내문_{name}.pdf"
    save_anneam_pdf(ctx, page, preview_btn, anneam_path)
    print("완료! PDF:", anneam_path, flush=True)
