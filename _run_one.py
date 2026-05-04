"""
_run_one.py - 신규고객 1명: PDF 다운로드 + 파싱 (incometaxbot 서브프로세스용)

사용: python _run_one.py {이름} {홈택스ID} {홈택스PW} {주민번호}

전제: Edge 디버그 창 열려있어야 함 (python launch_edge.py 또는 launch_edge.bat)
"""
import sys, os
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from datetime import datetime
from playwright.sync_api import sync_playwright

import 신규고객처리 as m
import parse_and_sync_신규 as pm
from 종합소득세안내문조회 import ensure_output_workbook
from config import OUTPUT_DIR

if len(sys.argv) < 5:
    print("사용법: python _run_one.py {이름} {홈택스ID} {홈택스PW} {주민번호}")
    sys.exit(1)

name       = sys.argv[1]
hometax_id = sys.argv[2]
hometax_pw = sys.argv[3]
jumin_raw  = sys.argv[4]

customer = {
    "name":       name,
    "hometax_id": hometax_id,
    "hometax_pw": hometax_pw,
    "jumin_raw":  jumin_raw,
}

print(f"[_run_one] {name} 처리 시작 (ID: {hometax_id})", flush=True)

wb, ws_out = ensure_output_workbook()

def _ensure_edge_cdp():
    """Edge CDP 확인 후 없으면 자동 실행"""
    import urllib.request, subprocess, time
    from pathlib import Path
    try:
        urllib.request.urlopen("http://localhost:9222/json", timeout=2)
        return  # 이미 떠있음
    except Exception:
        pass
    print("[Edge] CDP 미실행 → 자동 실행 중...", flush=True)
    EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    USER_DATA = Path(r"F:\종소세2026\.edge_debug_profile")
    USER_DATA.mkdir(parents=True, exist_ok=True)
    subprocess.Popen([EDGE, "--remote-debugging-port=9222",
                      f"--user-data-dir={USER_DATA}", "https://hometax.go.kr"])
    for _ in range(15):
        time.sleep(1)
        try:
            urllib.request.urlopen("http://localhost:9222/json", timeout=2)
            print("[Edge] CDP 준비 완료", flush=True)
            return
        except Exception:
            pass
    print("[오류] Edge CDP 시작 실패", flush=True)
    sys.exit(1)

_ensure_edge_cdp()

with sync_playwright() as p:
    try:
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
    except Exception as e:
        print(f"[오류] Edge CDP 연결 실패: {e}")
        sys.exit(1)

    ctx  = browser.contexts[0]
    page = ctx.pages[0]
    page.bring_to_front()
    page.on("dialog", lambda d: d.dismiss())

    result = m.process_one_신규(ctx, page, customer)
    print(f"[다운로드] {result['status']}: {result.get('error_msg', '')}", flush=True)

    ws_out.append([
        name, jumin_raw, result["status"],
        result.get("anneam_pdf", ""),
        result.get("prev_income_xlsx", ""),
        result.get("biznos", ""),
        result.get("vat_xlsx_count", 0),
        result.get("error_msg", ""),
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    ])
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(str(OUTPUT_DIR / "결과.xlsx"))

if result["status"] in ("완료", "부분완료"):
    print(f"[파싱] {name} 파싱 시작...", flush=True)
    pm.NEW_NAMES = [name]
    pm.main()
    print(f"[파싱] {name} 완료", flush=True)
else:
    print(f"[파싱 스킵] 다운로드 실패로 파싱 건너뜀: {result.get('error_msg', '')}")
    sys.exit(1)
