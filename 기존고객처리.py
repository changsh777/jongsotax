"""
기존고객처리.py - 수임동의 완료(기존) 고객 종합소득세 안내문 조회 (건바이건)

전제조건:
  1. launch_edge.py 실행 → Edge 디버그 창 열기
  2. 홈택스에 세무사 계정으로 직접 로그인
  3. python 기존고객처리.py

처리대상: 구글시트 접수명단 중 고객구분=기존 + PDF 없는 고객 자동 감지
처리방식: 세무사 계정 세션 유지 → 고객 주민번호 입력 → 안내문 다운
완료 후: PDF 파싱 + 구글시트 자동 동기화
"""
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from datetime import datetime
from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import process_one, ensure_output_workbook
from gsheet_writer import get_credentials
from config import CUSTOMER_DIR
import gspread

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"


def load_기존_customers(only_names: list = None):
    """구글시트 접수명단에서 고객구분=기존 + PDF 없는 고객만 반환
    only_names 지정 시 해당 이름만 처리 (미처리 전체 방지)
    """
    creds = get_credentials()
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet("접수명단")
    rows = ws.get_all_records()

    targets = []
    for r in rows:
        name  = str(r.get("성명", "") or "").strip()
        구분   = str(r.get("고객구분", "") or "").strip()
        jumin = str(r.get("주민번호", "") or "").strip()
        phone = str(r.get("핸드폰번호", "") or "").strip()

        if not name or 구분 != "기존":
            continue

        # 이름 필터 (지정 시 해당 이름만)
        if only_names and name not in only_names:
            continue

        # PDF 존재 여부 확인
        folder_candidates = list(CUSTOMER_DIR.glob(f"{name}_*")) + [CUSTOMER_DIR / name]
        folder = next((f for f in folder_candidates if f.is_dir()), None)
        has_pdf = folder and bool(list(folder.glob("종소세안내문_*.pdf")))

        if not has_pdf:
            targets.append({
                "name":      name,
                "jumin_raw": jumin,
                "phone_raw": phone,
            })

    return targets


def main():
    # 커맨드라인 인수로 이름 지정 시 해당 고객만 처리
    only_names = sys.argv[1:] if len(sys.argv) > 1 else None
    if only_names:
        print(f"[기존고객처리] 지정 고객만 처리: {only_names}")
    else:
        print("[기존고객처리] 구글시트에서 처리 대상 조회 중...")
    customers = load_기존_customers(only_names=only_names)

    if not customers:
        print("  → 처리할 고객 없음 (기존 고객 PDF 모두 완료)")
        return

    print(f"  → {len(customers)}명 처리 대상: {[c['name'] for c in customers]}")
    print(f"\n  ※ Edge 디버그 창 + 세무사 계정 홈택스 로그인 확인 후 계속\n")

    wb, ws_out = ensure_output_workbook()

    pw = sync_playwright().start()
    try:
        browser = pw.chromium.connect_over_cdp("http://localhost:9222")
        ctx  = browser.contexts[0]
        page = ctx.pages[0]
        page.bring_to_front()
        page.on("dialog", lambda d: d.dismiss())

        for i, c in enumerate(customers, 1):
            print(f"[{i}/{len(customers)}] {c['name']}")
            r = process_one(ctx, page, c)
            ws_out.append([
                c["name"], str(c["jumin_raw"]), r["status"],
                r["anneam_pdf"], r["prev_income_xlsx"],
                r["biznos"], r["vat_xlsx_count"],
                r["error_msg"],
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ])
            wb.save(r"F:\종소세2026\output\결과.xlsx")
            print(f"    → {r['status']} {r['error_msg'] or ''}\n")
    finally:
        try:
            browser.disconnect()
        except Exception:
            pass
        try:
            pw.stop()
        except Exception:
            pass

    print(f"[완료] {len(customers)}명 처리")

    # 파싱 + 구글시트 동기화
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
