import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

import unicodedata
import subprocess
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import process_one, ensure_output_workbook
from jakupan_gen import make_jakupan
from config import CUSTOMER_DIR

VERIFY_SCRIPT = Path(r"F:\종소세2026\verify_folder_integrity.py")


def jakupan_exists(name: str) -> bool:
    nfc = unicodedata.normalize("NFC", name)
    for folder in CUSTOMER_DIR.iterdir():
        if not folder.is_dir():
            continue
        if not unicodedata.normalize("NFC", folder.name).startswith(nfc + "_"):
            continue
        for f in folder.iterdir():
            if f.is_file() and unicodedata.normalize("NFC", f.name).startswith("작업판_"):
                return True
    return False


customers = [
    {"name": "김인구", "jumin_raw": "8602151823521", "phone_raw": "1036780215"},
]

wb, ws = ensure_output_workbook()

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    ctx = browser.contexts[0]
    page = ctx.pages[0]
    page.bring_to_front()
    page.on("dialog", lambda d: d.dismiss())

    for i, c in enumerate(customers, 1):
        name = c["name"]
        print(f"[{i}/{len(customers)}] {name}")
        r = process_one(ctx, page, c)
        ws.append([
            name, str(c["jumin_raw"]), r["status"],
            r["anneam_pdf"], r["prev_income_xlsx"],
            r["biznos"], r["vat_xlsx_count"],
            r["error_msg"],
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ])
        wb.save(r"F:\종소세2026\output\결과.xlsx")
        print(f"    -> {r['status']} {r['error_msg'] or ''}")

        # 파싱 성공/부분완료 시 작업판 자동 생성
        if r["status"] in ("완료", "부분완료"):
            jumin6 = str(c["jumin_raw"]).replace("-", "").replace(" ", "")[:6]
            if jakupan_exists(name):
                print(f"    [작업판] 이미 있음 — 스킵")
            else:
                print(f"    [작업판] 생성 중...")
                out = make_jakupan(name, jumin6)
                if out:
                    print(f"    [작업판] 저장 완료 → {Path(out).name}")
                else:
                    print(f"    [작업판] 생성 실패 (수동 확인 필요)")
        else:
            print(f"    [작업판] 파싱 에러 — 스킵")
        print()

names = ", ".join(c["name"] for c in customers)
print(f"[완료] {names} 처리 완료")

# 배치 완료 후 혼입검증 자동 실행
print("\n[혼입검증] 시작...")
result = subprocess.run(
    [sys.executable, str(VERIFY_SCRIPT)],
    capture_output=True, text=True, encoding="utf-8", errors="replace"
)
print(result.stdout)
if result.returncode != 0:
    print(f"[혼입검증] 오류:\n{result.stderr}")
