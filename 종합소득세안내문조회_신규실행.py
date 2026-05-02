import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from datetime import datetime
from playwright.sync_api import sync_playwright
from 종합소득세안내문조회 import process_one, ensure_output_workbook

customers = [
    {"name": "신정숙", "jumin_raw": "6906302917114", "phone_raw": "1039406940"},
]

wb, ws = ensure_output_workbook()

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    ctx = browser.contexts[0]
    page = ctx.pages[0]
    page.bring_to_front()
    page.on("dialog", lambda d: d.dismiss())

    for i, c in enumerate(customers, 1):
        print(f"[{i}/1] {c['name']}")
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

print("[완료] 신정숙 처리 완료")
