import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from datetime import datetime
from playwright.sync_api import sync_playwright
from step4_full import process_one, ensure_output_workbook

customers = [
    {"name": "탁설환", "jumin_raw": "8012131520311", "phone_raw": "1028628979"},
    {"name": "유영주", "jumin_raw": "7906052528318", "phone_raw": "1092328979"},
]

wb, ws = ensure_output_workbook()

with sync_playwright() as p:
    browser = p.chromium.connect_over_cdp("http://localhost:9222")
    ctx = browser.contexts[0]
    page = ctx.pages[0]
    page.bring_to_front()
    page.on("dialog", lambda d: d.dismiss())

    for i, c in enumerate(customers, 1):
        print(f"[{i}/2] {c['name']}")
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

print("[완료] 2명 처리 완료")
