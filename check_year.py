import sys, os, random, re
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import pdfplumber, openpyxl
from config import PARSE_RESULT_XLSX
from pathlib import Path

wb = openpyxl.load_workbook(PARSE_RESULT_XLSX, data_only=True)
ws = wb.active
headers = [c.value for c in ws[1]]
idx = {h: i for i, h in enumerate(headers)}

rows = [r for r in ws.iter_rows(min_row=2, values_only=True)
        if r and r[0] and r[idx["PDF경로"]]]

random.seed(7)
samples = random.sample(rows, min(8, len(rows)))

print("=== PDF 귀속연도 확인 ===")
for row in samples:
    name = str(row[0]).strip()
    pdf_path = Path(str(row[idx["PDF경로"]]))
    if not pdf_path.exists():
        print(f"{name}: PDF 없음")
        continue
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""
    years = re.findall(r"(20\d\d)\s*년?\s*(귀속|신고)", text)
    fname = pdf_path.name
    print(f"{name} ({fname}): {years[:4]}")
