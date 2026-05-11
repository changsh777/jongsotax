import sys, io, os, re, unicodedata
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")
import pdfplumber
from pathlib import Path

base = Path(r"Z:\종소세2026\고객")
targets = ["최인경_600205", "박창환_710122", "김병수_840826"]
keywords = ["이자소득금액", "배당소득금액", "근로소득금액", "연금소득금액", "기타소득금액"]

for folder_name in targets:
    folder = base / folder_name
    pdfs = sorted([f for f in folder.iterdir()
                   if "신고서" in unicodedata.normalize("NFC", f.name) and f.suffix == ".pdf"])
    if not pdfs:
        print(f"{folder_name}: 신고서없음\n")
        continue
    pdf = pdfs[-1]
    print(f"=== {folder_name} ({pdf.name}) ===")
    with pdfplumber.open(str(pdf)) as p:
        pages = [unicodedata.normalize("NFC", pg.extract_text() or "") for pg in p.pages]
    print(f"  총 {len(pages)}페이지")
    for i, text in enumerate(pages):
        for line in text.split("\n"):
            if any(k in line for k in keywords):
                print(f"  p{i+1}: {repr(line.strip())}")
    print()
