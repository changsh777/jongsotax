"""
PDF 안내문에서 사업자번호 추출 테스트
- 어제 받은 3개 PDF로 검증
- 정규식 패턴: 000-00-00000
"""
import pdfplumber
import re
from pathlib import Path

PDF_DIR = Path(r"F:\종소세2026\output\PDF")

BIZNO_PATTERN = re.compile(r"\d{3}-\d{2}-\d{5}")


def extract_biznos(pdf_path):
    """PDF 전체 텍스트에서 사업자번호 패턴 추출 (중복 제거)"""
    biznos = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            text = page.extract_text() or ""
            found = BIZNO_PATTERN.findall(text)
            if found:
                print(f"  [페이지 {i}] {found}")
            biznos.extend(found)
    # 순서 유지 중복 제거
    seen = set()
    unique = []
    for b in biznos:
        if b not in seen:
            seen.add(b)
            unique.append(b)
    return unique


def main():
    pdfs = sorted(PDF_DIR.glob("*.pdf"))
    if not pdfs:
        print("PDF가 output/PDF 직하에 없습니다. 어제 받은 파일 위치 확인 필요.")
        return

    for pdf in pdfs:
        print(f"\n=== {pdf.name} ===")
        biznos = extract_biznos(pdf)
        print(f"  → 추출된 사업자번호 {len(biznos)}개: {biznos}")


if __name__ == "__main__":
    main()
