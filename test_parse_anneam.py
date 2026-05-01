"""
종소세 안내문 PDF 파싱 테스트
- 4개 PDF에서 핵심 필드 추출 시도
- 결과를 표로 출력 → 어떤 필드가 안정적으로 뽑히는지 판단
"""
import pdfplumber
import re
from pathlib import Path
import warnings
import logging

warnings.filterwarnings("ignore")
logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfplumber").setLevel(logging.ERROR)

BASE = Path(r"F:\종소세2026\output\PDF")

BIZNO = re.compile(r"\d{3}-\d{2}-\d{5}")
NUM = re.compile(r"-?[\d,]+(?:\.\d+)?")  # 숫자(콤마/소수점/음수)


def first_match(pattern, text, default=""):
    m = re.search(pattern, text)
    return m.group(1).strip() if m else default


def parse_anneam(pdf_path):
    """안내문 PDF 1개 → 핵심 필드 dict"""
    out = {
        "성명": "",
        "생년월일": "",
        "기장의무": "",
        "추계시적용경비율": "",
        "수입금액총계": "",
        "이자": "",
        "배당": "",
        "근로(단일)": "",
        "근로(복수)": "",
        "연금": "",
        "기타": "",
    }

    with pdfplumber.open(pdf_path) as pdf:
        all_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    # 페이지1: 기본정보
    out["성명"] = first_match(r"성명\s+(\S+)\s+생년월일", all_text)
    out["생년월일"] = first_match(r"생년월일\s+([\d.]+)", all_text)
    out["기장의무"] = first_match(r"기장의무\s+([^\n]+?)\s+추계시", all_text)
    out["추계시적용경비율"] = first_match(r"추계시적용경비율\s+(\S+)", all_text)
    out["수입금액총계"] = first_match(r"총계\s+([\d,]+)", all_text).replace(",", "")

    # 타소득(합산대상): 이자/배당/근로단일/근로복수/연금/기타 각각 별도 컬럼 (O/X)
    m = re.search(
        r"해당여부\s+([XO])\s+([XO])\s+([XO])\s+([XO])\s+([XO])\s+([XO])",
        all_text,
    )
    if m:
        labels = ["이자", "배당", "근로(단일)", "근로(복수)", "연금", "기타"]
        for i, label in enumerate(labels):
            out[label] = m.group(i + 1)

    return out


def main():
    rows = []
    for folder in sorted(BASE.iterdir()):
        if not folder.is_dir() or folder.name == "테스트":
            continue
        pdfs = list(folder.glob("종소세안내문_*.pdf"))
        if not pdfs:
            continue
        for pdf in pdfs:
            print(f"\n=== {pdf.name} ===")
            try:
                data = parse_anneam(pdf)
                rows.append(data)
                for k, v in data.items():
                    print(f"  {k:20s} : {v}")
            except Exception as e:
                print(f"  파싱 실패: {e}")

    print("\n" + "=" * 80)
    print(f"총 {len(rows)}건 파싱 완료")


if __name__ == "__main__":
    main()
