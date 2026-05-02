import sys, os, random, re
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

import xlrd, openpyxl
from config import CUSTOMER_DIR
from pathlib import Path

random.seed(3)

# 전년도종소세신고내역.xlsx 파일 수집
files = list(CUSTOMER_DIR.glob("*/전년도종소세신고내역.xlsx"))
print(f"전년도 xlsx 파일 {len(files)}개 발견")

samples = random.sample(files, min(8, len(files)))

print("\n=== 전년도 종소세 신고내역 연도 확인 ===")
for f in samples:
    folder = f.parent.name
    try:
        wb = xlrd.open_workbook(f)
        sh = wb.sheet_by_index(0)
        # 1행(헤더) 확인
        headers = sh.row_values(0)
        years = [str(int(v)) for v in headers if isinstance(v, float) and 2000 < v < 2030]
        # 또는 텍스트에서 연도 추출
        header_str = ' '.join(str(h) for h in headers[:10])
        yr = re.findall(r"20\d\d", header_str)
        # 데이터 1행
        data_row = sh.row_values(1) if sh.nrows > 1 else []
        print(f"{folder}: 헤더연도={yr} | 첫데이터={[str(v)[:10] for v in data_row[:5]]}")
    except Exception:
        try:
            wb2 = openpyxl.load_workbook(f, data_only=True)
            sh2 = wb2.active
            headers = [str(c.value or '') for c in sh2[1]]
            yr = re.findall(r"20\d\d", ' '.join(headers))
            data = [str(sh2.cell(2, i+1).value or '') for i in range(5)]
            print(f"{folder}: 헤더연도={yr} | 첫데이터={data}")
        except Exception as e:
            print(f"{folder}: 읽기 실패 {e}")
