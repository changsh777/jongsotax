"""
tax_cross_verify.py  —  종소세 교차검증 보고서 생성기
세무회계창연 | 2026

사용법:
  python tax_cross_verify.py 박현민_870529        (폴더명)
  python tax_cross_verify.py 박현민 870529        (이름 + 주민앞6)

결과:
  고객폴더/검증보고서_YYYYMMDD_HHMM.html
"""

from __future__ import annotations
import sys, io, os, re, unicodedata, fnmatch
from pathlib import Path
from datetime import datetime


def nfc_glob(folder, pattern: str):
    """Mac SMB NFD 파일명 대응 glob"""
    nfc_pat = unicodedata.normalize("NFC", pattern)
    return [p for p in folder.iterdir()
            if fnmatch.fnmatch(unicodedata.normalize("NFC", p.name), nfc_pat)]

# stdout utf-8 강제 (한글 터미널 오류 방지)
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
elif sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

_THIS_DIR = Path(__file__).resolve().parent
if str(_THIS_DIR) not in sys.path:
    sys.path.insert(0, str(_THIS_DIR))
os.environ.setdefault("SEOTAX_ENV", "nas")

import pdfplumber
import xlrd

# ── 경로 (플랫폼 자동 선택) ──────────────────────────────────────
def _default_customer_dir() -> Path:
    candidates = [
        Path(r"Z:\종소세2026\고객"),                        # Windows NAS
        Path("/Users/changmini/NAS/종소세2026/고객"),       # Mac Mini
        Path("/mnt/nas/종소세2026/고객"),                   # Linux
    ]
    for c in candidates:
        if c.exists():
            return c
    return Path(r"Z:\종소세2026\고객")   # fallback

CUSTOMER_DIR = _default_customer_dir()

# ── 파일 역할 키워드 ──────────────────────────────────────────────
ROLES = {
    "신고서":      ["신고서", "과세표준확정신고"],
    "안내문":      ["안내문", "신고안내정보", "종소세안내문"],
    "영수증":      ["원천징수영수증"],
    "지급명세서":  ["사업소득"],     # xlsx
    "신한카드":    ["신한카드"],
    "비씨카드":    ["비씨카드", "bc카드", "salelst"],
}


# ═══════════════════════════════════════════════════════════════════
# 1. 파일 인식
# ═══════════════════════════════════════════════════════════════════

def classify_files(folder: Path) -> dict:
    """폴더 내 파일 자동 분류 → {역할: [Path, ...]}"""
    result = {k: [] for k in ROLES}
    result["기타"] = []

    def _match(fname: str) -> str | None:
        import unicodedata
        fl = unicodedata.normalize("NFC", fname.lower())
        for role, kws in ROLES.items():
            if any(unicodedata.normalize("NFC", kw.lower()) in fl for kw in kws):
                return role
        return None

    # 루트 파일
    for f in folder.iterdir():
        if not f.is_file():
            continue
        role = _match(f.name)
        if role:
            result[role].append(f)
        else:
            result["기타"].append(f)

    # 하위 폴더 (지급명세서/, 간이용역소득/, 자료/)
    # _archive/, _temp/ 등 밑줄로 시작하는 폴더는 스캔 제외
    for sub in folder.iterdir():
        if not sub.is_dir():
            continue
        if sub.name.startswith("_"):   # _archive, _temp 등 제외
            continue
        for f in sub.rglob("*"):
            if not f.is_file():
                continue
            role = _match(f.name)
            if role:
                result[role].append(f)
            elif f.suffix.lower() in (".pdf", ".xlsx", ".xls"):
                result["기타"].append(f)

    # 신고서: PDF만
    result["신고서"] = [f for f in result["신고서"] if f.suffix.lower() == ".pdf"]
    # 연도 순 정렬 (파일명 내 연도 숫자 기준)
    def _year(p):
        m = re.search(r'(\d{4})', p.stem)
        return int(m.group(1)) if m else 0
    result["신고서"].sort(key=_year)

    return result


# ═══════════════════════════════════════════════════════════════════
# 2. 파싱 함수
# ═══════════════════════════════════════════════════════════════════

def parse_tax_return(pdf_path: Path) -> dict:
    """신고서 PDF → 주요 수치 dict
    keys: 귀속연도, 업종코드, 총수입금액, 필요경비, 소득금액,
          소득공제, 과세표준, 세율, 산출세액, 세액공제,
          결정세액, 기납부세액, 납부세액
    """
    result = {"파일": pdf_path.name}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            p1 = pdf.pages[0].extract_text() or ""
            p2 = pdf.pages[1].extract_text() if len(pdf.pages) > 1 else ""
    except Exception as e:
        result["오류"] = str(e)
        return result

    def _int(text, pattern):
        m = re.search(pattern, text, re.DOTALL)
        if not m:
            return None
        raw = m.group(1).replace(",", "").replace(" ", "").strip()
        try:
            v = int(float(raw))
            return -v if raw.startswith("-") and v > 0 else v
        except Exception:
            return None

    # 귀속연도 — PDF에 따라 글자 사이 공백 삽입 케이스 대응
    p1_nsp = re.sub(r'\s', '', p1)   # 공백 제거본으로 검색
    m_yr = re.search(r'\((\d{4})년귀속\)', p1_nsp)
    result["귀속연도"] = int(m_yr.group(1)) if m_yr else None

    # 기장의무·신고유형
    m_gij = re.search(r'기\s*장\s*의\s*무.*?(간편장부대상자|복식부기의무자)', p1)
    result["기장의무"] = m_gij.group(1) if m_gij else ""

    # 업종코드 (p2) — 기호와 한글 사이 공백 대응
    m_up = re.search(r'⑧\s*주\s*업\s*종\s*코\s*드\s*([\d]{6})', p2)
    result["업종코드"] = m_up.group(1) if m_up else ""

    def _parse_multi_col(raw: str) -> int | None:
        """복식부기 다중사업장 숫자 파싱
        - 마지막 숫자 == 앞 숫자들의 합 → 합계열 있음 → 마지막 반환
        - 그 외(합계열 없음 또는 단일 사업장) → 모두 합산
        """
        nums_raw = re.findall(r'[\d,]+', raw)
        if not nums_raw:
            return None
        nums_int = [int(n.replace(",", "")) for n in nums_raw]
        if len(nums_int) >= 2 and nums_int[-1] == sum(nums_int[:-1]):
            return nums_int[-1]   # 합계열 있음
        return sum(nums_int)      # 합계열 없음 → 전체 합산

    # 총수입금액 (p2 ⑨)
    m_rev = re.search(r'⑨\s*총\s*수\s*입\s*금\s*액([^\n]+)', p2)
    result["총수입금액"] = _parse_multi_col(m_rev.group(1)) if m_rev else None

    # 필요경비 (p2 ⑩)
    m_exp = re.search(r'⑩\s*필\s*요\s*경\s*비([^\n]+)', p2)
    result["필요경비"] = _parse_multi_col(m_exp.group(1)) if m_exp else None

    # 소득금액 ⑪
    m_inc = re.search(r'⑪\s*소\s*득\s*금\s*액([^\n]+)', p2)
    if m_inc:
        result["소득금액"] = _parse_multi_col(m_inc.group(1))
    else:
        m_inc2 = re.search(r'종\s+합\s+소\s+득\s+금\s+액\s+\d+\s+([\d,]+)', p1)
        result["소득금액"] = int(m_inc2.group(1).replace(",", "")) if m_inc2 else None

    # 소득공제 (p1 20)
    m_ded = re.search(r'소\s+득\s+공\s+제\s+20\s+([\d,]+)', p1)
    result["소득공제"] = int(m_ded.group(1).replace(",", "")) if m_ded else None

    # 과세표준 (p1 21)
    m_base = re.search(r'과\s+세\s+표\s+준.*?21\s+([\d,]+)', p1)
    result["과세표준"] = int(m_base.group(1).replace(",", "")) if m_base else None

    # 세율 (p1 22)
    m_rate = re.search(r'세\s+율\s+22\s+([\d.]+)', p1)
    result["세율"] = float(m_rate.group(1)) if m_rate else None

    # 산출세액 (p1 23)
    m_calc = re.search(r'산\s+출\s+세\s+액\s+23\s+([\d,]+)', p1)
    result["산출세액"] = int(m_calc.group(1).replace(",", "")) if m_calc else None

    # 세액공제 (p1 25)
    m_cr = re.search(r'세\s+액\s+공\s+제\s+25\s+([\d,]+)', p1)
    result["세액공제"] = int(m_cr.group(1).replace(",", "")) if m_cr else None

    # 결정세액 (합계 28) — 공백 삽입 PDF 포함 3단계 패턴
    m_det = re.search(r'합\s+계\s*\(.*?26.*?27.*?\)\s+28\s+([\d,]+)', p1)
    if not m_det:
        m_det = re.search(r'합\s+계\(26\s*\+27\s*\)\s+28\s+([\d,]+)', p1)
    if not m_det:  # 숫자 사이 공백 있는 케이스 (복식부기 PDF)
        m_det = re.search(r'합\s*계\s*\(\s*2\s*6\s*\+\s*2\s*7\s*\)\s*2\s*8\s*([\d,]+)', p1)
    result["결정세액"] = int(m_det.group(1).replace(",", "")) if m_det else None

    # 기납부세액 (32)
    m_pre = re.search(r'기\s+납\s+부\s+세\s+액\s+32\s+([\d,]+)', p1)
    result["기납부세액"] = int(m_pre.group(1).replace(",", "")) if m_pre else None

    # 납부(환급)세액 (33)
    m_pay = re.search(r'납\s+부.*?총\s+세\s+액.*?33\s+(-?[\d,]+)', p1)
    if m_pay:
        raw = m_pay.group(1).replace(",", "")
        result["납부세액"] = int(raw) if raw else None
    else:
        result["납부세액"] = None

    return result


def parse_anneam(pdf_path: Path) -> dict:
    """안내문 PDF → 수입금액, 기장의무, 추계경비율, 기납부세액, 공제참고자료"""
    result = {}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)
            text0 = pdf.pages[0].extract_text() or ""
    except Exception:
        return result

    # 수입금액
    m_total = re.search(r'총계\s+([\d,]+)', text0)
    if m_total:
        result["수입금액"] = int(m_total.group(1).replace(",", ""))

    # 기장의무
    m_gij = re.search(r'기장의무\s+(간편장부대상자|복식부기의무자)', text0)
    result["기장의무"] = m_gij.group(1) if m_gij else ""

    # 추계경비율
    m_choegye = re.search(r'추계시적용경비율\s+(기준경비율|단순경비율)', text0)
    result["추계경비율"] = m_choegye.group(1) if m_choegye else ""

    # 원천징수세액 (기납부세액) — 줄바꿈 포함 탐색
    m_pre = re.search(r'원천징수세액\D{0,20}?([\d,]{4,})원', text, re.DOTALL)
    if m_pre:
        result["기납부세액"] = int(m_pre.group(1).replace(",", ""))

    # ── 공제 참고자료 ─────────────────────────────────────────────
    def _amt(pattern):
        m = re.search(pattern, text, re.DOTALL)
        if m:
            v = m.group(1).replace(",", "")
            try: return int(v)
            except: return None
        return None

    # 소득공제
    result["국민연금보험료"]   = _amt(r'국민연금보험료\s*([\d,]+)원')
    result["개인연금저축"]     = _amt(r'개인연금저축\s*([\d,]+)원')
    # 노란우산공제: PDF에서 설명문(한도 표기) 다음 줄에 금액이 오는 구조
    # "소기업소상공인공제부금 이하),500만원...\n6,500,000원" 형태 대응
    result["노란우산공제"] = _amt(r'소기업소상공인공제부금[^\n]+\n\s*([\d,]+)원')
    if result["노란우산공제"] is None:          # fallback: 같은 줄에 금액이 있는 PDF
        result["노란우산공제"] = _amt(r'소기업소상공인공제부금[^\d]*([\d,]+)원')

    # 세액공제
    result["퇴직연금세액공제"] = _amt(r'퇴직연금세액공제\s*([\d,]+)원')
    result["연금계좌세액공제"] = _amt(r'연금계좌세[액금]?세액공제\s*([\d,]+)원')

    # 중간예납세액
    result["중간예납세액"]     = _amt(r'중간예납세액\s*([\d,]+)원')

    return result


def parse_income_excel(xlsx_path: Path, jumin6: str = "") -> list:
    """지급명세서 xlsx → list of dict {사업자번호, 징수의무자, 지급총액, 소득세, 지방소득세}"""
    import io as _io
    try:
        import msoffcrypto as _msoffcrypto
    except ImportError:
        _msoffcrypto = None

    rows = []
    try:
        with open(xlsx_path, "rb") as f:
            raw = f.read()

        is_enc = False
        if _msoffcrypto:
            try:
                of = _msoffcrypto.OfficeFile(_io.BytesIO(raw))
                is_enc = of.is_encrypted()
            except Exception:
                is_enc = False

        if is_enc and jumin6 and _msoffcrypto:
            of2 = _msoffcrypto.OfficeFile(_io.BytesIO(raw))
            dec = _io.BytesIO()
            of2.load_key(password=jumin6)
            of2.decrypt(dec)
            wb = xlrd.open_workbook(file_contents=dec.getvalue())
        else:
            wb = xlrd.open_workbook(str(xlsx_path))

        sheet = wb.sheet_by_index(0)
        if sheet.nrows < 2:
            return []
        hdrs = sheet.row_values(0)
        col = {h: i for i, h in enumerate(hdrs)}

        for i in range(1, sheet.nrows):
            vals = sheet.row_values(i)
            rows.append({
                "사업자번호": str(vals[col.get("사업자(주민)등록번호", 1)]).strip(),
                "징수의무자":  str(vals[col.get("징수의무자", 2)]).strip(),
                "지급총액":    int(float(vals[col.get("총지급액", 7)] or 0)),
                "소득세":      int(float(vals[col.get("소득세", 8)] or 0)),
                "지방소득세":  int(float(vals[col.get("지방소득세", 9)] or 0)),
            })
    except Exception as e:
        print(f"  [경고] 지급명세서 파싱 오류: {e}")
    return rows


def parse_card_excel(xls_path: Path) -> dict:
    """카드 xls → {합계, 건수, 카드사}"""
    result = {"합계": 0, "건수": 0, "카드사": ""}
    fname = xls_path.name.lower()

    # 카드사 구분
    if "신한" in fname:
        result["카드사"] = "신한카드"
    elif "bc" in fname or "비씨" in fname or "salelst" in fname:
        result["카드사"] = "비씨카드"

    try:
        wb = xlrd.open_workbook(str(xls_path))
        sheet = wb.sheet_by_index(0)
    except Exception:
        return result

    # 헤더 행 탐색 (금액 컬럼 찾기)
    amt_col = None
    header_row = 0
    for r in range(min(sheet.nrows, 10)):
        for c in range(sheet.ncols):
            v = str(sheet.cell_value(r, c)).strip()
            if "이용금액" in v or "매출금액" in v or "사용금액" in v or "거래금액" in v:
                amt_col = c
                header_row = r
                break
        if amt_col is not None:
            break

    if amt_col is None:
        return result

    total = 0
    count = 0
    for r in range(header_row + 1, sheet.nrows):
        v = sheet.cell_value(r, amt_col)
        # 숫자 또는 "1,450 원" 형식 문자열 처리
        amt = None
        if isinstance(v, (int, float)) and v > 0:
            amt = int(v)
        elif isinstance(v, str):
            clean = re.sub(r'[^\d]', '', v)
            if clean:
                try:
                    amt = int(clean)
                except Exception:
                    pass
        if amt and amt > 0:
            total += amt
            count += 1

    result["합계"] = total
    result["건수"] = count
    return result


# ═══════════════════════════════════════════════════════════════════
# 3. 교차검증 로직
# ═══════════════════════════════════════════════════════════════════

def cross_verify(
    당기신고서: dict,
    전기신고서: dict | None,
    안내문: dict,
    지급명세서: list,
    카드_목록: list,
) -> list:
    """
    검증 항목별 결과 list of dict:
    {섹션, 항목, 값1_label, 값1, 값2_label, 값2, 차이, 상태, 메모}
    상태: 'pass' | 'warn' | 'fail'
    """
    results = []

    _AUTO = object()   # sentinel: diff 자동계산
    # diff=None → 차이 칸 숨김, diff=_AUTO(기본) → v1-v2 자동계산, diff=숫자 → 명시값

    def add(섹션, 항목, lbl1, v1, lbl2, v2, 상태, 메모="", diff=_AUTO):
        if diff is None:
            computed = None          # 명시적 숨김
        elif diff is not _AUTO:
            computed = diff          # 명시값
        else:
            computed = (             # 자동계산
                (v1 or 0) - (v2 or 0)
                if isinstance(v1, (int, float)) and isinstance(v2, (int, float))
                else None
            )
        results.append({
            "섹션": 섹션,
            "항목": 항목,
            "값1_label": lbl1,
            "값1": v1,
            "값2_label": lbl2,
            "값2": v2,
            "차이": computed,
            "상태": 상태,
            "메모": 메모,
        })

    # ── A. 수입금액 3방향 검증 ─────────────────────────────────────
    rev_신고 = 당기신고서.get("총수입금액")
    rev_안내 = 안내문.get("수입금액")
    rev_지급 = sum(d["지급총액"] for d in 지급명세서) if 지급명세서 else None

    # A1. 신고서 vs 지급명세서
    if rev_신고 is not None and rev_지급 is not None:
        diff = rev_신고 - rev_지급
        add("수입금액", "신고서 vs 지급명세서",
            "[신고서] 총수입금액(⑨)", rev_신고, "[지급명세서] 지급총액 합계", rev_지급,
            "pass" if diff == 0 else "fail",
            메모="" if diff == 0 else f"차이 {diff:,}원",
            diff=diff)
    elif rev_신고 is None:
        add("수입금액", "신고서 vs 지급명세서",
            "[신고서] 총수입금액", "파싱실패", "[지급명세서]", rev_지급, "warn", "신고서 파싱 실패")

    # A2. 신고서 vs 안내문
    if rev_신고 is not None and rev_안내 is not None:
        diff = rev_신고 - rev_안내
        add("수입금액", "신고서 vs 안내문",
            "[신고서] 총수입금액(⑨)", rev_신고, "[안내문] 수입금액 총계", rev_안내,
            "pass" if diff == 0 else "fail",
            메모="" if diff == 0 else f"차이 {diff:,}원",
            diff=diff)

    # ── B. 원천징수세액 ────────────────────────────────────────────
    pre_신고 = 당기신고서.get("기납부세액")
    pre_안내 = 안내문.get("기납부세액")
    pre_지급 = sum(d["소득세"] for d in 지급명세서) if 지급명세서 else None

    # B0. 신고서(32번) vs 안내문 — 직원 입력 체크
    if pre_신고 is not None and pre_안내 is not None:
        diff = pre_신고 - pre_안내
        add("원천징수세액", "신고서 vs 안내문 (직원 입력 체크)",
            "[신고서] 기납부세액(32번)", pre_신고,
            "[안내문] 원천징수세액", pre_안내,
            "pass" if diff == 0 else "fail",
            메모="" if diff == 0 else f"차이 {diff:,}원 — 직원 입력값 확인 필요",
            diff=diff)

    # B1. 신고서(32번) vs 지급명세서 소득세 합계 — 원천징수 완전성
    if pre_신고 is not None and pre_지급 is not None:
        diff = pre_신고 - pre_지급
        add("원천징수세액", "신고서 vs 지급명세서 (원천징수 완전성)",
            "[신고서] 기납부세액(32번)", pre_신고,
            "[지급명세서] 소득세 합계", pre_지급,
            "pass" if diff == 0 else "fail",
            메모="" if diff == 0 else f"차이 {diff:,}원",
            diff=diff)
    # B2 제거: 안내문 vs 지급명세서 (신고서==안내문이면 중복)

    # ── C. 거래처별 — 지급명세서 내역 (참고용) ─────────────────────
    # 개별 row: 지급총액·소득세 참고용 (지급총액-소득세 차이는 무의미 → diff=None)
    # 합산 row 제거: 원천징수세액 B1에서 동일 비교 수행
    for d in 지급명세서:
        biz = d["사업자번호"]
        amt = d["지급총액"]
        tax = d["소득세"]
        징수 = d["징수의무자"]
        add("거래처별", f"{징수}({biz})",
            "[지급명세서] 지급총액", amt, "[지급명세서] 소득세", tax,
            "pass",
            메모=f"소득세율 {tax/amt*100:.1f}%" if amt else "",
            diff=None)   # 지급총액-소득세 차이는 무의미

    # ── D. 세액 계산 검증 ─────────────────────────────────────────
    과세표준 = 당기신고서.get("과세표준")
    세율 = 당기신고서.get("세율")
    산출세액_신고 = 당기신고서.get("산출세액")

    if 과세표준 and 세율 and 산출세액_신고:
        # 누진공제 테이블 — 귀속연도별 세율 (2023년 세법개정 반영)
        # 2022년 이전: 1,200만/4,600만 구간
        # 2023년 이후: 1,400만/5,000만 구간 (소득세법 §55 개정)
        귀속 = 당기신고서.get("귀속연도") or 2024
        if 귀속 >= 2023:
            세율_테이블 = [
                (14_000_000, 0.06, 0),
                (50_000_000, 0.15, 1_260_000),
                (88_000_000, 0.24, 5_760_000),
                (150_000_000, 0.35, 15_440_000),
                (300_000_000, 0.38, 19_940_000),
                (500_000_000, 0.40, 25_940_000),
                (1_000_000_000, 0.42, 35_940_000),
                (float("inf"), 0.45, 65_940_000),
            ]
        else:
            세율_테이블 = [
                (12_000_000, 0.06, 0),
                (46_000_000, 0.15, 1_080_000),
                (88_000_000, 0.24, 5_220_000),
                (150_000_000, 0.35, 14_900_000),
                (300_000_000, 0.38, 19_400_000),
                (500_000_000, 0.40, 25_400_000),
                (1_000_000_000, 0.42, 35_400_000),
                (float("inf"), 0.45, 65_400_000),
            ]
        for 구간상한, 세율값, 누진공제 in 세율_테이블:
            if 과세표준 <= 구간상한:
                산출세액_계산 = int(과세표준 * 세율값 - 누진공제)
                break

        diff = 산출세액_신고 - 산출세액_계산
        add("세액계산", f"과세표준×세율-누진공제",
            "[신고서] 산출세액(23번)", 산출세액_신고, "[계산] 과세표준×세율-누진공제", 산출세액_계산,
            "pass" if abs(diff) <= 1 else "fail",
            메모="" if abs(diff) <= 1 else f"차이 {diff:,}원 (반올림 허용 ±1)",
            diff=diff)

    # ── E. 환급율 ─────────────────────────────────────────────────
    결정세액 = 당기신고서.get("결정세액")
    pre_val = pre_신고

    if 결정세액 is not None and pre_val:
        환급액 = pre_val - 결정세액
        환급율 = 환급액 / pre_val * 100
        상태 = "pass"
        if 환급율 < 0:
            상태 = "warn"
            메모 = f"납부 {abs(환급액):,}원"
        else:
            메모 = f"환급 {환급액:,}원"
        add("환급", "환급율",
            "[신고서] 기납부세액(32번)", pre_val, "[신고서] 결정세액(28번)", 결정세액,
            상태, 메모=f"{메모} | 환급율 {환급율:.1f}%",
            diff=환급액 if 환급율 >= 0 else -abs(환급액))

    # ── F. 경비 분석 (카드) ───────────────────────────────────────
    필요경비 = 당기신고서.get("필요경비")
    카드_합계 = sum(c["합계"] for c in 카드_목록)

    if 필요경비 and 카드_합계 > 0 and rev_신고:
        카드율 = 카드_합계 / rev_신고 * 100
        경비율 = 필요경비 / rev_신고 * 100
        add("경비분석", "카드경비 / 필요경비",
            f"[카드] 합계({len(카드_목록)}개사)", 카드_합계,
            "[신고서] 필요경비(⑩)", 필요경비,
            "pass" if 카드_합계 <= 필요경비 else "warn",
            메모=f"카드 {카드율:.1f}% / 경비율 {경비율:.1f}%")

    # ── G. 전기 vs 당기 비교 ──────────────────────────────────────
    if 전기신고서:
        rev_전 = 전기신고서.get("총수입금액")
        rev_당 = 당기신고서.get("총수입금액")
        inc_전 = 전기신고서.get("소득금액")
        inc_당 = 당기신고서.get("소득금액")
        exp_전 = 전기신고서.get("필요경비")
        exp_당 = 당기신고서.get("필요경비")

        # G1. 수입금액 증감
        if rev_전 and rev_당:
            변동율 = (rev_당 - rev_전) / rev_전 * 100
            add("전기당기", f"수입금액 증감 ({전기신고서.get('귀속연도')}→{당기신고서.get('귀속연도')})",
                f"[신고서] {전기신고서.get('귀속연도')}년 수입", rev_전,
                f"[신고서] {당기신고서.get('귀속연도')}년 수입", rev_당,
                "warn" if abs(변동율) >= 30 else "pass",
                메모=f"{변동율:+.1f}%  {'⚠ 30% 이상 변동' if abs(변동율) >= 30 else ''}",
                diff=rev_당 - rev_전)

        # G2. 소득률 변동
        if rev_전 and inc_전 and rev_당 and inc_당:
            소득률_전 = inc_전 / rev_전 * 100
            소득률_당 = inc_당 / rev_당 * 100
            변동 = 소득률_당 - 소득률_전
            add("전기당기", "소득률 변동",
                f"[신고서] {전기신고서.get('귀속연도')}년 소득률", round(소득률_전, 2),
                f"[신고서] {당기신고서.get('귀속연도')}년 소득률", round(소득률_당, 2),
                "warn" if abs(변동) >= 5 else "pass",
                메모=f"변동 {변동:+.2f}%p  {'⚠ 5%p 초과' if abs(변동) >= 5 else ''}",
                diff=round(변동, 2))

        # G3. 경비율 변동
        if rev_전 and exp_전 and rev_당 and exp_당:
            경비율_전 = exp_전 / rev_전 * 100
            경비율_당 = exp_당 / rev_당 * 100
            변동 = 경비율_당 - 경비율_전
            add("전기당기", "경비율 변동",
                f"[신고서] {전기신고서.get('귀속연도')}년 경비율", round(경비율_전, 2),
                f"[신고서] {당기신고서.get('귀속연도')}년 경비율", round(경비율_당, 2),
                "warn" if abs(변동) >= 5 else "pass",
                메모=f"변동 {변동:+.2f}%p  {'⚠ 5%p 초과' if abs(변동) >= 5 else ''}",
                diff=round(변동, 2))

    return results


# ═══════════════════════════════════════════════════════════════════
# 4. HTML 보고서 생성
# ═══════════════════════════════════════════════════════════════════

_CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Malgun Gothic', '맑은 고딕', sans-serif; font-size: 13px;
       background: #f5f5f5; color: #222; }
.wrap { max-width: 960px; margin: 0 auto; padding: 16px; }
.title-bar { background: #1F4E79; color: #fff; padding: 14px 20px; border-radius: 6px 6px 0 0; }
.title-bar h1 { font-size: 17px; }
.title-bar .sub { font-size: 12px; opacity: .8; margin-top: 4px; }
.summary-bar { display: flex; gap: 12px; background: #fff; padding: 12px 16px;
               border: 1px solid #ddd; border-top: none; margin-bottom: 12px; }
.sum-box { flex: 1; text-align: center; border-radius: 4px; padding: 10px; }
.sum-pass { background: #e8f5e9; color: #2e7d32; }
.sum-warn { background: #fff3e0; color: #e65100; }
.sum-fail { background: #ffebee; color: #c62828; }
.sum-box .num { font-size: 28px; font-weight: bold; }
.sum-box .lbl { font-size: 11px; margin-top: 2px; }
.section { background: #fff; border: 1px solid #ddd; border-radius: 4px;
           margin-bottom: 10px; overflow: hidden; }
.sec-hdr { background: #2E75B6; color: #fff; padding: 8px 14px; font-size: 13px; font-weight: bold; }
table { width: 100%; border-collapse: collapse; }
th { background: #D9E1F2; padding: 7px 10px; font-size: 12px;
     border-bottom: 1px solid #bbb; text-align: center; }
td { padding: 7px 10px; border-bottom: 1px solid #eee; font-size: 12px; }
td.num { text-align: right; font-family: monospace; }
td.pct { text-align: right; font-size: 12px; color: #1F4E79; font-weight: bold; font-family: monospace; }
.pass td.status { color: #2e7d32; font-weight: bold; }
.warn td.status { color: #e65100; font-weight: bold; }
.fail td.status { color: #c62828; font-weight: bold; }
.pass { background: #f9fff9; }
.warn { background: #fffdf5; }
.fail { background: #fff5f5; }
.file-list { padding: 10px 14px; }
.file-list li { padding: 3px 0; font-size: 12px; }
.badge { display: inline-block; padding: 1px 7px; border-radius: 10px; font-size: 11px; font-weight: bold; }
.badge-신고서 { background: #e3f2fd; color: #1565c0; }
.badge-안내문 { background: #f3e5f5; color: #6a1b9a; }
.badge-지급명세서 { background: #e8f5e9; color: #2e7d32; }
.badge-카드 { background: #fff8e1; color: #f57f17; }
.badge-기타 { background: #eceff1; color: #546e7a; }
.footer { text-align: center; color: #999; font-size: 11px; margin-top: 16px; }
@media print {
  body { background: #fff; print-color-adjust: exact; -webkit-print-color-adjust: exact; }
  .wrap { max-width: 100%; padding: 8px; }
  .section { page-break-inside: avoid; }
}
"""


def _fmt(v) -> str:
    if v is None:
        return "—"
    if isinstance(v, float):
        if v == int(v):
            return f"{int(v):,}"
        return f"{v:,.2f}"
    if isinstance(v, int):
        return f"{v:,}"
    return str(v)


def _fmt_label(lbl: str) -> str:
    """[문서명] 태그를 색상 뱃지 + 항목명으로 변환 (HTML 안전)"""
    if not lbl:
        return ""
    badge_map = {
        "[신고서]":    ("신고서",    "#dbeafe", "#1d4ed8"),
        "[안내문]":    ("안내문",    "#f3e5f5", "#6a1b9a"),
        "[지급명세서]": ("지급명세서","#dcfce7", "#15803d"),
        "[카드]":      ("카드",      "#fef9c3", "#a16207"),
        "[계산]":      ("계산",      "#ffedd5", "#c2410c"),
    }
    for tag, (text, bg, fg) in badge_map.items():
        if lbl.startswith(tag):
            badge = (f'<span style="display:inline-block;padding:1px 6px;'
                     f'border-radius:3px;font-size:10px;font-weight:bold;'
                     f'background:{bg};color:{fg};margin-right:4px">{text}</span>')
            rest = lbl[len(tag):].strip()
            return badge + rest
    return lbl


def generate_html(
    name: str,
    jumin6: str,
    files: dict,
    당기신고서: dict,
    전기신고서: dict | None,
    안내문_data: dict,
    지급명세서: list,
    카드_목록: list,
    verify_results: list,
) -> str:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    귀속 = 당기신고서.get("귀속연도", "")

    # 요약 카운트
    n_pass = sum(1 for r in verify_results if r["상태"] == "pass")
    n_warn = sum(1 for r in verify_results if r["상태"] == "warn")
    n_fail = sum(1 for r in verify_results if r["상태"] == "fail")

    def status_icon(s):
        return {"pass": "✓", "warn": "⚠", "fail": "✗"}.get(s, "?")

    # 파일 목록 HTML
    file_rows = ""
    badge_map = {
        "신고서": "신고서", "안내문": "안내문", "지급명세서": "지급명세서",
        "신한카드": "카드", "비씨카드": "카드",
    }
    for role, paths in files.items():
        if role == "기타" or not paths:
            continue
        for p in paths:
            badge = badge_map.get(role, "기타")
            file_rows += (f'<li><span class="badge badge-{badge}">{role}</span> '
                          f'&nbsp; {p.name}</li>\n')

    # 검증 결과 테이블
    섹션별 = {}
    for r in verify_results:
        섹션별.setdefault(r["섹션"], []).append(r)

    verify_html = ""
    for 섹션, rows in 섹션별.items():
        rows_html = ""
        for r in rows:
            cls = r["상태"]
            diff_str = ""
            if r["차이"] is not None:
                if isinstance(r["차이"], float):
                    diff_str = f"{r['차이']:+.2f}"
                else:
                    diff_str = f"{r['차이']:+,}" if r["차이"] != 0 else "0"

            rows_html += f"""
<tr class="{cls}">
  <td class="status">{status_icon(r['상태'])}</td>
  <td>{r['항목']}</td>
  <td>{_fmt_label(r['값1_label'])}</td>
  <td class="num">{_fmt(r['값1'])}</td>
  <td>{_fmt_label(r['값2_label'])}</td>
  <td class="num">{_fmt(r['값2'])}</td>
  <td class="num" style="color:{'red' if r['차이'] and r['차이']!=0 and not isinstance(r['차이'],float) else '#555'}">{diff_str}</td>
  <td style="color:#666">{r.get('메모','')}</td>
</tr>"""

        verify_html += f"""
<div class="section">
  <div class="sec-hdr">{섹션}</div>
  <table>
    <tr>
      <th style="width:30px"></th>
      <th style="text-align:left">항목</th>
      <th>출처 1</th>
      <th>값 1</th>
      <th>출처 2</th>
      <th>값 2</th>
      <th>차이</th>
      <th style="text-align:left">메모</th>
    </tr>
    {rows_html}
  </table>
</div>"""

    # 신고서 비교표 + 안내문 공제 체크
    당기_yr = 당기신고서.get("귀속연도", "당기")
    전기_yr = 전기신고서.get("귀속연도", "전기") if 전기신고서 else "—"
    compare_fields = [
        ("총수입금액", "총수입금액"),
        ("필요경비", "필요경비"),
        ("소득금액", "소득금액"),
        ("소득공제", "소득공제"),
        ("과세표준", "과세표준"),
        ("산출세액", "산출세액"),
        ("세액공제", "세액공제"),
        ("결정세액", "결정세액"),
        ("기납부세액", "기납부세액"),
        ("납부세액", "납부세액"),
    ]
    # 비율 계산용 수입금액
    rev_전_base = 전기신고서.get("총수입금액") if 전기신고서 else None
    rev_당_base = 당기신고서.get("총수입금액")

    def _pct(v, base):
        """총수입금액 대비 비율 문자열"""
        if isinstance(v, (int, float)) and isinstance(base, (int, float)) and base != 0:
            return f"{v / base * 100:.1f}%"
        return ""

    # 안내문 자동 체크 가능 항목: 신고서 key → 안내문 key
    안내문_매핑 = {
        "총수입금액": "수입금액",
        "기납부세액": "기납부세액",
    }
    # 소득공제·세액공제 세부항목 (안내문 값 표시 + 수동 확인)
    공제_서브항목 = {
        "소득공제": [
            ("↳ 국민연금보험료", "국민연금보험료", "❓ 확인"),
            ("↳ 개인연금저축",   "개인연금저축",   "❓ 확인"),
            ("↳ 노란우산공제",   "노란우산공제",   "❓ 한도 500만"),
        ],
        "세액공제": [
            ("↳ 퇴직연금세액공제", "퇴직연금세액공제", "❓ 확인"),
            ("↳ 연금계좌세액공제", "연금계좌세액공제", "❓ 확인"),
        ],
    }

    compare_rows = ""
    for label, key in compare_fields:
        v_당 = 당기신고서.get(key)
        v_전 = 전기신고서.get(key) if 전기신고서 else None
        diff = None
        if isinstance(v_당, (int, float)) and isinstance(v_전, (int, float)):
            diff = v_당 - v_전

        pct_전 = _pct(v_전, rev_전_base)
        pct_당 = _pct(v_당, rev_당_base)
        pct_style = ' style="color:#c62828"' if key == "소득금액" else ""

        # 증감 색상
        if diff is None:
            diff_color = "#555"
        elif diff < 0:
            diff_color = "#c62828"
        elif diff > 0:
            diff_color = "#2e7d32"
        else:
            diff_color = "#555"

        # 안내문 값 + 체크 아이콘
        안내_key = 안내문_매핑.get(key)
        v_안내 = 안내문_data.get(안내_key) if 안내_key else None
        안내_str = f"{v_안내:,}" if isinstance(v_안내, int) else "—"
        if v_안내 is not None and isinstance(v_당, int):
            체크_str = ('<span style="color:#2e7d32;font-weight:bold">✓</span>'
                        if v_당 == v_안내 else
                        '<span style="color:#c62828;font-weight:bold">✗</span>')
        else:
            체크_str = ""

        compare_rows += f"""
<tr>
  <td>{label}</td>
  <td class="num">{_fmt(v_전)}</td>
  <td class="num pct"{pct_style}>{pct_전}</td>
  <td class="num">{_fmt(v_당)}</td>
  <td class="num pct"{pct_style}>{pct_당}</td>
  <td class="num" style="color:{diff_color}">
    {f'{diff:+,}' if isinstance(diff, int) else _fmt(diff)}
  </td>
  <td class="num">{안내_str}</td>
  <td style="text-align:center">{체크_str}</td>
</tr>"""

        # 공제 세부항목 sub-rows (안내문에 값이 있을 때만)
        for sub_label, sub_key, sub_check in 공제_서브항목.get(key, []):
            v_sub = 안내문_data.get(sub_key)
            if v_sub is not None:
                compare_rows += f"""
<tr style="background:#f5f5ff">
  <td style="padding-left:22px; color:#5c6bc0; font-size:11px">{sub_label}</td>
  <td class="num" style="color:#bbb">—</td>
  <td class="num"></td>
  <td class="num" style="color:#bbb">—</td>
  <td class="num"></td>
  <td class="num"></td>
  <td class="num" style="color:#5c6bc0; font-weight:bold">{v_sub:,}</td>
  <td style="text-align:center; color:#e65100; font-size:11px">{sub_check}</td>
</tr>"""

    신고서_html = f"""
<div class="section">
  <div class="sec-hdr">신고서 전기↔당기 비교 + 안내문 공제 체크</div>
  <table>
    <tr>
      <th style="text-align:left">항목</th>
      <th>{전기_yr}년 (전기)</th>
      <th style="color:#888;font-weight:normal">%</th>
      <th>{당기_yr}년 (당기)</th>
      <th style="color:#888;font-weight:normal">%</th>
      <th>증감</th>
      <th>안내문</th>
      <th>체크</th>
    </tr>
    {compare_rows}
  </table>
</div>"""

    # 지급명세서 테이블
    지급_rows = ""
    for d in 지급명세서:
        지급_rows += f"""
<tr>
  <td>{d['사업자번호']}</td>
  <td>{d['징수의무자']}</td>
  <td class="num">{d['지급총액']:,}</td>
  <td class="num">{d['소득세']:,}</td>
  <td class="num">{d['지방소득세']:,}</td>
</tr>"""

    지급_합계 = sum(d["지급총액"] for d in 지급명세서)
    세_합계 = sum(d["소득세"] for d in 지급명세서)
    지방_합계 = sum(d["지방소득세"] for d in 지급명세서)
    지급_rows += f"""
<tr style="font-weight:bold; background:#D9E1F2">
  <td colspan="2">합계</td>
  <td class="num">{지급_합계:,}</td>
  <td class="num">{세_합계:,}</td>
  <td class="num">{지방_합계:,}</td>
</tr>"""

    지급_html = f"""
<div class="section">
  <div class="sec-hdr">지급명세서 상세</div>
  <table>
    <tr>
      <th>사업자번호</th><th>징수의무자</th>
      <th>지급총액</th><th>소득세</th><th>지방소득세</th>
    </tr>
    {지급_rows}
  </table>
</div>"""

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>교차검증 보고서 — {name} {귀속}년귀속</title>
<style>{_CSS}</style>
</head>
<body>
<div class="wrap">

<div class="title-bar">
  <h1>종합소득세 교차검증 보고서</h1>
  <div class="sub">
    고객: {name}&nbsp;({jumin6})&nbsp; | &nbsp;{귀속}년 귀속&nbsp; | &nbsp;생성: {ts}&nbsp; | &nbsp;세무회계창연
  </div>
</div>

<div class="summary-bar">
  <div class="sum-box sum-pass"><div class="num">{n_pass}</div><div class="lbl">✓ 일치 (pass)</div></div>
  <div class="sum-box sum-warn"><div class="num">{n_warn}</div><div class="lbl">⚠ 주의 (warn)</div></div>
  <div class="sum-box sum-fail"><div class="num">{n_fail}</div><div class="lbl">✗ 불일치 (fail)</div></div>
  <div class="sum-box" style="background:#eceff1; color:#37474f; flex:2">
    <div style="font-size:13px; margin-top:4px">
      {"✅ 불일치 없음 — 신고 준비 완료" if n_fail == 0 else f"🚨 불일치 {n_fail}건 — 즉시 원인 확인 필요"}
    </div>
  </div>
</div>

<div class="section">
  <div class="sec-hdr">파일 인식 결과</div>
  <ul class="file-list">{file_rows}</ul>
</div>

{신고서_html}
{verify_html}
{지급_html}

<div class="footer">
  세무회계창연 &nbsp;|&nbsp; 장성환 세무사 &nbsp;|&nbsp; 생성: {ts}
</div>

</div>
</body>
</html>"""

    return html


# ═══════════════════════════════════════════════════════════════════
# 5. 메인
# ═══════════════════════════════════════════════════════════════════

def find_folder(name: str, jumin6: str = "") -> Path | None:
    nfc_name = unicodedata.normalize("NFC", name)
    candidates = [p for p in CUSTOMER_DIR.iterdir()
                  if p.is_dir() and unicodedata.normalize("NFC", p.name).startswith(f"{nfc_name}_")]
    if not candidates:
        p = CUSTOMER_DIR / name
        return p if p.is_dir() else None
    if jumin6:
        exact = [c for c in candidates
                 if unicodedata.normalize("NFC", c.name).endswith(f"_{jumin6}")]
        if exact:
            return exact[0]
    return candidates[0] if candidates else None


def run(name: str, jumin6: str = "", folder: Path = None) -> Path | None:
    if folder is None:
        folder = find_folder(name, jumin6)
    if not folder:
        print(f"  [오류] 폴더 없음: {name}")
        return None
    # name/jumin6 폴더명에서 자동 추출 (봇 호환)
    if not name and folder:
        parts = folder.name.rsplit("_", 1)
        name   = parts[0]
        jumin6 = parts[1] if len(parts) > 1 else ""

    print(f"\n{'─'*54}")
    print(f"  종합소득세 교차검증 보고서 생성기")
    print(f"{'─'*54}")
    print(f"  폴더: {folder}")

    # ── 1. 파일 인식 + 신고서 귀속연도 기준 중복 제거 ───────────────
    print(f"\n[1] 파일 인식 결과")
    files = classify_files(folder)

    # 신고서: 파싱 후 귀속연도 중복 제거 (같은 연도면 파일명 짧은 것 우선)
    raw_신고서 = []
    for p in files["신고서"]:
        sr = parse_tax_return(p)
        sr["_path"] = p
        raw_신고서.append(sr)

    seen_yr: dict[int, dict] = {}
    for sr in raw_신고서:
        yr = sr.get("귀속연도")
        if yr is None:
            continue
        prev = seen_yr.get(yr)
        if prev is None or len(sr.get("파일", "")) < len(prev.get("파일", "")):
            seen_yr[yr] = sr

    신고서_list = sorted(seen_yr.values(), key=lambda s: s.get("귀속연도", 0))

    # 파일 인식 출력 (신고서는 중복 제거 후만 표시)
    dedup_신고서_names = {sr["파일"] for sr in 신고서_list}
    for role, paths in files.items():
        if not paths:
            continue
        if role == "신고서":
            for p in paths:
                if p.name in dedup_신고서_names:
                    print(f"    [신고서      ] {p.name}")
                # else: 중복 파일 → 표시 생략
        elif role != "기타":
            for p in paths:
                print(f"    [{role:10s}] {p.name}")

    # ── 2. 신고서 파싱 결과 출력 ──────────────────────────────────
    print(f"\n[2] 신고서 PDF 파싱")
    for sr in 신고서_list:
        yr = sr.get("귀속연도", "?")
        rev = sr.get("총수입금액")
        det = sr.get("결정세액")
        print(f"    {yr}년귀속: 수입 {rev:,}원 / 결정세액 {det:,}원" if isinstance(rev, int) and isinstance(det, int) else f"    {sr.get('파일')} — 파싱 불완전")

    당기신고서 = 신고서_list[-1] if 신고서_list else {}
    전기신고서 = 신고서_list[-2] if len(신고서_list) >= 2 else None

    # 유효성 체크 제거: 근로+사업 복수소득 시 과세표준 > 사업소득 수입금액이 정상 케이스

    # ── 3. 안내문 파싱 ────────────────────────────────────────────
    print(f"\n[3] 안내문 PDF 파싱")
    안내문_data = {}
    if files["안내문"]:
        안내문_data = parse_anneam(files["안내문"][0])
        print(f"    수입금액: {안내문_data.get('수입금액','?'):,}   기납부: {안내문_data.get('기납부세액','?')}"
              if isinstance(안내문_data.get('수입금액'), int) else "    파싱 실패")

    # ── 4. 지급명세서 파싱 ────────────────────────────────────────
    print(f"\n[4] 지급명세서 파싱")
    지급명세서 = []
    for p in files["지급명세서"]:
        rows = parse_income_excel(p, jumin6)
        지급명세서.extend(rows)
        print(f"    {p.name}: {len(rows)}건 / 합계 {sum(r['지급총액'] for r in rows):,}원")

    # ── 5. 카드 파싱 ──────────────────────────────────────────────
    print(f"\n[5] 카드내역 파싱")
    카드_목록 = []
    for role in ["신한카드", "비씨카드"]:
        for p in files.get(role, []):
            r = parse_card_excel(p)
            카드_목록.append(r)
            print(f"    {r['카드사']}: {r['건수']}건 / {r['합계']:,}원")

    # ── 6. 교차검증 ───────────────────────────────────────────────
    print(f"\n[6] 교차검증 수행")
    results = cross_verify(당기신고서, 전기신고서, 안내문_data, 지급명세서, 카드_목록)
    for r in results:
        icon = {"pass": "✓", "warn": "⚠", "fail": "✗"}[r["상태"]]
        print(f"    {icon} [{r['섹션']}] {r['항목']}: {r.get('메모','')}")

    n_pass = sum(1 for r in results if r["상태"] == "pass")
    n_warn = sum(1 for r in results if r["상태"] == "warn")
    n_fail = sum(1 for r in results if r["상태"] == "fail")
    print(f"\n  완료: 일치 {n_pass}건 / 주의 {n_warn}건 / 불일치 {n_fail}건")

    # ── 7. HTML 보고서 생성 ───────────────────────────────────────
    print(f"\n[7] HTML 보고서 생성")

    # 파일 인식 결과에서 신고서는 중복 제거 후 목록만 표시
    dedup_신고서_paths = {sr["_path"] for sr in 신고서_list if "_path" in sr}
    files_for_html = dict(files)
    files_for_html["신고서"] = [p for p in files["신고서"] if p in dedup_신고서_paths]

    html = generate_html(
        name, jumin6, files_for_html,
        당기신고서, 전기신고서, 안내문_data,
        지급명세서, 카드_목록, results,
    )
    ts = datetime.now().strftime("%Y%m%d_%H%M")

    # 기존 검증보고서 → _archive 이동 (타임스탬프 파일명 그대로 보존)
    old_reports = nfc_glob(folder, "검증보고서_*.html")
    if old_reports:
        arch = folder / "_archive"
        arch.mkdir(exist_ok=True)
        for old in old_reports:
            try:
                old.rename(arch / old.name)
            except Exception:
                pass

    out_path = folder / f"검증보고서_{ts}.html"
    out_path.write_text(html, encoding="utf-8")
    print(f"    ✅ {out_path.name}")
    print(f"{'─'*54}\n")

    return out_path


def load_customers():
    """구글시트 접수명단 → [{name, jumin6}, ...]"""
    from gsheet_writer import get_credentials
    import gspread
    GSHEET_ID  = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
    SHEET_NAME = "접수명단"
    COL_NAME   = 2
    COL_JUMIN  = 4
    creds = get_credentials()
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(GSHEET_ID).worksheet(SHEET_NAME)
    rows = ws.get_all_values()
    customers = []
    for row in rows[1:]:
        name  = row[COL_NAME].strip()  if len(row) > COL_NAME  else ""
        jumin = row[COL_JUMIN].strip() if len(row) > COL_JUMIN else ""
        if not name:
            continue
        customers.append({"name": name, "jumin6": jumin.replace("-", "")[:6]})
    return customers


def main():
    args = sys.argv[1:]
    if not args:
        print("사용법: python 검증.py 이름 [주민앞6자리]")
        print("       python 검증.py --all")
        sys.exit(1)

    if args[0] == "--all":
        customers = load_customers()
        total = len(customers)
        ok = err = 0
        for i, c in enumerate(customers, 1):
            try:
                out = run(c["name"], c["jumin6"])
                if out:
                    ok += 1
                else:
                    err += 1
            except Exception as e:
                print(f"  [오류] {c['name']}: {e}")
                err += 1
            if i % 10 == 0 or i == total:
                print(f"  진행: {i}/{total} (성공 {ok} / 오류 {err})")
        print(f"\n[완료] {ok}/{total}명 검증보고서 생성")
        return

    # 폴더명 직접 넘긴 경우 ("박현민_870529")
    if len(args) == 1 and "_" in args[0]:
        parts = args[0].rsplit("_", 1)
        name, jumin6 = parts[0], parts[1] if len(parts) > 1 else ""
    else:
        name   = args[0]
        jumin6 = args[1] if len(args) > 1 else ""

    run(name, jumin6)


if __name__ == "__main__":
    main()
