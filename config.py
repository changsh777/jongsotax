"""
종소세 자동화 공통 설정
- 베이스 경로를 한 곳에서 관리
- local (테스트) → nas (운영) 이전 시 ENV만 변경
"""
from pathlib import Path
import os
import platform

# ----- 환경 선택 -----
# 환경변수 SEOTAX_ENV로 바꿀 수 있음 ('local' or 'nas')
ENV = os.environ.get("SEOTAX_ENV", "local")

# ----- 경로 정의 -----
if platform.system() == "Darwin":
    # 맥미니 — NAS SMB 마운트
    BASE = Path("/Volumes/장성환/종소세2026")
elif ENV == "nas":
    # Windows NAS (Z: 드라이브 마운트)
    BASE = Path(r"Z:\종소세2026")
else:
    # Windows 로컬 테스트
    BASE = Path(r"F:\종소세2026")

# NAS 폴더 구조
CUSTOMER_DIR  = BASE / "고객"       # 인별 폴더 (성명_주민앞6자리)
MYUNGDAN_DIR  = BASE / "_명단"      # 접수명단 CSV
LOG_DIR       = BASE / "_로그"      # 실행 로그

# 하위 호환 (기존 스크립트용)
INPUT_DIR     = BASE / "input"
OUTPUT_DIR    = BASE / "output"
PDF_DIR       = CUSTOMER_DIR        # 고객 폴더 = 기존 PDF_DIR 역할
LOGS_DIR      = LOG_DIR
TEMPLATES_DIR = BASE / "templates"
CREDENTIALS_DIR = BASE / ".credentials"

INPUT_XLSX       = INPUT_DIR  / "종소세신고도움서비스테스트.xlsx"
RESULT_XLSX      = OUTPUT_DIR / "결과.xlsx"
PARSE_RESULT_XLSX = OUTPUT_DIR / "파싱결과.xlsx"

# 구글시트
GSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
GSHEET_WORKSHEET = "안내문파싱"

# 홈택스 URL
HOMETAX_MAIN = "https://hometax.go.kr"

REPORT_HELP_URL_TAX_AGENT = (  # 세무대리인 전용 신고도움서비스 (Track A)
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml"
    "&tmIdx=06&tm2lIdx=0601000000&tm3lIdx=0601200000"
)

REPORT_HELP_URL_PERSONAL = (  # 개인 신고도움서비스 (Track B)
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml&tmIdx=41&tm2lIdx=4103000000&tm3lIdx=4103080000"
)

LOGIN_URL = (
    "https://hometax.go.kr/websquare/websquare.html"
    "?w2xPath=/ui/pp/index_pp.xml&menuCd=index3"
)


def customer_folder(name: str, jumin: str = "") -> Path:
    """고객 폴더 반환 + 자동 생성.

    Args:
        name: 성명 (예: '홍길동')
        jumin: 주민번호 전체 또는 앞 6자리 (예: '800101' or '800101-1234567')

    Returns:
        Path: 고객/홍길동_800101/

    모든 자동화 스크립트가 이 함수만 사용 (경로 직접 조립 금지)
    """
    import re
    jumin_front = re.sub(r'[^0-9]', '', str(jumin))[:6]  # 숫자만, 앞 6자리
    folder_name = f"{name}_{jumin_front}" if jumin_front else name
    folder = CUSTOMER_DIR / folder_name
    # 하위 폴더도 함께 생성
    (folder / "자료").mkdir(parents=True, exist_ok=True)
    return folder


if __name__ == "__main__":
    print(f"환경: {ENV}")
    print(f"BASE: {BASE}")
    print(f"PDF_DIR: {PDF_DIR}")
    print(f"존재함: {BASE.exists()}")
