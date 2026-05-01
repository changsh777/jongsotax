"""
구글시트 연동 모듈
- OAuth 첫 인증 시 브라우저 한 번 열림 → 자동으로 토큰 저장
- 이후 자동 갱신
- 시트에 행 단위로 누적 (덮어쓰지 않음, 동일 성명 행은 갱신)
"""
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from pathlib import Path
import pickle

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

CRED_DIR = Path(r"F:\종소세2026\.credentials")
CLIENT_SECRET = CRED_DIR / "client_secret.json"
TOKEN_FILE = CRED_DIR / "token.pickle"

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
WORKSHEET_NAME = "안내문파싱"  # 시트 탭 이름. 없으면 자동 생성

# 미동의명단 시트: step4/5 에러난 사람 → step7 수임동의 대상
CONSENT_SHEET_NAME = "미동의명단"
CONSENT_COLUMNS = [
    "성명", "주민번호", "핸드폰번호",       # 기본 정보 (step4에서 복사)
    "에러사유",                             # step4/5에서 기록
    "수임상태",                             # step7이 기록 (1트랙_동의요청 / 2트랙_해임후동의 / 우리수임완료 / ...)
    "시도일시",                             # step7이 기록
    "alert원문",                            # step7이 기록 (디버그용)
    "카카오발송문",                          # step7이 생성 (사람이 검토 후 발송)
    "비고",
]

COLUMNS = [
    "성명", "생년월일", "기장의무", "추계시적용경비율", "수입금액총계",
    "이자", "배당", "근로(단일)", "근로(복수)", "연금", "기타",
    "사업장부 정가", "타소득가산", "합산정가",
    "사전접수할인가", "일반접수가",
    "전년도_총수입금액", "전년도_필요경비", "전년도_종합소득금액", "전년도_소득공제",
    "전년도_과세표준", "전년도_산출세액", "전년도_세액감면공제", "전년도_결정세액",
    "전년도_가산세", "전년도_기납부세액", "전년도_납부할총세액",
    "부가세_매출", "부가세_매입", "부가세_납부",
    "처리상태", "처리일시", "PDF경로",
]


def get_credentials():
    """OAuth credentials 로드 또는 첫 인증"""
    creds = None
    if TOKEN_FILE.exists():
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CLIENT_SECRET), SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return creds


def get_worksheet():
    """시트 객체 반환 (없으면 워크시트 생성 + 헤더 작성)"""
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=100, cols=20)
        ws.append_row(COLUMNS, value_input_option="USER_ENTERED")

    # 헤더 항상 최신 COLUMNS로 동기화
    current_header = ws.row_values(1)
    if current_header != COLUMNS:
        from gspread.utils import rowcol_to_a1
        end_a1 = rowcol_to_a1(1, len(COLUMNS))
        ws.update(f"A1:{end_a1}", [COLUMNS])

    return ws


def _end_col_letter():
    from gspread.utils import rowcol_to_a1
    return rowcol_to_a1(1, len(COLUMNS)).rstrip("0123456789")


def upsert_row(data):
    """성명 기준 upsert: 있으면 업데이트, 없으면 추가"""
    ws = get_worksheet()
    name = data.get("성명", "")
    if not name:
        raise ValueError("성명이 없는 데이터")

    all_rows = ws.get_all_values()
    target_row_idx = None
    for i, row in enumerate(all_rows[1:], start=2):
        if row and row[0] == name:
            target_row_idx = i
            break

    new_row = [data.get(c, "") for c in COLUMNS]
    end_col = _end_col_letter()

    if target_row_idx:
        ws.update(f"A{target_row_idx}:{end_col}{target_row_idx}", [new_row])
        return f"업데이트 (행 {target_row_idx})"
    else:
        ws.append_row(new_row, value_input_option="USER_ENTERED")
        return f"신규 추가"


def write_all(rows):
    """전체 다시 쓰기 + 숫자 컬럼 number format 강제"""
    ws = get_worksheet()
    end_col = _end_col_letter()
    ws.batch_clear([f"A2:{end_col}"])
    if rows:
        values = [[r.get(c, "") for c in COLUMNS] for r in rows]
        ws.update(f"A2:{end_col}{len(values)+1}", values, value_input_option="USER_ENTERED")
        _enforce_number_format(ws, len(values))
    return len(rows)


def _enforce_number_format(ws, n_rows):
    """숫자 컬럼 (수입금액·수수료·전년도·부가세) format을 #,##0 으로 강제"""
    from gspread.utils import rowcol_to_a1
    numeric_cols = [
        "수입금액총계",
        "사업장부 정가", "타소득가산", "합산정가",
        "사전접수할인가", "일반접수가",
        "전년도_총수입금액", "전년도_필요경비", "전년도_종합소득금액", "전년도_소득공제",
        "전년도_과세표준", "전년도_산출세액", "전년도_세액감면공제", "전년도_결정세액",
        "전년도_가산세", "전년도_기납부세액", "전년도_납부할총세액",
        "부가세_매출", "부가세_매입", "부가세_납부",
    ]
    fmt = {"numberFormat": {"type": "NUMBER", "pattern": "#,##0"}}
    for col_name in numeric_cols:
        if col_name not in COLUMNS:
            continue
        col_idx = COLUMNS.index(col_name) + 1
        col_letter = rowcol_to_a1(1, col_idx).rstrip("0123456789")
        rng = f"{col_letter}2:{col_letter}{n_rows + 1}"
        try:
            ws.format(rng, fmt)
        except Exception as e:
            print(f"    [format 실패 {col_name}] {e}")


# ===== 미동의명단 시트 관련 함수 =====

def get_consent_worksheet():
    """미동의명단 시트 반환 (없으면 생성 + 헤더)"""
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet(CONSENT_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=CONSENT_SHEET_NAME, rows=500, cols=len(CONSENT_COLUMNS))
        ws.append_row(CONSENT_COLUMNS, value_input_option="USER_ENTERED")

    # 헤더 최신화
    current_header = ws.row_values(1)
    if current_header != CONSENT_COLUMNS:
        from gspread.utils import rowcol_to_a1
        end_a1 = rowcol_to_a1(1, len(CONSENT_COLUMNS))
        ws.update(f"A1:{end_a1}", [CONSENT_COLUMNS])
    return ws


def upsert_consent_row(data: dict):
    """미동의명단 upsert (성명 기준)

    data 키: CONSENT_COLUMNS 참조
    """
    ws = get_consent_worksheet()
    name = data.get("성명", "")
    if not name:
        raise ValueError("성명 없음")

    all_rows = ws.get_all_values()
    target_idx = None
    for i, row in enumerate(all_rows[1:], start=2):
        if row and row[0] == name:
            target_idx = i
            break

    new_row = [data.get(c, "") for c in CONSENT_COLUMNS]
    from gspread.utils import rowcol_to_a1
    end_col = rowcol_to_a1(1, len(CONSENT_COLUMNS)).rstrip("0123456789")

    if target_idx:
        ws.update(f"A{target_idx}:{end_col}{target_idx}", [new_row])
        return f"업데이트 (행 {target_idx})"
    else:
        ws.append_row(new_row, value_input_option="USER_ENTERED")
        return "신규 추가"


def load_consent_rows() -> list[dict]:
    """미동의명단 전체 로딩 (헤더 기반 dict 리스트)"""
    ws = get_consent_worksheet()
    return ws.get_all_records()


def update_consent_status(ws, row_idx: int, status: str,
                          alert_raw: str = "", kakao_msg: str = ""):
    """미동의명단 수임상태·시도일시·alert·카카오문 갱신"""
    from datetime import datetime
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    col = {c: i + 1 for i, c in enumerate(CONSENT_COLUMNS)}
    ws.update_cell(row_idx, col["수임상태"], status)
    ws.update_cell(row_idx, col["시도일시"], now)
    if alert_raw:
        ws.update_cell(row_idx, col["alert원문"], alert_raw[:300])
    if kakao_msg:
        ws.update_cell(row_idx, col["카카오발송문"], kakao_msg)


def read_customers_from_gsheet() -> list[dict]:
    """접수명단 시트에서 고객 목록 읽기 → step4 인풋 형식으로 반환

    Returns:
        list of {"name": str, "jumin_raw": str, "phone_raw": str}
        주민번호 없으면 스킵
    """
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.worksheet("접수명단")
    rows = ws.get_all_records()

    out = []
    for row in rows:
        name  = str(row.get("성명", "") or "").strip()
        jumin = str(row.get("주민번호", "") or "").strip()
        phone = str(row.get("핸드폰번호", "") or "").strip()
        if not name or not jumin:
            continue
        out.append({"name": name, "jumin_raw": jumin, "phone_raw": phone})
    print(f"[접수명단] {len(out)}건 로드")
    return out


if __name__ == "__main__":
    print("[1] OAuth 인증 시도 (첫 실행 시 브라우저 열림)")
    creds = get_credentials()
    print(f"    토큰: {'유효' if creds.valid else '만료'}")
    print(f"[2] 시트 접근")
    ws = get_worksheet()
    print(f"    워크시트: {ws.title} ({ws.row_count}행 x {ws.col_count}열)")
    print(f"[3] 헤더 확인: {ws.row_values(1)}")
    print("\n[테스트] 더미 데이터 1건 upsert")
    test_data = {c: "" for c in COLUMNS}
    test_data["성명"] = "테스트"
    test_data["기장의무"] = "테스트용"
    test_data["처리일시"] = "2026-04-25 22:50:00"
    result = upsert_row(test_data)
    print(f"    결과: {result}")
    print(f"\n시트 확인: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
