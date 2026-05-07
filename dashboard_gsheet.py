"""
dashboard_gsheet.py — 고객 폴더 현황 → 구글시트 대시보드
세무회계창연 | 2026

실행 (맥미니):
    python3 ~/종소세2026/dashboard_gsheet.py

→ 구글시트 "현황대시보드" 탭을 전체 갱신
"""

import os
import sys
import platform
import unicodedata
import pickle
from pathlib import Path
from datetime import datetime

import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, str(Path(__file__).parent))
from config import CUSTOMER_DIR

# ── 인증 경로 (OS별) ─────────────────────────────────────────────────────
if platform.system() == "Darwin":
    CRED_DIR = Path.home() / "종소세2026" / ".credentials"
else:
    CRED_DIR = Path(r"F:\종소세2026\.credentials")

CLIENT_SECRET = CRED_DIR / "client_secret.json"
TOKEN_FILE    = CRED_DIR / "token.pickle"
SCOPES        = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SPREADSHEET_ID       = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"  # 세무사 내부용
STAFF_SPREADSHEET_ID = "1ht7fk381ei8fJ33KpigeMSaEdjlOifVFzgN2Dyy-dpk"  # 직원 공유용
SHEET_NAME           = "현황대시보드"


# ── 컬럼 정의 ─────────────────────────────────────────────────────────────
# (key, 헤더명, glob 패턴 or lambda(folder,name)→Path|None, 중요도)
# 중요도: "high" = 없으면 빨강, "mid" = 없으면 주황, "low" = 없으면 흰색

def _g(folder, pattern):
    """NFC/NFD 양쪽 정규화 후 fnmatch — macOS SMB 마운트 대응"""
    nfc_pat = unicodedata.normalize("NFC", pattern)
    nfd_pat  = unicodedata.normalize("NFD", pattern)
    try:
        hits = sorted(
            f for f in folder.iterdir()
            if f.is_file() and (
                unicodedata.normalize("NFC", f.name) == nfc_pat
                or unicodedata.normalize("NFD", f.name) == nfd_pat
                or __import__("fnmatch").fnmatch(
                    unicodedata.normalize("NFC", f.name), nfc_pat)
            )
        )
        return hits[0] if hits else None
    except Exception:
        return None

def _subfile(folder, subdir):
    nfc_sub = unicodedata.normalize("NFC", subdir)
    # 서브디렉토리도 NFC/NFD 양방향 탐색
    d = None
    try:
        for item in folder.iterdir():
            if item.is_dir() and unicodedata.normalize("NFC", item.name) == nfc_sub:
                d = item
                break
    except Exception:
        pass
    if d is None:
        d = folder / subdir
    if not d.is_dir():
        return None
    files = [f for f in d.iterdir() if f.is_file()]
    return files[0] if files else None

COLUMNS_DEF = [
    # key               헤더            체크함수                                        중요도
    ("안내문",          "안내문",        lambda f,n: _g(f, f"종소세안내문_{n}.pdf"),     "mid"),
    ("전기신고내역",    "전기내역",      lambda f,n: _g(f, "전년도종소세신고내역.xls*"), "low"),
    ("전기신고서",      "전기신고서",    lambda f,n: _g(f, "2025*신고서*.pdf"),          "low"),
    ("지급명세서",      "지급명세서",    lambda f,n: _subfile(f, "지급명세서"),          "mid"),
    ("작업판",          "작업판",        lambda f,n: _g(f, f"작업판_{n}.xlsx"),          "mid"),
    ("작업결과",        "작업결과",      lambda f,n: _g(f, f"작업결과_{n}.xlsx"),        "mid"),
    ("당기신고서",      "당기신고서",    lambda f,n: _g(f, "신고서.pdf"),                "mid"),
    ("검증보고서",      "검증보고서",    lambda f,n: _g(f, "검증보고서_*.html"),         "mid"),
    ("출력패키지",      "출력패키지",    lambda f,n: _g(f, "출력패키지_*.pdf"),          "mid"),
    ("접수증",          "접수증",        lambda f,n: _g(f, f"종합소득세 접수증 {n}.pdf"),"high"),
    ("스크래핑신고서",  "스크래핑신고서",lambda f,n: _g(f, f"종합소득세 신고서 {n}.pdf"),"high"),
    ("소득세납부서",    "소득납부서",    lambda f,n: _g(f, f"종합소득세 납부서 {n}.pdf"),"high"),
    ("지방세납부서",    "지방납부서",    lambda f,n: _g(f, f"지방소득세 납부서 {n}.pdf"),"high"),
    ("랜딩HTML",        "랜딩",          lambda f,n: _g(f, f"신고결과_{n}.html"),        "high"),
]

HEADER_ROW = ["순번", "고객명", "주민앞6"] + [c[1] for c in COLUMNS_DEF] + ["갱신시각"]

# ── 색상 ─────────────────────────────────────────────────────────────────
COLOR_GREEN  = {"red": 0.851, "green": 0.918, "blue": 0.827}  # #d9ead3
COLOR_RED    = {"red": 0.957, "green": 0.800, "blue": 0.800}  # #f4cccc
COLOR_ORANGE = {"red": 0.988, "green": 0.914, "blue": 0.702}  # #fce5cd
COLOR_GREY   = {"red": 0.95,  "green": 0.95,  "blue": 0.95}
COLOR_WHITE  = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
COLOR_HEADER = {"red": 0.118, "green": 0.227, "blue": 0.369}  # #1e3a5f (네이비)


# ── 인증 ─────────────────────────────────────────────────────────────────

def get_credentials():
    creds = None
    if TOKEN_FILE.exists():
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CLIENT_SECRET), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return creds


def get_worksheet(spreadsheet_id=SPREADSHEET_ID):
    gc = gspread.authorize(get_credentials())
    sh = gc.open_by_key(spreadsheet_id)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=500, cols=len(HEADER_ROW) + 2)
    return ws


# ── 폴더 스캔 ─────────────────────────────────────────────────────────────

def _nfc(s):
    return unicodedata.normalize("NFC", s)

def _fmt(path):
    if path is None:
        return ""
    try:
        return datetime.fromtimestamp(path.stat().st_mtime).strftime("%m/%d")
    except Exception:
        return "✓"

def scan_folders():
    """고객 폴더 스캔 → [(folder, name, jumin6, {key: Path|None})]"""
    results = []
    folders = sorted(
        [p for p in CUSTOMER_DIR.iterdir()
         if p.is_dir() and not p.name.startswith("_")],
        key=lambda p: _nfc(p.name)
    )
    for folder in folders:
        parts  = _nfc(folder.name).split("_")
        name   = parts[0]
        jumin6 = parts[1] if len(parts) > 1 else ""
        checks = {}
        for key, _, checker, _ in COLUMNS_DEF:
            try:
                checks[key] = checker(folder, name)
            except Exception:
                checks[key] = None
        results.append((folder, name, jumin6, checks))
    return results


# ── 시트 쓰기 ─────────────────────────────────────────────────────────────

def build_requests(ws_id, rows_data):
    """색상 포맷 batch_update 요청 목록 생성"""
    requests = []
    n_cols = len(HEADER_ROW)

    # 헤더 행 배경색 (네이비)
    requests.append({
        "repeatCell": {
            "range": {
                "sheetId": ws_id, "startRowIndex": 0, "endRowIndex": 1,
                "startColumnIndex": 0, "endColumnIndex": n_cols,
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": COLOR_HEADER,
                    "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}},
                    "horizontalAlignment": "CENTER",
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
        }
    })

    # 데이터 행 색상
    for row_idx, (_, name, jumin6, checks) in enumerate(rows_data, start=1):
        col_offset = 3  # 순번, 고객명, 주민앞6 이후

        for col_j, (key, _, _, importance) in enumerate(COLUMNS_DEF):
            hit = checks.get(key)
            if hit is not None:
                bg = COLOR_GREEN
            else:
                if importance == "high":
                    bg = COLOR_RED
                elif importance == "mid":
                    bg = COLOR_ORANGE
                else:
                    bg = COLOR_GREY

            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": ws_id,
                        "startRowIndex": row_idx, "endRowIndex": row_idx + 1,
                        "startColumnIndex": col_offset + col_j,
                        "endColumnIndex": col_offset + col_j + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": bg,
                            "horizontalAlignment": "CENTER",
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)",
                }
            })

    return requests


def write_dashboard(ws, rows_data, generated, spreadsheet_id=SPREADSHEET_ID):
    print("  헤더 + 데이터 쓰는 중...")
    all_rows = [HEADER_ROW]

    for i, (folder, name, jumin6, checks) in enumerate(rows_data, start=1):
        row = [i, _nfc(name), jumin6]
        for key, _, _, _ in COLUMNS_DEF:
            row.append(_fmt(checks.get(key)))
        row.append(generated)
        all_rows.append(row)

    # 전체 클리어 후 쓰기
    ws.clear()
    ws.update(all_rows, "A1", value_input_option="USER_ENTERED")

    # 열 너비 자동 조정
    ws.columns_auto_resize(0, len(HEADER_ROW))

    print("  색상 포맷 적용 중...")
    requests = build_requests(ws.id, rows_data)

    # batch_update (100개씩 나눠서)
    gc = gspread.authorize(get_credentials())
    sh = gc.open_by_key(spreadsheet_id)
    for i in range(0, len(requests), 100):
        sh.batch_update({"requests": requests[i:i+100]})

    print(f"  완료: {len(rows_data)}명 기록")


# ── 메인 ─────────────────────────────────────────────────────────────────

def run():
    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"=== 현황대시보드 갱신 시작: {generated} ===")

    print("고객 폴더 스캔 중...")
    rows_data = scan_folders()
    print(f"  {len(rows_data)}명 발견")

    for sid, label in [
        (SPREADSHEET_ID,       "내부용"),
        (STAFF_SPREADSHEET_ID, "직원공유용"),
    ]:
        print(f"구글시트 연결 중... ({label})")
        ws = get_worksheet(sid)
        print(f"  시트: {SHEET_NAME} (id={ws.id})")
        write_dashboard(ws, rows_data, generated, spreadsheet_id=sid)
        print(f"  ✅ {label}: https://docs.google.com/spreadsheets/d/{sid}")


if __name__ == "__main__":
    run()
