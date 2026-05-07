"""
build_lego_sheet.py — 파이프라인 레고블럭 현황 → 구글시트 "레고블럭" 탭
세무회계창연 | 2026

실행:
    python F:\종소세2026\build_lego_sheet.py
    python ~/종소세2026/build_lego_sheet.py
"""
import sys, io, os, time
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
elif hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))

import pickle
from pathlib import Path
from datetime import datetime

import gspread
from google.auth.transport.requests import Request

from pipeline_runner import PIPELINE, load_all_customers, find_folder, get_status

# ── 인증 경로 ─────────────────────────────────────────────────────────
import platform
if platform.system() == "Darwin":
    CRED_DIR = Path.home() / "종소세2026/.credentials"
else:
    CRED_DIR = Path(r"F:\종소세2026\.credentials")

TOKEN_FILE     = CRED_DIR / "token.pickle"
SPREADSHEET_ID       = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
STAFF_SPREADSHEET_ID = "1ht7fk381ei8fJ33KpigeMSaEdjlOifVFzgN2Dyy-dpk"
SHEET_NAME     = "레고블럭"

# ── 색상 ─────────────────────────────────────────────────────────────
C_HEADER  = {"red": 0.118, "green": 0.227, "blue": 0.369}   # 네이비
C_GREEN   = {"red": 0.851, "green": 0.918, "blue": 0.827}   # done
C_GREY    = {"red": 0.90,  "green": 0.90,  "blue": 0.90}    # blocked
C_ORANGE  = {"red": 0.988, "green": 0.914, "blue": 0.702}   # waiting/manual
C_RED     = {"red": 0.957, "green": 0.800, "blue": 0.800}   # pending
C_WHITE   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}

# 상태 → (기호, 배경색)
STATUS_STYLE = {
    "done":    ("✓",  C_GREEN),
    "blocked": ("·",  C_GREY),
    "waiting": ("대기", C_ORANGE),
    "pending": ("✗",  C_RED),
}

# ── 인증 ─────────────────────────────────────────────────────────────
def get_credentials():
    with open(TOKEN_FILE, "rb") as f:
        creds = pickle.load(f)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return creds

def get_worksheet(spreadsheet_id):
    gc = gspread.authorize(get_credentials())
    sh = gc.open_by_key(spreadsheet_id)
    try:
        ws = sh.worksheet(SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=600, cols=len(PIPELINE) + 5)
    return ws

# ── 데이터 수집 ──────────────────────────────────────────────────────
def collect_data():
    customers = load_all_customers()
    print(f"  고객 {len(customers)}명 스캔 중...")
    rows = []
    for i, c in enumerate(customers, 1):
        name, jumin6 = c["name"], c["jumin6"]
        folder = find_folder(name, jumin6)
        statuses = {}
        for mod in PIPELINE:
            statuses[mod.id] = get_status(mod, folder, name)
        rows.append((name, jumin6, statuses))
        if i % 50 == 0:
            print(f"    {i}/{len(customers)}...", flush=True)
    return rows

# ── 시트 쓰기 ─────────────────────────────────────────────────────────
def write_sheet(ws, rows, generated, spreadsheet_id):
    # 헤더
    header = ["순번", "고객명", "주민앞6"] + [m.label for m in PIPELINE] + ["갱신시각"]
    n_cols = len(header)

    # 데이터 행 (값)
    all_rows = [header]
    for i, (name, jumin6, statuses) in enumerate(rows, 1):
        row = [i, name, jumin6]
        for mod in PIPELINE:
            sym, _ = STATUS_STYLE.get(statuses[mod.id], ("?", C_WHITE))
            row.append(sym)
        row.append(generated)
        all_rows.append(row)

    print("  값 쓰는 중...")
    ws.clear()
    ws.update(all_rows, "A1", value_input_option="USER_ENTERED")
    ws.columns_auto_resize(0, n_cols)

    # 색상 배치 요청
    print("  색상 적용 중...")
    requests = []

    # 헤더 행
    requests.append({
        "repeatCell": {
            "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": n_cols},
            "cell": {"userEnteredFormat": {
                "backgroundColor": C_HEADER,
                "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}},
                "horizontalAlignment": "CENTER",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
        }
    })

    # 데이터 행 색상 (파이프라인 컬럼만)
    col_offset = 3  # 순번, 고객명, 주민앞6
    for row_idx, (name, jumin6, statuses) in enumerate(rows, start=1):
        for col_j, mod in enumerate(PIPELINE):
            st = statuses[mod.id]
            _, bg = STATUS_STYLE.get(st, ("?", C_WHITE))
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": ws.id,
                        "startRowIndex": row_idx, "endRowIndex": row_idx + 1,
                        "startColumnIndex": col_offset + col_j,
                        "endColumnIndex": col_offset + col_j + 1,
                    },
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": bg,
                        "horizontalAlignment": "CENTER",
                    }},
                    "fields": "userEnteredFormat(backgroundColor,horizontalAlignment)",
                }
            })

    # 100개씩 batch_update
    gc = gspread.authorize(get_credentials())
    sh = gc.open_by_key(spreadsheet_id)
    for i in range(0, len(requests), 100):
        sh.batch_update({"requests": requests[i:i+100]})

    print(f"  완료: {len(rows)}명 × {len(PIPELINE)}단계")

# ── 메인 ─────────────────────────────────────────────────────────────
def run():
    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"=== 레고블럭 시트 갱신: {generated} ===")

    rows = collect_data()

    for idx, (sid, label) in enumerate([
        (SPREADSHEET_ID,       "내부용"),
        (STAFF_SPREADSHEET_ID, "직원공유용"),
    ]):
        if idx > 0:
            print("  API 쿼타 대기 (65초)...")
            time.sleep(65)
        print(f"구글시트 연결 ({label})...")
        ws = get_worksheet(sid)
        write_sheet(ws, rows, generated, sid)
        print(f"  ✅ {label}: https://docs.google.com/spreadsheets/d/{sid}")

if __name__ == "__main__":
    run()
