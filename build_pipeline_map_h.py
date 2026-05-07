"""
build_pipeline_map_h.py — 종소세 자동화 파이프라인 가로맵
A→B→C→...→J 단계를 왼쪽→오른쪽으로 시각화 + K 개선추천사항
세무회계창연 | 2026
"""
import os, sys, time
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))

import gspread
from gsheet_writer import get_credentials

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME     = "파이프라인가로맵"

# ── 색상 ──────────────────────────────────────────────────────────
C_TITLE     = {"red": 0.071, "green": 0.141, "blue": 0.243}   # 짙은 네이비
C_STAGE     = {"red": 0.173, "green": 0.341, "blue": 0.612}   # 단계 헤더 파랑
C_STAGE_SUB = {"red": 0.631, "green": 0.741, "blue": 0.898}   # 헤더 서브 라이트블루
C_AUTO      = {"red": 0.824, "green": 0.941, "blue": 0.800}   # 자동  연초록
C_BOT       = {"red": 0.784, "green": 0.878, "blue": 0.957}   # 봇    연파랑
C_MANUAL    = {"red": 1.000, "green": 0.949, "blue": 0.800}   # 수동  연노랑
C_PENDING   = {"red": 0.957, "green": 0.800, "blue": 0.800}   # 미구현 연빨강
C_ARROW     = {"red": 0.878, "green": 0.878, "blue": 0.878}   # 화살표 셀 회색
C_LEGEND    = {"red": 0.949, "green": 0.949, "blue": 0.949}   # 범례 배경
C_K_HEADER  = {"red": 0.239, "green": 0.239, "blue": 0.239}   # K 섹션 헤더 다크그레이
C_WHITE     = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
C_TEXT_W    = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
C_TEXT_B    = {"red": 0.1,   "green": 0.1,   "blue": 0.1}
C_TEXT_GRAY = {"red": 0.35,  "green": 0.35,  "blue": 0.35}

MODE_BG = {"auto": C_AUTO, "bot": C_BOT, "manual": C_MANUAL, "pending": C_PENDING}

# ── 파이프라인 단계 A~J ────────────────────────────────────────────
STAGES = [
    {
        "name": "A.  고객 등록",
        "sub":  "에어테이블 → 접수명단",
        "items": [
            {"mode": "manual",  "t": "A1  에어테이블",       "d": "마스터 DB 고객 등록"},
            {"mode": "auto",    "t": "A2  airtable_sync",    "d": "에어테이블→접수명단  (1분 크론)"},
            {"mode": "auto",    "t": "A3  SOLAPI 알림톡",    "d": "수임동의 카카오 자동 발송"},
        ]
    },
    {
        "name": "B.  자료 수집",
        "sub":  "홈택스 Playwright  ·  엣지 디버그",
        "items": [
            {"mode": "manual",  "t": "B1  기존고객처리.py",   "d": "안내문 PDF  (세무사 계정)"},
            {"mode": "manual",  "t": "B2  신규고객처리.py",   "d": "안내문 PDF  (고객 로그인)"},
            {"mode": "manual",  "t": "B3  jipgum_batch",      "d": "지급명세서 PDF 일괄"},
            {"mode": "manual",  "t": "B4  ganiyiyong_batch",  "d": "간이용역소득 xlsx 일괄"},
            {"mode": "manual",  "t": "B5  안내문조회.py",      "d": "전기신고내역 + 부가세"},
        ]
    },
    {
        "name": "C.  파싱 / 분석",
        "sub":  "PDF → 구글시트 동기화",
        "items": [
            {"mode": "auto",    "t": "C1  auto_parse.py",    "d": "PDF 감지→파싱  (2분 크론)"},
            {"mode": "auto",    "t": "C2  parse_to_xlsx.py", "d": "수입금액·기장의무 핵심 파싱"},
            {"mode": "manual",  "t": "C3  reparse_high",     "d": "1억↑ 재파싱 + 작업판 재생성"},
            {"mode": "auto",    "t": "C4  gsheet_writer.py", "d": "파싱결과→접수명단 동기화"},
        ]
    },
    {
        "name": "D.  작업판 생성",
        "sub":  "jakupan_gen",
        "items": [
            {"mode": "auto",    "t": "D1  jakupan_gen.py",   "d": "작업판_{이름}.xlsx 자동 생성"},
            {"mode": "bot",     "t": "D2  봇  /work 이름",    "d": "작업판+자료 직원 텔레그램 전송"},
        ]
    },
    {
        "name": "E.  신고 작업",
        "sub":  "직원 담당",
        "items": [
            {"mode": "bot",     "t": "E1  작업결과 업로드",   "d": "봇 자동감지→NAS 저장+알림"},
            {"mode": "manual",  "t": "E2  직원 작업판 작성",  "d": "세무사 검토 후 확정"},
        ]
    },
    {
        "name": "F.  검증 / 출력",
        "sub":  "교차검증  ·  PDF 패키지",
        "items": [
            {"mode": "auto",    "t": "F1  tax_cross_verify", "d": "교차검증→검증보고서.html"},
            {"mode": "auto",    "t": "F2  print_package.py", "d": "검증+소득+작업준비→PDF"},
            {"mode": "bot",     "t": "F3  봇  신고서 업로드", "d": "25이름신고서.pdf→자동처리"},
        ]
    },
    {
        "name": "G.  홈택스 신고",
        "sub":  "세무사 직접 제출",
        "items": [
            {"mode": "manual",  "t": "G1  홈택스 직접 신고",  "d": "자동화 대상 아님"},
        ]
    },
    {
        "name": "H.  신고결과 수집",
        "sub":  "홈택스·위택스 스크래핑",
        "items": [
            {"mode": "manual",  "t": "H1  hometax_result",   "d": "접수증·신고서·납부서 PDF"},
            {"mode": "pending", "t": "H2  wetax_scraper",    "d": "지방소득세 납부서  (미구현)"},
        ]
    },
    {
        "name": "I.  결과 안내 발송",
        "sub":  "FastAPI  ·  SOLAPI  ·  n8n",
        "items": [
            {"mode": "auto",    "t": "I1  landing_gen.py",   "d": "신고결과 랜딩 HTML 생성"},
            {"mode": "auto",    "t": "I2  landing_server",   "d": "FastAPI 8766  (맥미니 상시)"},
            {"mode": "auto",    "t": "I3  알림톡  신고결과",  "d": "홈택스 확인링크 카카오 발송"},
            {"mode": "auto",    "t": "I4  알림톡  세액계산",  "d": "신고 전 세액결과 안내"},
        ]
    },
    {
        "name": "J.  모니터링",
        "sub":  "구글시트  ·  HTML 대시보드",
        "items": [
            {"mode": "auto",    "t": "J1  dashboard_gsheet", "d": "14컬럼 현황  (30분 크론)"},
            {"mode": "auto",    "t": "J2  dashboard_gen.py", "d": "NAS HTML 대시보드"},
            {"mode": "pending", "t": "J3  실행 대시보드",     "d": "단계별 실행버튼+현황  (예정)"},
        ]
    },
]

# ── K 개선 추천사항 ─────────────────────────────────────────────────
K_ITEMS = [
    {"t": "K1  B단계 자동화  ★★★",       "d": "엣지 디버그 의존 제거 → 크론 자동 수집  (인력 최대 절감)"},
    {"t": "K2  위택스 스크래핑  ★★★",     "d": "지방소득세 납부서 자동화  (유일한 수작업 구간 H2)"},
    {"t": "K3  파싱 트리거 자동화  ★★☆",  "d": "PDF감지→파싱→작업판 end-to-end  (누락 방지)"},
    {"t": "K4  봇  /status  ★★☆",         "d": "텔레그램 전체 고객 진행현황 즉시 응답"},
    {"t": "K5  세액알림톡 자동화  ★★☆",   "d": "작업결과 업로드 감지 → SOLAPI 자동 발송"},
    {"t": "K6  랜딩 배포 자동화  ★★☆",    "d": "Cloudflare Tunnel 서빙 완전 자동화"},
    {"t": "K7  연도별 이월 구조  ★☆☆",    "d": "2026→2027  아카이브+신규폴더 자동 생성"},
]

# ── 레이아웃 상수 ──────────────────────────────────────────────────
N_STAGES   = len(STAGES)          # 10
N_COLS     = N_STAGES * 2 - 1    # 19  (stage, arrow, stage, ..., stage)
MAX_ITEMS  = max(len(s["items"]) for s in STAGES)  # 5

# 행 인덱스 (0-based)
R_TITLE    = 0
R_BLANK1   = 1
R_LEGEND   = 2
R_BLANK2   = 3
R_HEADER   = 4   # 단계 헤더 (파랑)
R_SUB      = 5   # 단계 서브타이틀
R_SEP      = 6   # 구분선
R_ITEM0    = 7   # 첫 번째 아이템 행
# 아이템 행: R_ITEM0 .. R_ITEM0+MAX_ITEMS-1  (7..11)
R_BLANK3   = R_ITEM0 + MAX_ITEMS          # 12
R_K_HEAD   = R_BLANK3 + 1                # 13
R_K_SUB    = R_K_HEAD + 1                # 14
R_K_ITEM0  = R_K_SUB + 1                 # 15  (K 아이템 행 시작)
N_K_ROWS   = 3                           # ceil(7/3) = 3행
R_BLANK4   = R_K_ITEM0 + N_K_ROWS        # 18
N_ROWS     = R_BLANK4 + 1               # 19

# K 아이템 열 그룹 (col start, col end exclusive)
K_GROUPS = [(0, 6), (6, 12), (12, N_COLS)]

# 열 너비 (짝수=단계, 홀수=화살표)
STAGE_W = 210
ARROW_W = 30


# ── 헬퍼 함수 ─────────────────────────────────────────────────────
def _range(ws_id, r1, r2, c1, c2):
    return {"sheetId": ws_id,
            "startRowIndex": r1, "endRowIndex": r2,
            "startColumnIndex": c1, "endColumnIndex": c2}

def _merge(ws_id, r1, r2, c1, c2):
    return {"mergeCells": {"range": _range(ws_id, r1, r2, c1, c2),
                           "mergeType": "MERGE_ALL"}}

def _fmt(ws_id, r1, r2, c1, c2,
         bg=C_WHITE, fg=C_TEXT_B, bold=False, size=10,
         halign="LEFT", valign="MIDDLE", wrap=False, italic=False):
    tf = {"bold": bold, "fontSize": size, "foregroundColor": fg}
    if italic:
        tf["italic"] = True
    fmt = {"backgroundColor": bg,
           "textFormat": tf,
           "horizontalAlignment": halign,
           "verticalAlignment": valign}
    if wrap:
        fmt["wrapStrategy"] = "WRAP"
    return {"repeatCell": {
        "range": _range(ws_id, r1, r2, c1, c2),
        "cell": {"userEnteredFormat": fmt},
        "fields": "userEnteredFormat",
    }}

def _row_h(ws_id, row, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": ws_id, "dimension": "ROWS",
                  "startIndex": row, "endIndex": row + 1},
        "properties": {"pixelSize": px}, "fields": "pixelSize"}}

def _col_w(ws_id, col, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": ws_id, "dimension": "COLUMNS",
                  "startIndex": col, "endIndex": col + 1},
        "properties": {"pixelSize": px}, "fields": "pixelSize"}}


# ── 데이터 조립 ────────────────────────────────────────────────────
def build_data():
    data = [[""] * N_COLS for _ in range(N_ROWS)]

    # 제목
    data[R_TITLE][0] = "종소세 자동화 파이프라인  가로맵  |  세무회계창연 2026"

    # 범례
    data[R_LEGEND][0] = (
        "  🟢 자동 (크론/n8n)      🔵 봇 명령어      "
        "🟡 수동 실행      🔴 미구현/예정"
    )

    # 단계 헤더 + 서브타이틀 + 화살표
    for s_idx, stage in enumerate(STAGES):
        col = s_idx * 2
        data[R_HEADER][col] = stage["name"]
        data[R_SUB][col]    = "  " + stage["sub"]
    for a_idx in range(N_STAGES - 1):
        data[R_HEADER][a_idx * 2 + 1] = "→"

    # 아이템
    for s_idx, stage in enumerate(STAGES):
        col = s_idx * 2
        for i, item in enumerate(stage["items"]):
            data[R_ITEM0 + i][col] = f"  {item['t']}\n  {item['d']}"

    # K 헤더 + 서브
    data[R_K_HEAD][0] = "  K.  개선 추천사항  —  Claude 제안"
    data[R_K_SUB][0]  = "  우선순위:  ★★★ 즉시 착수   ★★☆ 이번 시즌 후   ★☆☆ 다음 시즌"

    # K 아이템 (3열 × 3행)
    for k_idx, k in enumerate(K_ITEMS):
        row   = R_K_ITEM0 + k_idx // 3
        g_col = K_GROUPS[k_idx % 3][0]
        data[row][g_col] = f"  {k['t']}\n  {k['d']}"

    return data


# ── 서식 요청 조립 ─────────────────────────────────────────────────
def build_requests(ws_id):
    R = []

    # ── 제목 ─────────────────────────────────
    R.append(_merge(ws_id, R_TITLE, R_TITLE+1, 0, N_COLS))
    R.append(_fmt(ws_id, R_TITLE, R_TITLE+1, 0, N_COLS,
                  bg=C_TITLE, fg=C_TEXT_W, bold=True, size=15, halign="CENTER"))
    R.append(_row_h(ws_id, R_TITLE, 44))

    # ── 빈행 ─────────────────────────────────
    for row, px in [(R_BLANK1,5),(R_BLANK2,5),(R_SEP,4),(R_BLANK3,12),(R_BLANK4,12)]:
        R.append(_fmt(ws_id, row, row+1, 0, N_COLS, bg=C_WHITE))
        R.append(_row_h(ws_id, row, px))

    # ── 범례 ─────────────────────────────────
    R.append(_merge(ws_id, R_LEGEND, R_LEGEND+1, 0, N_COLS))
    R.append(_fmt(ws_id, R_LEGEND, R_LEGEND+1, 0, N_COLS,
                  bg=C_LEGEND, fg=C_TEXT_GRAY, bold=False, size=9, halign="LEFT"))
    R.append(_row_h(ws_id, R_LEGEND, 22))

    # ── 단계 헤더행 (R_HEADER) ────────────────
    R.append(_row_h(ws_id, R_HEADER, 40))
    for s_idx in range(N_STAGES):
        col = s_idx * 2
        R.append(_fmt(ws_id, R_HEADER, R_HEADER+1, col, col+1,
                      bg=C_STAGE, fg=C_TEXT_W, bold=True, size=11, halign="LEFT", valign="MIDDLE"))

    # ── 화살표 열: 단계헤더~마지막아이템 까지 세로 병합 ─
    arrow_row_end = R_ITEM0 + MAX_ITEMS   # 12
    for a_idx in range(N_STAGES - 1):
        col = a_idx * 2 + 1
        R.append(_merge(ws_id, R_HEADER, arrow_row_end, col, col+1))
        R.append(_fmt(ws_id, R_HEADER, arrow_row_end, col, col+1,
                      bg=C_ARROW, fg=C_TEXT_GRAY, bold=True, size=18,
                      halign="CENTER", valign="MIDDLE"))

    # ── 단계 서브타이틀행 (R_SUB) ─────────────
    R.append(_row_h(ws_id, R_SUB, 18))
    for s_idx in range(N_STAGES):
        col = s_idx * 2
        R.append(_fmt(ws_id, R_SUB, R_SUB+1, col, col+1,
                      bg=C_STAGE_SUB, fg=C_TEXT_B, bold=False, size=8,
                      halign="LEFT", valign="MIDDLE", italic=True))

    # ── 아이템 행 ─────────────────────────────
    for i in range(MAX_ITEMS):
        row = R_ITEM0 + i
        R.append(_row_h(ws_id, row, 34))
        for s_idx, stage in enumerate(STAGES):
            col = s_idx * 2
            if i < len(stage["items"]):
                mode = stage["items"][i]["mode"]
                bg   = MODE_BG[mode]
            else:
                bg = C_WHITE
            R.append(_fmt(ws_id, row, row+1, col, col+1,
                          bg=bg, fg=C_TEXT_B, bold=False, size=9,
                          halign="LEFT", valign="TOP", wrap=True))

    # ── K 헤더 ───────────────────────────────
    R.append(_merge(ws_id, R_K_HEAD, R_K_HEAD+1, 0, N_COLS))
    R.append(_fmt(ws_id, R_K_HEAD, R_K_HEAD+1, 0, N_COLS,
                  bg=C_K_HEADER, fg=C_TEXT_W, bold=True, size=11, halign="LEFT"))
    R.append(_row_h(ws_id, R_K_HEAD, 40))

    # ── K 서브타이틀 ─────────────────────────
    R.append(_merge(ws_id, R_K_SUB, R_K_SUB+1, 0, N_COLS))
    R.append(_fmt(ws_id, R_K_SUB, R_K_SUB+1, 0, N_COLS,
                  bg=C_LEGEND, fg=C_TEXT_GRAY, bold=False, size=9, halign="LEFT"))
    R.append(_row_h(ws_id, R_K_SUB, 20))

    # ── K 아이템 행 (3행 × 3열) ───────────────
    for k_row in range(N_K_ROWS):
        row = R_K_ITEM0 + k_row
        R.append(_row_h(ws_id, row, 46))
        for g_idx, (c1, c2) in enumerate(K_GROUPS):
            k_idx = k_row * 3 + g_idx
            if k_idx < len(K_ITEMS):
                bg = C_PENDING
            else:
                bg = C_WHITE
            R.append(_merge(ws_id, row, row+1, c1, c2))
            R.append(_fmt(ws_id, row, row+1, c1, c2,
                          bg=bg, fg=C_TEXT_B, bold=False, size=9,
                          halign="LEFT", valign="MIDDLE", wrap=True))

    # ── 열 너비 ──────────────────────────────
    for s_idx in range(N_STAGES):
        R.append(_col_w(ws_id, s_idx * 2, STAGE_W))
    for a_idx in range(N_STAGES - 1):
        R.append(_col_w(ws_id, a_idx * 2 + 1, ARROW_W))

    # ── 1행 고정 (스크롤 시 제목 유지) ──────
    R.append({"updateSheetProperties": {
        "properties": {"sheetId": ws_id,
                       "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount",
    }})

    return R


# ── 메인 ──────────────────────────────────────────────────────────
def main():
    print("  Google Sheets 연결 중...")
    creds = get_credentials()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(SPREADSHEET_ID)

    # 시트 준비
    try:
        ws = sh.worksheet(SHEET_NAME)
        ws.clear()
        sh.batch_update({"requests": [{"unmergeCells": {
            "range": {"sheetId": ws.id,
                      "startRowIndex": 0, "endRowIndex": 200,
                      "startColumnIndex": 0, "endColumnIndex": N_COLS + 2}
        }}]})
        print(f"  기존 시트 초기화 완료")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=N_ROWS + 5, cols=N_COLS + 2)
        print(f"  새 시트 생성: {SHEET_NAME}")

    # 데이터 쓰기
    data = build_data()
    print(f"  데이터 쓰는 중 ({N_ROWS}행 × {N_COLS}열)...")
    ws.update(data, "A1", value_input_option="USER_ENTERED")

    # 서식 적용
    print(f"  서식 적용 중...")
    reqs = build_requests(ws.id)
    for i in range(0, len(reqs), 100):
        sh.batch_update({"requests": reqs[i:i+100]})
        if i + 100 < len(reqs):
            time.sleep(0.5)

    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
    print(f"\n✅ 완료: {url}")
    print(f"   시트: '{SHEET_NAME}'")


if __name__ == "__main__":
    main()
