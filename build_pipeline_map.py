"""
build_pipeline_map.py — 종소세 자동화 파이프라인 맵 구글시트 생성
세무회계창연 | 2026
"""
import os, sys, time
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))

import gspread
from gsheet_writer import get_credentials

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME     = "파이프라인맵"
N_COLS         = 6   # A~F

# ── 색상 ──────────────────────────────────────────────────────────
C_TITLE    = {"red": 0.071, "green": 0.141, "blue": 0.243}   # 짙은네이비
C_STAGE    = {"red": 0.173, "green": 0.341, "blue": 0.612}   # 단계헤더 파랑
C_AUTO     = {"red": 0.824, "green": 0.941, "blue": 0.800}   # 자동 연초록
C_BOT      = {"red": 0.784, "green": 0.878, "blue": 0.957}   # 봇 연파랑
C_MANUAL   = {"red": 1.000, "green": 0.949, "blue": 0.800}   # 수동 연노랑
C_DONE     = {"red": 0.851, "green": 0.918, "blue": 0.827}   # 완료
C_PENDING  = {"red": 0.957, "green": 0.800, "blue": 0.800}   # 미구현
C_ARROW    = {"red": 0.850, "green": 0.850, "blue": 0.850}   # 화살표행
C_LEGEND   = {"red": 0.953, "green": 0.953, "blue": 0.953}   # 범례
C_WHITE    = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
C_TEXT_W   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
C_TEXT_B   = {"red": 0.1,   "green": 0.1,   "blue": 0.1}

# ── 행 타입 정의 ──────────────────────────────────────────────────
# type: title / stage / arrow / item / legend / blank
# cols: [흐름, 스크립트/도구, 설명, 실행방식, 상태, 비고]

ROWS = [
    # ── 제목 ──────────────────────────────────────────────────────
    {"type": "title",
     "cols": ["종소세 자동화 파이프라인  |  세무회계창연 2026", "", "", "", "", ""]},

    {"type": "blank"},

    # ── 범례 ──────────────────────────────────────────────────────
    {"type": "legend",
     "cols": ["🟢 자동(크론/n8n)", "🔵 봇 명령어", "🟡 수동 실행", "🔴 미구현", "✅ 완료", ""]},

    {"type": "blank"},

    # ══════════════════════════════════════════════════════════════
    # A. 고객 등록
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["A.  고객 등록", "", "", "", "", ""]},

    {"type": "item", "mode": "auto",
     "cols": ["  ├─ A1  에어테이블", "마스터 DB에 고객 등록", "수동 입력", "완료", "", ""]},
    {"type": "item", "mode": "auto",
     "cols": ["  ├─ A2  airtable_sync_mac.py", "에어테이블 → 구글시트 접수명단 동기화", "자동 (1분 크론)", "완료", "", "맥미니"]},
    {"type": "item", "mode": "auto",
     "cols": ["  └─ A3  SOLAPI 알림톡", "수임동의 요청 카카오 발송", "자동 (n8n)", "완료", "", "솔라피"]},

    {"type": "arrow"},

    # ══════════════════════════════════════════════════════════════
    # B. 자료 수집
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["B.  자료 수집", "", "홈택스 Playwright", "엣지 디버그 필요", "", ""]},

    {"type": "item", "mode": "manual",
     "cols": ["  ├─ B1  기존고객처리.py", "안내문 PDF  ·  세무사 계정으로 주민번호 입력", "수동 실행", "완료", "", ""]},
    {"type": "item", "mode": "manual",
     "cols": ["  ├─ B2  신규고객처리.py", "안내문 PDF  ·  고객 아이디/비번 자동 로그인", "수동 실행", "완료", "", ""]},
    {"type": "item", "mode": "manual",
     "cols": ["  ├─ B3  jipgum_batch.py", "지급명세서 PDF 일괄 다운로드", "수동 실행", "완료", "", "run_7명.py 타겟 가능"]},
    {"type": "item", "mode": "manual",
     "cols": ["  ├─ B4  ganiyiyong_batch.py", "간이용역소득 xlsx 일괄 다운로드", "수동 실행", "완료", "", "run_7명.py 타겟 가능"]},
    {"type": "item", "mode": "manual",
     "cols": ["  └─ B5  종합소득세안내문조회.py", "전기신고내역 + 부가세 자료  (안내문 조회 시 함께)", "수동 실행", "완료", "", ""]},

    {"type": "arrow"},

    # ══════════════════════════════════════════════════════════════
    # C. 파싱 / 분석
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["C.  파싱 / 분석", "", "", "", "", ""]},

    {"type": "item", "mode": "auto",
     "cols": ["  ├─ C1  auto_parse.py", "안내문 PDF 감지 → 수입금액/기장의무 자동 파싱", "자동 (2분 크론)", "완료", "", "맥미니"]},
    {"type": "item", "mode": "auto",
     "cols": ["  ├─ C2  parse_to_xlsx.py", "parse_anneam()  ·  PDF 핵심 파싱 로직", "라이브러리", "완료", "", "1억↑ 줄바꿈 버그 수정"]},
    {"type": "item", "mode": "manual",
     "cols": ["  ├─ C3  reparse_high_income.py", "1억↑ 재파싱 + 변경분 작업판 자동 재생성", "수동 실행", "완료", "", ""]},
    {"type": "item", "mode": "auto",
     "cols": ["  └─ C4  gsheet_writer.py", "파싱결과 → 접수명단 구글시트 동기화", "라이브러리", "완료", "", ""]},

    {"type": "arrow"},

    # ══════════════════════════════════════════════════════════════
    # D. 작업판 생성
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["D.  작업판 생성", "", "", "", "", ""]},

    {"type": "item", "mode": "auto",
     "cols": ["  ├─ D1  jakupan_gen.py", "작업판_{이름}.xlsx 생성  ·  템플릿 자동 입력", "봇 /work  or  수동", "완료", "", ""]},
    {"type": "item", "mode": "bot",
     "cols": ["  └─ D2  봇  /work 이름", "작업판 + 자료 → 직원에게 텔레그램 전송", "봇 명령어", "완료", "", "jongsotaxbot"]},

    {"type": "arrow"},

    # ══════════════════════════════════════════════════════════════
    # E. 신고 작업 (직원)
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["E.  신고 작업  (직원)", "", "", "", "", ""]},

    {"type": "item", "mode": "bot",
     "cols": ["  ├─ E1  작업결과_{이름}.xls 업로드", "봇 자동 감지 → NAS 저장 + 담당자 알림", "봇 자동 감지", "완료", "", ""]},
    {"type": "item", "mode": "manual",
     "cols": ["  └─ E2  직원 작업판 작성", "세무사가 검토 후 작업결과 확정", "수동", "해당없음", "", ""]},

    {"type": "arrow"},

    # ══════════════════════════════════════════════════════════════
    # F. 검증 / 출력패키지
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["F.  검증 / 출력패키지", "", "", "", "", ""]},

    {"type": "item", "mode": "auto",
     "cols": ["  ├─ F1  tax_cross_verify.py", "교차검증 → 검증보고서_{날짜}.html", "봇 자동  or  수동", "완료", "", ""]},
    {"type": "item", "mode": "auto",
     "cols": ["  ├─ F2  print_package.py", "검증+소득+작업준비+안내문1p → 출력패키지 PDF", "봇 자동  or  /pkg", "완료", "", ""]},
    {"type": "item", "mode": "bot",
     "cols": ["  └─ F3  봇  신고서 업로드 자동", "25이름신고서.pdf 업로드 → 검증+출력패키지 자동", "봇 자동 감지", "완료", "", "jongsotaxbot"]},

    {"type": "arrow"},

    # ══════════════════════════════════════════════════════════════
    # G. 홈택스 신고 (세무사)
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["G.  홈택스 신고  (세무사 직접)", "", "", "", "", ""]},

    {"type": "item", "mode": "manual",
     "cols": ["  └─ G1  홈택스 직접 신고 제출", "자동화 대상 아님", "수동", "해당없음", "", ""]},

    {"type": "arrow"},

    # ══════════════════════════════════════════════════════════════
    # H. 신고결과 수집
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["H.  신고결과 수집", "", "홈택스 / 위택스 Playwright", "", "", ""]},

    {"type": "item", "mode": "manual",
     "cols": ["  ├─ H1  hometax_result_scraper.py", "접수증 + 신고서 + 납부서 PDF 스크래핑", "수동 실행", "완료", "", "홈택스"]},
    {"type": "item", "mode": "pending",
     "cols": ["  └─ H2  wetax_scraper.py", "지방소득세 납부서 PDF 스크래핑", "수동 실행", "미구현", "", "위택스  ·  예정"]},

    {"type": "arrow"},

    # ══════════════════════════════════════════════════════════════
    # I. 결과 안내 발송
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["I.  결과 안내 발송", "", "", "", "", ""]},

    {"type": "item", "mode": "auto",
     "cols": ["  ├─ I1  landing_gen.py", "신고결과 랜딩 HTML 생성", "자동 (FastAPI)", "완료", "", "taxeng.co.kr/jongsotax/"]},
    {"type": "item", "mode": "auto",
     "cols": ["  ├─ I2  landing_server.py", "FastAPI 8766  ·  n8n 요청 수신", "상시 실행", "완료", "", "맥미니"]},
    {"type": "item", "mode": "auto",
     "cols": ["  ├─ I3  알림톡  신고결과 안내", "홈택스 신고결과 확인 버튼 발송", "자동 (n8n)", "완료", "", "RDihHy11DR"]},
    {"type": "item", "mode": "auto",
     "cols": ["  └─ I4  알림톡  세액계산 안내", "신고 전 세액계산결과 확인 발송", "자동 (n8n)", "완료", "", "신고전세액계산"]},

    {"type": "arrow"},

    # ══════════════════════════════════════════════════════════════
    # J. 현황 모니터링
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["J.  현황 모니터링", "", "", "", "", ""]},

    {"type": "item", "mode": "auto",
     "cols": ["  ├─ J1  dashboard_gsheet.py", "고객폴더 14컬럼 현황 → 구글시트 갱신", "자동 (30분 크론)", "완료", "", "내부용 + 직원공유용"]},
    {"type": "item", "mode": "auto",
     "cols": ["  ├─ J2  dashboard_gen.py", "고객폴더 현황 → HTML 대시보드", "수동 실행", "완료", "", "NAS 직접 열람"]},
    {"type": "item", "mode": "pending",
     "cols": ["  └─ J3  파이프라인 실행 대시보드", "단계별 실행 버튼 + 현황 한눈에", "웹 or 봇", "예정", "", "이 문서 기반"]},

    {"type": "arrow"},

    # ══════════════════════════════════════════════════════════════
    # K. 개선 추천사항
    # ══════════════════════════════════════════════════════════════
    {"type": "stage",
     "cols": ["K.  개선 추천사항  (Claude 제안)", "", "", "", "", ""]},

    {"type": "item", "mode": "pending",
     "cols": ["  ├─ K1  B단계 자동화 (Playwright 상시화)",
              "엣지 디버그 의존 → 세무사 계정 크론 자동 수집으로 전환. 매일 새 접수 고객 자동 감지",
              "자동 크론", "예정", "★★★", "인력 투입 최대 절감"]},

    {"type": "item", "mode": "pending",
     "cols": ["  ├─ K2  H2 위택스 스크래핑 구현",
              "지방소득세 납부서 자동 수집 → 출력패키지 완성도 100%",
              "수동→자동", "미구현", "★★★", "현재 수작업 남은 유일 구간"]},

    {"type": "item", "mode": "pending",
     "cols": ["  ├─ K3  C단계 트리거 자동화",
              "새 PDF 감지 → 파싱 → 작업판 생성까지 end-to-end 자동 연결 (현재 별도 수동 실행)",
              "자동 (watchdog)", "예정", "★★☆", "작업판 누락 방지"]},

    {"type": "item", "mode": "pending",
     "cols": ["  ├─ K4  봇 /status 실시간 현황",
              "텔레그램 /status → 전체 고객 단계별 진행 현황 요약 즉시 응답",
              "봇 명령어", "예정", "★★☆", "현재 구글시트 직접 확인"]},

    {"type": "item", "mode": "pending",
     "cols": ["  ├─ K5  세액계산 알림톡 자동 발송",
              "작업결과 파일 업로드 감지 → 세액 추출 → SOLAPI 알림톡 자동 발송",
              "자동 (n8n)", "예정", "★★☆", "현재 수동 발송"]},

    {"type": "item", "mode": "pending",
     "cols": ["  ├─ K6  고객 랜딩페이지 배포 자동화",
              "신고결과 HTML → Cloudflare Tunnel 서빙 자동화 (현재 수동 nginx 설정 필요)",
              "자동 (FastAPI)", "예정", "★★☆", "landing_server.py 연동"]},

    {"type": "item", "mode": "pending",
     "cols": ["  └─ K7  연도별 고객 데이터 누적 구조",
              "2026→2027 이월 시 전년도 작업판/신고서를 자동 _archive 이동 + 신규 폴더 생성",
              "자동 (크론)", "예정", "★☆☆", "파일버전 관리 규칙 확장"]},

    {"type": "blank"},
]

# ── 헤더행 ────────────────────────────────────────────────────────
HEADER = ["흐름 / 모듈", "설명", "실행방식", "구현상태", "비고", ""]


def row_bg(r):
    t = r["type"]
    m = r.get("mode", "")
    if t == "title":   return C_TITLE
    if t == "stage":   return C_STAGE
    if t == "arrow":   return C_ARROW
    if t == "legend":  return C_LEGEND
    if t == "blank":   return C_WHITE
    if m == "auto":    return C_AUTO
    if m == "bot":     return C_BOT
    if m == "pending": return C_PENDING
    return C_MANUAL


def build_requests(ws_id, row_objects, data_rows):
    reqs = []

    # ── 제목 행 merge + 서식 ──────────────────────────────────────
    reqs.append({"mergeCells": {
        "range": {"sheetId": ws_id, "startRowIndex": 0, "endRowIndex": 1,
                  "startColumnIndex": 0, "endColumnIndex": N_COLS},
        "mergeType": "MERGE_ALL",
    }})
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws_id, "startRowIndex": 0, "endRowIndex": 1,
                  "startColumnIndex": 0, "endColumnIndex": N_COLS},
        "cell": {"userEnteredFormat": {
            "backgroundColor": C_TITLE,
            "textFormat": {"bold": True, "fontSize": 14,
                           "foregroundColor": C_TEXT_W},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
        }},
        "fields": "userEnteredFormat",
    }})
    reqs.append({"updateDimensionProperties": {
        "range": {"sheetId": ws_id, "dimension": "ROWS",
                  "startIndex": 0, "endIndex": 1},
        "properties": {"pixelSize": 40},
        "fields": "pixelSize",
    }})

    # ── 각 행 서식 ────────────────────────────────────────────────
    for i, r in enumerate(row_objects, start=1):
        bg   = row_bg(r)
        bold = r["type"] in ("stage", "title", "legend")
        fg   = C_TEXT_W if r["type"] in ("stage", "title") else C_TEXT_B

        reqs.append({"repeatCell": {
            "range": {"sheetId": ws_id, "startRowIndex": i, "endRowIndex": i+1,
                      "startColumnIndex": 0, "endColumnIndex": N_COLS},
            "cell": {"userEnteredFormat": {
                "backgroundColor": bg,
                "textFormat": {"bold": bold, "foregroundColor": fg,
                               "fontSize": 11 if r["type"] == "stage" else 10},
                "verticalAlignment": "MIDDLE",
            }},
            "fields": "userEnteredFormat",
        }})

        # stage 행 → 첫 컬럼 merge
        if r["type"] == "stage":
            reqs.append({"mergeCells": {
                "range": {"sheetId": ws_id, "startRowIndex": i, "endRowIndex": i+1,
                          "startColumnIndex": 0, "endColumnIndex": N_COLS},
                "mergeType": "MERGE_ALL",
            }})
            reqs.append({"repeatCell": {
                "range": {"sheetId": ws_id, "startRowIndex": i, "endRowIndex": i+1,
                          "startColumnIndex": 0, "endColumnIndex": N_COLS},
                "cell": {"userEnteredFormat": {"horizontalAlignment": "LEFT"}},
                "fields": "userEnteredFormat(horizontalAlignment)",
            }})

        # arrow 행 → 전체 merge + 중앙
        if r["type"] == "arrow":
            reqs.append({"mergeCells": {
                "range": {"sheetId": ws_id, "startRowIndex": i, "endRowIndex": i+1,
                          "startColumnIndex": 0, "endColumnIndex": N_COLS},
                "mergeType": "MERGE_ALL",
            }})
            reqs.append({"repeatCell": {
                "range": {"sheetId": ws_id, "startRowIndex": i, "endRowIndex": i+1,
                          "startColumnIndex": 0, "endColumnIndex": N_COLS},
                "cell": {"userEnteredFormat": {
                    "horizontalAlignment": "CENTER",
                    "textFormat": {"fontSize": 14, "foregroundColor":
                                   {"red":0.4,"green":0.4,"blue":0.4}},
                }},
                "fields": "userEnteredFormat",
            }})

        # 상태 컬럼(3) 색상
        status = r["cols"][3] if len(r.get("cols", [])) > 3 else ""
        if status == "완료":
            sc = C_DONE
        elif status in ("미구현", "예정"):
            sc = C_PENDING
        else:
            sc = bg
        if status and r["type"] == "item":
            reqs.append({"repeatCell": {
                "range": {"sheetId": ws_id, "startRowIndex": i, "endRowIndex": i+1,
                          "startColumnIndex": 3, "endColumnIndex": 4},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": sc,
                    "horizontalAlignment": "CENTER",
                    "textFormat": {"bold": True},
                }},
                "fields": "userEnteredFormat",
            }})

        # 행 높이
        h = 40 if r["type"] == "stage" else (10 if r["type"] in ("blank","arrow") else 26)
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": ws_id, "dimension": "ROWS",
                      "startIndex": i, "endIndex": i+1},
            "properties": {"pixelSize": h},
            "fields": "pixelSize",
        }})

    # ── 열 너비 ───────────────────────────────────────────────────
    widths = [280, 330, 140, 80, 200, 10]
    for ci, w in enumerate(widths):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": ws_id, "dimension": "COLUMNS",
                      "startIndex": ci, "endIndex": ci+1},
            "properties": {"pixelSize": w},
            "fields": "pixelSize",
        }})

    # 좌측 고정
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": ws_id,
                       "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount",
    }})

    return reqs


def main():
    creds = get_credentials()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SHEET_NAME)
        # 기존 merge 해제
        ws.clear()
        sh.batch_update({"requests": [{"unmergeCells": {
            "range": {"sheetId": ws.id,
                      "startRowIndex": 0, "endRowIndex": 200,
                      "startColumnIndex": 0, "endColumnIndex": N_COLS}
        }}]})
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=200, cols=N_COLS)

    # 데이터 조립
    data = [["종소세 자동화 파이프라인  |  세무회계창연 2026"] + [""]*(N_COLS-1)]
    for r in ROWS:
        cols = r.get("cols", ["","","","","",""])
        if r["type"] == "arrow":
            cols = ["↓"] + [""]*(N_COLS-1)
        data.append(cols[:N_COLS] + [""]*(N_COLS-len(cols)))

    print("  데이터 쓰는 중...")
    ws.update(data, "A1", value_input_option="USER_ENTERED")

    print("  서식 적용 중...")
    reqs = build_requests(ws.id, ROWS, data)
    for i in range(0, len(reqs), 100):
        sh.batch_update({"requests": reqs[i:i+100]})
        time.sleep(1)

    print(f"\n✅ 완료: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")


if __name__ == "__main__":
    main()
