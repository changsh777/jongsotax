"""
build_pipeline_map.py — 종소세 자동화 파이프라인 맵 구글시트 생성
세무회계창연 | 2026

실행: python3 ~/종소세2026/build_pipeline_map.py
→ 기존 스프레드시트에 "파이프라인맵" 탭 생성/갱신
"""
import os, sys, platform
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))

import gspread
from gsheet_writer import get_credentials

SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME     = "파이프라인맵"

# ── 색상 ──────────────────────────────────────────────────────────
C_HEADER  = {"red": 0.118, "green": 0.227, "blue": 0.369}   # 네이비
C_STAGE   = {"red": 0.263, "green": 0.471, "blue": 0.847}   # 파랑
C_AUTO    = {"red": 0.851, "green": 0.918, "blue": 0.827}   # 연초록 (자동)
C_MANUAL  = {"red": 0.988, "green": 0.914, "blue": 0.702}   # 연주황 (수동/반자동)
C_BOT     = {"red": 0.800, "green": 0.878, "blue": 0.957}   # 연파랑 (봇)
C_PENDING = {"red": 0.957, "green": 0.800, "blue": 0.800}   # 연빨강 (미구현)
C_WHITE   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
C_LGREY   = {"red": 0.95,  "green": 0.95,  "blue": 0.95}
C_TEXT_W  = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
C_TEXT_B  = {"red": 0.0,   "green": 0.0,   "blue": 0.0}

# ── 데이터 ────────────────────────────────────────────────────────
# (단계, 구분, 모듈/스크립트, 설명, 실행방식, 구현상태)
ROWS = [
    # 헤더
    ["단계", "구분", "모듈 / 스크립트", "설명", "실행방식", "구현상태", "비고"],

    # A. 고객 등록
    ["A. 고객 등록", "", "", "", "", "", ""],
    ["", "A1", "에어테이블 (마스터)", "고객 기본정보 등록", "수동", "완료", ""],
    ["", "A2", "airtable_sync_mac.py", "에어테이블 → 구글시트 접수명단 동기화", "자동 (1분 크론)", "완료", ""],
    ["", "A3", "SOLAPI 알림톡", "수임동의 요청 카카오 발송", "자동", "완료", "솔라피"],

    # B. 자료 수집
    ["B. 자료 수집", "", "", "", "", "", "홈택스 Playwright (엣지 디버그 필요)"],
    ["", "B1", "기존고객처리.py", "안내문 PDF — 세무사 계정으로 주민번호 입력", "수동 실행", "완료", ""],
    ["", "B2", "신규고객처리.py", "안내문 PDF — 고객 아이디/비번 자동 로그인", "수동 실행", "완료", ""],
    ["", "B3", "jipgum_batch.py", "지급명세서 PDF 일괄 다운로드", "수동 실행", "완료", "run_7명.py 타겟 실행 가능"],
    ["", "B4", "ganiyiyong_batch.py", "간이용역소득 xlsx 일괄 다운로드", "수동 실행", "완료", "run_7명.py 타겟 실행 가능"],
    ["", "B5", "종합소득세안내문조회.py", "전기신고내역 + 부가세 자료 (안내문 조회 시 함께)", "수동 실행", "완료", ""],

    # C. 파싱/분석
    ["C. 파싱 / 분석", "", "", "", "", "", ""],
    ["", "C1", "auto_parse.py", "안내문 PDF 감지 → 자동 파싱 (수입금액/기장의무)", "자동 (2분 크론)", "완료", "맥미니"],
    ["", "C2", "parse_to_xlsx.py", "parse_anneam() — PDF 텍스트 추출 핵심 로직", "라이브러리", "완료", "1억↑ 줄바꿈 버그 수정완료"],
    ["", "C3", "reparse_high_income.py", "1억↑ 수입금액 재파싱 + 변경분 작업판 재생성", "수동 실행", "완료", ""],
    ["", "C4", "gsheet_writer.py", "파싱결과 → 접수명단 구글시트 동기화", "라이브러리", "완료", ""],

    # D. 작업판 생성
    ["D. 작업판 생성", "", "", "", "", "", ""],
    ["", "D1", "jakupan_gen.py", "작업판_{이름}.xlsx 생성 (템플릿 자동 입력)", "봇 /work or 수동", "완료", ""],
    ["", "D2", "jongsotaxbot /work", "작업판 + 자료 직원에게 텔레그램 전송", "봇 명령어", "완료", ""],

    # E. 신고 작업
    ["E. 신고 작업 (직원)", "", "", "", "", "", ""],
    ["", "E1", "jongsotaxbot (파일 업로드)", "작업결과_{이름}.xls 업로드 → NAS 저장 + 알림", "봇 자동 감지", "완료", ""],

    # F. 검증/출력
    ["F. 검증 / 출력패키지", "", "", "", "", "", ""],
    ["", "F1", "tax_cross_verify.py", "교차검증 → 검증보고서_{날짜}.html 생성", "봇 자동 or 수동", "완료", ""],
    ["", "F2", "print_package.py", "검증+소득+작업준비+안내문1p → 출력패키지 PDF", "봇 자동 or /pkg", "완료", ""],
    ["", "F3", "jongsotaxbot (신고서 업로드)", "25이름신고서.pdf → 검증 + 출력패키지 자동 실행", "봇 자동 감지", "완료", ""],

    # G. 홈택스 신고
    ["G. 홈택스 신고 (세무사)", "", "", "", "", "", ""],
    ["", "G1", "홈택스 직접 신고", "세무사가 홈택스에서 직접 제출", "수동", "해당없음", "자동화 대상 아님"],

    # H. 신고결과 수집
    ["H. 신고결과 수집", "", "", "", "", "", "홈택스/위택스 Playwright"],
    ["", "H1", "hometax_result_scraper.py", "접수증 + 신고서 + 납부서 PDF 스크래핑", "수동 실행", "완료", ""],
    ["", "H2", "wetax_scraper.py", "지방소득세 납부서 PDF (위택스)", "수동 실행", "미구현", "예정"],

    # I. 결과 안내
    ["I. 결과 안내 발송", "", "", "", "", "", ""],
    ["", "I1", "landing_gen.py", "신고결과 랜딩 HTML 생성", "자동 (FastAPI)", "완료", "taxeng.co.kr/jongsotax/"],
    ["", "I2", "landing_server.py", "FastAPI 8766 — n8n에서 랜딩 생성 요청 수신", "상시 실행", "완료", "맥미니"],
    ["", "I3", "SOLAPI 알림톡 (신고결과)", "카카오 알림톡 발송 — 홈택스 신고결과 확인 버튼", "자동 (n8n)", "완료", "템플릿 RDihHy11DR"],
    ["", "I4", "SOLAPI 알림톡 (세액안내)", "신고 전 세액계산결과 확인 알림톡", "자동 (n8n)", "완료", "템플릿 신고전세액계산"],

    # J. 모니터링
    ["J. 현황 모니터링", "", "", "", "", "", ""],
    ["", "J1", "dashboard_gsheet.py", "고객폴더 현황 → 구글시트 현황대시보드 (14컬럼)", "자동 (30분 크론)", "완료", "내부용 + 직원공유용"],
    ["", "J2", "dashboard_gen.py", "고객폴더 현황 → HTML 대시보드 (_대시보드.html)", "수동 실행", "완료", "NAS 직접 열람"],

    # 범례
    ["", "", "", "", "", "", ""],
    ["범례", "자동 (크론/n8n)", "봇 명령어/자동감지", "수동 실행", "미구현", "", ""],
]

COLOR_MAP = {
    "완료":     C_AUTO,
    "미구현":   C_PENDING,
    "해당없음": C_LGREY,
}

STAGE_COLOR = {"red": 0.851, "green": 0.851, "blue": 0.918}  # 연보라 (단계 헤더)


def cell_color(row):
    if row[0] and row[0] != "범례":
        return STAGE_COLOR
    status = row[5] if len(row) > 5 else ""
    run_type = row[4] if len(row) > 4 else ""
    if not row[1]:
        return C_WHITE
    if "자동" in run_type:
        return C_AUTO
    if "봇" in run_type:
        return C_BOT
    return C_MANUAL


def build_requests(ws_id, data_rows):
    reqs = []
    n_cols = 7

    # 헤더
    reqs.append({"repeatCell": {
        "range": {"sheetId": ws_id, "startRowIndex": 0, "endRowIndex": 1,
                  "startColumnIndex": 0, "endColumnIndex": n_cols},
        "cell": {"userEnteredFormat": {
            "backgroundColor": C_HEADER,
            "textFormat": {"bold": True, "foregroundColor": C_TEXT_W},
            "horizontalAlignment": "CENTER",
        }},
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
    }})

    for i, row in enumerate(data_rows[1:], start=1):
        bg = cell_color(row)

        # 상태 컬럼(5) 색상 오버라이드
        status = row[5] if len(row) > 5 else ""
        status_bg = COLOR_MAP.get(status, bg)

        # 행 전체
        reqs.append({"repeatCell": {
            "range": {"sheetId": ws_id, "startRowIndex": i, "endRowIndex": i+1,
                      "startColumnIndex": 0, "endColumnIndex": n_cols},
            "cell": {"userEnteredFormat": {
                "backgroundColor": bg,
                "textFormat": {"bold": bool(row[0] and row[0] != "범례")},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }})

        # 상태 컬럼만 별도 색상
        if status:
            reqs.append({"repeatCell": {
                "range": {"sheetId": ws_id, "startRowIndex": i, "endRowIndex": i+1,
                          "startColumnIndex": 5, "endColumnIndex": 6},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": status_bg,
                    "horizontalAlignment": "CENTER",
                    "textFormat": {"bold": True},
                }},
                "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
            }})

    # 열 너비 고정
    widths = [120, 50, 220, 320, 150, 80, 200]
    for ci, w in enumerate(widths):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": ws_id, "dimension": "COLUMNS",
                      "startIndex": ci, "endIndex": ci+1},
            "properties": {"pixelSize": w},
            "fields": "pixelSize",
        }})

    return reqs


def main():
    creds = get_credentials()
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet(SHEET_NAME)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHEET_NAME, rows=60, cols=7)

    print(f"  시트 '{SHEET_NAME}' 데이터 쓰는 중...")
    ws.update(ROWS, "A1", value_input_option="USER_ENTERED")

    print("  서식 적용 중...")
    reqs = build_requests(ws.id, ROWS)
    for i in range(0, len(reqs), 100):
        sh.batch_update({"requests": reqs[i:i+100]})

    print(f"✅ 완료: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")


if __name__ == "__main__":
    main()
