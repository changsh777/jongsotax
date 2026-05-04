"""
airtable_sync_mac.py - 에어테이블 → 구글시트 단방향 자동 동기화 (Mac Mini 전용)

- 1분 크론으로 실행
- 2026-06-01부터 자동 종료 (종소세 시즌 끝)
- 에어테이블 = 마스터 / 구글시트 = 작업 사본 (덮어쓰기)
- 대원칙: 에어테이블 행번호(Grid view 순서) 기준
- 안내문파싱 탭 데이터(기장의무/추계시적용경비율/O/X 등)를 접수명단에 병합
"""
import sys, json, time, urllib.request, urllib.parse, pickle
from datetime import datetime, date
from pathlib import Path

import gspread
from google.auth.transport.requests import Request

# ── 설정 ─────────────────────────────────────────────
SEASON_END     = date(2026, 6, 1)
CRED_DIR       = Path.home() / "종소세2026/.credentials"
TOKEN_FILE     = CRED_DIR / "token.pickle"
AIRTABLE_PAT   = open(Path.home() / "종소세2026/.credentials/airtable_pat.txt").read().strip()
BASE_ID        = "appSvDTDOmYfBeIFs"
TABLE_ID       = "tbl2f2h6GfSnLCQpt"
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME     = "접수명단"
PARSE_SHEET    = "안내문파싱"
SKIP_TYPES     = {"multipleRecordLinks", "multipleAttachments", "multipleLookupValues"}

# 자동회신↔발송관리 사이에 삽입할 컬럼
PARSE_COLS     = ["타소득(O/X)", "이자", "배당", "근로(단일)", "근로(복수)", "연금", "기타", "기장의무", "추계시적용경비율"]
# 타소득(O/X) 파생에 쓸 개별 O/X 원본 컬럼
TAXINCOME_COLS = ["이자", "배당", "근로(단일)", "근로(복수)", "연금", "기타"]

# 구글시트 값이 항상 우선 (에어테이블 값이 있어도 구글시트 값 유지)
GOOGLE_WINS_COLS = {"입금체크"}


# ── 인증 ─────────────────────────────────────────────
def get_creds():
    with open(TOKEN_FILE, "rb") as f:
        creds = pickle.load(f)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return creds


# ── 에어테이블 ────────────────────────────────────────
def airtable_get(path):
    url = f"https://api.airtable.com/v0/{path}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {AIRTABLE_PAT}"})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read().decode())

def fetch_meta():
    """필드 목록 + 첫 번째 뷰(Grid view) ID"""
    data = airtable_get(f"meta/bases/{BASE_ID}/tables")
    for t in data.get("tables", []):
        if t["id"] == TABLE_ID:
            fields  = [(f["name"], f["type"]) for f in t.get("fields", [])]
            views   = t.get("views", [])
            view_id = views[0]["id"] if views else None
            view_nm = views[0].get("name", "?") if views else "없음"
            print(f"  뷰: {view_nm} ({view_id})")
            return fields, view_id
    return [], None

def fetch_all_records(view_id=None):
    """에어테이블 행번호(Grid view) 기준 레코드 전체"""
    records, offset = [], None
    while True:
        params = []
        if view_id:
            params.append(f"view={urllib.parse.quote(view_id)}")
        if offset:
            params.append(f"offset={urllib.parse.quote(offset)}")
        qs   = ("?" + "&".join(params)) if params else ""
        path = f"{BASE_ID}/{TABLE_ID}{qs}"
        data = airtable_get(path)
        records.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset:
            break
        time.sleep(0.2)
    return records

def cell_value(val, ftype):
    if val is None: return ""
    if ftype in SKIP_TYPES: return f"[{len(val)}건]" if isinstance(val, list) else str(val)
    if isinstance(val, bool): return "O" if val else ""
    if isinstance(val, (int, float)): return val
    if isinstance(val, dict): return val.get("name") or val.get("email") or str(val)
    return str(val)


# ── 안내문파싱 탭 읽기 ─────────────────────────────────
def fetch_parse_data(gc):
    """안내문파싱 탭 → 성명 기준 dict 반환 (O/X 컬럼 포함)"""
    try:
        ws = gc.open_by_key(SPREADSHEET_ID).worksheet(PARSE_SHEET)
        rows = ws.get_all_records()
        result = {}
        for r in rows:
            name = str(r.get("성명", "")).strip()
            if name:
                result[name] = r
        print(f"  안내문파싱: {len(result)}명 로드")
        return result
    except Exception as e:
        print(f"  안내문파싱 로드 실패: {e}")
        return {}


# ── 메인 ─────────────────────────────────────────────
def main():
    if date.today() >= SEASON_END:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 종소세 시즌 종료({SEASON_END}) - 싱크 중단")
        sys.exit(0)

    # 에어테이블 Grid view 순서로 레코드 가져오기
    fields, view_id = fetch_meta()
    records = fetch_all_records(view_id)

    # 안내문파싱 탭 데이터 (성명 기준)
    gc = gspread.authorize(get_creds())
    parse_data = fetch_parse_data(gc)

    normal = [(n, t) for n, t in fields if t not in SKIP_TYPES]
    linked = [(n, t) for n, t in fields if t in SKIP_TYPES]

    # 자동회신 바로 뒤 (발송관리 앞)에 안내문파싱 컬럼 삽입
    normal_names = [n for n, _ in normal]
    insert_after = next((i for i, n in enumerate(normal_names) if n == "자동회신"),
                        len(normal_names) - 1)  # 없으면 맨 뒤

    header = (
        normal_names[:insert_after + 1]
        + PARSE_COLS
        + normal_names[insert_after + 1:]
        + ["[링크]" + n for n, _ in linked]
        + ["_sync_at"]
    )

    # ── 기존 구글시트 값 먼저 읽기 (FORMULA 렌더: 수식셀은 "=..." 문자열로 옴) ───
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    existing_all = ws.get_all_values(value_render_option="FORMULA")
    existing_by_name = {}
    if existing_all:
        ex_header = existing_all[0]
        if "성명" in ex_header:
            name_col_idx = ex_header.index("성명")
            for ex_row in existing_all[1:]:
                if len(ex_row) > name_col_idx and str(ex_row[name_col_idx]).strip():
                    rname = str(ex_row[name_col_idx]).strip()
                    existing_by_name[rname] = {
                        ex_header[i]: ex_row[i]
                        for i in range(len(ex_header))
                        if i < len(ex_row)
                    }

    now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [header]
    for r in records:
        name = str(r["fields"].get("성명", "")).strip()
        pd   = parse_data.get(name, {})
        ex   = existing_by_name.get(name, {})

        at_vals = [cell_value(r["fields"].get(n), t) for n, t in normal]

        # 타소득(O/X): 안내문파싱에 데이터 있을 때만 계산 (없으면 "" → 기존값 유지)
        if name in parse_data:
            타소득ox = "O" if any(str(pd.get(c, "")).strip() == "O" for c in TAXINCOME_COLS) else "X"
        else:
            타소득ox = ""
        parse_vals = [
            타소득ox,
            str(pd.get("이자", "")),
            str(pd.get("배당", "")),
            str(pd.get("근로(단일)", "")),
            str(pd.get("근로(복수)", "")),
            str(pd.get("연금", "")),
            str(pd.get("기타", "")),
            str(pd.get("기장의무", "")),
            str(pd.get("추계시적용경비율", "")),
        ]

        raw_row = (
            at_vals[:insert_after + 1]
            + parse_vals
            + at_vals[insert_after + 1:]
            + [cell_value(r["fields"].get(n), t) for n, t in linked]
            + [now]
        )

        # ── 셀별 우선순위 ──────────────────────────────────
        # 1) 기존 구글시트가 수식(=...)        → 항상 수식 유지 (에어테이블 무시)
        # 2) GOOGLE_WINS_COLS (입금체크 등)    → 구글시트에 값 있으면 항상 구글 우선
        # 3) 에어테이블 빈값 + 구글시트에 값  → 구글시트값 유지
        # 4) 그 외                             → 에어테이블/파싱 값 사용
        final_row = []
        for col_name, new_val in zip(header, raw_row):
            if col_name == "_sync_at":
                final_row.append(new_val)
            else:
                ex_val = ex.get(col_name, "")
                if str(ex_val).startswith("="):
                    final_row.append(ex_val)       # 수식 셀 → 수식 유지
                elif col_name in GOOGLE_WINS_COLS and str(ex_val).strip() != "":
                    final_row.append(ex_val)       # 구글 우선 컬럼 → 에어테이블 무시
                elif str(new_val).strip() == "" and str(ex_val).strip() != "":
                    final_row.append(ex_val)       # 에어테이블 빈값 → 구글시트값 유지
                else:
                    final_row.append(new_val)
        rows.append(final_row)

    n_formula = sum(
        1 for row in rows[1:] for v in row if str(v).startswith("=")
    )
    if n_formula:
        print(f"  수식 셀 유지: {n_formula}개")

    ws.clear()
    ws.update(range_name="A1", values=rows, value_input_option="USER_ENTERED")
    print(f"[{now}] 동기화 완료: {len(records)}건  (에어테이블 순서 + 안내문파싱 병합, 기존값·수식 보호)")


if __name__ == "__main__":
    main()
