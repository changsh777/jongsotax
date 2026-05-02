"""
auto_download.py - 구글시트 신규 감지 → Edge 자동실행 → 홈택스 PDF 다운 (Windows 전용)

Windows 작업 스케줄러 등록 (2분마다):
  python F:\종소세2026\auto_download.py

동작:
  1. 구글시트 신규 고객 중 NAS PDF 없는 건 감지 (최초 1건씩 처리)
  2. Edge CDP 자동 실행 (안 켜져 있으면 자동 시작)
  3. _run_one.py 로직 실행 → PDF 다운 + 파싱 + 구글시트/에어테이블 업데이트
  4. Mac Mini auto_parse.py 는 수입 채워진 거 보고 자동 스킵
"""

import sys, os, time, json, subprocess, urllib.request, pickle, unicodedata
from pathlib import Path
from datetime import date, datetime

sys.path.insert(0, r"F:\종소세2026")
os.environ.setdefault("SEOTAX_ENV", "nas")

import gspread
from google.auth.transport.requests import Request

# ── 설정 ─────────────────────────────────────────────
SEASON_END     = date(2026, 6, 1)
CRED_DIR       = Path(r"F:\종소세2026\.credentials")
TOKEN_FILE     = CRED_DIR / "token.pickle"
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME     = "접수명단"
NAS_BASE       = Path(r"Z:\종소세2026\고객")
EDGE_EXE       = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
EDGE_PROFILE   = Path(r"F:\종소세2026\.edge_debug_profile")
CDP_URL        = "http://localhost:9222"
SEEN_FILE      = CRED_DIR / "seen_download.json"


# ── 인증 ────────────────────────────────────────────
def get_creds():
    with open(TOKEN_FILE, "rb") as f:
        creds = pickle.load(f)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return creds


# ── 이미 시도한 고객 추적 ──────────────────────────────
def load_seen():
    if SEEN_FILE.exists():
        return set(json.loads(SEEN_FILE.read_text(encoding="utf-8")))
    return set()

def save_seen(seen: set):
    SEEN_FILE.write_text(json.dumps(list(seen)), encoding="utf-8")


# ── NAS PDF 확인 ──────────────────────────────────────
def has_pdf(name):
    name_nfc = unicodedata.normalize("NFC", name)
    try:
        for folder in NAS_BASE.iterdir():
            if unicodedata.normalize("NFC", folder.name).startswith(f"{name_nfc}_"):
                return bool(list(folder.glob("종소세안내문_*.pdf")))
    except Exception as e:
        print(f"  PDF 확인 실패: {e}")
    return False


# ── Edge CDP 실행 여부 확인 ───────────────────────────
def is_edge_running():
    try:
        urllib.request.urlopen(f"{CDP_URL}/json", timeout=2)
        return True
    except Exception:
        return False

def launch_edge():
    print("  Edge 자동 실행 중...")
    EDGE_PROFILE.mkdir(parents=True, exist_ok=True)
    subprocess.Popen([
        EDGE_EXE,
        "--remote-debugging-port=9222",
        f"--user-data-dir={EDGE_PROFILE}",
        "https://hometax.go.kr",
    ])
    time.sleep(4)  # Edge 시작 대기


# ── _run_one.py 실행 ──────────────────────────────────
def run_one(name, hid, pw, jumin):
    print(f"  [{name}] 다운로드 + 파싱 시작...")
    result = subprocess.run(
        [sys.executable, r"F:\종소세2026\_run_one.py", name, hid, pw, jumin],
        capture_output=True, text=True, timeout=300,
        cwd=r"F:\종소세2026"
    )
    out = (result.stdout + result.stderr).strip()
    print(out[-300:] if out else "출력 없음")
    return result.returncode == 0


# ── 메인 ─────────────────────────────────────────────
def main():
    if date.today() >= SEASON_END:
        return

    gc = gspread.authorize(get_creds())
    ws = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    rows = ws.get_all_records()

    seen = load_seen()

    for r in rows:
        if str(r.get("고객구분", "")).strip() != "신규":
            continue
        if str(r.get("수입", "")).strip():
            continue  # 이미 파싱 완료
        hid  = str(r.get("홈택스아이디", "") or "").strip()
        if not hid:
            continue
        name = str(r.get("성명", "")).strip()
        if not name:
            continue
        if has_pdf(name):
            continue  # PDF 이미 있음 (Mac Mini가 파싱 처리)
        if name in seen:
            continue  # 이미 시도함

        # 새 고객 발견
        print(f"[{datetime.now().strftime('%H:%M:%S')}] 신규 감지: {name}")

        pw    = str(r.get("홈택스비번", "") or "").strip()
        jumin = str(r.get("주민번호", "") or "").replace("-", "").strip()
        if jumin.isdigit() and len(jumin) < 13:
            jumin = jumin.zfill(13)

        # Edge 디버그 모드로 재시작
        subprocess.run(["taskkill", "/f", "/im", "msedge.exe"], capture_output=True)
        time.sleep(1)
        launch_edge()
        if not is_edge_running():
            print(f"  Edge 실행 실패 — 스킵")
            continue

        # seen 등록 (실패해도 재시도 안 함 — 수동 확인 필요)
        seen.add(name)
        save_seen(seen)

        ok = run_one(name, hid, pw, jumin)
        if ok:
            print(f"  [{name}] 완료")
        else:
            print(f"  [{name}] 실패 — 로그 확인 필요")

        break  # 한 번에 1명씩 처리 (안정성)


if __name__ == "__main__":
    main()
