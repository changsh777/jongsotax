"""
auto_parse.py - 구글시트 접수명단 감시 → 신규 미처리 건 자동 파싱 (Mac Mini 크론)

크론: */2 * * * * /usr/bin/python3 ~/종소세2026/auto_parse.py >> ~/auto_parse.log 2>&1

동작:
  - 구글시트 신규 고객 중 처음 감지된 건만 처리 (중복 알림 없음)
  - NAS PDF 있으면 → 파싱 자동 실행
  - PDF 없으면 → 최초 1회만 알림 (이후 무시)
"""

import sys, os, subprocess, pickle, logging, unicodedata, json
from pathlib import Path
from datetime import date, datetime

import gspread
from google.auth.transport.requests import Request

# ── 설정 ─────────────────────────────────────────────
SEASON_END     = date(2026, 6, 1)
CRED_DIR       = Path.home() / "종소세2026/.credentials"
TOKEN_FILE     = CRED_DIR / "token.pickle"
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
SHEET_NAME     = "접수명단"
NAS_BASE       = Path("/Users/changmini/NAS/종소세2026/고객")
TELEGRAM_TOKEN = "REDACTED_TOKEN_2"
ADMIN_CHAT_ID  = "5980411081"
LOCK_DIR       = Path.home() / "종소세2026/.parse_locks"
SEEN_FILE      = Path.home() / "종소세2026/.parse_locks/seen.json"  # 이미 처리/알림한 고객

logging.basicConfig(format="%(asctime)s [%(levelname)s] %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)


# ── 인증 ────────────────────────────────────────────
def get_creds():
    with open(TOKEN_FILE, "rb") as f:
        creds = pickle.load(f)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)
    return creds


# ── 텔레그램 ─────────────────────────────────────────
def send_telegram(text):
    import urllib.request, urllib.parse
    url  = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    data = urllib.parse.urlencode({"chat_id": ADMIN_CHAT_ID, "text": text, "parse_mode": "Markdown"}).encode()
    try:
        urllib.request.urlopen(url, data=data, timeout=10)
    except Exception as e:
        logger.warning("텔레그램 전송 실패: %s", e)


# ── 이미 처리/알림한 고객 목록 ────────────────────────
def load_seen():
    LOCK_DIR.mkdir(parents=True, exist_ok=True)
    if SEEN_FILE.exists():
        return set(json.loads(SEEN_FILE.read_text()))
    return set()

def save_seen(seen: set):
    SEEN_FILE.write_text(json.dumps(list(seen)))


# ── NAS PDF 확인 ──────────────────────────────────────
def has_pdf(name):
    name_nfc = unicodedata.normalize("NFC", name)
    try:
        for folder in NAS_BASE.iterdir():
            if unicodedata.normalize("NFC", folder.name).startswith(f"{name_nfc}_"):
                return bool(list(folder.glob("종소세안내문_*.pdf")))
    except Exception as e:
        logger.warning("PDF 확인 실패: %s", e)
    return False


# ── 파싱 실행 ─────────────────────────────────────────
def run_parse(name):
    try:
        result = subprocess.run(
            [sys.executable, "-c",
             f"import sys; sys.path.insert(0,'{os.path.expanduser('~/종소세2026')}'); "
             f"import parse_and_sync_신규 as pm; pm.NEW_NAMES=['{name}']; pm.main()"],
            capture_output=True, text=True, timeout=180,
            cwd=os.path.expanduser("~/종소세2026")
        )
        out = (result.stdout + result.stderr).strip()
        return out[-500:] if out else "출력 없음"
    except subprocess.TimeoutExpired:
        return "타임아웃 (3분 초과)"
    except Exception as e:
        return f"오류: {e}"


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
        if not name or name in seen:
            continue  # 이미 처리/알림한 고객

        pw    = str(r.get("홈택스비번", "") or "").strip()
        jumin = str(r.get("주민번호", "") or "").replace("-", "").strip()
        if jumin.isdigit() and len(jumin) < 13:
            jumin = jumin.zfill(13)

        # seen에 추가 (중복 알림 방지)
        seen.add(name)
        save_seen(seen)

        if not has_pdf(name):
            logger.info("%s NAS PDF 없음", name)
            send_telegram(
                f"📥 신규 접수: *{name}*\n"
                f"PDF 없음 — Windows에서 실행:\n"
                f"`python _run_one.py {name} {hid} {pw} {jumin}`"
            )
            continue

        logger.info("%s PDF 확인 → 파싱 시작", name)
        send_telegram(f"📄 *{name}* 파싱 시작...")
        out = run_parse(name)
        logger.info("%s 완료", name)
        send_telegram(f"✅ *{name}* 파싱 완료\n\n{out}")


if __name__ == "__main__":
    main()
