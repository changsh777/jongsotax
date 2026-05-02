"""
auto_parse.py - 구글시트 접수명단 감시 → 신규 미처리 건 자동 파싱 (Mac Mini 크론)

Mac Mini 크론 등록 (2분마다):
  */2 * * * * /usr/bin/python3 ~/종소세2026/auto_parse.py >> ~/auto_parse.log 2>&1

동작:
  1. 구글시트 접수명단 읽기
  2. 고객구분=신규 + 수입=빈칸 + 홈택스아이디 있음 → 미처리 건 추출
  3. NAS PDF 있으면 parse_and_sync_신규.py 실행
  4. 텔레그램 결과 알림 (chat_id 파일 있을 때)
"""

import sys, os, subprocess, pickle, logging, unicodedata
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
TELEGRAM_TOKEN = "8672211090:AAHecG0siKKAKm5jVUEzDHTfX5v5XSE7BHw"
CHAT_ID_FILE   = CRED_DIR / "telegram_chat_id.txt"
ADMIN_CHAT_ID  = "5980411081"
LOCK_DIR       = Path.home() / "종소세2026/.parse_locks"

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO,
)
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
def get_chat_id():
    if CHAT_ID_FILE.exists():
        return CHAT_ID_FILE.read_text().strip()
    return ADMIN_CHAT_ID

def send_telegram(text):
    import urllib.request, urllib.parse
    chat_id = get_chat_id()
    if not chat_id:
        logger.info("텔레그램 chat_id 없음 — 알림 생략 (chat_id 파일: %s)", CHAT_ID_FILE)
        return
    url  = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    data = urllib.parse.urlencode({"chat_id": chat_id, "text": text, "parse_mode": "Markdown"}).encode()
    try:
        urllib.request.urlopen(url, data=data, timeout=10)
    except Exception as e:
        logger.warning("텔레그램 전송 실패: %s", e)


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


# ── 처리 중 락 ────────────────────────────────────────
def is_locked(name):
    LOCK_DIR.mkdir(parents=True, exist_ok=True)
    return (LOCK_DIR / f"{name}.lock").exists()

def lock(name):
    LOCK_DIR.mkdir(parents=True, exist_ok=True)
    (LOCK_DIR / f"{name}.lock").write_text(datetime.now().isoformat())

def unlock(name):
    lf = LOCK_DIR / f"{name}.lock"
    if lf.exists():
        lf.unlink()


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

    # 미처리 신규 건 추출
    targets = []  # (name, hid, pw, jumin)
    for r in rows:
        구분 = str(r.get("고객구분", "")).strip()
        수입 = str(r.get("수입", "")).strip()
        hid  = str(r.get("홈택스아이디", "") or "").strip()
        pw   = str(r.get("홈택스비번", "") or "").strip()
        name = str(r.get("성명", "")).strip()
        jumin = str(r.get("주민번호", "") or "").replace("-", "").strip()
        if jumin.isdigit() and len(jumin) < 13:
            jumin = jumin.zfill(13)
        if 구분 == "신규" and not 수입 and hid and name:
            targets.append((name, hid, pw, jumin))

    if not targets:
        return  # 조용히 종료

    logger.info("미처리 신규: %s", targets)

    for name, hid, pw, jumin in targets:
        if is_locked(name):
            logger.info("%s 처리 중 (락) — 스킵", name)
            continue

        if not has_pdf(name):
            logger.info("%s NAS PDF 없음 — 스킵 (Windows에서 _run_one.py 필요)", name)
            send_telegram(
                f"⚠️ *{name}* — PDF 없음\n"
                f"Windows에서 실행:\n"
                f"`python _run_one.py {name} {hid} {pw} {jumin}`"
            )
            continue

        lock(name)
        logger.info("%s 파싱 시작", name)
        send_telegram(f"📄 *{name}* 자동 파싱 시작...")

        try:
            out = run_parse(name)
            logger.info("%s 완료: %s", name, out[:100])
            send_telegram(f"✅ *{name}* 파싱 완료\n\n{out}")
        except Exception as e:
            logger.error("%s 파싱 오류: %s", name, e)
            send_telegram(f"❌ *{name}* 파싱 오류: {e}")
        finally:
            unlock(name)


if __name__ == "__main__":
    main()
