"""
incometaxbot.py - 종소세 신규접수 자동 파싱 봇 (@incometax777_bot)

동작:
  n8n → Airtable 신규접수 → Telegram 메시지 수신
  "{이름}님 신규/기존 접수되었습니다." 패턴 감지
  → 신규/기존 구분 없이 항상 홈택스ID/PW 방식으로 처리
  → NAS에 PDF 있으면 바로 파싱
  → PDF 없으면 Telegram으로 Windows 실행 명령어 안내

실행: python3 ~/macmini-bots/incometaxbot.py  (Mac Mini)
"""

import sys, os, re, logging, subprocess
from pathlib import Path
from datetime import date

sys.path.insert(0, os.path.expanduser("~/종소세2026"))

from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes
from gsheet_writer import get_credentials
import gspread

# ===== 설정 =====
from config_secret import AUTO_PARSE_BOT_TOKEN
TOKEN          = AUTO_PARSE_BOT_TOKEN
SPREADSHEET_ID = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
NAS_BASE       = Path("/Users/changmini/NAS/종소세2026/고객")
SEASON_END     = date(2026, 6, 1)
CHAT_ID_FILE   = Path.home() / "종소세2026/.credentials/telegram_chat_id.txt"
ADMIN_CHAT_ID  = "5980411081"

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO,
    handlers=[
        logging.FileHandler(os.path.expanduser("~/incometaxbot.log")),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)


# ===== 구글시트 고객 정보 조회 =====
def get_customer_info(name: str):
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)
        ws = gc.open_by_key(SPREADSHEET_ID).worksheet("접수명단")
        for r in ws.get_all_records():
            if str(r.get("성명", "")).strip() == name:
                jumin = str(r.get("주민번호", "") or "").replace("-", "").strip()
                if jumin.isdigit() and len(jumin) < 13:
                    jumin = jumin.zfill(13)  # 구글시트 숫자 저장 시 앞자리 0 소실 복원
                return {
                    "name":       name,
                    "jumin_raw":  jumin,
                    "hometax_id": str(r.get("홈택스아이디", "") or "").strip(),
                    "hometax_pw": str(r.get("홈택스비번", "") or "").strip(),
                    "수입":       str(r.get("수입", "") or "").strip(),
                }
        return None
    except Exception as e:
        logger.error("구글시트 조회 실패: %s", e)
        return None


# ===== PDF 존재 여부 확인 =====
def has_pdf(name: str) -> bool:
    try:
        import unicodedata
        name_nfc = unicodedata.normalize("NFC", name)
        for folder in NAS_BASE.iterdir():
            if unicodedata.normalize("NFC", folder.name).startswith(f"{name_nfc}_"):
                return bool(list(folder.glob("종소세안내문_*.pdf")))
    except Exception as e:
        logger.warning("PDF 확인 실패: %s", e)
    return False


# ===== Mac Mini에서 파싱 실행 =====
def run_parse(name: str) -> str:
    try:
        result = subprocess.run(
            [sys.executable, "-c",
             f"import sys; sys.path.insert(0,'{os.path.expanduser('~/종소세2026')}'); "
             f"import parse_and_sync_신규 as pm; pm.NEW_NAMES=['{name}']; pm.main()"],
            capture_output=True, text=True, timeout=120,
            cwd=os.path.expanduser("~/종소세2026")
        )
        out = (result.stdout + result.stderr).strip()
        return out[-400:] if out else "출력 없음"
    except subprocess.TimeoutExpired:
        return "타임아웃"
    except Exception as e:
        return f"오류: {e}"


# ===== chat_id 자동 저장 =====
def save_chat_id(chat_id: int):
    try:
        CHAT_ID_FILE.parent.mkdir(parents=True, exist_ok=True)
        CHAT_ID_FILE.write_text(str(chat_id))
    except Exception as e:
        logger.warning("chat_id 저장 실패: %s", e)


# ===== 메시지 핸들러 =====
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (update.message.text or "").strip()
    logger.info("수신: %s", text[:100])

    # chat_id 자동 저장 (auto_parse.py 알림용)
    save_chat_id(update.message.chat_id)

    if date.today() >= SEASON_END:
        await update.message.reply_text("⏹ 종소세 시즌 종료(6/1)")
        return

    # 패턴: "{이름}님 신규 접수" 또는 "{이름}님 기존 접수" — 신규/기존 필수 (다른 메시지 오작동 방지)
    m = re.search(r"(.+?)님\s*(신규|기존)\s*접수", text)
    if not m:
        return

    name = m.group(1).strip()
    logger.info("접수 감지: %s", name)
    await update.message.reply_text(f"📥 {name}님 접수 확인\n구글시트 조회 중...")

    info = get_customer_info(name)
    if not info:
        await update.message.reply_text(f"❌ {name} — 구글시트 없음. 수동 처리 필요")
        return

    if not info["hometax_id"] or not info["jumin_raw"]:
        await update.message.reply_text(
            f"⚠️ {name} — 홈택스ID 또는 주민번호 없음\n"
            "구글시트 입력 후 재시도 필요"
        )
        return

    # 이미 파싱 완료
    if info["수입"]:
        await update.message.reply_text(
            f"ℹ️ {name} — 이미 파싱 완료 (수입: {info['수입']}원)"
        )
        return

    # PDF 있으면 바로 파싱
    if has_pdf(name):
        await update.message.reply_text(f"📄 PDF 확인 → 파싱 시작...")
        out = run_parse(name)
        await update.message.reply_text(f"✅ {name} 파싱 완료\n\n{out}")
        return

    # PDF 없으면 Windows 실행 명령어 안내
    await update.message.reply_text(
        f"🖥 {name} — NAS에 PDF 없음\n\n"
        f"Windows 데스크탑에서 실행:\n"
        f"`python _run_one.py {name} {info['hometax_id']} {info['hometax_pw']} {info['jumin_raw']}`\n\n"
        f"(Edge 디버그 창 먼저 열기)"
    )


async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if date.today() >= SEASON_END:
        await update.message.reply_text("⏹ 시즌 종료됨")
        return
    nas_ok = "✅" if NAS_BASE.exists() else "❌ NAS 미연결"
    await update.message.reply_text(
        f"✅ incometax777_bot 정상\n"
        f"NAS: {nas_ok}\n"
        f"시즌 종료: {SEASON_END}"
    )


# ===== 중복 실행 방지 + NAS 대기 =====
def acquire_pid_lock(name: str):
    import os, sys, atexit
    pid_file = Path.home() / f".{name}.pid"
    if pid_file.exists():
        try:
            old_pid = int(pid_file.read_text().strip())
            os.kill(old_pid, 0)
            print(f"[{name}] 이미 실행 중 (PID {old_pid}) — 종료")
            sys.exit(1)
        except (ValueError, ProcessLookupError):
            pid_file.unlink(missing_ok=True)
    pid_file.write_text(str(os.getpid()))
    atexit.register(pid_file.unlink, missing_ok=True)

def wait_for_nas(path: Path, timeout: int = 60) -> bool:
    import time
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            next(path.iterdir()); return True
        except StopIteration:
            return True
        except Exception:
            time.sleep(3)
    return False


# ===== 메인 =====
def main():
    acquire_pid_lock("incometaxbot")
    if date.today() >= SEASON_END:
        logger.info("시즌 종료 - 봇 시작 안 함")
        return

    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("status", cmd_status))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    logger.info("incometax777_bot 시작")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
