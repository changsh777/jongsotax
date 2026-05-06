"""
jongsotaxbot.py - 종소세 작업 전용 텔레그램 봇 (@jongsotax_bot)

명령어:
  /작업 장성환      NAS에서 파일 꺼내서 전송 (안내문+전년도+부가세+작업시트)
  /수임동의 장성환  진행 상태 조회
  /발송 장성환      접수증+납부서 링크 발송 (게이트 포함)
  장성환.pdf 업로드 → NAS 신고서.pdf 저장 (기존 있으면 _archive 이동)

실행: python3 ~/종소세2026/jongsotaxbot.py
"""

import os
import glob
import logging
from pathlib import Path
from datetime import datetime
from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters, ContextTypes
)

# ===== 설정 =====
TOKEN          = "REDACTED_TOKEN_1"
ADMIN_CHAT_ID  = 5980411081    # 세무사 (관리자)
NAS_BASE       = Path("/Users/changmini/NAS/종소세2026/고객")   # Mac Mini NAS 마운트 경로
NAS_URL        = "https://nas.taxenglab.com/종소세2026/고객"    # Cloudflare URL

ALLOWED_USERS: list[int] = []  # 빈 리스트 = 전체 허용

LOG_FILE = os.path.expanduser("~/종소세2026/jongsotaxbot.log")

logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(message)s",
    level=logging.INFO,
    handlers=[logging.FileHandler(LOG_FILE), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

# 동명이인 선택 대기 상태
user_pending: dict[int, dict] = {}


# ===== 유틸 =====
def nas_ok() -> bool:
    return NAS_BASE.exists()


def find_folders(name: str) -> list[Path]:
    """이름으로 고객 폴더 검색. 동명이인 있으면 여러 개 반환.
    NFD/NFC 양쪽 다 NFC로 정규화 후 비교 (Mac SMB 한글 인코딩 차이 회피)
    """
    import unicodedata
    name_nfc = unicodedata.normalize("NFC", name)
    prefix   = f"{name_nfc}_"
    try:
        return sorted(
            p for p in NAS_BASE.iterdir()
            if p.is_dir() and unicodedata.normalize("NFC", p.name).startswith(prefix)
        )
    except Exception:
        return []


def is_allowed(update: Update) -> bool:
    if not ALLOWED_USERS:
        return True
    return update.effective_user.id in ALLOWED_USERS


async def nas_fail(update: Update):
    await update.message.reply_text("⚠️ NAS 연결 끊김 — 관리자 확인 필요")
    logger.error("NAS 접근 실패: %s", NAS_BASE)


# ===== 동명이인 선택 =====
async def ask_choice(update: Update, user_id: int, folders: list[Path], action: str, extra=None):
    user_pending[user_id] = {"folders": folders, "action": action, "extra": extra}
    lines = "\n".join(f"{i+1}. {f.name}" for i, f in enumerate(folders))
    await update.message.reply_text(f"동명이인 확인:\n{lines}\n\n번호를 입력하세요.")


async def resolve_choice(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    user_id = update.effective_user.id
    if user_id not in user_pending:
        return False
    text = update.message.text.strip()
    try:
        idx = int(text) - 1
        pending = user_pending.pop(user_id)
        folders = pending["folders"]
        if not (0 <= idx < len(folders)):
            await update.message.reply_text("잘못된 번호입니다.")
            return True
        folder = folders[idx]
        action = pending["action"]
        extra  = pending["extra"]
        if action == "작업":
            await do_work(update, folder)
        elif action == "발송":
            await do_send(update, folder)
        elif action == "신고서":
            await do_save_singoser(update, context, folder, extra)
        elif action == "수임동의":
            await do_status(update, folder)
    except ValueError:
        pass
    return True


def parse_name_arg(update: Update, context: ContextTypes.DEFAULT_TYPE) -> str:
    """명령어에서 이름 추출. /work강동수 또는 /work 강동수 둘 다 처리"""
    # context.args 있으면 우선
    if context.args:
        return " ".join(context.args).strip()
    # 없으면 raw text에서 명령어 부분 제거
    text = update.message.text or ""
    # /work강동수 → 강동수
    parts = text.split(None, 1)
    if len(parts) >= 2:
        return parts[1].strip()
    # /work강동수 (공백 없음) → 명령어에서 / 뒤 prefix 제거
    cmd = parts[0].lstrip("/")
    for prefix in ("work", "작업", "agree", "수임동의", "send", "발송"):
        if cmd.startswith(prefix):
            return cmd[len(prefix):].strip()
    return ""


# ===== /작업 =====
async def cmd_work(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    name = parse_name_arg(update, context)
    if not name:
        await update.message.reply_text("사용법: /work 강동수  또는  /작업 강동수"); return
    if not nas_ok():
        await nas_fail(update); return

    folders = find_folders(name)

    if not folders:
        await update.message.reply_text(f"'{name}' 고객 자료가 없습니다. 홈택스 안내문 파싱이 완료됐나요?"); return
    if len(folders) > 1:
        await ask_choice(update, update.effective_user.id, folders, "작업"); return
    await do_work(update, folders[0])


async def do_work(update: Update, folder: Path):
    """안내문 + 전년도자료 + 작업판 + 지급명세서 + 간이용역소득 전송"""
    files_to_send = []

    # 1. 안내문 (최신 1개)
    ann = sorted(folder.glob("종소세안내문_*.pdf"), key=lambda f: f.stat().st_mtime, reverse=True)
    if ann:
        files_to_send.append(ann[0])

    # 2. 전년도 자료
    for pat in ["전년도종소세신고내역.*", "전년도*.xls*"]:
        files_to_send.extend(folder.glob(pat))

    # 3. 작업판 엑셀 (최신 1개)
    wp = sorted(folder.glob("작업판_*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    if wp:
        files_to_send.append(wp[0])

    # 4. 지급명세서 폴더
    for sub in ["지급명세서", "간이용역소득"]:
        d = folder / sub
        if d.is_dir():
            files_to_send.extend(sorted(d.iterdir()))

    files_to_send = [f for f in files_to_send if f.is_file()]
    if not files_to_send:
        await update.message.reply_text(f"{folder.name}: 작업 파일 없음 (홈택스 안내문 파싱 완료됐나요?)"); return

    await update.message.reply_text(f"📁 {folder.name} — {len(files_to_send)}개 전송 중...")
    for f in files_to_send:
        try:
            with open(f, "rb") as fp:
                await update.message.reply_document(document=fp, filename=f.name)
        except Exception as e:
            await update.message.reply_text(f"❌ {f.name} 실패: {e}")
    await update.message.reply_text("✅ 전송 완료")


# ===== /수임동의 =====
async def cmd_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    name = parse_name_arg(update, context)
    if not name:
        await update.message.reply_text("사용법: /agree 강동수  또는  /수임동의 강동수"); return
    if not nas_ok():
        await nas_fail(update); return

    folders = find_folders(name)

    if not folders:
        await update.message.reply_text(f"'{name}' 폴더 없음"); return
    if len(folders) > 1:
        await ask_choice(update, update.effective_user.id, folders, "수임동의"); return
    await do_status(update, folders[0])


async def do_status(update: Update, folder: Path):
    def chk(pattern):
        return bool(list(folder.glob(pattern)))

    items = {
        "안내문 파싱": chk("종소세안내문_*.pdf"),
        "신고서":      (folder / "신고서.pdf").exists(),
        "접수증":      (folder / "접수증.pdf").exists(),
        "납부서":      (folder / "납부서.pdf").exists(),
    }
    lines = "\n".join(f"{'✅' if v else '❌'} {k}" for k, v in items.items())
    await update.message.reply_text(f"📋 {folder.name}\n{lines}")


# ===== /발송 =====
async def cmd_send(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    name = parse_name_arg(update, context)
    if not name:
        await update.message.reply_text("사용법: /send 강동수  또는  /발송 강동수"); return
    if not nas_ok():
        await nas_fail(update); return

    folders = find_folders(name)

    if not folders:
        await update.message.reply_text(f"'{name}' 폴더 없음"); return
    if len(folders) > 1:
        await ask_choice(update, update.effective_user.id, folders, "발송"); return
    await do_send(update, folders[0])


async def do_send(update: Update, folder: Path):
    """접수증 + 납부서 링크 발송 — 게이트 ⑤⑥ 포함"""
    접수증 = folder / "접수증.pdf"
    납부서 = folder / "납부서.pdf"

    # 게이트 ⑥: 접수증 존재 확인
    if not 접수증.exists():
        await update.message.reply_text("❌ 접수증이 없습니다. 최종신고 완료 후 다시 시도하세요.")
        return

    # 게이트 ⑥: 당일 스크래핑 확인
    today = datetime.now().date()
    mtime = datetime.fromtimestamp(접수증.stat().st_mtime).date()
    if mtime != today:
        await update.message.reply_text(
            f"⚠️ 접수증이 오늘({today}) 파일이 아닙니다 (파일날짜: {mtime})\n"
            "오늘 최종신고 후 다시 시도하세요."
        )
        return

    url = f"{NAS_URL}/{folder.name}"
    lines = [f"📄 접수증: {url}/접수증.pdf"]

    # 게이트 ⑤: 납부서 존재 확인 (소득세>0인 경우 필수)
    if 납부서.exists():
        lines.append(f"💳 납부서: {url}/납부서.pdf")
    else:
        lines.append("ℹ️ 납부서 없음 (환급이거나 누락 — 세무사 확인)")

    await update.message.reply_text(
        f"✅ {folder.name} 발송 링크:\n\n" + "\n".join(lines) +
        "\n\n→ 솔라피 알림톡 발송하세요."
    )


# ===== 파일 수신: 이름.pdf → 신고서 저장 + 검증 =====
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return

    # 동명이인 선택 대기 중이면 먼저 처리
    if await resolve_choice(update, context):
        return

    doc = update.message.document
    if not doc or not (doc.file_name or "").endswith(".pdf"):
        return

    fname = doc.file_name or ""
    fname_lower = fname.lower()

    # 25년 신고서만 처리 (파일명에 "25"+"신고서" 둘 다 있어야)
    is_25_singoser = "25" in fname and "신고서" in fname_lower
    if not is_25_singoser:
        return  # 조건 불만족 → 무시

    # 고객명: 캡션 우선, 없으면 파일명에서 추출
    caption = (update.message.caption or "").strip()
    name    = caption if caption else Path(fname).stem.strip()

    if not nas_ok():
        await nas_fail(update); return

    folders = find_folders(name)
    if not folders:
        await update.message.reply_text(f"'{name}' 폴더 없음 — 이름 확인 필요"); return

    tg_file = await doc.get_file()

    if len(folders) > 1:
        await ask_choice(update, update.effective_user.id, folders, "신고서", tg_file)
        return

    await do_save_singoser(update, context, folders[0], tg_file)


async def do_save_singoser(update: Update, context: ContextTypes.DEFAULT_TYPE, folder: Path, tg_file):
    """신고서.pdf 저장 → 교차검증 → 발신자+관리자 보고서 발송 [ULTRA CRITICAL]"""
    target  = folder / "신고서.pdf"
    archive = folder / "_archive"

    # 기존 파일 archive 이동
    if target.exists():
        archive.mkdir(exist_ok=True)
        ts = datetime.fromtimestamp(target.stat().st_mtime).strftime("%Y%m%d_%H%M%S")
        target.rename(archive / f"신고서_{ts}.pdf")
        logger.info("archive 이동: 신고서_%s.pdf", ts)

    await tg_file.download_to_drive(str(target))
    logger.info("신고서 저장: %s", target)
    await update.message.reply_text(f"✅ {folder.name}/신고서.pdf 저장\n⏳ 교차검증 실행 중...")

    # 교차검증 실행
    try:
        import sys as _sys
        _bot_dir = str(Path(__file__).resolve().parent)
        if _bot_dir not in _sys.path:
            _sys.path.insert(0, _bot_dir)
        from tax_cross_verify import run as verify_run

        parts  = folder.name.rsplit("_", 1)
        _name  = parts[0]
        _jumin = parts[1] if len(parts) > 1 else ""

        html_path = verify_run(_name, _jumin, folder=folder)

        if html_path and html_path.exists():
            sender_id = update.effective_chat.id
            caption   = f"📊 {folder.name} 검증보고서"

            # 발신자에게 전송
            with open(html_path, "rb") as fp:
                await context.bot.send_document(chat_id=sender_id, document=fp,
                                                filename=html_path.name, caption=caption)

            # 관리자에게도 전송 (발신자 ≠ 관리자인 경우만)
            if sender_id != ADMIN_CHAT_ID:
                with open(html_path, "rb") as fp:
                    await context.bot.send_document(
                        chat_id=ADMIN_CHAT_ID, document=fp,
                        filename=html_path.name,
                        caption=f"{caption} (직원 업로드: {update.effective_user.full_name})"
                    )
        else:
            await update.message.reply_text("⚠️ 검증보고서 생성 실패 — 수동 실행 필요")

    except Exception as e:
        logger.error("검증 오류: %s", e, exc_info=True)
        await update.message.reply_text(f"⚠️ 검증 오류: {e}\n(신고서는 저장됐습니다)")


# ===== 텍스트: 동명이인 번호 선택 =====
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return
    await resolve_choice(update, context)


# ===== 메인 =====
def main():
    if not nas_ok():
        logger.warning("NAS 미연결: %s", NAS_BASE)

    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("work",    cmd_work))    # /work 강동수
    app.add_handler(CommandHandler("agree",   cmd_status))  # /agree 강동수
    app.add_handler(CommandHandler("send",    cmd_send))    # /send 강동수
    app.add_handler(MessageHandler(filters.Document.ALL,            handle_file))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info("jongsotax_bot 시작")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
