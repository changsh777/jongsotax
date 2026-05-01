"""
airtable_watcher.py - 에어테이블 첨부파일 감시 → NAS 저장 → 텔레그램 알림

실행:
  python airtable_watcher.py          # 30분 간격 데몬 모드
  python airtable_watcher.py --once   # 1회 실행 후 종료 (테스트용)

환경변수:
  BOT_TOKEN      텔레그램 봇 토큰
  ADMIN_CHAT_ID  텔레그램 채팅 ID
"""

import os
import sys
import time
import json
import glob
import logging
import argparse
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path

# ===== 설정 =====
AIRTABLE_PAT   = os.environ.get("AIRTABLE_PAT", "")   # .env 또는 launchd EnvironmentVariables
BASE_ID        = "appSvDTDOmYfBeIFs"
TABLE_ID       = "tbl2f2h6GfSnLCQpt"

NAS_CLIENT_ROOT = "/Users/changmini/NAS/종소세2026/고객"
LOG_PATH        = "/Users/changmini/NAS/종소세2026/_로그/airtable_watcher.log"
POLL_INTERVAL   = 30 * 60   # 30분 (초 단위)

# 에어테이블 "성명" 필드명 (실제 필드명과 일치해야 함)
FIELD_NAME      = "성명"
FIELD_ATTACHMENTS = "첨부파일"   # 에어테이블 첨부파일 필드명 — 실제 필드명으로 변경 필요

# ===== 로깅 =====
def setup_logging():
    log_dir = os.path.dirname(LOG_PATH)
    os.makedirs(log_dir, exist_ok=True)

    fmt = "%(asctime)s [%(levelname)s] %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S"

    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        datefmt=datefmt,
        handlers=[
            logging.FileHandler(LOG_PATH, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )

log = logging.getLogger(__name__)


# ===== 텔레그램 =====
def telegram_send(text: str):
    """텔레그램 메시지 전송. 실패 시 로그만 남기고 계속."""
    token   = os.environ.get("BOT_TOKEN", "")
    chat_id = os.environ.get("ADMIN_CHAT_ID", "")

    if not token or not chat_id:
        log.warning("텔레그램 환경변수 미설정 (BOT_TOKEN / ADMIN_CHAT_ID) — 알림 생략")
        return

    url  = f"https://api.telegram.org/bot{token}/sendMessage"
    body = json.dumps({"chat_id": chat_id, "text": text}).encode("utf-8")
    req  = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            r.read()
    except Exception as e:
        log.warning(f"텔레그램 전송 실패 (무시): {e}")


# ===== 에어테이블 API =====
def airtable_get(path: str) -> dict:
    url = f"https://api.airtable.com/v0/{path}"
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {AIRTABLE_PAT}"},
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def fetch_all_records() -> list:
    """페이지네이션 처리해서 전체 레코드 반환"""
    records = []
    offset  = None
    page    = 1
    while True:
        path = f"{BASE_ID}/{TABLE_ID}"
        if offset:
            path += f"?offset={offset}"
        data   = airtable_get(path)
        batch  = data.get("records", [])
        records.extend(batch)
        log.debug(f"페이지 {page}: {len(batch)}건 (누적 {len(records)}건)")
        offset = data.get("offset")
        if not offset:
            break
        page += 1
        time.sleep(0.2)
    return records


def fetch_attachment_fields() -> list[str]:
    """테이블 스키마에서 multipleAttachments 타입 필드명 목록 반환"""
    try:
        data = airtable_get(f"meta/bases/{BASE_ID}/tables")
        for t in data.get("tables", []):
            if t["id"] == TABLE_ID:
                return [
                    f["name"]
                    for f in t.get("fields", [])
                    if f.get("type") == "multipleAttachments"
                ]
    except Exception as e:
        log.warning(f"스키마 조회 실패 — 기본 필드명 사용: {e}")
    return [FIELD_ATTACHMENTS]


# ===== 고객 폴더 매칭 =====
def find_client_folder(name: str) -> str | None:
    """
    성명으로 NAS 고객 폴더 찾기.
    패턴: {NAS_CLIENT_ROOT}/{name}_XXXXXX/
    """
    pattern = os.path.join(NAS_CLIENT_ROOT, f"{name}_*")
    matches = glob.glob(pattern)
    if not matches:
        return None
    # 여러 개면 첫 번째 (경고는 로그로)
    if len(matches) > 1:
        log.warning(f"'{name}' 매칭 폴더 {len(matches)}개 → 첫 번째 사용: {matches[0]}")
    return matches[0]


def ensure_client_subdir(client_folder: str) -> str:
    """고객폴더/자료/ 디렉터리 생성 후 경로 반환"""
    subdir = os.path.join(client_folder, "자료")
    os.makedirs(subdir, exist_ok=True)
    return subdir


# ===== 파일 중복 체크 =====
def is_duplicate(target_dir: str, filename: str, filesize: int) -> bool:
    """
    파일명과 크기로 중복 판정.
    파일명이 같고 크기도 같으면 이미 저장된 것으로 간주.
    """
    candidate = os.path.join(target_dir, filename)
    if os.path.exists(candidate):
        existing_size = os.path.getsize(candidate)
        if existing_size == filesize:
            return True
    return False


def resolve_filename(target_dir: str, filename: str) -> str:
    """
    같은 이름 파일이 있을 경우 _1, _2 ... 붙여 충돌 없는 이름 반환.
    크기가 달라 중복이 아닌 경우에만 진입.
    """
    candidate = os.path.join(target_dir, filename)
    if not os.path.exists(candidate):
        return filename

    stem, ext = os.path.splitext(filename)
    idx = 1
    while True:
        new_name = f"{stem}_{idx}{ext}"
        if not os.path.exists(os.path.join(target_dir, new_name)):
            return new_name
        idx += 1


# ===== 파일 다운로드 =====
def download_file(url: str, dest_path: str) -> bool:
    """파일 다운로드. 실패 시 False 반환 (크래시 없음)."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "airtable-watcher/1.0"})
        with urllib.request.urlopen(req, timeout=60) as r:
            data = r.read()
        # 임시 경로에 먼저 쓰고 이동 (중간 실패 방지)
        tmp_path = dest_path + ".tmp"
        with open(tmp_path, "wb") as f:
            f.write(data)
        os.replace(tmp_path, dest_path)
        return True
    except Exception as e:
        log.error(f"다운로드 실패 ({url}): {e}")
        # 임시 파일 정리
        tmp_path = dest_path + ".tmp"
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        return False


# ===== 메인 로직 =====
def process_once(attachment_fields: list[str]):
    """1회 전체 스캔 실행"""
    log.info("--- 스캔 시작 ---")

    try:
        records = fetch_all_records()
    except Exception as e:
        log.error(f"에어테이블 레코드 조회 실패: {e}")
        telegram_send(f"[airtable_watcher] 에어테이블 조회 실패: {e}")
        return

    log.info(f"전체 레코드: {len(records)}건 / 첨부 필드: {attachment_fields}")

    new_count = 0

    for rec in records:
        fields      = rec.get("fields", {})
        record_id   = rec.get("id", "?")
        client_name = str(fields.get(FIELD_NAME, "")).strip()

        if not client_name:
            continue

        # 모든 첨부파일 필드를 순회
        for afield in attachment_fields:
            attachments = fields.get(afield)
            if not attachments or not isinstance(attachments, list):
                continue

            for att in attachments:
                filename = att.get("filename", "")
                filesize = att.get("size", 0)
                url      = att.get("url", "")

                if not filename or not url:
                    continue

                # 고객 폴더 찾기
                client_folder = find_client_folder(client_name)
                if client_folder is None:
                    msg = f"[폴더없음] {client_name} — '{filename}' 저장 불가"
                    log.warning(msg)
                    telegram_send(f"⚠️ {msg}")
                    continue

                target_dir = ensure_client_subdir(client_folder)

                # 중복 체크
                if is_duplicate(target_dir, filename, filesize):
                    log.debug(f"스킵(중복): {client_name} / {filename}")
                    continue

                # 파일명 충돌 해소
                save_name = resolve_filename(target_dir, filename)
                dest_path = os.path.join(target_dir, save_name)

                log.info(f"다운로드: {client_name} / {save_name} ({filesize:,} bytes)")

                ok = download_file(url, dest_path)
                if ok:
                    new_count += 1
                    msg = f"📎 {client_name} 자료 도착: {save_name}"
                    log.info(f"저장 완료: {dest_path}")
                    telegram_send(msg)
                else:
                    telegram_send(f"❌ {client_name} 파일 다운로드 실패: {save_name}")

    log.info(f"--- 스캔 완료: 신규 {new_count}건 ---")


def run_daemon(attachment_fields: list[str], once: bool):
    if once:
        process_once(attachment_fields)
        return

    log.info(f"데몬 시작 — {POLL_INTERVAL // 60}분 간격 폴링")
    telegram_send("[airtable_watcher] 감시 데몬 시작")

    while True:
        try:
            process_once(attachment_fields)
        except Exception as e:
            log.exception(f"process_once 예외 (무시하고 계속): {e}")

        next_run = datetime.now().strftime("%H:%M")
        log.info(f"다음 스캔까지 {POLL_INTERVAL // 60}분 대기 (현재 {next_run})")
        time.sleep(POLL_INTERVAL)


# ===== 진입점 =====
def main():
    setup_logging()

    parser = argparse.ArgumentParser(description="에어테이블 첨부파일 감시 데몬")
    parser.add_argument("--once", action="store_true", help="1회 실행 후 종료 (테스트용)")
    args = parser.parse_args()

    # 첨부파일 필드 목록을 스키마에서 자동 감지
    log.info("에어테이블 스키마 조회 중...")
    attachment_fields = fetch_attachment_fields()
    if not attachment_fields:
        log.warning(f"첨부파일 필드 자동 감지 실패 — 기본값 사용: [{FIELD_ATTACHMENTS}]")
        attachment_fields = [FIELD_ATTACHMENTS]
    else:
        log.info(f"감지된 첨부파일 필드: {attachment_fields}")

    run_daemon(attachment_fields, once=args.once)


if __name__ == "__main__":
    main()
