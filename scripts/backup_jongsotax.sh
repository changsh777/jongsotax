#!/bin/bash
# backup_jongsotax.sh — 종소세2026 NAS → 외장하드 백업 (하루 2회)
# 위치: ~/종소세2026/scripts/backup_jongsotax.sh
# 맥미니에서 실행

# ── 설정 (외장하드 이름 확인: ls /Volumes/) ──────────────────────────
NAS_SRC="/Users/changmini/NAS/종소세2026"
# 외장하드 경로: ls /Volumes/ 결과에서 이름 확인 후 수정
EXT_DEST="/Volumes/BACKUP/종소세2026_backup"
LOG_FILE="$HOME/종소세2026/backup.log"

# ── 실행 ─────────────────────────────────────────────────────────────
TS=$(date '+%Y-%m-%d %H:%M:%S')

# NAS 마운트 확인
if [ ! -d "$NAS_SRC" ]; then
    echo "[$TS] ❌ NAS 미마운트: $NAS_SRC" | tee -a "$LOG_FILE"
    exit 1
fi

# 외장하드 마운트 확인
if [ ! -d "$(dirname "$EXT_DEST")" ]; then
    echo "[$TS] ❌ 외장하드 없음: $(dirname "$EXT_DEST")" | tee -a "$LOG_FILE"
    exit 1
fi

mkdir -p "$EXT_DEST"

echo "[$TS] 백업 시작: $NAS_SRC → $EXT_DEST" | tee -a "$LOG_FILE"

rsync -av --delete \
    --exclude='.DS_Store' \
    --exclude='._*' \
    --exclude='.Trash*' \
    "$NAS_SRC/" "$EXT_DEST/" \
    >> "$LOG_FILE" 2>&1

STATUS=$?
TS2=$(date '+%Y-%m-%d %H:%M:%S')
if [ $STATUS -eq 0 ]; then
    echo "[$TS2] ✅ 백업 완료" | tee -a "$LOG_FILE"
else
    echo "[$TS2] ⚠️ 백업 오류 (exit $STATUS)" | tee -a "$LOG_FILE"
fi
