#!/bin/bash
# install_backup_cron.sh — 백업 cron 등록 (맥미니에서 1회 실행)
# 실행: bash ~/종소세2026/scripts/install_backup_cron.sh

SCRIPT="$HOME/종소세2026/scripts/backup_jongsotax.sh"

# 실행 권한 부여
chmod +x "$SCRIPT"

# 기존 cron에서 backup_jongsotax 라인 제거 후 새로 추가
(crontab -l 2>/dev/null | grep -v "backup_jongsotax"; \
 echo "0  8 * * * $SCRIPT   # 종소세2026 백업 08:00"; \
 echo "0 22 * * * $SCRIPT   # 종소세2026 백업 22:00" \
) | crontab -

echo "✅ cron 등록 완료:"
crontab -l | grep backup_jongsotax
echo ""
echo "⚠️  외장하드 경로 확인 후 backup_jongsotax.sh 상단 EXT_DEST 수정"
echo "   현재 연결된 볼륨: $(ls /Volumes/)"
