#!/bin/bash
# mount_nas.sh — NAS SMB 마운트 (재부팅 후 자동 복구용)
# LaunchAgent: com.taxeng.nas-mount (2분마다 실행)
NAS_IP="192.168.0.100"
LOG="/Users/changmini/종소세2026/mount_nas.log"
MNT_BASE="/Users/changmini/mnt"
log() { echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" >> "$LOG"; }

log "NAS 마운트 시도..."

# 네트워크 대기
for i in $(seq 1 15); do
    ping -c 1 -t 2 "$NAS_IP" &>/dev/null && break
    sleep 2
done

PASS=$(security find-internet-password -s "$NAS_IP" -a admin -w 2>/dev/null)
if [ -z "$PASS" ]; then log "키체인 비번 없음 — 종료"; exit 1; fi

mount_share() {
    local share="$1"
    local mnt_path="$MNT_BASE/$share"

    if mount | grep -q "$mnt_path"; then
        log "$share 이미 마운트됨: $mnt_path"
        return 0
    fi

    mkdir -p "$mnt_path"
    osascript -e "mount volume \"smb://admin:${PASS}@${NAS_IP}/${share}\"" 2>/dev/null
    sleep 3

    if mount | grep -q "$mnt_path"; then
        log "$share 마운트 성공: $mnt_path"
        return 0
    else
        log "$share 마운트 실패 (osascript 응답 없음 또는 NAS 거절)"
        return 1
    fi
}

mount_share "장성환"
mount_share "세무작업"
