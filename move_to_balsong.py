# -*- coding: utf-8 -*-
# 접수증 + 신고서를 올바른 발송용 폴더로 이동
# 잘못 만들어진 {name}\ 폴더는 빈 경우 삭제

import os, shutil, unicodedata

NAS        = r"Z:\종소세2026\고객"
LOCAL      = r"C:\Users\pc\종소세2026"

def nfc(s):
    return unicodedata.normalize("NFC", str(s))

def find_customer_folder(name):
    """NAS에서 {name}_XXXXXX 폴더 찾기. 복수 hit 시 None 반환(경고)."""
    name_nfc = nfc(name)
    hits = []
    for d in os.listdir(NAS):
        parts = nfc(d).split("_")
        if parts[0] == name_nfc and len(parts) >= 2 and os.path.isdir(os.path.join(NAS, d)):
            hits.append(d)
    if len(hits) == 1:
        return os.path.join(NAS, hits[0])
    if len(hits) > 1:
        print(f"  !! 동명이인 {name}: {hits} — 스킵 (수동 확인)")
        return None
    print(f"  !! {name} 폴더 미발견 — 스킵")
    return None

# 처리할 고객 목록 (홈택스 목록 기준)
names = [
    "강유진","김성준","김지은","나기은","박현민","변은지",
    "오상연","이근만","이명회","이선웅","정재호","지성환"
]

ok = 0
skip = 0

for name in names:
    print(f"\n[{name}]")
    cdir = find_customer_folder(name)
    if not cdir:
        skip += 1
        continue

    balsong = os.path.join(cdir, "발송용")
    os.makedirs(balsong, exist_ok=True)

    moved_any = False

    # ── 접수증 이동 ──────────────────────────────────────────────
    # 잘못된 위치: Z:\종소세2026\고객\{name}\{name}_접수증.pdf
    wrong_folder = os.path.join(NAS, name)
    jeup_src = os.path.join(wrong_folder, f"{name}_접수증.pdf")
    jeup_dst = os.path.join(balsong, f"{name}_접수증.pdf")
    if os.path.exists(jeup_src):
        shutil.move(jeup_src, jeup_dst)
        print(f"  접수증 이동: {jeup_dst}")
        moved_any = True
    elif os.path.exists(jeup_dst):
        print(f"  접수증 이미 있음: {jeup_dst}")
    else:
        print(f"  접수증 파일 없음")

    # 잘못된 폴더 비었으면 삭제
    if os.path.isdir(wrong_folder) and not os.listdir(wrong_folder):
        os.rmdir(wrong_folder)
        print(f"  빈 폴더 삭제: {wrong_folder}")

    # ── 신고서 이동 ──────────────────────────────────────────────
    # 로컬: C:\Users\pc\종소세2026\{name}_종합소득세.pdf
    shingo_src = os.path.join(LOCAL, f"{name}_종합소득세.pdf")
    shingo_dst = os.path.join(balsong, f"{name}_종합소득세.pdf")
    if os.path.exists(shingo_src):
        shutil.move(shingo_src, shingo_dst)
        print(f"  신고서 이동: {shingo_dst}")
        moved_any = True
    elif os.path.exists(shingo_dst):
        print(f"  신고서 이미 있음: {shingo_dst}")
    else:
        print(f"  신고서 파일 없음")

    if moved_any:
        ok += 1
    else:
        skip += 1

print(f"\n완료: 처리 {ok}명 / 스킵 {skip}명")
