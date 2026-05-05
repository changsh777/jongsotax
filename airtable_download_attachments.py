"""
airtable_download_attachments.py
에어테이블 자료접수 테이블 첨부파일 → NAS 고객 폴더로 다운로드

저장 경로: Z:\종소세2026\고객\성명_주민앞6자리\자료\{필드명}\{파일명}
"""
import sys, io, json, urllib.request, urllib.error, time
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import os
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, r"F:\종소세2026")

from config_secret import AIRTABLE_PAT
from config import CUSTOMER_DIR

BASE_ID  = "appSvDTDOmYfBeIFs"
TABLE_ID = "tbl39rqc8msQc7xDX"  # 자료접수

# 필드명 → 저장 서브폴더명
ATT_FIELDS = {
    "신용카드엑셀자료": "신용카드엑셀자료",
    "주민등록등본":     "주민등록등본",
    "연말정산pdf":      "연말정산pdf",
    "기타자료":         "기타자료",
}


def api_get(url):
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {AIRTABLE_PAT}"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def download_file(url, dest_path):
    """URL → 파일 저장. 성공 True, 실패 False."""
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as r:
            data = r.read()
        dest_path.write_bytes(data)
        return True
    except Exception as e:
        print(f"      [다운로드실패] {e}", flush=True)
        return False


def find_customer_folder(name, jumin_raw):
    """성명+주민번호로 NAS 폴더 찾기. 없으면 None."""
    jumin6 = str(jumin_raw).replace("-", "").replace(" ", "")[:6]

    # 1순위: 정확한 폴더명
    exact = CUSTOMER_DIR / f"{name}_{jumin6}"
    if exact.is_dir():
        return exact

    # 2순위: 성명만으로 glob (주민번호 없는 경우 대비)
    if jumin6:
        candidates = list(CUSTOMER_DIR.glob(f"{name}_{jumin6}*"))
        if len(candidates) == 1:
            return candidates[0]

    # 3순위: 성명으로만 찾기 (단일 매칭)
    by_name = [f for f in CUSTOMER_DIR.iterdir()
               if f.is_dir() and f.name.split("_")[0] == name
               and not f.name.startswith("_")]
    if len(by_name) == 1:
        return by_name[0]

    return None


def main():
    dry = "--dry" in sys.argv

    # 전체 레코드 조회 (pagination)
    records = []
    offset = None
    while True:
        url = f"https://api.airtable.com/v0/{BASE_ID}/{TABLE_ID}?pageSize=100"
        if offset:
            url += f"&offset={offset}"
        data = api_get(url)
        records.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset:
            break

    print(f"[에어테이블] 자료접수 레코드: {len(records)}개\n")

    ok = 0
    skip_exist = 0
    skip_no_folder = 0
    fail = 0
    no_folder_list = []

    for rec in records:
        fields = rec.get("fields", {})
        name  = str(fields.get("성명", "") or "").strip()
        jumin = str(fields.get("주민번호", "") or "").strip()

        # 첨부파일 없으면 스킵
        all_atts = []
        for field_name in ATT_FIELDS:
            for att in fields.get(field_name, []):
                all_atts.append((field_name, att))
        if not all_atts:
            continue

        # 고객 폴더 찾기
        folder = find_customer_folder(name, jumin)
        if not folder:
            skip_no_folder += len(all_atts)
            no_folder_list.append(f"{name}({jumin[:6] if jumin else '?'})")
            continue

        print(f"[{name}] → {folder.name} ({len(all_atts)}개)", flush=True)

        for field_name, att in all_atts:
            sub_dir = folder / "자료" / ATT_FIELDS[field_name]
            sub_dir.mkdir(parents=True, exist_ok=True)

            filename = att.get("filename", "unknown")
            dest = sub_dir / filename

            if dest.exists():
                print(f"  [스킵:기존] {ATT_FIELDS[field_name]}/{filename}", flush=True)
                skip_exist += 1
                continue

            url = att.get("url", "")
            if not url:
                print(f"  [URL없음] {filename}", flush=True)
                fail += 1
                continue

            if dry:
                print(f"  [DRY] {ATT_FIELDS[field_name]}/{filename} ({att.get('size',0):,}bytes)", flush=True)
                ok += 1
            else:
                size = att.get("size", 0)
                print(f"  [다운로드] {ATT_FIELDS[field_name]}/{filename} ({size:,}bytes)", flush=True)
                if download_file(url, dest):
                    ok += 1
                else:
                    fail += 1
                time.sleep(0.2)  # API 레이트 리밋 방지

    print()
    print(f"{'[DRY RUN] ' if dry else ''}결과:")
    print(f"  다운로드 성공: {ok}개")
    print(f"  기존 파일 스킵: {skip_exist}개")
    print(f"  NAS 폴더 없음 스킵: {skip_no_folder}개")
    print(f"  실패: {fail}개")
    if no_folder_list:
        print(f"\nNAS 폴더 없는 고객 ({len(no_folder_list)}명):")
        for n in no_folder_list:
            print(f"  {n}")


if __name__ == "__main__":
    main()
