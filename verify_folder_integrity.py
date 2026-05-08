"""
verify_folder_integrity.py — 고객 폴더 내용 기반 혼입 검증 + 자동 이동

검증 항목:
  1. 안내문 PDF  → '성명 XXX' 파싱 → 폴더명 이름과 불일치 시 경고 (마스킹 이름 제외)
  2. 지급명세서 PDF → '⑩주민등록번호 XXXXXX' → 폴더명 주민앞6자리와 불일치 시 자동 이동

실행:
  python verify_folder_integrity.py          # 검증만
  python verify_folder_integrity.py --fix    # 검증 + 지급명세서 자동 이동
"""
from __future__ import annotations
import sys, io, os, re, shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
os.environ.setdefault("SEOTAX_ENV", "nas")
_proj = str(__import__('pathlib').Path(__file__).resolve().parent)
if _proj not in sys.path:
    sys.path.insert(0, _proj)

from pathlib import Path
from datetime import datetime
from typing import Optional
import unicodedata
import pdfplumber
from config import CUSTOMER_DIR

FIX_MODE = "--fix" in sys.argv


def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", str(s))


def extract_name_from_anneam(pdf_path: Path) -> Optional[str]:
    """안내문 PDF에서 성명 추출"""
    try:
        with pdfplumber.open(str(pdf_path)) as p:
            text = p.pages[0].extract_text() or ""
        m = re.search(r'성명\s+(\S+)', text)
        return m.group(1) if m else None
    except Exception:
        return None


def extract_jumin6_from_jipgum(pdf_path: Path) -> Optional[str]:
    """지급명세서 PDF에서 소득자 주민번호 앞 6자리 추출"""
    try:
        with pdfplumber.open(str(pdf_path)) as p:
            text = ""
            for page in p.pages:
                text += (page.extract_text() or "") + "\n"
        m = re.search(r'⑩\s*주민등록번호\s+(\d{6})', text)
        if m:
            return m.group(1)
        m2 = re.search(r'생년월일\s+(\d{2})\.(\d{2})\.(\d{2})', text)
        if m2:
            return m2.group(1) + m2.group(2) + m2.group(3)
        return None
    except Exception:
        return None


def find_folder_by_jumin6(jumin6: str) -> Optional[Path]:
    """주민번호 앞 6자리로 올바른 고객 폴더 찾기"""
    for p in CUSTOMER_DIR.iterdir():
        if not p.is_dir(): continue
        parts = nfc(p.name).split("_")
        if len(parts) > 1 and parts[1] == jumin6:
            return p
    return None


def auto_fix_jipgum(src_file: Path, correct_folder: Path):
    """혼입된 지급명세서를 올바른 폴더로 이동"""
    dst_dir = correct_folder / "지급명세서"
    dst_dir.mkdir(exist_ok=True)
    # 파일명을 올바른 폴더명으로 변경
    new_name = f"{nfc(correct_folder.name)}.pdf"
    dst_file = dst_dir / new_name
    # 중복 방지
    if dst_file.exists():
        ts = datetime.now().strftime("%Y%m%d%H%M%S")
        new_name = f"{nfc(correct_folder.name)}_{ts}.pdf"
        dst_file = dst_dir / new_name
    shutil.move(str(src_file), str(dst_file))
    # 원본 폴더 비었으면 삭제
    parent = src_file.parent
    if parent.is_dir() and not any(parent.iterdir()):
        parent.rmdir()
    return dst_file


def main():
    if not CUSTOMER_DIR.exists():
        print(f"[오류] 고객 폴더 없음: {CUSTOMER_DIR}")
        sys.exit(1)

    mode_str = "검증 + 자동이동" if FIX_MODE else "검증만"
    print(f"고객 폴더 혼입 검증 [{mode_str}]  {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    folders = sorted(
        p for p in CUSTOMER_DIR.iterdir()
        if p.is_dir() and not p.name.startswith("_")
    )
    total = len(folders)
    print(f"폴더 {total}개 검사 중...\n")

    issues = []
    fixed = []
    ok_count = 0

    for i, folder in enumerate(folders, 1):
        folder_nfc = nfc(folder.name)
        parts = folder_nfc.split("_")
        my_name  = parts[0]
        my_jumin = parts[1] if len(parts) > 1 else ""

        folder_ok = True
        print(f"[{i}/{total}] {folder.name}", end=" ", flush=True)

        # ── 1. 안내문 PDF 검증 (보고만) ───────────────────────────
        anneam_files = [f for f in folder.iterdir()
                        if f.is_file() and "안내문" in nfc(f.name) and f.suffix == ".pdf"]
        if anneam_files:
            pdf_name = extract_name_from_anneam(anneam_files[0])
            if pdf_name and '*' not in pdf_name and nfc(pdf_name) != my_name:
                issues.append({
                    "folder": folder.name,
                    "file": anneam_files[0].name,
                    "type": "안내문",
                    "detail": f"폴더[{my_name}] ≠ PDF[{pdf_name}]",
                    "fixable": False,
                })
                folder_ok = False
                print(f"⚠안내문({pdf_name})", end=" ")

        # ── 2. 지급명세서 PDF 검증 + 자동이동 ─────────────────────
        jip_dir = folder / "지급명세서"
        jip_pdfs = []
        if jip_dir.is_dir():
            jip_pdfs = [f for f in jip_dir.iterdir() if f.is_file() and f.suffix == ".pdf"]
        jip_pdfs += [f for f in folder.iterdir()
                     if f.is_file() and f.suffix == ".pdf" and "지급명세서" in nfc(f.name)]

        for jp in jip_pdfs:
            pdf_jumin6 = extract_jumin6_from_jipgum(jp)
            if pdf_jumin6 and my_jumin and pdf_jumin6 != my_jumin:
                correct_folder = find_folder_by_jumin6(pdf_jumin6)
                issue = {
                    "folder": folder.name,
                    "file": str(jp.relative_to(folder)),
                    "type": "지급명세서",
                    "detail": f"폴더[{my_jumin}] ≠ PDF주민[{pdf_jumin6}] → 올바른폴더[{correct_folder.name if correct_folder else '미확인'}]",
                    "fixable": correct_folder is not None,
                    "_src": jp,
                    "_dst_folder": correct_folder,
                }
                if FIX_MODE and correct_folder:
                    dst = auto_fix_jipgum(jp, correct_folder)
                    fixed.append(f"{folder.name} → {correct_folder.name}/{dst.name}")
                    print(f"✅이동→{correct_folder.name}", end=" ")
                else:
                    issues.append(issue)
                    folder_ok = False
                    print(f"⚠지급명세서[{pdf_jumin6}]", end=" ")

        if folder_ok:
            ok_count += 1
            print("✓")
        else:
            print()

    # ── 결과 ──────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    if FIX_MODE and fixed:
        print(f"✅ 자동이동 {len(fixed)}건:")
        for f in fixed:
            print(f"   {f}")
        print()
    if not issues:
        print(f"✅ 이상 없음 — {ok_count}/{total}개 폴더 정상")
    else:
        print(f"⚠️  수동확인 필요 {len(issues)}건 (정상 {ok_count}/{total})")
        print(f"{'='*60}")
        for iss in issues:
            print(f"  [{iss['type']}] {iss['folder']} / {iss['file']}")
            print(f"    → {iss['detail']}")
    print(f"{'='*60}")
    print(f"완료: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


if __name__ == "__main__":
    main()
