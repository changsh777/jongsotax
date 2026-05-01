"""
세무 자료 안전 저장 유틸 [ULTRA CRITICAL]

원칙:
  - root 폴더: 항상 최신 파일 (직원·고객 보는 곳)
  - _archive 폴더: 옛 버전 timestamped 자동 백업
  - hash 비교로 동일 파일 중복 저장 방지
  - 모든 자동 다운로드는 이 함수만 사용 (직접 save_as 금지)

사용 예:
  # 1) 다운로드 객체 (Playwright)
  safe_download(download, folder, "안내문.pdf")

  # 2) popup PDF
  safe_popup_pdf(popup, folder, "안내문.pdf", format="A4")

  # 3) openpyxl 워크북
  safe_save_workbook(wb, folder, "부가세.xlsx")

  # 4) 일반 임시 파일에서 이동
  save_with_archive(folder, "신고서.pdf", tmp_path)
"""
import hashlib
import shutil
import tempfile
from pathlib import Path
from datetime import datetime


def file_hash(path):
    """MD5 해시"""
    h = hashlib.md5()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(8192)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def save_with_archive(folder, filename, src_path):
    """root 최신 + _archive 백업 패턴.

    Args:
        folder: 대상 폴더 (Path or str)
        filename: 저장할 파일명 (예: '안내문.pdf')
        src_path: 임시 저장 경로 (이 파일을 이동시킴)

    Returns:
        (status, target_path)
        status ∈ {'신규', '갱신', '동일파일_스킵'}
    """
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / filename
    src = Path(src_path)

    if not src.exists():
        raise FileNotFoundError(f"src 파일 없음: {src}")

    if not target.exists():
        # 신규
        shutil.move(str(src), str(target))
        return ("신규", target)

    # 동일 파일 검사
    try:
        same = file_hash(target) == file_hash(src)
    except Exception:
        same = False

    if same:
        # 동일 → 저장 스킵, 임시 파일만 정리
        try:
            src.unlink()
        except Exception:
            pass
        return ("동일파일_스킵", target)

    # 다른 파일 → 기존을 archive로 백업
    archive_dir = folder / "_archive"
    archive_dir.mkdir(exist_ok=True)
    old_mtime = datetime.fromtimestamp(target.stat().st_mtime)
    archived_name = f"{target.stem}_{old_mtime.strftime('%Y%m%d_%H%M%S')}{target.suffix}"
    shutil.move(str(target), str(archive_dir / archived_name))

    # 새 파일을 root로
    shutil.move(str(src), str(target))
    return ("갱신", target)


def _tmp_path(suffix):
    """임시 파일 경로 (닫혀있음)"""
    fd, name = tempfile.mkstemp(suffix=suffix)
    import os
    os.close(fd)
    return Path(name)


def safe_download(download, folder, filename):
    """Playwright download 객체 → archive 패턴 저장"""
    suffix = Path(filename).suffix
    tmp = _tmp_path(suffix)
    download.save_as(str(tmp))
    return save_with_archive(folder, filename, tmp)


def safe_popup_pdf(popup, folder, filename, **pdf_kwargs):
    """Playwright popup → page.pdf() 결과를 archive 패턴 저장"""
    suffix = Path(filename).suffix
    tmp = _tmp_path(suffix)
    popup.pdf(path=str(tmp), **pdf_kwargs)
    return save_with_archive(folder, filename, tmp)


def safe_save_workbook(wb, folder, filename):
    """openpyxl Workbook → archive 패턴 저장"""
    suffix = Path(filename).suffix
    tmp = _tmp_path(suffix)
    wb.save(str(tmp))
    return save_with_archive(folder, filename, tmp)
