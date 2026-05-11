"""
print_package.py  —  결제용 A4 출력 패키지 생성기
세무회계창연 | 2026

출력 순서 (결제용):
  1. 검증보고서    (검증보고서_*.html  → PDF)
  2. 작업결과 소득시트  ({이름}.xls 프리/사업자복식/... 시트 → PDF)
  3. 작업결과 작업준비  ({이름}.xls 작업준비_* 시트 → PDF)
  4. 안내문 1페이지 (종소세안내문_*.pdf 1페이지)

→ 합쳐서: 출력패키지_{이름}_{날짜}.pdf

※ {이름}.xls = 직원이 작업 완료 후 저장한 작업결과 파일

사용법:
  python print_package.py 박현민 870529
  python print_package.py --all
"""

import sys, io, os, re, shutil, tempfile, unicodedata, fnmatch
from pathlib import Path
from datetime import datetime


def nfc_glob(folder, pattern: str):
    """Mac SMB NFD 파일명 대응 glob"""
    nfc_pat = unicodedata.normalize("NFC", pattern)
    return [p for p in folder.iterdir()
            if fnmatch.fnmatch(unicodedata.normalize("NFC", p.name), nfc_pat)]

sys.path.insert(0, r"F:\종소세2026")
os.environ.setdefault("SEOTAX_ENV", "nas")

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import PyPDF2
import xlwings as xw
from playwright.sync_api import sync_playwright

# ── 경로 ───────────────────────────────────────────────────────────
CUSTOMER_DIR = Path(r"Z:\종소세2026\고객")

# 작업판 시트 중 출력 대상 (프리/복식 작업판)
WORKPAN_SHEETS = {"프리", "사업자복식", "프리복식", "사업자+프리", "사업자+사업자", "프리+프리"}


# ═══════════════════════════════════════════════════════════════════
# 유틸
# ═══════════════════════════════════════════════════════════════════

def find_folder(name: str, jumin6: str = "") -> Path | None:
    nfc_name = unicodedata.normalize("NFC", name)
    candidates = [p for p in CUSTOMER_DIR.iterdir()
                  if p.is_dir() and unicodedata.normalize("NFC", p.name).startswith(f"{nfc_name}_")]
    if not candidates:
        p = CUSTOMER_DIR / name
        return p if p.is_dir() else None
    if jumin6:
        exact = [c for c in candidates
                 if c.name.endswith(f"_{jumin6}") and c.is_dir()]
        if exact:
            return exact[0]
    dirs = [c for c in candidates if c.is_dir()]
    return dirs[0] if dirs else None


# ═══════════════════════════════════════════════════════════════════
# 변환 함수
# ═══════════════════════════════════════════════════════════════════

def html_to_pdf(html_path: Path, pdf_path: Path) -> bool:
    """playwright Chromium으로 HTML → A4 PDF 변환"""
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.goto(html_path.as_uri(), wait_until="networkidle")
            page.pdf(
                path=str(pdf_path),
                format="A4",
                print_background=True,
                margin={"top": "10mm", "bottom": "10mm",
                        "left": "10mm", "right": "10mm"},
            )
            browser.close()
        return True
    except Exception as e:
        print(f"  [오류] HTML→PDF 변환 실패: {e}")
        return False


def sheet_to_pdf(xlsx_path: Path, sheet_name: str, pdf_path: Path) -> bool:
    """xlwings로 특정 시트 → PDF 변환 (Windows + Excel 필요)"""
    app = None
    wb = None
    try:
        app = xw.App(visible=False)
        app.display_alerts = False
        wb = app.books.open(str(xlsx_path))
        sheet = wb.sheets[sheet_name]
        sheet.api.ExportAsFixedFormat(
            Type=0,                     # xlTypePDF
            Filename=str(pdf_path),
            Quality=0,                  # xlQualityStandard
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
        return True
    except Exception as e:
        print(f"  [오류] 시트 '{sheet_name}' PDF 변환 실패: {e}")
        return False
    finally:
        if wb:
            try:
                wb.close()
            except Exception:
                pass
        if app:
            try:
                app.quit()
            except Exception:
                pass


def extract_first_page(pdf_in: Path, pdf_out: Path) -> bool:
    """PDF 첫 페이지만 추출"""
    try:
        reader = PyPDF2.PdfReader(str(pdf_in))
        writer = PyPDF2.PdfWriter()
        writer.add_page(reader.pages[0])
        with open(pdf_out, "wb") as f:
            writer.write(f)
        return True
    except Exception as e:
        print(f"  [오류] PDF 첫 페이지 추출 실패: {e}")
        return False


def merge_pdfs(pdf_list: list[Path], out_path: Path) -> bool:
    """여러 PDF → 하나로 합치기"""
    try:
        merger = PyPDF2.PdfMerger()
        for p in pdf_list:
            merger.append(str(p))
        with open(out_path, "wb") as f:
            merger.write(f)
        merger.close()
        return True
    except Exception as e:
        print(f"  [오류] PDF 합치기 실패: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════
# 메인 처리
# ═══════════════════════════════════════════════════════════════════

def make_package(name: str, jumin6: str = "") -> Path | None:
    folder = find_folder(name, jumin6)
    if not folder:
        print(f"  [오류] 폴더 없음: {name}")
        return None

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    print(f"\n{'═'*54}")
    print(f"  출력 패키지 생성: {name} ({jumin6})")
    print(f"{'═'*54}")
    print(f"  폴더: {folder}")

    tmpdir = Path(tempfile.mkdtemp(prefix="print_pkg_"))
    pdf_parts: list[Path] = []

    try:
        # ── 1. 검증보고서 HTML → PDF ──────────────────────────────
        html_files = sorted(nfc_glob(folder, "검증보고서_*.html"),
                            key=lambda p: p.stat().st_mtime, reverse=True)
        if html_files:
            latest_html = html_files[0]
            pdf_check = tmpdir / "01_검증보고서.pdf"
            print(f"\n  [1] 검증보고서: {latest_html.name}")
            if html_to_pdf(latest_html, pdf_check):
                pdf_parts.append(pdf_check)
                print(f"      → PDF 변환 완료")
            else:
                print(f"      → 스킵 (변환 실패)")
        else:
            print(f"\n  [1] 검증보고서: 없음 — tax_cross_verify.py 먼저 실행")

        # ── 2. 작업판/작업결과 시트 → PDF ────────────────────────────
        # 작업결과_*.xlsx 우선, 없으면 작업판_*.xlsx fallback
        jakup_files = sorted(nfc_glob(folder, "작업결과_*.xlsx"),
                             key=lambda p: p.stat().st_mtime, reverse=True)
        if not jakup_files:
            jakup_files = sorted(nfc_glob(folder, "작업판_*.xlsx"),
                                 key=lambda p: p.stat().st_mtime, reverse=True)
        if jakup_files:
            xlsx = jakup_files[0]
            # xlwings로 시트 목록 확인
            app_tmp = xw.App(visible=False)
            app_tmp.display_alerts = False
            wb_tmp = app_tmp.books.open(str(xlsx))
            sheet_names = [s.name for s in wb_tmp.sheets]
            wb_tmp.close()
            app_tmp.quit()

            # 작업준비 시트
            jakupjunbi_sheets = [s for s in sheet_names if s.startswith("작업준비_")]
            # 작업판 시트 (WORKPAN_SHEETS 중 해당 것)
            workpan_sheet = next((s for s in sheet_names if s in WORKPAN_SHEETS), None)

            print(f"\n  [2] 작업판 엑셀: {xlsx.name}")
            print(f"      시트 목록: {sheet_names}")

            # 작업판 시트 PDF
            if workpan_sheet:
                pdf_wp = tmpdir / "02_작업판.pdf"
                print(f"      작업판 시트: '{workpan_sheet}'")
                if sheet_to_pdf(xlsx, workpan_sheet, pdf_wp):
                    pdf_parts.append(pdf_wp)
                    print(f"      → PDF 변환 완료")
            else:
                print(f"      → 작업판 시트 없음 (스킵)")

            # 작업준비 시트 PDF
            if jakupjunbi_sheets:
                sname = jakupjunbi_sheets[0]
                pdf_jj = tmpdir / "03_작업준비.pdf"
                print(f"\n  [3] 작업준비 시트: '{sname}'")
                if sheet_to_pdf(xlsx, sname, pdf_jj):
                    pdf_parts.append(pdf_jj)
                    print(f"      → PDF 변환 완료")
            else:
                print(f"\n  [3] 작업준비 시트: 없음")
        else:
            print(f"\n  [2/3] 작업판 xlsx 없음 — jakupan_gen.py 먼저 실행")

        # ── 4. 안내문 첫 페이지 ───────────────────────────────────
        ann_files = sorted(nfc_glob(folder, "종소세안내문_*.pdf"),
                           key=lambda p: p.stat().st_mtime, reverse=True)
        if ann_files:
            ann = ann_files[0]
            pdf_ann = tmpdir / "04_안내문1페이지.pdf"
            print(f"\n  [4] 안내문: {ann.name}")
            if extract_first_page(ann, pdf_ann):
                pdf_parts.append(pdf_ann)
                print(f"      → 1페이지 추출 완료")
        else:
            print(f"\n  [4] 안내문 PDF 없음")

        # ── 5. 신고서 PDF (당기) ──────────────────────────────────
        singoser = folder / "신고서.pdf"
        if singoser.exists():
            pdf_sg = tmpdir / "05_신고서.pdf"
            print(f"\n  [5] 신고서: {singoser.name}")
            try:
                import shutil as _sh
                _sh.copy2(str(singoser), str(pdf_sg))
                pdf_parts.append(pdf_sg)
                print(f"      → 추가 완료")
            except Exception as e:
                print(f"      → 복사 실패: {e}")
        else:
            print(f"\n  [5] 신고서.pdf 없음 — 스킵")

        # ── 6. 기존 출력패키지 → _archive 이동 ────────────────────
        old_pkgs = nfc_glob(folder, "출력패키지_*.pdf")
        if old_pkgs:
            arch = folder / "_archive"
            arch.mkdir(exist_ok=True)
            for op in old_pkgs:
                try:
                    op.rename(arch / op.name)
                    print(f"\n  [archive] {op.name} → _archive/")
                except Exception as e:
                    print(f"\n  [archive 실패] {op.name}: {e}")

        # ── 7. 합치기 ─────────────────────────────────────────────
        if not pdf_parts:
            print(f"\n  [오류] 합칠 PDF가 없습니다.")
            return None

        out_path = folder / f"출력패키지_{name}_{ts}.pdf"
        print(f"\n  [7] PDF 합치기 ({len(pdf_parts)}개)")
        for i, p in enumerate(pdf_parts, 1):
            print(f"      {i}. {p.name}")

        if merge_pdfs(pdf_parts, out_path):
            reader = PyPDF2.PdfReader(str(out_path))
            print(f"\n  ✅ 저장: {out_path.name}")
            print(f"     총 {len(reader.pages)}페이지")
            print(f"{'═'*54}\n")
            return out_path
        else:
            return None

    finally:
        # 임시 파일 정리
        shutil.rmtree(tmpdir, ignore_errors=True)


# ═══════════════════════════════════════════════════════════════════
# 구글시트 고객 목록
# ═══════════════════════════════════════════════════════════════════

def load_customers():
    from gsheet_writer import get_credentials
    import gspread
    GSHEET_ID  = "1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI"
    SHEET_NAME = "접수명단"
    creds = get_credentials()
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(GSHEET_ID).worksheet(SHEET_NAME)
    rows = ws.get_all_values()
    customers = []
    for row in rows[1:]:
        name  = row[2].strip() if len(row) > 2 else ""
        jumin = row[4].strip() if len(row) > 4 else ""
        if not name:
            continue
        customers.append({"name": name, "jumin6": jumin.replace("-", "")[:6]})
    return customers


# ═══════════════════════════════════════════════════════════════════
# 엔트리포인트
# ═══════════════════════════════════════════════════════════════════

def main():
    args = sys.argv[1:]

    if args and args[0] == "--all":
        customers = load_customers()
        ok = 0
        for i, c in enumerate(customers, 1):
            out = make_package(c["name"], c["jumin6"])
            if out:
                ok += 1
        print(f"\n[완료] {ok}/{len(customers)}명 패키지 생성")
    else:
        if not args:
            print("사용법: python print_package.py 이름 [주민앞6자리]")
            sys.exit(1)
        if len(args) == 1 and "_" in args[0]:
            parts = args[0].rsplit("_", 1)
            name, jumin6 = parts[0], parts[1] if len(parts) > 1 else ""
        else:
            name   = args[0]
            jumin6 = args[1] if len(args) > 1 else ""
        make_package(name, jumin6)


if __name__ == "__main__":
    main()
