"""
pipeline_runner.py — 종소세 파이프라인 레고블럭 실행기
세무회계창연 | 2026

각 모듈은 독립 블럭: is_done(폴더확인) → can_run(조건확인) → run_fn(실행)
파일시스템 상태 = 진행 상태 DB

사용법:
  python pipeline_runner.py 이혜주               # 한 명 전체 파이프라인 (auto 실행)
  python pipeline_runner.py 이혜주 --dry         # 한 명 상태만 확인
  python pipeline_runner.py 이혜주 --stages C,D  # 한 명 특정 단계만
  python pipeline_runner.py --all               # 전체 고객 auto 실행
  python pipeline_runner.py --all --dry         # 전체 고객 상태 테이블
  python pipeline_runner.py --all --stages D    # 전체 고객 D단계만
"""
import sys, io, os, subprocess, unicodedata, fnmatch, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)
os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, str(__import__("pathlib").Path(__file__).parent))

from pathlib import Path
from dataclasses import dataclass, field
from typing import Callable, Optional, List

from config import CUSTOMER_DIR, PARSE_RESULT_XLSX

PROJ = Path(__file__).parent


# ── NFC glob 헬퍼 (macOS SMB 마운트 대응) ────────────────────────────
def _g(folder: Path, pattern: str) -> Optional[Path]:
    nfc = unicodedata.normalize("NFC", pattern)
    try:
        hits = sorted(
            f for f in folder.iterdir()
            if f.is_file() and fnmatch.fnmatch(unicodedata.normalize("NFC", f.name), nfc)
        )
        return hits[0] if hits else None
    except Exception:
        return None

def _sub(folder: Path, subdir: str) -> Optional[Path]:
    """서브폴더 내 첫 번째 파일"""
    nfc_sub = unicodedata.normalize("NFC", subdir)
    try:
        for item in folder.iterdir():
            if item.is_dir() and unicodedata.normalize("NFC", item.name) == nfc_sub:
                files = [f for f in item.iterdir() if f.is_file()]
                return files[0] if files else None
    except Exception:
        pass
    return None

def find_folder(name: str, jumin6: str) -> Optional[Path]:
    nfc = unicodedata.normalize("NFC", f"{name}_{jumin6}")
    try:
        for f in CUSTOMER_DIR.iterdir():
            if f.is_dir() and unicodedata.normalize("NFC", f.name) == nfc:
                return f
    except Exception:
        pass
    # jumin6 없이 이름만으로도 탐색
    if not jumin6:
        nfc_name = unicodedata.normalize("NFC", name)
        try:
            for f in CUSTOMER_DIR.iterdir():
                fn = unicodedata.normalize("NFC", f.name)
                if f.is_dir() and fn.startswith(f"{nfc_name}_"):
                    return f
        except Exception:
            pass
    return None

def load_all_customers() -> list:
    result = []
    try:
        for folder in sorted(CUSTOMER_DIR.iterdir()):
            if not folder.is_dir() or folder.name.startswith("_"):
                continue
            nfc = unicodedata.normalize("NFC", folder.name)
            if "_" not in nfc:
                continue
            parts = nfc.rsplit("_", 1)
            if len(parts) == 2 and parts[1].isdigit() and len(parts[1]) == 6:
                result.append({"name": parts[0], "jumin6": parts[1]})
    except Exception:
        pass
    return result


# ── 파싱결과 캐시 (C_parse is_done 판단용) ───────────────────────────
_parse_data: Optional[dict] = None

def _load_parse_cache():
    global _parse_data
    if _parse_data is not None:
        return
    _parse_data = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(PARSE_RESULT_XLSX, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, max_col=5, values_only=True):
            if row and row[0]:
                n = str(row[0]).strip()
                try:
                    income = int(float(str(row[4] or 0).replace(",", "") or 0))
                except Exception:
                    income = 0
                _parse_data[n] = income
        wb.close()
    except Exception:
        pass

def _parse_done(name: str) -> bool:
    _load_parse_cache()
    return _parse_data.get(name, 0) > 0

def _invalidate_parse_cache(name: str, income: int = 1):
    global _parse_data
    if _parse_data is None:
        _parse_data = {}
    _parse_data[name] = income


# ── 실행 함수들 (auto 모듈) ──────────────────────────────────────────
def _run_parse(name: str, jumin6: str) -> tuple:
    """C단계: 안내문 PDF 파싱 → gsheet 업데이트"""
    try:
        from parse_to_xlsx import parse_anneam
        from gsheet_writer import write_parsed_to_접수명단

        folder = find_folder(name, jumin6)
        if not folder:
            return False, "폴더 없음"

        pdfs = sorted(
            [f for f in folder.iterdir()
             if f.is_file() and fnmatch.fnmatch(
                 unicodedata.normalize("NFC", f.name),
                 unicodedata.normalize("NFC", "종소세안내문_*.pdf"))],
            key=lambda p: p.stat().st_mtime
        )
        if not pdfs:
            return False, "안내문 PDF 없음"

        data = parse_anneam(pdfs[-1])
        data["성명"] = name
        write_parsed_to_접수명단([data])

        income = data.get("수입금액총계") or 0
        _invalidate_parse_cache(name, income)
        return True, f"수입: {income:,}" if income else "완료 (수입 없음)"
    except Exception as e:
        return False, str(e)


def _run_jakupan(name: str, jumin6: str) -> tuple:
    """D단계: 작업판 생성"""
    try:
        result = subprocess.run(
            [sys.executable, str(PROJ / "jakupan_gen.py"), name, jumin6],
            capture_output=True, text=True, timeout=120,
            cwd=str(PROJ)
        )
        out = (result.stdout + result.stderr).strip()
        return result.returncode == 0, (out[-200:] if out else "출력 없음")
    except subprocess.TimeoutExpired:
        return False, "타임아웃 (2분)"
    except Exception as e:
        return False, str(e)


def _stub(name: str, jumin6: str) -> tuple:
    """아직 연결 안된 자동 모듈 스텁"""
    return False, "구현예정 — run_fn 연결 필요"


# ── 모듈 정의 ──────────────────────────────────────────────────────────
@dataclass
class Module:
    id:      str
    stage:   str
    label:   str
    mode:    str        # "auto" | "manual" | "pending"
    is_done: Callable   # (folder, name) → bool
    can_run: Callable   # (folder, name) → bool
    run_fn:  Callable = None
    wip:     bool = False  # True = 구조만 있고 아직 run_fn 미연결

# ──────────────────────────────────────────────────────────────────────
# 파이프라인 정의  (블럭 추가/제거/순서 변경만 하면 됨)
# ──────────────────────────────────────────────────────────────────────
PIPELINE: List[Module] = [

    # ── B. 자료 수집 (수동: Playwright 필요) ─────────────────────
    Module("B_안내문",    "B", "안내문PDF",    "manual",
           lambda f, n: _g(f, "종소세안내문_*.pdf") is not None,
           lambda f, n: True),

    Module("B_지급명세서","B", "지급명세서",   "manual",
           lambda f, n: _sub(f, "지급명세서") is not None,
           lambda f, n: True),

    Module("B_전기내역",  "B", "전기신고내역", "manual",
           lambda f, n: _g(f, "전년도종소세신고내역.xls*") is not None,
           lambda f, n: True),

    # ── C. 파싱/분석 (자동) ──────────────────────────────────────
    Module("C_parse",    "C", "안내문파싱",   "auto",
           lambda f, n: _parse_done(n),
           lambda f, n: _g(f, "종소세안내문_*.pdf") is not None,
           _run_parse),

    # ── D. 작업판 생성 (자동) ────────────────────────────────────
    Module("D_작업판",   "D", "작업판생성",   "auto",
           lambda f, n: _g(f, f"작업판_{n}.xlsx") is not None,
           lambda f, n: _parse_done(n),
           _run_jakupan),

    # ── E. 신고 작업 (수동: 직원) ────────────────────────────────
    Module("E_작업결과", "E", "작업결과",     "manual",
           lambda f, n: _g(f, f"작업결과_{n}.xlsx") is not None,
           lambda f, n: _g(f, f"작업판_{n}.xlsx") is not None),

    # ── F. 검증/출력 (자동 예정) ─────────────────────────────────
    Module("F_검증",     "F", "검증보고서",   "auto",
           lambda f, n: _g(f, "검증보고서_*.html") is not None,
           lambda f, n: _g(f, f"작업결과_{n}.xlsx") is not None,
           _stub, wip=True),

    Module("F_출력패키지","F", "출력패키지",   "auto",
           lambda f, n: _g(f, "출력패키지_*.pdf") is not None,
           lambda f, n: _g(f, "검증보고서_*.html") is not None,
           _stub, wip=True),

    # ── G. 홈택스 신고 (수동: 세무사) ───────────────────────────
    Module("G_신고",     "G", "홈택스신고",   "manual",
           lambda f, n: _g(f, f"종합소득세 접수증 {n}.pdf") is not None,
           lambda f, n: _g(f, "출력패키지_*.pdf") is not None),

    # ── H. 신고결과 수집 (수동/미구현) ──────────────────────────
    Module("H1_접수증",  "H", "접수증PDF",    "manual",
           lambda f, n: _g(f, f"종합소득세 접수증 {n}.pdf") is not None,
           lambda f, n: True),

    Module("H2_지방세",  "H", "지방세납부서", "pending",
           lambda f, n: _g(f, f"지방소득세 납부서 {n}.pdf") is not None,
           lambda f, n: False),   # 미구현

    # ── I. 결과 안내 (자동 예정) ─────────────────────────────────
    Module("I_랜딩",     "I", "랜딩HTML",     "auto",
           lambda f, n: _g(f, f"신고결과_{n}.html") is not None,
           lambda f, n: _g(f, f"종합소득세 접수증 {n}.pdf") is not None,
           _stub, wip=True),
]


# ── 상태 판단 ──────────────────────────────────────────────────────────
# 상태값: "done" | "waiting" | "blocked" | "pending"
def get_status(mod: Module, folder: Optional[Path], name: str) -> str:
    if folder is None:
        return "blocked"
    try:
        if mod.is_done(folder, name):
            return "done"
    except Exception:
        pass
    if mod.mode == "pending":
        return "pending"
    try:
        if not mod.can_run(folder, name):
            return "blocked"
    except Exception:
        return "blocked"
    return "waiting"


# ── 한 명 실행 ─────────────────────────────────────────────────────────
def run_one(name: str, jumin6: str,
            stage_filter=None, dry=False, verbose=True) -> list:

    folder = find_folder(name, jumin6)
    if not jumin6 and folder:
        jumin6 = unicodedata.normalize("NFC", folder.name).split("_")[-1]

    if verbose:
        print(f"\n{'─'*54}")
        print(f"  {name}  ({jumin6})  {'[DRY]' if dry else ''}")
        print(f"{'─'*54}")

    results = []
    prev_stage = None

    for mod in PIPELINE:
        if stage_filter and mod.stage not in stage_filter:
            continue

        # 단계 구분선
        if verbose and mod.stage != prev_stage:
            if prev_stage is not None:
                print()
            prev_stage = mod.stage

        st   = get_status(mod, folder, name)
        mode_label = {"auto": "자동", "manual": "수동", "pending": "미구현"}.get(mod.mode, "")
        wip_tag    = " [예정]" if mod.wip else ""

        if st == "done":
            icon, action = "✅", "완료"

        elif st == "pending":
            icon, action = "🔴", "미구현"

        elif st == "blocked":
            icon, action = "⏸ ", "조건미충족"

        elif st == "waiting":
            if mod.mode == "manual":
                icon, action = "👤", "수동대기"

            elif mod.mode == "auto" and not mod.wip and not dry:
                # 실제 실행
                print(f"  ▶  [{mod.id}] {mod.label} 실행 중...", flush=True)
                ok, msg = mod.run_fn(name, jumin6)
                icon    = "✅" if ok else "❌"
                action  = f"완료  {msg}" if ok else f"오류: {msg[:60]}"
            else:
                icon, action = "⬜", f"대기{wip_tag}"
        else:
            icon, action = "❓", st

        if verbose:
            print(f"  {icon} [{mod.id:<15}] {mod.label:<12} {action:<28} ({mode_label})",
                  flush=True)
        results.append({"id": mod.id, "stage": mod.stage, "status": st, "action": action})

    return results


# ── 전체 고객 실행 / 상태 테이블 ─────────────────────────────────────
def run_all(stage_filter=None, dry=False):
    customers = load_all_customers()
    print(f"\n전체 고객 {len(customers)}명\n", flush=True)

    if dry:
        # 콤팩트 상태 테이블
        cols = [m for m in PIPELINE if not stage_filter or m.stage in stage_filter]
        # 헤더
        hdr = f"{'이름':<9}" + "".join(f"{m.label[:5]:^7}" for m in cols)
        print(hdr)
        print("─" * len(hdr))
        sym = {"done": "✓", "waiting": "·", "blocked": "⏸", "pending": "✗"}
        for c in customers:
            folder = find_folder(c["name"], c["jumin6"])
            row = f"{c['name']:<9}"
            for mod in cols:
                st = get_status(mod, folder, c["name"])
                row += f"{'  ' + sym.get(st, '?') + '  ':^7}"
            print(row, flush=True)
    else:
        for c in customers:
            run_one(c["name"], c["jumin6"],
                    stage_filter=stage_filter, dry=False, verbose=True)
            time.sleep(1)   # gsheet API 쿼타 여유


# ── CLI ────────────────────────────────────────────────────────────────
def _parse_args(argv):
    dry          = "--dry" in argv
    all_mode     = "--all" in argv
    stage_filter = None
    args_clean   = []

    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--stages" and i + 1 < len(argv):
            stage_filter = set(argv[i + 1].upper().split(","))
            i += 2
            continue
        if a.startswith("--stages="):
            stage_filter = set(a[9:].upper().split(","))
        elif not a.startswith("--"):
            args_clean.append(a)
        i += 1

    return all_mode, dry, stage_filter, args_clean


def main():
    all_mode, dry, stage_filter, args_clean = _parse_args(sys.argv[1:])

    if all_mode:
        run_all(stage_filter=stage_filter, dry=dry)
        return

    if args_clean:
        name   = args_clean[0]
        jumin6 = args_clean[1] if len(args_clean) > 1 else ""
        run_one(name, jumin6, stage_filter=stage_filter, dry=dry)
        return

    print(__doc__, flush=True)


if __name__ == "__main__":
    main()
