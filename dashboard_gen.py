"""
dashboard_gen.py — 종소세 고객 폴더 현황 대시보드 생성
세무회계창연 | 2026

실행 (맥미니):
    python3 ~/종소세2026/dashboard_gen.py
    → /Volumes/장성환/종소세2026/_대시보드.html 생성

브라우저에서 열기:
    open /Volumes/장성환/종소세2026/_대시보드.html
"""

import os
import sys
import unicodedata
from pathlib import Path
from datetime import datetime

os.environ.setdefault("SEOTAX_ENV", "nas")
sys.path.insert(0, str(Path(__file__).parent))
from config import CUSTOMER_DIR, BASE

# ── 출력 경로 ─────────────────────────────────────────────────────────────
OUT_HTML = BASE / "_대시보드.html"

# ── NAS 파일 열람 URL (Cloudflare Tunnel) ─────────────────────────────────
NAS_FILE_URL = "https://nas.taxenglab.com/종소세2026/고객"


def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


# ── 컬럼 정의 ─────────────────────────────────────────────────────────────
# (key, 표시명, 파일 패턴 또는 체크 함수)
#   패턴: 문자열이면 glob, callable이면 함수(folder, name) → Path|None

def _glob_first(folder: Path, pattern: str) -> Path | None:
    hits = sorted(folder.glob(pattern))
    return hits[0] if hits else None

def _subdir_count(folder: Path, subdir: str):
    d = folder / subdir
    if not d.is_dir():
        return None
    files = [f for f in d.iterdir() if f.is_file()]
    return files[0] if files else None   # 대표 파일 1개 반환

COLUMNS = [
    # (key,            표시명,        체크 lambda(folder, name) → Path|None)
    ("안내문",         "안내문",       lambda f, n: _glob_first(f, f"종소세안내문_{n}.pdf")),
    ("전기신고내역",   "전기내역",     lambda f, n: _glob_first(f, "전년도종소세신고내역.xls*")),
    ("전기신고서PDF",  "전기신고서",   lambda f, n: _glob_first(f, "2025*신고서*.pdf")),
    ("지급명세서",     "지급명세서",   lambda f, n: _subdir_count(f, "지급명세서")),
    ("작업판",         "작업판",       lambda f, n: _glob_first(f, f"작업판_{n}.xlsx")),
    ("작업결과",       "작업결과",     lambda f, n: _glob_first(f, f"작업결과_{n}.xlsx")),
    ("당기신고서",     "당기신고서",   lambda f, n: _glob_first(f, "신고서.pdf")),
    ("검증보고서",     "검증",         lambda f, n: _glob_first(f, "검증보고서_*.html")),
    ("출력패키지",     "출력패키지",   lambda f, n: _glob_first(f, "출력패키지_*.pdf")),
    ("접수증",         "접수증",       lambda f, n: _glob_first(f, f"종합소득세 접수증 {n}.pdf")),
    ("스크래핑신고서", "스크래핑신고서", lambda f, n: _glob_first(f, f"종합소득세 신고서 {n}.pdf")),
    ("소득세납부서",   "소득납부서",   lambda f, n: _glob_first(f, f"종합소득세 납부서 {n}.pdf")),
    ("지방세납부서",   "지방납부서",   lambda f, n: _glob_first(f, f"지방소득세 납부서 {n}.pdf")),
    ("랜딩HTML",       "랜딩",         lambda f, n: _glob_first(f, f"신고결과_{n}.html")),
]

# 신고결과 그룹 (없으면 빨간 강조)
RESULT_KEYS = {"접수증", "신고서", "소득세납부서", "지방세납부서", "랜딩HTML"}
# 신고 전 준비 그룹
PREP_KEYS   = {"안내문", "작업판", "작업결과", "전기신고내역", "전기신고서PDF", "지급명세서"}


def _fmt_date(path: Path) -> str:
    try:
        ts = path.stat().st_mtime
        return datetime.fromtimestamp(ts).strftime("%m/%d %H:%M")
    except Exception:
        return ""


def _cell(path: Path | None, key: str, folder_name: str) -> str:
    """파일 존재 여부에 따라 TD 반환"""
    if path is None:
        cls = "miss-result" if key in RESULT_KEYS else "miss"
        return f'<td class="{cls}">—</td>'

    dt  = _fmt_date(path)
    rel = path.relative_to(CUSTOMER_DIR / folder_name) if path.is_relative_to(CUSTOMER_DIR / folder_name) else path.name
    url = f"{NAS_FILE_URL}/{folder_name}/{rel}"
    tip = f'{path.name}\n{dt}'
    cls = "ok-result" if key in RESULT_KEYS else "ok"
    return f'<td class="{cls}"><a href="{url}" target="_blank" title="{tip}">✓</a><br><span class="dt">{dt}</span></td>'


def build_row(folder: Path) -> str:
    raw_name = folder.name
    parts    = _nfc(raw_name).split("_")
    name     = parts[0]
    jumin6   = parts[1] if len(parts) > 1 else ""

    cells = [f'<td class="name">{_nfc(name)}<br><span class="jumin">{jumin6}</span></td>']
    for key, _, checker in COLUMNS:
        try:
            hit = checker(folder, name)
        except Exception:
            hit = None
        cells.append(_cell(hit, key, raw_name))

    return "<tr>" + "".join(cells) + "</tr>\n"


def build_html(rows: list[str], total: int, generated: str) -> str:
    headers = "".join(f"<th>{col[1]}</th>" for col in COLUMNS)
    rows_html = "".join(rows)
    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>종소세 고객 현황 대시보드 — 세무회계창연</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: -apple-system, 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif;
         background: #f0f2f5; padding: 16px; font-size: 12px; }}
  h1   {{ font-size: 16px; font-weight: 800; color: #111; margin-bottom: 4px; }}
  .meta {{ font-size: 11px; color: #999; margin-bottom: 12px; }}
  .wrap {{ overflow-x: auto; }}
  table {{ border-collapse: collapse; white-space: nowrap; background: white;
           border-radius: 10px; overflow: hidden;
           box-shadow: 0 1px 8px rgba(0,0,0,0.08); }}
  th {{ background: #1e3a5f; color: white; padding: 7px 10px;
        font-size: 11px; font-weight: 700; text-align: center; }}
  td {{ padding: 5px 8px; border-bottom: 1px solid #f0f0f0;
        text-align: center; vertical-align: middle; }}
  td.name {{ text-align: left; font-weight: 700; color: #111;
             min-width: 90px; font-size: 12px; }}
  td.name .jumin {{ font-size: 10px; color: #aaa; font-weight: 400; }}
  td.ok        {{ background: #f0fdf4; color: #15803d; }}
  td.ok-result {{ background: #dcfce7; color: #15803d; font-weight: 700; }}
  td.miss      {{ background: #fafafa; color: #ccc; }}
  td.miss-result {{ background: #fff1f2; color: #f87171; font-weight: 700; }}
  td a {{ color: inherit; text-decoration: none; font-size: 14px; }}
  td a:hover {{ text-decoration: underline; }}
  .dt {{ font-size: 9px; color: #aaa; display: block; line-height: 1.2; }}
  tr:hover td {{ filter: brightness(0.96); }}
  .summary {{ margin-top: 10px; font-size: 11px; color: #666; }}
</style>
</head>
<body>
<h1>📊 종소세 고객 현황 대시보드</h1>
<div class="meta">세무회계창연 | 생성: {generated} | 총 {total}명</div>
<div class="wrap">
<table>
  <thead>
    <tr>
      <th>고객</th>
      {headers}
    </tr>
  </thead>
  <tbody>
    {rows_html}
  </tbody>
</table>
</div>
<div class="summary">
  ✅ 초록 = 파일 있음 &nbsp;|&nbsp;
  — 회색 = 없음 &nbsp;|&nbsp;
  — 연빨강 = 신고결과 미완료
</div>
</body>
</html>"""


def run():
    if not CUSTOMER_DIR.exists():
        print(f"[오류] CUSTOMER_DIR 없음: {CUSTOMER_DIR}")
        return

    folders = sorted(
        [p for p in CUSTOMER_DIR.iterdir()
         if p.is_dir() and not p.name.startswith("_")],
        key=lambda p: _nfc(p.name)
    )
    print(f"고객 {len(folders)}명 처리 중...")

    rows = []
    for i, folder in enumerate(folders, 1):
        try:
            rows.append(build_row(folder))
        except Exception as e:
            print(f"  [{folder.name}] 오류: {e}")
        if i % 50 == 0:
            print(f"  {i}/{len(folders)}...")

    generated = datetime.now().strftime("%Y-%m-%d %H:%M")
    html = build_html(rows, len(folders), generated)
    OUT_HTML.write_text(html, encoding="utf-8")
    print(f"\n✅ 대시보드 생성 완료: {OUT_HTML}")
    print(f"   브라우저: open '{OUT_HTML}'")


if __name__ == "__main__":
    run()
