"""
landing_gen.py — 종합소득세 신고결과 랜딩 페이지 생성기
세무회계창연 | 2026

파일명 규칙 (스크래핑 결과):
    종합소득세 접수증 {이름}.pdf
    종합소득세 신고서 {이름}.pdf
    종합소득세 납부서 {이름}.pdf   (없으면 환급)
    지방소득세 납부서 {이름}.pdf   (없으면 지방세 없음)

사용:
    from landing_gen import make_landing_html_auto

    html_path = make_landing_html_auto(
        folder   = Path("Z:/종소세2026/고객/홍길동_800101"),
        name     = "홍길동",
        year     = 2025,
        base_url = "https://taxeng.co.kr/jongsotax/홍길동_800101",
    )
    # → (html_path, public_url)
"""

from pathlib import Path
from datetime import datetime
import unicodedata


def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


# ── 파일 자동 감지 ────────────────────────────────────────────────────

FILE_MAP = {
    "접수증":           "종합소득세 접수증 {name}.pdf",
    "신고서":           "종합소득세 신고서 {name}.pdf",
    "종합소득세납부서": "종합소득세 납부서 {name}.pdf",
    "지방소득세납부서": "지방소득세 납부서 {name}.pdf",
}


def detect_docs(folder: Path, name: str) -> dict:
    """폴더에서 파일명 규칙으로 존재하는 서류만 반환 (상대경로)"""
    docs = {}
    for key, pattern in FILE_MAP.items():
        fname = pattern.format(name=name)
        # NFC/NFD 양쪽 확인 (macOS SMB 대응)
        for f in folder.iterdir():
            if _nfc(f.name) == _nfc(fname):
                docs[key] = fname
                break
    return docs


# ── 카드 HTML 조각 ────────────────────────────────────────────────────

def _pay_card(num: str, title: str, href: str) -> str:
    return f"""
      <div class="card-pay">
        <div class="card-left">
          <div class="card-sub">{num}</div>
          <div class="doc-title">{title}</div>
        </div>
        <a class="btn btn-red" href="{href}">납부하기</a>
      </div>"""


def _doc_card(num: str, icon: str, title: str, href: str) -> str:
    return f"""
      <div class="card">
        <div class="card-left">
          <div class="card-sub">{num}</div>
          <div class="doc-title">{icon} {title}</div>
        </div>
        <a class="btn btn-blue" href="{href}">확인하기</a>
      </div>"""


def _divider(text: str = "다음 서류") -> str:
    return f"""
      <div class="divider">
        <div class="divider-line"></div>
        <div class="divider-text">{text}</div>
        <div class="divider-line"></div>
      </div>"""


def _section_sep(text: str) -> str:
    return f"""
      <div class="sec-sep">
        <div class="sec-sep-line"></div>
        <div class="sec-sep-pill">{text}</div>
        <div class="sec-sep-line"></div>
      </div>"""


# ── 본문 카드 목록 생성 ───────────────────────────────────────────────

def _build_cards(docs: dict) -> str:
    """
    docs 순서: 종합소득세납부서 → 지방소득세납부서 → 접수증 → 신고서
    """
    pay_keys  = ["종합소득세납부서", "지방소득세납부서"]
    doc_keys  = ["접수증", "신고서"]
    icons     = {"접수증": "📄", "신고서": "📋"}
    pay_labels = {
        "종합소득세납부서": "종합소득세 납부서",
        "지방소득세납부서": "지방소득세 납부서",
    }
    doc_labels = {
        "접수증": "접수증",
        "신고서": "종합소득세 신고서",
    }

    pay_items = [(k, docs[k]) for k in pay_keys if docs.get(k)]
    doc_items = [(k, docs[k]) for k in doc_keys if docs.get(k)]

    total = len(pay_items) + len(doc_items)
    idx   = 1
    parts = []

    # ── 납부서 섹션 ──
    if pay_items:
        parts.append('\n      <div class="sec-label">💳 납부서</div>')
        for i, (key, href) in enumerate(pay_items):
            num = f"{idx} / {total}"
            parts.append(_pay_card(num, pay_labels[key], href))
            idx += 1
            if i < len(pay_items) - 1:
                parts.append(_divider("지방세도 함께"))
        parts.append(_section_sep("신고 서류"))

    # ── 서류 섹션 ──
    for i, (key, href) in enumerate(doc_items):
        num = f"{idx} / {total}"
        parts.append(_doc_card(num, icons[key], doc_labels[key], href))
        idx += 1
        if i < len(doc_items) - 1:
            parts.append(_divider())

    return "\n".join(parts)


# ── 전체 HTML ─────────────────────────────────────────────────────────

_CSS = """
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Apple SD Gothic Neo',
                 'Noto Sans KR', 'Malgun Gothic', sans-serif;
    background: #f0f2f5;
    min-height: 100vh;
    display: flex;
    justify-content: center;
    padding: 20px 16px 40px;
  }
  .container { width: 100%; max-width: 400px; }

  .header {
    background: white; border-radius: 18px;
    padding: 22px 20px 18px; margin-bottom: 14px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.07);
  }
  .brand {
    font-size: 11px; font-weight: 800; color: #2563eb;
    margin-bottom: 8px; display: flex; align-items: center; gap: 5px;
  }
  .brand::before {
    content:''; display:inline-block;
    width:6px; height:6px; border-radius:50%; background:#2563eb;
  }
  .title { font-size: 22px; font-weight: 900; color: #111; line-height: 1.3; margin-bottom: 6px; }
  .sub   { font-size: 13px; color: #666; }

  .sec-label { font-size: 11px; font-weight: 700; color: #aaa; padding: 0 4px 6px; }

  .card-pay {
    background: #fff5f5; border: 1.5px solid #fecaca;
    border-radius: 14px; padding: 14px 16px; margin-bottom: 6px;
    display: flex; align-items: center; justify-content: space-between; gap: 12px;
  }
  .card-pay .doc-title { font-size: 15px; font-weight: 800; color: #dc2626; }
  .card-pay .card-sub  { font-size: 10px; font-weight: 700; color: #fca5a5; margin-bottom: 2px; }

  .card {
    background: white; border-radius: 14px; padding: 14px 16px; margin-bottom: 6px;
    box-shadow: 0 1px 6px rgba(0,0,0,0.07);
    display: flex; align-items: center; justify-content: space-between; gap: 12px;
  }
  .card .doc-title { font-size: 15px; font-weight: 700; color: #111; }
  .card .card-sub  { font-size: 10px; font-weight: 700; color: #c0c0c0; margin-bottom: 2px; }

  .card-left { display: flex; flex-direction: column; gap: 1px; min-width: 0; }

  .btn {
    padding: 10px 16px; border-radius: 10px;
    font-size: 13px; font-weight: 700; text-decoration: none;
    white-space: nowrap; flex-shrink: 0; border: none; cursor: pointer;
  }
  .btn-red  { background: #ef4444; color: white; }
  .btn-blue { background: #2563eb; color: white; }

  .divider {
    display: flex; align-items: center; gap: 8px; margin: 4px 0 6px;
  }
  .divider-line { flex: 1; height: 1px; background: #e8e8e8; }
  .divider-text { font-size: 10px; font-weight: 600; color: #ccc; white-space: nowrap; }

  .sec-sep {
    display: flex; align-items: center; gap: 8px; margin: 6px 0 10px;
  }
  .sec-sep-line { flex: 1; height: 1px; background: #ddd; }
  .sec-sep-pill {
    font-size: 10px; font-weight: 700; color: #aaa;
    background: #e8e8e8; border-radius: 10px; padding: 3px 9px;
    white-space: nowrap;
  }

  .footer {
    text-align: center; font-size: 11px; color: #c0c0c0;
    font-weight: 600; padding: 20px 0 4px;
  }
"""


def make_landing_html(
    folder: Path,
    name: str,
    year: int,
    docs: dict,
) -> Path:
    """
    고객별 신고결과 랜딩 HTML 생성 후 folder에 저장.

    docs = {
        "접수증":           "파일명.pdf",   # 필수
        "신고서":           "파일명.pdf",   # 필수
        "종합소득세납부서": "파일명.pdf",   # 없으면 키 없거나 None
        "지방소득세납부서": "파일명.pdf",   # 없으면 키 없거나 None
    }
    반환: 생성된 HTML 파일 경로
    """
    cards_html = _build_cards(docs)

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>{year}년 종합소득세 신고결과 - 세무회계창연</title>
<style>{_CSS}</style>
</head>
<body>
<div class="container">
  <div class="header">
    <div class="brand">세무회계창연</div>
    <div class="title">{year}년 종합소득세<br>신고 완료</div>
    <div class="sub">{name}님의 신고가 완료되었습니다.</div>
  </div>
{cards_html}
  <div class="footer">세무회계창연</div>
</div>
</body>
</html>"""

    out = folder / f"신고결과_{name}.html"
    out.write_text(html, encoding="utf-8")
    return out


# ── 메인 진입점 (Airtable 트리거용) ──────────────────────────────────

def make_landing_html_auto(
    folder: Path,
    name: str,
    year: int,
    base_url: str,          # 예: "https://taxeng.co.kr/jongsotax/홍길동_800101"
) -> tuple[Path, str]:
    """
    폴더에서 파일 자동 감지 → HTML 생성 → (html_path, public_url) 반환

    n8n에서:
        POST /generate-landing  { name, folder_path, base_url }
        → { html_path, public_url }  ← 이 url을 SOLAPI #{링크}에 삽입
    """
    docs = detect_docs(folder, name)
    if not docs.get("접수증") and not docs.get("신고서"):
        raise FileNotFoundError(
            f"[{name}] 접수증/신고서 없음 — 폴더: {folder}"
        )

    html_path  = make_landing_html(folder, name, year, docs)
    html_fname = html_path.name
    public_url = f"{base_url.rstrip('/')}/{html_fname}"
    return html_path, public_url


# ── CLI 테스트 ─────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys, tempfile

    tmp = Path(tempfile.mkdtemp())
    name = "홍길동"

    # 테스트용 더미 PDF 생성
    for fname in [
        f"종합소득세 접수증 {name}.pdf",
        f"종합소득세 신고서 {name}.pdf",
        f"종합소득세 납부서 {name}.pdf",
        f"지방소득세 납부서 {name}.pdf",
    ]:
        (tmp / fname).write_bytes(b"%PDF-1.4 dummy")

    html_path, url = make_landing_html_auto(
        folder   = tmp,
        name     = name,
        year     = 2025,
        base_url = "https://taxeng.co.kr/jongsotax/홍길동_800101",
    )
    print(f"HTML: {html_path}")
    print(f"URL:  {url}")
    detected = detect_docs(tmp, name)
    print(f"감지: {list(detected.keys())}")
