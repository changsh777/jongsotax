"""
landing_server.py — 신고결과 랜딩 HTML 생성 FastAPI 서버
세무회계창연 | 2026

포트: 8766
n8n → http://host.docker.internal:8766/generate-landing
       POST { "name": "홍길동", "phone": "010-xxxx-xxxx" }
       → { "public_url": "https://...", "docs_found": [...] }

실행:
    python3 landing_server.py

Mac Mini 백그라운드:
    nohup python3 landing_server.py >> ~/종소세2026/landing_server.log 2>&1 &
"""

import os
import sys
import unicodedata
from pathlib import Path

from fastapi import FastAPI, HTTPException
from pydantic import BaseModel

sys.path.insert(0, str(Path(__file__).parent))
from landing_gen import make_landing_html_auto, detect_docs

# ── 설정 (환경변수로 오버라이드 가능) ─────────────────────────────────
CUSTOMER_DIR = Path(os.getenv(
    "CUSTOMER_DIR",
    "/Volumes/장성환/종소세2026/고객"
))
BASE_URL = os.getenv(
    "BASE_URL",
    "https://taxeng.co.kr/jongsotax"
)
TAX_YEAR = int(os.getenv("TAX_YEAR", "2025"))
PORT     = int(os.getenv("LANDING_PORT", "8766"))

app = FastAPI(title="신고결과 랜딩 생성기", version="1.0.0")


# ── 유틸 ──────────────────────────────────────────────────────────────

def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def find_folder(name: str) -> Path | None:
    """이름으로 고객 폴더 찾기 (이름_ 로 시작하는 폴더)"""
    nfc_name = _nfc(name)
    if not CUSTOMER_DIR.exists():
        return None
    candidates = [
        p for p in CUSTOMER_DIR.iterdir()
        if p.is_dir() and _nfc(p.name).startswith(f"{nfc_name}_")
    ]
    if not candidates:
        exact = CUSTOMER_DIR / name
        return exact if exact.is_dir() else None
    return sorted(candidates)[0]   # 여러 개면 첫 번째


# ── 스키마 ────────────────────────────────────────────────────────────

class LandingRequest(BaseModel):
    name:  str
    phone: str = ""   # Airtable에서 넘어오는 전화번호 (현재 로그용)


class LandingResponse(BaseModel):
    public_url: str
    html_path:  str
    docs_found: list[str]
    name:       str
    folder:     str


# ── 엔드포인트 ────────────────────────────────────────────────────────

@app.post("/generate-landing", response_model=LandingResponse)
def generate_landing(req: LandingRequest):
    """
    1. 이름으로 고객 폴더 찾기
    2. 폴더에서 서류 PDF 자동 감지
    3. 랜딩 HTML 생성
    4. public_url 반환 → n8n이 SOLAPI에 전달
    """
    folder = find_folder(req.name)
    if not folder:
        raise HTTPException(
            status_code=404,
            detail=f"고객 폴더 없음: {req.name} (CUSTOMER_DIR: {CUSTOMER_DIR})"
        )

    docs = detect_docs(folder, req.name)
    if not docs.get("접수증") and not docs.get("신고서"):
        raise HTTPException(
            status_code=404,
            detail=f"접수증/신고서 없음 — 폴더: {folder}, 감지된 서류: {list(docs.keys())}"
        )

    folder_name = folder.name
    base_url    = f"{BASE_URL}/{folder_name}"

    try:
        html_path, public_url = make_landing_html_auto(
            folder   = folder,
            name     = req.name,
            year     = TAX_YEAR,
            base_url = base_url,
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    return LandingResponse(
        public_url = public_url,
        html_path  = str(html_path),
        docs_found = list(docs.keys()),
        name       = req.name,
        folder     = str(folder),
    )


@app.get("/health")
def health():
    return {
        "status":       "ok",
        "customer_dir": str(CUSTOMER_DIR),
        "base_url":     BASE_URL,
        "tax_year":     TAX_YEAR,
    }


# ── 실행 ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=PORT, reload=False)
