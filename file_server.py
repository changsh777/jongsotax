"""
file_server.py - 토큰 기반 파일 다운로드 서버 (5월 종소세 시즌용)

실행: python3 file_server.py
포트: 8089

API:
  POST /api/token
    body: {"path": "/Users/changmini/NAS/종소세2026/고객/홍길동_800101/접수증.pdf", "name": "홍길동"}
    return: {"token": "xxx", "url": "https://files.taxenglab.com/download/xxx"}

  GET /download/{token}
    → 파일 다운로드 (7일 만료)

사용법:
  # 토큰 발급
  curl -X POST http://localhost:8089/api/token \
       -H "Content-Type: application/json" \
       -d '{"path": "/Users/changmini/NAS/.../접수증.pdf", "name": "홍길동"}'

  # 다운로드
  https://files.taxenglab.com/download/{token}
"""

import json
import uuid
import os
import mimetypes
from http.server import HTTPServer, BaseHTTPRequestHandler
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

# ===== 설정 =====
PORT = 8089
BASE_URL = "https://files.taxenglab.com"
TOKEN_EXPIRE_DAYS = 7
ALLOWED_BASE = "/Users/changmini/NAS/종소세2026/고객"  # 이 경로 밖은 서빙 금지

# ===== 토큰 저장소 (메모리 + JSON 파일로 영속) =====
TOKEN_DB_PATH = Path("/Users/changmini/NAS/종소세2026/_로그/tokens.json")

def load_tokens():
    if TOKEN_DB_PATH.exists():
        try:
            with open(TOKEN_DB_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_tokens(db):
    TOKEN_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(TOKEN_DB_PATH, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)

tokens = load_tokens()


def purge_expired():
    """만료된 토큰 정리"""
    now = datetime.now().isoformat()
    expired = [k for k, v in tokens.items() if v.get("expires", "") < now]
    for k in expired:
        del tokens[k]
    if expired:
        save_tokens(tokens)


def create_token(file_path: str, name: str) -> dict:
    """토큰 발급"""
    # 경로 보안 체크
    abs_path = str(Path(file_path).resolve())
    if not abs_path.startswith(ALLOWED_BASE):
        raise ValueError(f"허용 경로 외부: {abs_path}")
    if not Path(abs_path).exists():
        raise FileNotFoundError(f"파일 없음: {abs_path}")

    token = uuid.uuid4().hex
    expires = (datetime.now() + timedelta(days=TOKEN_EXPIRE_DAYS)).isoformat()
    tokens[token] = {
        "path": abs_path,
        "name": name,
        "filename": Path(abs_path).name,
        "created": datetime.now().isoformat(),
        "expires": expires,
    }
    save_tokens(tokens)
    return {
        "token": token,
        "url": f"{BASE_URL}/download/{token}",
        "expires": expires,
        "filename": Path(abs_path).name,
    }


class Handler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {self.address_string()} {format % args}")

    def send_json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if self.path == "/api/token":
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length)
            try:
                data = json.loads(body)
                file_path = data.get("path", "")
                name = data.get("name", "고객")
                result = create_token(file_path, name)
                self.send_json(200, result)
            except FileNotFoundError as e:
                self.send_json(404, {"error": str(e)})
            except ValueError as e:
                self.send_json(403, {"error": str(e)})
            except Exception as e:
                self.send_json(500, {"error": str(e)})
        else:
            self.send_json(404, {"error": "not found"})

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path.startswith("/download/"):
            token = path[len("/download/"):]
            purge_expired()

            info = tokens.get(token)
            if not info:
                self.send_response(404)
                self.end_headers()
                self.wfile.write("링크가 만료되었거나 존재하지 않습니다.".encode("utf-8"))
                return

            if info["expires"] < datetime.now().isoformat():
                self.send_response(410)
                self.end_headers()
                self.wfile.write("다운로드 링크가 만료되었습니다. (7일 초과)".encode("utf-8"))
                return

            file_path = Path(info["path"])
            if not file_path.exists():
                self.send_response(404)
                self.end_headers()
                self.wfile.write("파일을 찾을 수 없습니다.".encode("utf-8"))
                return

            # 파일 서빙
            mime, _ = mimetypes.guess_type(str(file_path))
            mime = mime or "application/octet-stream"
            filename = info["filename"]
            file_size = file_path.stat().st_size

            self.send_response(200)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Length", file_size)
            self.send_header(
                "Content-Disposition",
                f'attachment; filename="{filename}"'
            )
            self.end_headers()

            with open(file_path, "rb") as f:
                while chunk := f.read(65536):
                    self.wfile.write(chunk)

        elif path == "/health":
            self.send_json(200, {"status": "ok", "tokens": len(tokens)})

        else:
            self.send_response(404)
            self.end_headers()


if __name__ == "__main__":
    print(f"[파일 서버] 포트 {PORT} 시작")
    print(f"[파일 서버] 외부 URL: {BASE_URL}")
    print(f"[파일 서버] 허용 경로: {ALLOWED_BASE}")
    print(f"[파일 서버] 토큰 만료: {TOKEN_EXPIRE_DAYS}일")
    server = HTTPServer(("0.0.0.0", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[파일 서버] 종료")
