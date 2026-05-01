"""
디버그 모드 엣지 실행 (Python 런처)
- 배치파일 인코딩 이슈 우회
- 한글 경로 안전하게 처리
"""
import subprocess
from pathlib import Path

EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
USER_DATA = Path(r"F:\종소세2026\.edge_debug_profile")
USER_DATA.mkdir(parents=True, exist_ok=True)

args = [
    EDGE,
    "--remote-debugging-port=9222",
    f"--user-data-dir={USER_DATA}",
    "https://hometax.go.kr",
]

print("[엣지 실행 중...]")
subprocess.Popen(args)
print(f"[OK] 디버그 포트: 9222")
print(f"[OK] 프로파일: {USER_DATA}")
print()
print("이 창은 닫아도 됩니다. 엣지에서 홈택스 로그인 후")
print("다른 CMD에서 다음 실행:")
print("    python F:\\종소세2026\\step1_attach.py")
