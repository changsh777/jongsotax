"""스크래퍼를 새 프로세스로 시작"""
import subprocess, os, sys

env = os.environ.copy()
env["HOMETAX_LIMIT"] = "1"
env["SEOTAX_ENV"] = "nas"

p = subprocess.Popen(
    [sys.executable, "hometax_result_scraper.py"],
    cwd=r"F:\종소세2026",
    env=env,
    # stdout/stderr를 로그 파일로
    stdout=open(r"C:\Users\pc\종소세2026\scraper_stdout.log", "w", encoding="utf-8"),
    stderr=subprocess.STDOUT,
    creationflags=subprocess.CREATE_NEW_CONSOLE,  # 별도 콘솔 창
)
print(f"PID: {p.pid}")
