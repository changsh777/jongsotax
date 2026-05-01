@echo off
REM 홈택스 자동화용 크롬 실행 (디버그 포트 9222 + 전용 프로파일)
REM 평소 쓰는 크롬은 그대로 두고, 이 배치만 실행하면 됨

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="F:\종소세2026\.chrome_debug_profile" ^
  https://hometax.go.kr
