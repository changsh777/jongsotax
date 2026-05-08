@echo off
echo Edge를 CDP 모드로 재시작합니다...
echo 로그인 쿠키가 유지되므로 홈택스 재로그인 불필요합니다.
echo.

REM 현재 Edge 종료 (홈택스 탭 포함)
taskkill /IM msedge.exe /F >nul 2>&1
timeout /t 2 >nul

REM CDP + 팝업허용 + 원격허용 플래그로 재시작
start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" ^
    --remote-debugging-port=9222 ^
    --remote-allow-origins=* ^
    --disable-popup-blocking ^
    --user-data-dir="C:\EdgeDebug" ^
    "https://hometax.go.kr"

echo Edge 재시작 완료!
echo 홈택스 로그인 완료 후 hometax_result_scraper.py 실행하세요.
pause
