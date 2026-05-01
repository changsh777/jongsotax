@echo off
chcp 65001 >nul
start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="F:\종소세2026\.edge_debug_profile" ^
  https://hometax.go.kr
