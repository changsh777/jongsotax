@echo off
chcp 65001 > nul
set HOMETAX_LIMIT=2
set SEOTAX_ENV=nas
cd /d F:\종소세2026
python hometax_result_scraper.py
