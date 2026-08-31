@echo off
title Review Scraper - Walmart Export
cd /d "%~dp0"
echo Installing/checking dependency (openpyxl)...
python -m pip install --quiet openpyxl
echo Starting app... a browser tab will open at http://127.0.0.1:8765/
echo Keep this window open while you work. Close it (or Ctrl+C) to stop.
python app.py
pause
