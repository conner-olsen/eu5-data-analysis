@echo off
cd /d "%~dp0src"
python scraper.py && python analyze.py
pause
