@echo off
rem Double-click to open FS Review in your browser. Needs Python 3.10+ from python.org.
cd /d "%~dp0"
python -m pip install -q -r requirements.txt
python -m fsreview --serve
pause
