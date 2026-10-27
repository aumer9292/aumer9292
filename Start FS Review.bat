@echo off
title FS Review
rem Double-click this file to open FS Review in your web browser.
cd /d "%~dp0"

set PY=
where py >nul 2>nul && set PY=py -3
if not defined PY (where python >nul 2>nul && set PY=python)
if not defined PY (
  echo.
  echo  Python is not installed on this computer.
  echo  1. Download it from https://www.python.org/downloads/  ^(version 3.10 or later^)
  echo  2. During setup tick "Add python.exe to PATH".
  echo  3. Double-click "Start FS Review.bat" again.
  echo.
  start https://www.python.org/downloads/
  pause
  exit /b 1
)

if not exist ".installed" (
  echo  First run: installing the components FS Review needs. This takes a few minutes, once.
  %PY% -m pip install --upgrade pip >nul
  %PY% -m pip install -r requirements.txt || (
    echo.
    echo  Installation failed. Check the internet connection and try again.
    pause
    exit /b 1
  )
  echo done> ".installed"
)

echo.
echo  FS Review is starting. Your browser will open at http://127.0.0.1:8765
echo  Keep this window open while you use it. Close it to stop FS Review.
echo.
%PY% -m fsreview --serve
pause
