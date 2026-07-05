@echo off
cd /d "%~dp0"
py --version
if errorlevel 1 (
  echo Python is not installed or the Python launcher is unavailable.
  echo Install Python from https://www.python.org/downloads/windows/
  pause
  exit /b 1
)
py -m pip install --upgrade pip
py -m pip install -r requirements.txt
pause
