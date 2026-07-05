@echo off
cd /d "%~dp0"
py app.py
if errorlevel 1 (
  echo.
  echo The app could not start.
  echo First run Install_Requirements.bat, then try again.
  echo If it still fails, copy the error text and send it to ChatGPT.
  pause
)
