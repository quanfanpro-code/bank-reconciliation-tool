@echo off
chcp 65001 >nul
set "APP_PYTHON=C:\Users\27651\AppData\Local\Programs\Python\Python314\python.exe"
if /I "%~1"=="--check" (
  if exist "%APP_PYTHON%" (
    "%APP_PYTHON%" -c "import gui; print('READY')"
    exit /b %ERRORLEVEL%
  )
  echo Python not found
  exit /b 1
)
cd /d "%~dp0"
"%APP_PYTHON%" "%~dp0main.py"
if errorlevel 1 pause
