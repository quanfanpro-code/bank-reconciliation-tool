@echo off
chcp 65001 >nul
setlocal
set "APP_PYTHON=C:\Users\27651\AppData\Local\Programs\Python\Python314\python.exe"
cd /d "%~dp0"
if not exist "%LOCALAPPDATA%\银行流水核对工具\大模型默认配置.json" (
  echo 未找到本机大模型默认配置，无法生成内置配置版。
  exit /b 1
)
if not defined BANK_BUILD_ROOT set "BANK_BUILD_ROOT=%TEMP%\银行流水打包_%RANDOM%_%RANDOM%"
if not defined BANK_DIST_DIR set "BANK_DIST_DIR=%~dp0dist"
set "PYINSTALLER_CONFIG_DIR=%BANK_BUILD_ROOT%\cache"
"%APP_PYTHON%" -m PyInstaller --noconfirm --onefile --windowed --name "银行流水核对工具" --workpath "%BANK_BUILD_ROOT%\build" --specpath "%BANK_BUILD_ROOT%" --distpath "%BANK_DIST_DIR%" --add-data "%LOCALAPPDATA%\银行流水核对工具\大模型默认配置.json:." --collect-all customtkinter --hidden-import psutil ^
  --exclude-module torch --exclude-module torchvision --exclude-module torchaudio ^
  --exclude-module scipy --exclude-module matplotlib --exclude-module pyarrow ^
  --exclude-module pytest --exclude-module spacy --exclude-module thinc ^
  --exclude-module transformers --exclude-module datasets --exclude-module cv2 ^
  --exclude-module av --exclude-module yt_dlp --exclude-module boto3 ^
  --exclude-module botocore --exclude-module fsspec --exclude-module soundfile main.py
if errorlevel 1 (
  echo 打包失败
  pause
  exit /b 1
)
echo 打包完成：%BANK_DIST_DIR%\银行流水核对工具.exe
echo 本次构建临时目录：%BANK_BUILD_ROOT%
endlocal
