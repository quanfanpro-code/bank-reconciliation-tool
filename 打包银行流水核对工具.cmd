@echo off
chcp 65001 >nul
set "APP_PYTHON=C:\Users\27651\AppData\Local\Programs\Python\Python314\python.exe"
cd /d "%~dp0"
"%APP_PYTHON%" -m PyInstaller --noconfirm --onefile --windowed --name "银行流水核对工具" --collect-all customtkinter --hidden-import psutil ^
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
echo 打包完成：%~dp0dist\银行流水核对工具.exe
