@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

:: Kích hoạt môi trường ảo nếu đã cài đặt bằng install.bat
IF EXIST "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

python chat.py
pause
