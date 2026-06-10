@echo off
chcp 65001 >nul
echo ========================================================
echo        CÀI ĐẶT MÔI TRƯỜNG CHO AI CHAT RAG
echo ========================================================
echo.

:: 1. Kiểm tra Python
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [LỖI] Không tìm thấy Python! Vui lòng cài đặt Python 3.10+ và nhớ tích vào "Add Python to PATH" khi cài đặt.
    echo Mở trang tải: https://www.python.org/downloads/
    pause
    exit /b
)
echo [OK] Đã phát hiện Python.

:: 2. Tạo Virtual Environment
echo.
echo [1/3] Đang tạo môi trường ảo (Virtual Environment)...
IF NOT EXIST "venv" (
    python -m venv venv
    echo [OK] Đã tạo thư mục venv.
) ELSE (
    echo [OK] Thư mục venv đã tồn tại.
)

:: 3. Kích hoạt và cài đặt thư viện
echo.
echo [2/3] Đang tải và cài đặt thư viện từ requirements.txt (Có thể mất vài phút)...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo [LỖI] Xảy ra lỗi trong quá trình cài đặt thư viện! Vui lòng kiểm tra lại kết nối mạng.
    pause
    exit /b
)
echo [OK] Đã cài đặt thành công tất cả thư viện.

:: 4. Hoàn tất
echo.
echo [3/3] HOÀN TẤT CÀI ĐẶT!
echo ========================================================
echo Bạn có thể khởi động LM Studio Local Server, 
echo sau đó chạy file "start.bat" để bắt đầu sử dụng.
echo ========================================================
pause
