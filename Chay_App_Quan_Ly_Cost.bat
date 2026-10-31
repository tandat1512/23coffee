@echo off
title 23 COFFEE - He Thong Quan Ly Cost & Menu F&B
chcp 65001 >nul
cd /d "%~dp0"

echo ===============================================================
echo     ☕ 23 COFFEE — HE THONG QUAN LY COST & MENU TU DONG
echo ===============================================================
echo.
echo Đang khởi động ứng dụng...
echo Trình duyệt web sẽ tự động mở tại: http://localhost:5000
echo.
echo [TÀI KHOẢN ĐĂNG NHẬP]
echo  - Tài khoản 1: tandat  / Mật khẩu: tandat@23coffee
echo  - Tài khoản 2: partner / Mật khẩu: partner@23coffee
echo.
echo [LƯU Ý] Hãy giữ cửa sổ này chạy trong suốt quá trình sử dụng.
echo Khi muốn tắt ứng dụng, chỉ cần đóng cửa sổ này lại.
echo ===============================================================
echo.

start "" "http://localhost:5000"
python app.py

pause
