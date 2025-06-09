@echo off
chcp 65001 > nul
cd /d "d:\OneDrive\Desktop\thu nghien software\production"
echo ===== EXCEL VALIDATOR - CHI TIET PIPE TREATMENT =====
echo 🚀 Đang khởi động validation tool mới...
echo 📋 Hiển thị chi tiết Array Number + Pipe Treatment
echo.
python excel_validator_detailed.py
echo.
echo ===== HOÀN THÀNH =====
echo Nhấn phím bất kỳ để đóng...
pause > nul
