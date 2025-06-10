@echo off
chcp 65001 > nul
cd /d "d:\OneDrive\Desktop\thu nghien software\production"
echo ===== EXCEL VALIDATOR TOOL - CHI TIET PIPE TREATMENT =====
echo.
echo 🚀 Chạy validation tool mới với hiển thị chi tiết Array Number và Pipe Treatment
echo.
python excel_validator_detailed.py
echo.
pause
