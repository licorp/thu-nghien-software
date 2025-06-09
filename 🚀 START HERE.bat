@echo off
chcp 65001 > nul
title Excel Validator - 3 Rules Complete ✅
color 0A

echo.
echo =======================================================
echo    🚀 EXCEL DATA VALIDATION TOOL - PRODUCTION READY ✅
echo =======================================================
echo.
echo 🎯 LAUNCHING VALIDATION WITH 3 COMPLETE RULES...
echo    • Rule 1: Array Number format validation
echo    • Rule 2: Pipe Treatment validation  
echo    • Rule 3: CP-INTERNAL matching validation
echo.
echo ⚡ SMART PRIORITY: CP-INTERNAL có logic ưu tiên đặc biệt
echo 📊 ERROR DISPLAY: Màu đỏ cho giá trị SAI, màu trắng cho giá trị ĐÚNG
echo.

python excel_validator_final.py

echo.
echo =======================================================
echo 🎉 Validation complete! Check output file for results.
echo =======================================================
pause
