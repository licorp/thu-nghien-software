@echo off
chcp 65001 > nul
echo ===== TESTING PAP VALIDATION =====
echo Starting validation test...
echo.
python test_simple.py
echo.
echo ===== COMPLETE =====
pause
