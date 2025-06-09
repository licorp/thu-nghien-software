@echo off
chcp 65001 > nul
title Excel Validator - Quick Launcher
color 0A

echo.
echo ================================
echo    Excel Data Validation Tool
echo ================================
echo.
echo Starting Excel Validator...
echo.

cd production
call Excel_Validator.bat

echo.
echo Returning to main directory...
cd ..
pause
