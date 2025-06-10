@echo off
chcp 65001 > nul
cd /d "d:\OneDrive\Desktop\thu nghien software\production"
echo ===== EXCEL VALIDATOR - RESTORED VERSION =====
echo 🚀 Đang khởi động validation tool đã khôi phục...
echo 📋 Chỉ bao gồm 2 validation rules cốt lõi:
echo    • Array Number Validation (EXP6 pattern matching)
echo    • Pipe Treatment Validation (CP-INTERNAL→GAL, Others→BLACK)
echo ✅ Đã khôi phục thành công về trạng thái trước PAP validation!
echo ❌ PAP Validation đã loại bỏ hoàn toàn
echo ❌ FAB Pipe Validation đã loại bỏ (clean up)
echo.
python excel_validator_detailed.py
echo.
echo ===== KHÔI PHỤC HOÀN THÀNH =====
echo 📊 Tool đã sẵn sàng với 2 quy tắc validation cốt lõi
echo Nhấn phím bất kỳ để đóng...
pause > nul
