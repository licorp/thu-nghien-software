#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from excel_validator_detailed import ExcelValidatorDetailed

def test_validation():
    """Test PAP validation nhanh"""
    print("🚀 TESTING PAP VALIDATION...")
    
    # Khởi tạo validator
    validator = ExcelValidatorDetailed()
    
    # File Excel để test
    excel_file = r"d:\OneDrive\Desktop\thu nghien software\Xp03-Fabrication & Listing.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"❌ Không tìm thấy file: {excel_file}")
        return
    
    print(f"📁 File: {excel_file}")
    
    try:
        # Chạy validation
        validator.validate_excel_file(excel_file)
        print("✅ Validation hoàn thành!")
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_validation()
