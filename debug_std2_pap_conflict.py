#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DEBUG STD 2 PAP RANGE vs FABRICATION CONFLICT
=============================================

Debug case cụ thể: size 65, item 65-5295, RG BE
Kết quả mong muốn: STD 2 PAP RANGE (ưu tiên cao)
Kết quả hiện tại: Fabrication (có thể bị conflict)
"""

import pandas as pd
import sys
import os

# Import validator
sys.path.append('.')
from excel_validator_final import ExcelValidator

def debug_std2_pap_case():
    """
    Debug case cụ thể từ dòng 21 Excel
    """
    print("🐛 DEBUG: STD 2 PAP RANGE vs FABRICATION CONFLICT")
    print("=" * 60)
    
    # Tạo validator
    validator = ExcelValidator()
    
    # Case từ dòng 21 Excel
    test_case = {
        "name": "DÒNG 21 EXCEL: STD 2 PAP RANGE Case",
        "data": {
            "item_description": "65-5295",  # Chứa 5295
            "size": "65",                   # Size 65  
            "fab_pipe": "STD 2 PAP RANGE", # FAB Pipe hiện tại
            "end_1": "RG",                  # End-1 = RG
            "end_2": "BE"                   # End-2 = BE
        },
        "expected": "PASS"  # Phải PASS vì đúng STD 2 PAP RANGE
    }
    
    print(f"🔍 Test Case: {test_case['name']}")
    print(f"📊 Data: {test_case['data']}")
    print()
    
    # Tạo mock row
    row = pd.Series(test_case['data'])
    
    # Chạy validation
    print("🚀 Chạy validation...")
    result = validator._check_pipe_schedule_mapping(
        row, 
        'item_description',
        'size', 
        'fab_pipe',
        'end_1',
        'end_2'
    )
    
    print(f"📝 Kết quả: {result}")
    print()
    
    # Phân tích kết quả
    is_pass = result == "PASS"
    expected_pass = test_case['expected'] == "PASS"
    
    if is_pass == expected_pass:
        print("✅ CORRECT: Kết quả đúng theo mong muốn")
    else:
        print("❌ INCORRECT: Có conflict logic!")
        print(f"   Expected: {test_case['expected']}")
        print(f"   Actual: {'PASS' if is_pass else 'FAIL'}")
        print(f"   Chi tiết lỗi: {result}")
    
    print()
    print("🔍 PHÂN TÍCH LOGIC:")
    print("=" * 30)
    print("✅ ƯU TIÊN CAO: STD 2 PAP RANGE (size 65, 5295, RG BE)")
    print("❌ ƯU TIÊN THẤP: Fabrication (65, RG BE - nhưng không phải PAP)")
    print()
    print("👉 CASE NÀY PHẢI: STD 2 PAP RANGE vì có 5295 trong item description!")

if __name__ == "__main__":
    debug_std2_pap_case()
