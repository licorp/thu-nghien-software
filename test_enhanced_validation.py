#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TEST ENHANCED VALIDATION LOGIC
==============================

Test các trường hợp cụ thể theo yêu cầu:
1. STD 1 PAP RANGE: size 65, chiều dài 4730, RG BE
2. STD 2 PAP RANGE: size 65, chiều dài 5295, RG BE  
3. STD ARRAY TEE: size 150, chiều dài 900, RG RG
4. Groove_Thread: RG, RG (còn trường hợp ống 40 TH, TH)
5. Fabrication: chỉ dành cho ống 65, RG BE
"""

import pandas as pd
import sys
import os

# Import validator
sys.path.append('.')
from excel_validator_final import ExcelValidator

def test_enhanced_validation():
    """
    Test các trường hợp validation theo yêu cầu mới
    """
    print("🧪 TESTING ENHANCED VALIDATION LOGIC")
    print("=" * 60)
    
    # Tạo validator
    validator = ExcelValidator()
    
    # Test cases theo yêu cầu
    test_cases = [
        {
            "name": "STD 1 PAP RANGE - Correct",
            "data": {
                "item_description": "Pipe-65-4730", 
                "size": "65", 
                "fab_pipe": "STD 1 PAP RANGE",
                "end_1": "RG", 
                "end_2": "BE"
            },
            "expected": "PASS"
        },
        {
            "name": "STD 1 PAP RANGE - Wrong FAB Pipe",
            "data": {
                "item_description": "Pipe-65-4730", 
                "size": "65", 
                "fab_pipe": "Wrong Value",
                "end_1": "RG", 
                "end_2": "BE"
            },
            "expected": "FAIL"
        },
        {
            "name": "STD 2 PAP RANGE - Correct",
            "data": {
                "item_description": "Pipe-65-5295", 
                "size": "65", 
                "fab_pipe": "STD 2 PAP RANGE",
                "end_1": "RG", 
                "end_2": "BE"
            },
            "expected": "PASS"
        },
        {
            "name": "STD ARRAY TEE - Correct",
            "data": {
                "item_description": "Fitting-150-900", 
                "size": "150", 
                "fab_pipe": "STD ARRAY TEE",
                "end_1": "RG", 
                "end_2": "RG"
            },
            "expected": "PASS"
        },
        {
            "name": "Groove_Thread RG-RG - Correct",
            "data": {
                "item_description": "Other Item", 
                "size": "80", 
                "fab_pipe": "Groove_Thread",
                "end_1": "RG", 
                "end_2": "RG"
            },
            "expected": "PASS"
        },
        {
            "name": "Groove_Thread Size 40 TH-TH - Correct",
            "data": {
                "item_description": "Other Item", 
                "size": "40", 
                "fab_pipe": "Groove_Thread",
                "end_1": "TH", 
                "end_2": "TH"
            },
            "expected": "PASS"
        },
        {
            "name": "Fabrication Size 65 RG-BE (not PAP) - Correct",
            "data": {
                "item_description": "Regular Pipe", 
                "size": "65", 
                "fab_pipe": "Fabrication",
                "end_1": "RG", 
                "end_2": "BE"
            },
            "expected": "PASS"
        },
        {
            "name": "Priority Test: Size 65 + 4730 should be STD 1 PAP (not Fabrication)",
            "data": {
                "item_description": "Pipe-65-4730", 
                "size": "65", 
                "fab_pipe": "Fabrication",  # Wrong - should be STD 1 PAP RANGE
                "end_1": "RG", 
                "end_2": "BE"
            },
            "expected": "FAIL"
        }
    ]
    
    # Chạy test
    print("\n🔍 RUNNING TEST CASES:")
    print("-" * 60)
    
    for i, test_case in enumerate(test_cases, 1):
        print(f"\n{i}. {test_case['name']}")
        print(f"   Data: {test_case['data']}")
        
        # Tạo mock row
        row = pd.Series(test_case['data'])
        
        # Chạy validation
        result = validator._check_pipe_schedule_mapping(
            row, 
            'item_description',
            'size', 
            'fab_pipe',
            'end_1',
            'end_2'
        )
        
        # Kiểm tra kết quả
        is_pass = result == "PASS"
        expected_pass = test_case['expected'] == "PASS"
        
        if is_pass == expected_pass:
            status = "✅ CORRECT"
        else:
            status = "❌ INCORRECT"
        
        print(f"   Result: {result}")
        print(f"   Expected: {test_case['expected']} | Actual: {'PASS' if is_pass else 'FAIL'} | {status}")
    
    print("\n" + "=" * 60)
    print("🎯 TEST COMPLETED")

if __name__ == "__main__":
    test_enhanced_validation()
