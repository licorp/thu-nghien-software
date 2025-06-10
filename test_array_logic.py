#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TEST ARRAY NUMBER LOGIC
======================
"""

import pandas as pd
import re

def test_array_number_logic():
    """
    Test case cụ thể cho Array Number validation
    """
    print("🧪 TESTING ARRAY NUMBER LOGIC")
    print("=" * 50)
    
    # Test case từ ví dụ của bạn
    test_cases = [
        {
            'cross_passage': 'B', 
            'location_lanes': '03', 
            'array_number': 'EXP61103B',
            'expected': 'EXP61103',
            'should_pass': True
        },
        {
            'cross_passage': 'A', 
            'location_lanes': '02', 
            'array_number': 'EXP61102',
            'expected': 'EXP61102', 
            'should_pass': True
        },
        {
            'cross_passage': 'C', 
            'location_lanes': '04', 
            'array_number': 'EXP61205',  # Sai
            'expected': 'EXP61104',
            'should_pass': False
        },
        {
            'cross_passage': 'D', 
            'location_lanes': '05', 
            'array_number': 'EXP61105X',  # Có thêm ký tự nhưng chứa pattern đúng
            'expected': 'EXP61105',
            'should_pass': True
        }
    ]
    
    for i, case in enumerate(test_cases, 1):
        print(f"\n📋 TEST CASE {i}:")
        print(f"   Cross Passage: {case['cross_passage']}")
        print(f"   Location Lanes: {case['location_lanes']}")
        print(f"   Array Number: {case['array_number']}")
        print(f"   Expected Pattern: {case['expected']}")
        
        # Áp dụng logic mới
        cross_passage_str = str(case['cross_passage'])
        location_lanes_str = str(case['location_lanes'])
        array_number_str = str(case['array_number']).strip()
        
        # Extract 2 digits cuối
        cross_digits = re.findall(r'\d', cross_passage_str)[-2:] if len(re.findall(r'\d', cross_passage_str)) >= 2 else ['0', '0']
        location_digits = re.findall(r'\d', location_lanes_str)[-2:] if len(re.findall(r'\d', location_lanes_str)) >= 2 else ['0', '0']
        
        # Tạo expected array number
        expected_array = f"EXP6{''.join(location_digits)}{''.join(cross_digits)}"
        
        # Kiểm tra logic mới: có chứa pattern không?
        result = expected_array in array_number_str
        
        print(f"   Generated Expected: {expected_array}")
        print(f"   Contains Check: '{expected_array}' in '{array_number_str}' = {result}")
        
        if result == case['should_pass']:
            print(f"   ✅ PASS - Logic đúng!")
        else:
            print(f"   ❌ FAIL - Logic sai!")
    
    print("\n" + "=" * 50)
    print("✅ TEST HOÀN THÀNH!")

if __name__ == "__main__":
    test_array_number_logic()
