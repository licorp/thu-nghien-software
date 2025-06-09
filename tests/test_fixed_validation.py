#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re

def test_array_validation_fixed():
    """
    Test logic validation mới với 2 số cuối cột A
    """
    
    # Test cases dựa trên data thực tế
    test_cases = [
        {
            'name': 'Test case 1 - Dữ liệu thực tế',
            'cross_passage': 'EXP61002',    # Cột A
            'location_lanes': 'M110',       # Cột B  
            'array_number': 'EXP61002',     # Cột D
            'expected_pattern': 'EXP61002', # EXP6 + 10 + 02
        },
        {
            'name': 'Test case 2 - Dữ liệu thực tế khác',
            'cross_passage': 'EXP61002',    # Cột A  
            'location_lanes': 'M111',       # Cột B
            'array_number': 'EXP61102',     # Cột D
            'expected_pattern': 'EXP61102', # EXP6 + 11 + 02
        },
        {
            'name': 'Test case 3 - Pattern có thêm ký tự',
            'cross_passage': 'EXP61003',    # Cột A
            'location_lanes': 'M112',       # Cột B  
            'array_number': 'EXP61203ABC',  # Cột D có thêm ký tự
            'expected_pattern': 'EXP61203', # EXP6 + 12 + 03
        }
    ]
    
    print("=== TEST LOGIC VALIDATION MỚI (2 SỐ CUỐI COLT A) ===\n")
    
    for i, test in enumerate(test_cases, 1):
        print(f"🧪 {test['name']}")
        
        # Logic giống như trong validate_real.py
        cross_passage = test['cross_passage']
        location_lanes = test['location_lanes'] 
        array_number = test['array_number']
        
        # Lấy 2 số cuối của cột B
        location_str = str(location_lanes).strip()
        numbers_in_location = re.findall(r'\d+', location_str)
        if numbers_in_location:
            last_2_b = numbers_in_location[-1][-2:] if len(numbers_in_location[-1]) >= 2 else numbers_in_location[-1].zfill(2)
        else:
            last_2_b = "00"
        
        # Lấy 2 số cuối của cột A (đã sửa từ 3 thành 2)
        cross_str = str(cross_passage).strip()
        numbers_in_cross = re.findall(r'\d+', cross_str)  
        if numbers_in_cross:
            last_2_a = numbers_in_cross[-1][-2:] if len(numbers_in_cross[-1]) >= 2 else numbers_in_cross[-1].zfill(2)
        else:
            last_2_a = "00"
        
        # Tạo pattern bắt buộc
        required_pattern = f"EXP6{last_2_b}{last_2_a}"
        actual_array = str(array_number).strip()
        
        # Kiểm tra
        is_valid = required_pattern in actual_array
        
        print(f"   📍 Cột A: {cross_passage} → 2 số cuối: {last_2_a}")
        print(f"   📍 Cột B: {location_lanes} → 2 số cuối: {last_2_b}")
        print(f"   📍 Pattern bắt buộc: {required_pattern}")
        print(f"   📍 Cột D thực tế: {actual_array}")
        print(f"   📍 Kết quả: {'✅ PASS' if is_valid else '❌ FAIL'}")
        
        if test['expected_pattern'] == required_pattern:
            print(f"   📍 Logic đúng: Expected={test['expected_pattern']}, Got={required_pattern}")
        else:
            print(f"   📍 Logic sai: Expected={test['expected_pattern']}, Got={required_pattern}")
        
        print()

if __name__ == "__main__":
    test_array_validation_fixed()
