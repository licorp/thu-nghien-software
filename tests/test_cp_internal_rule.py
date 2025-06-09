#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TEST CP-INTERNAL RULE - Kiểm tra logic validation mới
=====================================================

Test cả 2 rules cho Array Number validation:
1. Rule CP-INTERNAL: Array Number = Cross Passage  
2. Rule Pattern: Array Number chứa "EXP6" + 2 số cuối cột B + 2 số cuối cột A

Tác giả: GitHub Copilot
Ngày: 2025-06-09
"""

import pandas as pd
import re

def test_cp_internal_rule():
    """
    Test logic validation mới cho CP-INTERNAL rule
    """
    
    print("=== TEST CP-INTERNAL RULE VALIDATION ===\n")
    
    # Test cases cho cả 2 rules
    test_cases = [
        # Rule 1: CP-INTERNAL cases
        {
            'name': 'CP-INTERNAL - PASS case',
            'cross_passage': 'EXP61001',        # Cột A
            'location_lanes': 'M110',           # Cột B  
            'system_type': 'CP-INTERNAL',       # Cột C
            'array_number': 'EXP61001',         # Cột D = Cột A
            'expected_result': 'PASS (Rule: CP-INTERNAL)',
        },
        {
            'name': 'CP-INTERNAL - FAIL case',
            'cross_passage': 'EXP61001',        # Cột A
            'location_lanes': 'M110',           # Cột B  
            'system_type': 'CP-INTERNAL',       # Cột C
            'array_number': 'EXP61003',         # Cột D ≠ Cột A
            'expected_result': 'FAIL (Rule: CP-INTERNAL)',
        },
        
        # Rule 2: Pattern cases (non CP-INTERNAL)
        {
            'name': 'CP-EXTERNAL - Pattern PASS',
            'cross_passage': 'EXP61002',        # Cột A → 02
            'location_lanes': 'M110',           # Cột B → 10
            'system_type': 'CP-EXTERNAL',       # Cột C
            'array_number': 'EXP61002',         # Cột D chứa EXP61002
            'expected_result': 'PASS (Rule: Pattern)',
        },
        {
            'name': 'CW-DISTRIBUTION - Pattern FAIL',
            'cross_passage': 'EXP61002',        # Cột A → 02
            'location_lanes': 'M111',           # Cột B → 11  
            'system_type': 'CW-DISTRIBUTION',   # Cột C
            'array_number': 'EXP61002',         # Cột D cần EXP61102 nhưng có EXP61002
            'expected_result': 'FAIL (Rule: Pattern)',
        },
        {
            'name': 'No System Type - Pattern logic',
            'cross_passage': 'EXP61003',        # Cột A → 03
            'location_lanes': 'M112',           # Cột B → 12
            'system_type': None,                # Cột C = NA
            'array_number': 'EXP61203',         # Cột D chứa EXP61203  
            'expected_result': 'PASS (Rule: Pattern)',
        }
    ]
    
    for i, test in enumerate(test_cases, 1):
        print(f"🧪 Test {i}: {test['name']}")
        print(f"   📍 Cross Passage (A): {test['cross_passage']}")
        print(f"   📍 Location Lanes (B): {test['location_lanes']}")
        print(f"   📍 System Type (C): {test['system_type']}")
        print(f"   📍 Array Number (D): {test['array_number']}")
        
        # Simulate validation logic từ excel_validator_detailed.py
        result = simulate_array_validation(
            test['cross_passage'], 
            test['location_lanes'], 
            test['system_type'], 
            test['array_number']
        )
        
        print(f"   📍 Expected: {test['expected_result']}")
        print(f"   📍 Result: {result}")
        
        # Check if result contains expected rule type
        if test['expected_result'].startswith('PASS') and result.startswith('PASS'):
            status = "✅ PASS"
        elif test['expected_result'].startswith('FAIL') and result.startswith('FAIL'):
            status = "✅ PASS" 
        else:
            status = "❌ FAIL"
            
        print(f"   📍 Status: {status}")
        print()

def simulate_array_validation(cross_passage, location_lanes, system_type, array_number):
    """
    Simulate logic từ _check_array_number_detailed function
    """
    try:
        if pd.isna(cross_passage) or pd.isna(location_lanes) or pd.isna(array_number):
            return "SKIP: Thiếu dữ liệu"
        
        actual_array = str(array_number).strip()
        cross_passage_str = str(cross_passage).strip()
        
        # RULE 1: CP-INTERNAL check
        if pd.notna(system_type):
            system_type_str = str(system_type).upper().strip()
            if system_type_str == 'CP-INTERNAL':
                # Rule mới: Array Number phải bằng Cross Passage
                if actual_array == cross_passage_str:
                    return "PASS (Rule: CP-INTERNAL)"
                else:
                    return f"FAIL (Rule: CP-INTERNAL): cần Array=Cross, có '{actual_array}' ≠ '{cross_passage_str}'"
        
        # RULE 2: Pattern check (cho tất cả các case khác)
        # Lấy 2 số cuối của cột B
        location_str = str(location_lanes).strip()
        numbers_in_location = re.findall(r'\d+', location_str)
        if numbers_in_location:
            last_2_b = numbers_in_location[-1][-2:] if len(numbers_in_location[-1]) >= 2 else numbers_in_location[-1].zfill(2)
        else:
            last_2_b = "00"
        
        # Lấy 2 số cuối của cột A
        cross_str = str(cross_passage).strip()
        numbers_in_cross = re.findall(r'\d+', cross_str)
        if numbers_in_cross:
            last_2_a = numbers_in_cross[-1][-2:] if len(numbers_in_cross[-1]) >= 2 else numbers_in_cross[-1].zfill(2)
        else:
            last_2_a = "00"
        
        # Tạo pattern bắt buộc
        required_pattern = f"EXP6{last_2_b}{last_2_a}"
        
        if required_pattern in actual_array:
            return "PASS (Rule: Pattern)"
        else:
            return f"FAIL (Rule: Pattern): cần '{required_pattern}', có '{actual_array}'"
            
    except Exception as e:
        return f"ERROR: {str(e)}"

if __name__ == "__main__":
    test_cp_internal_rule()
