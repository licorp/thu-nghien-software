import pandas as pd
import re

def test_array_number_validation():
    """
    Test validation logic cho EE_Array Number
    """
    
    # Test cases mẫu
    test_cases = [
        {
            'EE_Cross Passage': 'Cross123',  # Cột A - lấy 3 số cuối: 123
            'EE_Location and Lanes': 'Lane45',  # Cột B - lấy 2 số cuối: 45  
            'EE_Array Number': 'EXP645123',  # Cột D - phải chứa EXP645123
            'expected': 'PASS'
        },
        {
            'EE_Cross Passage': 'Cross123',  
            'EE_Location and Lanes': 'Lane45',  
            'EE_Array Number': 'EXP645123-Version2',  # Có thêm text sau
            'expected': 'PASS'
        },
        {
            'EE_Cross Passage': 'Cross123',  
            'EE_Location and Lanes': 'Lane45',  
            'EE_Array Number': 'EXP645999',  # Sai số cuối
            'expected': 'FAIL'
        },
        {
            'EE_Cross Passage': 'ABC789',  # 3 số cuối: 789
            'EE_Location and Lanes': 'XYZ12',  # 2 số cuối: 12
            'EE_Array Number': 'EXP612789ABC',  # Chứa EXP612789 + thêm ABC
            'expected': 'PASS'
        },
        {
            'EE_Cross Passage': 'Test56',  # Chỉ có 2 số -> pad thành 056
            'EE_Location and Lanes': 'Data7',  # Chỉ có 1 số -> pad thành 07
            'EE_Array Number': 'EXP607056',  
            'expected': 'PASS'
        }
    ]
    
    print("=== TEST ARRAY NUMBER VALIDATION ===\n")
    
    for i, case in enumerate(test_cases, 1):
        print(f"Test case {i}:")
        print(f"  Cross Passage (A): {case['EE_Cross Passage']}")
        print(f"  Location Lanes (B): {case['EE_Location and Lanes']}")
        print(f"  Array Number (D): {case['EE_Array Number']}")
        
        # Thực hiện validation logic
        result = validate_array_number(case)
        
        print(f"  Expected: {case['expected']}")
        print(f"  Result: {result}")
        print(f"  Status: {'✅ PASS' if result == case['expected'] else '❌ FAIL'}")
        print()

def validate_array_number(row):
    """
    Test validation logic cho EE_Array Number
    """
    try:
        cross_passage = row.get('EE_Cross Passage', '')
        location_lanes = row.get('EE_Location and Lanes', '')
        array_number = row.get('EE_Array Number', '')
        
        if pd.notna(cross_passage) and pd.notna(location_lanes) and pd.notna(array_number):
            # Lấy 2 số cuối của cột B (EE_Location and Lanes)
            location_str = str(location_lanes).strip()
            numbers_in_location = re.findall(r'\d+', location_str)
            if numbers_in_location:
                last_2_b = numbers_in_location[-1][-2:] if len(numbers_in_location[-1]) >= 2 else numbers_in_location[-1].zfill(2)
            else:
                last_2_b = "00"
            
            # Lấy 3 số cuối của cột A (EE_Cross Passage)
            cross_str = str(cross_passage).strip()
            numbers_in_cross = re.findall(r'\d+', cross_str)
            if numbers_in_cross:
                last_3_a = numbers_in_cross[-1][-3:] if len(numbers_in_cross[-1]) >= 3 else numbers_in_cross[-1].zfill(3)
            else:
                last_3_a = "000"
            
            # Tạo expected pattern
            required_pattern = f"EXP6{last_2_b}{last_3_a}"
            actual_array = str(array_number).strip()
            
            print(f"    Computed pattern: {required_pattern}")
            print(f"    Contains check: '{required_pattern}' in '{actual_array}' = {required_pattern in actual_array}")
            
            # Kiểm tra xem array number có chứa pattern bắt buộc không
            if required_pattern in actual_array:
                return 'PASS'
            else:
                return 'FAIL'
        else:
            return 'SKIP (Missing data)'
            
    except Exception as e:
        return f'ERROR: {str(e)}'

if __name__ == "__main__":
    test_array_number_validation()
