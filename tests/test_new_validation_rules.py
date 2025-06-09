#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import sys
import os
sys.path.append('.')

# Import tool mới
from excel_validator_enhanced import ExcelValidatorEnhanced

def test_new_validation_rules():
    """
    Test các validation rules mới: FAB Pipe, Pap 1, Pap 2
    """
    print("=" * 80)
    print("🧪 TEST CÁC VALIDATION RULES MỚI")
    print("=" * 80)
    
    # Tạo test data
    test_data = {
        'EE_Cross Passage': ['EXP61001', 'EXP61002', 'EXP61003', 'EXP61004', 'EXP61005'],
        'EE_Location and Lanes': ['M110-M111', 'M110-M111', 'M110-M111', 'M110-M111', 'M110-M111'],
        'EE_System Type': ['CP-INTERNAL', 'CP-EXTERNAL', 'CW-ARRAY', 'CP-INTERNAL', 'CW-ARRAY'],
        'EE_Array Number': ['EXP61001', 'EXP61111', 'EXP61111', 'EXP61002', 'EXP61111'],
        'Item': ['A', 'B', 'C', 'D', 'E'],
        'Item Description': [
            'Steel Pipe, SCH 40, Groove Thread End',
            'Steel 90° Elbow, SCH 40, Thread End', 
            'Steel Tee, SCH 40, Flange End',
            'Steel Coupling, SCH 40',
            'Steel Pipe, SCH 40'
        ],
        'Type': ['Pipe', 'Fitting', 'Fitting', 'Fitting', 'Pipe'],
        'Size': [150.0, 100.0, 65.0, 80.0, 65.0],
        'Qty': [1, 2, 1, 1, 1],
        'Length': [5000.0, 100.0, 200.0, 150.0, 5295.0],
        'EE_FAB Pipe': ['Groove_Thread', 'Thread', 'Wrong_Value', 'Coupling', 'Thread'],
        'EE_PIPE END-1': ['Groove', 'Thread', 'Flange', 'Thread', 'Thread'],
        'EE_PIPE END-2': ['Thread', 'Thread', 'Flange', 'Thread', 'Thread'],
        'EE_Run Dim 1': [150, 100, 65, 80, 65],
        'EE_Pap 1': ['Straight_Pipe', '90_Elbow', 'Wrong_Value', 'Coupling', ''],
        'EE_Run Dim 2': [0, 0, 0, 0, 0],
        'EE_Pap 2': ['', '', '', '', 'Wrong_Special'],
        'EE_Run Dim 3': [0, 0, 0, 0, 0],
        'EE_Pap 3': ['', '', '', '', ''],
        'EE_Pipe Treatment': ['GAL', 'BLACK', 'BLACK', 'GAL', 'BLACK'],
        'Family': ['Pipe', 'Fitting', 'Fitting', 'Fitting', 'Pipe'],
        'Type2': ['Steel', 'Steel', 'Steel', 'Steel', 'Steel'],
        'ID': ['P001', 'F001', 'F002', 'F003', 'P002']
    }
    
    # Tạo DataFrame
    df = pd.DataFrame(test_data)
    
    print(f"📋 TEST DATA ({len(df)} dòng):")
    print(df[['Item Description', 'Size', 'Length', 'EE_FAB Pipe', 'EE_Pap 1', 'EE_Pap 2']].to_string())
    
    # Tạo validator instance
    validator = ExcelValidatorEnhanced()
    
    # Test FAB Pipe validation
    print(f"\n🏭 TEST FAB PIPE VALIDATION:")
    print("-" * 50)
    
    for idx, row in df.iterrows():
        result = validator._check_fab_pipe_detailed(
            row, 'Item Description', 'Size', 'EE_PIPE END-1', 'EE_PIPE END-2', 'EE_FAB Pipe'
        )
        print(f"Dòng {idx+1}: {row['Item Description'][:30]:30s} → {result}")
    
    # Test Pap 1 validation
    print(f"\n📏 TEST PAP 1 VALIDATION:")
    print("-" * 50)
    
    for idx, row in df.iterrows():
        result = validator._check_pap1_detailed(row, 'Item Description', 'EE_Pap 1')
        print(f"Dòng {idx+1}: {row['Item Description'][:30]:30s} → {result}")
    
    # Test Pap 2 validation
    print(f"\n📏 TEST PAP 2 VALIDATION:")
    print("-" * 50)
    
    for idx, row in df.iterrows():
        result = validator._check_pap2_detailed(row, 'Size', 'Length', 'EE_Pap 2')
        print(f"Dòng {idx+1}: Size={row['Size']}, Length={row['Length']} → {result}")
    
    print(f"\n🎉 TEST HOÀN THÀNH!")

if __name__ == "__main__":
    test_new_validation_rules()
