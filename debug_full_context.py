#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DEBUG FULL VALIDATION CONTEXT
==============================

Test với full validation context để tìm ra logic conflict
"""

import pandas as pd
import sys
import os

# Import validator
sys.path.append('.')
from excel_validator_final import ExcelValidator

def debug_full_validation():
    """
    Debug với full validation context như trong Excel thực tế
    """
    print("🐛 DEBUG: FULL VALIDATION CONTEXT")
    print("=" * 60)
    
    # Tạo validator
    validator = ExcelValidator()
    
    # Dữ liệu từ dòng 21 Excel (theo description của user)
    row_data = {
        'EE_Cross Passage': 'CP001',              # Col A
        'EE_Location and Lanes': 'Lane1002',      # Col B  
        'EE_System Type': 'CW-ARRAY',             # Col C
        'EE_Array Number': 'EXP61053A',           # Col D
        'EE_Empty1': '',                          # Col E
        'EE_Item Description': '65-5295',         # Col F
        'EE_Size': 65.0,                          # Col G
        'EE_Empty2': '',                          # Col H
        'EE_Empty3': '',                          # Col I
        'EE_Empty4': '',                          # Col J
        'EE_FAB Pipe': 'STD 2 PAP RANGE',         # Col K
        'EE_End-1': 'RG',                         # Col L
        'EE_End-2': 'BE',                         # Col M
        'EE_Empty5': '',                          # Col N
        'EE_Empty6': '',                          # Col O
        'EE_Empty7': '',                          # Col P
        'EE_Empty8': '',                          # Col Q
        'EE_Empty9': '',                          # Col R
        'EE_Empty10': '',                         # Col S
        'EE_Pipe Treatment': 'BLACK'              # Col T
    }
    
    print("📊 DỮ LIỆU DÒNG 21 EXCEL:")
    print("-" * 30)
    for key, value in row_data.items():
        if value != '':  # Chỉ hiển thị cột có data
            print(f"  {key}: {value}")
    print()
    
    # Tạo mock row
    row = pd.Series(row_data)
    
    # Tạo column mapping như trong code thực tế
    columns = list(row_data.keys())
    col_a_name = columns[0]   # EE_Cross Passage
    col_b_name = columns[1]   # EE_Location and Lanes
    col_c_name = columns[2]   # EE_System Type
    col_d_name = columns[3]   # EE_Array Number
    col_f_name = columns[5]   # EE_Item Description
    col_g_name = columns[6]   # EE_Size
    col_k_name = columns[10]  # EE_FAB Pipe
    col_l_name = columns[11]  # EE_End-1
    col_m_name = columns[12]  # EE_End-2
    col_t_name = columns[19]  # EE_Pipe Treatment
    
    print("🚀 CHẠY FULL VALIDATION (tất cả 4 rules)...")
    print()
    
    # Chạy _validate_row như trong code thực tế
    result = validator._validate_row(
        row, 
        col_a_name, col_b_name, col_c_name, col_d_name, 
        col_f_name, col_g_name, col_k_name, col_l_name, col_m_name, col_t_name,
        apply_array_validation=True,           # Pipe Schedule worksheet
        apply_pipe_treatment_validation=True,  # Pipe Schedule worksheet  
        apply_cp_internal_validation=True,     # Pipe Schedule worksheet
        apply_pipe_schedule_mapping_validation=True  # Pipe Schedule worksheet
    )
    
    print(f"📝 KẾT QUẢ FULL VALIDATION: {result}")
    print()
    
    # Phân tích từng rule riêng lẻ
    print("🔍 PHÂN TÍCH TỪNG RULE:")
    print("-" * 40)
    
    # Rule 1: Array Number (không áp dụng vì CW-ARRAY không phải CP-INTERNAL)
    print("1. Array Number Rule:")
    array_result = validator._check_array_number(row, col_a_name, col_b_name, col_d_name)
    print(f"   Kết quả: {array_result}")
    
    # Rule 2: Pipe Treatment
    print("2. Pipe Treatment Rule:")
    treatment_result = validator._check_pipe_treatment(row, col_c_name, col_t_name)
    print(f"   Kết quả: {treatment_result}")
    
    # Rule 3: CP-INTERNAL (không áp dụng vì CW-ARRAY)
    print("3. CP-INTERNAL Rule:")
    cp_result = validator._check_cp_internal_array(row, col_a_name, col_c_name, col_d_name)
    print(f"   Kết quả: {cp_result}")
    
    # Rule 4: Pipe Schedule Mapping
    print("4. Pipe Schedule Mapping Rule:")
    mapping_result = validator._check_pipe_schedule_mapping(row, col_f_name, col_g_name, col_k_name, col_l_name, col_m_name)
    print(f"   Kết quả: {mapping_result}")
    
    print()
    print("🎯 PHÂN TÍCH:")
    if result == "PASS":
        print("✅ PASS: Logic hoạt động đúng!")
    else:
        print("❌ FAIL: Có conflict logic!")
        print(f"   Chi tiết: {result}")

if __name__ == "__main__":
    debug_full_validation()
