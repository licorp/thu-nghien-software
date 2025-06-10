#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DEBUG DÒNG 27 - STD ARRAY TEE CONFLICT
=====================================

Dòng 27: F=150-900 | G=150.0 | K=STD ARRAY TEE
Lỗi: "Groove_Thread cần FAB Pipe 'Groove_Thread', có 'STD ARRAY TEE'"

Vấn đề: Tại sao STD ARRAY TEE lại bị check như Groove_Thread?
"""

import pandas as pd
import sys
import os

# Import validator
sys.path.append('.')
from excel_validator_final import ExcelValidator

def debug_dong_27():
    """
    Debug dòng 27 để tìm hiểu tại sao STD ARRAY TEE lại conflict
    """
    print("🐛 DEBUG DÒNG 27 - STD ARRAY TEE CONFLICT")
    print("=" * 60)
    
    # Tạo validator
    validator = ExcelValidator()
    
    # Dữ liệu từ dòng 27
    row_data = {
        'item_description': '150-900',          # F=150-900
        'size': '150.0',                        # G=150.0  
        'fab_pipe': 'STD ARRAY TEE',           # K=STD ARRAY TEE
        'end_1': 'UNKNOWN',                    # L=? (cần tìm hiểu)
        'end_2': 'UNKNOWN'                     # M=? (cần tìm hiểu)
    }
    
    print("📊 DỮ LIỆU DÒNG 27:")
    print("-" * 30)
    for key, value in row_data.items():
        print(f"  {key}: {value}")
    print()
    
    # Test case 1: Với End-1, End-2 không xác định
    print("🔍 TEST 1: End-1, End-2 = UNKNOWN")
    row = pd.Series(row_data)
    result1 = validator._check_pipe_schedule_mapping(
        row, 'item_description', 'size', 'fab_pipe', 'end_1', 'end_2'
    )
    print(f"Kết quả: {result1}")
    print()
    
    # Test case 2: Với End-1=RG, End-2=RG (có thể gây conflict)
    print("🔍 TEST 2: End-1=RG, End-2=RG (có thể gây conflict)")
    row_data_rg = row_data.copy()
    row_data_rg['end_1'] = 'RG'
    row_data_rg['end_2'] = 'RG'
    
    row = pd.Series(row_data_rg)
    result2 = validator._check_pipe_schedule_mapping(
        row, 'item_description', 'size', 'fab_pipe', 'end_1', 'end_2'
    )
    print(f"Kết quả: {result2}")
    print()
    
    # Test case 3: Với End-1=BE, End-2=BE
    print("🔍 TEST 3: End-1=BE, End-2=BE")
    row_data_be = row_data.copy()
    row_data_be['end_1'] = 'BE'
    row_data_be['end_2'] = 'BE'
    
    row = pd.Series(row_data_be)
    result3 = validator._check_pipe_schedule_mapping(
        row, 'item_description', 'size', 'fab_pipe', 'end_1', 'end_2'
    )
    print(f"Kết quả: {result3}")
    print()
    
    # Test case 4: Với End-1=TH, End-2=TH
    print("🔍 TEST 4: End-1=TH, End-2=TH")
    row_data_th = row_data.copy()
    row_data_th['end_1'] = 'TH'
    row_data_th['end_2'] = 'TH'
    
    row = pd.Series(row_data_th)
    result4 = validator._check_pipe_schedule_mapping(
        row, 'item_description', 'size', 'fab_pipe', 'end_1', 'end_2'
    )
    print(f"Kết quả: {result4}")
    print()
    
    print("🎯 PHÂN TÍCH:")
    print("-" * 30)
    print("✅ LOGIC ĐÚNG: size 150 + '900' → STD ARRAY TEE (ưu tiên cao)")
    print("❌ VẤN ĐỀ: Có thể End-1/End-2 đang gây conflict với logic ưu tiên thấp")
    print()
    print("🔍 CẦN KIỂM TRA:")
    print("1. End-1, End-2 thực tế của dòng 27 là gì?")
    print("2. Logic có đang check ưu tiên thấp trước ưu tiên cao?")

if __name__ == "__main__":
    debug_dong_27()
