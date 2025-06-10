#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KIỂM TRA DỮ LIỆU THỰC TẾ DÒNG 27
=================================
"""

import pandas as pd

def check_dong_27_real_data():
    """
    Kiểm tra dữ liệu thực tế của dòng 27
    """
    print("🔍 KIỂM TRA DỮ LIỆU THỰC TẾ DÒNG 27")
    print("=" * 50)
    
    # File Excel gốc
    file = 'MEP_Schedule_Table_20250610_154246.xlsx'
    
    try:
        # Đọc Pipe Schedule worksheet  
        df = pd.read_excel(file, sheet_name='Pipe Schedule')
        
        # Lấy dòng 27 (index 26 vì 0-based)
        if len(df) >= 27:
            row = df.iloc[26]  # Dòng 27
            
            print("📊 DỮ LIỆU DÒNG 27 (THỰC TẾ):")
            print("-" * 40)
            
            # Các cột quan trọng
            item_desc = row.get('EE_Item Description', 'N/A')
            size = row.get('EE_Size', 'N/A')
            fab_pipe = row.get('EE_FAB Pipe', 'N/A')
            end_1 = row.get('EE_End-1', 'N/A')
            end_2 = row.get('EE_End-2', 'N/A')
            system_type = row.get('EE_System Type', 'N/A')
            
            print(f"  Item Description: {item_desc}")
            print(f"  Size: {size}")
            print(f"  FAB Pipe: {fab_pipe}")
            print(f"  End-1: {end_1}")
            print(f"  End-2: {end_2}")
            print(f"  System Type: {system_type}")
            print()
            
            # Phân tích case
            print("🎯 PHÂN TÍCH CASE:")
            print("-" * 20)
            if str(item_desc) == '150-900' and str(size) == '150.0':
                print("✅ Đây là case STD ARRAY TEE (ưu tiên cao)")
                print(f"   - Item: {item_desc} (chứa '900')")
                print(f"   - Size: {size} (= 150)")
                
                if str(end_1) == 'RG' and str(end_2) == 'RG':
                    print("🔴 VẤN ĐỀ: End-1=RG, End-2=RG")
                    print("   → Logic đang nhảy vào Groove_Thread (ưu tiên thấp)")
                    print("   → Thay vì STD ARRAY TEE (ưu tiên cao)")
                    print()
                    print("🛠️ GIẢI PHÁP: Cần sửa logic để ưu tiên cao chạy TRƯỚC ưu tiên thấp")
            
        else:
            print("❌ Không đủ dữ liệu (< 27 dòng)")
                
    except Exception as e:
        print(f"❌ Lỗi đọc file: {e}")

if __name__ == "__main__":
    check_dong_27_real_data()
