#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KIỂM TRA LỖI THỰC TẾ DÒNG 27
============================
"""

import pandas as pd
import glob

def check_dong_27_errors():
    """
    Kiểm tra lỗi thực tế của dòng 27 trong file validation mới nhất
    """
    print("🔍 KIỂM TRA LỖI THỰC TẾ DÒNG 27")
    print("=" * 50)
    
    # Tìm file validation mới nhất
    validation_files = glob.glob("validation_4rules_*.xlsx")
    if not validation_files:
        print("❌ Không tìm thấy file validation")
        return
    
    latest_file = sorted(validation_files)[-1]
    print(f"📁 File validation mới nhất: {latest_file}")
    print()
    
    try:
        # Đọc Pipe Schedule worksheet từ file validation
        df = pd.read_excel(latest_file, sheet_name='Pipe Schedule')
        
        # Lấy dòng 27 (index 26)
        if len(df) >= 27:
            row = df.iloc[26]  # Dòng 27
            
            print("📊 DỮ LIỆU DÒNG 27 TRONG FILE VALIDATION:")
            print("-" * 50)
            
            # Các cột quan trọng
            item_desc = row.get('EE_Item Description', 'N/A')
            size = row.get('EE_Size', 'N/A')
            fab_pipe = row.get('EE_FAB Pipe', 'N/A')
            end_1 = row.get('EE_End-1', 'N/A')
            end_2 = row.get('EE_End-2', 'N/A')
            validation_check = row.get('Validation_Check', 'N/A')
            
            print(f"  Item Description: {item_desc}")
            print(f"  Size: {size}")
            print(f"  FAB Pipe: {fab_pipe}")
            print(f"  End-1: {end_1}")
            print(f"  End-2: {end_2}")
            print(f"  Validation Check: {validation_check}")
            print()
            
            # Phân tích kết quả
            if validation_check != 'PASS':
                print("🔴 DÒNG 27 CÓ LỖI:")
                print(f"   {validation_check}")
                print()
                
                # Phân tích nguyên nhân
                if 'Groove_Thread' in str(validation_check):
                    print("🎯 PHÂN TÍCH NGUYÊN NHÂN:")
                    print("   1. Logic có thể đang check End-1/End-2 trước size+item")
                    print("   2. Cần sửa thứ tự ưu tiên trong code")
            else:
                print("✅ DÒNG 27 PASS - Không có lỗi")
                
        # Tìm tất cả dòng có lỗi Groove_Thread với 150-900
        print("\n🔍 TÌM TẤT CẢ DÒNG LỖI GROOVE_THREAD VỚI 150-900:")
        print("-" * 50)
        
        mask_150_900 = df['EE_Item Description'].astype(str).str.contains('150-900', na=False)
        mask_groove_error = df['Validation_Check'].astype(str).str.contains('Groove_Thread', na=False)
        
        error_rows = df[mask_150_900 & mask_groove_error]
        
        if len(error_rows) > 0:
            print(f"📊 TÌM THẤY {len(error_rows)} DÒNG LỖI:")
            for idx, row in error_rows.head(5).iterrows():
                item_desc = row.get('EE_Item Description', 'N/A')
                size = row.get('EE_Size', 'N/A')
                fab_pipe = row.get('EE_FAB Pipe', 'N/A')
                end_1 = row.get('EE_End-1', 'N/A')
                end_2 = row.get('EE_End-2', 'N/A')
                validation_check = row.get('Validation_Check', 'N/A')
                
                print(f"  Dòng {idx+2}: {item_desc} | Size={size} | End1={end_1} | End2={end_2}")
                print(f"    Lỗi: {validation_check}")
                print()
        else:
            print("✅ Không tìm thấy dòng nào có lỗi Groove_Thread với 150-900")
                
    except Exception as e:
        print(f"❌ Lỗi đọc file: {e}")

if __name__ == "__main__":
    check_dong_27_errors()
