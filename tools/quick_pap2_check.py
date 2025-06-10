#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re
import sys
import os

# Add production directory to path
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'production'))

# Import validation functions
from excel_validator_detailed import validate_pap_2

def quick_pap2_check():
    """Kiểm tra nhanh PAP 2 validation"""
    excel_file = r'd:\OneDrive\Desktop\thu nghien software\production\Xp03-Fabrication & Listing.xlsx'
    
    print("🔍 KIỂM TRA NHANH PAP 2 VALIDATION")
    print("="*50)
    
    try:
        # Đọc Pipe Schedule worksheet
        df = pd.read_excel(excel_file, sheet_name='Pipe Schedule')
        print(f"✅ Đã đọc worksheet Pipe Schedule: {len(df)} dòng")
        
        # Kiểm tra các cột cần thiết
        required_cols = ['EE_PIPE END-1', 'EE_PIPE END-2', 'Size', 'Length']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            print(f"❌ Thiếu cột: {missing_cols}")
            return
            
        print("✅ Tất cả cột cần thiết đều có")
        
        # Kiểm tra dòng 57 (index 56)
        if len(df) > 56:
            row = df.iloc[56]  # Dòng 57
            print(f"\n🔍 KIỂM TRA DÒNG 57:")
            print(f"  EE_PIPE END-1: {row['EE_PIPE END-1']}")
            print(f"  EE_PIPE END-2: {row['EE_PIPE END-2']}")
            print(f"  Size: {row['Size']}")
            print(f"  Length: {row['Length']}")
            
            # Test validation function
            result = validate_pap_2(row, 57)
            print(f"  Kết quả validation: {result}")
        
        # Kiểm tra tất cả các dòng PAP 2 có lỗi
        print(f"\n🔍 KIỂM TRA TẤT CẢ DÒNG PAP 2:")
        fail_count = 0
        
        for index, row in df.iterrows():
            pap1 = str(row['EE_PIPE END-1']) if pd.notna(row['EE_PIPE END-1']) else ""
            pap2 = str(row['EE_PIPE END-2']) if pd.notna(row['EE_PIPE END-2']) else ""
            
            # Chỉ kiểm tra những dòng có PAP 2 data
            if pap2 and pap2.upper() != 'NAN':
                result = validate_pap_2(row, index + 2)
                if 'FAIL' in result:
                    fail_count += 1
                    if fail_count <= 5:  # Chỉ hiển thị 5 lỗi đầu
                        print(f"  Dòng {index + 2}: {result}")
                        print(f"    PAP2: {pap2}")
                        print(f"    Size: {row['Size']}")
                        print(f"    Length: {row['Length']}")
                        print()
        
        print(f"📊 Tổng số dòng FAIL: {fail_count}")
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    quick_pap2_check()
