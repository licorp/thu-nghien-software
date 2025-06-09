#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

def check_pap2_data():
    """Kiểm tra dữ liệu PAP 2 thực tế"""
    excel_file = r'd:\OneDrive\Desktop\thu nghien software\Xp03-Fabrication & Listing.xlsx'
    
    print("🔍 KIỂM TRA DỮ LIỆU PAP 2 THỰC TẾ")
    print("="*60)
    
    # Đọc Pipe Schedule worksheet
    df = pd.read_excel(excel_file, sheet_name='Pipe Schedule')
    print(f"✅ Đã đọc worksheet Pipe Schedule: {len(df)} dòng")
    
    # Kiểm tra dòng 57 (index 56)
    if len(df) > 56:
        row = df.iloc[56]  # Dòng 57
        print(f"\n🔍 DÒNG 57 CHI TIẾT:")
        print(f"  EE_PIPE END-1: '{row['EE_PIPE END-1']}'")
        print(f"  EE_PIPE END-2: '{row['EE_PIPE END-2']}'")
        print(f"  Size: '{row['Size']}'")
        print(f"  Length: '{row['Length']}'")
        print(f"  Type của Size: {type(row['Size'])}")
        print(f"  Type của Length: {type(row['Length'])}")
        
        # Kiểm tra điều kiện 65mm và 5295mm
        try:
            size_val = float(row['Size']) if pd.notna(row['Size']) else 0
            length_val = float(row['Length']) if pd.notna(row['Length']) else 0
            print(f"  Size value: {size_val}")
            print(f"  Length value: {length_val}")
            print(f"  Size condition (65mm): {abs(size_val - 65.0) < 0.1}")
            print(f"  Length condition (5295mm): {abs(length_val - 5295.0) < 5.0}")
        except Exception as e:
            print(f"  Lỗi convert: {e}")
    
    # Kiểm tra tất cả các giá trị PAP 2 unique
    print(f"\n📊 TẤT CẢ GIÁ TRỊ PAP 2 UNIQUE:")
    pap2_values = df['EE_PIPE END-2'].dropna().unique()
    print(f"Số giá trị unique: {len(pap2_values)}")
    for i, val in enumerate(sorted(pap2_values)):
        print(f"  {i+1:2d}. '{val}' (type: {type(val)})")
        if i >= 10:  # Giới hạn hiển thị
            print(f"  ... và {len(pap2_values) - 11} giá trị khác")
            break
    
    # Kiểm tra những dòng có lỗi PAP 2
    print(f"\n❌ CÁC DÒNG CÓ VẤN ĐỀ:")
    problem_count = 0
    for index, row in df.iterrows():
        pap2 = row['EE_PIPE END-2']
        if pd.notna(pap2):
            pap2_str = str(pap2).strip()
            
            # Kiểm tra pattern
            import re
            dimension_pattern = r'\d+x\d+(?:x\d+)?'
            size_code_pattern = r'\d+[A-Z]+\d*'
            
            if not (re.search(dimension_pattern, pap2_str) or re.search(size_code_pattern, pap2_str)):
                problem_count += 1
                if problem_count <= 10:  # Hiển thị 10 lỗi đầu
                    print(f"  Dòng {index + 2}: PAP2='{pap2_str}', Size='{row['Size']}', Length='{row['Length']}'")
    
    print(f"\n📊 Tổng số dòng có vấn đề: {problem_count}")

if __name__ == "__main__":
    check_pap2_data()
