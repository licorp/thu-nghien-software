#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

def check_actual_pap_data():
    """Kiểm tra dữ liệu PAP thực tế trong file Excel"""
    excel_file = r'..\Xp03-Fabrication & Listing.xlsx'
    
    print("🔍 KIỂM TRA DỮ LIỆU PAP THỰC TẾ")
    print("="*50)
    
    try:
        # Đọc Pipe Schedule worksheet
        df = pd.read_excel(excel_file, sheet_name='Pipe Schedule')
        print(f"✅ Đã đọc worksheet: {len(df)} dòng")
        
        # Kiểm tra các cột PAP
        pap_cols = ['EE_PIPE END-1', 'EE_PIPE END-2']
        for col in pap_cols:
            if col in df.columns:
                unique_values = df[col].dropna().unique()
                print(f"\n📊 CỘT {col}:")
                print(f"  Số giá trị unique: {len(unique_values)}")
                print(f"  Các giá trị: {sorted([str(v) for v in unique_values])}")
                
                # Phân loại theo pattern
                size_codes = []
                numbers = []
                others = []
                
                for val in unique_values:
                    val_str = str(val).strip()
                    if val_str.replace('.', '').replace('-', '').isdigit():
                        numbers.append(val_str)
                    elif any(c.isalpha() for c in val_str) and any(c.isdigit() for c in val_str):
                        size_codes.append(val_str)
                    else:
                        others.append(val_str)
                
                print(f"  📏 Size codes (40B, 20T): {size_codes}")
                print(f"  🔢 Numbers: {numbers[:10]}{'...' if len(numbers) > 10 else ''}")
                print(f"  ❓ Others: {others}")
        
        # Kiểm tra dòng 57 cụ thể
        if len(df) > 56:
            row = df.iloc[56]  # Dòng 57
            print(f"\n🔍 DÒNG 57 (Index 56):")
            print(f"  EE_PIPE END-1: '{row['EE_PIPE END-1']}'")
            print(f"  EE_PIPE END-2: '{row['EE_PIPE END-2']}'")
            print(f"  Size: {row['Size']}")
            print(f"  Length: {row['Length']}")
        
        # Kiểm tra một vài dòng có PAP 2 FAIL
        print(f"\n🔍 MẪU CÁC DÒNG CÓ PAP 2 DATA:")
        count = 0
        for index, row in df.iterrows():
            pap2 = row['EE_PIPE END-2']
            if pd.notna(pap2) and str(pap2).strip() != '':
                count += 1
                if count <= 10:  # Chỉ hiển thị 10 dòng đầu
                    print(f"  Dòng {index + 2}: PAP2='{pap2}', Size={row['Size']}, Length={row['Length']}")
        
        print(f"\n📊 Tổng số dòng có PAP 2 data: {count}")
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_actual_pap_data()
