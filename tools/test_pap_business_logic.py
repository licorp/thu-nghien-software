#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

def test_pap_business_logic():
    """Test PAP validation với business logic mới"""
    excel_file = r'..\Xp03-Fabrication & Listing.xlsx'
    
    print("🧪 TEST PAP VALIDATION BUSINESS LOGIC")
    print("="*50)
    
    try:
        # Đọc dữ liệu
        df = pd.read_excel(excel_file, sheet_name='Pipe Schedule')
        print(f"✅ Đã đọc worksheet: {len(df)} dòng")
        
        # Test PAP 1 Logic
        print(f"\n📋 TEST PAP 1 LOGIC:")
        test_count = 0
        correct_count = 0
        
        for index, row in df.iterrows():
            size = row['Size']
            length = row['Length']
            pap1 = row['EE_PIPE END-1']
            
            if pd.notna(size) and pd.notna(length):
                size_val = float(size)
                length_val = float(length)
                pap1_str = str(pap1).strip() if pd.notna(pap1) else ""
                
                test_count += 1
                expected = ""
                
                # Business rules cho PAP 1:
                if abs(size_val - 150.0) < 0.1 and abs(length_val - 900.0) < 0.1:
                    expected = "65LR"
                elif abs(size_val - 65.0) < 0.1 and abs(length_val - 4730.0) < 5.0:
                    expected = "40B"
                elif abs(size_val - 65.0) < 0.1 and abs(length_val - 5295.0) < 5.0:
                    expected = "40B"
                else:
                    expected = ""  # Để trống
                
                # Kiểm tra
                if (expected == "" and (pap1_str == "" or pd.isna(pap1))) or (expected != "" and pap1_str == expected):
                    correct_count += 1
                elif test_count <= 10:  # Chỉ hiển thị 10 lỗi đầu
                    print(f"  ❌ Dòng {index + 2}: Size={size_val}, Length={length_val}")
                    print(f"     Expected: '{expected}', Actual: '{pap1_str}'")
        
        print(f"  📊 PAP 1: {correct_count}/{test_count} PASS ({correct_count/test_count*100:.1f}%)")
        
        # Test PAP 2 Logic
        print(f"\n📋 TEST PAP 2 LOGIC:")
        test_count = 0
        correct_count = 0
        
        for index, row in df.iterrows():
            size = row['Size']
            length = row['Length']
            pap2 = row['EE_PIPE END-2']
            
            if pd.notna(size) and pd.notna(length):
                size_val = float(size)
                length_val = float(length)
                pap2_str = str(pap2).strip() if pd.notna(pap2) else ""
                
                test_count += 1
                expected = ""
                
                # Business rules cho PAP 2:
                if abs(size_val - 65.0) < 0.1 and abs(length_val - 5295.0) < 5.0:
                    expected = "40B"
                else:
                    expected = ""  # Để trống
                
                # Kiểm tra
                if (expected == "" and (pap2_str == "" or pd.isna(pap2))) or (expected != "" and pap2_str == expected):
                    correct_count += 1
                elif test_count <= 10:  # Chỉ hiển thị 10 lỗi đầu
                    print(f"  ❌ Dòng {index + 2}: Size={size_val}, Length={length_val}")
                    print(f"     Expected: '{expected}', Actual: '{pap2_str}'")
        
        print(f"  📊 PAP 2: {correct_count}/{test_count} PASS ({correct_count/test_count*100:.1f}%)")
        
        # Phân tích dữ liệu thực tế
        print(f"\n📊 PHÂN TÍCH DỮ LIỆU THỰC TẾ:")
        unique_pap1 = df['EE_PIPE END-1'].dropna().unique()
        unique_pap2 = df['EE_PIPE END-2'].dropna().unique()
        print(f"  PAP 1 values: {sorted([str(v) for v in unique_pap1])}")
        print(f"  PAP 2 values: {sorted([str(v) for v in unique_pap2])}")
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_pap_business_logic()
