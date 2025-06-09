#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

def analyze_pap_data():
    """
    Phân tích dữ liệu PAP 1 và PAP 2 để hiểu format thực tế
    """
    excel_file = '../Xp03-Fabrication & Listing.xlsx'
    
    try:
        # Đọc Pipe Schedule worksheet
        df = pd.read_excel(excel_file, sheet_name='Pipe Schedule')
        print(f"Đã đọc Pipe Schedule: {len(df)} dòng")
        
        # Kiểm tra các cột PAP
        pap_cols = ['EE_Pap 1', 'EE_Pap 2', 'Item Description', 'Size', 'Length']
        print(f"\nKiểm tra các cột:")
        for col in pap_cols:
            if col in df.columns:
                print(f"✅ {col}")
            else:
                print(f"❌ {col} - THIẾU")
        
        # Phân tích EE_Pap 1
        if 'EE_Pap 1' in df.columns:
            print(f"\n" + "="*60)
            print("PHÂN TÍCH EE_PAP 1")
            print("="*60)
            
            pap1_series = df['EE_Pap 1'].dropna()
            print(f"Tổng số giá trị không null: {len(pap1_series)}")
            
            # Lấy mẫu dữ liệu đầu tiên
            print(f"\n20 giá trị đầu tiên:")
            for i, val in enumerate(pap1_series.head(20)):
                item_desc = df.iloc[pap1_series.index[i]]['Item Description'] if 'Item Description' in df.columns else "N/A"
                print(f"{i+1:2d}. '{val}' (Item: {item_desc})")
            
            # Phân loại các pattern
            print(f"\nPhân loại pattern:")
            dimension_pattern = 0  # chứa 'x'
            size_letter = 0       # số + chữ cái
            number_only = 0       # chỉ số
            other = 0             # khác
            
            for val in pap1_series:
                val_str = str(val).strip()
                if 'x' in val_str.lower():
                    dimension_pattern += 1
                elif any(c.isalpha() for c in val_str) and any(c.isdigit() for c in val_str):
                    size_letter += 1
                elif val_str.replace('.', '').isdigit():
                    number_only += 1
                else:
                    other += 1
            
            print(f"  - Dimension pattern (NxN): {dimension_pattern}")
            print(f"  - Size + Letter (40B, 65LR): {size_letter}")
            print(f"  - Number only: {number_only}")
            print(f"  - Other formats: {other}")
            
        # Phân tích EE_Pap 2
        if 'EE_Pap 2' in df.columns:
            print(f"\n" + "="*60)
            print("PHÂN TÍCH EE_PAP 2")
            print("="*60)
            
            pap2_series = df['EE_Pap 2'].dropna()
            print(f"Tổng số giá trị không null: {len(pap2_series)}")
            
            # Lấy mẫu dữ liệu đầu tiên
            print(f"\n20 giá trị đầu tiên:")
            for i, val in enumerate(pap2_series.head(20)):
                if i < len(df):
                    size_val = df.iloc[pap2_series.index[i]]['Size'] if 'Size' in df.columns else "N/A"
                    length_val = df.iloc[pap2_series.index[i]]['Length'] if 'Length' in df.columns else "N/A"
                    print(f"{i+1:2d}. '{val}' (Size: {size_val}, Length: {length_val})")
            
    except Exception as e:
        print(f"Lỗi: {e}")

if __name__ == "__main__":
    analyze_pap_data()
