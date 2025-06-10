#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

def check_all_pap2():
    """
    Kiểm tra tất cả giá trị PAP 2 để hiểu tại sao vẫn có lỗi
    """
    excel_file = '../Xp03-Fabrication & Listing.xlsx'
    
    try:
        # Đọc Pipe Schedule worksheet
        df = pd.read_excel(excel_file, sheet_name='Pipe Schedule')
        print(f"Đã đọc Pipe Schedule: {len(df)} dòng")
        
        # Phân tích EE_Pap 2
        if 'EE_Pap 2' in df.columns:
            pap2_series = df['EE_Pap 2'].dropna()
            print(f"Tổng số giá trị PAP 2 không null: {len(pap2_series)}")
            
            print(f"\nTất cả {len(pap2_series)} giá trị PAP 2:")
            for i, val in enumerate(pap2_series):
                size_val = df.iloc[pap2_series.index[i]]['Size'] if 'Size' in df.columns else "N/A"
                length_val = df.iloc[pap2_series.index[i]]['Length'] if 'Length' in df.columns else "N/A"
                print(f"{i+1:2d}. '{val}' (Size: {size_val}, Length: {length_val})")
            
            print(f"\nUnique PAP 2 values: {pap2_series.unique()}")
            print(f"Unique PAP 2 value types: {[type(x).__name__ for x in pap2_series.unique()]}")
            
            # Kiểm tra pattern
            import re
            size_code_pattern = r'\d+[A-Z]+\d*'
            dimension_pattern = r'\d+x\d+(?:x\d+)?'
            
            pattern_stats = {
                'size_codes': 0,
                'dimensions': 0, 
                'numbers': 0,
                'other': 0
            }
            
            for val in pap2_series:
                val_str = str(val).strip()
                if re.search(size_code_pattern, val_str):
                    pattern_stats['size_codes'] += 1
                elif re.search(dimension_pattern, val_str):
                    pattern_stats['dimensions'] += 1
                elif val_str.replace('.', '').replace('-', '').isdigit():
                    pattern_stats['numbers'] += 1
                else:
                    pattern_stats['other'] += 1
                    print(f"Unknown pattern: '{val_str}'")
            
            print(f"\nPattern Statistics:")
            for pattern, count in pattern_stats.items():
                print(f"  {pattern}: {count}")
                
    except Exception as e:
        print(f"Lỗi: {e}")

if __name__ == "__main__":
    check_all_pap2()
