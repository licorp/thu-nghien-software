#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re

def debug_pap2_validation():
    """
    Debug PAP 2 validation để hiểu tại sao vẫn FAIL
    """
    excel_file = '../Xp03-Fabrication & Listing.xlsx'
    
    # Đọc dữ liệu
    df = pd.read_excel(excel_file, sheet_name='Pipe Schedule')
    print(f"Loaded {len(df)} rows")
    
    # Tìm cột
    col_p_name = None
    size_col = None
    length_col = None
    
    for col in df.columns:
        if 'pap 2' in col.lower():
            col_p_name = col
        elif col.lower() == 'size':
            size_col = col
        elif col.lower() == 'length':
            length_col = col
    
    print(f"PAP 2 column: {col_p_name}")
    print(f"Size column: {size_col}")
    print(f"Length column: {length_col}")
    
    if not col_p_name:
        print("ERROR: Không tìm thấy cột PAP 2")
        return
    
    # Test validation logic trên từng dòng có PAP 2 data
    pap2_data = df[df[col_p_name].notna()]
    print(f"\nFound {len(pap2_data)} rows with PAP 2 data")
    
    for idx, row in pap2_data.head(10).iterrows():
        print(f"\nRow {idx+1}:")
        size = row[size_col] if size_col else None
        length = row[length_col] if length_col else None
        pap2 = row[col_p_name]
        
        print(f"  Size: {size} (type: {type(size)})")
        print(f"  Length: {length} (type: {type(length)})")
        print(f"  PAP 2: {pap2} (type: {type(pap2)})")
        
        # Simulate validation logic
        pap2_str = str(pap2).strip()
        print(f"  PAP 2 str: '{pap2_str}'")
        
        # Check 65mm + 5295mm condition
        if size and pd.notna(size) and length and pd.notna(length):
            size_val = float(size) if not pd.isna(size) else 0
            length_val = float(length) if not pd.isna(length) else 0
            print(f"  Size val: {size_val}, Length val: {length_val}")
            
            if abs(size_val - 65.0) < 0.1 and abs(length_val - 5295.0) < 5.0:
                print(f"  ✅ Matches 65mm + 5295mm condition")
                
                # Check patterns
                dimension_pattern = r'\d+x\d+(?:x\d+)?'
                size_code_pattern = r'\d+[A-Z]+\d*'
                
                if re.search(dimension_pattern, pap2_str):
                    result = "PASS (Rule: 65mm-5295mm Dimension)"
                elif re.search(size_code_pattern, pap2_str):
                    result = "PASS (Rule: 65mm-5295mm Size Code)"
                else:
                    result = f"FAIL (Rule: 65mm-5295mm): ống 65mm dài 5295mm cần dimension format (NxN) hoặc size code (40B, 65LR), có '{pap2_str}'"
                
                print(f"  Result: {result}")
            else:
                print(f"  ❌ Does not match 65mm + 5295mm condition")
        
        print(f"  " + "-"*50)

if __name__ == "__main__":
    debug_pap2_validation()
