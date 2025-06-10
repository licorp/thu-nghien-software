#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re

def debug_validation_with_tracing():
    """
    Debug PAP validation với tracing chi tiết để hiểu vấn đề
    """
    excel_file = '../Xp03-Fabrication & Listing.xlsx'
    
    # Đọc dữ liệu
    df = pd.read_excel(excel_file, sheet_name='Pipe Schedule')
    print(f"Loaded {len(df)} rows")
    print(f"Columns: {list(df.columns)}")
    
    # Tìm các cột cần thiết
    col_o_name = None  # PAP 1
    col_p_name = None  # PAP 2
    size_col = None
    length_col = None
    item_desc_col = None
    
    for col in df.columns:
        if 'pap 1' in col.lower() or ('ee' in col.lower() and 'pap' in col.lower() and '1' in col):
            col_o_name = col
        elif 'pap 2' in col.lower() or ('ee' in col.lower() and 'pap' in col.lower() and '2' in col):
            col_p_name = col
        elif col.lower() == 'size':
            size_col = col
        elif col.lower() == 'length':
            length_col = col
        elif 'item' in col.lower() and 'description' in col.lower():
            item_desc_col = col
    
    print(f"\nColumn mapping:")
    print(f"  PAP 1 (O): {col_o_name}")
    print(f"  PAP 2 (P): {col_p_name}")
    print(f"  Size: {size_col}")
    print(f"  Length: {length_col}")
    print(f"  Item Desc: {item_desc_col}")
    
    # Kiểm tra data type của các cột PAP
    if col_p_name:
        print(f"\nPAP 2 column details:")
        print(f"  Data type: {df[col_p_name].dtype}")
        print(f"  Non-null count: {df[col_p_name].notna().sum()}")
        print(f"  Unique values: {df[col_p_name].dropna().unique()}")
        print(f"  Value types: {[type(x).__name__ for x in df[col_p_name].dropna().unique()]}")
        
        # Check if any values could be misinterpreted as length
        pap2_series = df[col_p_name].dropna()
        print(f"\nChecking for potential data confusion:")
        for i, val in enumerate(pap2_series.head(10)):
            row_idx = pap2_series.index[i]
            length_val = df.iloc[row_idx][length_col] if length_col else None
            print(f"  Row {row_idx+1}: PAP2='{val}' (type: {type(val).__name__}), Length={length_val}")
            
            # Check if PAP2 value could be confused with length
            val_str = str(val).strip()
            try:
                val_numeric = float(val_str)
                if abs(val_numeric - 5250) < 10 or abs(val_numeric - 5295) < 10:
                    print(f"    ⚠️ PAP2 value '{val}' is close to length values!")
            except:
                pass
    
    print(f"\n" + "="*60)
    print("SIMULATING PAP 2 VALIDATION")
    print("="*60)
    
    # Simulate the exact validation logic from the main code
    if col_p_name and size_col and length_col:
        for idx, row in df.iterrows():
            # Skip rows without PAP 2 data
            if pd.isna(row[col_p_name]):
                continue
                
            # Only show first 5 rows with data
            if df[:idx+1][col_p_name].notna().sum() > 5:
                break
                
            print(f"\nRow {idx+1} validation:")
            
            size = row[size_col] if size_col else None
            length = row[length_col] if length_col else None
            pap2 = row[col_p_name]
            
            print(f"  Raw data: size={size}, length={length}, pap2={pap2}")
            print(f"  Types: size={type(size)}, length={type(length)}, pap2={type(pap2)}")
            
            # Simulate validation
            if pd.isna(pap2):
                result = "SKIP: Pap 2 trống"
            else:
                pap2_str = str(pap2).strip()
                print(f"  pap2_str = '{pap2_str}'")
                
                # Check 65mm + ~5295mm condition
                if size and pd.notna(size) and length and pd.notna(length):
                    size_val = float(size) if not pd.isna(size) else 0
                    length_val = float(length) if not pd.isna(length) else 0
                    print(f"  Converted: size_val={size_val}, length_val={length_val}")
                    
                    if abs(size_val - 65.0) < 0.1 and abs(length_val - 5295.0) < 5.0:
                        print(f"  ✅ Matches 65mm + ~5295mm condition")
                        
                        # Check patterns
                        dimension_pattern = r'\d+x\d+(?:x\d+)?'
                        size_code_pattern = r'\d+[A-Z]+\d*'
                        
                        if re.search(dimension_pattern, pap2_str):
                            result = "PASS (Rule: 65mm-5295mm Dimension)"
                        elif re.search(size_code_pattern, pap2_str):
                            result = "PASS (Rule: 65mm-5295mm Size Code)"
                        else:
                            result = f"FAIL (Rule: 65mm-5295mm): ống 65mm dài 5295mm cần dimension format (NxN) hoặc size code (40B, 65LR), có '{pap2_str}'"
                    else:
                        print(f"  ❌ Does not match 65mm + ~5295mm condition")
                        
                        # Fall through to general pattern check
                        dimension_pattern = r'\d+x\d+(?:x\d+)?'
                        size_code_pattern = r'\d+[A-Z]+\d*'
                        
                        if re.search(dimension_pattern, pap2_str):
                            result = "PASS (Rule: Valid Dimension Format)"
                        elif re.search(size_code_pattern, pap2_str):
                            result = "PASS (Rule: Valid Size Code)"
                        else:
                            result = f"FAIL (Rule: Valid Format): cần format NxN, NxNxN hoặc Size Code (ví dụ: 40B, 65LR), có '{pap2_str}'"
                else:
                    print(f"  ❌ Missing size or length data")
                    result = "SKIP: Thiếu Size hoặc Length"
            
            print(f"  RESULT: {result}")

if __name__ == "__main__":
    debug_validation_with_tracing()
