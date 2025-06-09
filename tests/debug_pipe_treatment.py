#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

def debug_pipe_treatment_validation():
    """
    Debug tại sao Pipe Treatment validation không áp dụng cho tất cả worksheet
    """
    excel_file = 'Xp02-Fabrication & Listing.xlsx'
    
    # Các worksheet cần check Pipe Treatment
    pipe_treatment_worksheets = [
        'Pipe Schedule', 
        'Pipe Fitting Schedule', 
        'Pipe Accessory Schedule'
    ]
    
    for sheet_name in pipe_treatment_worksheets:
        print(f"{'='*60}")
        print(f"DEBUG WORKSHEET: {sheet_name}")
        print(f"{'='*60}")
        
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        print(f"Số dòng: {len(df)}, Số cột: {len(df.columns)}")
        
        # Kiểm tra cột System Type và Pipe Treatment
        col_c_name = df.columns[2] if len(df.columns) > 2 else None  # EE_System Type
        col_t_name = df.columns[19] if len(df.columns) > 19 else None  # EE_Pipe Treatment
        
        print(f"Cột C (System Type): {col_c_name}")
        print(f"Cột T (Pipe Treatment): {col_t_name}")
        
        if col_c_name and col_t_name:
            # Thống kê System Type
            system_types = df[col_c_name].value_counts().to_dict()
            print(f"\nSystem Types trong {sheet_name}:")
            for st, count in system_types.items():
                print(f"  {st}: {count} dòng")
            
            # Thống kê Pipe Treatment
            treatments = df[col_t_name].value_counts().to_dict()
            print(f"\nPipe Treatments trong {sheet_name}:")
            for tr, count in treatments.items():
                print(f"  {tr}: {count} dòng")
            
            # Kiểm tra missing data
            system_na = df[col_c_name].isna().sum()
            treatment_na = df[col_t_name].isna().sum()
            print(f"\nMissing data:")
            print(f"  System Type NA: {system_na}")
            print(f"  Pipe Treatment NA: {treatment_na}")
            
            # Test validation cho 3 dòng đầu
            print(f"\n🔧 TEST VALIDATION (3 dòng đầu):")
            for i in range(min(3, len(df))):
                row = df.iloc[i]
                system_type = row[col_c_name]
                pipe_treatment = row[col_t_name]
                
                print(f"Dòng {i+2}:")
                print(f"  System Type: {system_type} (NA: {pd.isna(system_type)})")
                print(f"  Pipe Treatment: {pipe_treatment} (NA: {pd.isna(pipe_treatment)})")
                
                # Logic validation
                if pd.isna(system_type) or pd.isna(pipe_treatment):
                    result = "SKIP: Thiếu dữ liệu"
                else:
                    system_type_str = str(system_type).strip()
                    pipe_treatment_str = str(pipe_treatment).strip()
                    
                    if system_type_str == "CP-INTERNAL":
                        expected = "GAL"
                    elif system_type_str in ["CP-EXTERNAL", "CW-DISTRIBUTION", "CW-ARRAY"]:
                        expected = "BLACK"
                    else:
                        result = "PASS (no rule)"
                        expected = None
                    
                    if expected:
                        if pipe_treatment_str == expected:
                            result = "PASS"
                        else:
                            result = f"FAIL: cần '{expected}', có '{pipe_treatment_str}'"
                    
                print(f"  Validation: {result}")
                print()
        else:
            print("❌ THIẾU CỘT QUAN TRỌNG!")
            
        print("\n")

if __name__ == "__main__":
    debug_pipe_treatment_validation()
