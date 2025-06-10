#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

def analyze_worksheet_structures():
    """
    Phân tích cấu trúc chi tiết của từng worksheet
    """
    excel_file = '../Xp03-Fabrication & Listing.xlsx'
    xl_file = pd.ExcelFile(excel_file)
    
    target_worksheets = [
        'Pipe Schedule', 
        'Pipe Fitting Schedule', 
        'Pipe Accessory Schedule', 
        'Sprinkler Schedule'
    ]
    
    for sheet_name in xl_file.sheet_names:
        print(f"{'='*60}")
        print(f"WORKSHEET: {sheet_name}")
        print(f"{'='*60}")
        
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        print(f"Số dòng: {len(df)}, Số cột: {len(df.columns)}")
        
        print(f"\nCẤU TRÚC CỘT:")
        for i, col in enumerate(df.columns):
            sample_data = df[col].dropna().head(2).tolist()
            print(f"{i+1:2d}. {col:30s} | Mẫu: {sample_data}")
        
        # Kiểm tra các cột quan trọng cho Array Number validation
        array_validation_cols = ['EE_Cross Passage', 'EE_Location and Lanes', 'EE_Array Number']
        print(f"\nCÁC CỘT CHO ARRAY NUMBER VALIDATION:")
        for col in array_validation_cols:
            if col in df.columns:
                sample = df[col].dropna().head(3).tolist()
                print(f"✅ {col:25s} | Mẫu: {sample}")
            else:
                print(f"❌ {col:25s} | THIẾU")
          # Kiểm tra các cột quan trọng cho Pipe Treatment validation
        pipe_treatment_cols = ['EE_System Type', 'EE_Pipe Treatment']
        print(f"\nCÁC CỘT CHO PIPE TREATMENT VALIDATION:")
        for col in pipe_treatment_cols:
            if col in df.columns:
                sample = df[col].dropna().head(3).tolist()
                print(f"✅ {col:25s} | Mẫu: {sample}")
            else:
                print(f"❌ {col:25s} | THIẾU")
        
        # Kiểm tra các cột khác cho validation
        other_validation_cols = ['EE_FAB Pipe', 'EE_PIPE END-1', 'EE_PIPE END-2', 'Size', 'Length']
        print(f"\nCÁC CỘT KHÁC CHO VALIDATION:")
        for col in other_validation_cols:
            if col in df.columns:
                sample = df[col].dropna().head(3).tolist()
                print(f"✅ {col:25s} | Mẫu: {sample}")
            else:
                print(f"❌ {col:25s} | THIẾU")
        
        # Kiểm tra các cột cho FAB Pipe, Pap validation
        new_validation_cols = ['EE_FAB Pipe', 'EE_Pap 1', 'EE_Pap 2', 'Item Description']
        print(f"\nCÁC CỘT CHO NEW VALIDATION RULES:")
        for col in new_validation_cols:
            if col in df.columns:
                sample = df[col].dropna().head(5).tolist()
                print(f"✅ {col:25s} | Mẫu: {sample}")
            else:
                print(f"❌ {col:25s} | THIẾU")
        
        # Phân tích mẫu dữ liệu cho Array Number validation
        if all(col in df.columns for col in array_validation_cols):
            print(f"\n📊 PHÂN TÍCH ARRAY NUMBER VALIDATION (5 dòng đầu):")
            for i in range(min(5, len(df))):
                row = df.iloc[i]
                col_a = row['EE_Cross Passage']
                col_b = row['EE_Location and Lanes'] 
                col_d = row['EE_Array Number']
                
                print(f"Dòng {i+2}:")
                print(f"  Cross Passage: {col_a}")
                print(f"  Location Lanes: {col_b}")
                print(f"  Array Number: {col_d}")
                
                # Thử tính pattern
                if pd.notna(col_a) and pd.notna(col_b):
                    try:
                        import re
                        # Lấy số từ cột B
                        numbers_b = re.findall(r'\d+', str(col_b))
                        last_2_b = numbers_b[-1][-2:] if numbers_b and len(numbers_b[-1]) >= 2 else "00"
                        
                        # Lấy số từ cột A  
                        numbers_a = re.findall(r'\d+', str(col_a))
                        last_2_a = numbers_a[-1][-2:] if numbers_a and len(numbers_a[-1]) >= 2 else "00"
                        
                        expected = f"EXP6{last_2_b}{last_2_a}"
                        actual = str(col_d)
                        
                        print(f"  Expected: {expected}")
                        print(f"  Match: {'✅' if expected in actual else '❌'}")
                    except Exception as e:
                        print(f"  Expected: ERROR - {e}")
                print()
        
        # Phân tích mẫu dữ liệu cho Pipe Treatment validation
        if all(col in df.columns for col in pipe_treatment_cols):
            print(f"\n🔧 PHÂN TÍCH PIPE TREATMENT VALIDATION (5 dòng đầu):")
            for i in range(min(5, len(df))):
                row = df.iloc[i]
                system_type = row['EE_System Type']
                pipe_treatment = row['EE_Pipe Treatment']
                
                print(f"Dòng {i+2}:")
                print(f"  System Type: {system_type}")
                print(f"  Pipe Treatment: {pipe_treatment}")
                
                # Kiểm tra rule
                if pd.notna(system_type):
                    system_str = str(system_type).upper()
                    treatment_str = str(pipe_treatment).upper() if pd.notna(pipe_treatment) else ""
                    
                    expected = ""
                    if "CP-INTERNAL" in system_str:
                        expected = "GAL"
                    elif any(x in system_str for x in ["CP-EXTERNAL", "CW-DISTRIBUTION", "CW-ARRAY"]):
                        expected = "BLACK"
                    
                    if expected:
                        match = expected in treatment_str
                        print(f"  Expected: {expected}")
                        print(f"  Match: {'✅' if match else '❌'}")
                    else:
                        print(f"  Expected: NO RULE")
                print()
        
        # Phân tích chi tiết dữ liệu Pap 1 để hiểu format
        if 'EE_Pap 1' in df.columns:
            print(f"\n📏 PHÂN TÍCH EE_PAP 1 VALIDATION (10 dòng đầu):")
            pap1_data = df['EE_Pap 1'].dropna().head(10)
            item_desc_data = df['Item Description'].head(10) if 'Item Description' in df.columns else None
            
            for i, pap1_val in enumerate(pap1_data):
                item_desc = item_desc_data.iloc[i] if item_desc_data is not None and i < len(item_desc_data) else "N/A"
                print(f"Dòng {i+2}:")
                print(f"  Item Description: {item_desc}")
                print(f"  EE_Pap 1: {pap1_val}")
                
                # Phân tích pattern
                pap1_str = str(pap1_val)
                if 'x' in pap1_str:
                    print(f"  → Format: Dimension (chứa 'x')")
                elif any(char.isalpha() for char in pap1_str):
                    print(f"  → Format: Size + Letter (như 40B, 65LR)")
