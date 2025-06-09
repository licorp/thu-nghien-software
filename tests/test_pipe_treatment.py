#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
from pathlib import Path

def test_pipe_treatment_validation():
    """
    Test chi tiết Pipe Treatment validation
    """
    excel_file = 'Xp02-Fabrication & Listing.xlsx'
    
    # 3 worksheet cần kiểm tra Pipe Treatment
    target_worksheets = [
        'Pipe Schedule', 
        'Pipe Fitting Schedule', 
        'Pipe Accessory Schedule'
    ]
    
    xl_file = pd.ExcelFile(excel_file)
    
    print("=== TEST CHI TIẾT PIPE TREATMENT VALIDATION ===")
    print(f"File: {excel_file}")
    print("Quy tắc:")
    print("  - CP-INTERNAL → GAL")
    print("  - CP-EXTERNAL, CW-DISTRIBUTION, CW-ARRAY → BLACK")
    print()
    
    for sheet_name in target_worksheets:
        print(f"{'='*60}")
        print(f"WORKSHEET: {sheet_name}")
        print(f"{'='*60}")
        
        # Đọc worksheet
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # Cột C (index 2) = EE_System Type
        # Cột T (index 19) = EE_Pipe Treatment
        col_c_name = df.columns[2]  # EE_System Type
        col_t_name = df.columns[19]  # EE_Pipe Treatment
        
        print(f"Cột C: {col_c_name}")
        print(f"Cột T: {col_t_name}")
        print()
        
        # Phân tích các giá trị unique
        print("📊 PHÂN TÍCH GIÁ TRỊ UNIQUE:")
        system_types = df[col_c_name].value_counts()
        treatments = df[col_t_name].value_counts()
        
        print(f"System Types (Cột C):")
        for sys_type, count in system_types.items():
            print(f"  {sys_type}: {count} dòng")
        
        print(f"\nPipe Treatments (Cột T):")
        for treatment, count in treatments.items():
            print(f"  {treatment}: {count} dòng")
        
        # Test validation chi tiết
        print(f"\n🧪 TEST VALIDATION CHI TIẾT:")
        
        pass_count = 0
        fail_count = 0
        skip_count = 0
        error_details = []
        
        for idx, row in df.iterrows():
            system_type = row[col_c_name]
            pipe_treatment = row[col_t_name]
            
            # Kiểm tra dữ liệu
            if pd.isna(system_type) or pd.isna(pipe_treatment):
                skip_count += 1
                continue
            
            system_type_str = str(system_type).strip()
            pipe_treatment_str = str(pipe_treatment).strip()
            
            # Áp dụng quy tắc
            expected_treatment = None
            if system_type_str == "CP-INTERNAL":
                expected_treatment = "GAL"
            elif system_type_str in ["CP-EXTERNAL", "CW-DISTRIBUTION", "CW-ARRAY"]:
                expected_treatment = "BLACK"
            
            if expected_treatment:
                if pipe_treatment_str == expected_treatment:
                    pass_count += 1
                else:
                    fail_count += 1
                    error_details.append({
                        'row': idx + 2,
                        'system_type': system_type_str,
                        'expected': expected_treatment,
                        'actual': pipe_treatment_str
                    })
            else:
                # Không áp dụng rule cho system type khác
                pass_count += 1
        
        # Thống kê
        total_checked = pass_count + fail_count
        print(f"✅ PASS: {pass_count}/{total_checked} ({pass_count/total_checked*100:.1f}%)")
        print(f"❌ FAIL: {fail_count}/{total_checked} ({fail_count/total_checked*100:.1f}%)")
        print(f"⏭️ SKIP: {skip_count} dòng (thiếu dữ liệu)")
        
        # Hiển thị lỗi chi tiết
        if error_details:
            print(f"\n❌ CHI TIẾT CÁC LỖI (tối đa 10 dòng đầu):")
            for error in error_details[:10]:
                print(f"  Dòng {error['row']:3d}: {error['system_type']} → cần '{error['expected']}', có '{error['actual']}'")
        
        # Hiển thị ma trận kết hợp
        print(f"\n📋 MA TRẬN SYSTEM TYPE vs PIPE TREATMENT:")
        matrix = df.groupby([col_c_name, col_t_name]).size().reset_index(name='count')
        for _, row in matrix.iterrows():
            sys_type = row[col_c_name]
            treatment = row[col_t_name]
            count = row['count']
            
            # Kiểm tra có đúng quy tắc không
            is_correct = False
            if str(sys_type) == "CP-INTERNAL" and str(treatment) == "GAL":
                is_correct = True
            elif str(sys_type) in ["CP-EXTERNAL", "CW-DISTRIBUTION", "CW-ARRAY"] and str(treatment) == "BLACK":
                is_correct = True
            elif str(sys_type) not in ["CP-INTERNAL", "CP-EXTERNAL", "CW-DISTRIBUTION", "CW-ARRAY"]:
                is_correct = True  # Không áp dụng rule
            
            status = "✅" if is_correct else "❌"
            print(f"  {status} {sys_type} + {treatment}: {count} dòng")
        
        print()

if __name__ == "__main__":
    test_pipe_treatment_validation()
