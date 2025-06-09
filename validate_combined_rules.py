#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re
from pathlib import Path
from datetime import datetime

def validate_array_number_and_pipe_treatment(excel_file_path):
    """
    Validation cho:
    1. Array Number (4 worksheet)
    2. Pipe Treatment (3 worksheet)
    """
    try:
        # Worksheet áp dụng Array Number validation
        array_number_worksheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule', 
            'Pipe Accessory Schedule', 
            'Sprinkler Schedule'
        ]
        
        # Worksheet áp dụng Pipe Treatment validation  
        pipe_treatment_worksheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule', 
            'Pipe Accessory Schedule'
        ]
        
        xl_file = pd.ExcelFile(excel_file_path)
        
        print("=== VALIDATION ARRAY NUMBER + PIPE TREATMENT ===")
        print(f"File: {excel_file_path}")
        print(f"Array Number worksheets: {array_number_worksheets}")
        print(f"Pipe Treatment worksheets: {pipe_treatment_worksheets}")
        print()
        
        all_results = {}
        total_pass = 0
        total_fail = 0
        total_rows = 0
        
        # Xử lý từng worksheet
        for sheet_name in xl_file.sheet_names:
            print(f"=== WORKSHEET: {sheet_name} ===")
            
            # Đọc worksheet
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
            print(f"Số dòng: {len(df)}, Số cột: {len(df.columns)}")
            
            # Kiểm tra worksheet nào cần validation gì
            apply_array_validation = sheet_name in array_number_worksheets
            apply_pipe_treatment_validation = sheet_name in pipe_treatment_worksheets
            
            print(f"Array Number validation: {'✅ ÁP DỤNG' if apply_array_validation else '❌ KHÔNG ÁP DỤNG'}")
            print(f"Pipe Treatment validation: {'✅ ÁP DỤNG' if apply_pipe_treatment_validation else '❌ KHÔNG ÁP DỤNG'}")
            
            # Lấy tên cột (theo vị trí)
            col_a_name = df.columns[0] if len(df.columns) > 0 else None  # EE_Cross Passage
            col_b_name = df.columns[1] if len(df.columns) > 1 else None  # EE_Location and Lanes  
            col_c_name = df.columns[2] if len(df.columns) > 2 else None  # EE_System Type
            col_d_name = df.columns[3] if len(df.columns) > 3 else None  # EE_Array Number
            col_t_name = df.columns[19] if len(df.columns) > 19 else None  # EE_Pipe Treatment (cột T = index 19)
            
            print(f"Cột A: {col_a_name}")
            print(f"Cột B: {col_b_name}")
            print(f"Cột C: {col_c_name}")
            print(f"Cột D: {col_d_name}")
            print(f"Cột T: {col_t_name}")
            
            # Áp dụng validation
            df['Validation_Check'] = df.apply(
                lambda row: validate_combined_rules(
                    row, 
                    col_a_name, col_b_name, col_c_name, col_d_name, col_t_name,
                    apply_array_validation, apply_pipe_treatment_validation
                ), 
                axis=1
            )
            
            # Thống kê worksheet
            sheet_total = len(df)
            sheet_pass = len(df[df['Validation_Check'] == 'PASS'])
            sheet_fail = len(df[df['Validation_Check'] != 'PASS'])
            
            print(f"✅ PASS: {sheet_pass}/{sheet_total} ({sheet_pass/sheet_total*100:.1f}%)")
            print(f"❌ FAIL: {sheet_fail}/{sheet_total} ({sheet_fail/sheet_total*100:.1f}%)")
            
            # Cộng dồn
            total_rows += sheet_total
            total_pass += sheet_pass  
            total_fail += sheet_fail
            
            # Lưu kết quả
            all_results[sheet_name] = df
            
            # Hiển thị một số lỗi mẫu
            fail_rows = df[df['Validation_Check'] != 'PASS']
            if not fail_rows.empty:
                print(f"Lỗi mẫu (5 dòng đầu):")
                for idx, row in fail_rows.head(5).iterrows():
                    col_c = row[col_c_name] if col_c_name else 'N/A'
                    col_d = row[col_d_name] if col_d_name else 'N/A' 
                    col_t = row[col_t_name] if col_t_name else 'N/A'
                    check_result = row['Validation_Check']
                    print(f"  Dòng {idx+2:3d}: C={col_c} | D={col_d} | T={col_t} | {check_result}")
            
            print()
        
        # Thống kê tổng
        print("=== TỔNG KẾT VALIDATION ===")
        print(f"✅ PASS: {total_pass}/{total_rows} ({total_pass/total_rows*100:.1f}%)")
        print(f"❌ FAIL: {total_fail}/{total_rows} ({total_fail/total_rows*100:.1f}%)")
        
        # Xuất file kết quả
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"validation_combined_{timestamp}.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in all_results.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n📁 File kết quả đã lưu: {output_file}")
        
        return all_results
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return None

def validate_combined_rules(row, col_a_name, col_b_name, col_c_name, col_d_name, col_t_name, 
                           apply_array_validation, apply_pipe_treatment_validation):
    """
    Áp dụng kết hợp các rule validation
    """
    errors = []
    
    try:
        # Rule 1: Array Number validation (nếu áp dụng)
        if apply_array_validation and col_a_name and col_b_name and col_d_name:
            array_result = check_array_number_rule(row, col_a_name, col_b_name, col_d_name)
            if array_result != "PASS" and not array_result.startswith("SKIP"):
                errors.append(f"Array: {array_result}")
        
        # Rule 2: Pipe Treatment validation (nếu áp dụng)
        if apply_pipe_treatment_validation and col_c_name and col_t_name:
            pipe_treatment_result = check_pipe_treatment_rule(row, col_c_name, col_t_name)
            if pipe_treatment_result != "PASS" and not pipe_treatment_result.startswith("SKIP"):
                errors.append(f"Treatment: {pipe_treatment_result}")
        
        # Trả về kết quả
        if errors:
            return f"FAIL: {'; '.join(errors[:2])}"  # Chỉ hiển thị 2 lỗi đầu
        else:
            return "PASS"
            
    except Exception as e:
        return f"ERROR: {str(e)}"

def check_array_number_rule(row, col_a_name, col_b_name, col_d_name):
    """
    Kiểm tra Array Number rule: EE_Array Number phải chứa "EXP6" + 2 số cuối cột B + 2 số cuối cột A
    """
    try:
        cross_passage = row[col_a_name]  # Cột A
        location_lanes = row[col_b_name]  # Cột B  
        array_number = row[col_d_name]  # Cột D
        
        # Kiểm tra dữ liệu có hợp lệ không
        if pd.isna(cross_passage) or pd.isna(location_lanes) or pd.isna(array_number):
            return "SKIP: Thiếu dữ liệu Array"
        
        # Lấy 2 số cuối của cột B
        location_str = str(location_lanes).strip()
        numbers_in_location = re.findall(r'\d+', location_str)
        if numbers_in_location:
            last_2_b = numbers_in_location[-1][-2:] if len(numbers_in_location[-1]) >= 2 else numbers_in_location[-1].zfill(2)
        else:
            last_2_b = "00"
        
        # Lấy 2 số cuối của cột A
        cross_str = str(cross_passage).strip()
        numbers_in_cross = re.findall(r'\d+', cross_str)
        if numbers_in_cross:
            last_2_a = numbers_in_cross[-1][-2:] if len(numbers_in_cross[-1]) >= 2 else numbers_in_cross[-1].zfill(2)
        else:
            last_2_a = "00"
        
        # Tạo pattern bắt buộc
        required_pattern = f"EXP6{last_2_b}{last_2_a}"
        actual_array = str(array_number).strip()
        
        # Kiểm tra
        if required_pattern in actual_array:
            return "PASS"
        else:
            return f"cần '{required_pattern}', có '{actual_array}'"
            
    except Exception as e:
        return f"ERROR Array: {str(e)}"

def check_pipe_treatment_rule(row, col_c_name, col_t_name):
    """
    Kiểm tra Pipe Treatment rule:
    - CP-INTERNAL → GAL
    - CP-EXTERNAL, CW-DISTRIBUTION, CW-ARRAY → BLACK
    """
    try:
        system_type = row[col_c_name]  # Cột C
        pipe_treatment = row[col_t_name]  # Cột T
        
        # Kiểm tra dữ liệu có hợp lệ không
        if pd.isna(system_type) or pd.isna(pipe_treatment):
            return "SKIP: Thiếu dữ liệu Treatment"
        
        system_type_str = str(system_type).strip()
        pipe_treatment_str = str(pipe_treatment).strip()
        
        # Quy tắc validation
        if system_type_str == "CP-INTERNAL":
            expected_treatment = "GAL"
        elif system_type_str in ["CP-EXTERNAL", "CW-DISTRIBUTION", "CW-ARRAY"]:
            expected_treatment = "BLACK"
        else:
            # Không áp dụng rule cho các system type khác
            return "PASS"
        
        # Kiểm tra
        if pipe_treatment_str == expected_treatment:
            return "PASS"
        else:
            return f"System '{system_type_str}' cần '{expected_treatment}', có '{pipe_treatment_str}'"
            
    except Exception as e:
        return f"ERROR Treatment: {str(e)}"

if __name__ == "__main__":
    # Tự động tìm file Excel
    current_dir = Path(".")
    excel_files = [f for f in current_dir.glob("*.xlsx") if not f.name.startswith('~') and 'validation' not in f.name.lower()]
    
    if not excel_files:
        print("❌ Không tìm thấy file Excel!")
        exit()
    
    print("📁 File Excel có sẵn:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {file.name}")
    
    try:
        choice = int(input(f"\nChọn file (1-{len(excel_files)}): ")) - 1
        selected_file = excel_files[choice]
        validate_array_number_and_pipe_treatment(selected_file)
    except (ValueError, IndexError):
        print("❌ Lựa chọn không hợp lệ!")
    except KeyboardInterrupt:
        print("\n⏹️ Đã hủy!")
