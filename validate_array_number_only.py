#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re
from pathlib import Path
from datetime import datetime

def validate_array_number_only(excel_file_path):
    """
    Chỉ validate Array Number cho 4 worksheet cụ thể
    """
    try:
        # 4 worksheet cần kiểm tra Array Number
        target_worksheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule', 
            'Pipe Accessory Schedule', 
            'Sprinkler Schedule'
        ]
        
        xl_file = pd.ExcelFile(excel_file_path)
        
        print("=== KIỂM TRA ARRAY NUMBER CHO 4 WORKSHEET ===")
        print(f"File: {excel_file_path}")
        print(f"Target worksheets: {target_worksheets}")
        print()
        
        all_results = {}
        total_pass = 0
        total_fail = 0
        total_rows = 0
        
        # Xử lý từng worksheet
        for sheet_name in target_worksheets:
            if sheet_name not in xl_file.sheet_names:
                print(f"❌ Không tìm thấy worksheet: {sheet_name}")
                continue
                
            print(f"=== WORKSHEET: {sheet_name} ===")
            
            # Đọc worksheet
            df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
            print(f"Số dòng: {len(df)}, Số cột: {len(df.columns)}")
            
            # Lấy cột A, B, D (index 0, 1, 3)
            col_a_name = df.columns[0]  # EE_Cross Passage
            col_b_name = df.columns[1]  # EE_Location and Lanes  
            col_d_name = df.columns[3]  # EE_Array Number
            
            print(f"Cột A: {col_a_name}")
            print(f"Cột B: {col_b_name}")
            print(f"Cột D: {col_d_name}")
            
            # Áp dụng Array Number validation
            df['Array_Number_Check'] = df.apply(
                lambda row: check_array_number_rule(row, col_a_name, col_b_name, col_d_name), 
                axis=1
            )
            
            # Thống kê worksheet
            sheet_total = len(df)
            sheet_pass = len(df[df['Array_Number_Check'] == 'PASS'])
            sheet_fail = len(df[df['Array_Number_Check'] != 'PASS'])
            
            print(f"✅ PASS: {sheet_pass}/{sheet_total} ({sheet_pass/sheet_total*100:.1f}%)")
            print(f"❌ FAIL: {sheet_fail}/{sheet_total} ({sheet_fail/sheet_total*100:.1f}%)")
            
            # Cộng dồn
            total_rows += sheet_total
            total_pass += sheet_pass  
            total_fail += sheet_fail
            
            # Lưu kết quả
            all_results[sheet_name] = df
            
            # Hiển thị một số lỗi mẫu
            fail_rows = df[df['Array_Number_Check'] != 'PASS']
            if not fail_rows.empty:
                print(f"Lỗi mẫu (5 dòng đầu):")
                for idx, row in fail_rows.head(5).iterrows():
                    col_a = row[col_a_name]
                    col_b = row[col_b_name] 
                    col_d = row[col_d_name]
                    check_result = row['Array_Number_Check']
                    print(f"  Dòng {idx+2:3d}: A={col_a} | B={col_b} | D={col_d} | {check_result}")
            
            print()
        
        # Thống kê tổng
        print("=== TỔNG KẾT ARRAY NUMBER VALIDATION ===")
        print(f"✅ PASS: {total_pass}/{total_rows} ({total_pass/total_rows*100:.1f}%)")
        print(f"❌ FAIL: {total_fail}/{total_rows} ({total_fail/total_rows*100:.1f}%)")
        
        # Xuất file kết quả
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"array_number_validation_{timestamp}.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in all_results.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n📁 File kết quả đã lưu: {output_file}")
        
        return all_results
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return None

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
            return "SKIP: Thiếu dữ liệu"
        
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
            return f"FAIL: cần '{required_pattern}', có '{actual_array}'"
            
    except Exception as e:
        return f"ERROR: {str(e)}"

if __name__ == "__main__":
    # Tự động tìm file Excel
    current_dir = Path(".")
    excel_files = [f for f in current_dir.glob("*.xlsx") if not f.name.startswith('~') and 'validation' not in f.name.lower() and 'array_number' not in f.name.lower()]
    
    if not excel_files:
        print("❌ Không tìm thấy file Excel!")
        exit()
    
    print("📁 File Excel có sẵn:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {file.name}")
    
    try:
        choice = int(input(f"\nChọn file (1-{len(excel_files)}): ")) - 1
        selected_file = excel_files[choice]
        validate_array_number_only(selected_file)
    except (ValueError, IndexError):
        print("❌ Lựa chọn không hợp lệ!")
    except KeyboardInterrupt:
        print("\n⏹️ Đã hủy!")
