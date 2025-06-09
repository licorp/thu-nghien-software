#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import os
from pathlib import Path
from datetime import datetime

def validate_excel_multi_worksheets(excel_file_path):
    """
    Validate Excel file với nhiều worksheet, áp dụng Array Number validation chỉ cho 4 worksheet cụ thể
    """
    try:
        # Các worksheet cần áp dụng Array Number validation
        target_worksheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule', 
            'Pipe Accessory Schedule', 
            'Sprinkler Schedule'
        ]
        
        xl_file = pd.ExcelFile(excel_file_path)
        
        print("=== PHÂN TÍCH FILE EXCEL MULTI-WORKSHEET ===")
        print(f"File: {excel_file_path}")
        print(f"Số worksheet: {len(xl_file.sheet_names)}")
        print(f"Target worksheets: {target_worksheets}")
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
            
            # Kiểm tra có áp dụng Array Number validation không
            apply_array_validation = sheet_name in target_worksheets
            print(f"Array Number validation: {'✅ ÁP DỤNG' if apply_array_validation else '❌ KHÔNG ÁP DỤNG'}")
            
            # Áp dụng validation
            df['Validation_Check'] = df.apply(
                lambda row: validate_row_conditions(row, apply_array_validation=apply_array_validation), 
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
                    item_desc = row.get('EE_Item Description', 'N/A')
                    fab_pipe = row.get('EE_FAB Pipe', 'N/A')
                    validation_result = row['Validation_Check']
                    print(f"  Dòng {idx+2:3d}: {item_desc} | {fab_pipe} | {validation_result}")
            
            print()
        
        # Thống kê tổng
        print("=== TỔNG KẾT TẤT CẢ WORKSHEET ===")
        print(f"✅ PASS: {total_pass}/{total_rows} ({total_pass/total_rows*100:.1f}%)")
        print(f"❌ FAIL: {total_fail}/{total_rows} ({total_fail/total_rows*100:.1f}%)")
        
        # Xuất file kết quả
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"validation_multi_worksheet_{timestamp}.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in all_results.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n📁 File kết quả đã lưu: {output_file}")
        
        return all_results
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return None

def validate_row_conditions(row, apply_array_validation=True):
    """
    Validate từng dòng dữ liệu theo business rules
    """
    try:
        # Lấy dữ liệu từ các cột
        item_desc = str(row.get('EE_Item Description', '')).strip() if pd.notna(row.get('EE_Item Description')) else ''
        size = row.get('Size', '')
        fab_pipe = str(row.get('EE_FAB Pipe', '')).strip() if pd.notna(row.get('EE_FAB Pipe')) else ''
        pipe_end1 = str(row.get('EE_PIPE END-1', '')).strip() if pd.notna(row.get('EE_PIPE END-1')) else ''
        pipe_end2 = str(row.get('EE_PIPE END-2', '')).strip() if pd.notna(row.get('EE_PIPE END-2')) else ''
        
        errors = []
        
        # Rule 1: Kiểm tra các trường bắt buộc
        if not fab_pipe:
            errors.append("EE_FAB Pipe trống")
        if not pipe_end1:
            errors.append("EE_PIPE END-1 trống") 
        if not pipe_end2:
            errors.append("EE_PIPE END-2 trống")
        
        # Rule 2: Kiểm tra Size hợp lệ
        if pd.isna(size) or (isinstance(size, str) and size.strip() == ''):
            errors.append("Size trống")
        elif isinstance(size, (int, float)) and size <= 0:
            errors.append("Size ≤ 0")
        
        # Rule 3: Business logic cho Groove_Thread
        if 'Groove_Thread' in fab_pipe:
            if pipe_end1 != pipe_end2:
                errors.append(f"Groove_Thread: END-1({pipe_end1}) ≠ END-2({pipe_end2})")
        
        # Rule 4: Business logic cho STD PAP RANGE
        if 'STD' in fab_pipe and 'PAP RANGE' in fab_pipe:
            if pipe_end1 != 'RG':
                errors.append(f"STD PAP RANGE: END-1 cần RG, có {pipe_end1}")
            if pipe_end2 != 'BE':
                errors.append(f"STD PAP RANGE: END-2 cần BE, có {pipe_end2}")
        
        # Rule 5: Business logic cho STD ARRAY TEE
        if 'STD ARRAY TEE' in fab_pipe:
            if pipe_end1 != 'RG' or pipe_end2 != 'RG':
                errors.append(f"STD ARRAY TEE: cần RG-RG, có {pipe_end1}-{pipe_end2}")
        
        # Rule 6: Business logic cho Fabrication
        if 'Fabrication' in fab_pipe:
            if pipe_end1 != 'RG':
                errors.append(f"Fabrication: END-1 cần RG, có {pipe_end1}")
            if pipe_end2 != 'BE':
                errors.append(f"Fabrication: END-2 cần BE, có {pipe_end2}")
        
        # Rule 8: Kiểm tra EE_Item Description = Size + "-" + Length (làm tròn 5)
        length = row.get('Length', '')
        if pd.notna(length) and pd.notna(size) and length != '' and size != '':
            try:
                # Làm tròn Length với bội số của 5
                length_rounded = round(float(length) / 5) * 5
                # Tạo expected value: Size + "-" + Length_rounded
                expected_item_desc = f"{int(size)}-{int(length_rounded)}"
                
                # So sánh với EE_Item Description thực tế
                actual_item_desc = str(row.get('EE_Item Description', '')).strip()
                if actual_item_desc != expected_item_desc:
                    errors.append(f"Item Description: cần '{expected_item_desc}', có '{actual_item_desc}'")
            except (ValueError, TypeError):
                errors.append("Không thể tính Item Description (Size/Length lỗi)")
        
        # Rule 9: CHỈ ÁP DỤNG CHO 4 WORKSHEET CỤ THỂ - Kiểm tra EE_Array Number
        if apply_array_validation:
            cross_passage = row.get('EE_Cross Passage', '')  # Cột A
            location_lanes = row.get('EE_Location and Lanes', '')  # Cột B  
            array_number = row.get('EE_Array Number', '')  # Cột D
            
            if pd.notna(cross_passage) and pd.notna(location_lanes) and pd.notna(array_number):
                try:
                    # Lấy 2 số cuối của cột B (EE_Location and Lanes)
                    location_str = str(location_lanes).strip()
                    # Tìm số trong string, lấy 2 số cuối
                    import re
                    numbers_in_location = re.findall(r'\d+', location_str)
                    if numbers_in_location:
                        last_2_b = numbers_in_location[-1][-2:] if len(numbers_in_location[-1]) >= 2 else numbers_in_location[-1].zfill(2)
                    else:
                        last_2_b = "00"
                    
                    # Lấy 2 số cuối của cột A (EE_Cross Passage)
                    cross_str = str(cross_passage).strip()
                    numbers_in_cross = re.findall(r'\d+', cross_str)
                    if numbers_in_cross:
                        last_2_a = numbers_in_cross[-1][-2:] if len(numbers_in_cross[-1]) >= 2 else numbers_in_cross[-1].zfill(2)
                    else:
                        last_2_a = "00"
                    
                    # Tạo expected pattern (phải chứa trong array number)
                    required_pattern = f"EXP6{last_2_b}{last_2_a}"
                    actual_array = str(array_number).strip()
                    
                    # Kiểm tra xem array number có chứa pattern bắt buộc không
                    if required_pattern not in actual_array:
                        errors.append(f"Array Number: phải chứa '{required_pattern}', có '{actual_array}'")
                        
                except Exception as e:
                    errors.append(f"Không thể tính Array Number: {str(e)}")
        
        # Trả về kết quả
        if errors:
            return f"FAIL: {'; '.join(errors[:2])}"  # Chỉ hiển thị 2 lỗi đầu để không quá dài
        else:
            return "PASS"
            
    except Exception as e:
        return f"ERROR: {str(e)}"

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
        validate_excel_multi_worksheets(selected_file)
    except (ValueError, IndexError):
        print("❌ Lựa chọn không hợp lệ!")
    except KeyboardInterrupt:
        print("\n⏹️ Đã hủy!")
