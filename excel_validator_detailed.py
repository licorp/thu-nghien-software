#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re
import os
from pathlib import Path
from datetime import datetime

class ExcelValidatorDetailed:
    """
    Excel Validator với hiển thị chi tiết từng loại validation
    """
    
    def __init__(self):
        # Cấu hình worksheet
        self.array_number_worksheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule', 
            'Pipe Accessory Schedule', 
            'Sprinkler Schedule'
        ]
        
        self.pipe_treatment_worksheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule', 
            'Pipe Accessory Schedule'
        ]
        
        # Thống kê validation chi tiết
        self.total_rows = 0
        self.array_pass = 0
        self.array_fail = 0
        self.array_skip = 0
        self.treatment_pass = 0
        self.treatment_fail = 0
        self.treatment_skip = 0
        
        self.validation_results = {}
        
    def validate_excel_file(self, excel_file_path):
        """
        Validate toàn bộ Excel file với chi tiết từng rule
        """
        try:
            print("=" * 80)
            print("🚀 EXCEL VALIDATION TOOL - CHI TIẾT TỪNG RULE")
            print("=" * 80)
            print(f"📁 File: {Path(excel_file_path).name}")
            print(f"🕐 Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            # Hiển thị cấu hình
            print(f"\n📋 CẤU HÌNH VALIDATION:")
            print(f"1. Array Number Validation:")
            print(f"   - Worksheets: {', '.join(self.array_number_worksheets)}")
            print(f"   - Quy tắc: Cột D phải chứa 'EXP6' + 2 số cuối cột B + 2 số cuối cột A")
            print(f"2. Pipe Treatment Validation:")
            print(f"   - Worksheets: {', '.join(self.pipe_treatment_worksheets)}")
            print(f"   - Quy tắc: CP-INTERNAL→GAL, CP-EXTERNAL/CW-DISTRIBUTION/CW-ARRAY→BLACK")
            
            # Đọc file Excel
            xl_file = pd.ExcelFile(excel_file_path)
            
            # Validate từng worksheet
            for sheet_name in xl_file.sheet_names:
                self._validate_worksheet_detailed(excel_file_path, sheet_name)
            
            # Tổng kết chi tiết
            self._generate_detailed_summary()
            
        except Exception as e:
            print(f"❌ Lỗi validation: {e}")
            return None
    
    def _validate_worksheet_detailed(self, excel_file_path, sheet_name):
        """
        Validate worksheet với hiển thị chi tiết từng rule
        """
        print(f"\n📊 WORKSHEET: {sheet_name}")
        print("-" * 60)
        
        # Đọc worksheet
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        print(f"Số dòng: {len(df)}, Số cột: {len(df.columns)}")
        
        # Kiểm tra worksheet nào cần validation gì
        apply_array_validation = sheet_name in self.array_number_worksheets
        apply_pipe_treatment_validation = sheet_name in self.pipe_treatment_worksheets
        
        print(f"🔢 Array Number validation: {'✅ ÁP DỤNG' if apply_array_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"🔧 Pipe Treatment validation: {'✅ ÁP DỤNG' if apply_pipe_treatment_validation else '❌ KHÔNG ÁP DỤNG'}")
        
        if not apply_array_validation and not apply_pipe_treatment_validation:
            print("⏭️ Bỏ qua worksheet (không có quy tắc nào áp dụng)")
            return
        
        # Lấy tên cột theo vị trí
        col_a_name = df.columns[0] if len(df.columns) > 0 else None  # EE_Cross Passage
        col_b_name = df.columns[1] if len(df.columns) > 1 else None  # EE_Location and Lanes  
        col_c_name = df.columns[2] if len(df.columns) > 2 else None  # EE_System Type
        col_d_name = df.columns[3] if len(df.columns) > 3 else None  # EE_Array Number
        col_t_name = df.columns[19] if len(df.columns) > 19 else None  # EE_Pipe Treatment
        
        # Áp dụng validation chi tiết
        array_results = []
        treatment_results = []
        
        for idx, row in df.iterrows():
            # Array Number validation
            if apply_array_validation:
                array_result = self._check_array_number_detailed(row, col_a_name, col_b_name, col_d_name)
                array_results.append(array_result)
            else:
                array_results.append("N/A")
            
            # Pipe Treatment validation  
            if apply_pipe_treatment_validation:
                treatment_result = self._check_pipe_treatment_detailed(row, col_c_name, col_t_name)
                treatment_results.append(treatment_result)
            else:
                treatment_results.append("N/A")
        
        # Thêm kết quả vào DataFrame
        df['Array_Check'] = array_results
        df['Treatment_Check'] = treatment_results
        
        # Thống kê chi tiết
        self._report_detailed_stats(df, sheet_name, apply_array_validation, apply_pipe_treatment_validation)
        
        # Hiển thị lỗi mẫu
        self._show_detailed_errors(df, sheet_name, col_c_name, col_d_name, col_t_name, 
                                 apply_array_validation, apply_pipe_treatment_validation)
        
        # Lưu kết quả
        self.validation_results[sheet_name] = df
        self.total_rows += len(df)
    
    def _check_array_number_detailed(self, row, col_a_name, col_b_name, col_d_name):
        """
        Kiểm tra Array Number rule chi tiết
        """
        try:
            if not col_a_name or not col_b_name or not col_d_name:
                return "SKIP: Thiếu cột"
                
            cross_passage = row[col_a_name]
            location_lanes = row[col_b_name]
            array_number = row[col_d_name]
            
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
            
            if required_pattern in actual_array:
                return "PASS"
            else:
                return f"FAIL: cần '{required_pattern}', có '{actual_array}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_pipe_treatment_detailed(self, row, col_c_name, col_t_name):
        """
        Kiểm tra Pipe Treatment rule chi tiết
        """
        try:
            if not col_c_name or not col_t_name:
                return "SKIP: Thiếu cột"
                
            system_type = row[col_c_name]
            pipe_treatment = row[col_t_name]
            
            if pd.isna(system_type) or pd.isna(pipe_treatment):
                return "SKIP: Thiếu dữ liệu"
            
            system_type_str = str(system_type).strip()
            pipe_treatment_str = str(pipe_treatment).strip()
            
            # Áp dụng quy tắc
            if system_type_str == "CP-INTERNAL":
                expected_treatment = "GAL"
            elif system_type_str in ["CP-EXTERNAL", "CW-DISTRIBUTION", "CW-ARRAY"]:
                expected_treatment = "BLACK"
            else:
                return "PASS: Không áp dụng rule"
            
            if pipe_treatment_str == expected_treatment:
                return "PASS"
            else:
                return f"FAIL: '{system_type_str}' cần '{expected_treatment}', có '{pipe_treatment_str}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _report_detailed_stats(self, df, sheet_name, apply_array, apply_treatment):
        """
        Báo cáo thống kê chi tiết cho worksheet
        """
        # Thống kê Array Number
        if apply_array:
            array_pass = len(df[df['Array_Check'] == 'PASS'])
            array_fail = len(df[df['Array_Check'].str.startswith('FAIL', na=False)])
            array_skip = len(df[df['Array_Check'].str.startswith('SKIP', na=False)])
            
            print(f"\n🔢 ARRAY NUMBER VALIDATION:")
            print(f"   ✅ PASS: {array_pass}")
            print(f"   ❌ FAIL: {array_fail}")
            print(f"   ⏭️ SKIP: {array_skip}")
            
            self.array_pass += array_pass
            self.array_fail += array_fail
            self.array_skip += array_skip
        
        # Thống kê Pipe Treatment
        if apply_treatment:
            treatment_pass = len(df[df['Treatment_Check'] == 'PASS']) + len(df[df['Treatment_Check'].str.contains('PASS:', na=False)])
            treatment_fail = len(df[df['Treatment_Check'].str.startswith('FAIL', na=False)])
            treatment_skip = len(df[df['Treatment_Check'].str.startswith('SKIP', na=False)])
            
            print(f"\n🔧 PIPE TREATMENT VALIDATION:")
            print(f"   ✅ PASS: {treatment_pass}")
            print(f"   ❌ FAIL: {treatment_fail}")
            print(f"   ⏭️ SKIP: {treatment_skip}")
            
            self.treatment_pass += treatment_pass
            self.treatment_fail += treatment_fail
            self.treatment_skip += treatment_skip
    
    def _show_detailed_errors(self, df, sheet_name, col_c_name, col_d_name, col_t_name, apply_array, apply_treatment):
        """
        Hiển thị lỗi chi tiết theo từng loại validation
        """
        # Lỗi Array Number
        if apply_array:
            array_errors = df[df['Array_Check'].str.startswith('FAIL', na=False)]
            if not array_errors.empty:
                print(f"\n❌ LỖI ARRAY NUMBER (5 dòng đầu):")
                for idx, row in array_errors.head(5).iterrows():
                    col_d = row[col_d_name] if col_d_name else 'N/A'
                    error_msg = row['Array_Check']
                    print(f"   Dòng {idx+2:3d}: D={col_d} | {error_msg}")
        
        # Lỗi Pipe Treatment  
        if apply_treatment:
            treatment_errors = df[df['Treatment_Check'].str.startswith('FAIL', na=False)]
            if not treatment_errors.empty:
                print(f"\n❌ LỖI PIPE TREATMENT (5 dòng đầu):")
                for idx, row in treatment_errors.head(5).iterrows():
                    col_c = row[col_c_name] if col_c_name else 'N/A'
                    col_t = row[col_t_name] if col_t_name else 'N/A'
                    error_msg = row['Treatment_Check']
                    print(f"   Dòng {idx+2:3d}: C={col_c} | T={col_t} | {error_msg}")
    
    def _generate_detailed_summary(self):
        """
        Tạo báo cáo tổng kết chi tiết
        """
        print("\n" + "=" * 80)
        print("📈 TỔNG KẾT VALIDATION CHI TIẾT")
        print("=" * 80)
        
        print(f"📊 Tổng số dòng đã kiểm tra: {self.total_rows:,}")
        
        # Tổng kết Array Number
        array_total = self.array_pass + self.array_fail + self.array_skip
        if array_total > 0:
            print(f"\n🔢 ARRAY NUMBER VALIDATION:")
            print(f"   ✅ PASS: {self.array_pass:,}/{array_total:,} ({self.array_pass/array_total*100:.1f}%)")
            print(f"   ❌ FAIL: {self.array_fail:,}/{array_total:,} ({self.array_fail/array_total*100:.1f}%)")
            print(f"   ⏭️ SKIP: {self.array_skip:,}/{array_total:,} ({self.array_skip/array_total*100:.1f}%)")
        
        # Tổng kết Pipe Treatment
        treatment_total = self.treatment_pass + self.treatment_fail + self.treatment_skip
        if treatment_total > 0:
            print(f"\n🔧 PIPE TREATMENT VALIDATION:")
            print(f"   ✅ PASS: {self.treatment_pass:,}/{treatment_total:,} ({self.treatment_pass/treatment_total*100:.1f}%)")
            print(f"   ❌ FAIL: {self.treatment_fail:,}/{treatment_total:,} ({self.treatment_fail/treatment_total*100:.1f}%)")
            print(f"   ⏭️ SKIP: {self.treatment_skip:,}/{treatment_total:,} ({self.treatment_skip/treatment_total*100:.1f}%)")
        
        print(f"\n🎉 VALIDATION HOÀN THÀNH!")

def main():
    """
    Main function để chạy validation
    """
    # Tìm file Excel
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~') and not f.startswith('validation')]
    
    if not excel_files:
        print("❌ Không tìm thấy file Excel nào!")
        return
    
    if len(excel_files) == 1:
        selected_file = excel_files[0]
    else:
        print("🔍 FILE EXCEL CÓ SẴN:")
        for i, file in enumerate(excel_files, 1):
            size = os.path.getsize(file) // 1024
            print(f" {i}. {file:40s} ({size} KB)")
        
        while True:
            choice = input("✏️ Chọn file (1-{}) hoặc 'q' để thoát: ".format(len(excel_files)))
            if choice.lower() == 'q':
                return
            try:
                file_idx = int(choice) - 1
                if 0 <= file_idx < len(excel_files):
                    selected_file = excel_files[file_idx]
                    break
                else:
                    print("❌ Số không hợp lệ!")
            except ValueError:
                print("❌ Vui lòng nhập số!")
    
    # Chạy validation
    validator = ExcelValidatorDetailed()
    validator.validate_excel_file(selected_file)

if __name__ == "__main__":
    main()
