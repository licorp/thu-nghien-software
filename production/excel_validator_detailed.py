#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EXCEL VALIDATION TOOL - PHIÊN BẢN HOÀN CHỈNH
============================================

Tool validation Excel cho dự án pipe/equipment data với các quy tắc:

1. Array Number Validation (4 worksheet):
   - Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule, Sprinkler Schedule
   - Quy tắc: Cột D phải chứa "EXP6" + 2 số cuối cột B + 2 số cuối cột A

2. Pipe Treatment Validation (3 worksheet):
   - Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule  
   - Quy tắc:
     * CP-INTERNAL → GAL
     * CP-EXTERNAL, CW-DISTRIBUTION, CW-ARRAY → BLACK

Tác giả: GitHub Copilot
Ngày tạo: 2025-06-09
"""

import pandas as pd
import re
import os
from pathlib import Path
from datetime import datetime

class ExcelValidator:
    """
    Class chính cho Excel validation với đầy đủ các quy tắc
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
        
        # Thống kê validation
        self.total_rows = 0
        self.total_pass = 0
        self.total_fail = 0
        self.validation_results = {}
    
    def validate_excel_file(self, excel_file_path):
        """
        Validate toàn bộ file Excel với tất cả các quy tắc
        """
        try:
            print("=" * 80)
            print("🚀 EXCEL VALIDATION TOOL - PHIÊN BẢN HOÀN CHỈNH")
            print("=" * 80)
            print(f"📁 File: {excel_file_path}")
            print(f"🕐 Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print()
            
            # Đọc Excel file
            xl_file = pd.ExcelFile(excel_file_path)
            
            print("📋 CÁC QUY TẮC VALIDATION:")
            print("1. Array Number Validation:")
            print(f"   - Worksheets: {', '.join(self.array_number_worksheets)}")
            print("   - Quy tắc: Cột D phải chứa 'EXP6' + 2 số cuối cột B + 2 số cuối cột A")
            print("2. Pipe Treatment Validation:")
            print(f"   - Worksheets: {', '.join(self.pipe_treatment_worksheets)}")
            print("   - Quy tắc: CP-INTERNAL→GAL, CP-EXTERNAL/CW-DISTRIBUTION/CW-ARRAY→BLACK")
            print()
            
            # Xử lý từng worksheet
            for sheet_name in xl_file.sheet_names:
                self._validate_worksheet(excel_file_path, sheet_name)
            
            # Tổng kết và xuất file
            self._generate_summary()
            output_file = self._export_results(excel_file_path)
            
            return output_file
            
        except Exception as e:
            print(f"❌ Lỗi validation: {e}")
            return None
    
    def _validate_worksheet(self, excel_file_path, sheet_name):
        """
        Validate một worksheet cụ thể
        """
        print(f"📊 WORKSHEET: {sheet_name}")
        print("-" * 50)
        
        # Đọc worksheet
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        print(f"Số dòng: {len(df)}, Số cột: {len(df.columns)}")
        
        # Kiểm tra worksheet nào cần validation gì
        apply_array_validation = sheet_name in self.array_number_worksheets
        apply_pipe_treatment_validation = sheet_name in self.pipe_treatment_worksheets
        
        print(f"Array Number validation: {'✅ ÁP DỤNG' if apply_array_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"Pipe Treatment validation: {'✅ ÁP DỤNG' if apply_pipe_treatment_validation else '❌ KHÔNG ÁP DỤNG'}")
        
        if not apply_array_validation and not apply_pipe_treatment_validation:
            print("⏭️ Bỏ qua worksheet (không có quy tắc nào áp dụng)")
            print()
            return
        
        # Lấy tên cột theo vị trí
        col_a_name = df.columns[0] if len(df.columns) > 0 else None  # EE_Cross Passage
        col_b_name = df.columns[1] if len(df.columns) > 1 else None  # EE_Location and Lanes  
        col_c_name = df.columns[2] if len(df.columns) > 2 else None  # EE_System Type
        col_d_name = df.columns[3] if len(df.columns) > 3 else None  # EE_Array Number
        col_t_name = df.columns[19] if len(df.columns) > 19 else None  # EE_Pipe Treatment
        
        # Áp dụng validation
        df['Validation_Check'] = df.apply(
            lambda row: self._validate_row(
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
        
        # Cộng dồn thống kê
        self.total_rows += sheet_total
        self.total_pass += sheet_pass  
        self.total_fail += sheet_fail
        
        # Lưu kết quả
        self.validation_results[sheet_name] = df
        
        # Hiển thị lỗi mẫu
        self._show_sample_errors(df, sheet_name, col_c_name, col_d_name, col_t_name)
        print()
    
    def _validate_row(self, row, col_a_name, col_b_name, col_c_name, col_d_name, col_t_name, 
                     apply_array_validation, apply_pipe_treatment_validation):
        """
        Validate một dòng dữ liệu với tất cả các quy tắc
        """
        errors = []
        
        try:
            # Rule 1: Array Number validation
            if apply_array_validation and col_a_name and col_b_name and col_d_name:
                array_result = self._check_array_number(row, col_a_name, col_b_name, col_d_name)
                if array_result != "PASS" and not array_result.startswith("SKIP"):
                    errors.append(f"Array: {array_result}")
            
            # Rule 2: Pipe Treatment validation
            if apply_pipe_treatment_validation and col_c_name and col_t_name:
                treatment_result = self._check_pipe_treatment(row, col_c_name, col_t_name)
                if treatment_result != "PASS" and not treatment_result.startswith("SKIP"):
                    errors.append(f"Treatment: {treatment_result}")
            
            # Trả về kết quả
            if errors:
                return f"FAIL: {'; '.join(errors[:2])}"  # Chỉ hiển thị 2 lỗi đầu
            else:
                return "PASS"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_array_number(self, row, col_a_name, col_b_name, col_d_name):
        """
        Kiểm tra Array Number rule
        """
        try:
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
                return f"cần '{required_pattern}', có '{actual_array}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_pipe_treatment(self, row, col_c_name, col_t_name):
        """
        Kiểm tra Pipe Treatment rule
        """
        try:
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
                return "PASS"  # Không áp dụng rule cho system type khác
            
            if pipe_treatment_str == expected_treatment:
                return "PASS"
            else:                return f"'{system_type_str}' cần '{expected_treatment}', có '{pipe_treatment_str}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _show_sample_errors(self, df, sheet_name, col_c_name, col_d_name, col_t_name):
        """
        Hiển thị lỗi với tùy chọn xem tất cả
        """
        fail_rows = df[df['Validation_Check'] != 'PASS']
        if not fail_rows.empty:
            total_errors = len(fail_rows)
            
            # Nếu ít lỗi (<= 20), hiển thị tất cả
            if total_errors <= 20:
                print(f"📋 TẤT CẢ {total_errors} LỖI:")
                for idx, row in fail_rows.iterrows():
                    col_c = row[col_c_name] if col_c_name else 'N/A'
                    col_d = row[col_d_name] if col_d_name else 'N/A' 
                    col_t = row[col_t_name] if col_t_name else 'N/A'
                    check_result = row['Validation_Check']
                    print(f"  Dòng {idx+2:3d}: C={col_c} | D={col_d} | T={col_t}")
                    print(f"           {check_result}")
            else:
                # Nếu nhiều lỗi, hiển thị 15 đầu + 5 cuối
                print(f"📋 Tổng cộng {total_errors} lỗi - Hiển thị 15 đầu + 5 cuối:")
                print(f"\n🔺 15 LỖI ĐẦU TIÊN:")
                for idx, row in fail_rows.head(15).iterrows():
                    col_c = row[col_c_name] if col_c_name else 'N/A'
                    col_d = row[col_d_name] if col_d_name else 'N/A' 
                    col_t = row[col_t_name] if col_t_name else 'N/A'
                    check_result = row['Validation_Check']
                    print(f"  Dòng {idx+2:3d}: C={col_c} | D={col_d} | T={col_t}")
                    print(f"           {check_result}")
                
                if total_errors > 15:
                    print(f"\n⋮⋮⋮ ... Bỏ qua {total_errors - 20} lỗi ở giữa ... ⋮⋮⋮")
                    print(f"\n🔻 5 LỖI CUỐI CÙNG:")
                    for idx, row in fail_rows.tail(5).iterrows():
                        col_c = row[col_c_name] if col_c_name else 'N/A'
                        col_d = row[col_d_name] if col_d_name else 'N/A' 
                        col_t = row[col_t_name] if col_t_name else 'N/A'
                        check_result = row['Validation_Check']
                        print(f"  Dòng {idx+2:3d}: C={col_c} | D={col_d} | T={col_t}")
                        print(f"           {check_result}")
    
    def _generate_summary(self):
        """
        Tạo báo cáo tổng kết
        """
        print("=" * 80)
        print("📈 TỔNG KẾT VALIDATION")
        print("=" * 80)
        
        if self.total_rows > 0:
            pass_rate = self.total_pass / self.total_rows * 100
            fail_rate = self.total_fail / self.total_rows * 100
            
            print(f"✅ PASS: {self.total_pass:,}/{self.total_rows:,} ({pass_rate:.1f}%)")
            print(f"❌ FAIL: {self.total_fail:,}/{self.total_rows:,} ({fail_rate:.1f}%)")
            
            # Phân tích theo worksheet
            print(f"\n📊 CHI TIẾT THEO WORKSHEET:")
            for sheet_name, df in self.validation_results.items():
                sheet_total = len(df)
                sheet_pass = len(df[df['Validation_Check'] == 'PASS'])
                sheet_rate = sheet_pass / sheet_total * 100
                print(f"  {sheet_name:25s}: {sheet_pass:3d}/{sheet_total:3d} ({sheet_rate:5.1f}%)")
        else:
            print("❌ Không có dữ liệu để validation")
    
    def _export_results(self, excel_file_path):
        """
        Xuất file kết quả
        """
        if not self.validation_results:
            return None
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = Path(excel_file_path).stem
        output_file = f"validation_final_{base_name}_{timestamp}.xlsx"
        
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, df in self.validation_results.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"\n📁 File kết quả đã lưu: {output_file}")
            return output_file
            
        except Exception as e:
            print(f"❌ Lỗi xuất file: {e}")
            return None

def main():
    """
    Hàm main để chạy validation tool
    """
    try:
        # Tìm file Excel
        current_dir = Path(".")
        excel_files = [f for f in current_dir.glob("*.xlsx") 
                      if not f.name.startswith('~') 
                      and 'validation' not in f.name.lower()
                      and 'array_number' not in f.name.lower()]
        
        if not excel_files:
            print("❌ Không tìm thấy file Excel để validation!")
            return
        
        print("🔍 FILE EXCEL CÓ SẴN:")
        for i, file in enumerate(excel_files, 1):
            file_size = file.stat().st_size / 1024  # KB
            print(f"{i:2d}. {file.name:40s} ({file_size:,.0f} KB)")
        
        # Chọn file
        while True:
            try:
                choice = input(f"\n✏️ Chọn file (1-{len(excel_files)}) hoặc 'q' để thoát: ").strip()
                if choice.lower() == 'q':
                    print("👋 Đã thoát!")
                    return
                
                choice_idx = int(choice) - 1
                if 0 <= choice_idx < len(excel_files):
                    selected_file = excel_files[choice_idx]
                    break
                else:
                    print(f"❌ Vui lòng chọn số từ 1 đến {len(excel_files)}")
            except ValueError:
                print("❌ Vui lòng nhập số hợp lệ hoặc 'q'")
        
        # Chạy validation
        validator = ExcelValidator()
        output_file = validator.validate_excel_file(selected_file)
        
        if output_file:
            print(f"\n🎉 VALIDATION HOÀN THÀNH THÀNH CÔNG!")
            print(f"📁 Kết quả: {output_file}")
        else:
            print(f"\n❌ VALIDATION THẤT BẠI!")
            
    except KeyboardInterrupt:
        print("\n⏹️ Đã hủy bởi người dùng!")
    except Exception as e:
        print(f"\n❌ Lỗi không mong muốn: {e}")

if __name__ == "__main__":
    main()
