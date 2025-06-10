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

3. CP-INTERNAL Array Number Validation (mới):
   - Áp dụng tất cả worksheet
   - Quy tắc: Khi EE_System Type = "CP-INTERNAL" thì EE_Array Number phải trùng với EE_Cross Passage

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
            'Sprinkler Schedule'        ]
        
        self.pipe_treatment_worksheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule', 
            'Pipe Accessory Schedule'
        ]
          # CP-INTERNAL Array Number worksheets (Rule 3)
        self.cp_internal_worksheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule', 
            'Pipe Accessory Schedule'
        ]
        
        # Pipe Schedule Mapping worksheets (Rule 4) - MỚI
        self.pipe_schedule_mapping_worksheets = [
            'Pipe Schedule'
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
            print("3. CP-INTERNAL Array Number Validation:")
            print("   - Worksheets: Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule")
            print("   - Quy tắc: Khi EE_System Type = 'CP-INTERNAL' thì EE_Array Number phải trùng EE_Cross Passage")
            print("4. Pipe Schedule Mapping Validation (MỚI):")
            print("   - Worksheet: Pipe Schedule")
            print("   - Quy tắc: Kiểm tra mapping giữa Item Description, FAB Pipe, và Size")
            print("     * '150-900' → 'STD ARRAY TEE'")
            print("     * '65-4730' → 'STD 1 PAP RANGE'") 
            print("     * '65-5295' → 'STD 2 PAP RANGE'")
            print("     * Size 40 → 'Groove_Thread'")
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
        print(f"Số dòng: {len(df)}, Số cột: {len(df.columns)}")        # Kiểm tra worksheet nào cần validation gì
        apply_array_validation = sheet_name in self.array_number_worksheets
        apply_pipe_treatment_validation = sheet_name in self.pipe_treatment_worksheets
        apply_cp_internal_validation = sheet_name in self.cp_internal_worksheets
        apply_pipe_schedule_mapping_validation = sheet_name in self.pipe_schedule_mapping_worksheets
        
        print(f"Array Number validation: {'✅ ÁP DỤNG' if apply_array_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"Pipe Treatment validation: {'✅ ÁP DỤNG' if apply_pipe_treatment_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"CP-INTERNAL Array validation: {'✅ ÁP DỤNG' if apply_cp_internal_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"Pipe Schedule Mapping validation: {'✅ ÁP DỤNG' if apply_pipe_schedule_mapping_validation else '❌ KHÔNG ÁP DỤNG'}")
        
        if not apply_array_validation and not apply_pipe_treatment_validation and not apply_cp_internal_validation and not apply_pipe_schedule_mapping_validation:
            print("⏭️ Bỏ qua worksheet (không có quy tắc nào áp dụng)")
            print()
            return
          # Lấy tên cột theo vị trí
        col_a_name = df.columns[0] if len(df.columns) > 0 else None  # EE_Cross Passage
        col_b_name = df.columns[1] if len(df.columns) > 1 else None  # EE_Location and Lanes  
        col_c_name = df.columns[2] if len(df.columns) > 2 else None  # EE_System Type
        col_d_name = df.columns[3] if len(df.columns) > 3 else None  # EE_Array Number
        col_f_name = df.columns[5] if len(df.columns) > 5 else None  # EE_Item Description
        col_g_name = df.columns[6] if len(df.columns) > 6 else None  # EE_Size
        col_k_name = df.columns[10] if len(df.columns) > 10 else None  # EE_FAB Pipe
        col_t_name = df.columns[19] if len(df.columns) > 19 else None  # EE_Pipe Treatment        # Áp dụng validation
        df['Validation_Check'] = df.apply(
            lambda row: self._validate_row(
                row, 
                col_a_name, col_b_name, col_c_name, col_d_name, col_f_name, col_g_name, col_k_name, col_t_name,
                apply_array_validation, apply_pipe_treatment_validation, apply_cp_internal_validation, apply_pipe_schedule_mapping_validation
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
        self.total_fail += sheet_fail        # Lưu kết quả
        self.validation_results[sheet_name] = df
        
        # Hiển thị lỗi mẫu
        self._show_sample_errors(df, sheet_name, col_c_name, col_d_name, col_f_name, col_g_name, col_k_name, col_t_name)
        print()

    def _validate_row(self, row, col_a_name, col_b_name, col_c_name, col_d_name, col_f_name, col_g_name, col_k_name, col_t_name, 
                     apply_array_validation, apply_pipe_treatment_validation, apply_cp_internal_validation, apply_pipe_schedule_mapping_validation):
        """
        Validate một dòng dữ liệu với tất cả các quy tắc
        LOGIC: CP-INTERNAL ưu tiên Rule 3, Rule 4 áp dụng cho Pipe Schedule
        """
        errors = []
        
        try:
            # Kiểm tra xem có phải CP-INTERNAL không
            is_cp_internal = False
            if col_c_name and not pd.isna(row[col_c_name]):
                system_type = str(row[col_c_name]).strip()
                is_cp_internal = (system_type == "CP-INTERNAL")
            
            # Rule 1: Array Number validation (BỎ QUA nếu CP-INTERNAL)
            if apply_array_validation and col_a_name and col_b_name and col_d_name and not is_cp_internal:
                array_result = self._check_array_number(row, col_a_name, col_b_name, col_d_name)
                if array_result != "PASS" and not array_result.startswith("SKIP"):
                    errors.append(f"Array: {array_result}")
            
            # Rule 2: Pipe Treatment validation
            if apply_pipe_treatment_validation and col_c_name and col_t_name:
                treatment_result = self._check_pipe_treatment(row, col_c_name, col_t_name)
                if treatment_result != "PASS" and not treatment_result.startswith("SKIP"):
                    errors.append(f"Treatment: {treatment_result}")
                # Rule 3: CP-INTERNAL Array Number validation (ƯU TIÊN cao hơn Rule 1)
            if apply_cp_internal_validation and col_a_name and col_c_name and col_d_name:
                cp_internal_result = self._check_cp_internal_array(row, col_a_name, col_c_name, col_d_name)
                if cp_internal_result != "PASS" and not cp_internal_result.startswith("SKIP"):
                    errors.append(f"CP-Internal: {cp_internal_result}")
            
            # Rule 4: Pipe Schedule Mapping validation (MỚI)
            if apply_pipe_schedule_mapping_validation and col_f_name and col_g_name and col_k_name:
                mapping_result = self._check_pipe_schedule_mapping(row, col_f_name, col_g_name, col_k_name)
                if mapping_result != "PASS" and not mapping_result.startswith("SKIP"):
                    errors.append(f"Mapping: {mapping_result}")
            
            # Trả về kết quả
            if errors:
                return f"FAIL: {'; '.join(errors[:3])}"  # Hiển thị 3 lỗi đầu
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
            else:
                return f"'{system_type_str}' cần '{expected_treatment}', có '{pipe_treatment_str}'"
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_cp_internal_array(self, row, col_a_name, col_c_name, col_d_name):
        """
        Kiểm tra CP-INTERNAL Array Number rule:
        Khi EE_System Type = "CP-INTERNAL" thì EE_Array Number phải trùng với EE_Cross Passage
        """
        try:
            cross_passage = row[col_a_name]
            system_type = row[col_c_name]
            array_number = row[col_d_name]
            
            # Kiểm tra dữ liệu có đầy đủ không
            if pd.isna(system_type):
                return "SKIP: Thiếu System Type"
            
            system_type_str = str(system_type).strip()
            
            # Chỉ áp dụng rule cho CP-INTERNAL
            if system_type_str != "CP-INTERNAL":
                return "PASS"  # Không áp dụng rule cho system type khác
            
            # Kiểm tra dữ liệu cho CP-INTERNAL
            if pd.isna(cross_passage) or pd.isna(array_number):
                return "SKIP: Thiếu dữ liệu Cross Passage hoặc Array Number"
            
            cross_passage_str = str(cross_passage).strip()
            array_number_str = str(array_number).strip()
            
            # So sánh Cross Passage với Array Number
            if cross_passage_str == array_number_str:
                return "PASS"
            else:                return f"Array Number phải trùng Cross Passage: cần '{cross_passage_str}', có '{array_number_str}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_pipe_schedule_mapping(self, row, col_f_name, col_g_name, col_k_name):
        """
        Kiểm tra Pipe Schedule Mapping rule:
        - Item Description mapping với FAB Pipe
        - Size 40 mapping với FAB Pipe
        """
        try:
            item_description = row[col_f_name]
            size = row[col_g_name]
            fab_pipe = row[col_k_name]
            
            # Skip nếu thiếu dữ liệu
            if pd.isna(item_description) and pd.isna(size):
                return "SKIP: Thiếu Item Description và Size"
            
            errors = []
            
            # Rule 4.1: Item Description mapping
            if not pd.isna(item_description):
                item_desc_str = str(item_description).strip()
                expected_fab_pipe = None
                
                if "150-900" in item_desc_str:
                    expected_fab_pipe = "STD ARRAY TEE"
                elif "65-4730" in item_desc_str:
                    expected_fab_pipe = "STD 1 PAP RANGE"
                elif "65-5295" in item_desc_str:
                    expected_fab_pipe = "STD 2 PAP RANGE"
                
                if expected_fab_pipe:
                    if pd.isna(fab_pipe):
                        errors.append(f"Item '{item_desc_str}' cần FAB Pipe '{expected_fab_pipe}', nhưng thiếu")
                    else:
                        fab_pipe_str = str(fab_pipe).strip()
                        if fab_pipe_str != expected_fab_pipe:
                            errors.append(f"Item '{item_desc_str}' cần FAB Pipe '{expected_fab_pipe}', có '{fab_pipe_str}'")
            
            # Rule 4.2: Size 40 mapping
            if not pd.isna(size):
                size_str = str(size).strip()
                if size_str == "40":
                    expected_fab_pipe = "Groove_Thread"
                    if pd.isna(fab_pipe):
                        errors.append(f"Size 40 cần FAB Pipe '{expected_fab_pipe}', nhưng thiếu")
                    else:
                        fab_pipe_str = str(fab_pipe).strip()
                        if fab_pipe_str != expected_fab_pipe:
                            errors.append(f"Size 40 cần FAB Pipe '{expected_fab_pipe}', có '{fab_pipe_str}'")
            
            # Trả về kết quả
            if errors:
                return f"{'; '.join(errors)}"
            else:
                return "PASS"
                  except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _show_sample_errors(self, df, sheet_name, col_c_name, col_d_name, col_f_name, col_g_name, col_k_name, col_t_name):
        """
        Hiển thị TẤT CẢ lỗi với màu sắc: ĐỎ cho giá trị SAI, TRẮNG cho giá trị ĐÚNG
        """
        fail_rows = df[df['Validation_Check'] != 'PASS']
        if not fail_rows.empty:
            total_errors = len(fail_rows)
              # Hiển thị TẤT CẢ lỗi với màu sắc
            print(f"📋 TẤT CẢ {total_errors} LỖI (MÀU ĐỎ=SAI, TRẮNG=ĐÚNG):")
            for idx, row in fail_rows.iterrows():
                col_c = row[col_c_name] if col_c_name else 'N/A'
                col_d = row[col_d_name] if col_d_name else 'N/A' 
                col_f = row[col_f_name] if col_f_name else 'N/A'
                col_g = row[col_g_name] if col_g_name else 'N/A'
                col_k = row[col_k_name] if col_k_name else 'N/A'
                col_t = row[col_t_name] if col_t_name else 'N/A'
                check_result = row['Validation_Check']
                
                print(f"  Dòng {idx+2:3d}: C={col_c} | D={col_d} | F={col_f} | G={col_g} | K={col_k} | T={col_t}")
                
                # Hiển thị lỗi với màu sắc
                if "cần '" in check_result and "', có '" in check_result:
                    # Tách expected và actual values để tô màu
                    parts = check_result.split("cần '")
                    if len(parts) > 1:
                        prefix = parts[0]
                        remaining = parts[1]
                        if "', có '" in remaining:
                            expected_and_actual = remaining.split("', có '")
                            expected = expected_and_actual[0]
                            actual = expected_and_actual[1].rstrip("'")
                            
                            # In với màu: TRẮNG cho expected (đúng), ĐỎ cho actual (sai)
                            print(f"           {prefix}cần '\033[97m{expected}\033[0m', có '\033[91m{actual}\033[0m'")
                        else:
                            print(f"           {check_result}")
                    else:
                        print(f"           {check_result}")
                else:
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
