#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EXCEL VALIDATION TOOL - ENHANCED WITH END-1/END-2 RULES
=======================================================

Tool validation Excel cho dự án pipe/equipment data với 4 quy tắc ENHANCED:

1. Array Number Validation (4 worksheet):
   - Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule, Sprinkler Schedule
   - Quy tắc: Array Number phải CHỨA Cross Passage value (Fixed logic)

2. Pipe Treatment Validation (3 worksheet):
   - Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule  
   - Quy tắc:
     * CP-INTERNAL → GAL
     * CP-EXTERNAL, CW-DISTRIBUTION, CW-ARRAY → BLACK

3. CP-INTERNAL Array Number Validation (3 worksheet):
   - Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule
   - Quy tắc: Khi EE_System Type = "CP-INTERNAL" thì EE_Array Number phải trùng với EE_Cross Passage

4. Pipe Schedule Mapping Validation - ENHANCED (1 worksheet):
   - Pipe Schedule
   - Quy tắc:
     * Item Description "150-900" → FAB Pipe "STD ARRAY TEE"
     * Item Description "65-4730" → FAB Pipe "STD 1 PAP RANGE"
     * Item Description "65-5295" → FAB Pipe "STD 2 PAP RANGE"
     * Size "40" → FAB Pipe "Groove_Thread"
     * 🆕 Nếu cột (L) End-1 = "BE" hoặc cột (M) End-2 = "BE": "Fabrication"
     * 🆕 Nếu cả End-1 và End-2 đều thuộc ["RG", "TH"]: "Groove_Thread"

Tác giả: GitHub Copilot
Ngày tạo: 2025-06-10 - ENHANCED VERSION
"""

import pandas as pd
import re
import os
from pathlib import Path
from datetime import datetime

class ExcelValidatorEnhanced:
    """
    Class chính cho Excel validation với đầy đủ các quy tắc ENHANCED
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
        
        # CP-INTERNAL Array Number worksheets (Rule 3)
        self.cp_internal_worksheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule', 
            'Pipe Accessory Schedule'
        ]
        
        # Pipe Schedule Mapping worksheets (Rule 4) - ENHANCED
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
        Validate toàn bộ file Excel với tất cả các quy tắc ENHANCED
        """
        try:
            print("=" * 80)
            print("🚀 EXCEL VALIDATION TOOL - 4 RULES ENHANCED")
            print("=" * 80)
            print(f"📁 File: {excel_file_path}")
            print(f"🕐 Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print()
            
            # Đọc Excel file
            xl_file = pd.ExcelFile(excel_file_path)
            
            print("📋 CÁC QUY TẮC VALIDATION:")
            print("1. Array Number Validation:")
            print(f"   - Worksheets: {', '.join(self.array_number_worksheets)}")
            print("   - Quy tắc: Array Number phải CHỨA Cross Passage value")
            print("2. Pipe Treatment Validation:")
            print(f"   - Worksheets: {', '.join(self.pipe_treatment_worksheets)}")
            print("   - Quy tắc: CP-INTERNAL→GAL, CP-EXTERNAL/CW-DISTRIBUTION/CW-ARRAY→BLACK")
            print("3. CP-INTERNAL Array Number Validation:")
            print("   - Worksheets: Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule")
            print("   - Quy tắc: Khi EE_System Type = 'CP-INTERNAL' thì EE_Array Number phải trùng EE_Cross Passage")
            print("4. Pipe Schedule Mapping Validation - 🆕 ENHANCED:")
            print("   - Worksheet: Pipe Schedule")
            print("   - Quy tắc cũ: Item Description/Size mapping với FAB Pipe")
            print("   - 🆕 MỚI: End-1/End-2 = 'BE' → 'Fabrication'")
            print("   - 🆕 MỚI: End-1 & End-2 thuộc ['RG','TH'] → 'Groove_Thread'")
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
        Validate một worksheet cụ thể với tất cả các quy tắc ENHANCED
        """
        print(f"📊 WORKSHEET: {sheet_name}")
        print("-" * 50)
        
        # Đọc worksheet
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        print(f"Số dòng: {len(df)}, Số cột: {len(df.columns)}")
        
        # Kiểm tra worksheet nào cần validation gì
        apply_array_validation = sheet_name in self.array_number_worksheets
        apply_pipe_treatment_validation = sheet_name in self.pipe_treatment_worksheets
        apply_cp_internal_validation = sheet_name in self.cp_internal_worksheets
        apply_pipe_schedule_mapping_validation = sheet_name in self.pipe_schedule_mapping_worksheets
        
        print(f"Array Number validation: {'✅ ÁP DỤNG' if apply_array_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"Pipe Treatment validation: {'✅ ÁP DỤNG' if apply_pipe_treatment_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"CP-INTERNAL Array validation: {'✅ ÁP DỤNG' if apply_cp_internal_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"Pipe Schedule Mapping validation: {'✅ ÁP DỤNG - ENHANCED' if apply_pipe_schedule_mapping_validation else '❌ KHÔNG ÁP DỤNG'}")
        
        if not any([apply_array_validation, apply_pipe_treatment_validation, apply_cp_internal_validation, apply_pipe_schedule_mapping_validation]):
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
        col_l_name = df.columns[11] if len(df.columns) > 11 else None  # EE_End-1
        col_m_name = df.columns[12] if len(df.columns) > 12 else None  # EE_End-2
        col_t_name = df.columns[19] if len(df.columns) > 19 else None  # EE_Pipe Treatment
        
        # Áp dụng validation
        df['Validation_Check'] = df.apply(
            lambda row: self._validate_row(
                row, 
                col_a_name, col_b_name, col_c_name, col_d_name, col_f_name, col_g_name, col_k_name, col_l_name, col_m_name, col_t_name,
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
        self.total_fail += sheet_fail
        
        # Lưu kết quả
        self.validation_results[sheet_name] = df
        
        # Hiển thị lỗi mẫu
        self._show_sample_errors(df, sheet_name, col_c_name, col_d_name, col_f_name, col_g_name, col_k_name, col_l_name, col_m_name, col_t_name)
        print()

    def _validate_row(self, row, col_a_name, col_b_name, col_c_name, col_d_name, col_f_name, col_g_name, col_k_name, col_l_name, col_m_name, col_t_name, 
                     apply_array_validation, apply_pipe_treatment_validation, apply_cp_internal_validation, apply_pipe_schedule_mapping_validation):
        """
        Validate một dòng dữ liệu với tất cả các quy tắc ENHANCED
        LOGIC: CP-INTERNAL ưu tiên Rule 3 thay vì Rule 1
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
            
            # Rule 4: Pipe Schedule Mapping validation - ENHANCED
            if apply_pipe_schedule_mapping_validation and col_f_name and col_g_name and col_k_name:
                mapping_result = self._check_pipe_schedule_mapping_enhanced(row, col_f_name, col_g_name, col_k_name, col_l_name, col_m_name)
                if mapping_result != "PASS" and not mapping_result.startswith("SKIP"):
                    errors.append(f"Mapping: {mapping_result}")
            
            # Trả về kết quả
            if errors:
                return f"FAIL: {'; '.join(errors[:4])}"  # Hiển thị 4 lỗi đầu
            else:
                return "PASS"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_array_number(self, row, col_a_name, col_b_name, col_d_name):
        """
        Rule 1: Kiểm tra Array Number format
        UPDATED LOGIC: Array Number should contain Cross Passage value
        """
        try:
            cross_passage = row[col_a_name]
            location_lanes = row[col_b_name]
            array_number = row[col_d_name]
            
            if pd.isna(cross_passage) or pd.isna(location_lanes) or pd.isna(array_number):
                return "SKIP: Thiếu dữ liệu Cross Passage hoặc Array Number"
            
            # Chuyển thành string
            cross_passage_str = str(cross_passage).strip()
            actual_array = str(array_number).strip()
            
            # UPDATED LOGIC: Check if array number contains cross passage
            if cross_passage_str in actual_array:
                return "PASS"
            else:
                return f"Array Number '{actual_array}' cần chứa '{cross_passage_str}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"

    def _check_pipe_treatment(self, row, col_c_name, col_t_name):
        """
        Rule 2: Kiểm tra Pipe Treatment
        """
        try:
            system_type = row[col_c_name]
            pipe_treatment = row[col_t_name]
            
            if pd.isna(system_type) or pd.isna(pipe_treatment):
                return "SKIP: Thiếu dữ liệu System Type hoặc Pipe Treatment"
            
            system_type_str = str(system_type).strip()
            pipe_treatment_str = str(pipe_treatment).strip()
            
            # Quy tắc validation
            if system_type_str == "CP-INTERNAL":
                expected = "GAL"
                if pipe_treatment_str == expected:
                    return "PASS"
                else:
                    return f"CP-INTERNAL cần '{expected}', nhận '{pipe_treatment_str}'"
            
            elif system_type_str in ["CP-EXTERNAL", "CW-DISTRIBUTION", "CW-ARRAY"]:
                expected = "BLACK"
                if pipe_treatment_str == expected:
                    return "PASS"
                else:
                    return f"{system_type_str} cần '{expected}', nhận '{pipe_treatment_str}'"
            
            else:
                return "SKIP: System Type không thuộc quy tắc"
                
        except Exception as e:
            return f"ERROR: {str(e)}"

    def _check_cp_internal_array(self, row, col_a_name, col_c_name, col_d_name):
        """
        Rule 3: Kiểm tra CP-INTERNAL Array Number matching
        """
        try:
            cross_passage = row[col_a_name]
            system_type = row[col_c_name]
            array_number = row[col_d_name]
            
            if pd.isna(system_type):
                return "SKIP: Thiếu System Type"
            
            system_type_str = str(system_type).strip()
            
            # Chỉ áp dụng với CP-INTERNAL
            if system_type_str != "CP-INTERNAL":
                return "SKIP: Không phải CP-INTERNAL"
            
            if pd.isna(cross_passage) or pd.isna(array_number):
                return "SKIP: Thiếu Cross Passage hoặc Array Number"
            
            cross_passage_str = str(cross_passage).strip()
            array_number_str = str(array_number).strip()
            
            if cross_passage_str == array_number_str:
                return "PASS"
            else:
                return f"CP-INTERNAL: Array Number '{array_number_str}' phải trùng Cross Passage '{cross_passage_str}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"

    def _check_pipe_schedule_mapping_enhanced(self, row, col_f_name, col_g_name, col_k_name, col_l_name, col_m_name):
        """
        Rule 4: Kiểm tra Pipe Schedule Mapping - ENHANCED với End-1/End-2
        """
        try:
            item_description = row[col_f_name] if col_f_name else None
            size = row[col_g_name] if col_g_name else None
            fab_pipe = row[col_k_name] if col_k_name else None
            end_1 = row[col_l_name] if col_l_name else None
            end_2 = row[col_m_name] if col_m_name else None
            
            if pd.isna(fab_pipe):
                return "SKIP: Thiếu FAB Pipe"
            
            fab_pipe_str = str(fab_pipe).strip()
            
            # Check End-1/End-2 rules FIRST (NEW ENHANCED LOGIC)
            
            # Rule 4.3: If End-1 = "BE" OR End-2 = "BE" → FAB Pipe should be "Fabrication"
            end_1_str = str(end_1).strip() if not pd.isna(end_1) else ""
            end_2_str = str(end_2).strip() if not pd.isna(end_2) else ""
            
            if end_1_str == "BE" or end_2_str == "BE":
                if fab_pipe_str == "Fabrication":
                    return "PASS"
                else:
                    return f"End-1/End-2 = 'BE' cần FAB Pipe 'Fabrication', nhận '{fab_pipe_str}'"
            
            # Rule 4.4: If BOTH End-1 AND End-2 are in ["RG", "TH"] → FAB Pipe should be "Groove_Thread"  
            if (end_1_str in ["RG", "TH"]) and (end_2_str in ["RG", "TH"]):
                if fab_pipe_str == "Groove_Thread":
                    return "PASS"
                else:
                    return f"End-1 & End-2 thuộc ['RG','TH'] cần FAB Pipe 'Groove_Thread', nhận '{fab_pipe_str}'"
            
            # Original mapping rules (if End-1/End-2 rules don't apply)
            
            # Rule 4.1: Item Description mapping
            if not pd.isna(item_description):
                item_description_str = str(item_description).strip()
                
                if "150-900" in item_description_str:
                    expected = "STD ARRAY TEE"
                    if fab_pipe_str == expected:
                        return "PASS"
                    else:
                        return f"Item '150-900' cần FAB Pipe '{expected}', nhận '{fab_pipe_str}'"
                
                elif "65-4730" in item_description_str:
                    expected = "STD 1 PAP RANGE"
                    if fab_pipe_str == expected:
                        return "PASS"
                    else:
                        return f"Item '65-4730' cần FAB Pipe '{expected}', nhận '{fab_pipe_str}'"
                
                elif "65-5295" in item_description_str:
                    expected = "STD 2 PAP RANGE"
                    if fab_pipe_str == expected:
                        return "PASS"
                    else:
                        return f"Item '65-5295' cần FAB Pipe '{expected}', nhận '{fab_pipe_str}'"
            
            # Rule 4.2: Size mapping
            if not pd.isna(size):
                size_str = str(size).strip()
                if size_str == "40":
                    expected = "Groove_Thread"
                    if fab_pipe_str == expected:
                        return "PASS"
                    else:
                        return f"Size '40' cần FAB Pipe '{expected}', nhận '{fab_pipe_str}'"
            
            return "SKIP: Không thuộc quy tắc mapping"
            
        except Exception as e:
            return f"ERROR: {str(e)}"    def _show_sample_errors(self, df, sheet_name, col_c_name, col_d_name, col_f_name, col_g_name, col_k_name, col_l_name, col_m_name, col_t_name):
        """
        Hiển thị TẤT CẢ lỗi với thông tin vị trí chi tiết (Worksheet + Row + Column)
        """
        fail_df = df[df['Validation_Check'] != 'PASS']
        
        if len(fail_df) == 0:
            print("✨ Không có lỗi nào!")
            return
        
        print(f"⚠️  HIỂN THỊ TẤT CẢ {len(fail_df)} LỖI CHO WORKSHEET '{sheet_name}':")
        print()
        
        # Chọn cột hiển thị
        display_cols = []
        if col_c_name: display_cols.append(col_c_name)
        if col_d_name: display_cols.append(col_d_name)  
        if col_f_name: display_cols.append(col_f_name)
        if col_g_name: display_cols.append(col_g_name)
        if col_k_name: display_cols.append(col_k_name)
        if col_l_name: display_cols.append(col_l_name)  # End-1
        if col_m_name: display_cols.append(col_m_name)  # End-2
        if col_t_name: display_cols.append(col_t_name)
        display_cols.append('Validation_Check')
          for idx, (original_idx, row) in enumerate(fail_df.iterrows(), 1):
            # Excel row number = pandas index + 2 (header = row 1, data starts from row 2)
            excel_row = original_idx + 2
            print(f"🔸 Lỗi {idx} - WORKSHEET: '{sheet_name}' - HÀNG {excel_row} (Excel):")
            
            # Map column names to Excel column letters for easier identification
            col_mapping = {
                col_c_name: 'C (System Type)',
                col_d_name: 'D (Array Number)', 
                col_f_name: 'F (Item Description)',
                col_g_name: 'G (Size)',
                col_k_name: 'K (FAB Pipe)',
                col_l_name: 'L (End-1)',
                col_m_name: 'M (End-2)',
                col_t_name: 'T (Pipe Treatment)',
                'Validation_Check': 'Validation_Check'
            }
            
            for col in display_cols:
                if col in row:
                    value = row[col]
                    col_display = col_mapping.get(col, col) if col in col_mapping else col
                    if col == 'Validation_Check' and str(value).startswith('FAIL'):
                        print(f"   🔴 {col_display}: {value}")
                    else:
                        print(f"   ⚪ {col_display}: {value}")
            print()    def _generate_summary(self):
        """
        Tạo báo cáo tổng kết ENHANCED với hướng dẫn sửa lỗi
        """
        print("=" * 80)
        print("📊 TỔNG KẾT VALIDATION - ENHANCED VERSION")
        print("=" * 80)
        print(f"Tổng số dòng đã kiểm tra: {self.total_rows}")
        print(f"✅ PASS: {self.total_pass} ({self.total_pass/self.total_rows*100:.1f}%)")
        print(f"❌ FAIL: {self.total_fail} ({self.total_fail/self.total_rows*100:.1f}%)")
        print()
        
        # Chi tiết theo worksheet
        for sheet_name, df in self.validation_results.items():
            total = len(df)
            passes = len(df[df['Validation_Check'] == 'PASS'])
            fails = len(df[df['Validation_Check'] != 'PASS'])
            print(f"📋 {sheet_name}: {passes}/{total} PASS ({passes/total*100:.1f}%)")
        
        # Hướng dẫn sửa lỗi
        if self.total_fail > 0:
            print()
            print("🔧 HƯỚNG DẪN SỬA LỖI:")
            print("   1. Xem thông tin lỗi ở trên: WORKSHEET + HÀNG + CỘT")
            print("   2. Mở file Excel gốc")
            print("   3. Đi đến Worksheet được chỉ định")
            print("   4. Đi đến Hàng được chỉ định (ví dụ: HÀNG 218)")
            print("   5. Sửa giá trị trong cột được chỉ định (A, C, D, F, G, K, L, M, T)")
            print("   6. Lưu file và chạy lại validation")
            print()
            print("📝 VÍ DỤ SỬA LỖI:")
            print("   🔸 Lỗi: WORKSHEET 'Pipe Schedule' - HÀNG 218")
            print("   🔴 L (End-1): RG, M (End-2): RG → cần FAB Pipe 'Groove_Thread'")
            print("   → Sửa: Đi đến Pipe Schedule, hàng 218, cột K, đổi thành 'Groove_Thread'")
        print()

    def _export_results(self, excel_file_path):
        """
        Xuất kết quả validation ra file Excel
        """
        try:
            # Tạo tên file output
            input_path = Path(excel_file_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"validation_enhanced_{input_path.stem}_{timestamp}.xlsx"
            output_path = input_path.parent / output_filename
            
            print(f"💾 Xuất kết quả ra: {output_filename}")
            
            # Xuất từng sheet
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in self.validation_results.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print("✅ Xuất file thành công!")
            return str(output_path)
            
        except Exception as e:
            print(f"❌ Lỗi xuất file: {e}")
            return None

def main():
    """
    Hàm main để chạy validation tool ENHANCED
    """
    print("🔍 Tìm file Excel...")
    
    # Tìm file Excel trong thư mục hiện tại
    current_dir = Path(".")
    excel_files = list(current_dir.glob("*.xlsx"))
    
    if not excel_files:
        print("❌ Không tìm thấy file Excel (.xlsx) nào!")
        input("Nhấn Enter để thoát...")
        return
    
    if len(excel_files) == 1:
        selected_file = excel_files[0]
        print(f"📁 Tìm thấy 1 file: {selected_file.name}")
    else:
        print(f"📁 Tìm thấy {len(excel_files)} file Excel:")
        for i, file in enumerate(excel_files, 1):
            print(f"   {i}. {file.name}")
        
        while True:
            try:
                choice = int(input("\nChọn file (nhập số): ")) - 1
                if 0 <= choice < len(excel_files):
                    selected_file = excel_files[choice]
                    break
                else:
                    print("❌ Lựa chọn không hợp lệ!")
            except ValueError:
                print("❌ Vui lòng nhập số!")
    
    # Chạy validation ENHANCED
    validator = ExcelValidatorEnhanced()
    output_file = validator.validate_excel_file(str(selected_file))
    
    if output_file:
        print("🎉 VALIDATION ENHANCED HOÀN THÀNH!")
        print(f"📄 Kết quả đã được lưu tại: {output_file}")
    else:
        print("❌ Validation thất bại!")
    
    input("\nNhấn Enter để thoát...")

if __name__ == "__main__":
    main()
