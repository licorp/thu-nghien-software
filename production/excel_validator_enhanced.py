#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import re
import os
from pathlib import Path
from datetime import datetime

class ExcelValidatorEnhanced:
    """
    Excel Validator với validation mở rộng cho FAB Pipe, Pap 1, Pap 2
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
        
        # Cấu hình validation cho các cột mới
        self.fab_pipe_worksheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule'
        ]
        
        self.pap_validation_worksheets = [
            'Pipe Schedule'
        ]
        
        # Thống kê validation chi tiết
        self.total_rows = 0
        self.array_pass = 0
        self.array_fail = 0
        self.array_skip = 0
        self.treatment_pass = 0
        self.treatment_fail = 0
        self.treatment_skip = 0
        self.fab_pipe_pass = 0
        self.fab_pipe_fail = 0
        self.fab_pipe_skip = 0
        self.pap1_pass = 0
        self.pap1_fail = 0
        self.pap1_skip = 0
        self.pap2_pass = 0
        self.pap2_fail = 0
        self.pap2_skip = 0
        
        self.validation_results = {}
        
    def validate_excel_file(self, excel_file_path):
        """
        Validate toàn bộ Excel file với chi tiết từng rule
        """
        try:
            print("=" * 80)
            print("🚀 EXCEL VALIDATION TOOL - ENHANCED VERSION")
            print("=" * 80)
            print(f"📁 File: {Path(excel_file_path).name}")
            print(f"🕐 Thời gian: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            # Hiển thị cấu hình            
            print(f"\n📋 CẤU HÌNH VALIDATION:")
            print(f"1. Array Number Validation:")
            print(f"   - Worksheets: {', '.join(self.array_number_worksheets)}")
            print(f"   - Quy tắc 1: Khi System Type = CP-INTERNAL → Array Number = Cross Passage")
            print(f"   - Quy tắc 2: Các trường hợp khác → Array Number chứa 'EXP6' + 2 số cuối cột B + 2 số cuối cột A")
            print(f"2. Pipe Treatment Validation:")
            print(f"   - Worksheets: {', '.join(self.pipe_treatment_worksheets)}")
            print(f"   - Quy tắc: CP-INTERNAL→GAL, CP-EXTERNAL/CW-DISTRIBUTION/CW-ARRAY→BLACK")
            print(f"3. FAB Pipe Validation:")
            print(f"   - Worksheets: {', '.join(self.fab_pipe_worksheets)}")
            print(f"   - Quy tắc: Dựa trên Item Description, Size, End-1/End-2")
            print(f"4. Pap 1 & Pap 2 Validation:")
            print(f"   - Worksheets: {', '.join(self.pap_validation_worksheets)}")
            print(f"   - Quy tắc: Pap 1 theo Item Description, Pap 2 cho ống 65mm dài 5295mm")
            
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
        apply_fab_pipe_validation = sheet_name in self.fab_pipe_worksheets
        apply_pap_validation = sheet_name in self.pap_validation_worksheets
        
        print(f"🔢 Array Number validation: {'✅ ÁP DỤNG' if apply_array_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"🔧 Pipe Treatment validation: {'✅ ÁP DỤNG' if apply_pipe_treatment_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"🏭 FAB Pipe validation: {'✅ ÁP DỤNG' if apply_fab_pipe_validation else '❌ KHÔNG ÁP DỤNG'}")
        print(f"📏 Pap validation: {'✅ ÁP DỤNG' if apply_pap_validation else '❌ KHÔNG ÁP DỤNG'}")
        
        if not any([apply_array_validation, apply_pipe_treatment_validation, 
                   apply_fab_pipe_validation, apply_pap_validation]):
            print("⏭️ Bỏ qua worksheet (không có quy tắc nào áp dụng)")
            return
        
        # Lấy tên cột theo vị trí
        col_a_name = df.columns[0] if len(df.columns) > 0 else None  # EE_Cross Passage
        col_b_name = df.columns[1] if len(df.columns) > 1 else None  # EE_Location and Lanes  
        col_c_name = df.columns[2] if len(df.columns) > 2 else None  # EE_System Type
        col_d_name = df.columns[3] if len(df.columns) > 3 else None  # EE_Array Number
        col_k_name = df.columns[10] if len(df.columns) > 10 else None  # EE_FAB Pipe
        col_o_name = df.columns[14] if len(df.columns) > 14 else None  # EE_Pap 1
        col_p_name = df.columns[15] if len(df.columns) > 15 else None  # EE_Pap 2
        col_t_name = df.columns[19] if len(df.columns) > 19 else None  # EE_Pipe Treatment
        
        # Thêm cột để dễ truy cập
        item_desc_col = None
        size_col = None
        length_col = None
        end1_col = None
        end2_col = None
        
        for i, col in enumerate(df.columns):
            if 'Item Description' in str(col):
                item_desc_col = df.columns[i]
            elif 'Size' in str(col):
                size_col = df.columns[i]
            elif 'Length' in str(col):
                length_col = df.columns[i]
            elif 'END-1' in str(col):
                end1_col = df.columns[i]
            elif 'END-2' in str(col):
                end2_col = df.columns[i]
        
        # Áp dụng validation chi tiết
        array_results = []
        treatment_results = []
        fab_pipe_results = []
        pap1_results = []
        pap2_results = []
        
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
                
            # FAB Pipe validation
            if apply_fab_pipe_validation:
                fab_result = self._check_fab_pipe_detailed(row, item_desc_col, size_col, end1_col, end2_col, col_k_name)
                fab_pipe_results.append(fab_result)
            else:
                fab_pipe_results.append("N/A")
                
            # Pap 1 validation
            if apply_pap_validation:
                pap1_result = self._check_pap1_detailed(row, item_desc_col, col_o_name)
                pap1_results.append(pap1_result)
            else:
                pap1_results.append("N/A")
                
            # Pap 2 validation
            if apply_pap_validation:
                pap2_result = self._check_pap2_detailed(row, size_col, length_col, col_p_name)
                pap2_results.append(pap2_result)
            else:
                pap2_results.append("N/A")
        
        # Thêm kết quả vào DataFrame
        df['Array_Check'] = array_results
        df['Treatment_Check'] = treatment_results
        df['FAB_Pipe_Check'] = fab_pipe_results
        df['Pap1_Check'] = pap1_results
        df['Pap2_Check'] = pap2_results
        
        # Thống kê chi tiết
        self._report_detailed_stats(df, sheet_name, apply_array_validation, apply_pipe_treatment_validation,
                                  apply_fab_pipe_validation, apply_pap_validation)
        
        # Hiển thị lỗi mẫu
        self._show_detailed_errors(df, sheet_name, col_c_name, col_d_name, col_t_name, col_k_name, col_o_name, col_p_name,
                                 apply_array_validation, apply_pipe_treatment_validation, 
                                 apply_fab_pipe_validation, apply_pap_validation)
        
        # Lưu kết quả
        self.validation_results[sheet_name] = df
        self.total_rows += len(df)
    
    def _check_array_number_detailed(self, row, col_a_name, col_b_name, col_d_name):
        """
        Kiểm tra Array Number rule chi tiết với 2 rules:
        1. Rule mới: Khi EE_System Type = CP-INTERNAL thì EE_Array Number = EE_Cross Passage
        2. Rule cũ: Các trường hợp khác → EE_Array Number = "EXP6" + 2 số cuối cột B + 2 số cuối cột A
        """
        try:
            if not col_a_name or not col_b_name or not col_d_name:
                return "SKIP: Thiếu cột"
                
            cross_passage = row[col_a_name]
            location_lanes = row[col_b_name]
            array_number = row[col_d_name]
            
            # Lấy System Type (cột thứ 3, index 2)
            system_type = row.iloc[2] if len(row) > 2 else None
            
            if pd.isna(cross_passage) or pd.isna(location_lanes) or pd.isna(array_number):
                return "SKIP: Thiếu dữ liệu"
            
            actual_array = str(array_number).strip()
            cross_passage_str = str(cross_passage).strip()
            
            # RULE MỚI: Chỉ kiểm tra CP-INTERNAL
            if pd.notna(system_type):
                system_type_str = str(system_type).upper().strip()
                if system_type_str == 'CP-INTERNAL':
                    # Rule mới: Array Number phải bằng Cross Passage
                    if actual_array == cross_passage_str:
                        return "PASS (Rule: CP-INTERNAL)"
                    else:
                        return f"FAIL (Rule: CP-INTERNAL): cần Array=Cross, có '{actual_array}' ≠ '{cross_passage_str}'"
            
            # RULE CŨ: Pattern EXP6 + digits (cho tất cả các case khác)
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
            
            if required_pattern in actual_array:
                return "PASS (Rule: Pattern)"
            else:
                return f"FAIL (Rule: Pattern): cần '{required_pattern}', có '{actual_array}'"
                
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
    
    def _check_fab_pipe_detailed(self, row, item_desc_col, size_col, end1_col, end2_col, col_k_name):
        """
        Kiểm tra EE_FAB Pipe (cột K) dựa trên Item Description, Size, End-1, End-2
        """
        try:
            if not col_k_name or not item_desc_col:
                return "SKIP: Thiếu cột"
                
            fab_pipe = row[col_k_name] if col_k_name else None
            item_desc = row[item_desc_col] if item_desc_col else None
            size = row[size_col] if size_col else None
            end1 = row[end1_col] if end1_col else None
            end2 = row[end2_col] if end2_col else None
            
            if pd.isna(fab_pipe):
                return "SKIP: FAB Pipe trống"
            
            if pd.isna(item_desc):
                return "SKIP: Item Description trống"
                
            fab_pipe_str = str(fab_pipe).strip()
            item_desc_str = str(item_desc).strip()
            size_str = str(size).strip() if pd.notna(size) else ""
            end1_str = str(end1).strip() if pd.notna(end1) else ""
            end2_str = str(end2).strip() if pd.notna(end2) else ""
            
            # Quy tắc FAB Pipe theo Item Description
            expected_fab = None
            
            # Rule 1: Nếu Item Description chứa "Groove" → FAB Pipe = "Groove_Thread"
            if "Groove" in item_desc_str:
                expected_fab = "Groove_Thread"
            # Rule 2: Nếu Item Description chứa "Thread" → FAB Pipe = "Thread"  
            elif "Thread" in item_desc_str:
                expected_fab = "Thread"
            # Rule 3: Nếu Item Description chứa "Flange" → FAB Pipe = "Flange"
            elif "Flange" in item_desc_str:
                expected_fab = "Flange"
            # Rule 4: Nếu Item Description chứa "Coupling" → FAB Pipe = "Coupling"
            elif "Coupling" in item_desc_str:
                expected_fab = "Coupling"
            # Rule 5: Dựa trên Size và End-1/End-2
            elif size_str and (end1_str or end2_str):
                # Nếu size >= 100 và có End-1 hoặc End-2 → "Groove_Thread"
                try:
                    size_num = float(size_str)
                    if size_num >= 100 and (end1_str or end2_str):
                        expected_fab = "Groove_Thread"
                    else:
                        expected_fab = "Thread"
                except:
                    expected_fab = "Thread"
            else:
                return "SKIP: Không đủ thông tin để xác định FAB Pipe"
            
            if fab_pipe_str == expected_fab:
                return "PASS"
            else:
                return f"FAIL: cần '{expected_fab}', có '{fab_pipe_str}' (Item: {item_desc_str})"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_pap1_detailed(self, row, item_desc_col, col_o_name):
        """
        Kiểm tra EE_Pap 1 (cột O) dựa trên Item Description
        """
        try:
            if not col_o_name or not item_desc_col:
                return "SKIP: Thiếu cột"
                
            pap1 = row[col_o_name] if col_o_name else None
            item_desc = row[item_desc_col] if item_desc_col else None
            
            if pd.isna(item_desc):
                return "SKIP: Item Description trống"
                
            item_desc_str = str(item_desc).strip()
            pap1_str = str(pap1).strip() if pd.notna(pap1) else ""
            
            # Quy tắc Pap 1 theo Item Description
            expected_pap1 = None
            
            # Rule mapping cho Pap 1
            if "90° Elbow" in item_desc_str:
                expected_pap1 = "90_Elbow"
            elif "45° Elbow" in item_desc_str:
                expected_pap1 = "45_Elbow"
            elif "Tee" in item_desc_str:
                expected_pap1 = "Tee"
            elif "Cross" in item_desc_str:
                expected_pap1 = "Cross"
            elif "Reducer" in item_desc_str:
                expected_pap1 = "Reducer"
            elif "Cap" in item_desc_str:
                expected_pap1 = "Cap"
            elif "Coupling" in item_desc_str:
                expected_pap1 = "Coupling"
            elif "Flange" in item_desc_str:
                expected_pap1 = "Flange"
            elif "Pipe" in item_desc_str and "Schedule" not in item_desc_str:
                expected_pap1 = "Straight_Pipe"
            else:
                return "SKIP: Item Description không khớp rule Pap 1"
            
            # Nếu không có giá trị Pap 1 nhưng có expected → FAIL
            if not pap1_str and expected_pap1:
                return f"FAIL: cần '{expected_pap1}', có rỗng (Item: {item_desc_str})"
            
            # Nếu không có expected → SKIP
            if not expected_pap1:
                return "SKIP: Không áp dụng rule"
                
            if pap1_str == expected_pap1:
                return "PASS"
            else:
                return f"FAIL: cần '{expected_pap1}', có '{pap1_str}' (Item: {item_desc_str})"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_pap2_detailed(self, row, size_col, length_col, col_p_name):
        """
        Kiểm tra EE_Pap 2 (cột P) - đặc biệt cho ống 65mm dài 5295mm
        """
        try:
            if not col_p_name:
                return "SKIP: Thiếu cột"
                
            pap2 = row[col_p_name] if col_p_name else None
            size = row[size_col] if size_col else None
            length = row[length_col] if length_col else None
            
            size_str = str(size).strip() if pd.notna(size) else ""
            length_str = str(length).strip() if pd.notna(length) else ""
            pap2_str = str(pap2).strip() if pd.notna(pap2) else ""
            
            # Rule đặc biệt: ống 65mm dài 5295mm
            try:
                size_num = float(size_str) if size_str else 0
                length_num = float(length_str) if length_str else 0
                
                # Kiểm tra điều kiện đặc biệt
                if size_num == 65.0 and length_num == 5295.0:
                    expected_pap2 = "Special_65mm_5295"
                    if pap2_str == expected_pap2:
                        return "PASS"
                    else:
                        return f"FAIL: ống 65mm dài 5295mm cần '{expected_pap2}', có '{pap2_str}'"
                else:
                    # Các case khác có thể để trống
                    return "SKIP: Không phải ống 65mm dài 5295mm"
                    
            except ValueError:
                return "SKIP: Size hoặc Length không phải số"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _report_detailed_stats(self, df, sheet_name, apply_array, apply_treatment, apply_fab_pipe, apply_pap):
        """
        Báo cáo thống kê chi tiết cho worksheet
        """
        # Thống kê Array Number
        if apply_array:
            array_pass = len(df[df['Array_Check'].str.contains('PASS', na=False)])
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
            treatment_pass = len(df[df['Treatment_Check'].str.contains('PASS', na=False)])
            treatment_fail = len(df[df['Treatment_Check'].str.startswith('FAIL', na=False)])
            treatment_skip = len(df[df['Treatment_Check'].str.startswith('SKIP', na=False)])
            
            print(f"\n🔧 PIPE TREATMENT VALIDATION:")
            print(f"   ✅ PASS: {treatment_pass}")
            print(f"   ❌ FAIL: {treatment_fail}")
            print(f"   ⏭️ SKIP: {treatment_skip}")
            
            self.treatment_pass += treatment_pass
            self.treatment_fail += treatment_fail
            self.treatment_skip += treatment_skip
            
        # Thống kê FAB Pipe
        if apply_fab_pipe:
            fab_pass = len(df[df['FAB_Pipe_Check'].str.contains('PASS', na=False)])
            fab_fail = len(df[df['FAB_Pipe_Check'].str.startswith('FAIL', na=False)])
            fab_skip = len(df[df['FAB_Pipe_Check'].str.startswith('SKIP', na=False)])
            
            print(f"\n🏭 FAB PIPE VALIDATION:")
            print(f"   ✅ PASS: {fab_pass}")
            print(f"   ❌ FAIL: {fab_fail}")
            print(f"   ⏭️ SKIP: {fab_skip}")
            
            self.fab_pipe_pass += fab_pass
            self.fab_pipe_fail += fab_fail
            self.fab_pipe_skip += fab_skip
            
        # Thống kê Pap 1 & 2
        if apply_pap:
            pap1_pass = len(df[df['Pap1_Check'].str.contains('PASS', na=False)])
            pap1_fail = len(df[df['Pap1_Check'].str.startswith('FAIL', na=False)])
            pap1_skip = len(df[df['Pap1_Check'].str.startswith('SKIP', na=False)])
            
            pap2_pass = len(df[df['Pap2_Check'].str.contains('PASS', na=False)])
            pap2_fail = len(df[df['Pap2_Check'].str.startswith('FAIL', na=False)])
            pap2_skip = len(df[df['Pap2_Check'].str.startswith('SKIP', na=False)])
            
            print(f"\n📏 PAP 1 VALIDATION:")
            print(f"   ✅ PASS: {pap1_pass}")
            print(f"   ❌ FAIL: {pap1_fail}")
            print(f"   ⏭️ SKIP: {pap1_skip}")
            
            print(f"\n📏 PAP 2 VALIDATION:")
            print(f"   ✅ PASS: {pap2_pass}")
            print(f"   ❌ FAIL: {pap2_fail}")
            print(f"   ⏭️ SKIP: {pap2_skip}")
            
            self.pap1_pass += pap1_pass
            self.pap1_fail += pap1_fail
            self.pap1_skip += pap1_skip
            self.pap2_pass += pap2_pass
            self.pap2_fail += pap2_fail
            self.pap2_skip += pap2_skip
    
    def _show_detailed_errors(self, df, sheet_name, col_c_name, col_d_name, col_t_name, col_k_name, col_o_name, col_p_name, 
                            apply_array, apply_treatment, apply_fab_pipe, apply_pap):
        """
        Hiển thị lỗi chi tiết theo từng loại validation
        """        
        # Lỗi Array Number
        if apply_array:
            array_errors = df[df['Array_Check'].str.startswith('FAIL', na=False)]
            if not array_errors.empty:
                print(f"\n❌ LỖI ARRAY NUMBER ({len(array_errors)} lỗi):")
                for idx, row in array_errors.iterrows():
                    col_d = row[col_d_name] if col_d_name else 'N/A'
                    error_msg = row['Array_Check']
                    print(f"   Dòng {idx+2:3d}: D={col_d} | {error_msg}")
        
        # Lỗi Pipe Treatment  
        if apply_treatment:
            treatment_errors = df[df['Treatment_Check'].str.startswith('FAIL', na=False)]
            if not treatment_errors.empty:
                print(f"\n❌ LỖI PIPE TREATMENT ({len(treatment_errors)} lỗi):")
                for idx, row in treatment_errors.iterrows():
                    col_c = row[col_c_name] if col_c_name else 'N/A'
                    col_t = row[col_t_name] if col_t_name else 'N/A'
                    error_msg = row['Treatment_Check']
                    print(f"   Dòng {idx+2:3d}: C={col_c} | T={col_t} | {error_msg}")
                    
        # Lỗi FAB Pipe
        if apply_fab_pipe:
            fab_errors = df[df['FAB_Pipe_Check'].str.startswith('FAIL', na=False)]
            if not fab_errors.empty:
                print(f"\n❌ LỖI FAB PIPE ({len(fab_errors)} lỗi):")
                for idx, row in fab_errors.iterrows():
                    col_k = row[col_k_name] if col_k_name else 'N/A'
                    error_msg = row['FAB_Pipe_Check']
                    print(f"   Dòng {idx+2:3d}: K={col_k} | {error_msg}")
                    
        # Lỗi Pap 1
        if apply_pap:
            pap1_errors = df[df['Pap1_Check'].str.startswith('FAIL', na=False)]
            if not pap1_errors.empty:
                print(f"\n❌ LỖI PAP 1 ({len(pap1_errors)} lỗi):")
                for idx, row in pap1_errors.iterrows():
                    col_o = row[col_o_name] if col_o_name and col_o_name in row.index else 'N/A'
                    error_msg = row['Pap1_Check']
                    print(f"   Dòng {idx+2:3d}: O={col_o} | {error_msg}")
                    
        # Lỗi Pap 2
        if apply_pap:
            pap2_errors = df[df['Pap2_Check'].str.startswith('FAIL', na=False)]
            if not pap2_errors.empty:
                print(f"\n❌ LỖI PAP 2 ({len(pap2_errors)} lỗi):")
                for idx, row in pap2_errors.iterrows():
                    col_p = row[col_p_name] if col_p_name and col_p_name in row.index else 'N/A'
                    error_msg = row['Pap2_Check']
                    print(f"   Dòng {idx+2:3d}: P={col_p} | {error_msg}")
    
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
            
        # Tổng kết FAB Pipe
        fab_total = self.fab_pipe_pass + self.fab_pipe_fail + self.fab_pipe_skip
        if fab_total > 0:
            print(f"\n🏭 FAB PIPE VALIDATION:")
            print(f"   ✅ PASS: {self.fab_pipe_pass:,}/{fab_total:,} ({self.fab_pipe_pass/fab_total*100:.1f}%)")
            print(f"   ❌ FAIL: {self.fab_pipe_fail:,}/{fab_total:,} ({self.fab_pipe_fail/fab_total*100:.1f}%)")
            print(f"   ⏭️ SKIP: {self.fab_pipe_skip:,}/{fab_total:,} ({self.fab_pipe_skip/fab_total*100:.1f}%)")
            
        # Tổng kết Pap 1
        pap1_total = self.pap1_pass + self.pap1_fail + self.pap1_skip
        if pap1_total > 0:
            print(f"\n📏 PAP 1 VALIDATION:")
            print(f"   ✅ PASS: {self.pap1_pass:,}/{pap1_total:,} ({self.pap1_pass/pap1_total*100:.1f}%)")
            print(f"   ❌ FAIL: {self.pap1_fail:,}/{pap1_total:,} ({self.pap1_fail/pap1_total*100:.1f}%)")
            print(f"   ⏭️ SKIP: {self.pap1_skip:,}/{pap1_total:,} ({self.pap1_skip/pap1_total*100:.1f}%)")
            
        # Tổng kết Pap 2
        pap2_total = self.pap2_pass + self.pap2_fail + self.pap2_skip
        if pap2_total > 0:
            print(f"\n📏 PAP 2 VALIDATION:")
            print(f"   ✅ PASS: {self.pap2_pass:,}/{pap2_total:,} ({self.pap2_pass/pap2_total*100:.1f}%)")
            print(f"   ❌ FAIL: {self.pap2_fail:,}/{pap2_total:,} ({self.pap2_fail/pap2_total*100:.1f}%)")
            print(f"   ⏭️ SKIP: {self.pap2_skip:,}/{pap2_total:,} ({self.pap2_skip/pap2_total*100:.1f}%)")
        
        print(f"\n🎉 VALIDATION HOÀN THÀNH!")

def main():
    """
    Main function để chạy validation
    """
    # Tìm file Excel trong thư mục hiện tại và thư mục cha
    current_dir_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~') and not f.startswith('validation')]
    parent_dir_files = [f for f in os.listdir('..') if f.endswith('.xlsx') and not f.startswith('~') and not f.startswith('validation')]
    
    excel_files = []
    # Thêm file từ thư mục hiện tại
    for f in current_dir_files:
        excel_files.append(f)
    # Thêm file từ thư mục cha
    for f in parent_dir_files:
        excel_files.append(f"../{f}")
    
    if not excel_files:
        print("❌ Không tìm thấy file Excel nào!")
        input("Nhấn Enter để thoát...")
        return
    
    # LUÔN LUÔN CHO USER CHỌN FILE (không tự động chọn)
    print("🔍 FILE EXCEL CÓ SẴN:")
    for i, file in enumerate(excel_files, 1):
        file_path = file if not file.startswith('..') else file
        try:
            size = os.path.getsize(file) // 1024
            display_name = os.path.basename(file)
            print(f" {i}. {display_name:40s} ({size} KB)")
        except:
            display_name = os.path.basename(file) 
            print(f" {i}. {display_name:40s}")
    
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
    validator = ExcelValidatorEnhanced()
    validator.validate_excel_file(selected_file)

if __name__ == "__main__":
    main()
