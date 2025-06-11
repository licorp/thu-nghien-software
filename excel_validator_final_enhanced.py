#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EXCEL VALIDATION TOOL - PRODUCTION VERSION WITH 6 RULES (ENHANCED ERROR REPORTING)
================================================================================

Tool validation Excel cho dự án pipe/equipment data với 6 quy tắc và báo cáo lỗi chi tiết theo cột:
1. Array Number Validation
2. Pipe Treatment Validation  
3. CP-INTERNAL Array Number Validation
4. Priority-based Pipe Schedule Mapping Validation
5. EE_Run Dim & EE_Pap Validation (với báo cáo chi tiết cột N, O)
6. Item Description = Family Validation (Pipe Accessory Schedule)

Cập nhật: Báo cáo chi tiết cột K (FAB Pipe), L (End-1), M (End-2), N (EE_Run Dim 1), O (EE_Pap 1)

Tác giả: GitHub Copilot
Ngày: 2025-06-11
"""

import pandas as pd
import re
from pathlib import Path
from datetime import datetime

class ExcelValidator:
    """Excel validation với 6 quy tắc hoàn chỉnh và báo cáo lỗi chi tiết"""
    
    def __init__(self):
        self.worksheets_config = {
            'array_number': ['Pipe Schedule', 'Pipe Fitting Schedule', 'Pipe Accessory Schedule', 'Sprinkler Schedule'],
            'pipe_treatment': ['Pipe Schedule', 'Pipe Fitting Schedule', 'Pipe Accessory Schedule'],
            'cp_internal': ['Pipe Schedule', 'Pipe Fitting Schedule', 'Pipe Accessory Schedule'],
            'pipe_mapping': ['Pipe Schedule'],
            'ee_run_pap': ['Pipe Schedule'],
            'item_family_match': ['Pipe Accessory Schedule']  # New validation for Item Description = Family
        }
        
        self.total_rows = 0
        self.total_pass = 0
        self.total_fail = 0
        self.validation_results = {}
    
    def validate_excel_file(self, excel_file_path):
        """Validate toàn bộ file Excel với 6 quy tắc"""
        try:
            print("=" * 80)
            print("🚀 EXCEL VALIDATION TOOL - ENHANCED WITH 6 RULES & DETAILED ERROR REPORTING")
            print("=" * 80)
            print(f"📁 File: {excel_file_path}")
            print(f"🕐 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print()
            
            xl_file = pd.ExcelFile(excel_file_path)
            
            for sheet_name in xl_file.sheet_names:
                self._validate_worksheet(excel_file_path, sheet_name)
            
            self._generate_summary()
            
            # Export results
            output_file = self._export_results(excel_file_path)
            print(f"📁 Kết quả: {output_file}")
            return output_file
            
        except Exception as e:
            print(f"❌ Lỗi: {e}")
            return None

    def _validate_worksheet(self, excel_file_path, sheet_name):
        """Validate một worksheet với tất cả rules"""
        print(f"📊 WORKSHEET: {sheet_name}")
        print("-" * 50)
        
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        print(f"Số dòng: {len(df)}, Số cột: {len(df.columns)}")
        
        # Kiểm tra rule nào áp dụng
        rules = {
            'array_number': sheet_name in self.worksheets_config['array_number'],
            'pipe_treatment': sheet_name in self.worksheets_config['pipe_treatment'],
            'cp_internal': sheet_name in self.worksheets_config['cp_internal'],
            'pipe_mapping': sheet_name in self.worksheets_config['pipe_mapping'],
            'ee_run_pap': sheet_name in self.worksheets_config['ee_run_pap'],
            'item_family_match': sheet_name in self.worksheets_config['item_family_match']
        }
        
        for rule, apply in rules.items():
            status = "✅ ÁP DỤNG" if apply else "❌ KHÔNG ÁP DỤNG"
            rule_name = rule.replace('_', ' ').title().replace('Ee Run Pap', 'EE Run Dim/Pap').replace('Item Family Match', 'Item-Family Match')
            print(f"{rule_name} validation: {status}")
        
        if not any(rules.values()):
            print("⏭️ Bỏ qua worksheet")
            print()
            return
        
        # Lấy tên cột theo vị trí
        cols = {chr(65+i): df.columns[i] if len(df.columns) > i else None for i in range(26)}
        
        # Áp dụng validation
        df['Validation_Check'] = df.apply(lambda row: self._validate_row(row, cols, rules), axis=1)
        
        # Thống kê
        sheet_total = len(df)
        sheet_pass = len(df[df['Validation_Check'] == 'PASS'])
        sheet_fail = sheet_total - sheet_pass
        
        print(f"✅ PASS: {sheet_pass}/{sheet_total} ({sheet_pass/sheet_total*100:.1f}%)")
        print(f"❌ FAIL: {sheet_fail}/{sheet_total} ({sheet_fail/sheet_total*100:.1f}%)")
        
        self.total_rows += sheet_total
        self.total_pass += sheet_pass  
        self.total_fail += sheet_fail
        self.validation_results[sheet_name] = df
        
        self._show_sample_errors(df, cols)
        self._check_empty_cells(df, sheet_name, cols, rules)
        print()

    def _validate_row(self, row, cols, rules):
        """Validate một dòng với tất cả rules"""
        errors = []
        
        try:
            # Kiểm tra CP-INTERNAL
            is_cp_internal = False
            if cols['C'] and not pd.isna(row[cols['C']]):
                is_cp_internal = str(row[cols['C']]).strip() == "CP-INTERNAL"
            
            # Rule 1: Array Number (skip nếu CP-INTERNAL)
            if rules['array_number'] and not is_cp_internal and all(cols[c] for c in ['A', 'B', 'D']):
                result = self._check_array_number(row, cols['A'], cols['B'], cols['D'])
                if result != "PASS" and not result.startswith("SKIP"):
                    errors.append(f"Array: {result}")
            
            # Rule 2: Pipe Treatment
            if rules['pipe_treatment'] and all(cols[c] for c in ['C', 'T']):
                result = self._check_pipe_treatment(row, cols['C'], cols['T'])
                if result != "PASS" and not result.startswith("SKIP"):
                    errors.append(f"Treatment: {result}")
            
            # Rule 3: CP-INTERNAL Array
            if rules['cp_internal'] and all(cols[c] for c in ['A', 'C', 'D']):
                result = self._check_cp_internal_array(row, cols['A'], cols['C'], cols['D'])
                if result != "PASS" and not result.startswith("SKIP"):
                    errors.append(f"CP-Internal: {result}")
            
            # Rule 4: Pipe Schedule Mapping
            if rules['pipe_mapping'] and all(cols[c] for c in ['F', 'G', 'K']):
                result = self._check_pipe_schedule_mapping(row, cols['F'], cols['G'], cols['K'], cols.get('L'), cols.get('M'))
                if result != "PASS" and not result.startswith("SKIP"):
                    errors.append(f"Mapping: {result}")
            
            # Rule 5: EE_Run Dim & EE_Pap Validation
            if rules['ee_run_pap'] and all(cols[c] for c in ['F', 'G']):
                result = self._check_ee_run_pap(row, cols['F'], cols['G'], cols.get('N'), cols.get('O'), cols.get('P'), cols.get('Q'), cols.get('R'), cols.get('S'), cols.get('L'), cols.get('M'))
                if result != "PASS" and not result.startswith("SKIP"):
                    errors.append(f"EE_Run/Pap: {result}")
            
            # Rule 6: Item Description = Family Validation
            if rules['item_family_match'] and all(cols[c] for c in ['F', 'U']):
                result = self._check_item_family_match(row, cols['F'], cols['U'])
                if result != "PASS" and not result.startswith("SKIP"):
                    errors.append(f"Item-Family: {result}")
            
            return "PASS" if not errors else "; ".join(errors)
            
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_array_number(self, row, col_a, col_b, col_d):
        """Rule 1: Array Number validation"""
        try:
            cross_passage, location_lanes, array_number = row[col_a], row[col_b], row[col_d]
            
            if any(pd.isna(x) for x in [cross_passage, location_lanes, array_number]):
                return "SKIP: Thiếu dữ liệu"
            
            # Lấy 2 số cuối
            def get_last_2_digits(text):
                numbers = re.findall(r'\d+', str(text).strip())
                if numbers:
                    return numbers[-1][-2:] if len(numbers[-1]) >= 2 else numbers[-1].zfill(2)
                return "00"
            
            last_2_b = get_last_2_digits(location_lanes)
            last_2_a = get_last_2_digits(cross_passage)
            required_pattern = f"EXP6{last_2_b}{last_2_a}"
            actual_array = str(array_number).strip()
            
            return "PASS" if required_pattern in actual_array else f"cần '{required_pattern}', có '{actual_array}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_pipe_treatment(self, row, col_c, col_t):
        """Rule 2: Pipe Treatment"""
        try:
            system_type, pipe_treatment = row[col_c], row[col_t]
            
            if any(pd.isna(x) for x in [system_type, pipe_treatment]):
                return "SKIP: Thiếu dữ liệu"
            
            system_type_str = str(system_type).strip()
            pipe_treatment_str = str(pipe_treatment).strip()
            
            expected_map = {
                "CP-INTERNAL": "GAL",
                "CP-EXTERNAL": "BLACK",
                "CW-DISTRIBUTION": "BLACK", 
                "CW-ARRAY": "BLACK"
            }
            
            expected = expected_map.get(system_type_str)
            if not expected:
                return "PASS"
            
            return "PASS" if pipe_treatment_str == expected else f"'{system_type_str}' cần '{expected}', có '{pipe_treatment_str}'"
        
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_cp_internal_array(self, row, col_a, col_c, col_d):
        """Rule 3: CP-INTERNAL Array matching"""
        try:
            cross_passage, system_type, array_number = row[col_a], row[col_c], row[col_d]
            
            if pd.isna(system_type):
                return "SKIP: Thiếu System Type"
            
            system_type_str = str(system_type).strip()
            if system_type_str != "CP-INTERNAL":
                return "PASS"
            
            if any(pd.isna(x) for x in [cross_passage, array_number]):
                return "SKIP: Thiếu dữ liệu Cross Passage hoặc Array Number"
            
            cross_passage_str = str(cross_passage).strip()
            array_number_str = str(array_number).strip()
            
            return "PASS" if cross_passage_str == array_number_str else f"Array Number phải trùng Cross Passage: cần '{cross_passage_str}', có '{array_number_str}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_pipe_schedule_mapping(self, row, col_f, col_g, col_k, col_l, col_m):
        """Rule 4: Priority-based Pipe Schedule Mapping"""
        try:
            item_description, size, fab_pipe = row[col_f], row[col_g], row[col_k]
            end_1 = row[col_l] if col_l else None
            end_2 = row[col_m] if col_m else None
            
            # Skip nếu thiếu toàn bộ dữ liệu
            if all(pd.isna(x) for x in [item_description, size, end_1, end_2]):
                return "SKIP: Thiếu Item Description, Size, End-1, và End-2"
            
            # Chuẩn bị dữ liệu
            def safe_str(val):
                return str(val).strip() if not pd.isna(val) else ""
            
            item_desc_str, size_str, fab_pipe_str = map(safe_str, [item_description, size, fab_pipe])
            end_1_str, end_2_str = map(safe_str, [end_1, end_2])
            
            # HIGH PRIORITY RULES
            priority_rules = [
                # STD 1 PAP RANGE
                {
                    'condition': (size_str in ["65.0", "65"]) and "4730" in item_desc_str,
                    'expected': ("STD 1 PAP RANGE", "RG", "BE"),
                    'name': "STD 1 PAP RANGE (size 65, 4730)"
                },
                # STD 2 PAP RANGE  
                {
                    'condition': (size_str in ["65.0", "65"]) and "5295" in item_desc_str,
                    'expected': ("STD 2 PAP RANGE", "RG", "BE"),
                    'name': "STD 2 PAP RANGE (size 65, 5295)"
                },
                # STD ARRAY TEE
                {
                    'condition': ((size_str in ["150.0", "150"]) and "900" in item_desc_str) or "150-900" in item_desc_str,
                    'expected': ("STD ARRAY TEE", "RG", "RG"),
                    'name': "STD ARRAY TEE (150-900)"
                }
            ]
            
            # Kiểm tra high priority rules
            for rule in priority_rules:
                if rule['condition']:
                    return self._validate_mapping_rule(fab_pipe, fab_pipe_str, end_1_str, end_2_str, rule['expected'], rule['name'])
            
            # LOW PRIORITY RULES (chỉ khi không match high priority)
            # Groove_Thread
            if ((end_1_str == "RG" and end_2_str == "RG") or 
                (size_str == "40" and end_1_str == "TH" and end_2_str == "TH")):
                return self._validate_fab_pipe_only(fab_pipe, fab_pipe_str, "Groove_Thread")
            
            # Fabrication
            if (size_str == "65" and end_1_str == "RG" and end_2_str == "BE" and 
                "4730" not in item_desc_str and "5295" not in item_desc_str):
                return self._validate_fab_pipe_only(fab_pipe, fab_pipe_str, "Fabrication")
            
            return "PASS"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_ee_run_pap(self, row, col_f, col_g, col_n, col_o, col_p, col_q, col_r, col_s, col_l, col_m):
        """Rule 5: EE_Run Dim & EE_Pap Validation với báo cáo chi tiết cột N, O"""
        try:
            item_description = row[col_f] if col_f else None
            size = row[col_g] if col_g else None
            
            # Skip nếu thiếu Item Description và Size
            if pd.isna(item_description) and pd.isna(size):
                return "SKIP: Thiếu Item Description và Size"
            
            def safe_str(val):
                return str(val).strip() if not pd.isna(val) else ""
            
            item_desc_str = safe_str(item_description)
            size_str = safe_str(size)
            
            # Get all EE_Run Dim and EE_Pap values
            ee_run_dim_1 = row[col_n] if col_n else None
            ee_pap_1 = row[col_o] if col_o else None
            ee_run_dim_2 = row[col_p] if col_p else None
            ee_pap_2 = row[col_q] if col_q else None
            ee_run_dim_3 = row[col_r] if col_r else None
            ee_pap_3 = row[col_s] if col_s else None
            
            # Get End-1 and End-2 for Fabrication check
            end_1 = row[col_l] if col_l else None
            end_2 = row[col_m] if col_m else None
            end_1_str = safe_str(end_1)
            end_2_str = safe_str(end_2)
            
            errors = []
            
            # HIGH PRIORITY RULES with specific EE_Run Dim & EE_Pap requirements
            
            # STD 1 PAP RANGE: ống 65 dài 4730 → EE_Run Dim 1: 4685, EE_Pap 1: 40B
            if (size_str in ["65.0", "65"]) and "4730" in item_desc_str:
                if pd.isna(ee_run_dim_1) or safe_str(ee_run_dim_1) not in ["4685", "4685.0"]:
                    errors.append(f"Cột N (EE_Run Dim 1): STD 1 PAP RANGE cần '4685', có '{safe_str(ee_run_dim_1)}'")
                if pd.isna(ee_pap_1) or safe_str(ee_pap_1) != "40B":
                    errors.append(f"Cột O (EE_Pap 1): STD 1 PAP RANGE cần '40B', có '{safe_str(ee_pap_1)}'")
            
            # STD 2 PAP RANGE: ống 65 dài 5295 → EE_Run Dim 1: 150, EE_Pap 1: 40B, EE_Run Dim 2: 5250, EE_Pap 2: 40B
            elif (size_str in ["65.0", "65"]) and "5295" in item_desc_str:
                if pd.isna(ee_run_dim_1) or safe_str(ee_run_dim_1) not in ["150", "150.0"]:
                    errors.append(f"Cột N (EE_Run Dim 1): STD 2 PAP RANGE cần '150', có '{safe_str(ee_run_dim_1)}'")
                if pd.isna(ee_pap_1) or safe_str(ee_pap_1) != "40B":
                    errors.append(f"Cột O (EE_Pap 1): STD 2 PAP RANGE cần '40B', có '{safe_str(ee_pap_1)}'")
                if pd.isna(ee_run_dim_2) or safe_str(ee_run_dim_2) not in ["5250", "5250.0"]:
                    errors.append(f"Cột P (EE_Run Dim 2): STD 2 PAP RANGE cần '5250', có '{safe_str(ee_run_dim_2)}'")
                if pd.isna(ee_pap_2) or safe_str(ee_pap_2) != "40B":
                    errors.append(f"Cột Q (EE_Pap 2): STD 2 PAP RANGE cần '40B', có '{safe_str(ee_pap_2)}'")
            
            # STD ARRAY TEE: ống 150 dài 900 → EE_Run Dim 1: 150, EE_Pap 1: 65LR
            elif ((size_str in ["150.0", "150"]) and "900" in item_desc_str) or "150-900" in item_desc_str:
                if pd.isna(ee_run_dim_1) or safe_str(ee_run_dim_1) not in ["150", "150.0"]:
                    errors.append(f"Cột N (EE_Run Dim 1): STD ARRAY TEE cần '150', có '{safe_str(ee_run_dim_1)}'")
                if pd.isna(ee_pap_1) or safe_str(ee_pap_1) != "65LR":
                    errors.append(f"Cột O (EE_Pap 1): STD ARRAY TEE cần '65LR', có '{safe_str(ee_pap_1)}'")
            
            # Fabrication case: ống 65 RG BE (không phải PAP RANGE) - cần tối thiểu EE_Run Dim 1 và EE_Pap 1
            elif (size_str == "65" and end_1_str == "RG" and end_2_str == "BE" and 
                  "4730" not in item_desc_str and "5295" not in item_desc_str):
                if pd.isna(ee_run_dim_1) or safe_str(ee_run_dim_1) == "":
                    errors.append(f"Cột N (EE_Run Dim 1): Fabrication (ống 65 RG BE) thiếu EE_Run Dim 1")
                if pd.isna(ee_pap_1) or safe_str(ee_pap_1) == "":
                    errors.append(f"Cột O (EE_Pap 1): Fabrication (ống 65 RG BE) thiếu EE_Pap 1")
            
            # Check for "Thiếu" or "Sai" values in any EE_Run Dim or EE_Pap columns
            all_ee_values = [
                (ee_run_dim_1, "EE_Run Dim 1", "N"),
                (ee_pap_1, "EE_Pap 1", "O"),
                (ee_run_dim_2, "EE_Run Dim 2", "P"),
                (ee_pap_2, "EE_Pap 2", "Q"),
                (ee_run_dim_3, "EE_Run Dim 3", "R"),
                (ee_pap_3, "EE_Pap 3", "S")
            ]
            
            for value, col_name, col_letter in all_ee_values:
                if not pd.isna(value):
                    value_str = safe_str(value).upper()
                    if value_str in ["THIẾU", "SAI"]:
                        errors.append(f"Cột {col_letter} ({col_name}): có giá trị '{value_str}' - cần kiểm tra và sửa")
            
            if errors:
                return "; ".join(errors)
            else:
                return "PASS"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_item_family_match(self, row, col_f, col_u):
        """Rule 6: Item Description phải trùng với Family (Pipe Accessory Schedule)"""
        try:
            item_description = row[col_f] if col_f else None
            family = row[col_u] if col_u else None
            
            # Skip nếu thiếu dữ liệu
            if pd.isna(item_description) and pd.isna(family):
                return "SKIP: Thiếu Item Description và Family"
            
            def safe_str(val):
                return str(val).strip() if not pd.isna(val) else ""
            
            item_desc_str = safe_str(item_description)
            family_str = safe_str(family)
            
            # Cả hai đều trống thì PASS
            if item_desc_str == "" and family_str == "":
                return "PASS"
            
            # Một trong hai trống thì FAIL
            if item_desc_str == "" or family_str == "":
                return f"Cột F (Item Description) '{item_desc_str}' và Cột U (Family) '{family_str}' phải cùng có giá trị hoặc cùng trống"
            
            # So sánh giá trị
            if item_desc_str == family_str:
                return "PASS"
            else:
                return f"Cột F (Item Description) phải trùng Cột U (Family): cần '{family_str}', có '{item_desc_str}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _validate_mapping_rule(self, fab_pipe, fab_pipe_str, end_1_str, end_2_str, expected, rule_name):
        """Validate một rule mapping cụ thể với báo cáo chi tiết cột"""
        errors = []
        expected_fab_pipe, expected_end_1, expected_end_2 = expected
        
        # Kiểm tra FAB Pipe (Cột K)
        if pd.isna(fab_pipe):
            errors.append(f"Cột K (FAB Pipe): {rule_name} cần '{expected_fab_pipe}', nhưng thiếu")
        elif fab_pipe_str != expected_fab_pipe:
            errors.append(f"Cột K (FAB Pipe): {rule_name} cần '{expected_fab_pipe}', có '{fab_pipe_str}'")
        
        # Kiểm tra End-1 (Cột L), End-2 (Cột M) (skip N/A)
        end_mappings = [
            (end_1_str, expected_end_1, "L", "End-1"), 
            (end_2_str, expected_end_2, "M", "End-2")
        ]
        
        for end_str, expected_end, col_letter, end_name in end_mappings:
            if end_str not in ["", "N/A", "nan"] and end_str != expected_end:
                errors.append(f"Cột {col_letter} ({end_name}): {rule_name.split('(')[0].strip()} cần '{expected_end}', có '{end_str}'")
        
        return f"{'; '.join(errors)}" if errors else "PASS"
    
    def _validate_fab_pipe_only(self, fab_pipe, fab_pipe_str, expected_fab_pipe):
        """Validate chỉ FAB Pipe với báo cáo chi tiết về cột"""
        if pd.isna(fab_pipe):
            return f"Cột K (FAB Pipe): {expected_fab_pipe} cần '{expected_fab_pipe}', nhưng thiếu"
        elif fab_pipe_str != expected_fab_pipe:
            return f"Cột K (FAB Pipe): {expected_fab_pipe} cần '{expected_fab_pipe}', có '{fab_pipe_str}'"
        else:
            return "PASS"

    def _show_sample_errors(self, df, cols):
        """Hiển thị lỗi với màu sắc và thông tin chi tiết cột"""
        fail_rows = df[df['Validation_Check'] != 'PASS']
        if fail_rows.empty:
            return
            
        print(f"📋 {len(fail_rows)} LỖI (ĐỎ=SAI, TRẮNG=ĐÚNG):")
        for idx, row in fail_rows.iterrows():
            # Hiển thị thông tin dòng với nhiều cột hơn
            info_cols = ['C', 'D', 'F', 'G', 'K', 'L', 'M', 'N', 'O', 'T']
            col_info = " | ".join([f"{c}={str(row[cols[c]])[:15] if cols[c] and not pd.isna(row[cols[c]]) else 'N/A'}" for c in info_cols])
            print(f"  Dòng {idx+2:3d}: {col_info}")
            
            # Hiển thị lỗi với màu sắc
            check_result = row['Validation_Check']
            if "cần '" in check_result and "', có '" in check_result:
                parts = check_result.split("cần '")
                if len(parts) > 1:
                    prefix = parts[0]
                    remaining = parts[1]
                    if "', có '" in remaining:
                        expected_and_actual = remaining.split("', có '")
                        expected = expected_and_actual[0]
                        actual = expected_and_actual[1].rstrip("'")
                        print(f"           {prefix}cần '\033[97m{expected}\033[0m', có '\033[91m{actual}\033[0m'")
                        continue
            print(f"           {check_result}")
    
    def _generate_summary(self):
        """Tạo báo cáo tổng kết"""
        print("=" * 80)
        print("📈 TỔNG KẾT VALIDATION")
        print("=" * 80)
        print()
        print(f"✅ PASS: {self.total_pass:,}/{self.total_rows:,} ({self.total_pass/self.total_rows*100:.1f}%)")
        print(f"❌ FAIL: {self.total_fail:,}/{self.total_rows:,} ({self.total_fail/self.total_rows*100:.1f}%)")
        print("📊 CHI TIẾT THEO WORKSHEET:")
        
        for sheet_name, df in self.validation_results.items():
            sheet_total = len(df)
            sheet_pass = len(df[df['Validation_Check'] == 'PASS'])
            print(f"  {sheet_name:25}: {sheet_pass:3d}/{sheet_total:3d} ({sheet_pass/sheet_total*100:5.1f}%)")
    
    def _export_results(self, excel_file_path):
        """Xuất file kết quả"""
        try:
            file_path = Path(excel_file_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = file_path.parent / f"validation_6rules_enhanced_{file_path.stem}_{timestamp}.xlsx"
            
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, df in self.validation_results.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            return output_file.name
            
        except Exception as e:
            print(f"⚠️ Không thể xuất file: {e}")
            return None
    
    def _check_empty_cells(self, df, sheet_name, cols, rules):
        """Kiểm tra ô trống cho từng validation rule"""
        print(f"📋 BÁO CÁO Ô TRỐNG - {sheet_name}:")
        
        # Kiểm tra ô trống cho từng rule
        rule_configs = {
            'array_number': (['A', 'B', 'D'], ['Cross Passage', 'Location Lanes', 'Array Number']),
            'pipe_treatment': (['C', 'T'], ['System Type', 'Pipe Treatment']),
            'cp_internal': (['A', 'C', 'D'], ['Cross Passage', 'System Type', 'Array Number']),
            'pipe_mapping': (['F', 'G', 'K', 'L', 'M'], ['Item Description', 'Size', 'FAB Pipe', 'End-1', 'End-2']),
            'ee_run_pap': (['N', 'O', 'P', 'Q', 'R', 'S'], ['EE_Run Dim 1', 'EE_Pap 1', 'EE_Run Dim 2', 'EE_Pap 2', 'EE_Run Dim 3', 'EE_Pap 3']),
            'item_family_match': (['F', 'U'], ['Item Description', 'Family'])
        }
        
        for rule_name, (col_letters, col_descriptions) in rule_configs.items():
            if rules[rule_name]:  # Chỉ kiểm tra rule được áp dụng
                print(f"  🔍 {rule_name.replace('_', ' ').title().replace('Ee Run Pap', 'EE Run Dim/Pap').replace('Item Family Match', 'Item-Family Match')}:")
                
                for col_letter, col_desc in zip(col_letters, col_descriptions):
                    if cols[col_letter]:  # Cột tồn tại
                        col_name = cols[col_letter]
                        empty_count = df[col_name].isna().sum()
                        total_count = len(df)
                        if empty_count > 0:
                            print(f"    ❌ Cột {col_letter} ({col_desc}): {empty_count}/{total_count} ô trống ({empty_count/total_count*100:.1f}%)")

def main():
    """Main function"""
    try:
        print("🔍 EXCEL VALIDATION TOOL - ENHANCED VERSION WITH 6 RULES & DETAILED ERROR REPORTING")
        print("=" * 80)
        
        # Tìm file Excel trong thư mục hiện tại
        current_dir = Path('.')
        excel_files = list(current_dir.glob('*.xlsx'))
        excel_files = [f for f in excel_files if not f.name.startswith('~$') and not f.name.startswith('validation_')]
        
        if not excel_files:
            print("❌ Không tìm thấy file Excel nào trong thư mục hiện tại")
            return
        
        # Hiển thị danh sách file
        print("🔍 FILE EXCEL CÓ SẴN:")
        for i, file in enumerate(excel_files, 1):
            file_size = file.stat().st_size / 1024  # KB
            print(f" {i}. {file.name}  ({file_size:.0f} KB)")
        
        # Cho user chọn file
        while True:
            choice = input(f"✏️ Chọn file (1-{len(excel_files)}) hoặc 'q' để thoát: ").strip()
            
            if choice.lower() == 'q':
                print("👋 Tạm biệt!")
                return
                
            try:
                file_index = int(choice) - 1
                if 0 <= file_index < len(excel_files):
                    selected_file = excel_files[file_index]
                    break
                else:
                    print(f"❌ Vui lòng chọn số từ 1 đến {len(excel_files)}")
            except ValueError:
                print("❌ Vui lòng nhập số hợp lệ")
        
        # Chạy validation
        validator = ExcelValidator()
        output_file = validator.validate_excel_file(selected_file)
        
        if output_file:
            print("🎉 VALIDATION HOÀN THÀNH!")
            print(f"📁 Kết quả: {output_file}")
        
    except KeyboardInterrupt:
        print("\n⚠️ Đã hủy bởi người dùng")
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        import traceback
        traceback.print_exc()
        
def run_validation_on_file(filepath):
    """Run validation on a specific file - utility function for testing"""
    validator = ExcelValidator()
    return validator.validate_excel_file(filepath)

if __name__ == "__main__":
    main()
