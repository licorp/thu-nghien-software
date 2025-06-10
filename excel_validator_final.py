#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EXCEL VALIDATION TOOL - PRODUCTION VERSION
==========================================

Tool validation Excel cho dự án pipe/equipment data với 4 quy tắc:
1. Array Number Validation
2. Pipe Treatment Validation  
3. CP-INTERNAL Array Number Validation
4. Priority-based Pipe Schedule Mapping Validation

Tác giả: GitHub Copilot
Ngày: 2025-06-10
"""

import pandas as pd
import re
from pathlib import Path
from datetime import datetime

class ExcelValidator:
    """Excel validation với 4 quy tắc hoàn chỉnh"""
    
    def __init__(self):
        self.worksheets_config = {
            'array_number': ['Pipe Schedule', 'Pipe Fitting Schedule', 'Pipe Accessory Schedule', 'Sprinkler Schedule'],
            'pipe_treatment': ['Pipe Schedule', 'Pipe Fitting Schedule', 'Pipe Accessory Schedule'],
            'cp_internal': ['Pipe Schedule', 'Pipe Fitting Schedule', 'Pipe Accessory Schedule'],
            'pipe_mapping': ['Pipe Schedule']
        }
        
        self.total_rows = 0
        self.total_pass = 0
        self.total_fail = 0
        self.validation_results = {}
    
    def validate_excel_file(self, excel_file_path):
        """Validate toàn bộ file Excel với 4 quy tắc"""
        try:
            print("=" * 80)
            print("🚀 EXCEL VALIDATION TOOL - PRODUCTION VERSION")
            print("=" * 80)
            print(f"📁 File: {excel_file_path}")
            print(f"🕐 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print()
            
            xl_file = pd.ExcelFile(excel_file_path)
            
            for sheet_name in xl_file.sheet_names:
                self._validate_worksheet(excel_file_path, sheet_name)
            
            self._generate_summary()
            return self._export_results(excel_file_path)
            
        except Exception as e:
            print(f"❌ Lỗi validation: {e}")
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
            'pipe_mapping': sheet_name in self.worksheets_config['pipe_mapping']
        }
        
        for rule, apply in rules.items():
            status = "✅ ÁP DỤNG" if apply else "❌ KHÔNG ÁP DỤNG"
            print(f"{rule.replace('_', ' ').title()} validation: {status}")
        
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
                result = self._check_pipe_schedule_mapping(row, cols['F'], cols['G'], cols['K'], cols['L'], cols['M'])
                if result != "PASS" and not result.startswith("SKIP"):
                    errors.append(f"Mapping: {result}")
            
            return f"FAIL: {'; '.join(errors[:4])}" if errors else "PASS"
                
        except Exception as e:
            return f"ERROR: {str(e)}"
    
    def _check_array_number(self, row, col_a, col_b, col_d):
        """Rule 1: Array Number format"""
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
    
    def _validate_mapping_rule(self, fab_pipe, fab_pipe_str, end_1_str, end_2_str, expected, rule_name):
        """Validate một rule mapping cụ thể"""
        errors = []
        expected_fab_pipe, expected_end_1, expected_end_2 = expected
        
        # Kiểm tra FAB Pipe
        if pd.isna(fab_pipe):
            errors.append(f"{rule_name} cần FAB Pipe '{expected_fab_pipe}', nhưng thiếu")
        elif fab_pipe_str != expected_fab_pipe:
            errors.append(f"{rule_name} cần FAB Pipe '{expected_fab_pipe}', có '{fab_pipe_str}'")
        
        # Kiểm tra End-1, End-2 (skip N/A)
        for end_str, expected_end, end_name in [(end_1_str, expected_end_1, "End-1"), (end_2_str, expected_end_2, "End-2")]:
            if end_str not in ["", "N/A", "nan"] and end_str != expected_end:
                errors.append(f"{rule_name.split('(')[0].strip()} cần {end_name} '{expected_end}', có '{end_str}'")
        
        return f"{'; '.join(errors)}" if errors else "PASS"
    
    def _validate_fab_pipe_only(self, fab_pipe, fab_pipe_str, expected_fab_pipe):
        """Validate chỉ FAB Pipe"""
        if pd.isna(fab_pipe):
            return f"{expected_fab_pipe} cần FAB Pipe '{expected_fab_pipe}', nhưng thiếu"
        elif fab_pipe_str != expected_fab_pipe:
            return f"{expected_fab_pipe} cần FAB Pipe '{expected_fab_pipe}', có '{fab_pipe_str}'"
        return "PASS"
    
    def _show_sample_errors(self, df, cols):
        """Hiển thị lỗi với màu sắc"""
        fail_rows = df[df['Validation_Check'] != 'PASS']
        if fail_rows.empty:
            return
            
        print(f"📋 {len(fail_rows)} LỖI (ĐỎ=SAI, TRẮNG=ĐÚNG):")
        for idx, row in fail_rows.iterrows():
            # Hiển thị thông tin dòng
            info_cols = ['C', 'D', 'F', 'G', 'K', 'T']
            col_info = " | ".join([f"{c}={row[cols[c]] if cols[c] else 'N/A'}" for c in info_cols])
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
        
        if self.total_rows > 0:
            pass_rate = self.total_pass / self.total_rows * 100
            print(f"✅ PASS: {self.total_pass:,}/{self.total_rows:,} ({pass_rate:.1f}%)")
            print(f"❌ FAIL: {self.total_fail:,}/{self.total_rows:,} ({100-pass_rate:.1f}%)")
            
            print(f"\n📊 CHI TIẾT THEO WORKSHEET:")
            for sheet_name, df in self.validation_results.items():
                sheet_total = len(df)
                sheet_pass = len(df[df['Validation_Check'] == 'PASS'])
                sheet_rate = sheet_pass / sheet_total * 100
                print(f"  {sheet_name:25s}: {sheet_pass:3d}/{sheet_total:3d} ({sheet_rate:5.1f}%)")
    
    def _export_results(self, excel_file_path):
        """Xuất file kết quả"""
        if not self.validation_results:
            return None
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = Path(excel_file_path).stem
        output_file = f"validation_4rules_{base_name}_{timestamp}.xlsx"
        
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
    """Hàm main để chạy validation tool"""
    try:
        current_dir = Path(".")
        excel_files = [f for f in current_dir.glob("*.xlsx") 
                      if not f.name.startswith('~') 
                      and 'validation' not in f.name.lower()]
        
        if not excel_files:
            print("❌ Không tìm thấy file Excel để validation!")
            return
        
        print("🔍 FILE EXCEL CÓ SẴN:")
        for i, file in enumerate(excel_files, 1):
            file_size = file.stat().st_size / 1024
            print(f"{i:2d}. {file.name:40s} ({file_size:,.0f} KB)")
        
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
        
        validator = ExcelValidator()
        output_file = validator.validate_excel_file(selected_file)
        
        if output_file:
            print(f"\n🎉 VALIDATION HOÀN THÀNH!")
            print(f"📁 Kết quả: {output_file}")
        else:
            print(f"\n❌ VALIDATION THẤT BẠI!")
            
    except KeyboardInterrupt:
        print("\n⏹️ Đã hủy bởi người dùng!")
    except Exception as e:
        print(f"\n❌ Lỗi không mong muốn: {e}")

if __name__ == "__main__":
    main()
