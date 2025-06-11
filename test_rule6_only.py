#!/usr/bin/env python3
"""
Direct test script that doesn't import the module
"""

import pandas as pd
import re
from pathlib import Path
from datetime import datetime

# Copy the class definition directly
class ExcelValidator:
    """Excel validation với 6 quy tắc hoàn chỉnh"""
    
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
            print("🚀 EXCEL VALIDATION TOOL - ENHANCED WITH 6 RULES")
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
            # Rule 6: Item Description = Family Validation
            if rules['item_family_match'] and all(cols[c] for c in ['F', 'U']):
                result = self._check_item_family_match(row, cols['F'], cols['U'])
                if result != "PASS" and not result.startswith("SKIP"):
                    errors.append(f"Item-Family: {result}")
            
            return "PASS" if not errors else "; ".join(errors)
            
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
                return f"Item Description '{item_desc_str}' và Family '{family_str}' phải cùng có giá trị hoặc cùng trống"
            
            # So sánh giá trị
            if item_desc_str == family_str:
                return "PASS"
            else:
                return f"Item Description phải trùng Family: cần '{family_str}', có '{item_desc_str}'"
                
        except Exception as e:
            return f"ERROR: {str(e)}"

    def _show_sample_errors(self, df, cols):
        """Hiển thị lỗi với màu sắc"""
        fail_rows = df[df['Validation_Check'] != 'PASS']
        if fail_rows.empty:
            print("✅ Không có lỗi nào!")
            return
            
        print(f"📋 {len(fail_rows)} LỖI ĐƯỢC PHÁT HIỆN:")
        for idx, row in fail_rows.head(10).iterrows():  # Show only first 10 errors
            check_result = row['Validation_Check']
            print(f"  Dòng {idx+2:3d}: {check_result}")
    
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
            output_file = file_path.parent / f"validation_6rules_{file_path.stem}_{timestamp}.xlsx"
            
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
        
        # Kiểm tra ô trống cho Rule 6 nếu áp dụng
        if rules['item_family_match']:
            rule_configs = {
                'item_family_match': (['F', 'U'], ['Item Description', 'Family'])
            }
            
            for rule_name, (col_letters, col_descriptions) in rule_configs.items():
                print(f"  🔍 {rule_name.replace('_', ' ').title().replace('Item Family Match', 'Item-Family Match')}:")
                
                for col_letter, col_desc in zip(col_letters, col_descriptions):
                    if cols[col_letter]:  # Cột tồn tại
                        col_name = cols[col_letter]
                        empty_count = df[col_name].isna().sum()
                        total_count = len(df)
                        if empty_count > 0:
                            print(f"    ❌ Cột {col_letter} ({col_desc}): {empty_count}/{total_count} ô trống ({empty_count/total_count*100:.1f}%)")
                        else:
                            print(f"    ✅ Cột {col_letter} ({col_desc}): Không có ô trống")

# Test function
def test_rule6_validation():
    """Test Rule 6: Item Description = Family validation"""
    print("🧪 TESTING RULE 6: ITEM DESCRIPTION = FAMILY VALIDATION")
    print("=" * 70)
    
    # Test file
    excel_file = r"MEP_Schedule_Table_20250610_154246.xlsx"
    
    if not Path(excel_file).exists():
        print(f"❌ File not found: {excel_file}")
        return
    
    print(f"📁 Testing file: {excel_file}")
    print()
    
    # Run validation
    validator = ExcelValidator()
    output_file = validator.validate_excel_file(excel_file)
    
    if output_file:
        print(f"✅ Rule 6 validation test completed successfully!")
        print(f"📁 Output: {output_file}")
    else:
        print("❌ Rule 6 validation test failed!")

if __name__ == "__main__":
    test_rule6_validation()
