#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from production.excel_validator_detailed import ExcelValidator
import pandas as pd

def demo_error_display():
    """
    Demo chức năng hiển thị lỗi mới
    """
    print("🚀 DEMO CHỨC NĂNG HIỂN THỊ LỖI MỚI")
    print("=" * 50)
    
    # Tạo fake data để test
    fake_data = {
        'EE_Cross Passage': ['CP01', 'CP02', 'CP03'] * 10,
        'EE_Location': ['Lane1', 'Lane2', 'Lane3'] * 10, 
        'EE_System Type': ['CP-INTERNAL'] * 30,
        'EE_Array Number': ['EXP61001', 'WRONG', 'EXP61003'] * 10,
        **{f'Col_{i}': [f'Data_{i}'] * 30 for i in range(5, 20)},
        'EE_Pipe Treatment': ['GAL'] * 30
    }
    
    df = pd.DataFrame(fake_data)
    
    # Tạo validation check giả
    df['Validation_Check'] = ['PASS', 'FAIL: Array error', 'FAIL: Treatment error'] * 10
    
    # Test hàm hiển thị lỗi
    validator = ExcelValidator()
    validator._show_sample_errors(df, 'Test Sheet', 'EE_System Type', 'EE_Array Number', 'EE_Pipe Treatment')

if __name__ == "__main__":
    demo_error_display()
