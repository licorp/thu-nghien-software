#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

def check_excel_worksheets():
    """
    Kiểm tra tất cả worksheet trong file Excel
    """
    excel_file = 'Xp02-Fabrication & Listing.xlsx'
    
    try:
        # Đọc tất cả worksheet names
        xl_file = pd.ExcelFile(excel_file)
        
        print("=== DANH SÁCH WORKSHEET TRONG FILE EXCEL ===")
        print(f"File: {excel_file}")
        print(f"Tổng số worksheet: {len(xl_file.sheet_names)}")
        print()
        
        for i, sheet_name in enumerate(xl_file.sheet_names, 1):
            print(f"{i:2d}. {sheet_name}")
            
            # Đọc worksheet để xem cấu trúc
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                print(f"    📊 Số dòng: {len(df)}, Số cột: {len(df.columns)}")
                
                # Kiểm tra có các cột cần thiết không
                required_cols = ['EE_Cross Passage', 'EE_Location and Lanes', 'EE_Array Number']
                has_required = all(col in df.columns for col in required_cols)
                
                if has_required:
                    print(f"    ✅ Có đầy đủ cột cần thiết cho Array Number validation")
                else:
                    print(f"    ❌ Thiếu cột cần thiết cho Array Number validation")
                    
            except Exception as e:
                print(f"    ⚠️ Không thể đọc worksheet: {str(e)}")
                
            print()
            
        # Kiểm tra các worksheet mục tiêu
        target_sheets = [
            'Pipe Schedule', 
            'Pipe Fitting Schedule', 
            'Pipe Accessory Schedule', 
            'Sprinkler Schedule'
        ]
        
        print("=== KIỂM TRA WORKSHEET MỤC TIÊU ===")
        for sheet in target_sheets:
            if sheet in xl_file.sheet_names:
                print(f"✅ Tìm thấy: {sheet}")
            else:
                print(f"❌ Không tìm thấy: {sheet}")
                
    except Exception as e:
        print(f"❌ Lỗi đọc file: {e}")

if __name__ == "__main__":
    check_excel_worksheets()
