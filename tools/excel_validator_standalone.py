#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel Validator Tool
Double-click để chạy trên Windows
"""

import pandas as pd
import os
import sys
from pathlib import Path
from datetime import datetime

def main():
    """
    Hàm main để chạy Excel validator
    """
    try:
        print("="*50)
        print("          EXCEL VALIDATOR TOOL")
        print("="*50)
        print()
        
        # Tự động tìm file Excel trong thư mục hiện tại
        current_dir = Path(__file__).parent
        os.chdir(current_dir)  # Đảm bảo working directory đúng
        
        excel_files = [f for f in current_dir.glob("*.xlsx") 
                      if not f.name.startswith('~') 
                      and 'validation' not in f.name.lower()]
        
        if not excel_files:
            print("❌ Không tìm thấy file Excel!")
            input("Nhấn Enter để thoát...")
            return
        
        print("📁 File Excel có sẵn:")
        for i, file in enumerate(excel_files, 1):
            print(f"   {i}. {file.name}")
        print()
        
        # Chọn file
        while True:
            try:
                choice = input(f"Chọn file (1-{len(excel_files)}) hoặc 'q' để thoát: ").strip()
                if choice.lower() == 'q':
                    return
                choice = int(choice) - 1
                if 0 <= choice < len(excel_files):
                    break
                else:
                    print("❌ Số không hợp lệ!")
            except ValueError:
                print("❌ Vui lòng nhập số!")
        
        selected_file = excel_files[choice]
        print(f"\n🔍 Đang xử lý file: {selected_file.name}")
        print("-" * 50)
        
        # Validate file
        result = validate_excel_file(selected_file)
        
        if result:
            print("\n" + "="*50)
            print("✅ XỬ LÝ THÀNH CÔNG!")
            print("="*50)
        else:
            print("\n" + "="*50)
            print("❌ CÓ LỖI XẢY RA!")
            print("="*50)
            
    except KeyboardInterrupt:
        print("\n⏹️ Đã hủy bởi người dùng!")
    except Exception as e:
        print(f"\n❌ Lỗi: {e}")
    finally:
        input("\nNhấn Enter để đóng...")

def validate_excel_file(excel_file_path):
    """
    Validate Excel file và trả về kết quả
    """
    try:
        # Đọc file Excel
        df = pd.read_excel(excel_file_path)
        
        print(f"📊 Thông tin file:")
        print(f"   - Số dòng: {len(df)}")
        print(f"   - Số cột: {len(df.columns)}")
        
        # Kiểm tra các cột cần thiết
        required_columns = ['EE_Item Description', 'Size', 'EE_FAB Pipe', 'EE_PIPE END-1', 'EE_PIPE END-2']
        missing_cols = [col for col in required_columns if col not in df.columns]
        
        if missing_cols:
            print(f"❌ Thiếu cột: {missing_cols}")
            return False
        
        print("✅ Tìm thấy tất cả cột cần thiết!")
        print("\n🔍 Đang thực hiện validation...")
        
        # Áp dụng validation
        df['Validation_Check'] = df.apply(validate_row, axis=1)
        
        # Thống kê kết quả
        total_rows = len(df)
        pass_rows = df[df['Validation_Check'] == 'PASS']
        fail_rows = df[df['Validation_Check'] != 'PASS']
        
        print(f"\n📈 KẾT QUẢ VALIDATION:")
        print(f"   ✅ PASS: {len(pass_rows)}/{total_rows} ({len(pass_rows)/total_rows*100:.1f}%)")
        print(f"   ❌ FAIL: {len(fail_rows)}/{total_rows} ({len(fail_rows)/total_rows*100:.1f}%)")
        
        # Hiển thị một số lỗi mẫu
        if not fail_rows.empty:
            print(f"\n⚠️ Một số lỗi tìm thấy:")
            for idx, row in fail_rows.head(3).iterrows():
                item_desc = str(row.get('EE_Item Description', 'N/A'))[:20]
                fab_pipe = str(row.get('EE_FAB Pipe', 'N/A'))[:20]
                validation_result = str(row['Validation_Check'])[:60]
                print(f"   Dòng {idx+2}: {item_desc} | {validation_result}")
        
        # Xuất file kết quả
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"validation_result_{timestamp}.xlsx"
        df.to_excel(output_file, index=False)
        print(f"\n💾 File kết quả: {output_file}")
        
        return True
        
    except Exception as e:
        print(f"❌ Lỗi khi xử lý: {e}")
        return False

def validate_row(row):
    """
    Validate từng dòng theo business rules
    """
    try:
        # Lấy dữ liệu
        fab_pipe = str(row.get('EE_FAB Pipe', '')).strip() if pd.notna(row.get('EE_FAB Pipe')) else ''
        pipe_end1 = str(row.get('EE_PIPE END-1', '')).strip() if pd.notna(row.get('EE_PIPE END-1')) else ''
        pipe_end2 = str(row.get('EE_PIPE END-2', '')).strip() if pd.notna(row.get('EE_PIPE END-2')) else ''
        size = row.get('Size', '')
        
        errors = []
        
        # Kiểm tra trường bắt buộc
        if not fab_pipe:
            errors.append("FAB Pipe trống")
        if not pipe_end1:
            errors.append("END-1 trống") 
        if not pipe_end2:
            errors.append("END-2 trống")
        
        # Kiểm tra Size
        if pd.isna(size) or (isinstance(size, (int, float)) and size <= 0):
            errors.append("Size không hợp lệ")
        
        # Business rules
        if 'Groove_Thread' in fab_pipe and pipe_end1 != pipe_end2:
            errors.append(f"Groove_Thread: {pipe_end1}≠{pipe_end2}")
        
        if 'STD' in fab_pipe and 'PAP RANGE' in fab_pipe:
            if pipe_end1 != 'RG' or pipe_end2 != 'BE':
                errors.append(f"STD PAP: cần RG-BE, có {pipe_end1}-{pipe_end2}")
        
        if 'STD ARRAY TEE' in fab_pipe and (pipe_end1 != 'RG' or pipe_end2 != 'RG'):
            errors.append(f"STD ARRAY: cần RG-RG, có {pipe_end1}-{pipe_end2}")
        
        if 'Fabrication' in fab_pipe and (pipe_end1 != 'RG' or pipe_end2 != 'BE'):
            errors.append(f"Fabrication: cần RG-BE, có {pipe_end1}-{pipe_end2}")
        
        return "PASS" if not errors else f"FAIL: {'; '.join(errors[:2])}"
        
    except Exception as e:
        return f"ERROR: {str(e)}"

if __name__ == "__main__":
    main()
