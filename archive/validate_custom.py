import pandas as pd
import os
from pathlib import Path
from datetime import datetime

def validate_with_custom_conditions(excel_file_path):
    """
    Validate Excel với điều kiện cụ thể từ hình ảnh
    """
    try:
        # Đọc file Excel
        df = pd.read_excel(excel_file_path)
        
        print("=== BẮT ĐẦU VALIDATION ===")
        print(f"File: {excel_file_path}")
        print(f"Số dòng: {len(df)}")
        print(f"Các cột: {list(df.columns)}")
        
        # Tìm các cột cần thiết
        required_columns = ['EE_Item Description', 'Size', 'EE_FAB Pipe', 'EE_PIPE END-1', 'EE_PIPE END-2', 'Check', 'Ghi chú']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"❌ Thiếu các cột: {missing_columns}")
            print("Đang tìm cột tương tự...")
            # Hiển thị tất cả cột để user chọn
            for i, col in enumerate(df.columns):
                print(f"{i+1}. {col}")
            return None
        
        # Áp dụng validation
        df['Validation_Result'] = df.apply(lambda row: check_validation_conditions(row), axis=1)
        
        # Tạo file kết quả với timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"validation_result_{timestamp}.xlsx"
        
        # Xuất kết quả
        df.to_excel(output_file, index=False)
        
        # Thống kê kết quả
        total_rows = len(df)
        pass_count = len(df[df['Validation_Result'] == 'PASS'])
        fail_count = total_rows - pass_count
        
        print(f"\n=== KẾT QUẢ VALIDATION ===")
        print(f"✅ PASS: {pass_count}/{total_rows} ({pass_count/total_rows*100:.1f}%)")
        print(f"❌ FAIL: {fail_count}/{total_rows} ({fail_count/total_rows*100:.1f}%)")
        print(f"📁 File kết quả: {output_file}")
        
        # Hiển thị một số lỗi mẫu
        failed_rows = df[df['Validation_Result'] != 'PASS']
        if not failed_rows.empty:
            print(f"\nMột số lỗi phổ biến:")
            for idx, row in failed_rows.head(3).iterrows():
                print(f"Dòng {idx+2}: {row['Validation_Result']}")
        
        return df
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return None

def check_validation_conditions(row):
    """
    Kiểm tra điều kiện validation theo business logic
    """
    try:
        # Lấy giá trị
        item_desc = str(row.get('EE_Item Description', '')).strip() if pd.notna(row.get('EE_Item Description')) else ''
        size = row.get('Size', '')
        fab_pipe = str(row.get('EE_FAB Pipe', '')).strip() if pd.notna(row.get('EE_FAB Pipe')) else ''
        pipe_end1 = str(row.get('EE_PIPE END-1', '')).strip() if pd.notna(row.get('EE_PIPE END-1')) else ''
        pipe_end2 = str(row.get('EE_PIPE END-2', '')).strip() if pd.notna(row.get('EE_PIPE END-2')) else ''
        check_status = row.get('Check', False)
        ghi_chu = str(row.get('Ghi chú', '')).strip() if pd.notna(row.get('Ghi chú')) else ''
        
        errors = []
        
        # Rule 1: Groove_Thread - END-1 và END-2 phải giống nhau
        if 'Groove_Thread' in fab_pipe and pipe_end1 != pipe_end2:
            errors.append(f"Groove_Thread: END-1({pipe_end1}) ≠ END-2({pipe_end2})")
        
        # Rule 2: STD PAP RANGE - RG-BE pattern
        if 'STD' in fab_pipe and 'PAP RANGE' in fab_pipe:
            if pipe_end1 != 'RG':
                errors.append(f"STD PAP: END-1 cần RG, có {pipe_end1}")
            if pipe_end2 != 'BE':
                errors.append(f"STD PAP: END-2 cần BE, có {pipe_end2}")
        
        # Rule 3: STD ARRAY TEE - RG-RG pattern  
        if 'STD ARRAY TEE' in fab_pipe:
            if pipe_end1 != 'RG' or pipe_end2 != 'RG':
                errors.append(f"STD ARRAY: cần RG-RG, có {pipe_end1}-{pipe_end2}")
        
        # Rule 4: Fabrication - RG-BE + ghi chú
        if 'Fabrication' in fab_pipe:
            if pipe_end1 != 'RG':
                errors.append(f"Fabrication: END-1 cần RG, có {pipe_end1}")
            if pipe_end2 != 'BE':
                errors.append(f"Fabrication: END-2 cần BE, có {pipe_end2}")
            if 'không tâm Cốt G' not in ghi_chu:
                errors.append("Fabrication: thiếu ghi chú 'không tâm Cốt G'")
        
        # Rule 5: Groove cần ghi chú
        if 'Groove' in fab_pipe and 'không tâm Cốt G' not in ghi_chu:
            errors.append("Groove: cần ghi chú 'không tâm Cốt G'")
        
        # Rule 6: Check phải TRUE
        if not check_status:
            errors.append("Check ≠ TRUE")
            
        # Rule 7: Size validation
        if pd.isna(size) or (isinstance(size, str) and size.strip() == '') or (isinstance(size, (int, float)) and size <= 0):
            errors.append("Size không hợp lệ")
        
        return "PASS" if not errors else f"FAIL: {'; '.join(errors)}"
        
    except Exception as e:
        return f"ERROR: {str(e)}"

if __name__ == "__main__":
    # Tìm file Excel
    current_dir = Path(".")
    excel_files = list(current_dir.glob("*.xlsx")) + list(current_dir.glob("*.xls"))
    excel_files = [f for f in excel_files if not f.name.startswith('~') and 'validation' not in f.name.lower()]
    
    if not excel_files:
        print("Không tìm thấy file Excel!")
        exit()
    
    print("File Excel có sẵn:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {file.name}")
    
    try:
        choice = int(input(f"Chọn file (1-{len(excel_files)}): ")) - 1
        selected_file = excel_files[choice]
        validate_with_custom_conditions(selected_file)
    except (ValueError, IndexError):
        print("Lựa chọn không hợp lệ!")
