import pandas as pd
import os
from pathlib import Path
from datetime import datetime

def validate_actual_excel(excel_file_path):
    """
    Validate Excel file thực tế với cấu trúc cột có sẵn
    """
    try:
        # Đọc file Excel
        df = pd.read_excel(excel_file_path)
        
        print("=== PHÂN TÍCH FILE EXCEL ===")
        print(f"File: {excel_file_path}")
        print(f"Số dòng: {len(df)}")
        print(f"Số cột: {len(df.columns)}")
        
        # Hiển thị cấu trúc dữ liệu
        print("\nCấu trúc dữ liệu:")
        for i, col in enumerate(df.columns):
            sample_data = df[col].dropna().head(2).tolist()
            print(f"{i+1:2d}. {col:25s} | Mẫu: {sample_data}")
        
        # Kiểm tra các cột quan trọng có tồn tại không
        key_columns = {
            'EE_Item Description': 'EE_Item Description',
            'Size': 'Size', 
            'EE_FAB Pipe': 'EE_FAB Pipe',
            'EE_PIPE END-1': 'EE_PIPE END-1',
            'EE_PIPE END-2': 'EE_PIPE END-2'
        }
        
        missing_cols = []
        for key, col_name in key_columns.items():
            if col_name not in df.columns:
                missing_cols.append(col_name)
        
        if missing_cols:
            print(f"\n❌ Thiếu cột quan trọng: {missing_cols}")
            return None
        
        print(f"\n✅ Tìm thấy tất cả cột cần thiết!")
        
        # Áp dụng validation
        print("\n=== BẮT ĐẦU VALIDATION ===")
        df['Validation_Check'] = df.apply(lambda row: validate_row_conditions(row), axis=1)
        
        # Thống kê
        total_rows = len(df)
        pass_rows = df[df['Validation_Check'] == 'PASS']
        fail_rows = df[df['Validation_Check'] != 'PASS']
        
        print(f"\n=== KẾT QUẢ VALIDATION ===")
        print(f"✅ PASS: {len(pass_rows)}/{total_rows} ({len(pass_rows)/total_rows*100:.1f}%)")
        print(f"❌ FAIL: {len(fail_rows)}/{total_rows} ({len(fail_rows)/total_rows*100:.1f}%)")
        
        # Hiển thị một số lỗi mẫu
        if not fail_rows.empty:
            print(f"\nCác lỗi phổ biến:")
            for idx, row in fail_rows.head(5).iterrows():
                item_desc = row.get('EE_Item Description', 'N/A')
                fab_pipe = row.get('EE_FAB Pipe', 'N/A')
                validation_result = row['Validation_Check']
                print(f"Dòng {idx+2:3d}: {item_desc} | {fab_pipe} | {validation_result}")
        
        # Xuất file kết quả
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"validation_result_{timestamp}.xlsx"
        df.to_excel(output_file, index=False)
        print(f"\n📁 File kết quả đã lưu: {output_file}")
        
        return df
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        return None

def validate_row_conditions(row):
    """
    Validate từng dòng dữ liệu theo business rules
    """
    try:
        # Lấy dữ liệu từ các cột
        item_desc = str(row.get('EE_Item Description', '')).strip() if pd.notna(row.get('EE_Item Description')) else ''
        size = row.get('Size', '')
        fab_pipe = str(row.get('EE_FAB Pipe', '')).strip() if pd.notna(row.get('EE_FAB Pipe')) else ''
        pipe_end1 = str(row.get('EE_PIPE END-1', '')).strip() if pd.notna(row.get('EE_PIPE END-1')) else ''
        pipe_end2 = str(row.get('EE_PIPE END-2', '')).strip() if pd.notna(row.get('EE_PIPE END-2')) else ''
        
        errors = []
        
        # Rule 1: Kiểm tra các trường bắt buộc
        if not fab_pipe:
            errors.append("EE_FAB Pipe trống")
        if not pipe_end1:
            errors.append("EE_PIPE END-1 trống") 
        if not pipe_end2:
            errors.append("EE_PIPE END-2 trống")
        
        # Rule 2: Kiểm tra Size hợp lệ
        if pd.isna(size) or (isinstance(size, str) and size.strip() == ''):
            errors.append("Size trống")
        elif isinstance(size, (int, float)) and size <= 0:
            errors.append("Size ≤ 0")
        
        # Rule 3: Business logic cho Groove_Thread
        if 'Groove_Thread' in fab_pipe:
            if pipe_end1 != pipe_end2:
                errors.append(f"Groove_Thread: END-1({pipe_end1}) ≠ END-2({pipe_end2})")
        
        # Rule 4: Business logic cho STD PAP RANGE
        if 'STD' in fab_pipe and 'PAP RANGE' in fab_pipe:
            if pipe_end1 != 'RG':
                errors.append(f"STD PAP RANGE: END-1 cần RG, có {pipe_end1}")
            if pipe_end2 != 'BE':
                errors.append(f"STD PAP RANGE: END-2 cần BE, có {pipe_end2}")
        
        # Rule 5: Business logic cho STD ARRAY TEE
        if 'STD ARRAY TEE' in fab_pipe:
            if pipe_end1 != 'RG' or pipe_end2 != 'RG':
                errors.append(f"STD ARRAY TEE: cần RG-RG, có {pipe_end1}-{pipe_end2}")
        
        # Rule 6: Business logic cho Fabrication
        if 'Fabrication' in fab_pipe:
            if pipe_end1 != 'RG':
                errors.append(f"Fabrication: END-1 cần RG, có {pipe_end1}")
            if pipe_end2 != 'BE':
                errors.append(f"Fabrication: END-2 cần BE, có {pipe_end2}")
          # Rule 7: Kiểm tra kết hợp logic
        # Nếu có các pattern đặc biệt khác có thể thêm ở đây
          # Rule 8: ĐIỀU KIỆN MỚI - Kiểm tra EE_Item Description = Size + "-" + Length (làm tròn 5)
        length = row.get('Length', '')
        if pd.notna(length) and pd.notna(size) and length != '' and size != '':
            try:
                # Làm tròn Length với bội số của 5
                length_rounded = round(float(length) / 5) * 5
                # Tạo expected value: Size + "-" + Length_rounded
                expected_item_desc = f"{int(size)}-{int(length_rounded)}"
                
                # So sánh với EE_Item Description thực tế
                actual_item_desc = str(row.get('EE_Item Description', '')).strip()
                if actual_item_desc != expected_item_desc:
                    errors.append(f"Item Description: cần '{expected_item_desc}', có '{actual_item_desc}'")
            except (ValueError, TypeError):
                errors.append("Không thể tính Item Description (Size/Length lỗi)")        # Rule 9: ĐIỀU KIỆN MỚI - Kiểm tra EE_Array Number chứa "EXP6" + 2 số cuối cột B + 2 số cuối cột A
        cross_passage = row.get('EE_Cross Passage', '')  # Cột A
        location_lanes = row.get('EE_Location and Lanes', '')  # Cột B  
        array_number = row.get('EE_Array Number', '')  # Cột D
        
        if pd.notna(cross_passage) and pd.notna(location_lanes) and pd.notna(array_number):
            try:
                # Lấy 2 số cuối của cột B (EE_Location and Lanes)
                location_str = str(location_lanes).strip()
                # Tìm số trong string, lấy 2 số cuối
                import re
                numbers_in_location = re.findall(r'\d+', location_str)
                if numbers_in_location:
                    last_2_b = numbers_in_location[-1][-2:] if len(numbers_in_location[-1]) >= 2 else numbers_in_location[-1].zfill(2)
                else:
                    last_2_b = "00"
                
                # Lấy 2 số cuối của cột A (EE_Cross Passage) - SỬA ĐỔI TỪ 3 SỐ THÀNH 2 SỐ
                cross_str = str(cross_passage).strip()
                numbers_in_cross = re.findall(r'\d+', cross_str)
                if numbers_in_cross:
                    last_2_a = numbers_in_cross[-1][-2:] if len(numbers_in_cross[-1]) >= 2 else numbers_in_cross[-1].zfill(2)
                else:
                    last_2_a = "00"
                
                # Tạo expected pattern (phải chứa trong array number)
                required_pattern = f"EXP6{last_2_b}{last_2_a}"
                actual_array = str(array_number).strip()
                
                # Kiểm tra xem array number có chứa pattern bắt buộc không
                if required_pattern not in actual_array:
                    errors.append(f"Array Number: phải chứa '{required_pattern}', có '{actual_array}'")
                    
            except Exception as e:
                errors.append(f"Không thể tính Array Number: {str(e)}")
        
        # Trả về kết quả
        if errors:
            return f"FAIL: {'; '.join(errors[:2])}"  # Chỉ hiển thị 2 lỗi đầu để không quá dài
        else:
            return "PASS"
            
    except Exception as e:
        return f"ERROR: {str(e)}"

if __name__ == "__main__":
    # Tự động tìm file Excel
    current_dir = Path(".")
    excel_files = [f for f in current_dir.glob("*.xlsx") if not f.name.startswith('~') and 'validation' not in f.name.lower()]
    
    if not excel_files:
        print("❌ Không tìm thấy file Excel!")
        exit()
    
    print("📁 File Excel có sẵn:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {file.name}")
    
    try:
        choice = int(input(f"\nChọn file (1-{len(excel_files)}): ")) - 1
        selected_file = excel_files[choice]
        validate_actual_excel(selected_file)
    except (ValueError, IndexError):
        print("❌ Lựa chọn không hợp lệ!")
    except KeyboardInterrupt:
        print("\n⏹️ Đã hủy!")
