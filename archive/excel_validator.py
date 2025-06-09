import pandas as pd
import openpyxl
from openpyxl import load_workbook
import os

def validate_excel_data(excel_file_path, output_file_path=None):
    """
    Đọc file Excel và thêm cột validation ở cột N
    """
    try:
        # Đọc file Excel
        df = pd.read_excel(excel_file_path)
        
        print("Dữ liệu hiện tại:")
        print(df.head())
        print(f"\nCác cột hiện có: {list(df.columns)}")
        print(f"Số dòng: {len(df)}")
          # Tạo cột check validation ở cột N
        # Áp dụng điều kiện check cho từng dòng
        df['Validation_Check'] = df.apply(lambda row: check_conditions(row), axis=1)
        
        # Xuất file kết quả
        if output_file_path is None:
            output_file_path = excel_file_path.replace('.xlsx', '_with_validation.xlsx')
        
        df.to_excel(output_file_path, index=False)
        print(f"\nFile đã được lưu: {output_file_path}")
        
        return df
        
    except Exception as e:
        print(f"Lỗi khi xử lý file: {str(e)}")
        return None

def check_conditions(row):
    """
    Hàm kiểm tra điều kiện validation dựa trên business logic từ hình ảnh
    """    try:
        # Lấy giá trị các cột cần kiểm tra
        item_desc = str(row.get('EE_Item Description', '')).strip() if pd.notna(row.get('EE_Item Description')) else ''
        size = row.get('Size', '')
        length = row.get('Length', '')  # Thêm cột Length (cột J)
        fab_pipe = str(row.get('EE_FAB Pipe', '')).strip() if pd.notna(row.get('EE_FAB Pipe')) else ''
        pipe_end1 = str(row.get('EE_PIPE END-1', '')).strip() if pd.notna(row.get('EE_PIPE END-1')) else ''
        pipe_end2 = str(row.get('EE_PIPE END-2', '')).strip() if pd.notna(row.get('EE_PIPE END-2')) else ''
        check_status = row.get('Check', False)
        ghi_chu = str(row.get('Ghi chú', '')).strip() if pd.notna(row.get('Ghi chú')) else ''
        
        # Danh sách lỗi tìm thấy
        errors = []
        
        # ĐIỀU KIỆN 1: Kiểm tra Groove_Thread - END-1 và END-2 phải giống nhau
        if 'Groove_Thread' in fab_pipe:
            if pipe_end1 != pipe_end2:
                errors.append(f"Groove_Thread: END-1({pipe_end1}) ≠ END-2({pipe_end2})")
        
        # ĐIỀU KIỆN 2: Kiểm tra STD PAP RANGE - END-1 phải là RG, END-2 phải là BE
        if 'STD' in fab_pipe and 'PAP RANGE' in fab_pipe:
            if pipe_end1 != 'RG':
                errors.append(f"STD PAP RANGE: END-1 phải là RG, hiện tại là {pipe_end1}")
            if pipe_end2 != 'BE':
                errors.append(f"STD PAP RANGE: END-2 phải là BE, hiện tại là {pipe_end2}")
        
        # ĐIỀU KIỆN 3: Kiểm tra STD ARRAY TEE - END-1 và END-2 phải là RG
        if 'STD ARRAY TEE' in fab_pipe:
            if pipe_end1 != 'RG':
                errors.append(f"STD ARRAY TEE: END-1 phải là RG, hiện tại là {pipe_end1}")
            if pipe_end2 != 'RG':
                errors.append(f"STD ARRAY TEE: END-2 phải là RG, hiện tại là {pipe_end2}")
        
        # ĐIỀU KIỆN 4: Kiểm tra Fabrication - END-1 phải là RG, END-2 phải là BE
        if 'Fabrication' in fab_pipe:
            if pipe_end1 != 'RG':
                errors.append(f"Fabrication: END-1 phải là RG, hiện tại là {pipe_end1}")
            if pipe_end2 != 'BE':
                errors.append(f"Fabrication: END-2 phải là BE, hiện tại là {pipe_end2}")
            # Fabrication phải có ghi chú "không tâm Cốt G"
            if 'không tâm Cốt G' not in ghi_chu:
                errors.append("Fabrication: thiếu ghi chú 'không tâm Cốt G'")
        
        # ĐIỀU KIỆN 5: Kiểm tra Size hợp lệ
        if pd.isna(size) or size == '' or (isinstance(size, (int, float)) and size <= 0):
            errors.append("Size không hợp lệ hoặc ≤ 0")
        
        # ĐIỀU KIỆN 6: Kiểm tra các trường bắt buộc
        if not pipe_end1 or pipe_end1 == '':
            errors.append("EE_PIPE END-1 không được trống")
        if not pipe_end2 or pipe_end2 == '':
            errors.append("EE_PIPE END-2 không được trống")
        if not fab_pipe or fab_pipe == '':
            errors.append("EE_FAB Pipe không được trống")
            
        # ĐIỀU KIỆN 7: Check status phải là TRUE
        if not check_status:
            errors.append("Check status phải là TRUE")
          # ĐIỀU KIỆN 8: Kiểm tra Groove có ghi chú "không tâm Cốt G"
        if 'Groove' in fab_pipe and 'không tâm Cốt G' not in ghi_chu:
            errors.append("Groove cần có ghi chú 'không tâm Cốt G'")
          # ĐIỀU KIỆN 9: Kiểm tra EE_Item Description = Size + "-" + Length (làm tròn 5)
        # Lấy giá trị Length từ cột J
        length = row.get('Length', '')
        if pd.notna(length) and pd.notna(size) and length != '' and size != '':
            try:
                # Làm tròn Length với bội số của 5
                length_rounded = round(float(length) / 5) * 5
                # Tạo expected value: Size + "-" + Length_rounded
                expected_item_desc = f"{int(size)}-{int(length_rounded)}"
                
                # So sánh với EE_Item Description thực tế
                if item_desc != expected_item_desc:
                    errors.append(f"Item Description: mong đợi '{expected_item_desc}', có '{item_desc}'")
            except (ValueError, TypeError):
                errors.append("Không thể tính toán Item Description (Size hoặc Length không hợp lệ)")
        
        # ĐIỀU KIỆN 10: Kiểm tra EE_Array Number = "EXP6" + 2 số cuối cột B + 3 số cuối cột A
        cross_passage = row.get('EE_Cross Passage', '')  # Cột A
        location_lanes = row.get('EE_Location and Lanes', '')  # Cột B  
        array_number = row.get('EE_Array Number', '')  # Cột D
        
        if pd.notna(cross_passage) and pd.notna(location_lanes) and pd.notna(array_number):
            try:
                import re
                # Lấy 2 số cuối của cột B (EE_Location and Lanes)
                location_str = str(location_lanes).strip()
                numbers_in_location = re.findall(r'\d+', location_str)
                if numbers_in_location:
                    last_2_b = numbers_in_location[-1][-2:] if len(numbers_in_location[-1]) >= 2 else numbers_in_location[-1]
                else:
                    last_2_b = "00"
                
                # Lấy 3 số cuối của cột A (EE_Cross Passage)
                cross_str = str(cross_passage).strip()
                numbers_in_cross = re.findall(r'\d+', cross_str)
                if numbers_in_cross:
                    last_3_a = numbers_in_cross[-1][-3:] if len(numbers_in_cross[-1]) >= 3 else numbers_in_cross[-1].zfill(3)
                else:
                    last_3_a = "000"
                
                # Tạo expected value
                expected_array = f"EXP6{last_2_b}{last_3_a}"
                actual_array = str(array_number).strip()
                
                if actual_array != expected_array:
                    errors.append(f"Array Number: mong đợi '{expected_array}', có '{actual_array}'")
                    
            except Exception as e:
                errors.append(f"Không thể tính Array Number: {str(e)}")
            
        # Trả về kết quả
        if errors:
            return f"FAIL: {'; '.join(errors)}"
        else:
            return "PASS"
            
    except Exception as e:
        return f"ERROR: {str(e)}"

if __name__ == "__main__":
    # Đường dẫn file Excel của bạn
    excel_file = input("Nhập đường dẫn file Excel: ")
    
    if os.path.exists(excel_file):
        result = validate_excel_data(excel_file)
        if result is not None:
            print("Xử lý thành công!")
    else:
        print("File không tồn tại!")
