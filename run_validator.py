import pandas as pd
import os
from pathlib import Path

def main():
    print("=== EXCEL DATA VALIDATOR ===")
    print("Tìm kiếm file Excel trong thư mục hiện tại...")
    
    # Tìm tất cả file Excel trong thư mục hiện tại
    current_dir = Path(".")
    excel_files = list(current_dir.glob("*.xlsx")) + list(current_dir.glob("*.xls"))
    
    if not excel_files:
        print("Không tìm thấy file Excel nào!")
        print("Hãy copy file Excel vào thư mục:", os.getcwd())
        return
    
    # Hiển thị danh sách file
    print("\nFile Excel tìm thấy:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {file.name}")
    
    # Chọn file
    try:
        choice = int(input(f"\nChọn file (1-{len(excel_files)}): ")) - 1
        selected_file = excel_files[choice]
    except (ValueError, IndexError):
        print("Lựa chọn không hợp lệ!")
        return
    
    # Đọc và hiển thị dữ liệu
    try:
        print(f"\nĐang đọc file: {selected_file.name}")
        df = pd.read_excel(selected_file)
        
        print(f"\nThông tin file:")
        print(f"- Số dòng: {len(df)}")
        print(f"- Số cột: {len(df.columns)}")
        print(f"- Tên các cột: {list(df.columns)}")
        
        print(f"\nDữ liệu mẫu (5 dòng đầu):")
        print(df.head())
        
        # Tạo cột validation (tạm thời để PASS tất cả)
        df['Validation_Check'] = 'PASS'
        
        # Xuất file mới
        output_file = selected_file.stem + "_validated.xlsx"
        df.to_excel(output_file, index=False)
        
        print(f"\n✅ Đã tạo file mới: {output_file}")
        print("Cột 'Validation_Check' đã được thêm vào!")
        
    except Exception as e:
        print(f"❌ Lỗi: {e}")

if __name__ == "__main__":
    main()
