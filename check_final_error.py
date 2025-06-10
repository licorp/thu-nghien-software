import pandas as pd
import numpy as np

# Kiểm tra file validation result mới nhất
result_file = r"d:\OneDrive\Desktop\thu nghien software\validation_4rules_MEP_Schedule_Table_20250610_154246_20250610_170036.xlsx"

try:
    # Đọc tất cả worksheets
    all_sheets = pd.read_excel(result_file, sheet_name=None, engine='openpyxl')
    
    print("=== KIỂM TRA DÒNG LỖI CUỐI CÙNG ===")
    
    # Tìm dòng lỗi trong từng worksheet
    for sheet_name, df in all_sheets.items():
        print(f"\n📊 WORKSHEET: {sheet_name}")
        print(f"Số dòng: {len(df)}")
        
        # Tìm cột validation
        validation_col = None
        for col in df.columns:
            if 'validation' in col.lower() or 'check' in col.lower():
                validation_col = col
                break
        
        if validation_col:
            # Tìm các dòng FAIL
            fail_rows = df[df[validation_col].str.contains('FAIL|Groove_Thread|Fabrication', na=False, case=False)]
            
            if len(fail_rows) > 0:
                print(f"❌ FAIL: {len(fail_rows)} dòng")
                for idx, row in fail_rows.iterrows():
                    print(f"   Dòng {idx+1}: {row[validation_col]}")
                    
                    # Hiển thị thông tin chi tiết
                    for col in df.columns:
                        if any(keyword in col.upper() for keyword in ['SIZE', 'ITEM', 'FAB', 'END']):
                            print(f"     {col}: {row[col]}")
                    print()
            else:
                print("✅ Tất cả dòng PASS")
        else:
            print("❓ Không tìm thấy cột validation")
            
except Exception as e:
    print(f"Lỗi: {e}")
