import pandas as pd
import numpy as np

# Kiểm tra file validation result mới nhất
result_file = r"d:\OneDrive\Desktop\thu nghien software\validation_4rules_MEP_Schedule_Table_20250610_154246_20250610_170036.xlsx"

try:
    # Đọc worksheet Pipe Accessory Schedule (nơi có lỗi)
    df = pd.read_excel(result_file, sheet_name='Pipe Accessory Schedule', engine='openpyxl')
    
    print(f"=== WORKSHEET: Pipe Accessory Schedule ===")
    print(f"Số dòng: {len(df)}")
    print(f"Tên các cột:")
    for i, col in enumerate(df.columns):
        print(f"  {i+1}. {col}")
    
    # Kiểm tra dòng 393 (index 392)
    if len(df) >= 393:
        print(f"\n=== DÒNG 393 (dòng cuối - có thể là dòng lỗi) ===")
        row_393 = df.iloc[392]  # index 392 = dòng 393
        
        for col in df.columns:
            print(f"{col}: '{row_393[col]}'")
    
    # Tìm cột validation
    validation_cols = [col for col in df.columns if 'validation' in col.lower() or 'check' in col.lower()]
    if validation_cols:
        validation_col = validation_cols[0]
        print(f"\n=== KIỂM TRA CỘT VALIDATION: {validation_col} ===")
        
        # Đếm PASS/FAIL
        pass_count = df[validation_col].str.contains('PASS', na=False).sum()
        fail_count = len(df) - pass_count
        
        print(f"✅ PASS: {pass_count}")
        print(f"❌ FAIL: {fail_count}")
        
        # Hiển thị các dòng FAIL
        if fail_count > 0:
            fail_mask = ~df[validation_col].str.contains('PASS', na=False) | df[validation_col].isna()
            fail_rows = df[fail_mask]
            print(f"\n=== DÒNG FAIL ===")
            for idx, row in fail_rows.iterrows():
                print(f"Dòng {idx+1}: Validation = '{row[validation_col]}'")
                
except Exception as e:
    print(f"Lỗi: {e}")
    import traceback
    traceback.print_exc()
