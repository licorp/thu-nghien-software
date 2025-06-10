import pandas as pd
import numpy as np

# Kiểm tra cấu trúc file validation result
result_file = r"d:\OneDrive\Desktop\thu nghien software\validation_4rules_MEP_Schedule_Table_20250610_154246_20250610_164604.xlsx"

try:
    # Đọc file validation result
    df_result = pd.read_excel(result_file, engine='openpyxl')
    
    print("=== TÊN CÁC CỘT TRONG FILE VALIDATION RESULT ===")
    for i, col in enumerate(df_result.columns):
        print(f"{i+1}. '{col}'")
    
    print(f"\n=== DÒNG 27 (index 26) ===")
    if len(df_result) > 26:
        row_27 = df_result.iloc[26]
        for col in df_result.columns:
            print(f"{col}: '{row_27[col]}'")
    else:
        print("File không có đủ 27 dòng")
        
except Exception as e:
    print(f"Lỗi: {e}")
