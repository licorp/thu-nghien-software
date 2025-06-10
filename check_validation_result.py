import pandas as pd
import numpy as np
import openpyxl

# Kiểm tra file validation result mới nhất
result_file = r"d:\OneDrive\Desktop\thu nghien software\validation_4rules_MEP_Schedule_Table_20250610_154246_20250610_164604.xlsx"

try:
    # Đọc file validation result
    df_result = pd.read_excel(result_file, engine='openpyxl')
    
    print(f"=== KIỂM TRA DÒNG 27 TRONG FILE VALIDATION RESULT ===")
    
    # Kiểm tra dòng 27 (index 26)
    row_27 = df_result.iloc[26]
    
    print(f"SIZE: '{row_27.get('SIZE', 'N/A')}'")
    print(f"ITEM DESCRIPTION: '{row_27.get('ITEM DESCRIPTION', 'N/A')}'")
    print(f"FAB PIPE: '{row_27.get('FAB PIPE', 'N/A')}'")
    print(f"END-1: '{row_27.get('END-1', 'N/A')}'")
    print(f"END-2: '{row_27.get('END-2', 'N/A')}'")
    print(f"VALIDATION RESULT: '{row_27.get('VALIDATION RESULT', 'N/A')}'")
    
    # Đếm các dòng có lỗi "Groove_Thread"
    groove_errors = df_result[df_result['VALIDATION RESULT'].str.contains('Groove_Thread', na=False)]
    print(f"\n=== TỔNG SỐ DÒNG BÁO LỖI 'Groove_Thread': {len(groove_errors)} ===")
    
    for idx, row in groove_errors.iterrows():
        print(f"Dòng {idx+1}: SIZE={row.get('SIZE', 'N/A')}, ITEM_DESC='{str(row.get('ITEM DESCRIPTION', 'N/A'))[:50]}...', END-1={row.get('END-1', 'N/A')}, END-2={row.get('END-2', 'N/A')}")
        print(f"   Error: {row.get('VALIDATION RESULT', 'N/A')}")
        print()
        
except Exception as e:
    print(f"Lỗi đọc file: {e}")
