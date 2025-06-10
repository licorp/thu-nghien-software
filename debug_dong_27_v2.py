import pandas as pd
import numpy as np
import openpyxl

# Debug dòng 27 để hiểu tại sao lại ra "Groove_Thread" thay vì "STD ARRAY TEE"
input_file = r"d:\OneDrive\Desktop\thu nghien software\MEP_Schedule_Table_20250610_154246.xlsx"

# Đọc file Excel
df = pd.read_excel(input_file, engine='openpyxl')

# Debug dòng 27 (index 26)
row_idx = 26
row = df.iloc[row_idx]

print(f"=== DEBUG DÒNG 27 (index {row_idx}) ===")
print(f"SIZE: '{row.get('SIZE', 'N/A')}'")
print(f"ITEM DESCRIPTION: '{row.get('ITEM DESCRIPTION', 'N/A')}'")
print(f"FAB PIPE: '{row.get('FAB PIPE', 'N/A')}'") 
print(f"END-1: '{row.get('END-1', 'N/A')}'")
print(f"END-2: '{row.get('END-2', 'N/A')}'")

# Tạo string versions
size_val = row.get('SIZE', np.nan)
item_desc = row.get('ITEM DESCRIPTION', np.nan)
fab_pipe = row.get('FAB PIPE', np.nan)
end_1 = row.get('END-1', np.nan)
end_2 = row.get('END-2', np.nan)

size_str = "N/A" if pd.isna(size_val) else str(size_val).strip()
item_desc_str = "N/A" if pd.isna(item_desc) else str(item_desc).strip()
fab_pipe_str = "N/A" if pd.isna(fab_pipe) else str(fab_pipe).strip()
end_1_str = "N/A" if pd.isna(end_1) else str(end_1).strip()
end_2_str = "N/A" if pd.isna(end_2) else str(end_2).strip()

print(f"\n=== PROCESSED VALUES ===")
print(f"size_str: '{size_str}'")
print(f"item_desc_str: '{item_desc_str}'")
print(f"fab_pipe_str: '{fab_pipe_str}'")
print(f"end_1_str: '{end_1_str}'")
print(f"end_2_str: '{end_2_str}'")

print(f"\n=== VALIDATION CHECKS ===")

# Kiểm tra các rule theo thứ tự
print(f"1. STD 1 PAP RANGE check:")
print(f"   - size_str == '65': {size_str == '65'}")
print(f"   - '4730' in item_desc_str: {'4730' in item_desc_str}")

print(f"2. STD 2 PAP RANGE check:")
print(f"   - size_str == '65': {size_str == '65'}")
print(f"   - '5295' in item_desc_str: {'5295' in item_desc_str}")

print(f"3. STD ARRAY TEE check:")
print(f"   - size_str == '150': {size_str == '150'}")
print(f"   - '900' in item_desc_str: {'900' in item_desc_str}")
print(f"   - '150-900' in item_desc_str: {'150-900' in item_desc_str}")
print(f"   - MATCH: {((size_str == '150.0' or size_str == '150') and '900' in item_desc_str) or '150-900' in item_desc_str}")

print(f"4. Groove_Thread check:")
print(f"   - end_1_str == 'RG' and end_2_str == 'RG': {end_1_str == 'RG' and end_2_str == 'RG'}")
print(f"   - size_str == '40' and TH TH: {size_str == '40' and end_1_str == 'TH' and end_2_str == 'TH'}")
print(f"   - MATCH: {((end_1_str == 'RG' and end_2_str == 'RG') or (size_str == '40' and end_1_str == 'TH' and end_2_str == 'TH'))}")

print(f"\n=== CONCLUSION ===")
if "150-900" in item_desc_str:
    print("✅ Dòng này PHẢI là STD ARRAY TEE vì có '150-900' trong ITEM DESCRIPTION")
elif (size_str == "150" and "900" in item_desc_str):
    print("✅ Dòng này PHẢI là STD ARRAY TEE vì size=150 và có '900'")
elif end_1_str == "RG" and end_2_str == "RG":
    print("❌ Dòng này đang bị fallback thành Groove_Thread vì RG-RG")
else:
    print("❓ Không rõ tại sao lại báo lỗi")
