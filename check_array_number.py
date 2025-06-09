import pandas as pd

# Đọc file Excel để kiểm tra cấu trúc cột A, B, D
df = pd.read_excel('Xp02-Fabrication & Listing.xlsx')

print("=== KIỂM TRA CẤU TRÚC DỮ LIỆU ===")
print("Cột A:", df.columns[0])
print("Cột B:", df.columns[1]) 
print("Cột D:", df.columns[3])
print()

print("=== MẪU DỮ LIỆU ===")
for i in range(5):
    row = df.iloc[i]
    col_a = row.iloc[0]  # Cột A
    col_b = row.iloc[1]  # Cột B  
    col_d = row.iloc[3]  # Cột D
    
    print(f"Dòng {i+2}:")
    print(f"  Cột A ({df.columns[0]}): {col_a}")
    print(f"  Cột B ({df.columns[1]}): {col_b}")
    print(f"  Cột D ({df.columns[3]}): {col_d}")
    
    # Tính toán expected value
    if pd.notna(col_a) and pd.notna(col_b):
        try:
            # Lấy 2 số cuối của cột B
            col_b_str = str(col_b)
            last_2_b = col_b_str[-2:] if len(col_b_str) >= 2 else col_b_str
            
            # Lấy 3 số cuối của cột A  
            col_a_str = str(col_a)
            last_3_a = col_a_str[-3:] if len(col_a_str) >= 3 else col_a_str
            
            expected = f"EXP6{last_2_b}{last_3_a}"
            actual = str(col_d)
            
            status = "✅ PASS" if actual == expected else "❌ FAIL"
            print(f"  Expected: EXP6 + {last_2_b} + {last_3_a} = {expected}")
            print(f"  Actual: {actual}")
            print(f"  Status: {status}")
            
        except Exception as e:
            print(f"  Lỗi: {e}")
    print()
