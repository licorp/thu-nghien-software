import pandas as pd

# Đọc file Excel để kiểm tra điều kiện mới
df = pd.read_excel('Xp02-Fabrication & Listing.xlsx')

print("=== KIỂM TRA ĐIỀU KIỆN ITEM DESCRIPTION ===")
print("Cột F: EE_Item Description")
print("Cột G: Size") 
print("Cột J: Length")
print("Công thức: EE_Item Description = Size + '-' + round(Length/5)*5")
print()

# Kiểm tra một vài dòng đầu
for i in range(5):
    row = df.iloc[i]
    item_desc = row.get('EE_Item Description', '')
    size = row.get('Size', '')
    length = row.get('Length', '')
    
    if pd.notna(length) and pd.notna(size):
        length_rounded = round(float(length) / 5) * 5
        expected = f"{int(size)}-{int(length_rounded)}"
        
        status = "✅ PASS" if str(item_desc) == expected else "❌ FAIL"
        
        print(f"Dòng {i+2:2d}: {item_desc:10s} = {size} + '-' + {length_rounded} = {expected:10s} | {status}")
        if str(item_desc) != expected:
            print(f"        Chi tiết: Length gốc = {length}, làm tròn = {length_rounded}")
    print()

print("\n=== THỐNG KÊ ĐIỀU KIỆN ITEM DESCRIPTION ===")
# Đếm số lỗi về Item Description
fail_count = 0
total_count = 0

for i in range(len(df)):
    row = df.iloc[i]
    item_desc = str(row.get('EE_Item Description', '')).strip()
    size = row.get('Size', '')
    length = row.get('Length', '')
    
    if pd.notna(length) and pd.notna(size) and length != '' and size != '':
        total_count += 1
        try:
            length_rounded = round(float(length) / 5) * 5
            expected = f"{int(size)}-{int(length_rounded)}"
            if item_desc != expected:
                fail_count += 1
        except:
            fail_count += 1

print(f"Tổng số dòng có đủ dữ liệu: {total_count}")
print(f"Số dòng FAIL điều kiện Item Description: {fail_count}")
print(f"Số dòng PASS điều kiện Item Description: {total_count - fail_count}")
print(f"Tỷ lệ PASS: {(total_count - fail_count)/total_count*100:.1f}%")
