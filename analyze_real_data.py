import pandas as pd
import re

def analyze_real_excel_data():
    """
    Phân tích dữ liệu thực tế để hiểu pattern
    """
    try:
        # Đọc file Excel
        df = pd.read_excel("Xp02-Fabrication & Listing.xlsx")
        
        print("=== PHÂN TÍCH PATTERN THỰC TẾ ===\n")
        
        # Lấy 10 dòng đầu để phân tích
        for i in range(10):
            if i >= len(df):
                break
                
            row = df.iloc[i]
            cross_passage = str(row.get('EE_Cross Passage', '')).strip()
            location_lanes = str(row.get('EE_Location and Lanes', '')).strip() 
            array_number = str(row.get('EE_Array Number', '')).strip()
            
            print(f"Dòng {i+1}:")
            print(f"  Cross Passage (A): '{cross_passage}'")
            print(f"  Location Lanes (B): '{location_lanes}'")
            print(f"  Array Number (D): '{array_number}'")
            
            # Phân tích số trong từng cột
            numbers_in_a = re.findall(r'\d+', cross_passage)
            numbers_in_b = re.findall(r'\d+', location_lanes)
            
            print(f"  Số trong A: {numbers_in_a}")
            print(f"  Số trong B: {numbers_in_b}")
            
            if numbers_in_a:
                print(f"  Số cuối trong A: {numbers_in_a[-1]}")
                print(f"  3 ký tự cuối: {numbers_in_a[-1][-3:] if len(numbers_in_a[-1]) >= 3 else numbers_in_a[-1].zfill(3)}")
            
            if numbers_in_b:
                print(f"  Số cuối trong B: {numbers_in_b[-1]}")
                print(f"  2 ký tự cuối: {numbers_in_b[-1][-2:] if len(numbers_in_b[-1]) >= 2 else numbers_in_b[-1].zfill(2)}")
            
            # Thử các pattern khác nhau
            if numbers_in_a and numbers_in_b:
                pattern1 = f"EXP6{numbers_in_b[-1][-2:] if len(numbers_in_b[-1]) >= 2 else numbers_in_b[-1].zfill(2)}{numbers_in_a[-1][-3:] if len(numbers_in_a[-1]) >= 3 else numbers_in_a[-1].zfill(3)}"
                pattern2 = f"EXP6{numbers_in_a[-1][-3:] if len(numbers_in_a[-1]) >= 3 else numbers_in_a[-1].zfill(3)}"
                pattern3 = f"EXP6{numbers_in_a[-1]}"
                
                print(f"  Pattern 1 (B+A): '{pattern1}' - Match: {pattern1 in array_number}")
                print(f"  Pattern 2 (A only): '{pattern2}' - Match: {pattern2 in array_number}")
                print(f"  Pattern 3 (Full A): '{pattern3}' - Match: {pattern3 in array_number}")
            
            print()

if __name__ == "__main__":
    analyze_real_excel_data()
