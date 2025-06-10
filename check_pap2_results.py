#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KIỂM TRA KẾT QUẢ STD 2 PAP RANGE
================================

Kiểm tra xem case 65-5295 có được xử lý đúng logic ưu tiên không
"""

import pandas as pd

def check_pap2_results():
    """
    Kiểm tra kết quả STD 2 PAP RANGE trong file validation mới nhất
    """
    print("🔍 KIỂM TRA KẾT QUẢ STD 2 PAP RANGE")
    print("=" * 50)
    
    # File kết quả mới nhất
    file = 'validation_4rules_MEP_Schedule_Table_20250610_154246_20250610_164009.xlsx'
    
    try:
        # Đọc Pipe Schedule worksheet  
        df = pd.read_excel(file, sheet_name='Pipe Schedule')
        print(f"📊 Pipe Schedule: {len(df)} dòng")
        
        # Tìm dòng có 65-5295 (STD 2 PAP RANGE case)
        mask = df['EE_Item Description'].astype(str).str.contains('65-5295', na=False)
        pap2_rows = df[mask]
        
        print(f"\n🎯 TÌM THẤY {len(pap2_rows)} DÒNG CÓ '65-5295':")
        print("-" * 40)
        
        if len(pap2_rows) == 0:
            print("❌ Không tìm thấy dòng nào có '65-5295'")
            
            # Thử tìm với pattern khác
            print("\n🔍 Thử tìm các pattern khác:")
            patterns = ['5295', '65-', 'PAP']
            for pattern in patterns:
                mask2 = df['EE_Item Description'].astype(str).str.contains(pattern, na=False)
                found = df[mask2]
                print(f"  Pattern '{pattern}': {len(found)} dòng")
                if len(found) > 0:
                    for idx, row in found.head(3).iterrows():
                        item_desc = row.get('EE_Item Description', 'N/A')
                        print(f"    Dòng {idx+2}: {item_desc}")
        else:
            for idx, row in pap2_rows.iterrows():
                result = row.get('Validation_Check', 'N/A')
                size = row.get('EE_Size', 'N/A') 
                fab_pipe = row.get('EE_FAB Pipe', 'N/A')
                end1 = row.get('EE_End-1', 'N/A')
                end2 = row.get('EE_End-2', 'N/A')
                item_desc = row.get('EE_Item Description', 'N/A')
                
                print(f"  Dòng {idx+2}: {item_desc}")
                print(f"    Size={size} | FAB={fab_pipe} | End1={end1} | End2={end2}")
                if result == 'PASS':
                    print(f"    ✅ PASS: Logic ưu tiên hoạt động đúng!")
                else:
                    print(f"    ❌ FAIL: {result}")
                print()
                
    except Exception as e:
        print(f"❌ Lỗi đọc file: {e}")
        
    # Kiểm tra các file validation khác
    print("\n📁 KIỂM TRA CÁC FILE VALIDATION KHÁC:")
    print("-" * 40)
    import glob
    validation_files = glob.glob("validation_4rules_*.xlsx")
    for file in sorted(validation_files):
        print(f"  {file}")

if __name__ == "__main__":
    check_pap2_results()
