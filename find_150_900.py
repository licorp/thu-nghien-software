#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TÌM DÒNG CÓ 150-900 TRONG FILE
==============================
"""

import pandas as pd

def find_150_900_rows():
    """
    Tìm tất cả dòng có Item Description = 150-900
    """
    print("🔍 TÌM DÒNG CÓ ITEM DESCRIPTION = '150-900'")
    print("=" * 50)
    
    # File Excel gốc
    file = 'MEP_Schedule_Table_20250610_154246.xlsx'
    
    try:
        # Đọc Pipe Schedule worksheet  
        df = pd.read_excel(file, sheet_name='Pipe Schedule')
        
        # Tìm dòng có 150-900
        mask = df['EE_Item Description'].astype(str).str.contains('150-900', na=False)
        found_rows = df[mask]
        
        print(f"📊 TÌM THẤY {len(found_rows)} DÒNG CÓ '150-900':")
        print("-" * 40)
        
        if len(found_rows) == 0:
            print("❌ Không tìm thấy dòng nào có '150-900'")
            
            # Thử tìm pattern khác
            print("\n🔍 Thử tìm pattern '150' và '900':")
            patterns = ['150', '900', '150-', '-900']
            for pattern in patterns:
                mask2 = df['EE_Item Description'].astype(str).str.contains(pattern, na=False)
                found = df[mask2]
                print(f"  Pattern '{pattern}': {len(found)} dòng")
                if len(found) > 0:
                    for idx, row in found.head(5).iterrows():
                        item_desc = row.get('EE_Item Description', 'N/A')
                        print(f"    Dòng {idx+2}: {item_desc}")
        else:
            for idx, row in found_rows.iterrows():
                item_desc = row.get('EE_Item Description', 'N/A')
                size = row.get('EE_Size', 'N/A')
                fab_pipe = row.get('EE_FAB Pipe', 'N/A')
                end_1 = row.get('EE_End-1', 'N/A')
                end_2 = row.get('EE_End-2', 'N/A')
                
                print(f"  Dòng {idx+2}: {item_desc}")
                print(f"    Size={size} | FAB={fab_pipe} | End1={end_1} | End2={end_2}")
                print()
                
    except Exception as e:
        print(f"❌ Lỗi đọc file: {e}")

if __name__ == "__main__":
    find_150_900_rows()
