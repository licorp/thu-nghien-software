#!/usr/bin/env python3
"""
Completely isolated test for Rule 6 validation
No imports from other modules
"""

try:
    import pandas as pd
    import sys
    print("🧪 ISOLATED RULE 6 TEST")
    print("=" * 30)
    
    # Test file
    excel_file = "MEP_Schedule_Table_20250610_154246.xlsx"
    
    # Read Excel file
    xl_file = pd.ExcelFile(excel_file)
    print(f"📁 File: {excel_file}")
    print(f"📊 Worksheets: {len(xl_file.sheet_names)}")
    
    # Check for Pipe Accessory Schedule
    target_sheet = "Pipe Accessory Schedule"
    if target_sheet in xl_file.sheet_names:
        print(f"✅ Found {target_sheet}")
        
        # Read the worksheet
        df = pd.read_excel(excel_file, sheet_name=target_sheet)
        print(f"📋 Rows: {len(df)}, Columns: {len(df.columns)}")
        
        # Check if we have enough columns for F and U
        if len(df.columns) > 20:
            col_f_name = df.columns[5]   # Column F (0-based index 5)
            col_u_name = df.columns[20]  # Column U (0-based index 20)
            
            print(f"🔍 Column F (6th): '{col_f_name}'")
            print(f"🔍 Column U (21st): '{col_u_name}'")
            
            # Quick validation check
            matches = 0
            fails = 0
            
            for i, row in df.head(20).iterrows():  # Test first 20 rows
                f_val = str(row[col_f_name]).strip() if not pd.isna(row[col_f_name]) else ""
                u_val = str(row[col_u_name]).strip() if not pd.isna(row[col_u_name]) else ""
                
                if (f_val == "" and u_val == "") or (f_val == u_val):
                    matches += 1
                else:
                    fails += 1
                    if fails <= 3:  # Show first 3 failures
                        print(f"  ❌ Row {i+2}: F='{f_val[:30]}...' vs U='{u_val[:30]}...'")
            
            print(f"📊 Sample Results (first 20 rows):")
            print(f"✅ Pass: {matches}/20")
            print(f"❌ Fail: {fails}/20")
        else:
            print(f"❌ Only {len(df.columns)} columns, need at least 21 for column U")
    else:
        print(f"❌ {target_sheet} not found")
        print(f"Available: {xl_file.sheet_names}")
        
except Exception as e:
    print(f"❌ Error: {e}")
    import traceback
    traceback.print_exc()

print("✅ Test completed")
