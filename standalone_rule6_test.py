#!/usr/bin/env python3
"""
Standalone Rule 6 Validation Test
"""

import pandas as pd
from pathlib import Path

def test_rule6():
    """Test Rule 6: Item Description = Family validation"""
    print("🧪 RULE 6 VALIDATION TEST")
    print("=" * 50)
    
    excel_file = "MEP_Schedule_Table_20250610_154246.xlsx"
    
    try:
        xl_file = pd.ExcelFile(excel_file)
        print(f"📁 File: {excel_file}")
        print(f"📊 Worksheets: {xl_file.sheet_names}")
        
        if 'Pipe Accessory Schedule' in xl_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name='Pipe Accessory Schedule')
            print(f"✅ Found Pipe Accessory Schedule with {len(df)} rows, {len(df.columns)} columns")
            
            # Check columns F and U
            if len(df.columns) >= 21:  # Column U is 21st column (0-based index 20)
                col_f = df.columns[5]  # Column F (0-based index 5)
                col_u = df.columns[20] # Column U (0-based index 20)
                print(f"📋 Column F (index 5): {col_f}")
                print(f"📋 Column U (index 20): {col_u}")
                
                # Sample validation
                matches = 0
                mismatches = 0
                empty_pairs = 0
                mismatch_details = []
                
                for idx, row in df.iterrows():
                    f_val = str(row[col_f]).strip() if not pd.isna(row[col_f]) else ""
                    u_val = str(row[col_u]).strip() if not pd.isna(row[col_u]) else ""
                    
                    if f_val == "" and u_val == "":
                        empty_pairs += 1
                    elif f_val == u_val:
                        matches += 1
                    else:
                        mismatches += 1
                        if len(mismatch_details) < 10:  # Store first 10 mismatches
                            mismatch_details.append((idx+2, f_val, u_val))
                
                total = len(df)
                print()
                print("📊 RULE 6 VALIDATION RESULTS:")
                print(f"✅ Matches: {matches}/{total} ({matches/total*100:.1f}%)")
                print(f"❌ Mismatches: {mismatches}/{total} ({mismatches/total*100:.1f}%)")
                print(f"⚪ Empty pairs: {empty_pairs}/{total} ({empty_pairs/total*100:.1f}%)")
                
                if mismatch_details:
                    print()
                    print("🔍 SAMPLE MISMATCHES (first 10):")
                    for row_num, f_val, u_val in mismatch_details:
                        print(f"  Row {row_num:3d}: F='{f_val}' vs U='{u_val}'")
                
                # Export results with validation column
                def validate_item_family(row):
                    f_val = str(row[col_f]).strip() if not pd.isna(row[col_f]) else ""
                    u_val = str(row[col_u]).strip() if not pd.isna(row[col_u]) else ""
                    
                    if f_val == "" and u_val == "":
                        return "PASS"
                    elif f_val == "" or u_val == "":
                        return f"Item Description '{f_val}' và Family '{u_val}' phải cùng có giá trị hoặc cùng trống"
                    elif f_val == u_val:
                        return "PASS"
                    else:
                        return f"Item Description phải trùng Family: cần '{u_val}', có '{f_val}'"
                
                df['Rule6_Item_Family_Check'] = df.apply(validate_item_family, axis=1)
                
                # Export test results
                timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
                output_file = f"rule6_test_results_{timestamp}.xlsx"
                df.to_excel(output_file, index=False)
                print(f"📁 Test results exported to: {output_file}")
                
            else:
                print(f"❌ Not enough columns. Found {len(df.columns)}, need at least 21 for column U")
        else:
            print("❌ Pipe Accessory Schedule worksheet not found")
            print(f"Available worksheets: {xl_file.sheet_names}")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_rule6()
