#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Debug script to specifically trace PAP 2 validation failures
"""

import pandas as pd
import re
from pathlib import Path

def debug_pap2_validation():
    """Debug PAP 2 validation with detailed tracing"""
    
    # Load the Excel file
    excel_path = Path("d:/OneDrive/Desktop/thu nghien software/production/Xp03-Fabrication & Listing.xlsx")
    
    if not excel_path.exists():
        print(f"❌ File not found: {excel_path}")
        return
    
    try:
        # Read the Pipe Schedule worksheet
        df = pd.read_excel(excel_path, sheet_name='Pipe Schedule')
        print(f"📊 Loaded {len(df)} rows from Pipe Schedule")
        
        # Find rows with PAP values
        pap_rows = []
        for idx, row in df.iterrows():
            pap1_val = row.get('Pap1', '')
            pap2_val = row.get('Pap2', '')
            
            if pd.notna(pap2_val) and str(pap2_val).strip():
                pap_rows.append({
                    'Index': idx,
                    'Pap1': pap1_val,
                    'Pap2': pap2_val,
                    'Length': row.get('Length', ''),
                    'Size': row.get('Size', '')
                })
        
        print(f"📏 Found {len(pap_rows)} rows with PAP 2 values")
        
        # Test PAP 2 validation logic
        pap2_results = []
        
        for pap_row in pap_rows:
            result = validate_pap2_debug(pap_row['Pap2'], pap_row['Length'], pap_row['Size'])
            pap2_results.append({
                'Index': pap_row['Index'],
                'Pap2': pap_row['Pap2'],
                'Length': pap_row['Length'], 
                'Size': pap_row['Size'],
                'Result': result
            })
            
        # Count results
        pass_count = sum(1 for r in pap2_results if "PASS" in r['Result'])
        fail_count = sum(1 for r in pap2_results if "FAIL" in r['Result'])
        
        print(f"\n📋 PAP 2 VALIDATION RESULTS:")
        print(f"   ✅ PASS: {pass_count}")
        print(f"   ❌ FAIL: {fail_count}")
        print(f"   📊 Total: {len(pap2_results)}")
        
        # Show first few failures
        failures = [r for r in pap2_results if "FAIL" in r['Result']]
        if failures:
            print(f"\n❌ FIRST 10 FAILURES:")
            for i, failure in enumerate(failures[:10]):
                print(f"   {i+1}. Row {failure['Index']}: {failure['Pap2']} -> {failure['Result']}")
        
        # Show sample passes
        passes = [r for r in pap2_results if "PASS" in r['Result']]
        if passes:
            print(f"\n✅ FIRST 5 PASSES:")
            for i, success in enumerate(passes[:5]):
                print(f"   {i+1}. Row {success['Index']}: {success['Pap2']} -> {success['Result']}")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()

def validate_pap2_debug(pap2_val, length_val, size_val):
    """Debug version of PAP 2 validation logic"""
    try:
        pap2_str = str(pap2_val).strip()
        
        print(f"\n🔍 Debugging PAP 2: '{pap2_str}'")
        
        if not pap2_str or pap2_str.lower() in ['nan', 'none', '']:
            return "SKIP (Empty PAP2)"
        
        # Check if it's a numeric length and size
        try:
            length_num = float(length_val) if pd.notna(length_val) else None
            size_num = float(size_val) if pd.notna(size_val) else None
            
            print(f"   📏 Length: {length_num}, Size: {size_num}")
            
            # Special rule: 65mm pipes with 5295mm length
            if size_num == 65.0 and length_num is not None:
                print(f"   🔧 Checking 65mm special rule, length={length_num}")
                if abs(length_num - 5295.0) < 5.0:  # Updated tolerance
                    print(f"   ✅ 65mm-5295mm rule applies")
                    
                    # Pattern 1: Dimension format (NxN, NxNxN)
                    dimension_pattern = r'\d+x\d+(?:x\d+)?'
                    # Pattern 2: Size codes (like 40B, 65LR, 100A, etc.)
                    size_code_pattern = r'\d+[A-Z]+\d*'
                    
                    if re.search(dimension_pattern, pap2_str):
                        print(f"   ✅ Dimension pattern matched")
                        return "PASS (Rule: 65mm-5295mm Dimension)"
                    elif re.search(size_code_pattern, pap2_str):
                        print(f"   ✅ Size code pattern matched")
                        return "PASS (Rule: 65mm-5295mm Size Code)"
                    else:
                        print(f"   ❌ No pattern matched")
                        return f"FAIL (Rule: 65mm-5295mm): ống 65mm dài 5295mm cần dimension format (NxN) hoặc size code (40B, 65LR), có '{pap2_str}'"
        except (ValueError, TypeError):
            print(f"   ⚠️ Could not parse length/size as numbers")
            pass
            
        # General validation patterns
        print(f"   🔧 Checking general patterns")
        dimension_pattern = r'\d+x\d+(?:x\d+)?'
        size_code_pattern = r'\d+[A-Z]+\d*'
        
        if re.search(dimension_pattern, pap2_str):
            print(f"   ✅ General dimension pattern matched")
            return "PASS (Rule: Valid Dimension Format)"
        elif re.search(size_code_pattern, pap2_str):
            print(f"   ✅ General size code pattern matched")
            return "PASS (Rule: Valid Size Code)"
        else:
            print(f"   ❌ No general pattern matched")
            return f"FAIL (Rule: Valid Format): cần format NxN, NxNxN hoặc Size Code (ví dụ: 40B, 65LR), có '{pap2_str}'"
                
    except Exception as e:
        print(f"   ❌ Exception: {str(e)}")
        return f"ERROR: {str(e)}"

if __name__ == "__main__":
    debug_pap2_validation()
