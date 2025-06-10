#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Simple PAP 2 debug script to identify the issue
"""

import pandas as pd
import re
from pathlib import Path

def simple_pap2_debug():
    """Simple debug to see PAP 2 values"""
    
    # Load the Excel file
    excel_path = Path("d:/OneDrive/Desktop/thu nghien software/production/Xp03-Fabrication & Listing.xlsx")
    
    try:
        # Read the Pipe Schedule worksheet
        df = pd.read_excel(excel_path, sheet_name='Pipe Schedule')
        print(f"Loaded {len(df)} rows")
        
        # Find PAP 2 values
        pap2_values = []
        for idx, row in df.iterrows():
            pap2_val = row.get('Pap2', '')
            if pd.notna(pap2_val) and str(pap2_val).strip():
                pap2_values.append(str(pap2_val).strip())
        
        print(f"Found {len(pap2_values)} PAP 2 values")
        print("Unique PAP 2 values:")
        unique_vals = list(set(pap2_values))
        for val in unique_vals[:10]:  # Show first 10
            print(f"  '{val}'")
            
        # Test size code pattern
        size_code_pattern = r'\d+[A-Z]+\d*'
        
        print(f"\nTesting size code pattern on first few values:")
        for val in unique_vals[:5]:
            match = re.search(size_code_pattern, val)
            print(f"  '{val}' -> {'MATCH' if match else 'NO MATCH'}")
        
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    simple_pap2_debug()
