#!/usr/bin/env python3
# Check actual data in Excel file to understand the pattern

import pandas as pd

try:
    # Read a few rows from the Excel file to see the pattern
    df = pd.read_excel("Xp03-Fabrication & Listing.xlsx", sheet_name="Pipe Schedule", nrows=10)
    
    print("First 10 rows of Pipe Schedule:")
    print("="*50)
    
    # Show columns A, B, D (Cross Passage, Location, Array Number)
    if len(df.columns) >= 4:
        col_a = df.columns[0]  # Cross Passage
        col_b = df.columns[1]  # Location and Lanes
        col_d = df.columns[3]  # Array Number
        
        print(f"Column A ({col_a}):")
        print(df[col_a].head(10).tolist())
        print()
        
        print(f"Column B ({col_b}):")
        print(df[col_b].head(10).tolist())
        print()
        
        print(f"Column D ({col_d}):")
        print(df[col_d].head(10).tolist())
        print()
        
        print("Sample combinations:")
        for i in range(min(5, len(df))):
            a_val = df.iloc[i][col_a]
            b_val = df.iloc[i][col_b] 
            d_val = df.iloc[i][col_d]
            print(f"Row {i+2}: A='{a_val}' | B='{b_val}' | D='{d_val}'")
    else:
        print("Not enough columns found")
        print(f"Available columns: {df.columns.tolist()}")
        
except Exception as e:
    print(f"Error reading Excel file: {e}")
