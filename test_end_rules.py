#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test script to verify Enhanced End-1/End-2 validation rules
"""

import pandas as pd

def test_end_rules():
    print("🧪 TESTING ENHANCED END-1/END-2 RULES")
    print("=" * 50)
    
    # Read the Pipe Schedule sheet to look for End-1/End-2 data
    try:
        df = pd.read_excel("Xp03-Fabrication & Listing.xlsx", sheet_name="Pipe Schedule")
        
        # Check column L (End-1) and M (End-2)
        col_l = df.columns[11] if len(df.columns) > 11 else None  # End-1
        col_m = df.columns[12] if len(df.columns) > 12 else None  # End-2
        col_k = df.columns[10] if len(df.columns) > 10 else None  # FAB Pipe
        
        print(f"Column L (End-1): {col_l}")
        print(f"Column M (End-2): {col_m}")
        print(f"Column K (FAB Pipe): {col_k}")
        print()
        
        if col_l and col_m and col_k:
            # Look for rows with BE values
            be_rows = df[(df[col_l] == "BE") | (df[col_m] == "BE")]
            print(f"📊 Rows with End-1='BE' or End-2='BE': {len(be_rows)}")
            
            if len(be_rows) > 0:
                print("Sample BE rows:")
                for idx, (_, row) in enumerate(be_rows.head(3).iterrows(), 1):
                    print(f"  {idx}. End-1: {row[col_l]}, End-2: {row[col_m]}, FAB Pipe: {row[col_k]}")
            print()
            
            # Look for rows with RG/TH values
            rg_th_rows = df[(df[col_l].isin(["RG", "TH"])) & (df[col_m].isin(["RG", "TH"]))]
            print(f"📊 Rows with both End-1 & End-2 in ['RG','TH']: {len(rg_th_rows)}")
            
            if len(rg_th_rows) > 0:
                print("Sample RG/TH rows:")
                for idx, (_, row) in enumerate(rg_th_rows.head(3).iterrows(), 1):
                    print(f"  {idx}. End-1: {row[col_l]}, End-2: {row[col_m]}, FAB Pipe: {row[col_k]}")
            print()
            
            # Show unique values in End-1 and End-2 columns
            end1_values = df[col_l].dropna().unique()
            end2_values = df[col_m].dropna().unique()
            
            print(f"📋 Unique End-1 values: {sorted(end1_values) if len(end1_values) > 0 else 'None'}")
            print(f"📋 Unique End-2 values: {sorted(end2_values) if len(end2_values) > 0 else 'None'}")
            
        else:
            print("❌ Cannot find End-1/End-2 columns")
            
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    test_end_rules()
