#!/usr/bin/env python3
# Test script for 6-rule validation system

import sys
import os
sys.path.append(r'd:\OneDrive\Desktop\thu nghien software')

from excel_validator_final import run_validation_on_file

def test_validation():
    """Test the complete 6-rule validation system"""
    print("🧪 TESTING 6-RULE VALIDATION SYSTEM")
    print("=" * 60)
    
    # Test file with EE columns
    excel_file = r"d:\OneDrive\Desktop\thu nghien software\Xp54-Fabrication & Listing.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"❌ File not found: {excel_file}")
        return
    
    print(f"📁 Testing file: Xp54-Fabrication & Listing.xlsx")
    print()
    
    # Run validation
    output_file = run_validation_on_file(excel_file)
    
    if output_file:
        print(f"✅ Test completed successfully!")
        print(f"📁 Output: {output_file}")
    else:
        print("❌ Test failed!")

if __name__ == "__main__":
    test_validation()
