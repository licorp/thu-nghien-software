#!/usr/bin/env python3
"""
Test script to run PAP validation directly
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from excel_validator_detailed import ExcelValidatorDetailed

if __name__ == "__main__":
    validator = ExcelValidatorDetailed()
    excel_file = "d:\\OneDrive\\Desktop\\thu nghien software\\Xp03-Fabrication & Listing.xlsx"
    
    print("🚀 Testing PAP validation on Xp03-Fabrication & Listing.xlsx")
    print("=" * 60)
    
    # Call validation directly without the interactive menu
    try:
        validator.validate_excel_file(excel_file)
        print("✅ Validation completed successfully!")
    except Exception as e:
        print(f"❌ Error during validation: {e}")
        import traceback
        traceback.print_exc()
