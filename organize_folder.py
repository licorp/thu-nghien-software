#!/usr/bin/env python3
"""
Folder Organization Script
Organize Excel validation project files into clean structure
"""

import os
import shutil
from pathlib import Path

def organize_folder():
    """Organize files into proper folder structure"""
    base_path = Path("d:/OneDrive/Desktop/thu nghien software")
    
    # Create organized folder structure
    folders = {
        'production': 'Production-ready tools for end users',
        'tools': 'Development and utility tools',
        'archive': 'Old/deprecated files',
        'tests': 'Test and debug scripts',
        'docs': 'Documentation files',
        'results': 'Output files and results'
    }
    
    for folder in folders:
        folder_path = base_path / folder
        folder_path.mkdir(exist_ok=True)
        print(f"Created: {folder_path}")
    
    # File organization mapping
    file_mapping = {
        'production': [
            'excel_validator_detailed.py',  # Main production tool
            'Excel_Validator.bat',          # User-friendly runner
            'run_excel_validator.bat',      # Alternative runner
            'requirements.txt'              # Dependencies
        ],
        'tools': [
            'analyze_all_worksheets.py',    # Analysis tool
            'excel_validator_final.py',     # Alternative validator
            'excel_validator_standalone.py' # Standalone version
        ],
        'archive': [
            'excel_validator.py',           # Old version
            'validate_real.py',             # Replaced by detailed version
            'run_validator.py',             # Old runner
            'validate_custom.py',           # Custom validation
            'validate_multi_worksheet.py',  # Multi-worksheet validator
            'check_worksheets.py',          # Worksheet checker
            'create_sample.py'              # Sample creator
        ],
        'tests': [
            'debug_pipe_treatment.py',      # Debug tool
            'test_array_validation.py',     # Array validation test
            'test_fixed_validation.py',     # Fixed validation test
            'test_pipe_treatment.py',       # Pipe treatment test
            'validate_array_number_only.py', # Array number only test
            'validate_combined_rules.py',   # Combined rules test
            'check_array_number.py',        # Array number checker
            'check_item_desc.py',           # Item description checker
            'analyze_real_data.py'          # Real data analyzer
        ],
        'docs': [
            'README.md',                    # Original README
            'README_UPDATED.md'             # Updated README
        ],
        'results': [
            'array_number_validation_20250608_213150.xlsx',
            'validation_result_20250609_050036.xlsx',
            'validation_result_20250609_054611.xlsx',
            '~$validation_result_20250609_050036.xlsx'
        ]
    }
    
    # Move files to organized folders
    for folder, files in file_mapping.items():
        folder_path = base_path / folder
        for file in files:
            source = base_path / file
            target = folder_path / file
            
            if source.exists():
                try:
                    shutil.move(str(source), str(target))
                    print(f"Moved: {file} → {folder}/")
                except Exception as e:
                    print(f"Error moving {file}: {e}")
            else:
                print(f"File not found: {file}")
    
    # Keep important files in root
    keep_in_root = [
        'Xp02-Fabrication & Listing.xlsx',  # Test data file
        '__pycache__',                       # Python cache
        '.github'                           # GitHub folder
    ]
    
    print("\nFiles kept in root:")
    for item in keep_in_root:
        if (base_path / item).exists():
            print(f"  - {item}")
    
    print("\n✅ Folder organization complete!")
    
    # Create a quick reference file
    create_quick_reference(base_path)

def create_quick_reference(base_path):
    """Create a quick reference guide"""
    ref_content = """# Excel Validator - Quick Reference

## 🚀 How to Use (For End Users)
1. Double-click `production/Excel_Validator.bat`
2. Select your Excel file when prompted
3. Results will be saved automatically

## 📁 Folder Structure
- **production/**: Ready-to-use tools
  - `excel_validator_detailed.py` - Main validation tool
  - `Excel_Validator.bat` - Easy-to-use runner
  - `requirements.txt` - Required packages

- **tools/**: Development utilities
  - Analysis and alternative validation tools

- **archive/**: Old/deprecated files
  - Previous versions kept for reference

- **tests/**: Testing and debugging
  - Scripts used during development

- **docs/**: Documentation
  - README files and guides

- **results/**: Output files
  - Validation results and reports

## 🔧 For Developers
- Main code: `production/excel_validator_detailed.py`
- Test with: `tests/debug_pipe_treatment.py`
- Documentation: `docs/README_UPDATED.md`

## 📋 Validation Rules
1. **Array Number**: EE_Array Number = "EXP6" + last 2 digits of Column A + last 2 digits of Column B
2. **Pipe Treatment**: CP-INTERNAL → GAL, CP-EXTERNAL/CW-DISTRIBUTION/CW-ARRAY → BLACK
"""
    
    with open(base_path / "QUICK_START.md", "w", encoding="utf-8") as f:
        f.write(ref_content)
    
    print("Created: QUICK_START.md")

if __name__ == "__main__":
    organize_folder()
