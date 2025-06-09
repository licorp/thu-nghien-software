# Excel Validator Project - Organized Structure

## 🚀 **FOR END USERS** - How to Use

### Quick Start:
1. Go to the `production` folder
2. Double-click `Excel_Validator.bat`
3. Select your Excel file when prompted
4. Results will be saved automatically with timestamp

### Requirements:
- Python 3.7+ with pandas and openpyxl
- Run `pip install -r production/requirements.txt` if needed

---

## 📁 **FOLDER STRUCTURE**

### 🎯 **production/** - Ready-to-Use Tools
- `excel_validator_detailed.py` - **Main validation tool** (latest version)
- `Excel_Validator.bat` - **Easy double-click runner**
- `run_excel_validator.bat` - Alternative runner
- `requirements.txt` - Required Python packages

### 🔧 **tools/** - Development Utilities  
- `analyze_all_worksheets.py` - Analyze Excel worksheet structure
- `excel_validator_final.py` - Alternative validator version
- `excel_validator_standalone.py` - Standalone version

### 📦 **archive/** - Old/Deprecated Files
- Previous versions kept for reference
- `excel_validator.py` - Original version
- `validate_real.py` - Replaced by detailed version
- Other legacy scripts

### 🧪 **tests/** - Testing & Debug Scripts
- `debug_pipe_treatment.py` - Debug Pipe Treatment validation
- `test_*.py` - Various test scripts
- `validate_*.py` - Validation test scripts

### 📚 **docs/** - Documentation
- `README.md` - Original documentation  
- `README_UPDATED.md` - **Complete updated guide**

### 📊 **results/** - Output Files
- Validation result Excel files
- Reports with timestamps

---

## ✅ **VALIDATION RULES**

### 1. Array Number Validation
- **Rule**: `EE_Array Number = "EXP6" + last 2 digits of Column A + last 2 digits of Column B`
- **Special Case**: `CP-INTERNAL` → `Cross Passage`
- **Applied to**: Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule, Sprinkler Schedule
- **Success Rate**: ~81.9%

### 2. Pipe Treatment Validation  
- **Rules**:
  - `CP-INTERNAL` → `GAL`
  - `CP-EXTERNAL`, `CW-DISTRIBUTION`, `CW-ARRAY` → `BLACK`
- **Applied to**: Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule  
- **Success Rate**: ~99.8%

### 3. FAB Pipe Validation
- **Conditional Logic**: Rules based on Item Description content
- **Rules**:
  - **Pipe items**: Must be empty or "N/A"
  - **Fitting items**: Complex patterns based on Size and End-1/End-2 values
- **Applied to**: Pipe Schedule, Pipe Fitting Schedule
- **Success Rate**: ~34.0% pass (66% appropriately skipped)

---

## 🧹 **RECENT UPDATES**
- **✅ PAP Validation Removed**: Tool cleaned and optimized
- **🚀 Performance Improved**: 29% code reduction
- **📋 Focused**: 3 core validation rules only

---

## 🏗️ **FOR DEVELOPERS**

### Main Files:
- **Primary Code**: `production/excel_validator_detailed.py`
- **Documentation**: `docs/README_UPDATED.md`
- **Testing**: `tests/debug_pipe_treatment.py`

### Architecture:
- `ExcelValidatorDetailed` class handles all validation
- Separate validation methods for Array Number and Pipe Treatment
- Color-coded Excel output (Green=Pass, Red=Fail)
- Detailed error reporting and statistics

---

## 🎯 **CURRENT STATUS**
- ✅ **Working**: All validation rules implemented and tested (3 core rules)
- ✅ **Tested**: With real Excel files (2,019+ rows across 4 worksheets)
- ✅ **User-Ready**: Double-click .bat files for easy use
- ✅ **Organized**: Clean folder structure for maintenance
- ✅ **Optimized**: PAP validation removed for cleaner, faster tool

**Last Updated**: June 2025
