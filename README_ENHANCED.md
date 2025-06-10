# 📊 Excel Data Validation Tool - ENHANCED

**A comprehensive tool for validating pipe/equipment data in Excel files with business-specific rules. Now with ENHANCED Rule 4 features!**

---

## 🚀 **Quick Start** (For End Users)

1. **Double-click `🚀 START HERE.bat`**
2. **Select your Excel file when prompted**
3. **View validation results with enhanced End-1/End-2 rules**

---

## ✅ **Validation Rules - ENHANCED**

### 1. **Array Number Validation** 
- **Formula**: `EE_Array Number = "EXP6" + last 2 digits of Column A + last 2 digits of Column B`
- **Special Rule**: `CP-INTERNAL` → `Cross Passage`
- **Applied to**: 4 worksheets (Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule, Sprinkler Schedule)

### 2. **Pipe Treatment Validation**
- **Rules**: 
  - `CP-INTERNAL` → `GAL`
  - `CP-EXTERNAL`, `CW-DISTRIBUTION`, `CW-ARRAY` → `BLACK`
- **Applied to**: 3 worksheets (Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule)

### 3. **CP-INTERNAL Array Number Validation**
- **Rule**: When `EE_System Type = "CP-INTERNAL"`, then `EE_Array Number` must equal `EE_Cross Passage`
- **Priority**: Overrides Rule 1 for CP-INTERNAL records
- **Applied to**: 3 worksheets (Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule)

### 4. **Pipe Schedule Mapping Validation - 🆕 ENHANCED**
- **Applied to**: Pipe Schedule worksheet only
- **Original Rules**:
  - Item Description "150-900" → FAB Pipe "STD ARRAY TEE"
  - Item Description "65-4730" → FAB Pipe "STD 1 PAP RANGE"
  - Item Description "65-5295" → FAB Pipe "STD 2 PAP RANGE"
  - Size "40" → FAB Pipe "Groove_Thread"
- **🆕 NEW ENHANCED Rules**:
  - **End-1 or End-2 = "BE"** → FAB Pipe **"Fabrication"**
  - **End-1 AND End-2 both in ["RG", "TH"]** → FAB Pipe **"Groove_Thread"**

---

## 📁 **Project Structure - ENHANCED**

```
📂 Main Folder/
├── 📄 excel_validator_enhanced.py     🆕 ENHANCED validation tool
├── 📄 🚀 START HERE.bat              🎯 Easy launcher (UPDATED)
├── 📄 README.md                      📚 This documentation (ENHANCED)
├── 📄 requirements.txt               📋 Dependencies
├── 📄 excel_validator_backup.py      💾 Backup of original version
├── 📄 Xp03-Fabrication & Listing.xlsx 🧪 Test data
└── 📄 Xp04-Fabrication & Listing.xlsx 🧪 Test data
```

---

## 🧹 **Latest Updates - ENHANCED**
- **🆕 Rule 4 Enhanced**: Added End-1/End-2 validation logic  
- **✅ Complete Testing**: All 4 rules working with new features
- **🚀 Enhanced Display**: Shows L (End-1) and M (End-2) columns in error output
- **📋 Updated Documentation**: Reflects all new capabilities

---

## 📋 **Requirements**

- **Python 3.7+** with pandas and openpyxl
- **Windows OS** (batch files provided)
- Run `pip install pandas openpyxl` if needed

---

## 🎯 **Current Status - ENHANCED**

✅ **Working**: All 4 validation rules implemented and tested with ENHANCED features  
✅ **Tested**: With real Excel files (2,000+ rows across 4 worksheets)  
✅ **User-Ready**: Double-click .bat files for easy use  
✅ **Enhanced**: New End-1/End-2 validation logic in Rule 4
✅ **Complete**: Production-ready enhanced validation tool

**Enhanced Features**: 
- **Rule 4**: Now includes End-1/End-2 column validation
- **Display**: Shows columns C, D, F, G, K, L, M, T in error output
- **Logic**: Smart priority handling for all 4 validation rules
- **Output**: Files named `validation_enhanced_*.xlsx`

---

*Last Updated: June 10, 2025 - ENHANCED VERSION*
