# 📊 Excel Data Validation Tool

**A comprehensive tool for validating pipe/equipment data in Excel files with business-specific rules.**

---

## 🚀 **Quick Start** (For End Users)

1. **Go to the `production/` folder**
2. **Double-click `Excel_Validator.bat`**
3. **Select your Excel file when prompted**
4. **Results will be saved automatically with timestamp**

---

## 📁 **Project Structure**

```
📂 production/     🎯 Ready-to-use tools (START HERE)
📂 tools/          🔧 Development utilities
📂 archive/        📦 Old/deprecated files
📂 tests/          🧪 Test and debug scripts
📂 docs/           📚 Detailed documentation
📂 results/        📊 Output files and reports
```

---

## ✅ **Validation Rules**

### 1. **Array Number Validation** 
- **Formula**: `EE_Array Number = "EXP6" + last 2 digits of Column A + last 2 digits of Column B`
- **Applied to**: 4 worksheets (Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule, Sprinkler Schedule)

### 2. **Pipe Treatment Validation**
- **Rules**: 
  - `CP-INTERNAL` → `GAL`
  - `CP-EXTERNAL`, `CW-DISTRIBUTION`, `CW-ARRAY` → `BLACK`
- **Applied to**: 3 worksheets (Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule)

---

## 📋 **Requirements**

- **Python 3.7+** with pandas and openpyxl
- **Windows OS** (batch files provided)
- Run `pip install -r production/requirements.txt` if needed

---

## 📖 **Documentation**

- 📋 **Quick Guide**: `PROJECT_OVERVIEW.md`
- 📚 **Detailed Guide**: `docs/README_UPDATED.md`
- 🔧 **For Developers**: See `tools/` and `tests/` folders

---

## 🎯 **Current Status**

✅ **Working**: All validation rules implemented and tested  
✅ **Tested**: With real Excel file (898 rows across 4 worksheets)  
✅ **User-Ready**: Double-click .bat files for easy use  
✅ **Organized**: Clean folder structure for maintenance  

**Success Rates**: Array Number ~88.6% | Pipe Treatment ~99.4%

---

*Last Updated: June 2025*
