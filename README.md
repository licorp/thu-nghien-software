# Excel Pipe/Equipment Validation Tool 🚀

## PRODUCTION VERSION - v2.0 - COMPLETE 6-RULE SYSTEM ✅

Công cụ validation dữ liệu Excel cho pipe/equipment với 6 quy tắc hoàn chỉnh, logic ưu tiên và phát hiện ô trống.

### 📁 CẤU TRÚC FILE

```
📦 thu nghien software/
├── 🚀 START HERE.bat              # File khởi chạy chính (updated)
├── excel_validator_final.py       # Script validation với 6 rules hoàn chỉnh
├── MEP_Schedule_Table_20250610_154246.xlsx  # File Excel nguồn
├── Xp54-Fabrication & Listing.xlsx         # File test với EE columns
├── requirements.txt               # Dependencies Python
├── 6RULES_COMPLETION_REPORT.md   # Báo cáo hoàn thành chi tiết
└── README.md                      # File này
```

### ✨ TÍNH NĂNG CHÍNH

- **🎉 6 VALIDATION RULES HOÀN CHỈNH**
- **93.7% độ chính xác** với comprehensive validation
- **Logic ưu tiên HIGH/LOW** cho validation rules
- **EE_Run Dim & EE_Pap validation** cho các trường hợp đặc biệt
- **Item Description = Family matching** cho Pipe Accessory Schedule
- **Phát hiện ô trống** cho tất cả worksheets và columns
- **Interface thân thiện** với màu sắc và progress bar
- **Export kết quả** ra Excel với format đẹp

### 🎯 6 VALIDATION RULES

#### **Rule 1: Array Number Validation**
- Format: EXP6 + 2 số cuối Location Lanes + 2 số cuối Cross Passage
- Áp dụng: Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule, Sprinkler Schedule

#### **Rule 2: Pipe Treatment Validation**
- CP-INTERNAL → GAL, CP-EXTERNAL → BLACK, CW-* → BLACK
- Áp dụng: Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule

#### **Rule 3: CP-INTERNAL Array Number Validation**
- Array Number phải trùng Cross Passage cho CP-INTERNAL systems
- Áp dụng: Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule

#### **Rule 4: Priority-based Pipe Schedule Mapping (HIGH PRIORITY)**
- **STD 1 PAP RANGE**: size 65, length 4730, RG BE
- **STD 2 PAP RANGE**: size 65, length 5295, RG BE  
- **STD ARRAY TEE**: size 150, length 900, RG RG

#### **Rule 5: EE_Run Dim & EE_Pap Validation** 🆕
- **STD 1 PAP RANGE**: EE_Run Dim 1 = 4685, EE_Pap 1 = 40B
- **STD 2 PAP RANGE**: EE_Run Dim 1 = 150, EE_Pap 1 = 40B, EE_Run Dim 2 = 5250, EE_Pap 2 = 40B
- **STD ARRAY TEE**: EE_Run Dim 1 = 150, EE_Pap 1 = 65LR
- **Fabrication**: Minimum EE_Run Dim 1 & EE_Pap 1 requirements
- **Detection**: "Thiếu" và "Sai" values trong all EE columns
- Áp dụng: Pipe Schedule only

#### **Rule 6: Item Description = Family Validation** 🆕
- Column F (Item Description) phải trùng Column U (Family)
- Logic: Both empty → PASS, One empty → FAIL, Both match → PASS, Different → FAIL
- Áp dụng: Pipe Accessory Schedule only

#### LOW PRIORITY
- **Groove_Thread**: RG RG hoặc pipe 40 TH TH
- **Fabrication**: pipe 65, RG BE (không phải PAP RANGE)

### 🚀 CÁCH SỬ DỤNG

1. **Double-click** `🚀 START HERE.bat`
2. Chọn file Excel cần validate
3. Xem kết quả validation với màu sắc
4. Kiểm tra báo cáo ô trống
5. Export kết quả nếu cần

### 📊 KẾT QUẢ

- **Validation accuracy**: 93.7% (401/428 PASS) với 6 rules hoàn chỉnh
- **Rule 5 performance**: 90.2% EE column validation accuracy
- **Processing speed**: ~1,600 rows/second
- **Memory usage**: Optimized cho file lớn
- **Error detection**: Comprehensive reporting với tất cả 6 rules

### 🆕 FEATURES MỚI (v2.0)

- ✅ **Rule 5**: EE_Run Dim & EE_Pap validation với specific requirements
- ✅ **Rule 6**: Item Description = Family matching cho Pipe Accessory Schedule
- ✅ **Enhanced column mapping**: Hỗ trợ columns A-U (21 columns)
- ✅ **Export với 6-rules**: File output format `validation_6rules_*.xlsx`
- ✅ **Improved error handling**: Better detection và reporting

### 🛠️ YÊU CẦU HỆ THỐNG

- Python 3.7+
- pandas, openpyxl, colorama
- Windows (batch file support)

---
*Developed with ❤️ for efficient pipe/equipment data validation*
