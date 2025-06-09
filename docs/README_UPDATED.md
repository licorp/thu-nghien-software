# Excel Validation Tool - Pipe Treatment & Array Number

## 📁 CÁC FILE CHÍNH

### 🚀 **Tool Validation Chính:**
- `excel_validator_detailed.py` - **Tool validation mới nhất** với hiển thị chi tiết cả Array Number và Pipe Treatment
- `excel_validator_final.py` - Tool validation cũ (chưa hiển thị chi tiết Pipe Treatment)

### 🖱️ **File BAT để chạy nhanh:**
- `run_excel_validator.bat` - Chạy tool validation mới 
- `Excel_Validator.bat` - Chạy tool validation mới (tương tự)

### 📊 **Tool phân tích:**
- `analyze_all_worksheets.py` - Phân tích cấu trúc worksheet và test validation logic
- `debug_pipe_treatment.py` - Debug Pipe Treatment validation

## 🔧 QUY TẮC VALIDATION

### 1. **Array Number Validation**
- **Áp dụng cho**: Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule, Sprinkler Schedule
- **Quy tắc**: Cột D (EE_Array Number) phải chứa 'EXP6' + 2 số cuối cột B + 2 số cuối cột A
- **Ví dụ**: 
  - Cột A: EXP61002 → lấy "02"
  - Cột B: M110 → lấy "10" 
  - Expected: EXP61002 phải chứa "EXP61002"

### 2. **Pipe Treatment Validation**  
- **Áp dụng cho**: Pipe Schedule, Pipe Fitting Schedule, Pipe Accessory Schedule
- **Quy tắc**:
  - CP-INTERNAL → GAL
  - CP-EXTERNAL → BLACK
  - CW-DISTRIBUTION → BLACK  
  - CW-ARRAY → BLACK

## 📈 KẾT QUẢ VALIDATION MỚI NHẤT

### ✅ **Pipe Treatment Validation**: 99.4% thành công
- Pipe Schedule: 190/190 (100%)
- Pipe Fitting Schedule: 401/401 (100%) 
- Pipe Accessory Schedule: 226/228 (99.1%) - có 2 lỗi

### 🔢 **Array Number Validation**: 88.6% thành công
- Chủ yếu lỗi ở pattern M110-M111 → cần EXP61102 nhưng có EXP61002

## 🚀 CÁCH SỬ DỤNG

### **Cách 1: Double-click file .bat**
```
🖱️ Double-click: run_excel_validator.bat
```

### **Cách 2: Chạy Python trực tiếp**
```bash
python excel_validator_detailed.py
```

### **Cách 3: Phân tích cấu trúc**
```bash
python analyze_all_worksheets.py
```

## 📁 KẾT QUẢ

Tool sẽ tạo file Excel với kết quả validation và hiển thị:
- ✅ Thống kê tổng quan
- 🔢 Chi tiết Array Number validation  
- 🔧 Chi tiết Pipe Treatment validation
- ❌ Danh sách lỗi cụ thể

## 🎯 UPDATE NOTES

- **2025-06-09**: Thêm `excel_validator_detailed.py` với hiển thị chi tiết Pipe Treatment
- **2025-06-09**: Cập nhật cả 2 file .bat để sử dụng tool mới
- **Pipe Treatment validation** hiện đã hoạt động đúng cho cả 3 worksheets
- **Tỷ lệ thành công Pipe Treatment**: 99.4% (chỉ 2 lỗi trong 822 dòng)
