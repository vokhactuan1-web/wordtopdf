# 📄 Document to PDF Converter

Ứng dụng chuyển đổi Word và Excel sang PDF với hỗ trợ **Unicode đầy đủ** cho tiếng Việt.

## ✨ Tính năng

- ✅ Chuyển đổi **Word** (.docx, .doc) sang PDF
- ✅ Chuyển đổi **Excel** (.xlsx, .xls) sang PDF
- ✅ Hỗ trợ **tiếng Việt** hoàn toàn
- ✅ Tự động tải font Unicode nếu cần
- ✅ Giao diện đơn giản, dễ sử dụng
- ✅ Hỗ trợ kéo thả file (drag & drop)
- ✅ Xử lý hàng loạt nhiều file
- ✅ Logging chi tiết

## 📁 Cấu trúc dự án

```
docxtopdf/
├── README.md                 # File này
├── requirements.txt          # Danh sách thư viện
├── main_word.py              # Chạy Word Converter
├── main_excel.py             # Chạy Excel Converter
├── logs/                     # Thư mục chứa log files
└── src/
    ├── __init__.py
    ├── logging/
    │   └── logger_setup.py   # Thiết lập logging
    ├── interface/
    │   └── tkinter_ui.py     # Giao diện người dùng
    ├── io/
    │   └── file_handler.py   # Xử lý file I/O
    └── converters/
        ├── word_to_pdf.py    # Chuyển đổi Word
        └── excel_to_pdf.py   # Chuyển đổi Excel
```

## 🚀 Cài đặt

### 1. Clone hoặc tải project

```bash
git clone <repository-url>
cd docxtopdf
```

### 2. Cài đặt dependencies

```bash
pip install -r requirements.txt
```

### 3. Chạy ứng dụng

**Chuyển đổi Word:**
```bash
python main_word.py
```

**Chuyển đổi Excel:**
```bash
python main_excel.py
```

## 📦 Dependencies

- `openpyxl` - Đọc file Excel
- `reportlab` - Tạo file PDF
- `xlrd` - Hỗ trợ file .xls cũ
- `python-docx` - Đọc file Word
- `tkinterdnd2` - Hỗ trợ drag & drop (tùy chọn)

## 💡 Cách sử dụng

### Giao diện

1. **Thêm file:**
   - Kéo thả file vào vùng drop
   - Hoặc nhấn "📂 Chọn File"
   - Hoặc nhấn "📁 Chọn Thư Mục" để thêm hàng loạt

2. **Chuyển đổi:**
   - Nhấn "🔄 CHUYỂN ĐỔI SANG PDF"
   - File PDF sẽ được tạo cùng thư mục với file gốc

3. **Xem kết quả:**
   - Nhấn "📥 Mở Downloads" để mở thư mục Downloads
   - Hoặc mở thư mục chứa file gốc

### Code API

**Chuyển đổi Word:**
```python
from src.converters.word_to_pdf import convert_word_to_pdf
from pathlib import Path

input_file = Path("document.docx")
output_file = Path("document.pdf")

convert_word_to_pdf(input_file, output_file)
```

**Chuyển đổi Excel:**
```python
from src.converters.excel_to_pdf import convert_excel_to_pdf
from pathlib import Path

input_file = Path("data.xlsx")
output_file = Path("data.pdf")

convert_excel_to_pdf(input_file, output_file)
```

## 🎨 Tính năng Excel Converter

- ✅ Hỗ trợ nhiều sheets
- ✅ Tự động điều chỉnh độ rộng cột
- ✅ Format bảng đẹp với header màu
- ✅ Zebra striping (dòng xen kẽ màu)
- ✅ Giới hạn 500 dòng mỗi sheet (có thể điều chỉnh)
- ✅ Landscape mode cho bảng rộng

## 🎨 Tính năng Word Converter

- ✅ Giữ nguyên formatting (bold, italic)
- ✅ Hỗ trợ headings (H1, H2, H3)
- ✅ Hỗ trợ bullet lists
- ✅ Chuyển đổi tables
- ✅ Giữ nguyên cấu trúc document

## 🔧 Tùy chỉnh

### Thay đổi font

Trong `src/converters/excel_to_pdf.py` hoặc `word_to_pdf.py`:

```python
class FontManager:
    SYSTEM_FONTS = [
        'path/to/your/font.ttf',
        # Thêm font paths ở đây
    ]
```

### Thay đổi page size

```python
# Excel - Landscape A4
doc = SimpleDocTemplate(
    str(output_path),
    pagesize=landscape(A4),  # Thay đổi ở đây
)

# Word - Portrait A4
doc = SimpleDocTemplate(
    str(output_path),
    pagesize=A4,  # Hoặc letter, A3, etc.
)
```

### Thay đổi màu sắc bảng

```python
# Trong _get_table_style():
('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498DB')),  # Header color
('ROWBACKGROUNDS', (0, 1), (-1, -1),
 [colors.white, colors.HexColor('#ECF0F1')]),  # Zebra colors
```

## 📝 Logging

Log files được lưu trong thư mục `logs/`:
- `converter_YYYYMMDD.log` - Log theo ngày

Cấu hình logging trong `src/logging/logger_setup.py`

## ⚠️ Lưu ý

- Font Unicode sẽ được tự động tải từ GitHub nếu không tìm thấy trên hệ thống
- File PDF sẽ được tạo cùng thư mục với file gốc (trừ khi chỉ định output path)
- Excel files lớn (>500 dòng) sẽ bị giới hạn để tránh PDF quá lớn
- Word files phức tạp có thể mất formatting một số

## 🐛 Troubleshooting

**Lỗi font tiếng Việt:**
- Đảm bảo có kết nối internet để tải font
- Hoặc cài đặt font DejaVu Sans thủ công

**Lỗi import:**
```bash
pip install --upgrade -r requirements.txt
```

**Drag & drop không hoạt động:**
```bash
pip install tkinterdnd2
```

## 📜 License

MIT License - Tự do sử dụng và chỉnh sửa

## 🤝 Đóng góp

Mọi đóng góp đều được hoan nghênh! Vui lòng:
1. Fork project
2. Tạo branch mới
3. Commit changes
4. Push và tạo Pull Request

## 📧 Liên hệ

Có câu hỏi? Tạo issue trên GitHub!

---

**Made with ❤️ for Vietnamese users**