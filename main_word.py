#!/usr/bin/env python3
"""
main_word.py - Word to PDF Converter

Ứng dụng chuyển đổi Word sang PDF với hỗ trợ Unicode đầy đủ
"""
import sys
from pathlib import Path

# Thêm src vào Python path
sys.path.insert(0, str(Path(__file__).parent))

from src.converters.word_to_pdf import convert_word_to_pdf
from src.interface.tkinter_ui import create_app
from src.logging.logger_setup import setup_logger


def main():
    """Main function"""
    # Setup logging
    logger = setup_logger("word_converter")
    logger.info("Khởi động Word to PDF Converter")
    
    # Cấu hình
    title = "Word to PDF Converter"
    file_types = [("Word Files", "*.docx *.doc")]
    patterns = ['*.docx', '*.doc']
    
    # Tạo app
    app = create_app(title, file_types, patterns, convert_word_to_pdf)
    
    # Chạy
    logger.info("Ứng dụng đã sẵn sàng")
    app.root.mainloop()


if __name__ == '__main__':
    main()