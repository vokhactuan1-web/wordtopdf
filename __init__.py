"""
Docx to PDF Converter Package
Hỗ trợ chuyển đổi Word và Excel sang PDF với Unicode
"""

__version__ = "1.0.0"
__author__ = "Your Name"

# Import các module chính
from src.converters.word_to_pdf import convert_word_to_pdf
from src.converters.excel_to_pdf import convert_excel_to_pdf

__all__ = ['convert_word_to_pdf', 'convert_excel_to_pdf']