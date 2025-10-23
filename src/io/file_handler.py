"""
Xử lý các thao tác file I/O
"""
import os
import sys
import platform
import subprocess
from pathlib import Path
from typing import List, Optional
from tkinter import filedialog

from ..logging.logger_setup import get_logger

logger = get_logger(__name__)


class FileHandler:
    """Class xử lý các thao tác file"""
    
    @staticmethod
    def select_files(file_types: List[tuple], title: str = "Chọn file") -> List[Path]:
        """
        Mở dialog chọn nhiều file
        
        Args:
            file_types: Danh sách tuple (tên, pattern) VD: [("Word Files", "*.docx *.doc")]
            title: Tiêu đề dialog
            
        Returns:
            List[Path]: Danh sách đường dẫn file đã chọn
        """
        files = filedialog.askopenfilenames(
            title=title,
            filetypes=file_types + [("All Files", "*.*")]
        )
        return [Path(f) for f in files]
    
    @staticmethod
    def select_folder(title: str = "Chọn thư mục") -> Optional[Path]:
        """
        Mở dialog chọn thư mục
        
        Args:
            title: Tiêu đề dialog
            
        Returns:
            Optional[Path]: Đường dẫn thư mục hoặc None
        """
        folder = filedialog.askdirectory(title=title)
        return Path(folder) if folder else None
    
    @staticmethod
    def get_files_from_folder(folder: Path, patterns: List[str]) -> List[Path]:
        """
        Lấy tất cả file matching patterns từ folder và subfolder
        
        Args:
            folder: Đường dẫn thư mục
            patterns: Danh sách pattern VD: ['*.docx', '*.doc']
            
        Returns:
            List[Path]: Danh sách file tìm được
        """
        files = []
        for pattern in patterns:
            files.extend(folder.rglob(pattern))
        logger.info(f"Tìm thấy {len(files)} file trong {folder}")
        return files
    
    @staticmethod
    def validate_file(file_path: Path, valid_extensions: List[str]) -> bool:
        """
        Kiểm tra file có extension hợp lệ không
        
        Args:
            file_path: Đường dẫn file
            valid_extensions: Danh sách extension hợp lệ VD: ['.docx', '.doc']
            
        Returns:
            bool: True nếu hợp lệ
        """
        if not file_path.exists():
            logger.warning(f"File không tồn tại: {file_path}")
            return False
        
        if not file_path.is_file():
            logger.warning(f"Không phải file: {file_path}")
            return False
        
        if file_path.suffix.lower() not in valid_extensions:
            logger.warning(f"Extension không hợp lệ: {file_path.suffix}")
            return False
        
        return True
    
    @staticmethod
    def open_folder(folder_path: Path):
        """
        Mở thư mục trong file explorer
        
        Args:
            folder_path: Đường dẫn thư mục
        """
        if not folder_path.exists():
            logger.error(f"Thư mục không tồn tại: {folder_path}")
            return
        
        try:
            system = platform.system()
            
            if system == "Windows":
                os.startfile(folder_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", str(folder_path)])
            else:  # Linux
                subprocess.run(["xdg-open", str(folder_path)])
            
            logger.info(f"Đã mở thư mục: {folder_path}")
        except Exception as e:
            logger.error(f"Không thể mở thư mục: {e}")
    
    @staticmethod
    def get_downloads_folder() -> Path:
        """
        Lấy đường dẫn thư mục Downloads
        
        Returns:
            Path: Đường dẫn Downloads folder
        """
        home = Path.home()
        downloads = home / "Downloads"
        
        if not downloads.exists():
            downloads.mkdir(exist_ok=True)
            logger.info(f"Đã tạo thư mục Downloads: {downloads}")
        
        return downloads
    
    @staticmethod
    def ensure_output_path(output_path: Optional[Path], input_path: Path, 
                          new_suffix: str = '.pdf') -> Path:
        """
        Đảm bảo output path hợp lệ, tạo từ input nếu không có
        
        Args:
            output_path: Đường dẫn output (có thể None)
            input_path: Đường dẫn input
            new_suffix: Suffix mới cho file output
            
        Returns:
            Path: Đường dẫn output hợp lệ
        """
        if output_path is None:
            output_path = input_path.with_suffix(new_suffix)
        
        # Tạo thư mục parent nếu chưa có
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        return output_path
    
    @staticmethod
    def get_safe_filename(filename: str, max_length: int = 200) -> str:
        """
        Tạo filename an toàn (loại bỏ ký tự đặc biệt)
        
        Args:
            filename: Tên file gốc
            max_length: Độ dài tối đa
            
        Returns:
            str: Tên file an toàn
        """
        # Loại bỏ ký tự không hợp lệ
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        
        # Giới hạn độ dài
        if len(filename) > max_length:
            name, ext = os.path.splitext(filename)
            name = name[:max_length - len(ext) - 3] + '...'
            filename = name + ext
        
        return filename