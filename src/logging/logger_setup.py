"""
Thiết lập logging cho ứng dụng
"""
import logging
import os
from pathlib import Path
from datetime import datetime


def setup_logger(name: str = "converter", log_dir: str = "logs") -> logging.Logger:
    """
    Thiết lập logger với cả file và console output
    
    Args:
        name: Tên logger
        log_dir: Thư mục chứa log files
        
    Returns:
        logging.Logger: Logger đã được cấu hình
    """
    # Tạo thư mục logs nếu chưa có
    log_path = Path(log_dir)
    log_path.mkdir(exist_ok=True)
    
    # Tạo logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    # Tránh duplicate handlers
    if logger.handlers:
        return logger
    
    # Format cho log messages
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # File handler - ghi vào file
    log_file = log_path / f"converter_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Console handler - hiển thị trên terminal
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger


def get_logger(name: str = "converter") -> logging.Logger:
    """
    Lấy logger đã được setup hoặc tạo mới nếu chưa có
    
    Args:
        name: Tên logger
        
    Returns:
        logging.Logger: Logger instance
    """
    logger = logging.getLogger(name)
    if not logger.handlers:
        return setup_logger(name)
    return logger