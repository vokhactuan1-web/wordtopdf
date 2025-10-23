"""
Giao diện Tkinter cho ứng dụng converter
"""
import threading
from pathlib import Path
from typing import List, Callable, Optional
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

from ..logging.logger_setup import get_logger
from ..io.file_handler import FileHandler

logger = get_logger(__name__)


class ConverterUI:
    """Giao diện chung cho converter"""
    
    def __init__(self, root: tk.Tk, title: str, file_types: List[tuple], 
                 patterns: List[str], converter_func: Callable):
        """
        Args:
            root: Tkinter root window
            title: Tiêu đề ứng dụng
            file_types: Danh sách file types cho dialog
            patterns: Danh sách pattern cho file search
            converter_func: Hàm chuyển đổi (input_path, output_path) -> Path
        """
        self.root = root
        self.title = title
        self.file_types = file_types
        self.patterns = patterns
        self.converter_func = converter_func
        
        self.file_list: List[Path] = []
        
        # FIX: Tạo valid_extensions đúng cách - loại bỏ dấu * và chỉ lấy extension
        self.valid_extensions = []
        for _, pattern in file_types:
            # Split pattern và xử lý từng extension
            for ext in pattern.split():
                # Loại bỏ dấu * và thêm vào list
                clean_ext = ext.replace('*', '').lower()
                if clean_ext and clean_ext not in self.valid_extensions:
                    self.valid_extensions.append(clean_ext)
        
        logger.info(f"Valid extensions: {self.valid_extensions}")
        
        self._setup_window()
        self._create_widgets()
        self._log_system_info()
    
    def _setup_window(self):
        """Thiết lập cửa sổ"""
        self.root.title(self.title)
        self.root.geometry("900x700")
        
        # Icon (nếu có)
        try:
            # self.root.iconbitmap('icon.ico')
            pass
        except:
            pass
    
    def _create_widgets(self):
        """Tạo các widget"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(
            main_frame,
            text=f"📄 {self.title}",
            font=('Arial', 18, 'bold')
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        subtitle = ttk.Label(
            main_frame,
            text="✨ Hỗ trợ hoàn toàn tiếng Việt - Tự động tải font Unicode",
            font=('Arial', 9, 'italic'),
            foreground='#27AE60'
        )
        subtitle.grid(row=1, column=0, columnspan=3, pady=(0, 15))
        
        # Drop zone
        self._create_drop_zone(main_frame)
        
        # Buttons
        self._create_buttons(main_frame)
        
        # File list
        self._create_file_list(main_frame)
        
        # Convert button
        self._create_convert_button(main_frame)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Log area
        self._create_log_area(main_frame)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        main_frame.rowconfigure(7, weight=2)
        
        # Setup drag & drop
        self._setup_drag_drop()
    
    def _create_drop_zone(self, parent):
        """Tạo vùng drop file"""
        drop_frame = ttk.LabelFrame(parent, text="Thêm file", padding="25")
        drop_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        self.drop_label = ttk.Label(
            drop_frame,
            text=f"🖱️ Kéo thả file vào đây\nhoặc nhấn nút bên dưới",
            font=('Arial', 11),
            justify=tk.CENTER
        )
        self.drop_label.pack(expand=True, fill=tk.BOTH, pady=15)
    
    def _create_buttons(self, parent):
        """Tạo các nút chức năng"""
        btn_frame = ttk.Frame(parent)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        ttk.Button(
            btn_frame,
            text="📂 Chọn File",
            command=self.select_files,
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="📁 Chọn Thư Mục",
            command=self.select_folder,
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="🗑️ Xóa",
            command=self.clear_list,
            width=10
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="📥 Mở Downloads",
            command=self.open_downloads,
            width=15
        ).pack(side=tk.LEFT, padx=5)
    
    def _create_file_list(self, parent):
        """Tạo danh sách file"""
        list_frame = ttk.LabelFrame(parent, text="📋 Danh sách file", padding="10")
        list_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        scroll = ttk.Scrollbar(list_frame)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(
            list_frame,
            height=8,
            yscrollcommand=scroll.set,
            font=('Consolas', 9)
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.config(command=self.file_listbox.yview)
    
    def _create_convert_button(self, parent):
        """Tạo nút chuyển đổi"""
        self.convert_btn = ttk.Button(
            parent,
            text="🔄 CHUYỂN ĐỔI SANG PDF",
            command=self.convert_files
        )
        self.convert_btn.grid(row=5, column=0, columnspan=3, pady=15, ipadx=20, ipady=5)
    
    def _create_log_area(self, parent):
        """Tạo vùng log"""
        log_frame = ttk.LabelFrame(parent, text="📝 Log", padding="10")
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=10,
            state='disabled',
            font=('Consolas', 9),
            wrap=tk.WORD
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
    
    def _setup_drag_drop(self):
        """Thiết lập drag & drop"""
        try:
            from tkinterdnd2 import DND_FILES
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind('<<Drop>>', self.on_drop)
            logger.info("Drag & drop đã được kích hoạt")
        except ImportError:
            logger.warning("tkinterdnd2 chưa cài đặt - không có drag & drop")
    
    def _log_system_info(self):
        """Ghi thông tin hệ thống"""
        self.log("=" * 60)
        self.log(f"📊 {self.title.upper()}")
        self.log("=" * 60)
        self.log("")
        self.log("🚀 Sẵn sàng chuyển đổi!")
        self.log("=" * 60)
        self.log("")
    
    def log(self, msg: str):
        """Ghi log"""
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"{msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        logger.info(msg)
    
    def on_drop(self, event):
        """Xử lý drop file"""
        files = self.root.tk.splitlist(event.data)
        for f in files:
            path = Path(f)
            if FileHandler.validate_file(path, self.valid_extensions):
                self.add_file(path)
    
    def select_files(self):
        """Chọn file"""
        files = FileHandler.select_files(self.file_types, "Chọn file")
        for f in files:
            self.add_file(f)
    
    def select_folder(self):
        """Chọn thư mục"""
        folder = FileHandler.select_folder()
        if folder:
            files = FileHandler.get_files_from_folder(folder, self.patterns)
            for f in files:
                self.add_file(f)
    
    def add_file(self, path: Path):
        """Thêm file vào danh sách"""
        if path not in self.file_list:
            if FileHandler.validate_file(path, self.valid_extensions):
                self.file_list.append(path)
                self.file_listbox.insert(tk.END, str(path))
                self.log(f"➕ {path.name}")
            else:
                self.log(f"⚠️ File không hợp lệ: {path.name}")
                self.log(f"   Extension: {path.suffix} | Cho phép: {self.valid_extensions}")
    
    def clear_list(self):
        """Xóa danh sách"""
        self.file_list.clear()
        self.file_listbox.delete(0, tk.END)
        self.log("🗑️ Đã xóa danh sách\n")
    
    def open_downloads(self):
        """Mở thư mục Downloads"""
        downloads = FileHandler.get_downloads_folder()
        FileHandler.open_folder(downloads)
        self.log(f"📥 Đã mở: {downloads}")
    
    def convert_files(self):
        """Chuyển đổi các file"""
        if not self.file_list:
            messagebox.showwarning("Cảnh báo", "Chưa có file nào!")
            return
        
        self.convert_btn.config(state='disabled')
        self.progress.start()
        
        threading.Thread(target=self._convert_thread, daemon=True).start()
    
    def _convert_thread(self):
        """Thread chuyển đổi"""
        success = error = 0
        
        self.log("\n" + "=" * 60)
        self.log(f"🚀 BẮT ĐẦU CHUYỂN ĐỔI {len(self.file_list)} FILE")
        self.log("=" * 60 + "\n")
        
        for file_path in self.file_list:
            try:
                self.log(f"⏳ Đang xử lý: {file_path.name}")
                output = file_path.with_suffix('.pdf')
                
                # Gọi hàm converter
                result = self.converter_func(file_path, output)
                
                self.log(f"   ✅ → {result.name}\n")
                success += 1
            except Exception as e:
                self.log(f"   ❌ LỖI: {str(e)}\n")
                logger.error(f"Lỗi chuyển đổi {file_path}: {e}", exc_info=True)
                error += 1
        
        self.progress.stop()
        self.convert_btn.config(state='normal')
        
        self.log("=" * 60)
        self.log(f"🎉 KẾT QUẢ: ✅ {success} | ❌ {error}")
        self.log("=" * 60 + "\n")
        
        messagebox.showinfo(
            "Hoàn tất",
            f"✅ Thành công: {success}\n❌ Lỗi: {error}"
        )


def create_app(title: str, file_types: List[tuple], patterns: List[str],
               converter_func: Callable) -> ConverterUI:
    """
    Tạo ứng dụng converter
    
    Args:
        title: Tiêu đề
        file_types: Danh sách file types
        patterns: Danh sách pattern
        converter_func: Hàm chuyển đổi
        
    Returns:
        ConverterUI: UI instance
    """
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()
        logger.warning("tkinterdnd2 chưa cài đặt - không có drag & drop")
    
    app = ConverterUI(root, title, file_types, patterns, converter_func)
    return app