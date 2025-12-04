"""
Frame nhập dữ liệu từ file
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from utils.file_reader import read_office_files
from gui.theme import MILITARY_COLORS, get_button_style, get_label_style


class ImportFrame(tk.Frame):
    """Frame nhập dữ liệu từ file"""
    
    def __init__(self, parent, db: DatabaseService):
        """
        Args:
            parent: Parent widget
            db: DatabaseService instance
        """
        super().__init__(parent)
        self.db = db
        self.personnel_list = []
        self.setup_ui()
    
    def setup_ui(self):
        """Thiết lập giao diện"""
        # Configure frame background
        self.configure(bg=MILITARY_COLORS['bg_light'])
        
        # Title
        title_frame = tk.Frame(self, bg=MILITARY_COLORS['primary'], height=50)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(
            title_frame,
            text="📥 NHẬP DỮ LIỆU TỪ FILE",
            font=('Arial', 16, 'bold'),
            bg=MILITARY_COLORS['primary'],
            fg=MILITARY_COLORS['text_light']
        )
        title_label.pack(expand=True)
        
        # Hướng dẫn
        guide_frame = tk.LabelFrame(self, text="Hướng Dẫn", font=('Arial', 12, 'bold'), padx=10, pady=10)
        guide_frame.pack(fill=tk.X, padx=10, pady=5)
        
        guide_text = """
1. Chọn thư mục chứa các file Word/Excel (thư mục 'noidung')
2. Click 'Chọn Thư Mục' để chọn thư mục
3. Click 'Đọc File' để đọc dữ liệu
4. Xem preview và click 'Import Tất Cả' để lưu vào database
        """
        tk.Label(guide_frame, text=guide_text.strip(), font=('Arial', 10), justify=tk.LEFT).pack(anchor=tk.W)
        
        # Chọn thư mục
        folder_frame = tk.LabelFrame(self, text="Chọn Thư Mục", font=('Arial', 12, 'bold'), padx=10, pady=10)
        folder_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.folder_path_var = tk.StringVar(value="../noidung")
        
        tk.Label(folder_frame, text="Đường dẫn:", font=('Arial', 10)).pack(side=tk.LEFT, padx=5)
        
        folder_entry = tk.Entry(folder_frame, textvariable=self.folder_path_var, width=50, font=('Arial', 10))
        folder_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        browse_btn = tk.Button(
            folder_frame,
            text="📂 Chọn Thư Mục",
            command=self.browse_folder,
            bg='#3498db',
            fg='white',
            font=('Arial', 10),
            padx=10,
            pady=5,
            cursor='hand2'
        )
        browse_btn.pack(side=tk.LEFT, padx=5)
        
        # Nút đọc file
        read_btn = tk.Button(
            folder_frame,
            text="📖 Đọc File",
            command=self.read_files,
            bg='#4CAF50',
            fg='white',
            font=('Arial', 10, 'bold'),
            padx=15,
            pady=5,
            cursor='hand2'
        )
        read_btn.pack(side=tk.LEFT, padx=5)
        
        # Preview
        preview_frame = tk.LabelFrame(self, text="Preview Dữ Liệu", font=('Arial', 12, 'bold'), padx=10, pady=10)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Treeview để hiển thị preview
        columns = ('STT', 'Họ và Tên', 'Cấp Bậc', 'Đơn Vị', 'Dân Tộc')
        self.preview_tree = ttk.Treeview(preview_frame, columns=columns, show='headings', height=10)
        
        for col in columns:
            self.preview_tree.heading(col, text=col)
            if col == 'Họ và Tên':
                self.preview_tree.column(col, width=200)
            else:
                self.preview_tree.column(col, width=150, anchor=tk.CENTER)
        
        # Scrollbar
        preview_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=preview_scrollbar.set)
        
        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        preview_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Nút import
        import_btn = tk.Button(
            self,
            text="💾 Import Tất Cả Vào Database",
            command=self.import_data,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 12, 'bold'),
            padx=20,
            pady=10,
            cursor='hand2'
        )
        import_btn.pack(pady=10)
    
    def browse_folder(self):
        """Chọn thư mục"""
        folder = filedialog.askdirectory(title="Chọn thư mục chứa file")
        if folder:
            self.folder_path_var.set(folder)
    
    def read_files(self):
        """Đọc file từ thư mục"""
        folder_path = self.folder_path_var.get().strip()
        
        if not folder_path:
            messagebox.showerror("Lỗi", "Vui lòng chọn thư mục")
            return
        
        folder = Path(folder_path)
        if not folder.exists():
            messagebox.showerror("Lỗi", f"Thư mục không tồn tại: {folder_path}")
            return
        
        try:
            # Đọc file
            self.personnel_list = read_office_files(str(folder))
            
            if not self.personnel_list:
                messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu trong các file")
                return
            
            # Hiển thị preview
            self.preview_tree.delete(*self.preview_tree.get_children())
            
            for idx, person in enumerate(self.personnel_list[:50], 1):  # Chỉ hiển thị 50 đầu
                self.preview_tree.insert('', tk.END, values=(
                    idx,
                    person.hoTen or 'Chưa có tên',
                    person.capBac or '',
                    person.donVi or '',
                    person.danToc or ''
                ))
            
            if len(self.personnel_list) > 50:
                messagebox.showinfo("Thông báo", f"Đã đọc {len(self.personnel_list)} hồ sơ. Hiển thị 50 đầu tiên.")
            else:
                messagebox.showinfo("Thành công", f"Đã đọc được {len(self.personnel_list)} hồ sơ")
                
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi đọc file: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def import_data(self):
        """Import dữ liệu vào database"""
        if not self.personnel_list:
            messagebox.showwarning("Cảnh báo", "Chưa có dữ liệu để import. Vui lòng đọc file trước.")
            return
        
        # Xác nhận
        result = messagebox.askyesno(
            "Xác nhận",
            f"Bạn có chắc muốn import {len(self.personnel_list)} hồ sơ vào database?"
        )
        
        if not result:
            return
        
        try:
            imported = 0
            skipped = 0
            
            for person in self.personnel_list:
                # Kiểm tra xem đã tồn tại chưa (theo tên)
                existing = self.db.search(person.hoTen or "")
                if existing and any(p.hoTen == person.hoTen and person.hoTen for p in existing):
                    skipped += 1
                    continue
                
                self.db.create(person)
                imported += 1
            
            messagebox.showinfo(
                "Thành công",
                f"Đã import {imported} hồ sơ.\nBỏ qua {skipped} hồ sơ trùng lặp."
            )
            
            # Xóa preview
            self.personnel_list = []
            self.preview_tree.delete(*self.preview_tree.get_children())
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi import: {str(e)}")
            import traceback
            traceback.print_exc()