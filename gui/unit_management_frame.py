"""
Frame quản lý đơn vị (Đại đội, Trung đội, Xe...)
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
# from models.unit import Unit  # Unit model not available
from gui.theme import MILITARY_COLORS


class UnitManagementFrame(tk.Frame):
    """Frame quản lý đơn vị"""
    
    def __init__(self, parent, db: DatabaseService):
        """
        Args:
            parent: Parent widget
            db: DatabaseService instance
        """
        super().__init__(parent)
        self.db = db
        self.bg_color = '#FAFAFA'
        self.setup_ui()
        self.load_units()
    
    def setup_ui(self):
        """Thiết lập giao diện"""
        self.configure(bg=self.bg_color)
        
        # Title
        title_frame = tk.Frame(self, bg='#388E3C', height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        tk.Label(
            title_frame,
            text="⚙️ QUẢN LÝ ĐƠN VỊ",
            font=('Segoe UI', 18, 'bold'),
            bg='#388E3C',
            fg='white'
        ).pack(expand=True, pady=18)
        
        # Toolbar
        toolbar = tk.Frame(self, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        tk.Button(
            toolbar,
            text="➕ Tạo Đại Đội",
            command=lambda: self.create_unit('dai_doi'),
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            toolbar,
            text="➕ Tạo Trung Đội",
            command=lambda: self.create_unit('trung_doi'),
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            toolbar,
            text="➕ Tạo Xe",
            command=lambda: self.create_unit('xe'),
            font=('Segoe UI', 10),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            toolbar,
            text="➕ Tạo Tổ",
            command=lambda: self.create_unit('to'),
            font=('Segoe UI', 10),
            bg='#9C27B0',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        # Split frame: bên trái là danh sách đơn vị, bên phải là danh sách quân nhân
        main_split = tk.Frame(self, bg=self.bg_color)
        main_split.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Bên trái: Danh sách đơn vị
        left_frame = tk.LabelFrame(main_split, text="📋 Danh Sách Đơn Vị", 
                                  bg=self.bg_color, fg='#388E3C', 
                                  font=('Segoe UI', 11, 'bold'))
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        tree_frame = tk.Frame(left_frame, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        columns = ('STT', 'Tên Đơn Vị', 'Loại', 'Số Quân Nhân', 'Ghi Chú')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=20)
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor=tk.W)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree = tree
        
        # Bind selection để hiển thị quân nhân
        tree.bind('<<TreeviewSelect>>', self.on_unit_select)
        
        # Bên phải: Danh sách quân nhân trong đơn vị
        right_frame = tk.LabelFrame(main_split, text="👥 Quân Nhân Trong Đơn Vị", 
                                    bg=self.bg_color, fg='#388E3C', 
                                    font=('Segoe UI', 11, 'bold'))
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        personnel_tree_frame = tk.Frame(right_frame, bg=self.bg_color)
        personnel_tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        personnel_columns = ('STT', 'Họ và Tên', 'Cấp Bậc', 'Chức Vụ', 'Ngày Sinh')
        self.personnel_tree = ttk.Treeview(personnel_tree_frame, columns=personnel_columns, 
                                          show='headings', height=20)
        
        for col in personnel_columns:
            self.personnel_tree.heading(col, text=col)
            self.personnel_tree.column(col, width=120, anchor=tk.W)
        
        self.personnel_tree.column('STT', width=50, anchor=tk.CENTER)
        self.personnel_tree.column('Họ và Tên', width=200)
        
        personnel_scrollbar = ttk.Scrollbar(personnel_tree_frame, orient=tk.VERTICAL, 
                                           command=self.personnel_tree.yview)
        self.personnel_tree.configure(yscrollcommand=personnel_scrollbar.set)
        
        self.personnel_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        personnel_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Label thông báo khi chưa chọn đơn vị
        self.personnel_info_label = tk.Label(
            right_frame,
            text="👉 Chọn một đơn vị để xem danh sách quân nhân",
            font=('Segoe UI', 10, 'italic'),
            bg=self.bg_color,
            fg='#666666'
        )
        self.personnel_info_label.pack(pady=20)
        
        # Buttons
        btn_frame = tk.Frame(self, bg=self.bg_color, pady=10)
        btn_frame.pack(fill=tk.X, padx=10)
        
        tk.Button(
            btn_frame,
            text="✏️ Sửa",
            command=self.edit_unit,
            font=('Segoe UI', 10),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="🗑️ Xóa",
            command=self.delete_unit,
            font=('Segoe UI', 10),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="👥 Quản Lý Quân Nhân",
            command=self.manage_personnel,
            font=('Segoe UI', 10),
            bg='#388E3C',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
    
    def load_units(self):
        """Load danh sách đơn vị"""
        # Xóa dữ liệu cũ
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Load từ database
        try:
            units = self.db.get_all_units()
            for idx, unit in enumerate(units, 1):
                self.tree.insert('', tk.END, iid=unit.id, values=(
                    idx,
                    unit.ten,
                    self._get_loai_name(unit.loai),
                    len(unit.personnelIds),
                    unit.ghiChu or ''
                ))
        except Exception as e:
            # Nếu chưa có hàm get_all_units, hiển thị thông báo
            messagebox.showinfo("Thông báo", "Chức năng quản lý đơn vị đang được phát triển")
    
    def on_unit_select(self, event):
        """Khi chọn đơn vị, hiển thị danh sách quân nhân"""
        selection = self.tree.selection()
        if not selection:
            self.clear_personnel_list()
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            self.clear_personnel_list()
            return
        
        # Ẩn label thông báo
        self.personnel_info_label.pack_forget()
        
        # Xóa danh sách cũ
        for item in self.personnel_tree.get_children():
            self.personnel_tree.delete(item)
        
        # Load quân nhân trong đơn vị
        try:
            personnel_list = self.db.get_personnel_by_unit(unit_id)
            
            if not personnel_list:
                self.personnel_info_label.config(
                    text=f"Đơn vị '{unit.ten}' chưa có quân nhân nào"
                )
                self.personnel_info_label.pack(pady=20)
                return
            
            # Hiển thị danh sách
            for idx, person in enumerate(personnel_list, 1):
                self.personnel_tree.insert('', tk.END, iid=person.id, values=(
                    idx,
                    person.hoTen or '',
                    person.capBac or '',
                    person.chucVu or '',
                    person.ngaySinh or ''
                ))
        except Exception as e:
            self.personnel_info_label.config(
                text=f"Lỗi khi load quân nhân: {str(e)}"
            )
            self.personnel_info_label.pack(pady=20)
    
    def clear_personnel_list(self):
        """Xóa danh sách quân nhân"""
        for item in self.personnel_tree.get_children():
            self.personnel_tree.delete(item)
        
        self.personnel_info_label.config(
            text="👉 Chọn một đơn vị để xem danh sách quân nhân"
        )
        self.personnel_info_label.pack(pady=20)
    
    def _get_loai_name(self, loai: str) -> str:
        """Chuyển loại sang tên hiển thị"""
        mapping = {
            'dai_doi': 'Đại Đội',
            'trung_doi': 'Trung Đội',
            'xe': 'Xe',
            'to': 'Tổ'
        }
        return mapping.get(loai, loai)
    
    def create_unit(self, loai: str):
        """Tạo đơn vị mới"""
        dialog = tk.Toplevel(self)
        dialog.title(f"Tạo {self._get_loai_name(loai)}")
        dialog.geometry("400x200")
        dialog.configure(bg=self.bg_color)
        
        tk.Label(
            dialog,
            text=f"Tên {self._get_loai_name(loai)}:",
            font=('Segoe UI', 11),
            bg=self.bg_color
        ).pack(pady=10)
        
        name_entry = tk.Entry(dialog, font=('Segoe UI', 11), width=30)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        def save():
            ten = name_entry.get().strip()
            if not ten:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập tên đơn vị")
                return
            
            try:
                unit = Unit(ten=ten, loai=loai)
                self.db.create_unit(unit)
                messagebox.showinfo("Thành công", f"Đã tạo {self._get_loai_name(loai)}: {ten}")
                dialog.destroy()
                self.load_units()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể tạo đơn vị:\n{str(e)}")
        
        tk.Button(
            dialog,
            text="Lưu",
            command=save,
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5
        ).pack(pady=20)
        
        name_entry.bind('<Return>', lambda e: save())
    
    def edit_unit(self):
        """Sửa đơn vị"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn đơn vị cần sửa")
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            messagebox.showerror("Lỗi", "Không tìm thấy đơn vị")
            return
        
        # TODO: Mở dialog sửa đơn vị
        messagebox.showinfo("Thông báo", "Chức năng sửa đơn vị đang được phát triển")
    
    def delete_unit(self):
        """Xóa đơn vị"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn đơn vị cần xóa")
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            return
        
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa đơn vị '{unit.ten}'?"):
            try:
                self.db.delete_unit(unit_id)
                messagebox.showinfo("Thành công", "Đã xóa đơn vị")
                self.load_units()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xóa đơn vị:\n{str(e)}")
    
    def manage_personnel(self):
        """Quản lý quân nhân trong đơn vị"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn đơn vị")
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            messagebox.showerror("Lỗi", "Không tìm thấy đơn vị")
            return
        
        # Mở dialog chọn quân nhân
        dialog = tk.Toplevel(self)
        dialog.title(f"Quản Lý Quân Nhân - {unit.ten}")
        dialog.geometry("700x600")
        dialog.configure(bg=self.bg_color)
        
        # Title
        tk.Label(
            dialog,
            text=f"Chọn quân nhân cho: {unit.ten}",
            font=('Segoe UI', 14, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        ).pack(pady=10)
        
        # Treeview với checkbox
        tree_frame = tk.Frame(dialog, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        tree = ttk.Treeview(
            tree_frame,
            columns=('hoTen', 'capBac', 'chucVu'),
            show='tree headings',
            yscrollcommand=scrollbar.set,
            height=20
        )
        
        tree.heading('#0', text='Chọn')
        tree.heading('hoTen', text='Họ và Tên')
        tree.heading('capBac', text='Cấp Bậc')
        tree.heading('chucVu', text='Chức Vụ')
        
        tree.column('#0', width=50)
        tree.column('hoTen', width=250)
        tree.column('capBac', width=120)
        tree.column('chucVu', width=150)
        
        scrollbar.config(command=tree.yview)
        
        # Load tất cả quân nhân
        all_personnel = self.db.get_all()
        selected_ids = set(unit.personnelIds)
        
        for person in all_personnel:
            is_selected = person.id in selected_ids
            tree.insert('', tk.END, iid=person.id, 
                       text='✓' if is_selected else '',
                       values=(person.hoTen or '', person.capBac or '', person.chucVu or ''))
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind click để toggle
        def toggle_selection(event):
            item = tree.identify_row(event.y)
            if item:
                current_text = tree.item(item, 'text')
                if current_text == '✓':
                    tree.item(item, text='')
                else:
                    tree.item(item, text='✓')
        
        tree.bind('<Button-1>', toggle_selection)
        
        # Buttons
        btn_frame = tk.Frame(dialog, bg=self.bg_color, pady=10)
        btn_frame.pack(fill=tk.X, padx=10)
        
        def save():
            # Lấy danh sách ID đã chọn
            selected_personnel_ids = []
            for item in tree.get_children():
                if tree.item(item, 'text') == '✓':
                    selected_personnel_ids.append(item)
            
            # Cập nhật unit
            unit.personnelIds = selected_personnel_ids
            try:
                self.db.update_unit(unit)
                
                # Cập nhật unitId cho quân nhân
                all_personnel = self.db.get_all()
                for person in all_personnel:
                    if person.id in selected_personnel_ids:
                        person.unitId = unit.id
                    elif person.unitId == unit.id:
                        person.unitId = None
                    self.db.update(person)
                
                messagebox.showinfo("Thành công", f"Đã cập nhật {len(selected_personnel_ids)} quân nhân vào đơn vị")
                dialog.destroy()
                self.load_units()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể cập nhật:\n{str(e)}")
        
        tk.Button(
            btn_frame,
            text="💾 Lưu",
            command=save,
            font=('Segoe UI', 10, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5
        ).pack(side=tk.LEFT, padx=5)