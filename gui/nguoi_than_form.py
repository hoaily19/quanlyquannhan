"""
Form thêm/sửa người thân
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from models.nguoi_than import NguoiThan
from gui.date_picker import DatePicker
from gui.theme import MILITARY_COLORS


class NguoiThanFormDialog:
    """Dialog form thêm/sửa người thân"""
    
    def __init__(self, parent, db: DatabaseService, personnel_id: str, nguoi_than_id: str = None):
        """
        Args:
            parent: Parent widget
            db: DatabaseService instance
            personnel_id: ID quân nhân
            nguoi_than_id: ID người thân nếu là sửa
        """
        self.parent = parent
        self.db = db
        self.personnel_id = personnel_id
        self.nguoi_than_id = nguoi_than_id
        self.is_new = nguoi_than_id is None
        
        if self.is_new:
            self.nguoi_than = NguoiThan(personnelId=personnel_id)
        else:
            self.nguoi_than = db.get_nguoi_than_by_id(nguoi_than_id)
            if not self.nguoi_than:
                self.nguoi_than = NguoiThan(personnelId=personnel_id)
                self.is_new = True
        
        self.bg_color = '#FAFAFA'
        self.section_bg = '#FFFFFF'
        self.border_color = '#E0E0E0'
        self.text_color = '#424242'
        self.title_color = '#388E3C'
    
    def show(self):
        """Hiển thị dialog"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Thêm Người Thân" if self.is_new else "Sửa Người Thân")
        self.dialog.geometry("600x700")
        self.dialog.configure(bg=self.bg_color)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Title
        title_frame = tk.Frame(self.dialog, bg=self.title_color, height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        tk.Label(
            title_frame,
            text="➕ Thêm Người Thân" if self.is_new else "✏️ Sửa Người Thân",
            font=('Segoe UI', 16, 'bold'),
            bg=self.title_color,
            fg='white'
        ).pack(expand=True, pady=18)
        
        # Canvas với scrollbar để form có thể scroll
        canvas_frame = tk.Frame(self.dialog, bg=self.bg_color)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(canvas_frame, bg=self.bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.bg_color)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind mouse wheel
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Form
        form_frame = scrollable_frame
        
        # Họ và tên
        tk.Label(
            form_frame,
            text="Họ và Tên *",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg='#E53935'
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.ho_ten_var = tk.StringVar(value=self.nguoi_than.hoTen or "")
        ho_ten_entry = tk.Entry(
            form_frame,
            textvariable=self.ho_ten_var,
            font=('Segoe UI', 11),
            relief=tk.FLAT,
            bd=1,
            bg=self.section_bg,
            highlightthickness=1,
            highlightcolor=self.title_color,
            highlightbackground=self.border_color
        )
        ho_ten_entry.pack(fill=tk.X, pady=(0, 15), ipady=8)
        
        # Ngày sinh
        tk.Label(
            form_frame,
            text="Ngày Sinh",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(anchor=tk.W, pady=(0, 5))
        
        date_frame = tk.Frame(form_frame, bg=self.bg_color)
        date_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.ngay_sinh_picker = DatePicker(date_frame, self.nguoi_than.ngaySinh or "")
        self.ngay_sinh_picker.pack(fill=tk.X)
        
        # Địa chỉ
        tk.Label(
            form_frame,
            text="Địa Chỉ",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.dia_chi_var = tk.StringVar(value=self.nguoi_than.diaChi or "")
        dia_chi_entry = tk.Entry(
            form_frame,
            textvariable=self.dia_chi_var,
            font=('Segoe UI', 11),
            relief=tk.FLAT,
            bd=1,
            bg=self.section_bg,
            highlightthickness=1,
            highlightcolor=self.title_color,
            highlightbackground=self.border_color
        )
        dia_chi_entry.pack(fill=tk.X, pady=(0, 15), ipady=8)
        
        # Số điện thoại
        tk.Label(
            form_frame,
            text="Số Điện Thoại",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.sdt_var = tk.StringVar(value=self.nguoi_than.soDienThoai or "")
        sdt_entry = tk.Entry(
            form_frame,
            textvariable=self.sdt_var,
            font=('Segoe UI', 11),
            relief=tk.FLAT,
            bd=1,
            bg=self.section_bg,
            highlightthickness=1,
            highlightcolor=self.title_color,
            highlightbackground=self.border_color
        )
        sdt_entry.pack(fill=tk.X, pady=(0, 15), ipady=8)
        
        # Mối quan hệ
        tk.Label(
            form_frame,
            text="Mối Quan Hệ",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.moi_quan_he_var = tk.StringVar(value=self.nguoi_than.moiQuanHe or "")
        moi_quan_he_combo = ttk.Combobox(
            form_frame,
            textvariable=self.moi_quan_he_var,
            values=['Bố', 'Mẹ', 'Anh', 'Chị', 'Em', 'Vợ', 'Chồng', 'Con', 'Khác'],
            font=('Segoe UI', 11),
            state='readonly'
        )
        moi_quan_he_combo.pack(fill=tk.X, pady=(0, 15), ipady=8)
        
        # Nội dung
        tk.Label(
            form_frame,
            text="Nội Dung",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.noi_dung_text = tk.Text(
            form_frame,
            font=('Segoe UI', 10),
            relief=tk.FLAT,
            bd=1,
            bg=self.section_bg,
            highlightthickness=1,
            highlightcolor=self.title_color,
            highlightbackground=self.border_color,
            wrap=tk.WORD,
            height=4
        )
        self.noi_dung_text.pack(fill=tk.X, pady=(0, 15))
        self.noi_dung_text.insert('1.0', self.nguoi_than.noiDung or "")
        
        # Buttons - đặt ở cuối form
        btn_frame = tk.Frame(form_frame, bg=self.bg_color)
        btn_frame.pack(fill=tk.X, pady=(20, 20))
        
        tk.Button(
            btn_frame,
            text="💾 Lưu",
            command=self.save,
            font=('Segoe UI', 11, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=30,
            pady=10,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=self.close_dialog,
            font=('Segoe UI', 11),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=30,
            pady=10,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
    
    def close_dialog(self):
        """Đóng dialog một cách an toàn"""
        try:
            # Release grab trước khi destroy
            self.dialog.grab_release()
        except:
            pass
        try:
            self.dialog.destroy()
        except:
            pass
    
    def save(self):
        """Lưu người thân"""
        if not self.ho_ten_var.get().strip():
            messagebox.showerror("Lỗi", "Vui lòng nhập Họ và Tên")
            return
        
        # Cập nhật object
        self.nguoi_than.hoTen = self.ho_ten_var.get().strip()
        self.nguoi_than.ngaySinh = self.ngay_sinh_picker.get_date()
        self.nguoi_than.diaChi = self.dia_chi_var.get().strip()
        self.nguoi_than.soDienThoai = self.sdt_var.get().strip()
        self.nguoi_than.moiQuanHe = self.moi_quan_he_var.get().strip()
        self.nguoi_than.noiDung = self.noi_dung_text.get('1.0', tk.END).strip()
        self.nguoi_than.personnelId = self.personnel_id
        
        try:
            if self.is_new:
                self.db.create_nguoi_than(self.nguoi_than)
                messagebox.showinfo("Thành công", "Đã thêm người thân")
            else:
                self.nguoi_than.id = self.nguoi_than_id
                self.db.update_nguoi_than(self.nguoi_than)
                messagebox.showinfo("Thành công", "Đã cập nhật người thân")
            
            self.close_dialog()
            # Trigger callback để parent reload danh sách
            if hasattr(self.parent, 'load_nguoi_than_list'):
                self.parent.load_nguoi_than_list()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lưu:\n{str(e)}")

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from models.nguoi_than import NguoiThan
from gui.date_picker import DatePicker
from gui.theme import MILITARY_COLORS


class NguoiThanFormDialog:
    """Dialog form thêm/sửa người thân"""
    
    def __init__(self, parent, db: DatabaseService, personnel_id: str, nguoi_than_id: str = None):
        """
        Args:
            parent: Parent widget
            db: DatabaseService instance
            personnel_id: ID quân nhân
            nguoi_than_id: ID người thân nếu là sửa
        """
        self.parent = parent
        self.db = db
        self.personnel_id = personnel_id
        self.nguoi_than_id = nguoi_than_id
        self.is_new = nguoi_than_id is None
        
        if self.is_new:
            self.nguoi_than = NguoiThan(personnelId=personnel_id)
        else:
            self.nguoi_than = db.get_nguoi_than_by_id(nguoi_than_id)
            if not self.nguoi_than:
                self.nguoi_than = NguoiThan(personnelId=personnel_id)
                self.is_new = True
        
        self.bg_color = '#FAFAFA'
        self.section_bg = '#FFFFFF'
        self.border_color = '#E0E0E0'
        self.text_color = '#424242'
        self.title_color = '#388E3C'
    
    def show(self):
        """Hiển thị dialog"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Thêm Người Thân" if self.is_new else "Sửa Người Thân")
        self.dialog.geometry("600x700")
        self.dialog.configure(bg=self.bg_color)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Title
        title_frame = tk.Frame(self.dialog, bg=self.title_color, height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        tk.Label(
            title_frame,
            text="➕ Thêm Người Thân" if self.is_new else "✏️ Sửa Người Thân",
            font=('Segoe UI', 16, 'bold'),
            bg=self.title_color,
            fg='white'
        ).pack(expand=True, pady=18)
        
        # Canvas với scrollbar để form có thể scroll
        canvas_frame = tk.Frame(self.dialog, bg=self.bg_color)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(canvas_frame, bg=self.bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.bg_color)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind mouse wheel
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Form
        form_frame = scrollable_frame
        
        # Họ và tên
        tk.Label(
            form_frame,
            text="Họ và Tên *",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg='#E53935'
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.ho_ten_var = tk.StringVar(value=self.nguoi_than.hoTen or "")
        ho_ten_entry = tk.Entry(
            form_frame,
            textvariable=self.ho_ten_var,
            font=('Segoe UI', 11),
            relief=tk.FLAT,
            bd=1,
            bg=self.section_bg,
            highlightthickness=1,
            highlightcolor=self.title_color,
            highlightbackground=self.border_color
        )
        ho_ten_entry.pack(fill=tk.X, pady=(0, 15), ipady=8)
        
        # Ngày sinh
        tk.Label(
            form_frame,
            text="Ngày Sinh",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(anchor=tk.W, pady=(0, 5))
        
        date_frame = tk.Frame(form_frame, bg=self.bg_color)
        date_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.ngay_sinh_picker = DatePicker(date_frame, self.nguoi_than.ngaySinh or "")
        self.ngay_sinh_picker.pack(fill=tk.X)
        
        # Địa chỉ
        tk.Label(
            form_frame,
            text="Địa Chỉ",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.dia_chi_var = tk.StringVar(value=self.nguoi_than.diaChi or "")
        dia_chi_entry = tk.Entry(
            form_frame,
            textvariable=self.dia_chi_var,
            font=('Segoe UI', 11),
            relief=tk.FLAT,
            bd=1,
            bg=self.section_bg,
            highlightthickness=1,
            highlightcolor=self.title_color,
            highlightbackground=self.border_color
        )
        dia_chi_entry.pack(fill=tk.X, pady=(0, 15), ipady=8)
        
        # Số điện thoại
        tk.Label(
            form_frame,
            text="Số Điện Thoại",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.sdt_var = tk.StringVar(value=self.nguoi_than.soDienThoai or "")
        sdt_entry = tk.Entry(
            form_frame,
            textvariable=self.sdt_var,
            font=('Segoe UI', 11),
            relief=tk.FLAT,
            bd=1,
            bg=self.section_bg,
            highlightthickness=1,
            highlightcolor=self.title_color,
            highlightbackground=self.border_color
        )
        sdt_entry.pack(fill=tk.X, pady=(0, 15), ipady=8)
        
        # Mối quan hệ
        tk.Label(
            form_frame,
            text="Mối Quan Hệ",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.moi_quan_he_var = tk.StringVar(value=self.nguoi_than.moiQuanHe or "")
        moi_quan_he_combo = ttk.Combobox(
            form_frame,
            textvariable=self.moi_quan_he_var,
            values=['Bố', 'Mẹ', 'Anh', 'Chị', 'Em', 'Vợ', 'Chồng', 'Con', 'Khác'],
            font=('Segoe UI', 11),
            state='readonly'
        )
        moi_quan_he_combo.pack(fill=tk.X, pady=(0, 15), ipady=8)
        
        # Nội dung
        tk.Label(
            form_frame,
            text="Nội Dung",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.noi_dung_text = tk.Text(
            form_frame,
            font=('Segoe UI', 10),
            relief=tk.FLAT,
            bd=1,
            bg=self.section_bg,
            highlightthickness=1,
            highlightcolor=self.title_color,
            highlightbackground=self.border_color,
            wrap=tk.WORD,
            height=4
        )
        self.noi_dung_text.pack(fill=tk.X, pady=(0, 15))
        self.noi_dung_text.insert('1.0', self.nguoi_than.noiDung or "")
        
        # Buttons - đặt ở cuối form
        btn_frame = tk.Frame(form_frame, bg=self.bg_color)
        btn_frame.pack(fill=tk.X, pady=(20, 20))
        
        tk.Button(
            btn_frame,
            text="💾 Lưu",
            command=self.save,
            font=('Segoe UI', 11, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=30,
            pady=10,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=self.close_dialog,
            font=('Segoe UI', 11),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=30,
            pady=10,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
    
    def close_dialog(self):
        """Đóng dialog một cách an toàn"""
        try:
            # Release grab trước khi destroy
            self.dialog.grab_release()
        except:
            pass
        try:
            self.dialog.destroy()
        except:
            pass
    
    def save(self):
        """Lưu người thân"""
        if not self.ho_ten_var.get().strip():
            messagebox.showerror("Lỗi", "Vui lòng nhập Họ và Tên")
            return
        
        # Cập nhật object
        self.nguoi_than.hoTen = self.ho_ten_var.get().strip()
        self.nguoi_than.ngaySinh = self.ngay_sinh_picker.get_date()
        self.nguoi_than.diaChi = self.dia_chi_var.get().strip()
        self.nguoi_than.soDienThoai = self.sdt_var.get().strip()
        self.nguoi_than.moiQuanHe = self.moi_quan_he_var.get().strip()
        self.nguoi_than.noiDung = self.noi_dung_text.get('1.0', tk.END).strip()
        self.nguoi_than.personnelId = self.personnel_id
        
        try:
            if self.is_new:
                self.db.create_nguoi_than(self.nguoi_than)
                messagebox.showinfo("Thành công", "Đã thêm người thân")
            else:
                self.nguoi_than.id = self.nguoi_than_id
                self.db.update_nguoi_than(self.nguoi_than)
                messagebox.showinfo("Thành công", "Đã cập nhật người thân")
            
            self.close_dialog()
            # Trigger callback để parent reload danh sách
            if hasattr(self.parent, 'load_nguoi_than_list'):
                self.parent.load_nguoi_than_list()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lưu:\n{str(e)}")