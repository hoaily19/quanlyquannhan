"""
Form nhập thông tin yếu tố nước ngoài và đào tạo
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from models.personnel import Personnel
from services.database import DatabaseService
from gui.theme import MILITARY_COLORS, get_button_style


class YeuToNuocNgoaiFormDialog:
    """Dialog form nhập thông tin yếu tố nước ngoài và đào tạo"""
    
    def __init__(self, parent, db: DatabaseService, personnel: Personnel):
        """
        Args:
            parent: Parent window
            db: DatabaseService instance
            personnel: Personnel object cần chỉnh sửa
        """
        self.parent = parent
        self.db = db
        self.personnel = personnel
        self.result = False
        
        # Tạo dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Thông Tin Yếu Tố Nước Ngoài - {personnel.hoTen}")
        self.dialog.geometry("800x700")
        self.dialog.resizable(True, True)
        self.dialog.configure(bg='#FAFAFA')
        
        # Đảm bảo dialog hiển thị trên cùng
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")
        
        self.setup_ui()
        self.load_data()
    
    def setup_ui(self):
        """Thiết lập giao diện"""
        # Title
        title_frame = tk.Frame(self.dialog, bg=MILITARY_COLORS['primary'], height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        tk.Label(
            title_frame,
            text=f"🌍 THÔNG TIN YẾU TỐ NƯỚC NGOÀI - {self.personnel.hoTen}",
            font=('Arial', 14, 'bold'),
            bg=MILITARY_COLORS['primary'],
            fg='white'
        ).pack(expand=True, pady=15)
        
        # Buttons - Pack trước để ở dưới cùng
        btn_frame = tk.Frame(self.dialog, bg='#FAFAFA', pady=15)
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # Lấy style và override font
        save_style = get_button_style('success')
        save_style['font'] = ('Arial', 11, 'bold')
        save_btn = tk.Button(
            btn_frame,
            text="💾 Lưu",
            command=self.save,
            width=15,
            **save_style
        )
        save_btn.pack(side=tk.LEFT, padx=10, expand=True)
        
        # Lấy style và override font
        cancel_style = get_button_style('danger')
        cancel_style['font'] = ('Arial', 11, 'bold')
        cancel_btn = tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=self.cancel,
            width=15,
            **cancel_style
        )
        cancel_btn.pack(side=tk.LEFT, padx=10, expand=True)
        
        # Scrollable content - Pack sau buttons
        canvas = tk.Canvas(self.dialog, bg='#FAFAFA', highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#FAFAFA')
        
        # Tạo window trong canvas
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Form fields - Tạo trước khi bind events
        self.create_form_fields(scrollable_frame)
        
        # Bind để resize scrollable_frame theo canvas width
        def configure_scroll_region(event):
            canvas_width = event.width
            if canvas_width > 1:
                canvas.itemconfig(canvas_window, width=canvas_width)
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def on_frame_configure(event):
            """Cập nhật scrollregion khi frame thay đổi kích thước"""
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        scrollable_frame.bind("<Configure>", on_frame_configure)
        canvas.bind('<Configure>', configure_scroll_region)
        
        # Pack canvas và scrollbar
        canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scrollbar.pack(side="right", fill="y", padx=(0, 5), pady=5)
        
        # Update để đảm bảo scrollregion được tính đúng
        self.dialog.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))
        
        # Bind mouse wheel
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind("<MouseWheel>", on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", on_mousewheel)
        self.dialog.bind("<MouseWheel>", on_mousewheel)
    
    def create_form_fields(self, parent):
        """Tạo các trường form"""
        # Section 1: Yếu tố nước ngoài
        section1 = tk.LabelFrame(
            parent,
            text="🌍 Thông Tin Yếu Tố Nước Ngoài",
            font=('Arial', 12, 'bold'),
            bg='#FAFAFA',
            fg=MILITARY_COLORS['primary_dark'],
            padx=15,
            pady=15
        )
        section1.pack(fill=tk.X, padx=20, pady=10)
        
        # Nội dung yếu tố nước ngoài
        tk.Label(
            section1,
            text="Nội Dung Yếu Tố Nước Ngoài:",
            font=('Arial', 10, 'bold'),
            bg='#FAFAFA',
            fg='#424242'
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.noi_dung_var = tk.StringVar()
        noi_dung_entry = tk.Text(section1, height=4, width=70, font=('Arial', 10), wrap=tk.WORD)
        noi_dung_entry.pack(fill=tk.X, padx=5, pady=5)
        self.noi_dung_entry = noi_dung_entry
        
        # Mối quan hệ
        tk.Label(
            section1,
            text="Mối Quan Hệ:",
            font=('Arial', 10, 'bold'),
            bg='#FAFAFA',
            fg='#424242'
        ).pack(anchor=tk.W, pady=(10, 5))
        
        self.moi_quan_he_var = tk.StringVar()
        moi_quan_he_entry = tk.Entry(
            section1,
            textvariable=self.moi_quan_he_var,
            font=('Arial', 10),
            width=50
        )
        moi_quan_he_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # Tên nước
        tk.Label(
            section1,
            text="Tên Nước:",
            font=('Arial', 10, 'bold'),
            bg='#FAFAFA',
            fg='#424242'
        ).pack(anchor=tk.W, pady=(10, 5))
        
        self.ten_nuoc_var = tk.StringVar()
        ten_nuoc_entry = tk.Entry(
            section1,
            textvariable=self.ten_nuoc_var,
            font=('Arial', 10),
            width=50
        )
        ten_nuoc_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # Section 2: Thông tin đào tạo
        section2 = tk.LabelFrame(
            parent,
            text="🎓 Thông Tin Đào Tạo",
            font=('Arial', 12, 'bold'),
            bg='#FAFAFA',
            fg=MILITARY_COLORS['primary_dark'],
            padx=15,
            pady=15
        )
        section2.pack(fill=tk.X, padx=20, pady=10)
        
        # Qua trường
        tk.Label(
            section2,
            text="Qua Trường:",
            font=('Arial', 10, 'bold'),
            bg='#FAFAFA',
            fg='#424242'
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.qua_truong_var = tk.StringVar()
        qua_truong_entry = tk.Entry(
            section2,
            textvariable=self.qua_truong_var,
            font=('Arial', 10),
            width=50
        )
        qua_truong_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # Ngành học
        tk.Label(
            section2,
            text="Ngành Học:",
            font=('Arial', 10, 'bold'),
            bg='#FAFAFA',
            fg='#424242'
        ).pack(anchor=tk.W, pady=(10, 5))
        
        self.nganh_hoc_var = tk.StringVar()
        nganh_hoc_entry = tk.Entry(
            section2,
            textvariable=self.nganh_hoc_var,
            font=('Arial', 10),
            width=50
        )
        nganh_hoc_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # Cấp học
        tk.Label(
            section2,
            text="Cấp Học:",
            font=('Arial', 10, 'bold'),
            bg='#FAFAFA',
            fg='#424242'
        ).pack(anchor=tk.W, pady=(10, 5))
        
        self.cap_hoc_var = tk.StringVar()
        cap_hoc_entry = tk.Entry(
            section2,
            textvariable=self.cap_hoc_var,
            font=('Arial', 10),
            width=50
        )
        cap_hoc_entry.pack(fill=tk.X, padx=5, pady=5)
        
        # Thời gian đào tạo
        tk.Label(
            section2,
            text="Thời Gian Đào Tạo:",
            font=('Arial', 10, 'bold'),
            bg='#FAFAFA',
            fg='#424242'
        ).pack(anchor=tk.W, pady=(10, 5))
        
        self.thoi_gian_var = tk.StringVar()
        thoi_gian_entry = tk.Entry(
            section2,
            textvariable=self.thoi_gian_var,
            font=('Arial', 10),
            width=50
        )
        thoi_gian_entry.pack(fill=tk.X, padx=5, pady=5)
    
    def load_data(self):
        """Load dữ liệu hiện tại"""
        # Yếu tố nước ngoài
        self.noi_dung_entry.delete('1.0', tk.END)
        self.noi_dung_entry.insert('1.0', self.personnel.thongTinKhac.noiDungYeuToNN or '')
        self.moi_quan_he_var.set(self.personnel.thongTinKhac.moiQuanHeYeuToNN or '')
        self.ten_nuoc_var.set(self.personnel.thongTinKhac.tenNuoc or '')
        
        # Đào tạo
        self.qua_truong_var.set(self.personnel.quaTruong or '')
        self.nganh_hoc_var.set(self.personnel.nganhHoc or '')
        self.cap_hoc_var.set(self.personnel.capHoc or '')
        self.thoi_gian_var.set(self.personnel.thoiGianDaoTao or '')
    
    def save(self):
        """Lưu dữ liệu"""
        try:
            # Cập nhật yếu tố nước ngoài
            self.personnel.thongTinKhac.noiDungYeuToNN = self.noi_dung_entry.get('1.0', tk.END).strip()
            self.personnel.thongTinKhac.moiQuanHeYeuToNN = self.moi_quan_he_var.get().strip()
            self.personnel.thongTinKhac.tenNuoc = self.ten_nuoc_var.get().strip()
            
            # Cập nhật đào tạo
            self.personnel.quaTruong = self.qua_truong_var.get().strip()
            self.personnel.nganhHoc = self.nganh_hoc_var.get().strip()
            self.personnel.capHoc = self.cap_hoc_var.get().strip()
            self.personnel.thoiGianDaoTao = self.thoi_gian_var.get().strip()
            
            # Lưu vào database
            if self.db.update(self.personnel):
                messagebox.showinfo("Thành công", "Đã lưu thông tin yếu tố nước ngoài và đào tạo")
                self.result = True
                self.close_dialog()
            else:
                messagebox.showerror("Lỗi", "Không thể lưu dữ liệu")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Lỗi khi lưu: {str(e)}")
            import traceback
            traceback.print_exc()
    
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
    
    def cancel(self):
        """Hủy"""
        self.close_dialog()
