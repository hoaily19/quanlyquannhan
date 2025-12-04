"""
Frame quản lý các danh sách báo cáo đại đội
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from services.export import ExportService
from gui.theme import MILITARY_COLORS, get_button_style, get_label_style


class ReportsListFrame(tk.Frame):
    """Frame quản lý các danh sách báo cáo"""
    
    def __init__(self, parent, db: DatabaseService):
        """
        Args:
            parent: Parent widget
            db: DatabaseService instance
        """
        super().__init__(parent)
        self.db = db
        self.export_service = ExportService()
        self.bg_color = '#FAFAFA'
        self._setup_treeview_style()
        self.setup_ui()
    
    def _setup_treeview_style(self):
        """Thiết lập style cho treeview - viền và căn thẳng"""
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Treeview", 
                       borderwidth=1,
                       relief=tk.SOLID,
                       rowheight=25,
                       font=('Segoe UI', 9))
        style.configure("Treeview.Heading", 
                       font=('Segoe UI', 9, 'bold'),
                       background='#388E3C',
                       foreground='white',
                       borderwidth=1,
                       relief=tk.SOLID)
        style.map("Treeview.Heading",
                 background=[('active', '#2E7D32')])
        style.map("Treeview",
                 background=[('selected', '#C8E6C9')],
                 foreground=[('selected', 'black')])
    
    def setup_ui(self):
        """Thiết lập giao diện"""
        self.configure(bg=self.bg_color)
        
        # Title
        title_frame = tk.Frame(self, bg='#388E3C', height=80)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)
        
        # Title chính
        tk.Label(
            title_frame,
            text="📋 DANH SÁCH BÁO CÁO ĐẠI ĐỘI",
            font=('Segoe UI', 18, 'bold'),
            bg='#388E3C',
            fg='white'
        ).pack(expand=True, pady=(10, 0))
        
        # Ghi chú: có thể chỉnh sửa
        tk.Label(
            title_frame,
            text="💡 Tổng hợp từ dữ liệu quân nhân. Có thể chỉnh sửa bằng cách chọn hàng và click 'Chỉnh Sửa'",
            font=('Segoe UI', 9, 'italic'),
            bg='#388E3C',
            fg='#E8F5E9'
        ).pack(expand=True, pady=(0, 10))
        
        # Nút quay lại danh sách quân nhân
        back_btn = tk.Button(
            title_frame,
            text="← Danh Sách Quân Nhân",
            command=self.go_to_personnel_list,
            font=('Segoe UI', 10, 'bold'),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        )
        back_btn.pack(side=tk.RIGHT, padx=15, pady=5)
        
        # Notebook với các tab
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Style cho notebook
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TNotebook.Tab', padding=[20, 10], font=('Segoe UI', 10, 'bold'))
        
        # Tab 1: Trích ngang Đại đội
        tab1 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab1, text="📄 Trích Ngang Đại Đội")
        self.create_trich_ngang_tab(tab1)
        
        # Tab 2: Vị trí Cán bộ
        tab2 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab2, text="👔 Vị Trí Cán Bộ")
        self.create_vi_tri_can_bo_tab(tab2)
        
        # Tab 3: Đảng viên diễn tập
        tab3 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab3, text="🎯 Đảng Viên Diễn Tập")
        self.create_dang_vien_dien_tap_tab(tab3)
        
        # Tab 4: Tổ 3 người
        tab4 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab4, text="👥 Tổ 3 Người")
        self.create_to_3_nguoi_tab(tab4)
        
        # Tab 5: Tổ công tác dân vận
        tab5 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab5, text="🤝 Tổ Công Tác Dân Vận")
        self.create_to_dan_van_tab(tab5)
        
        # Tab 6: Ban chấp hành Chi đoàn
        tab6 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab6, text="🏛️ Ban Chấp Hành Chi Đoàn")
        self.create_ban_chap_hanh_tab(tab6)
        
        # Tab 7: Tổng hợp số liệu
        tab7 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab7, text="📊 Tổng Hợp Số Liệu")
        self.create_tong_hop_tab(tab7)
        
        # Tab 8: Quân nhân theo tôn giáo
        tab8 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab8, text="🕌 Quân Nhân Theo Tôn Giáo")
        self.create_ton_giao_tab(tab8)
        
        # Tab 9: Người thân đảng phái phản động
        tab9 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab9, text="⚠️ Người Thân Đảng Phái Phản Động")
        self.create_dang_phai_phan_dong_tab(tab9)
        
        # Tab 10: Yếu tố nước ngoài
        tab10 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab10, text="🌍 Yếu Tố Nước Ngoài")
        self.create_yeu_to_nuoc_ngoai_tab(tab10)
        
        # Tab 11: Bảo vệ an ninh
        tab11 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab11, text="🛡️ Bảo Vệ An Ninh")
        self.create_bao_ve_an_ninh_tab(tab11)
    
    def create_common_list_view(self, parent, columns, get_data_func, title="", get_id_func=None):
        """
        Tạo view danh sách chung với khả năng chỉnh sửa
        
        Args:
            parent: Parent widget
            columns: Danh sách cột
            get_data_func: Hàm trả về list of tuples (values) hoặc list of dict với 'values' và 'id'
            title: Tiêu đề
            get_id_func: Hàm để lấy ID từ row data (optional, nếu get_data_func không trả về dict)
        """
        # Toolbar
        toolbar = tk.Frame(parent, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        if title:
            title_label = tk.Label(
                toolbar,
                text=title,
                font=('Segoe UI', 12, 'bold'),
                bg=self.bg_color,
                fg='#388E3C'
            )
            title_label.pack(side=tk.LEFT, padx=10)
        
        # Treeview
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns - căn thẳng với header
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            else:
                tree.column(col, width=150, anchor=tk.W)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Buttons toolbar - Thêm nút chỉnh sửa và thêm mới
        btn_container = tk.Frame(toolbar, bg=self.bg_color)
        btn_container.pack(side=tk.RIGHT, padx=5)
        
        # Nút Thêm Mới
        tk.Button(
            btn_container,
            text="➕ Thêm Mới",
            command=self.add_new_personnel,
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Chỉnh Sửa
        edit_btn = tk.Button(
            btn_container,
            text="✏️ Chỉnh Sửa",
            command=lambda: self.edit_selected_from_tree(tree, get_id_func),
            font=('Segoe UI', 10),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        )
        edit_btn.pack(side=tk.LEFT, padx=3)
        
        # Nút Làm Mới
        tk.Button(
            btn_container,
            text="🔄 Làm Mới",
            command=lambda: self.refresh_list(get_data_func, tree, get_id_func),
            font=('Segoe UI', 10),
            bg='#388E3C',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xuất Excel
        tk.Button(
            btn_container,
            text="📥 Xuất Excel",
            command=lambda: self.export_excel(get_data_func, title),
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Bind events - Cho phép double-click để chỉnh sửa
        tree.bind('<Double-1>', lambda e: self.on_double_click_edit(tree, get_id_func))
        tree.bind('<Button-1>', lambda e: self.on_single_click_select(tree))
        
        # Load data
        self.refresh_list(get_data_func, tree, get_id_func)
        
        return tree
    
    def create_common_list_view_with_handlers(self, parent, columns, get_data_func, title="", 
                                               get_id_func=None, custom_edit_handler=None, 
                                               custom_double_click_handler=None):
        """
        Tạo view danh sách với custom handlers cho edit và double-click
        
        Args:
            parent: Parent widget
            columns: Danh sách cột
            get_data_func: Hàm trả về list of dict với 'values' và 'id'
            title: Tiêu đề
            get_id_func: Hàm để lấy ID từ row data
            custom_edit_handler: Handler tùy chỉnh cho nút edit
            custom_double_click_handler: Handler tùy chỉnh cho double-click
        """
        # Toolbar
        toolbar = tk.Frame(parent, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        if title:
            title_label = tk.Label(
                toolbar,
                text=title,
                font=('Segoe UI', 12, 'bold'),
                bg=self.bg_color,
                fg='#388E3C'
            )
            title_label.pack(side=tk.LEFT, padx=10)
        
        # Treeview
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns - căn thẳng với header
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            else:
                tree.column(col, width=150, anchor=tk.W)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Buttons toolbar
        btn_container = tk.Frame(toolbar, bg=self.bg_color)
        btn_container.pack(side=tk.RIGHT, padx=5)
        
        # Nút Thêm Mới
        tk.Button(
            btn_container,
            text="➕ Thêm Mới",
            command=self.add_new_personnel,
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Chỉnh Sửa - sử dụng custom handler nếu có
        edit_command = (lambda: custom_edit_handler(tree, get_id_func)) if custom_edit_handler else (lambda: self.edit_selected_from_tree(tree, get_id_func))
        edit_btn = tk.Button(
            btn_container,
            text="✏️ Chỉnh Sửa",
            command=edit_command,
            font=('Segoe UI', 10),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        )
        edit_btn.pack(side=tk.LEFT, padx=3)
        
        # Nút Làm Mới
        tk.Button(
            btn_container,
            text="🔄 Làm Mới",
            command=lambda: self.refresh_list(get_data_func, tree, get_id_func),
            font=('Segoe UI', 10),
            bg='#388E3C',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xuất Excel
        tk.Button(
            btn_container,
            text="📥 Xuất Excel",
            command=lambda: self.export_excel(get_data_func, title),
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Bind events - sử dụng custom handler nếu có
        if custom_double_click_handler:
            tree.bind('<Double-1>', lambda e: custom_double_click_handler(tree, get_id_func))
        else:
            tree.bind('<Double-1>', lambda e: self.on_double_click_edit(tree, get_id_func))
        tree.bind('<Button-1>', lambda e: self.on_single_click_select(tree))
        
        # Load data
        self.refresh_list(get_data_func, tree, get_id_func)
        
        return tree
    
    def refresh_list(self, get_data_func, tree, get_id_func=None):
        """Làm mới danh sách"""
        # Xóa dữ liệu cũ
        for item in tree.get_children():
            tree.delete(item)
        
        # Load dữ liệu mới - Sửa lỗi Item already exists
        try:
            data = get_data_func()
            inserted_ids = set()  # Track các ID đã insert
            
            for row_data in data:
                try:
                    # Nếu row_data là dict có 'values' và 'id'
                    if isinstance(row_data, dict) and 'values' in row_data:
                        values = row_data['values']
                        personnel_id = row_data.get('id')
                        if personnel_id:
                            # Kiểm tra xem ID đã được insert chưa
                            if personnel_id in inserted_ids:
                                continue
                            # Kiểm tra xem item đã tồn tại trong tree chưa
                            if personnel_id in tree.get_children():
                                tree.delete(personnel_id)
                            tree.insert('', tk.END, iid=personnel_id, values=values)
                            inserted_ids.add(personnel_id)
                        else:
                            tree.insert('', tk.END, values=values)
                    # Nếu row_data là tuple và có get_id_func
                    elif isinstance(row_data, tuple) and get_id_func:
                        personnel_id = get_id_func(row_data)
                        if personnel_id:
                            # Kiểm tra xem ID đã được insert chưa
                            if personnel_id in inserted_ids:
                                continue
                            # Kiểm tra xem item đã tồn tại trong tree chưa
                            if personnel_id in tree.get_children():
                                tree.delete(personnel_id)
                            tree.insert('', tk.END, iid=personnel_id, values=row_data)
                            inserted_ids.add(personnel_id)
                        else:
                            tree.insert('', tk.END, values=row_data)
                    # Nếu row_data là tuple thông thường
                    else:
                        tree.insert('', tk.END, values=row_data)
                except tk.TclError as e:
                    # Bỏ qua lỗi "Item already exists"
                    if "already exists" in str(e):
                        continue
                    else:
                        # Log lỗi khác nhưng không dừng
                        print(f"Lỗi khi insert vào tree: {str(e)}")
                        continue
        except Exception as e:
            # Xử lý lỗi tổng quát - không để giao diện bị nát
            import traceback
            traceback.print_exc()
            # Hiển thị lỗi nhưng không crash
            try:
                messagebox.showerror("Lỗi", f"Không thể tải dữ liệu:\n{str(e)}")
            except:
                # Nếu không thể hiển thị messagebox, chỉ in ra console
                print(f"Lỗi khi tải dữ liệu: {str(e)}")
    
    def on_single_click_select(self, tree):
        """Xử lý single click để highlight row"""
        item = tree.selection()
        if item:
            tree.selection_set(item)
    
    def on_double_click_edit(self, tree, get_id_func):
        """Xử lý double click để chỉnh sửa"""
        selection = tree.selection()
        if selection:
            item_id = selection[0]
            self.edit_personnel_by_id(item_id)
    
    def edit_selected_from_tree(self, tree, get_id_func):
        """Chỉnh sửa quân nhân đã chọn từ tree"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn một quân nhân để chỉnh sửa")
            return
        
        item_id = selection[0]
        self.edit_personnel_by_id(item_id)
    
    def edit_personnel_by_id(self, personnel_id: str):
        """Mở form chỉnh sửa quân nhân theo ID"""
        # Kiểm tra xem ID có phải là ID thực sự hay chỉ là index
        if not personnel_id or personnel_id.isdigit():
            # Có thể là STT, cần tìm lại ID thực
            messagebox.showwarning("Cảnh báo", "Không thể xác định quân nhân. Vui lòng quay lại 'Danh Sách Quân Nhân' để chỉnh sửa.")
            return
        
        # Lưu personnel_id vào main window
        if hasattr(self.master, 'master') and hasattr(self.master.master, 'edit_personnel_id'):
            self.master.master.edit_personnel_id = personnel_id
            self.master.master.show_frame('edit')
        else:
            # Fallback: mở form trực tiếp
            from gui.personnel_form_frame import PersonnelFormFrame
            from tkinter import messagebox
            try:
                person = self.db.get_by_id(personnel_id)
                if person:
                    # Tạo window mới để chỉnh sửa
                    edit_window = tk.Toplevel(self)
                    edit_window.title(f"Chỉnh Sửa: {person.hoTen}")
                    edit_window.geometry("1000x700")
                    form_frame = PersonnelFormFrame(edit_window, self.db, personnel_id=personnel_id)
                    form_frame.pack(fill=tk.BOTH, expand=True)
                else:
                    messagebox.showerror("Lỗi", "Không tìm thấy quân nhân với ID này")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở form chỉnh sửa:\n{str(e)}")
    
    def add_new_personnel(self):
        """Thêm quân nhân mới"""
        if hasattr(self.master, 'master') and hasattr(self.master.master, 'show_frame'):
            self.master.master.show_frame('add')
        else:
            from gui.personnel_form_frame import PersonnelFormFrame
            from tkinter import messagebox
            try:
                # Tạo window mới để thêm
                add_window = tk.Toplevel(self)
                add_window.title("Thêm Quân Nhân Mới")
                add_window.geometry("1000x700")
                form_frame = PersonnelFormFrame(add_window, self.db, is_new=True)
                form_frame.pack(fill=tk.BOTH, expand=True)
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể mở form thêm mới:\n{str(e)}")
    
    def go_to_personnel_list(self):
        """Quay lại danh sách quân nhân để chỉnh sửa"""
        if hasattr(self.master, 'master') and hasattr(self.master.master, 'show_frame'):
            self.master.master.show_frame('list')
        else:
            messagebox.showinfo(
                "Thông báo", 
                "Vui lòng sử dụng menu 'Quản Lý' → 'Danh Sách Quân Nhân' để chỉnh sửa dữ liệu"
            )
    
    def export_excel(self, get_data_func, title):
        """Xuất ra Excel"""
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=f"{title}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            )
            if filename:
                data = get_data_func()
                # TODO: Implement Excel export
                messagebox.showinfo("Thành công", f"Đã xuất file: {filename}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file:\n{str(e)}")
    
    # ========== Các hàm tạo tab ==========
    
    def create_trich_ngang_tab(self, parent):
        """Tab Trích ngang Đại đội - Đầy đủ cột như ảnh"""
        # Toolbar với nút quản lý đơn vị
        toolbar_top = tk.Frame(parent, bg=self.bg_color, pady=10)
        toolbar_top.pack(fill=tk.X, padx=10)
        
        tk.Label(
            toolbar_top,
            text="DANH SÁCH TRÍCH NGANG ĐẠI ĐỘI",
            font=('Segoe UI', 14, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Button(
            toolbar_top,
            text="⚙️ Quản Lý Đơn Vị",
            command=self.manage_units,
            font=('Segoe UI', 10, 'bold'),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.RIGHT, padx=5)
        
        # Cột theo ảnh + thêm cột đơn vị
        columns = (
            'STT', 
            'Họ và tên khai sinh / Họ và tên thường dùng',
            'Ngày tháng năm sinh / Cấp bậc / Ngày nhận',
            'Chức vụ / Ngày nhận',
            'Nhập ngũ / Xuất ngũ',
            'N. vào đoàn / N. vào đảng Chính thức',
            'Thành phần GĐ / Dân tộc / Tôn giáo',
            'Văn hóa',
            'Qua trường / Ngành học / Cấp học / Thời gian',
            'Quê quán / Trú quán / Khi cần báo tin cho ai SĐT',
            'Họ tên cha / Họ tên mẹ / Họ tên vợ',
            'Đơn vị đang làm nhiệm vụ',
            'Ghi chú'
        )
        
        def get_data():
            all_personnel = self.db.get_all()
            
            # Cache đơn vị để tránh load nhiều lần
            units_cache = {}
            try:
                all_units = self.db.get_all_units()
                for unit in all_units:
                    units_cache[unit.id] = unit.ten
            except:
                pass
            
            result = []
            for idx, p in enumerate(all_personnel, 1):
                # Tính tuổi
                try:
                    if p.ngaySinh:
                        birth_year = int(p.ngaySinh.split('/')[-1])
                        age = 2025 - birth_year
                    else:
                        age = ''
                except:
                    age = ''
                
                # Họ và tên
                ho_ten = f"{p.hoTen or ''}"
                if p.hoTenThuongDung:
                    ho_ten += f" / {p.hoTenThuongDung}"
                
                # Ngày sinh / Cấp bậc / Ngày nhận
                ngay_sinh_cb = f"{p.ngaySinh or ''}"
                if p.capBac:
                    ngay_sinh_cb += f" / {p.capBac}"
                if p.ngayNhanCapBac:
                    ngay_sinh_cb += f" / {p.ngayNhanCapBac}"
                
                # Chức vụ / Ngày nhận
                chuc_vu = f"{p.chucVu or ''}"
                if p.ngayNhanChucVu:
                    chuc_vu += f" / {p.ngayNhanChucVu}"
                
                # Nhập ngũ / Xuất ngũ
                nhap_xuat_ngu = f"{p.nhapNgu or ''}"
                if p.xuatNgu:
                    nhap_xuat_ngu += f" / {p.xuatNgu}"
                
                # Vào đoàn / Vào đảng
                doan_dang = f"{p.thongTinKhac.doan.ngayVao or ''}"
                if p.thongTinKhac.dang.ngayChinhThuc:
                    doan_dang += f" / {p.thongTinKhac.dang.ngayChinhThuc}"
                
                # Thành phần GĐ / Dân tộc / Tôn giáo
                thanh_phan = f"{p.thanhPhanGiaDinh or ''}"
                if p.danToc:
                    thanh_phan += f" / {p.danToc}"
                if p.tonGiao:
                    thanh_phan += f" / {p.tonGiao}"
                
                # Qua trường
                qua_truong = f"{p.quaTruong or ''}"
                if p.nganhHoc:
                    qua_truong += f" / {p.nganhHoc}"
                if p.capHoc:
                    qua_truong += f" / {p.capHoc}"
                if p.thoiGianDaoTao:
                    qua_truong += f" / {p.thoiGianDaoTao}"
                
                # Quê quán / Trú quán / Liên hệ
                que_tru_lien_he = f"{p.queQuan or ''}"
                if p.truQuan:
                    que_tru_lien_he += f" / {p.truQuan}"
                if p.lienHeKhiCan or p.soDienThoaiLienHe:
                    lien_he = f"{p.lienHeKhiCan or ''}"
                    if p.soDienThoaiLienHe:
                        lien_he += f": {p.soDienThoaiLienHe}"
                    que_tru_lien_he += f" / {lien_he}"
                
                # Lấy tên đơn vị từ cache
                ten_don_vi = units_cache.get(p.unitId, '') if p.unitId else ''
                
                # Họ tên cha / mẹ / vợ
                ho_ten_gia_dinh = f"{p.hoTenCha or ''}"
                if p.hoTenMe:
                    ho_ten_gia_dinh += f" / {p.hoTenMe}"
                if p.hoTenVo:
                    ho_ten_gia_dinh += f" / {p.hoTenVo}"
                
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        ho_ten,
                        ngay_sinh_cb,
                        chuc_vu,
                        nhap_xuat_ngu,
                        doan_dang,
                        thanh_phan,
                        p.trinhDoVanHoa or '',
                        qua_truong,
                        que_tru_lien_he,
                        ho_ten_gia_dinh,
                        ten_don_vi,  # Đơn vị đang làm nhiệm vụ
                        p.ghiChu or ''  # Ghi chú gốc
                    )
                })
            return result
        
        # Tạo treeview với scrollbar ngang
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns với width phù hợp
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            elif col == 'Đơn vị đang làm nhiệm vụ':
                tree.column(col, width=200, anchor=tk.W, minwidth=150)  # Đảm bảo cột này luôn hiển thị
            elif col == 'Họ và tên khai sinh / Họ và tên thường dùng':
                tree.column(col, width=250, anchor=tk.W)
            elif col == 'Quê quán / Trú quán / Khi cần báo tin cho ai SĐT':
                tree.column(col, width=250, anchor=tk.W)
            elif col == 'Qua trường / Ngành học / Cấp học / Thời gian':
                tree.column(col, width=250, anchor=tk.W)
            else:
                tree.column(col, width=180, anchor=tk.W)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Buttons
        btn_container = tk.Frame(parent, bg=self.bg_color)
        btn_container.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Button(
            btn_container,
            text="➕ Thêm Mới",
            command=self.add_new_personnel,
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        tk.Button(
            btn_container,
            text="✏️ Chỉnh Sửa",
            command=lambda: self.edit_selected_from_tree(tree, None),
            font=('Segoe UI', 10),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        tk.Button(
            btn_container,
            text="🔄 Làm Mới",
            command=lambda: self.refresh_list(get_data, tree, None),
            font=('Segoe UI', 10),
            bg='#388E3C',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Bind events
        tree.bind('<Double-1>', lambda e: self.on_double_click_edit(tree, None))
        tree.bind('<Button-1>', lambda e: self.on_single_click_select(tree))
        
        # Load data
        self.refresh_list(get_data, tree, None)
    
    def manage_units(self):
        """Mở cửa sổ quản lý đơn vị"""
        from gui.unit_management_frame import UnitManagementFrame
        
        unit_window = tk.Toplevel(self)
        unit_window.title("Quản Lý Đơn Vị")
        unit_window.geometry("900x700")
        unit_frame = UnitManagementFrame(unit_window, self.db)
        unit_frame.pack(fill=tk.BOTH, expand=True)
    
    def create_vi_tri_can_bo_tab(self, parent):
        """Tab Vị trí Cán bộ năm 2025"""
        columns = ('STT', 'Họ và Tên', 'Sinh (Tuổi)', 'Quê Quán', 'Cấp Bậc',
                  'Chức Vụ', 'Đơn Vị', 'Vào Đảng', 'Qua Trường')
        
        def get_data():
            all_personnel = self.db.get_all()
            result = []
            for idx, p in enumerate(all_personnel, 1):
                # Tính tuổi
                try:
                    if p.ngaySinh:
                        birth_year = int(p.ngaySinh.split('/')[-1])
                        age = 2025 - birth_year
                    else:
                        age = ''
                except:
                    age = ''
                
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        f"{p.ngaySinh or ''} ({age})" if age else p.ngaySinh or '',
                        p.queQuan or '',
                        p.capBac or '',
                        p.chucVu or '',
                        p.donVi or '',
                        p.thongTinKhac.dang.ngayChinhThuc or '',
                        p.quaTruong or ''  # Qua trường
                    )
                })
            return result
        
        self.create_common_list_view(parent, columns, get_data, "DANH SÁCH VỊ TRÍ CÁN BỘ NĂM 2025")
    
    def create_dang_vien_dien_tap_tab(self, parent):
        """Tab Đảng viên tham gia diễn tập năm 2025"""
        columns = ('STT', 'Họ và Tên', 'Ngày Sinh', 'Cấp Bậc/Chức Vụ', 
                  'Đơn Vị', 'Văn Hóa', 'Dân Tộc', 'Tôn Giáo', 'Chức Vụ Đảng')
        
        def get_data():
            all_personnel = self.db.get_all()
            # Lọc chỉ đảng viên
            dang_vien = [p for p in all_personnel 
                        if p.thongTinKhac.dang.ngayVao or p.thongTinKhac.dang.ngayChinhThuc]
            result = []
            for idx, p in enumerate(dang_vien, 1):
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        p.ngaySinh or '',
                        f"{p.capBac or ''}/{p.chucVu or ''}",
                        p.donVi or '',
                        p.trinhDoVanHoa or '',
                        p.danToc or '',
                        p.tonGiao or '',
                        p.thongTinKhac.dang.chucVuDang or ''
                    )
                })
            return result
        
        self.create_common_list_view(parent, columns, get_data, "ĐẢNG VIÊN THAM GIA DIỄN TẬP NĂM 2025")
    
    def create_to_3_nguoi_tab(self, parent):
        """Tab Tổ 3 người"""
        columns = ('STT', 'Họ và Tên', 'Cấp Bậc', 'Chức Vụ', 'Ghi Chú')
        
        def get_data():
            all_personnel = self.db.get_all()
            result = []
            for idx, p in enumerate(all_personnel, 1):
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        p.capBac or '',
                        p.chucVu or '',
                        p.ghiChu or ''  # Ghi chú
                    )
                })
            return result
        
        self.create_common_list_view(parent, columns, get_data, "DANH SÁCH TỔ 3 NGƯỜI")
    
    def create_to_dan_van_tab(self, parent):
        """Tab Tổ công tác dân vận"""
        columns = ('STT', 'Họ và Tên', 'Cấp Bậc/Chức Vụ', 'Đơn Vị', 
                  'Dân Tộc', 'Tôn Giáo', 'Văn Hóa', 'Ghi Chú')
        
        def get_data():
            all_personnel = self.db.get_all()
            result = []
            for idx, p in enumerate(all_personnel, 1):
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        f"{p.capBac or ''}/{p.chucVu or ''}",
                        p.donVi or '',
                        p.danToc or '',
                        p.tonGiao or '',
                        p.trinhDoVanHoa or '',
                        p.ghiChu or ''  # Ghi chú
                    )
                })
            return result
        
        self.create_common_list_view(parent, columns, get_data, "DANH SÁCH TỔ CÔNG TÁC DÂN VẬN")
    
    def create_ban_chap_hanh_tab(self, parent):
        """Tab Ban chấp hành Chi đoàn"""
        columns = ('STT', 'Họ và Tên', 'Ngày Sinh', 'Cấp Bậc', 'Chức Vụ',
                  'Nhập Ngũ', 'Đơn Vị', 'Ngày Vào Đoàn', 'Chức Vụ Đoàn')
        
        def get_data():
            all_personnel = self.db.get_all()
            # Lọc chỉ đoàn viên
            doan_vien = [p for p in all_personnel if p.thongTinKhac.doan.ngayVao]
            result = []
            for idx, p in enumerate(doan_vien, 1):
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        p.ngaySinh or '',
                        p.capBac or '',
                        p.chucVu or '',
                        p.nhapNgu or '',
                        p.donVi or '',
                        p.thongTinKhac.doan.ngayVao or '',
                        p.thongTinKhac.doan.chucVuDoan or ''
                    )
                })
            return result
        
        self.create_common_list_view(parent, columns, get_data, "BAN CHẤP HÀNH CHI ĐOÀN ĐẠI ĐỘI 3")
    
    def create_tong_hop_tab(self, parent):
        """Tab Tổng hợp số liệu"""
        # Frame tổng hợp
        summary_frame = tk.Frame(parent, bg=self.bg_color)
        summary_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        tk.Label(
            summary_frame,
            text="TỔNG HỢP SỐ LIỆU THUỘC DIỆN QUẢN LÝ NỘI BỘ",
            font=('Segoe UI', 14, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        ).pack(pady=10)
        
        # Treeview cho tổng hợp
        tree_frame = tk.Frame(summary_frame, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ('STT', 'Nội Dung', 'Tổng Số', 'SQ', 'QNCN', 'HSQ/CS', 'Chi Tiết')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor=tk.W)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Load data
        all_personnel = self.db.get_all()
        
        # Tính toán số liệu
        dtts = [p for p in all_personnel if p.danToc and p.danToc.strip()]
        ton_giao = [p for p in all_personnel if p.tonGiao and p.tonGiao.strip()]
        cd_cu = [p for p in all_personnel if p.thongTinKhac.cdCu]
        yeu_to_nn = [p for p in all_personnel if p.thongTinKhac.yeuToNN]
        
        data = [
            (1, 'Quân nhân là người đồng bào DTTS', len(dtts), 0, 0, len(dtts), ''),
            (2, 'Quân nhân theo tôn giáo', len(ton_giao), 0, 0, len(ton_giao), ''),
            (3, 'Quân nhân có người thân tham gia chế độ cũ', len(cd_cu), 0, 0, len(cd_cu), ''),
            (4, 'Quân nhân có yếu tố nước ngoài', len(yeu_to_nn), 0, 0, len(yeu_to_nn), ''),
        ]
        
        for row in data:
            tree.insert('', tk.END, values=row)
    
    def create_ton_giao_tab(self, parent):
        """Tab Quân nhân theo tôn giáo"""
        columns = ('STT', 'Họ Tên', 'Ngày Sinh', 'Nhập Ngũ', 'Cấp Bậc-Chức Vụ',
                  'Đơn Vị', 'Quê Quán', 'Tôn Giáo')
        
        def get_data():
            all_personnel = self.db.get_all()
            # Lọc chỉ có tôn giáo
            ton_giao = [p for p in all_personnel if p.tonGiao and p.tonGiao.strip()]
            result = []
            for idx, p in enumerate(ton_giao, 1):
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        p.ngaySinh or '',
                        p.nhapNgu or '',
                        f"{p.capBac or ''}-{p.chucVu or ''}",
                        p.donVi or '',
                        p.queQuan or '',
                        p.tonGiao or ''
                    )
                })
            return result
        
        self.create_common_list_view(parent, columns, get_data, "QUÂN NHÂN THEO TÔN GIÁO")
    
    def create_dang_phai_phan_dong_tab(self, parent):
        """Tab Người thân đảng phái phản động"""
        columns = ('STT', 'Họ và Tên QN', 'Ngày Sinh', 'Cấp Bậc-Chức Vụ',
                  'Đơn Vị', 'Họ Tên Người Thân', 'Mối Quan Hệ', 'Nội Dung')
        
        def get_data():
            all_personnel = self.db.get_all()
            result = []
            stt = 1
            
            for p in all_personnel:
                # Lấy danh sách người thân từ bảng nguoi_than
                try:
                    nguoi_than_list = self.db.get_nguoi_than_by_personnel(p.id)
                    
                    if nguoi_than_list:
                        # Gom tất cả mối quan hệ lại, cách nhau bằng '/'
                        ho_ten_nguoi_than = ' / '.join([nt.hoTen or '' for nt in nguoi_than_list if nt.hoTen])
                        moi_quan_he = ' / '.join([nt.moiQuanHe or '' for nt in nguoi_than_list if nt.moiQuanHe])
                        noi_dung = ' / '.join([nt.noiDung or '' for nt in nguoi_than_list if nt.noiDung])
                        
                        result.append({
                            'id': p.id,
                            'values': (
                                stt,
                                p.hoTen or '',
                                p.ngaySinh or '',
                                f"{p.capBac or ''}-{p.chucVu or ''}",
                                p.donVi or '',
                                ho_ten_nguoi_than,
                                moi_quan_he,
                                noi_dung
                            )
                        })
                        stt += 1
                    else:
                        # Nếu không có người thân, vẫn hiển thị quân nhân với thông tin trống
                        result.append({
                            'id': p.id,
                            'values': (
                                stt,
                                p.hoTen or '',
                                p.ngaySinh or '',
                                f"{p.capBac or ''}-{p.chucVu or ''}",
                                p.donVi or '',
                                '',
                                '',
                                ''
                            )
                        })
                        stt += 1
                except Exception:
                    # Fallback nếu chưa có bảng nguoi_than
                    result.append({
                        'id': p.id,
                        'values': (
                            stt,
                            p.hoTen or '',
                            p.ngaySinh or '',
                            f"{p.capBac or ''}-{p.chucVu or ''}",
                            p.donVi or '',
                            '',
                            '',
                            ''
                        )
                    })
                    stt += 1
            
            return result
        
        self.create_common_list_view(parent, columns, get_data, 
                                    "QUÂN NHÂN CÓ NGƯỜI THÂN THAM GIA ĐẢNG PHÁI PHẢN ĐỘNG")
    
    def create_yeu_to_nuoc_ngoai_tab(self, parent):
        """Tab Yếu tố nước ngoài"""
        columns = ('STT', 'Họ và Tên', 'Ngày Sinh', 'Cấp Bậc-Chức Vụ',
                  'Đơn Vị', 'Nội Dung Yếu Tố NN', 'Mối Quan Hệ', 'Tên Nước')
        
        def get_data():
            all_personnel = self.db.get_all()
            # Lọc chỉ có yếu tố nước ngoài
            yeu_to_nn = [p for p in all_personnel if p.thongTinKhac.yeuToNN]
            result = []
            for idx, p in enumerate(yeu_to_nn, 1):
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        p.ngaySinh or '',
                        f"{p.capBac or ''}-{p.chucVu or ''}",
                        p.donVi or '',
                        p.thongTinKhac.noiDungYeuToNN or 'Có yếu tố nước ngoài',
                        p.thongTinKhac.moiQuanHeYeuToNN or '',
                        p.thongTinKhac.tenNuoc or ''
                    )
                })
            return result
        
        def get_id_func(item_id):
            """Lấy personnel ID từ item ID"""
            return item_id
        
        # Tạo view với custom edit handler
        def custom_edit_handler(tree, get_id_func):
            """Handler tùy chỉnh để mở form yếu tố nước ngoài"""
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn một quân nhân để chỉnh sửa")
                return
            
            item_id = selection[0]
            personnel_id = get_id_func(item_id) if get_id_func else item_id
            
            # Lấy personnel từ database
            personnel = self.db.get_by_id(personnel_id)
            if not personnel:
                messagebox.showerror("Lỗi", "Không tìm thấy quân nhân")
                return
            
            # Mở form yếu tố nước ngoài
            from gui.yeu_to_nuoc_ngoai_form import YeuToNuocNgoaiFormDialog
            dialog = YeuToNuocNgoaiFormDialog(parent, self.db, personnel)
            parent.wait_window(dialog.dialog)
            
            # Refresh list sau khi đóng dialog
            if dialog.result:
                self.refresh_list(get_data, tree, get_id_func)
        
        # Tạo view với custom double-click handler
        def custom_double_click_handler(tree, get_id_func):
            """Handler double-click tùy chỉnh"""
            selection = tree.selection()
            if selection:
                item_id = selection[0]
                personnel_id = get_id_func(item_id) if get_id_func else item_id
                personnel = self.db.get_by_id(personnel_id)
                if personnel:
                    from gui.yeu_to_nuoc_ngoai_form import YeuToNuocNgoaiFormDialog
                    dialog = YeuToNuocNgoaiFormDialog(parent, self.db, personnel)
                    parent.wait_window(dialog.dialog)
                    if dialog.result:
                        self.refresh_list(get_data, tree, get_id_func)
        
        # Tạo view với handlers tùy chỉnh
        self.create_common_list_view_with_handlers(
            parent, columns, get_data, "QUÂN NHÂN CÓ YẾU TỐ NƯỚC NGOÀI",
            get_id_func, custom_edit_handler, custom_double_click_handler
        )
    
    def create_bao_ve_an_ninh_tab(self, parent):
        """Tab Bảo vệ an ninh"""
        columns = ('STT', 'Họ và Tên', 'Cấp Bậc', 'Chức Vụ', 'Đơn Vị', 'Vị Trí')
        
        def get_data():
            all_personnel = self.db.get_all()
            result = []
            for idx, p in enumerate(all_personnel, 1):
                # TODO: Cần filter theo chức vụ bảo vệ
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        p.capBac or '',
                        p.chucVu or '',
                        p.donVi or '',
                        ''  # Vị trí - cần thêm field
                    )
                })
            return result
        
        self.create_common_list_view(parent, columns, get_data, 
                                    "BÍ THƯ CẤP UỶ, CHI BỘ PHỤ TRÁCH CÔNG TÁC BVAN VÀ CHIẾN SỸ BẢO VỆ")