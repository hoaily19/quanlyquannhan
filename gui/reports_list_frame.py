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
                       rowheight=80,  # Tăng từ 25 lên 80 để hiển thị nhiều dòng text (3-4 dòng)
                       font=('Segoe UI', 9))
        style.configure("Treeview.Heading", 
                       font=('Segoe UI', 9, 'bold'),
                       background='#388E3C',
                       foreground='white',
                       borderwidth=1,
                       relief=tk.SOLID)
        # Tăng padding cho header để hiển thị 2 hàng tối thiểu
        try:
            style.configure("Treeview.Heading", padding=(5, 15))
        except:
            pass
        style.map("Treeview.Heading",
                 background=[('active', '#2E7D32')])
        style.map("Treeview",
                 background=[('selected', '#C8E6C9')],
                 foreground=[('selected', 'black')])
    
    def _add_treeview_border(self, tree):
        """Thêm border cho các hàng trong treeview"""
        try:
            tree.tag_configure('evenrow', background='#FFFFFF')
            tree.tag_configure('oddrow', background='#F5F5F5')
        except:
            pass
    
    def _add_search_toolbar(self, parent, get_data_func, tree_ref, get_id_func=None):
        """
        Thêm toolbar tìm kiếm vào parent frame
        
        Args:
            parent: Parent widget để thêm search toolbar
            get_data_func: Hàm lấy dữ liệu gốc
            tree_ref: List chứa tree reference [tree] hoặc None nếu chưa tạo tree
            get_id_func: Hàm lấy ID (optional)
        
        Returns:
            get_filtered_data: Hàm trả về dữ liệu đã lọc
        """
        # Toolbar tìm kiếm
        search_toolbar = tk.Frame(parent, bg=self.bg_color)
        search_toolbar.pack(fill=tk.X, padx=10, pady=(5, 5))
        
        tk.Label(
            search_toolbar,
            text="🔍 Tìm kiếm:",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        ).pack(side=tk.LEFT, padx=5)
        
        search_var = tk.StringVar()
        search_entry = tk.Entry(
            search_toolbar,
            textvariable=search_var,
            width=40,
            font=('Segoe UI', 10),
            relief=tk.SOLID,
            bd=1,
            highlightthickness=1,
            highlightcolor='#4CAF50',
            highlightbackground='#CCCCCC'
        )
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Hàm lọc dữ liệu
        original_get_data = get_data_func
        
        def get_filtered_data():
            data = original_get_data()
            search_text = search_var.get().strip().lower()
            if not search_text:
                return data
            filtered = []
            for item in data:
                # Tìm kiếm trong tất cả các cột
                if isinstance(item, dict):
                    searchable_text = ' '.join(str(v) for v in item.get('values', [])).lower()
                else:
                    searchable_text = ' '.join(str(v) for v in item).lower()
                if search_text in searchable_text:
                    filtered.append(item)
            return filtered
        
        def on_search_change(*args):
            tree = tree_ref[0] if tree_ref and len(tree_ref) > 0 else None
            if tree:
                self.refresh_list(get_filtered_data, tree, get_id_func)
        
        search_var.trace('w', on_search_change)
        
        return get_filtered_data
    
    def _parse_cap_bac_rank(self, cap_bac: str) -> int:
        """
        Parse cấp bậc thành số để so sánh
        Thứ tự từ cao xuống thấp:
        Đại tá (100) > Trung tá (90) > Thiếu tá (80) > Đại úy (70) > Thượng úy (60) > 
        Trung úy (50) > Thiếu úy (40) > Thượng sĩ (30) > Trung sĩ (20) > Hạ sĩ (10) > 
        H3 (3) > H2 (2) > H1 (1) > 4 (4) > 3 (3) > 2 (2) > 1 (1)
        """
        if not cap_bac:
            return 0
        
        cap_bac = cap_bac.strip().upper()
        
        # Sĩ quan
        if 'ĐẠI TÁ' in cap_bac or 'ĐẠI TÁ' == cap_bac:
            return 100
        elif 'TRUNG TÁ' in cap_bac or 'TRUNG TÁ' == cap_bac:
            return 90
        elif 'THIẾU TÁ' in cap_bac or 'THIẾU TÁ' == cap_bac:
            return 80
        elif 'ĐẠI ÚY' in cap_bac or 'ĐẠI ÚY' == cap_bac:
            return 70
        elif 'THƯỢNG ÚY' in cap_bac or 'THƯỢNG ÚY' == cap_bac:
            return 60
        elif 'TRUNG ÚY' in cap_bac or 'TRUNG ÚY' == cap_bac:
            return 50
        elif 'THIẾU ÚY' in cap_bac or 'THIẾU ÚY' == cap_bac:
            return 40
        # Hạ sĩ quan
        elif 'THƯỢNG SĨ' in cap_bac or 'THƯỢNG SĨ' == cap_bac:
            return 30
        elif 'TRUNG SĨ' in cap_bac or 'TRUNG SĨ' == cap_bac:
            return 20
        elif 'HẠ SĨ' in cap_bac or 'HẠ SĨ' == cap_bac:
            return 10
        # Binh sĩ - H1, H2, H3
        elif cap_bac.startswith('H'):
            try:
                num = int(cap_bac[1:])
                return num  # H1=1, H2=2, H3=3
            except:
                return 0
        # Binh sĩ - số thuần
        else:
            try:
                num = int(cap_bac)
                return num + 10  # 1=11, 2=12, 3=13, 4=14 (cao hơn H1, H2, H3)
            except:
                return 0
    
    def _sort_personnel_by_cap_bac(self, personnel_list):
        """
        Sắp xếp danh sách quân nhân theo cấp bậc (từ cao xuống thấp)
        Nếu cùng cấp bậc, sắp xếp theo tên
        """
        def sort_key(personnel):
            cap_bac_rank = self._parse_cap_bac_rank(personnel.capBac or '')
            ho_ten = (personnel.hoTen or '').lower()
            return (-cap_bac_rank, ho_ten)  # Dấu - để sắp xếp từ cao xuống thấp
        
        return sorted(personnel_list, key=sort_key)
    
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
        notebook.add(tab1, text="Trích Ngang Đại Đội")
        self.create_trich_ngang_tab(tab1)
        
        # Tab 2: Vị trí Cán bộ
        tab2 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab2, text="Vị Trí Cán Bộ")
        self.create_vi_tri_can_bo_tab(tab2)
        
        # Tab 3: Đảng viên diễn tập
        tab3 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab3, text="Đảng Viên Diễn Tập")
        self.create_dang_vien_dien_tap_tab(tab3)
        
        # Tab 4: Tổ 3 người
        tab4 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab4, text="Chế độ cũ")
        self.create_to_3_nguoi_tab(tab4)
        
        # Tab 5: Tổ công tác dân vận
        tab5 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab5, text="Tổ Công Tác Dân Vận")
        self.create_to_dan_van_tab(tab5)
        
        # Tab 6: Ban chấp hành Chi đoàn
        tab6 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab6, text="Ban Chấp Hành Chi Đoàn")
        self.create_ban_chap_hanh_tab(tab6)
        
        # Tab 7: Tổng hợp số liệu
        tab7 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab7, text="Tổng Hợp Số Liệu")
        self.create_tong_hop_tab(tab7)
        
        # Tab 8: Quân nhân theo tôn giáo
        tab8 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab8, text="Quân Nhân Theo Tôn Giáo")
        self.create_ton_giao_tab(tab8)
        
        # Tab 9: Người thân đảng phái phản động
        tab9 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab9, text="Người Thân Đảng Phái Phản Động")
        self.create_dang_phai_phan_dong_tab(tab9)
        
        # Tab 10: Yếu tố nước ngoài
        tab10 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab10, text="Yếu Tố Nước Ngoài")
        self.create_yeu_to_nuoc_ngoai_tab(tab10)
        
        # Tab 11: Bảo vệ an ninh
        tab11 = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab11, text="Bảo Vệ An Ninh")
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
                tree.column(col, width=200, anchor=tk.W)  # Tăng từ 150 lên 200 để hiển thị đầy đủ hơn
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
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
            command=lambda: self.refresh_list(get_filtered_data, tree, get_id_func),
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
            command=lambda: self.export_excel(get_filtered_data, title),
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
        self.refresh_list(get_filtered_data, tree, get_id_func)
        
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
                tree.column(col, width=200, anchor=tk.W)  # Tăng từ 150 lên 200 để hiển thị đầy đủ hơn
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
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
        
        # Load dữ liệu mới - Sửa lỗi Item already exists và thêm border
        try:
            data = get_data_func()
            inserted_ids = set()  # Track các ID đã insert
            
            # Đảm bảo tree có tag cho border
            try:
                tree.tag_configure('evenrow', background='#FFFFFF')
                tree.tag_configure('oddrow', background='#F5F5F5')
            except:
                pass
            
            for idx, row_data in enumerate(data):
                try:
                    # Tag cho hàng chẵn/lẻ để tạo border
                    tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
                    
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
                            tree.insert('', tk.END, iid=personnel_id, values=values, tags=(tag,))
                            inserted_ids.add(personnel_id)
                        else:
                            tree.insert('', tk.END, values=values, tags=(tag,))
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
                            tree.insert('', tk.END, iid=personnel_id, values=row_data, tags=(tag,))
                            inserted_ids.add(personnel_id)
                        else:
                            tree.insert('', tk.END, values=row_data, tags=(tag,))
                    # Nếu row_data là tuple thông thường
                    else:
                        tree.insert('', tk.END, values=row_data, tags=(tag,))
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
    
    def export_vi_tri_can_bo_word(self, get_data_func):
        """Xuất danh sách Vị Trí Cán Bộ ra Word"""
        try:
            from services.export_vi_tri_can_bo import to_word_docx_vi_tri_can_bo
            
            # Lấy dữ liệu từ get_data_func và chuyển thành personnel list
            data = get_data_func()
            if not data:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            personnel_list = []
            for item in data:
                personnel_id = item.get('id')
                if personnel_id:
                    personnel = self.db.get_by_id(personnel_id)
                    if personnel:
                        personnel_list.append(personnel)
            
            # Mở dialog nhập thông tin
            dialog = tk.Toplevel(self)
            dialog.title("Xuất File Word - Vị Trí Cán Bộ")
            dialog.geometry("500x350")
            dialog.transient(self)
            dialog.grab_set()
            
            main_container = tk.Frame(dialog, bg='#FAFAFA')
            main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            # Đại đội
            tk.Label(main_container, text="Đại đội:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dai_doi_var = tk.StringVar(value="Đại đội 3")
            tk.Entry(main_container, textvariable=dai_doi_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Năm
            tk.Label(main_container, text="Năm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            nam_var = tk.StringVar(value="2025")
            tk.Entry(main_container, textvariable=nam_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Chính trị viên
            tk.Label(main_container, text="Chính trị viên:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            chinh_tri_vien_var = tk.StringVar(value="Đại úy Triệu Văn Dũng")
            tk.Entry(main_container, textvariable=chinh_tri_vien_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            def save_and_export():
                try:
                    filename = filedialog.asksaveasfilename(
                        defaultextension=".docx",
                        filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
                        initialfile=f"Danh_sach_vi_tri_can_bo_{datetime.now().strftime('%Y%m%d')}.docx"
                    )
                    if filename:
                        # Xuất Word
                        word_bytes = to_word_docx_vi_tri_can_bo(
                            personnel_list=personnel_list,
                            don_vi=dai_doi_var.get(),
                            nam=nam_var.get(),
                            chinh_tri_vien=chinh_tri_vien_var.get(),
                            db_service=self.db
                        )
                        
                        # Lưu file
                        with open(filename, 'wb') as f:
                            f.write(word_bytes)
                        
                        messagebox.showinfo("Thành công", f"Đã xuất file Word:\n{filename}")
                        dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
            
            btn_frame = tk.Frame(main_container, bg='#FAFAFA', height=70)
            btn_frame.pack(fill=tk.X, padx=10, pady=10, side=tk.BOTTOM)
            btn_frame.pack_propagate(False)
            
            tk.Button(
                btn_frame,
                text="📄 Xuất File",
                command=save_and_export,
                font=('Segoe UI', 11, 'bold'),
                bg='#4CAF50',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=10)
            
            tk.Button(
                btn_frame,
                text="❌ Hủy",
                command=dialog.destroy,
                font=('Segoe UI', 11),
                bg='#757575',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=5)
            
        except ImportError as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
    
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
        
        # Định nghĩa get_data trước khi sử dụng
        def get_data():
            all_personnel = self.db.get_all()
            
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            all_personnel = self._sort_personnel_by_cap_bac(all_personnel)
            
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
        
        # Nút Xuất Word trong toolbar (sau khi đã định nghĩa get_data)
        tk.Button(
            toolbar_top,
            text="📄 Xuất Word",
            command=lambda: self.export_trich_ngang_word(get_data),
            font=('Segoe UI', 10, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.RIGHT, padx=5)
        
        # Toolbar tìm kiếm
        search_toolbar = tk.Frame(parent, bg=self.bg_color)
        search_toolbar.pack(fill=tk.X, padx=10, pady=(5, 5))
        
        tk.Label(
            search_toolbar,
            text="🔍 Tìm kiếm:",
            font=('Segoe UI', 10, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        ).pack(side=tk.LEFT, padx=5)
        
        search_var = tk.StringVar()
        search_entry = tk.Entry(
            search_toolbar,
            textvariable=search_var,
            width=40,
            font=('Segoe UI', 10),
            relief=tk.SOLID,
            bd=1,
            highlightthickness=1,
            highlightcolor='#4CAF50',
            highlightbackground='#CCCCCC'
        )
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Hàm lọc dữ liệu
        original_get_data = get_data
        
        def get_filtered_data():
            data = original_get_data()
            search_text = search_var.get().strip().lower()
            if not search_text:
                return data
            filtered = []
            for item in data:
                # Tìm kiếm trong tất cả các cột
                searchable_text = ' '.join(str(v) for v in item['values']).lower()
                if search_text in searchable_text:
                    filtered.append(item)
            return filtered
        
        def on_search_change(*args):
            self.refresh_list(get_filtered_data, tree, None)
        
        search_var.trace('w', on_search_change)
        
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
                tree.column(col, width=350, anchor=tk.W)  # Tăng từ 250 lên 350 để hiển thị đầy đủ
            elif col == 'Qua trường / Ngành học / Cấp học / Thời gian':
                tree.column(col, width=250, anchor=tk.W)
            else:
                tree.column(col, width=180, anchor=tk.W)
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
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
            command=lambda: self.refresh_list(get_filtered_data, tree, None),
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
        self.refresh_list(get_filtered_data, tree, None)
    
    def manage_units(self):
        """Mở cửa sổ quản lý đơn vị (chỉ mở một cửa sổ duy nhất)"""  
        from gui.unit_management_frame import UnitManagementFrame
        
        # Kiểm tra xem đã có cửa sổ quản lý đơn vị mở chưa
        if hasattr(self, '_unit_window') and self._unit_window.winfo_exists():
            # Nếu đã có, chỉ cần focus vào cửa sổ đó
            self._unit_window.lift()
            self._unit_window.focus_force()
            return
        
        # Tạo cửa sổ mới
        unit_window = tk.Toplevel(self)
        unit_window.title("Quản Lý Đơn Vị")
        unit_window.geometry("900x700")
        
        # Lưu reference để kiểm tra sau
        self._unit_window = unit_window
        
        # Xử lý khi đóng cửa sổ
        def on_close():
            if hasattr(self, '_unit_window'):
                delattr(self, '_unit_window')
            unit_window.destroy()
        
        unit_window.protocol("WM_DELETE_WINDOW", on_close)
        
        unit_frame = UnitManagementFrame(unit_window, self.db)
        unit_frame.pack(fill=tk.BOTH, expand=True)
    
    def export_trich_ngang_word(self, get_data_func):
        """Xuất danh sách Trích ngang Đại đội ra Word"""
        try:
            from services.export_trich_ngang import to_word_docx_trich_ngang
            from tkinter import filedialog
            from datetime import datetime
            
            # Lấy dữ liệu
            data = get_data_func()
            if not data:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            # Lấy danh sách quân nhân từ IDs - giữ nguyên thứ tự từ data (đã sắp xếp)
            personnel_list = []
            for item in data:
                personnel_id = item.get('id')
                if personnel_id:
                    personnel = self.db.get_by_id(personnel_id)
                    if personnel:
                        personnel_list.append(personnel)
            
            if not personnel_list:
                messagebox.showwarning("Cảnh báo", "Không có quân nhân để xuất!")
                return
            
            # Mở dialog nhập thông tin
            dialog = tk.Toplevel(self)
            dialog.title("Xuất File Word - Trích ngang Đại đội")
            dialog.geometry("500x400")
            dialog.transient(self)
            dialog.grab_set()
            
            main_container = tk.Frame(dialog, bg='#FAFAFA')
            main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            # Tiểu đoàn
            tk.Label(main_container, text="Tiểu đoàn:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
            tk.Entry(main_container, textvariable=tieu_doan_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Đại đội
            tk.Label(main_container, text="Đại đội:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dai_doi_var = tk.StringVar(value="ĐẠI ĐỘI 3")
            tk.Entry(main_container, textvariable=dai_doi_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Địa điểm
            tk.Label(main_container, text="Địa điểm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dia_diem_var = tk.StringVar(value="Đăk Lăk")
            tk.Entry(main_container, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Năm
            tk.Label(main_container, text="Năm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            nam_var = tk.StringVar(value="2025")
            tk.Entry(main_container, textvariable=nam_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            def save_and_export():
                try:
                    filename = filedialog.asksaveasfilename(
                        defaultextension=".docx",
                        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                        title="Lưu file Word",
                        initialfile=f"Trich_ngang_dai_doi_{datetime.now().strftime('%Y%m%d')}.docx"
                    )
                    
                    if filename:
                        word_bytes = to_word_docx_trich_ngang(
                            personnel_list=personnel_list,
                            tieu_doan=tieu_doan_var.get(),
                            dai_doi=dai_doi_var.get(),
                            dia_diem=dia_diem_var.get(),
                            nam=nam_var.get(),
                            db_service=self.db
                        )
                        
                        with open(filename, 'wb') as f:
                            f.write(word_bytes)
                        
                        messagebox.showinfo("Thành công", f"Đã xuất file Word:\n{filename}")
                        dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
            
            btn_frame = tk.Frame(main_container, bg='#FAFAFA', height=70)
            btn_frame.pack(fill=tk.X, padx=10, pady=10, side=tk.BOTTOM)
            btn_frame.pack_propagate(False)
            
            tk.Button(
                btn_frame,
                text="📄 Xuất File",
                command=save_and_export,
                font=('Segoe UI', 11, 'bold'),
                bg='#4CAF50',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=10)
            
            tk.Button(
                btn_frame,
                text="❌ Hủy",
                command=dialog.destroy,
                font=('Segoe UI', 11),
                bg='#757575',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=5)
            
        except ImportError as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
    
    def create_vi_tri_can_bo_tab(self, parent):
        """Tab Vị trí Cán bộ năm 2025 - Giống file Word với 11 cột"""
        columns = (
            'TT',
            'Họ và tên Sinh / Quê quán - trú quán SHSQ',
            'Cấp Bậc',
            'Chức, đơn vị',
            'CM Quân',
            'Vào Đảng: Chính thức',
            'Chức vụ chiến đấu / Chức vụ đã qua',
            'Qua trường / Ngành học / Cấp học / Thời gian',
            'VH SK',
            'DT TG',
            'Thông tin gia đình'
        )
        
        def get_data():
            all_personnel = self.db.get_all()
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            all_personnel = self._sort_personnel_by_cap_bac(all_personnel)
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
                
                # Cột 1: TT
                tt = f"{idx:02d}"
                
                # Cột 2: Họ và tên, Sinh, Quê quán - trú quán, SHSQ
                ho_ten = p.hoTen or ''
                ngay_sinh_str = p.ngaySinh or ''
                
                # Quê quán - trú quán
                que_quan = p.queQuan or ''
                tru_quan = p.truQuan or ''
                if que_quan and tru_quan:
                    que_tru = f"{que_quan} - {tru_quan}"
                elif que_quan:
                    que_tru = que_quan
                elif tru_quan:
                    que_tru = tru_quan
                else:
                    que_tru = ''
                
                # SHSQ (Số hiệu sĩ quan) - có thể lưu trong ghi chú hoặc cần thêm field
                shsq = p.ghiChu or ""
                
                # Format: Họ và tên Sinh (tuổi) Quê quán - trú quán SHSQ
                # Viết ngắn gọn và xuống dòng
                col2_parts = []
                if ho_ten:
                    col2_parts.append(ho_ten)
                if ngay_sinh_str:
                    # Bỏ phần tuổi trong ngoặc ở bảng danh sách
                    col2_parts.append(ngay_sinh_str)
                if que_tru:
                    # Rút ngắn: chỉ lấy tỉnh (phần cuối cùng)
                    que_tru_short = que_tru
                    if ',' in que_tru:
                        parts = que_tru.split(',')
                        if len(parts) >= 1:
                            que_tru_short = parts[-1].strip()  # Chỉ lấy tỉnh
                    # Giới hạn độ dài tối đa
                    if len(que_tru_short) > 30:
                        que_tru_short = que_tru_short[:30] + "..."
                    col2_parts.append(que_tru_short)
                if shsq:
                    col2_parts.append(f"SHSQ: {shsq}")
                
                col2 = "\n".join(col2_parts)  # Dùng \n thay vì space để xuống dòng
                
                # Cột 3: Cấp Bậc (Tháng năm nhận) - Viết tắt
                cap_bac = p.capBac or ''
                ngay_nhan_cb = p.ngayNhanCapBac or ''
                if cap_bac and ngay_nhan_cb:
                    # Format: cấp bậc/ MM/YYYY (rút ngắn)
                    if '/' in ngay_nhan_cb:
                        parts = ngay_nhan_cb.split('/')
                        if len(parts) == 3:
                            # DD/MM/YYYY -> MM/YYYY
                            thang_nam = f"{parts[1]}/{parts[2]}"
                            col3 = f"{cap_bac}\n{thang_nam}"
                        else:
                            col3 = f"{cap_bac}\n{ngay_nhan_cb}"
                    else:
                        col3 = f"{cap_bac}\n{ngay_nhan_cb}"
                elif cap_bac:
                    col3 = cap_bac
                else:
                    col3 = ''
                
                # Cột 4: Chức, đơn vị (Tháng năm nhận) - Format: "CTV C3.d15, 5/2019"
                chuc_vu = p.chucVu or ''
                don_vi = p.donVi or ''
                ngay_nhan_cv = p.ngayNhanChucVu or ''
                if chuc_vu and don_vi:
                    chuc_don_vi = f"{chuc_vu} {don_vi}"
                    if ngay_nhan_cv:
                        # Lấy tháng/năm từ ngày nhận (DD/MM/YYYY -> MM/YYYY)
                        if '/' in ngay_nhan_cv:
                            parts = ngay_nhan_cv.split('/')
                            if len(parts) == 3:
                                # DD/MM/YYYY -> MM/YYYY
                                thang_nam = f"{parts[1]}/{parts[2]}"
                                col4 = f"{chuc_don_vi}\n{thang_nam}"  # Xuống dòng
                            elif len(parts) >= 2:
                                # MM/YYYY
                                thang_nam = f"{parts[0]}/{parts[1]}"
                                col4 = f"{chuc_don_vi}\n{thang_nam}"  # Xuống dòng
                            else:
                                col4 = f"{chuc_don_vi}\n{ngay_nhan_cv}"  # Xuống dòng
                        else:
                            col4 = f"{chuc_don_vi}\n{ngay_nhan_cv}"  # Xuống dòng
                    else:
                        col4 = chuc_don_vi
                elif chuc_vu:
                    col4 = chuc_vu
                elif don_vi:
                    col4 = don_vi
                else:
                    col4 = ''
                
                # Cột 5: CM Quân (Tháng năm) - Format: "9/2012" hoặc "20/12/2024"
                cm_quan = p.cmQuan or p.nhapNgu or ''
                # Chuyển đổi format nếu cần (DD/MM/YYYY -> MM/YYYY hoặc giữ nguyên)
                if cm_quan and '/' in cm_quan:
                    parts = cm_quan.split('/')
                    if len(parts) >= 2:
                        # Có thể là DD/MM/YYYY hoặc MM/YYYY
                        if len(parts) == 3:
                            # DD/MM/YYYY -> MM/YYYY
                            cm_quan = f"{parts[1]}/{parts[2]}"
                        # Nếu là MM/YYYY thì giữ nguyên
                col5 = cm_quan
                
                # Cột 6: Vào Đảng: Chính thức - Viết tắt
                ngay_vao_dang = p.thongTinKhac.dang.ngayVao or ''
                ngay_chinh_thuc = p.thongTinKhac.dang.ngayChinhThuc or ''
                col6_parts = []
                if ngay_vao_dang:
                    # Chuyển DD/MM/YYYY -> MM/YYYY nếu cần
                    if '/' in ngay_vao_dang:
                        parts = ngay_vao_dang.split('/')
                        if len(parts) == 3:
                            ngay_vao_dang = f"{parts[1]}/{parts[2]}"
                    col6_parts.append(f"V: {ngay_vao_dang}")  # Viết tắt: "Vào" -> "V"
                if ngay_chinh_thuc:
                    if '/' in ngay_chinh_thuc:
                        parts = ngay_chinh_thuc.split('/')
                        if len(parts) == 3:
                            ngay_chinh_thuc = f"{parts[1]}/{parts[2]}"
                    col6_parts.append(f"CT: {ngay_chinh_thuc}")  # Viết tắt: "Chính thức" -> "CT"
                col6 = "\n".join(col6_parts) if col6_parts else ''  # Dùng \n để xuống dòng
                
                # Cột 7: Chức vụ chiến đấu (Thời gian) Chức vụ đã qua (Thời gian) - Viết tắt
                col7_parts = []
                if p.chucVuChienDau:
                    chien_dau_text = f"CD: {p.chucVuChienDau}"  # Viết tắt: "Chiến đấu" -> "CD"
                    if p.thoiGianChucVuChienDau:
                        chien_dau_text += f" ({p.thoiGianChucVuChienDau})"
                    col7_parts.append(chien_dau_text)
                elif chuc_vu and don_vi:
                    # Fallback: dùng chức vụ và đơn vị hiện tại
                    col7_parts.append(f"CD: {chuc_vu}/{don_vi}")
                
                if p.chucVuDaQua:
                    da_qua_text = f"ĐQ: {p.chucVuDaQua}"  # Viết tắt: "Đã qua" -> "ĐQ"
                    if p.thoiGianChucVuDaQua:
                        da_qua_text += f" ({p.thoiGianChucVuDaQua})"
                    col7_parts.append(da_qua_text)
                
                col7 = "\n".join(col7_parts) if col7_parts else ''  # Đã dùng \n
                
                # Cột 8: Qua trường (Ngành, thời gian, kết quả)
                qua_truong = p.quaTruong or ''
                nganh_hoc = p.nganhHoc or ''
                thoi_gian = p.thoiGianDaoTao or ''
                ket_qua = p.ketQuaDaoTao or ''
                if qua_truong:
                    qua_truong_info = qua_truong
                    if nganh_hoc or thoi_gian or ket_qua:
                        qua_truong_info += "\n("
                        parts = []
                        if nganh_hoc:
                            parts.append(nganh_hoc)
                        if thoi_gian:
                            parts.append(thoi_gian)
                        if ket_qua:
                            parts.append(f"-{ket_qua}")
                        qua_truong_info += ", ".join(parts)
                        qua_truong_info += ")"
                    col8 = qua_truong_info
                else:
                    col8 = ''
                
                # Cột 9: VH SK (Văn hóa, Sức khỏe)
                trinh_do_vh = p.trinhDoVanHoa or ''
                col9 = trinh_do_vh  # Sức khỏe có thể cần thêm field
                
                # Cột 10: DT TG (Dân tộc, Tôn giáo) - Xuống dòng để rõ ràng
                dan_toc = p.danToc or ''
                ton_giao = p.tonGiao or 'Không'
                col10_parts = []
                if dan_toc:
                    col10_parts.append(f"DT: {dan_toc}")
                if ton_giao and ton_giao != 'Không':
                    col10_parts.append(f"TG: {ton_giao}")
                col10 = "\n".join(col10_parts) if col10_parts else ''
                
                # Cột 11: Thông tin gia đình
                # Lấy từ bảng nguoi_than
                gia_dinh_info = []
                
                try:
                    # Lấy danh sách người thân từ database
                    nguoi_than_list = self.db.get_nguoi_than_by_personnel(p.id)
                    
                    # Nhóm theo mối quan hệ
                    bo_de = []
                    me_de = []
                    bo_vo = []
                    me_vo = []
                    vo = []
                    con = []
                    khac = []
                    
                    for nguoi_than in nguoi_than_list:
                        moi_quan_he = (nguoi_than.moiQuanHe or '').lower()
                        ho_ten = nguoi_than.hoTen or ''
                        ngay_sinh = nguoi_than.ngaySinh or ''
                        dia_chi = nguoi_than.diaChi or ''
                        so_dt = nguoi_than.soDienThoai or ''
                        noi_dung = nguoi_than.noiDung or ''
                        
                        # Format thông tin theo mẫu Word: "- Tên, năm sinh, nghề"
                        # Ví dụ: "- Triệu Văn Tung, 1968, LN"
                        info_parts = []
                        if ho_ten:
                            # Lấy năm sinh từ ngày sinh
                            nam_sinh = ""
                            if ngay_sinh:
                                try:
                                    # Format: DD/MM/YYYY hoặc YYYY
                                    if '/' in ngay_sinh:
                                        nam_sinh = ngay_sinh.split('/')[-1]
                                    else:
                                        nam_sinh = ngay_sinh[:4] if len(ngay_sinh) >= 4 else ngay_sinh
                                except:
                                    nam_sinh = ""
                            
                            # Nghề (có thể lấy từ nội dung hoặc để trống)
                            nghe = noi_dung if noi_dung else "LN"  # Mặc định LN nếu không có
                            
                            if nam_sinh:
                                info_str = f"- {ho_ten}, {nam_sinh}, {nghe}"
                            else:
                                info_str = f"- {ho_ten}, {nghe}"
                        else:
                            info_str = ""
                        
                        # Phân loại theo mối quan hệ - Viết tắt
                        if 'bố' in moi_quan_he or 'cha' in moi_quan_he:
                            if 'vợ' in moi_quan_he or 'vo' in moi_quan_he:
                                bo_vo.append(f"BV: {info_str}")  # Viết tắt: "Bố vợ" -> "BV"
                            else:
                                bo_de.append(f"BĐ: {info_str}")  # Viết tắt: "Bố đẻ" -> "BĐ"
                        elif 'mẹ' in moi_quan_he or 'me' in moi_quan_he:
                            if 'vợ' in moi_quan_he or 'vo' in moi_quan_he:
                                me_vo.append(f"MV: {info_str}")  # Viết tắt: "Mẹ vợ" -> "MV"
                            else:
                                me_de.append(f"MĐ: {info_str}")  # Viết tắt: "Mẹ đẻ" -> "MĐ"
                        elif 'vợ' in moi_quan_he or 'vo' in moi_quan_he:
                            vo.append(f"V: {info_str}")  # Viết tắt: "Vợ" -> "V"
                        elif 'con' in moi_quan_he:
                            con.append(f"C: {info_str}")  # Viết tắt: "Con" -> "C"
                        else:
                            if moi_quan_he:
                                # Viết tắt mối quan hệ
                                mqh_short = moi_quan_he[:2].upper() if len(moi_quan_he) >= 2 else moi_quan_he.upper()
                                khac.append(f"{mqh_short}: {info_str}")
                            else:
                                khac.append(info_str)
                    
                    # Thêm vào danh sách theo thứ tự (giống mẫu Word)
                    gia_dinh_info.extend(bo_de)
                    gia_dinh_info.extend(me_de)
                    
                    # Nơi ở hiện nay (của bố mẹ)
                    dia_chi_bo_me = None
                    for nguoi_than in nguoi_than_list:
                        moi_quan_he = (nguoi_than.moiQuanHe or '').lower()
                        if ('bố' in moi_quan_he or 'cha' in moi_quan_he or 'mẹ' in moi_quan_he or 'me' in moi_quan_he) and not ('vợ' in moi_quan_he or 'vo' in moi_quan_he):
                            if nguoi_than.diaChi:
                                dia_chi_bo_me = nguoi_than.diaChi
                                break
                    
                    if dia_chi_bo_me:
                        gia_dinh_info.append(dia_chi_bo_me)
                    elif p.queQuan:
                        gia_dinh_info.append(p.queQuan)
                    
                    gia_dinh_info.extend(bo_vo)
                    gia_dinh_info.extend(me_vo)
                    
                    # Nơi ở hiện nay (của bố mẹ vợ)
                    dia_chi_bo_me_vo = None
                    for nguoi_than in nguoi_than_list:
                        moi_quan_he = (nguoi_than.moiQuanHe or '').lower()
                        if ('bố' in moi_quan_he or 'cha' in moi_quan_he or 'mẹ' in moi_quan_he or 'me' in moi_quan_he) and ('vợ' in moi_quan_he or 'vo' in moi_quan_he):
                            if nguoi_than.diaChi:
                                dia_chi_bo_me_vo = nguoi_than.diaChi
                                break
                    
                    if dia_chi_bo_me_vo:
                        gia_dinh_info.append(dia_chi_bo_me_vo)
                    
                    gia_dinh_info.extend(vo)
                    gia_dinh_info.extend(con)
                    
                    # Nơi ở hiện nay (của vợ/con)
                    dia_chi_vo_con = None
                    for nguoi_than in nguoi_than_list:
                        moi_quan_he = (nguoi_than.moiQuanHe or '').lower()
                        if 'vợ' in moi_quan_he or 'vo' in moi_quan_he or 'con' in moi_quan_he:
                            if nguoi_than.diaChi:
                                dia_chi_vo_con = nguoi_than.diaChi
                                break
                    
                    if dia_chi_vo_con:
                        gia_dinh_info.append(dia_chi_vo_con)
                    
                    gia_dinh_info.extend(khac)
                    
                    # SĐT gia đình
                    sdt_gia_dinh = None
                    for nguoi_than in nguoi_than_list:
                        if nguoi_than.soDienThoai:
                            sdt_gia_dinh = nguoi_than.soDienThoai
                            break
                    
                    if sdt_gia_dinh:
                        gia_dinh_info.append(f"ĐT: {sdt_gia_dinh}")  # Viết tắt: "SĐT" -> "ĐT"
                    elif p.soDienThoaiLienHe:
                        gia_dinh_info.append(f"ĐT: {p.soDienThoaiLienHe}")  # Viết tắt
                    
                except Exception as e:
                    # Fallback: sử dụng các field cũ nếu có lỗi
                    if p.hoTenCha:
                        gia_dinh_info.append(f"Bố đẻ: {p.hoTenCha}")
                    if p.hoTenMe:
                        gia_dinh_info.append(f"Mẹ đẻ: {p.hoTenMe}")
                    if p.queQuan:
                        gia_dinh_info.append(f"Nơi ở: {p.queQuan}")
                    if p.hoTenVo:
                        gia_dinh_info.append(f"Vợ: {p.hoTenVo}")
                    if p.soDienThoaiLienHe:
                        gia_dinh_info.append(f"SĐT: {p.soDienThoaiLienHe}")
                
                # Format hiển thị: dùng \n để xuống dòng cho dễ đọc
                col11 = "\n".join(gia_dinh_info) if gia_dinh_info else ''
                
                result.append({
                    'id': p.id,
                    'values': (
                        tt,
                        col2,
                        col3,
                        col4,
                        col5,
                        col6,
                        col7,
                        col8,
                        col9,
                        col10,
                        col11
                    )
                })
            return result
        
        # Tạo view với custom column widths
        self.create_vi_tri_can_bo_list_view(parent, columns, get_data, "DANH SÁCH VỊ TRÍ CÁN BỘ NĂM 2025")
    
    def create_vi_tri_can_bo_list_view(self, parent, columns, get_data_func, title="", get_id_func=None):
        """
        Tạo view danh sách Vị Trí Cán Bộ với column widths tùy chỉnh
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
        
        # Thêm search toolbar - tạo tree_ref list để có thể cập nhật sau
        tree_ref = [None]
        get_filtered_data = self._add_search_toolbar(parent, get_data_func, tree_ref, get_id_func)
        
        # Treeview
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        # Cập nhật tree_ref
        tree_ref[0] = tree
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns với width tùy chỉnh - cân đối để hiển thị đầy đủ
        column_widths = {
            'TT': 40,
            'Họ và tên Sinh  Quê quán - trú quán SHSQ': 220,  # Giảm để ngắn gọn hơn
            'Cấp Bậc (Tháng năm nhận)': 140,
            'Chức, đơn vị (Tháng năm nhận)': 180,
            'CM Quân (Tháng năm)': 110,
            'Vào Đảng: Chính thức': 140,
            'Chức vụ chiến đấu / Chức vụ đã qua': 300,
            'Qua trường (Ngành, thời gian, kết quả)': 250,
            'VH SK': 80,
            'DT TG': 100,
            'Thông tin gia đình': 400  # Giảm xuống để không bị đẩy ra xa
        }
        
        for col in columns:
            # Thêm xuống dòng cho header text nếu dài quá
            header_text = col
            # Nếu header dài hơn 20 ký tự, thêm xuống dòng ở vị trí hợp lý
            if len(col) > 20:
                # Tìm vị trí để xuống dòng (ưu tiên sau dấu phẩy, ngoặc đơn, hoặc khoảng trắng)
                split_pos = -1
                # Tìm dấu phẩy đầu tiên sau 15 ký tự
                for i in range(15, len(col)):
                    if col[i] in [',', '(', '-']:
                        split_pos = i + 1
                        break
                # Nếu không tìm thấy, tìm khoảng trắng
                if split_pos == -1:
                    for i in range(15, len(col)):
                        if col[i] == ' ':
                            split_pos = i + 1
                            break
                # Nếu vẫn không tìm thấy, chia đôi
                if split_pos == -1:
                    split_pos = len(col) // 2
                
                if split_pos > 0:
                    header_text = col[:split_pos] + '\n' + col[split_pos:]
            
            tree.heading(col, text=header_text)
            width = column_widths.get(col, 150)
            if col == 'TT':
                tree.column(col, width=width, anchor=tk.CENTER)
            else:
                tree.column(col, width=width, anchor=tk.W)
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
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
            command=lambda: self.refresh_list(get_filtered_data, tree, get_id_func),
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
            text="📊 Xuất Excel",
            command=lambda: self.export_excel(get_filtered_data, title),
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xuất Word
        tk.Button(
            btn_container,
            text="📄 Xuất Word",
            command=lambda: self.export_vi_tri_can_bo_word(get_filtered_data),
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Bind events
        tree.bind('<Double-1>', lambda e: self.on_double_click_edit(tree, get_id_func))
        tree.bind('<Button-1>', lambda e: self.on_single_click_select(tree))
        
        # Load data
        self.refresh_list(get_filtered_data, tree, get_id_func)
    
    def create_dang_vien_dien_tap_tab(self, parent):
        """Tab Đảng viên tham gia diễn tập năm 2025"""
        columns = ('STT', 'Họ và Tên', 'Ngày Sinh', 'Cấp Bậc/Chức Vụ', 
                  'Đơn Vị', 'Văn Hóa', 'Dân Tộc', 'Tôn Giáo', 'Chức Vụ Đảng', 
                  'Quê Quán/Trú Quán', 'Ghi Chú')
        
        def get_data():
            # Chỉ lấy quân nhân đã được chọn vào danh sách
            selected_ids = set(self.db.get_dang_vien_dien_tap())
            all_personnel = self.db.get_all()
            filtered_personnel = [p for p in all_personnel if p.id in selected_ids]
            
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            filtered_personnel = self._sort_personnel_by_cap_bac(filtered_personnel)
            
            result = []
            for idx, p in enumerate(filtered_personnel, 1):
                # Quê quán/Trú quán
                que_quan = p.queQuan or ''
                tru_quan = p.truQuan or ''
                que_tru = f"{que_quan}; {tru_quan}".strip('; ').strip()
                
                # Lấy ghi chú riêng từ tab đảng viên diễn tập
                ghi_chu = self.db.get_dang_vien_dien_tap_ghi_chu(p.id)
                
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        p.ngaySinh or '',
                        f"{p.capBac or ''}/{p.chucVu or ''}".strip('/'),
                        p.donVi or '',
                        p.trinhDoVanHoa or '',
                        p.danToc or '',
                        p.tonGiao or '',
                        p.thongTinKhac.dang.chucVuDang or '',
                        que_tru,  # Quê quán/Trú quán
                        ghi_chu  # Ghi chú riêng
                    )
                })
            return result
        
        # Toolbar
        toolbar = tk.Frame(parent, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        title_label = tk.Label(
            toolbar,
            text="ĐẢNG VIÊN THAM GIA DIỄN TẬP NĂM 2025",
            font=('Segoe UI', 12, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        )
        title_label.pack(side=tk.LEFT, padx=10)
        
        # Thêm search toolbar - tạo tree_ref list để có thể cập nhật sau
        tree_ref = [None]
        get_filtered_data = self._add_search_toolbar(parent, get_data, tree_ref, None)
        
        # Buttons toolbar
        btn_container = tk.Frame(toolbar, bg=self.bg_color)
        btn_container.pack(side=tk.RIGHT, padx=5)
        
        # Nút Chọn Quân Nhân
        tk.Button(
            btn_container,
            text="👥 Chọn Quân Nhân",
            command=lambda: self.choose_dang_vien_dien_tap_personnel(parent),
            font=('Segoe UI', 10),
            bg='#9C27B0',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Treeview - tạo trước để có thể dùng trong các nút
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        # Cập nhật tree_ref
        tree_ref[0] = tree
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            elif col == 'Họ và Tên':
                tree.column(col, width=200, anchor=tk.W)  # Tăng từ 150 lên 200
            elif col == 'Ngày Sinh':
                tree.column(col, width=120, anchor=tk.CENTER)  # Tăng từ 100 lên 120
            elif col == 'Cấp Bậc/Chức Vụ':
                tree.column(col, width=150, anchor=tk.CENTER)  # Tăng từ 120 lên 150
            elif col == 'Đơn Vị':
                tree.column(col, width=120, anchor=tk.CENTER)  # Tăng từ 80 lên 120
            elif col == 'Văn Hóa':
                tree.column(col, width=120, anchor=tk.CENTER)  # Tăng từ 80 lên 120
            elif col == 'Dân Tộc':
                tree.column(col, width=130, anchor=tk.W)  # Tăng từ 100 lên 130
            elif col == 'Tôn Giáo':
                tree.column(col, width=120, anchor=tk.W)  # Tăng từ 100 lên 120
            elif col == 'Chức Vụ Đảng':
                tree.column(col, width=150, anchor=tk.W)  # Tăng từ 120 lên 150
            elif col == 'Quê Quán/Trú Quán':
                tree.column(col, width=250, anchor=tk.W)  # Tăng từ 200 lên 250
            elif col == 'Ghi Chú':
                tree.column(col, width=200, anchor=tk.W)  # Tăng từ 150 lên 200
            else:
                tree.column(col, width=150, anchor=tk.W)
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Nút Chỉnh Sửa
        edit_btn = tk.Button(
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
        )
        edit_btn.pack(side=tk.LEFT, padx=3)
        
        # Nút Xóa
        def delete_selected():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn quân nhân cần xóa!")
                return
            
            item_id = selection[0]
            values = tree.item(item_id, 'values')
            ho_ten = values[1] if len(values) > 1 else "quân nhân này"
            
            if messagebox.askyesno("Xác nhận", f"Bạn có chắc chắn muốn xóa {ho_ten} khỏi danh sách?"):
                if self.db.remove_dang_vien_dien_tap(item_id):
                    messagebox.showinfo("Thành công", f"Đã xóa {ho_ten} khỏi danh sách!")
                    self.refresh_list(get_filtered_data, tree, None)
                else:
                    messagebox.showerror("Lỗi", "Không thể xóa quân nhân!")
        
        tk.Button(
            btn_container,
            text="🗑️ Xóa",
            command=delete_selected,
            font=('Segoe UI', 10),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Làm Mới
        tk.Button(
            btn_container,
            text="🔄 Làm Mới",
            command=lambda: self.refresh_list(get_filtered_data, tree, None),
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
            command=lambda: self.export_excel(get_filtered_data, "ĐẢNG VIÊN THAM GIA DIỄN TẬP NĂM 2025"),
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xuất Word
        def export_word():
            self.export_dang_vien_dien_tap_word(get_filtered_data)
        
        tk.Button(
            btn_container,
            text="📄 Xuất Word",
            command=export_word,
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Biến để lưu trữ entry widget đang edit
        editing_cell = {'item': None, 'column': None, 'entry': None, 'buttons': None}
        
        def start_edit(event):
            """Bắt đầu chỉnh sửa cell"""
            region = tree.identify_region(event.x, event.y)
            if region != "cell":
                return
            
            item = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            
            if not item or not column:
                return
            
            # Chỉ cho phép edit cột Ghi Chú (cột 11, index 10)
            col_index = int(column.replace('#', '')) - 1
            if col_index != 10:  # Ghi Chú là cột cuối cùng
                return
            
            # Hủy edit cũ nếu có
            if editing_cell['entry']:
                cancel_edit()
            
            # Lấy giá trị hiện tại
            values = list(tree.item(item, 'values'))
            current_value = values[col_index] if col_index < len(values) else ''
            
            # Lấy vị trí cell
            bbox = tree.bbox(item, column)
            if not bbox:
                return
            
            # Tạo entry widget
            entry = tk.Entry(tree_frame, font=('Segoe UI', 9))
            entry.insert(0, current_value)
            entry.select_range(0, tk.END)
            entry.focus()
            
            # Đặt vị trí entry
            entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
            
            # Tạo nút Save và Cancel
            btn_frame = tk.Frame(tree_frame, bg='white')
            btn_frame.place(x=bbox[0] + bbox[2] + 5, y=bbox[1])
            
            save_btn = tk.Button(btn_frame, text="✓", command=lambda: save_edit(item, col_index, entry.get()), 
                                bg='#4CAF50', fg='white', width=3, font=('Segoe UI', 8))
            save_btn.pack(side=tk.LEFT, padx=2)
            
            cancel_btn = tk.Button(btn_frame, text="✗", command=cancel_edit,
                                  bg='#F44336', fg='white', width=3, font=('Segoe UI', 8))
            cancel_btn.pack(side=tk.LEFT, padx=2)
            
            editing_cell['item'] = item
            editing_cell['column'] = column
            editing_cell['entry'] = entry
            editing_cell['buttons'] = btn_frame
            
            def on_entry_return(event):
                save_edit(item, col_index, entry.get())
            
            def on_entry_escape(event):
                cancel_edit()
            
            entry.bind('<Return>', on_entry_return)
            entry.bind('<Escape>', on_entry_escape)
        
        def cancel_edit():
            """Hủy chỉnh sửa"""
            if editing_cell['entry']:
                editing_cell['entry'].destroy()
                editing_cell['entry'] = None
            if editing_cell['buttons']:
                editing_cell['buttons'].destroy()
                editing_cell['buttons'] = None
            editing_cell['item'] = None
            editing_cell['column'] = None
        
        def save_edit(item, col_index, new_value):
            """Lưu giá trị đã chỉnh sửa"""
            item_id = item
            
            if col_index == 10:  # Ghi Chú
                # Lưu ghi chú riêng cho tab đảng viên diễn tập
                if self.db.update_dang_vien_dien_tap_ghi_chu(item_id, new_value.strip()):
                    messagebox.showinfo("Thành công", "Đã cập nhật ghi chú!")
                else:
                    messagebox.showerror("Lỗi", "Không thể cập nhật ghi chú!")
                    cancel_edit()
                    return
            
            cancel_edit()
            # Refresh lại danh sách
            self.refresh_list(get_filtered_data, tree, None)
        
        # Bind events
        tree.bind('<Double-1>', start_edit)
        tree.bind('<Button-1>', lambda e: (cancel_edit(), self.on_single_click_select(tree)))
        
        # Load data
        self.refresh_list(get_filtered_data, tree, None)
    
    def create_to_3_nguoi_tab(self, parent):
        """Tab Quân nhân có người thân tham gia chế độ cũ"""
        columns = ('STT', 'Họ và Tên', 'Đơn Vị', 'Ngụy Quân', 'Ngụy Quyền', 
                  'Nợ Máu', 'Quê Quán', 'Chỗ Ở', 'Họ Tên Người Thân', 
                  'Quan Hệ', 'Đã Cải Tạo')
        
        def get_data():
            # Tự động lấy quân nhân có đánh dấu "Có người thân tham gia chế độ cũ"
            all_personnel = self.db.get_all()
            # Lọc quân nhân có checkbox cdCu được đánh dấu
            filtered_personnel = [p for p in all_personnel if p.thongTinKhac.cdCu]
            
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            filtered_personnel = self._sort_personnel_by_cap_bac(filtered_personnel)
            
            result = []
            for idx, p in enumerate(filtered_personnel, 1):
                # Lấy thông tin người thân
                nguoi_than_info = ""
                quan_he = ""
                try:
                    nguoi_than_list = self.db.get_nguoi_than_by_personnel(p.id)
                    if nguoi_than_list:
                        # Lấy người thân đầu tiên
                        nt = nguoi_than_list[0]
                        ho_ten_nt = nt.hoTen or ''
                        ngay_sinh_nt = nt.ngaySinh or ''
                        noi_dung_nt = nt.noiDung or ''
                        
                        # Lấy năm sinh
                        nam_sinh = ""
                        if ngay_sinh_nt:
                            try:
                                if '/' in ngay_sinh_nt:
                                    parts = ngay_sinh_nt.split('/')
                                    nam_sinh = parts[-1] if len(parts) >= 3 else (parts[1] if len(parts) == 2 else parts[0][:4])
                                else:
                                    nam_sinh = ngay_sinh_nt[:4] if len(ngay_sinh_nt) >= 4 else ngay_sinh_nt
                            except:
                                nam_sinh = ""
                        
                        # Tạo chuỗi thông tin người thân
                        if nam_sinh:
                            nguoi_than_info = f"{ho_ten_nt} ({nam_sinh}, {noi_dung_nt})"
                        else:
                            nguoi_than_info = f"{ho_ten_nt} ({noi_dung_nt})"
                        
                        quan_he = nt.moiQuanHe or ''
                except:
                    pass
                
                # Tạo chuỗi họ tên với thông tin bổ sung
                ho_ten_full = p.hoTen or ''
                if p.ngaySinh:
                    ho_ten_full += f"\n({p.ngaySinh})"
                if p.nhapNgu:
                    ho_ten_full += f"\nNhập ngũ: {p.nhapNgu}"
                cb_cv = f"{p.capBac or ''}-{p.chucVu or ''}".strip('-')
                if cb_cv:
                    ho_ten_full += f"\n{cb_cv}"
                
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        ho_ten_full,
                        p.donVi or '',
                        'X' if p.thamGiaNguyQuan else '',
                        'X' if p.thamGiaNguyQuyen else '',
                        p.thamGiaNoMau or '',
                        p.queQuan or '',
                        p.truQuan or '',
                        nguoi_than_info,
                        quan_he,
                        p.daCaiTao or ''
                    )
                })
            return result
        
        # Thêm search toolbar - tạo tree_ref list để có thể cập nhật sau
        tree_ref = [None]
        get_filtered_data = self._add_search_toolbar(parent, get_data, tree_ref, None)
        
        # Tạo view tùy chỉnh với nút Chọn Quân Nhân và Xuất Word
        # Toolbar
        toolbar = tk.Frame(parent, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        title_label = tk.Label(
            toolbar,
            text="DANH SÁCH QUÂN NHÂN CÓ NGƯỜI THÂN THAM GIA CHẾ ĐỘ CŨ",
            font=('Segoe UI', 12, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        )
        title_label.pack(side=tk.LEFT, padx=10)
        
        # Buttons toolbar
        btn_container = tk.Frame(toolbar, bg=self.bg_color)
        btn_container.pack(side=tk.RIGHT, padx=5)
        
        # Treeview - tạo trước để có thể dùng trong các nút
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        # Cập nhật tree_ref
        tree_ref[0] = tree
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            elif col == 'Họ và Tên':
                tree.column(col, width=200, anchor=tk.W)
            else:
                tree.column(col, width=150, anchor=tk.W)
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Nút Chỉnh Sửa
        edit_btn = tk.Button(
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
        )
        edit_btn.pack(side=tk.LEFT, padx=3)
        
        # Nút Làm Mới
        tk.Button(
            btn_container,
            text="🔄 Làm Mới",
            command=lambda: self.refresh_list(get_filtered_data, tree, None),
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
            command=lambda: self.export_excel(get_filtered_data, "DANH SÁCH QUÂN NHÂN CÓ NGƯỜI THÂN THAM GIA CHẾ ĐỘ CŨ"),
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xuất Word
        tk.Button(
            btn_container,
            text="📄 Xuất Word",
            command=lambda: self.export_nguoi_than_che_do_cu_word(get_filtered_data),
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            elif col == 'Họ và Tên':
                tree.column(col, width=200, anchor=tk.W)
            elif col == 'Đơn Vị':
                tree.column(col, width=100, anchor=tk.CENTER)
            elif col == 'Ngụy Quân':
                tree.column(col, width=80, anchor=tk.CENTER)
            elif col == 'Ngụy Quyền':
                tree.column(col, width=80, anchor=tk.CENTER)
            elif col == 'Nợ Máu':
                tree.column(col, width=120, anchor=tk.CENTER)
            elif col == 'Quê Quán':
                tree.column(col, width=150, anchor=tk.W)
            elif col == 'Chỗ Ở':
                tree.column(col, width=150, anchor=tk.W)
            elif col == 'Họ Tên Người Thân':
                tree.column(col, width=200, anchor=tk.W)
            elif col == 'Quan Hệ':
                tree.column(col, width=100, anchor=tk.CENTER)
            elif col == 'Đã Cải Tạo':
                tree.column(col, width=120, anchor=tk.CENTER)
            else:
                tree.column(col, width=150, anchor=tk.W)
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Bind events
        tree.bind('<Double-1>', lambda e: self.on_double_click_edit(tree, None))
        tree.bind('<Button-1>', lambda e: self.on_single_click_select(tree))
        
        # Load data
        self.refresh_list(get_filtered_data, tree, None)
    
    def create_to_dan_van_tab(self, parent):
        """Tab Tổ công tác dân vận"""
        columns = ('STT', 'Họ và Tên', 'Cấp Bậc/Chức Vụ', 'Đơn Vị', 
                  'Dân Tộc', 'Tôn Giáo', 'Văn Hóa', 'Ngoại Ngữ', 'Tiếng DTTS', 'Ghi Chú')
        
        def get_data():
            # Chỉ lấy quân nhân đã được chọn vào danh sách
            selected_ids = set(self.db.get_to_dan_van())
            all_personnel = self.db.get_all()
            filtered_personnel = [p for p in all_personnel if p.id in selected_ids]
            
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            filtered_personnel = self._sort_personnel_by_cap_bac(filtered_personnel)
            
            result = []
            for idx, p in enumerate(filtered_personnel, 1):
                # Tự động điền Tiếng DTTS nếu chưa có (dựa trên dân tộc)
                tieng_dtts = p.tiengDTTS or ''
                if not tieng_dtts and p.danToc:
                    tieng_dtts = p.danToc
                
                # Lấy ghi chú riêng từ tab tổ công tác dân vận
                ghi_chu = self.db.get_to_dan_van_ghi_chu(p.id)
                
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        f"{p.capBac or ''}/{p.chucVu or ''}".strip('/'),
                        p.donVi or '',
                        p.danToc or '',
                        p.tonGiao or '',
                        p.trinhDoVanHoa or '',
                        p.ngoaiNgu or '',  # Ngoại ngữ
                        tieng_dtts,  # Tiếng DTTS (tự động từ dân tộc nếu chưa có)
                        ghi_chu  # Ghi chú riêng
                    )
                })
            return result
        
        # Thêm search toolbar - tạo tree_ref list để có thể cập nhật sau
        tree_ref = [None]
        get_filtered_data = self._add_search_toolbar(parent, get_data, tree_ref, None)
        
        # Toolbar
        toolbar = tk.Frame(parent, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        title_label = tk.Label(
            toolbar,
            text="DANH SÁCH TỔ CÔNG TÁC DÂN VẬN",
            font=('Segoe UI', 12, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        )
        title_label.pack(side=tk.LEFT, padx=10)
        
        # Buttons toolbar
        btn_container = tk.Frame(toolbar, bg=self.bg_color)
        btn_container.pack(side=tk.RIGHT, padx=5)
        
        # Treeview - tạo trước để có thể dùng trong các nút
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        # Cập nhật tree_ref
        tree_ref[0] = tree
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            elif col == 'Họ và Tên':
                tree.column(col, width=150, anchor=tk.W)
            elif col == 'Cấp Bậc/Chức Vụ':
                tree.column(col, width=120, anchor=tk.CENTER)
            elif col == 'Đơn Vị':
                tree.column(col, width=80, anchor=tk.CENTER)
            elif col == 'Dân Tộc':
                tree.column(col, width=100, anchor=tk.W)
            elif col == 'Tôn Giáo':
                tree.column(col, width=100, anchor=tk.W)
            elif col == 'Văn Hóa':
                tree.column(col, width=100, anchor=tk.CENTER)
            elif col == 'Ngoại Ngữ':
                tree.column(col, width=150, anchor=tk.CENTER)  # Tăng từ 100 lên 150
            elif col == 'Tiếng DTTS':
                tree.column(col, width=150, anchor=tk.CENTER)  # Tăng từ 100 lên 150
            elif col == 'Ghi Chú' or col == 'Ghi chú':
                tree.column(col, width=200, anchor=tk.W)  # Tăng width cho cột ghi chú
            else:
                tree.column(col, width=150, anchor=tk.W)
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Nút Chọn Quân Nhân
        tk.Button(
            btn_container,
            text="👥 Chọn Quân Nhân",
            command=lambda: self.choose_to_dan_van_personnel(parent),
            font=('Segoe UI', 10),
            bg='#9C27B0',
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
            command=lambda: self.edit_selected_from_tree(tree, None),
            font=('Segoe UI', 10),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        )
        edit_btn.pack(side=tk.LEFT, padx=3)
        
        # Nút Xóa
        def delete_selected():
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn quân nhân cần xóa!")
                return
            
            # Xác nhận xóa
            item_id = selection[0]
            values = tree.item(item_id, 'values')
            ho_ten = values[1] if len(values) > 1 else "quân nhân này"
            
            if messagebox.askyesno("Xác nhận", f"Bạn có chắc chắn muốn xóa {ho_ten} khỏi danh sách tổ công tác dân vận?"):
                if self.db.remove_to_dan_van(item_id):
                    messagebox.showinfo("Thành công", f"Đã xóa {ho_ten} khỏi danh sách!")
                    self.refresh_list(get_filtered_data, tree, None)
                else:
                    messagebox.showerror("Lỗi", "Không thể xóa quân nhân!")
        
        tk.Button(
            btn_container,
            text="🗑️ Xóa",
            command=delete_selected,
            font=('Segoe UI', 10),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Làm Mới
        tk.Button(
            btn_container,
            text="🔄 Làm Mới",
            command=lambda: self.refresh_list(get_filtered_data, tree, None),
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
            command=lambda: self.export_excel(get_filtered_data, "DANH SÁCH TỔ CÔNG TÁC DÂN VẬN"),
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xuất Word
        def export_word():
            self.export_to_dan_van_word(get_filtered_data)
        
        tk.Button(
            btn_container,
            text="📄 Xuất Word",
            command=export_word,
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Biến để lưu trữ entry widget đang edit
        editing_cell = {'item': None, 'column': None, 'entry': None, 'buttons': None}
        
        def start_edit(event):
            """Bắt đầu chỉnh sửa cell"""
            region = tree.identify_region(event.x, event.y)
            if region != "cell":
                return
            
            item = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            
            if not item or not column:
                return
            
            # Chỉ cho phép edit các cột: Ngoại Ngữ (cột 8), Tiếng DTTS (cột 9), Ghi Chú (cột 10)
            # Treeview với show='headings' thì column bắt đầu từ '#1'
            # '#1' = STT (index 0), '#2' = Họ và Tên (index 1), ..., '#8' = Ngoại Ngữ (index 7)
            col_index = int(column.replace('#', '')) - 1
            editable_columns = [7, 8, 9]  # Ngoại Ngữ (cột 8), Tiếng DTTS (cột 9), Ghi Chú (cột 10) (0-indexed)
            
            if col_index not in editable_columns:
                return
            
            # Hủy edit cũ nếu có
            if editing_cell['entry']:
                cancel_edit()
            
            # Lấy giá trị hiện tại
            values = tree.item(item, 'values')
            current_value = values[col_index] if col_index < len(values) else ''
            
            # Lấy bounding box của cell
            bbox = tree.bbox(item, column)
            if not bbox:
                return
            
            # Tạo entry widget
            entry_frame = tk.Frame(tree_frame, bg='white', relief=tk.SOLID, bd=1)
            entry_frame.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
            
            entry = tk.Entry(entry_frame, font=('Segoe UI', 9))
            entry.insert(0, current_value)
            entry.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=1)
            entry.select_range(0, tk.END)
            entry.focus()
            
            # Tạo nút checkmark và cross
            btn_frame = tk.Frame(entry_frame, bg='white')
            btn_frame.pack(side=tk.RIGHT, padx=2)
            
            def on_save(event=None):
                save_edit(item, col_index, entry.get())
            
            def on_cancel(event=None):
                cancel_edit()
            
            check_btn = tk.Button(btn_frame, text='✓', font=('Segoe UI', 10, 'bold'),
                                 bg='#4CAF50', fg='white', relief=tk.FLAT,
                                 width=2, height=1, cursor='hand2',
                                 command=on_save)
            check_btn.pack(side=tk.LEFT, padx=1)
            
            cancel_btn = tk.Button(btn_frame, text='✕', font=('Segoe UI', 10, 'bold'),
                                  bg='#F44336', fg='white', relief=tk.FLAT,
                                  width=2, height=1, cursor='hand2',
                                  command=on_cancel)
            cancel_btn.pack(side=tk.LEFT, padx=1)
            
            editing_cell['item'] = item
            editing_cell['column'] = col_index
            editing_cell['entry'] = entry
            editing_cell['buttons'] = entry_frame
            
            def on_enter(event):
                on_save()
            
            def on_escape(event):
                on_cancel()
            
            entry.bind('<Return>', on_enter)
            entry.bind('<Escape>', on_escape)
            # Không bind FocusOut để tránh tự động hủy khi click vào nút
        
        def save_edit(item, col_index, new_value):
            """Lưu giá trị đã chỉnh sửa"""
            if not editing_cell['item']:
                return
            
            # Lấy ID quân nhân - item chính là ID vì refresh_list dùng iid=personnel_id
            item_id = item
            
            if not item_id:
                cancel_edit()
                return
            
            # Lấy quân nhân
            personnel = self.db.get_by_id(item_id)
            if not personnel:
                cancel_edit()
                messagebox.showerror("Lỗi", "Không tìm thấy quân nhân!")
                return
            
            # Cập nhật giá trị - Debug để kiểm tra
            print(f"DEBUG: col_index={col_index}, new_value='{new_value}', item_id={item_id}")
            print(f"DEBUG: personnel.ngoaiNgu trước khi cập nhật: '{personnel.ngoaiNgu}'")
            
            if col_index == 7:  # Ngoại Ngữ (cột thứ 8 trong treeview)
                personnel.ngoaiNgu = new_value.strip()
                print(f"DEBUG: Đã cập nhật personnel.ngoaiNgu = '{personnel.ngoaiNgu}'")
            elif col_index == 8:  # Tiếng DTTS (cột thứ 9 trong treeview)
                personnel.tiengDTTS = new_value.strip()
                print(f"DEBUG: Đã cập nhật personnel.tiengDTTS = '{personnel.tiengDTTS}'")
            elif col_index == 9:  # Ghi Chú (cột thứ 10 trong treeview)
                # Lưu ghi chú riêng cho tab tổ công tác dân vận
                if self.db.update_to_dan_van_ghi_chu(item_id, new_value.strip()):
                    print(f"DEBUG: Đã cập nhật ghi chú riêng cho to_dan_van: '{new_value.strip()}'")
                else:
                    messagebox.showerror("Lỗi", "Không thể cập nhật ghi chú!")
                    cancel_edit()
                    return
                # Không cần cập nhật personnel.ghiChu nữa vì dùng ghi chú riêng
                # personnel.ghiChu = new_value.strip()
                # print(f"DEBUG: Đã cập nhật personnel.ghiChu = '{personnel.ghiChu}'")
            
            # Tự động điền Tiếng DTTS từ dân tộc nếu trống
            if not personnel.tiengDTTS and personnel.danToc:
                personnel.tiengDTTS = personnel.danToc
            
            # Lưu vào database
            try:
                # Đảm bảo dữ liệu được cập nhật đúng
                print(f"DEBUG: Gọi db.update() với personnel.ngoaiNgu = '{personnel.ngoaiNgu}'")
                result = self.db.update(personnel)
                print(f"DEBUG: Kết quả update: {result}")
                
                if result:
                    # Kiểm tra lại sau khi update
                    personnel_after = self.db.get_by_id(item_id)
                    if personnel_after:
                        print(f"DEBUG: personnel.ngoaiNgu sau khi update: '{personnel_after.ngoaiNgu}'")
                    
                    # Hủy edit trước để tránh conflict
                    cancel_edit()
                    # Refresh lại toàn bộ tree từ database để đảm bảo dữ liệu đồng bộ
                    self.refresh_list(get_filtered_data, tree, None)
                    messagebox.showinfo("Thành công", "Đã lưu thành công!")
                else:
                    messagebox.showerror("Lỗi", "Không thể lưu dữ liệu! Có thể cột chưa được tạo trong database.\nVui lòng khởi động lại ứng dụng.")
            except Exception as e:
                error_msg = str(e)
                print(f"DEBUG: Lỗi khi lưu: {error_msg}")
                messagebox.showerror("Lỗi", f"Không thể lưu: {error_msg}\n\nVui lòng khởi động lại ứng dụng để cập nhật database.")
                import traceback
                traceback.print_exc()
                cancel_edit()
        
        def cancel_edit():
            """Hủy chỉnh sửa"""
            if editing_cell['buttons']:
                editing_cell['buttons'].destroy()
            editing_cell['item'] = None
            editing_cell['column'] = None
            editing_cell['entry'] = None
            editing_cell['buttons'] = None
        
        # Bind events - double click để edit
        tree.bind('<Double-1>', start_edit)
        def on_click(event):
            cancel_edit()
            # Chỉ select nếu không phải đang edit
            if not editing_cell['entry']:
                item = tree.identify_row(event.y)
                if item:
                    tree.selection_set(item)
        tree.bind('<Button-1>', on_click)
        
        # Load data
        self.refresh_list(get_filtered_data, tree, None)
    
    def create_ban_chap_hanh_tab(self, parent):
        """Tab Ban chấp hành Chi đoàn"""
        columns = ('STT', 'Họ và Tên', 'Ngày Sinh', 'Cấp Bậc', 'Chức Vụ',
                  'Nhập Ngũ', 'Đơn Vị', 'Ngày Vào Đoàn', 'Chức Vụ Đoàn')
        
        def get_data():
            # Lấy danh sách ID quân nhân trong ban chấp hành
            ban_chap_hanh_ids = self.db.get_ban_chap_hanh_chi_doan()
            if not ban_chap_hanh_ids:
                return []
            
            all_personnel = self.db.get_all()
            # Lọc chỉ những quân nhân trong ban chấp hành
            ban_chap_hanh = [p for p in all_personnel if p.id in ban_chap_hanh_ids]
            
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            ban_chap_hanh = self._sort_personnel_by_cap_bac(ban_chap_hanh)
            
            result = []
            for idx, p in enumerate(ban_chap_hanh, 1):
                # Lấy chức vụ đoàn từ bảng ban_chap_hanh_chi_doan
                chuc_vu_doan = self.db.get_chuc_vu_doan(p.id)
                if not chuc_vu_doan:
                    chuc_vu_doan = p.thongTinKhac.doan.chucVuDoan or ''
                
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
                        chuc_vu_doan
                    )
                })
            return result
        
        # Toolbar
        toolbar = tk.Frame(parent, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        title_label = tk.Label(
            toolbar,
            text="BAN CHẤP HÀNH CHI ĐOÀN ĐẠI ĐỘI 3",
            font=('Segoe UI', 12, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        )
        title_label.pack(side=tk.LEFT, padx=10)
        
        # Thêm search toolbar - tạo tree_ref list để có thể cập nhật sau
        tree_ref = [None]
        get_filtered_data = self._add_search_toolbar(parent, get_data, tree_ref, None)
        
        # Buttons toolbar
        btn_container = tk.Frame(toolbar, bg=self.bg_color)
        btn_container.pack(side=tk.RIGHT, padx=5)
        
        # Treeview - tạo trước để có thể dùng trong các nút
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        # Cập nhật tree_ref
        tree_ref[0] = tree
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            elif col == 'Họ và Tên':
                tree.column(col, width=200, anchor=tk.W)
            elif col == 'Chức Vụ Đoàn':
                tree.column(col, width=150, anchor=tk.W)
            else:
                tree.column(col, width=120, anchor=tk.W)
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Load data
        def refresh_list():
            tree = tree_ref[0] if tree_ref else None
            if tree:
                for item in tree.get_children():
                    tree.delete(item)
                data = get_filtered_data()
                for item_data in data:
                    tree.insert('', tk.END, values=item_data['values'], tags=(item_data.get('id'),))
        
        # Nút Chọn quân nhân chi đoàn
        def choose_personnel():
            self.choose_ban_chap_hanh_personnel(parent, tree_ref)
        
        tk.Button(
            btn_container,
            text="👥 Chọn Quân Nhân Chi Đoàn",
            command=choose_personnel,
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xóa khỏi ban chấp hành
        def remove_from_ban_chap_hanh():
            tree = tree_ref[0] if tree_ref else None
            if not tree:
                return
            
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn quân nhân cần xóa!")
                return
            
            # Lấy personnel_id từ tags của item
            item = selected[0]
            tags = tree.item(item, 'tags')
            personnel_id = tags[0] if tags else None
            
            if not personnel_id:
                messagebox.showerror("Lỗi", "Không tìm thấy ID quân nhân!")
                return
            
            if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa quân nhân này khỏi ban chấp hành?"):
                if self.db.remove_ban_chap_hanh_chi_doan(personnel_id):
                    messagebox.showinfo("Thành công", "Đã xóa quân nhân khỏi ban chấp hành!")
                    refresh_list()
                else:
                    messagebox.showerror("Lỗi", "Không thể xóa quân nhân!")
        
        tk.Button(
            btn_container,
            text="❌ Xóa Khỏi Ban Chấp Hành",
            command=remove_from_ban_chap_hanh,
            font=('Segoe UI', 10),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        tk.Button(
            btn_container,
            text="🔄 Làm Mới",
            command=refresh_list,
            font=('Segoe UI', 10),
            bg='#388E3C',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xuất Word
        def export_word():
            self.export_ban_chap_hanh_word(parent, get_filtered_data)
        
        tk.Button(
            btn_container,
            text="📄 Xuất Word",
            command=export_word,
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Load data ban đầu
        refresh_list()
    
    def export_ban_chap_hanh_word(self, parent, get_data_func):
        """Dialog xuất Word cho Ban chấp hành Chi đoàn"""
        dialog = tk.Toplevel(parent)
        dialog.title("Xuất File Word - Ban Chấp Hành Chi Đoàn")
        dialog.geometry("500x400")
        dialog.transient(parent)
        dialog.grab_set()
        
        # Frame chứa form
        form_frame = tk.Frame(dialog, bg='#FAFAFA', padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tiêu đề
        title_label = tk.Label(
            form_frame,
            text="Thiết lập thông tin xuất file",
            font=('Segoe UI', 12, 'bold'),
            bg='#FAFAFA',
            fg='#388E3C'
        )
        title_label.pack(pady=(0, 20))
        
        # Đơn vị
        tk.Label(
            form_frame,
            text="Đơn vị:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        don_vi_var = tk.StringVar(value="Đại đội 3")
        don_vi_entry = tk.Entry(form_frame, textvariable=don_vi_var, width=40, font=('Segoe UI', 10))
        don_vi_entry.pack(fill=tk.X, pady=5)
        
        # Tiểu đoàn
        tk.Label(
            form_frame,
            text="Tiểu đoàn:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
        tieu_doan_entry = tk.Entry(form_frame, textvariable=tieu_doan_var, width=40, font=('Segoe UI', 10))
        tieu_doan_entry.pack(fill=tk.X, pady=5)
        
        # Địa điểm
        tk.Label(
            form_frame,
            text="Địa điểm:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        dia_diem_var = tk.StringVar(value="Đăk Lăk")
        dia_diem_entry = tk.Entry(form_frame, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10))
        dia_diem_entry.pack(fill=tk.X, pady=5)
        
        # Tên Bí thư
        tk.Label(
            form_frame,
            text="Tên Bí thư:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        ten_bi_thu_var = tk.StringVar(value="Trần Quỳnh Thương")
        ten_bi_thu_entry = tk.Entry(form_frame, textvariable=ten_bi_thu_var, width=40, font=('Segoe UI', 10))
        ten_bi_thu_entry.pack(fill=tk.X, pady=5)
        
        # Buttons
        btn_frame = tk.Frame(form_frame, bg='#FAFAFA')
        btn_frame.pack(fill=tk.X, pady=20)
        
        def save_and_export():
            # Lấy dữ liệu
            don_vi = don_vi_var.get().strip()
            tieu_doan = tieu_doan_var.get().strip()
            dia_diem = dia_diem_var.get().strip()
            ten_bi_thu = ten_bi_thu_var.get().strip()
            
            if not don_vi:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập đơn vị!")
                return
            
            # Lấy danh sách quân nhân
            data = get_data_func()
            if not data:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            # Lấy danh sách Personnel từ IDs
            ban_chap_hanh_ids = self.db.get_ban_chap_hanh_chi_doan()
            all_personnel = self.db.get_all()
            personnel_list = [p for p in all_personnel if p.id in ban_chap_hanh_ids]
            
            if not personnel_list:
                messagebox.showwarning("Cảnh báo", "Không có quân nhân trong ban chấp hành!")
                return
            
            # Chọn file để lưu
            filename = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                initialfile=f"Ban_Chap_Hanh_Chi_Doan_{don_vi.replace(' ', '_')}.docx"
            )
            
            if not filename:
                return
            
            try:
                # Import và gọi hàm xuất
                from services.export_ban_chap_hanh_chi_doan import to_word_docx_ban_chap_hanh_chi_doan
                
                word_content = to_word_docx_ban_chap_hanh_chi_doan(
                    personnel_list=personnel_list,
                    don_vi=don_vi,
                    tieu_doan=tieu_doan,
                    dia_diem=dia_diem,
                    ten_bi_thu=ten_bi_thu,
                    db_service=self.db
                )
                
                # Lưu file
                with open(filename, 'wb') as f:
                    f.write(word_content)
                
                messagebox.showinfo("Thành công", f"Đã xuất file Word thành công!\n{filename}")
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        
        tk.Button(
            btn_frame,
            text="Xuất File",
            command=save_and_export,
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.RIGHT, padx=5)
        
        tk.Button(
            btn_frame,
            text="Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.RIGHT, padx=5)
    
    def choose_ban_chap_hanh_personnel(self, parent, tree_ref=None):
        """Dialog chọn quân nhân vào ban chấp hành chi đoàn"""
        dialog = tk.Toplevel(parent)
        dialog.title("Chọn Quân Nhân Chi Đoàn")
        dialog.geometry("900x650")
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(True, True)
        
        # Dùng grid để control layout tốt hơn
        dialog.grid_rowconfigure(0, weight=1)  # Row 0 (list_frame) có thể expand
        dialog.grid_rowconfigure(1, weight=0)  # Row 1 (chuc_vu_frame) không expand
        dialog.grid_rowconfigure(2, weight=0)  # Row 2 (btn_frame) không expand
        dialog.grid_columnconfigure(0, weight=1)
        
        # Frame chứa danh sách
        list_frame = tk.Frame(dialog, bg='#FAFAFA')
        list_frame.grid(row=0, column=0, padx=10, pady=10, sticky=tk.NSEW)
        
        # Label
        label = tk.Label(
            list_frame,
            text="Chọn quân nhân để thêm vào ban chấp hành chi đoàn:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        )
        label.pack(anchor=tk.W, pady=5)
        
        # Toolbar với tìm kiếm
        toolbar_frame = tk.Frame(list_frame, bg='#FAFAFA')
        toolbar_frame.pack(fill=tk.X, pady=5)
        
        # Tìm kiếm
        tk.Label(toolbar_frame, text="🔍 Tìm kiếm:", font=('Segoe UI', 9), bg='#FAFAFA').pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(toolbar_frame, textvariable=search_var, width=30, font=('Segoe UI', 9))
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Treeview với checkbox
        columns = ('Họ và Tên', 'Ngày Sinh', 'Cấp Bậc', 'Chức Vụ', 'Đơn Vị', 'Ngày Vào Đoàn')
        tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=15)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Configure columns
        tree.heading('#0', text='Chọn')
        tree.column('#0', width=50, anchor=tk.CENTER)
        for col in columns:
            tree.heading(col, text=col)
            if col == 'Họ và Tên':
                tree.column(col, width=200)
            else:
                tree.column(col, width=120)
        
        # Load data - chỉ hiển thị đoàn viên
        all_personnel = self.db.get_all()
        doan_vien = [p for p in all_personnel if p.thongTinKhac.doan.ngayVao]
        ban_chap_hanh_ids = set(self.db.get_ban_chap_hanh_chi_doan())
        
        selected_ids = set()
        for p in doan_vien:
            is_selected = p.id in ban_chap_hanh_ids
            if is_selected:
                selected_ids.add(p.id)
        
        def load_tree_data():
            """Load dữ liệu vào tree với filter"""
            # Xóa dữ liệu cũ
            for item in tree.get_children():
                tree.delete(item)
            
            # Lọc theo tìm kiếm
            search_text = search_var.get().lower()
            display_personnel = doan_vien
            if search_text:
                display_personnel = [p for p in doan_vien 
                                  if search_text in (p.hoTen or '').lower() or
                                     search_text in (p.capBac or '').lower() or
                                     search_text in (p.chucVu or '').lower()]
            
            # Sắp xếp theo cấp bậc
            def _parse_cap_bac_rank(cap_bac: str) -> int:
                if not cap_bac:
                    return 0
                cap_bac = cap_bac.strip().upper()
                if 'ĐẠI TÁ' in cap_bac: return 100
                elif 'TRUNG TÁ' in cap_bac: return 90
                elif 'THIẾU TÁ' in cap_bac: return 80
                elif 'ĐẠI ÚY' in cap_bac: return 70
                elif 'THƯỢNG ÚY' in cap_bac: return 60
                elif 'TRUNG ÚY' in cap_bac: return 50
                elif 'THIẾU ÚY' in cap_bac: return 40
                elif 'THƯỢNG SĨ' in cap_bac: return 30
                elif 'TRUNG SĨ' in cap_bac: return 20
                elif 'HẠ SĨ' in cap_bac: return 10
                elif cap_bac.startswith('H'):
                    try:
                        return int(cap_bac[1:])
                    except:
                        return 0
                else:
                    try:
                        return int(cap_bac) + 10
                    except:
                        return 0
            
            def sort_key(p):
                cap_bac_rank = _parse_cap_bac_rank(p.capBac or '')
                ho_ten = (p.hoTen or '').lower()
                return (-cap_bac_rank, ho_ten)
            
            display_personnel = sorted(display_personnel, key=sort_key)
            
            for p in display_personnel:
                is_selected = p.id in selected_ids
                tree.insert('', tk.END, 
                           text='✓' if is_selected else '',
                           values=(
                               p.hoTen or '',
                               p.ngaySinh or '',
                               p.capBac or '',
                               p.chucVu or '',
                               p.donVi or '',
                               p.thongTinKhac.doan.ngayVao or ''
                           ),
                           tags=(p.id,))
        
        # Bind click để toggle
        def toggle_selection(event):
            item = tree.identify_row(event.y)
            if not item:
                return
            
            item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
            if not item_id:
                return
            
            current_text = tree.item(item, 'text')
            if current_text == '✓':
                tree.item(item, text='')
                selected_ids.discard(item_id)
            else:
                tree.item(item, text='✓')
                selected_ids.add(item_id)
        
        tree.bind('<Button-1>', toggle_selection)
        search_var.trace('w', lambda *args: load_tree_data())
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Load dữ liệu ban đầu
        load_tree_data()
        
        # Frame chức vụ đoàn - Row 1
        chuc_vu_frame = tk.Frame(dialog, bg='#FAFAFA')
        chuc_vu_frame.grid(row=1, column=0, padx=10, pady=5, sticky=tk.EW)
        
        tk.Label(
            chuc_vu_frame,
            text="Chức vụ đoàn (nếu có):",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(side=tk.LEFT, padx=5)
        
        chuc_vu_var = tk.StringVar()
        chuc_vu_entry = tk.Entry(chuc_vu_frame, textvariable=chuc_vu_var, width=30, font=('Segoe UI', 10))
        chuc_vu_entry.pack(side=tk.LEFT, padx=5)
        chuc_vu_entry.insert(0, "Bí thư, UV, ...")
        
        # Buttons - Row 2, LUÔN HIỂN THỊ
        btn_frame = tk.Frame(dialog, bg='#FAFAFA', height=70)
        btn_frame.grid(row=2, column=0, padx=10, pady=10, sticky=tk.EW)
        btn_frame.grid_propagate(False)
        
        def save_selection():
            chuc_vu_doan = chuc_vu_var.get().strip()
            if chuc_vu_doan == "Bí thư, UV, ...":
                chuc_vu_doan = ""
            
            # Lưu tất cả quân nhân đã chọn
            success_count = 0
            for personnel_id in selected_ids:
                if self.db.add_ban_chap_hanh_chi_doan(personnel_id, chuc_vu_doan):
                    success_count += 1
            
            # Xóa những quân nhân không được chọn
            all_ban_chap_hanh_ids = set(self.db.get_ban_chap_hanh_chi_doan())
            to_remove = all_ban_chap_hanh_ids - selected_ids
            for personnel_id in to_remove:
                self.db.remove_ban_chap_hanh_chi_doan(personnel_id)
            
            messagebox.showinfo("Thành công", f"Đã cập nhật {success_count} quân nhân vào ban chấp hành!")
            dialog.destroy()
            # Refresh lại danh sách
            if tree_ref and tree_ref[0]:
                tree = tree_ref[0]
                # Clear và reload
                for item in tree.get_children():
                    tree.delete(item)
                # Reload data
                ban_chap_hanh_ids = self.db.get_ban_chap_hanh_chi_doan()
                if ban_chap_hanh_ids:
                    all_personnel = self.db.get_all()
                    ban_chap_hanh = [p for p in all_personnel if p.id in ban_chap_hanh_ids]
                    # Sắp xếp theo cấp bậc
                    def _parse_cap_bac_rank(cap_bac: str) -> int:
                        if not cap_bac:
                            return 0
                        cap_bac = cap_bac.strip().upper()
                        if 'ĐẠI TÁ' in cap_bac: return 100
                        elif 'TRUNG TÁ' in cap_bac: return 90
                        elif 'THIẾU TÁ' in cap_bac: return 80
                        elif 'ĐẠI ÚY' in cap_bac: return 70
                        elif 'THƯỢNG ÚY' in cap_bac: return 60
                        elif 'TRUNG ÚY' in cap_bac: return 50
                        elif 'THIẾU ÚY' in cap_bac: return 40
                        elif 'THƯỢNG SĨ' in cap_bac: return 30
                        elif 'TRUNG SĨ' in cap_bac: return 20
                        elif 'HẠ SĨ' in cap_bac: return 10
                        elif cap_bac.startswith('H'):
                            try:
                                return int(cap_bac[1:])
                            except:
                                return 0
                        else:
                            try:
                                return int(cap_bac) + 10
                            except:
                                return 0
                    
                    def sort_key(p):
                        cap_bac_rank = _parse_cap_bac_rank(p.capBac or '')
                        ho_ten = (p.hoTen or '').lower()
                        return (-cap_bac_rank, ho_ten)
                    
                    ban_chap_hanh = sorted(ban_chap_hanh, key=sort_key)
                    
                    for idx, p in enumerate(ban_chap_hanh, 1):
                        chuc_vu_doan = self.db.get_chuc_vu_doan(p.id)
                        if not chuc_vu_doan:
                            chuc_vu_doan = p.thongTinKhac.doan.chucVuDoan or ''
                        tree.insert('', tk.END, 
                                   values=(
                                       idx,
                                       p.hoTen or '',
                                       p.ngaySinh or '',
                                       p.capBac or '',
                                       p.chucVu or '',
                                       p.nhapNgu or '',
                                       p.donVi or '',
                                       p.thongTinKhac.doan.ngayVao or '',
                                       chuc_vu_doan
                                   ),
                                   tags=(p.id,))
        
        # Buttons layout với grid
        btn_frame.grid_columnconfigure(1, weight=1)
        
        # Nút Hủy
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor='hand2',
            width=10
        ).grid(row=0, column=0, padx=5, sticky=tk.W)
        
        # Spacer
        tk.Frame(btn_frame, bg='#FAFAFA').grid(row=0, column=1, sticky=tk.EW)
        
        # Nút XONG
        tk.Button(
            btn_frame,
            text="✅ XONG",
            command=save_selection,
            font=('Segoe UI', 11, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.RAISED,
            padx=25,
            pady=8,
            cursor='hand2',
            width=12,
            bd=2
        ).grid(row=0, column=2, padx=5, sticky=tk.E)
    
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
            
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            ton_giao = self._sort_personnel_by_cap_bac(ton_giao)
            
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
        
        # Toolbar
        toolbar = tk.Frame(parent, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        title_label = tk.Label(
            toolbar,
            text="QUÂN NHÂN THEO TÔN GIÁO",
            font=('Segoe UI', 12, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        )
        title_label.pack(side=tk.LEFT, padx=10)
        
        # Thêm search toolbar - tạo tree_ref list để có thể cập nhật sau
        tree_ref = [None]
        get_filtered_data = self._add_search_toolbar(parent, get_data, tree_ref, None)
        
        # Buttons toolbar
        btn_container = tk.Frame(toolbar, bg=self.bg_color)
        btn_container.pack(side=tk.RIGHT, padx=5)
        
        # Treeview
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        # Cập nhật tree_ref
        tree_ref[0] = tree
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            elif col == 'Họ Tên':
                tree.column(col, width=200, anchor=tk.W)
            elif col == 'Tôn Giáo':
                tree.column(col, width=150, anchor=tk.W)
            else:
                tree.column(col, width=120, anchor=tk.W)
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Nút Xuất Word
        def export_word():
            self.export_ton_giao_word(parent, get_filtered_data)
        
        tk.Button(
            btn_container,
            text="📄 Xuất Word",
            command=export_word,
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Load data
        def refresh_list():
            for item in tree.get_children():
                tree.delete(item)
            data = get_filtered_data()
            for item_data in data:
                tree.insert('', tk.END, values=item_data['values'], tags=(item_data.get('id'),))
        
        refresh_list()
    
    def export_ton_giao_word(self, parent, get_data_func):
        """Dialog xuất Word cho Quân Nhân Theo Tôn Giáo"""
        dialog = tk.Toplevel(parent)
        dialog.title("Xuất File Word - Quân Nhân Theo Tôn Giáo")
        dialog.geometry("500x400")
        dialog.transient(parent)
        dialog.grab_set()
        
        # Frame chứa form
        form_frame = tk.Frame(dialog, bg='#FAFAFA', padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tiêu đề
        title_label = tk.Label(
            form_frame,
            text="Thiết lập thông tin xuất file",
            font=('Segoe UI', 12, 'bold'),
            bg='#FAFAFA',
            fg='#388E3C'
        )
        title_label.pack(pady=(0, 20))
        
        # Đơn vị
        tk.Label(
            form_frame,
            text="Đơn vị:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        don_vi_var = tk.StringVar(value="Đại đội 3")
        don_vi_entry = tk.Entry(form_frame, textvariable=don_vi_var, width=40, font=('Segoe UI', 10))
        don_vi_entry.pack(fill=tk.X, pady=5)
        
        # Tiểu đoàn
        tk.Label(
            form_frame,
            text="Tiểu đoàn:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
        tieu_doan_entry = tk.Entry(form_frame, textvariable=tieu_doan_var, width=40, font=('Segoe UI', 10))
        tieu_doan_entry.pack(fill=tk.X, pady=5)
        
        # Địa điểm
        tk.Label(
            form_frame,
            text="Địa điểm:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        dia_diem_var = tk.StringVar(value="Đắk Lắk")
        dia_diem_entry = tk.Entry(form_frame, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10))
        dia_diem_entry.pack(fill=tk.X, pady=5)
        
        # Chính trị viên
        tk.Label(
            form_frame,
            text="Chính trị viên:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        chinh_tri_vien_var = tk.StringVar(value="Đại úy Triệu Văn Dũng")
        chinh_tri_vien_entry = tk.Entry(form_frame, textvariable=chinh_tri_vien_var, width=40, font=('Segoe UI', 10))
        chinh_tri_vien_entry.pack(fill=tk.X, pady=5)
        
        # Buttons
        btn_frame = tk.Frame(form_frame, bg='#FAFAFA')
        btn_frame.pack(fill=tk.X, pady=20)
        
        def save_and_export():
            # Lấy dữ liệu
            don_vi = don_vi_var.get().strip()
            tieu_doan = tieu_doan_var.get().strip()
            dia_diem = dia_diem_var.get().strip()
            chinh_tri_vien = chinh_tri_vien_var.get().strip()
            
            if not don_vi:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập đơn vị!")
                return
            
            # Lấy danh sách quân nhân
            data = get_data_func()
            if not data:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            # Lấy danh sách Personnel từ IDs
            all_personnel = self.db.get_all()
            ton_giao_personnel = [p for p in all_personnel if p.tonGiao and p.tonGiao.strip()]
            
            if not ton_giao_personnel:
                messagebox.showwarning("Cảnh báo", "Không có quân nhân theo tôn giáo!")
                return
            
            # Chọn file để lưu
            filename = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                initialfile=f"Quan_Nhan_Theo_Ton_Giao_{don_vi.replace(' ', '_')}.docx"
            )
            
            if not filename:
                return
            
            try:
                # Import và gọi hàm xuất
                from services.export_ton_giao import to_word_docx_ton_giao
                
                word_content = to_word_docx_ton_giao(
                    personnel_list=ton_giao_personnel,
                    don_vi=don_vi,
                    tieu_doan=tieu_doan,
                    dia_diem=dia_diem,
                    chinh_tri_vien=chinh_tri_vien
                )
                
                # Lưu file
                with open(filename, 'wb') as f:
                    f.write(word_content)
                
                messagebox.showinfo("Thành công", f"Đã xuất file Word thành công!\n{filename}")
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        
        tk.Button(
            btn_frame,
            text="Xuất File",
            command=save_and_export,
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.RIGHT, padx=5)
        
        tk.Button(
            btn_frame,
            text="Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.RIGHT, padx=5)
    
    def create_dang_phai_phan_dong_tab(self, parent):
        """Tab Người thân đảng phái phản động"""
        columns = ('STT', 'Họ và Tên QN', 'Ngày Sinh', 'Cấp Bậc-Chức Vụ',
                  'Đơn Vị', 'Họ Tên Người Thân', 'Mối Quan Hệ', 'Nội Dung')
        
        def get_data():
            # Chỉ lấy quân nhân có checkbox "Tham gia đảng phái phản động" được đánh dấu
            selected_ids = set(self.db.get_nguoi_than_dang_phai_phan_dong())
            all_personnel = self.db.get_all()
            filtered_personnel = [p for p in all_personnel if p.id in selected_ids]
            
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            filtered_personnel = self._sort_personnel_by_cap_bac(filtered_personnel)
            
            result = []
            stt = 1
            
            for p in filtered_personnel:
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
        
        # Toolbar
        toolbar = tk.Frame(parent, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        title_label = tk.Label(
            toolbar,
            text="QUÂN NHÂN CÓ NGƯỜI THÂN THAM GIA ĐẢNG PHÁI PHẢN ĐỘNG",
            font=('Segoe UI', 12, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        )
        title_label.pack(side=tk.LEFT, padx=10)
        
        # Thêm search toolbar - tạo tree_ref list để có thể cập nhật sau
        tree_ref = [None]
        get_filtered_data = self._add_search_toolbar(parent, get_data, tree_ref, None)
        
        # Buttons toolbar
        btn_container = tk.Frame(toolbar, bg=self.bg_color)
        btn_container.pack(side=tk.RIGHT, padx=5)
        
        # Treeview - tạo trước để có thể dùng trong các nút
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        # Cập nhật tree_ref
        tree_ref[0] = tree
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            elif col == 'Họ và Tên QN':
                tree.column(col, width=200, anchor=tk.W)
            elif col == 'Họ Tên Người Thân':
                tree.column(col, width=200, anchor=tk.W)
            elif col == 'Nội Dung':
                tree.column(col, width=250, anchor=tk.W)
            else:
                tree.column(col, width=150, anchor=tk.W)
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Nút Thêm Mới
        tk.Button(
            btn_container,
            text="➕ Thêm Mới",
            command=lambda: self.edit_selected_from_tree(tree, None),
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
            command=lambda: self.edit_selected_from_tree(tree, None),
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
            command=lambda: self.refresh_list(get_filtered_data, tree, None),
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
            command=lambda: self.export_excel(get_filtered_data, "QUÂN NHÂN CÓ NGƯỜI THÂN THAM GIA ĐẢNG PHÁI PHẢN ĐỘNG"),
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xuất Word
        def export_word():
            self.export_dang_phai_phan_dong_word(get_filtered_data)
        
        tk.Button(
            btn_container,
            text="📄 Xuất Word",
            command=export_word,
            font=('Segoe UI', 10),
            bg='#4CAF50',
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
        self.refresh_list(get_filtered_data, tree, None)
    
    def create_yeu_to_nuoc_ngoai_tab(self, parent):
        """Tab Yếu tố nước ngoài"""
        columns = ('STT', 'Họ và Tên', 'Ngày Sinh', 'Cấp Bậc-Chức Vụ',
                  'Đơn Vị', 'Nội Dung Yếu Tố NN', 'Mối Quan Hệ', 'Tên Nước')
        
        def get_data():
            all_personnel = self.db.get_all()
            # Lọc chỉ có yếu tố nước ngoài
            yeu_to_nn = [p for p in all_personnel if p.thongTinKhac.yeuToNN]
            
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            yeu_to_nn = self._sort_personnel_by_cap_bac(yeu_to_nn)
            
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
        columns = ('STT', 'Họ và Tên', 'Cấp Bậc', 'Chức Vụ', 'Đơn Vị', 'Thông tin người thân', 'Thời gian vào', 'Thời gian ra')
        
        def get_data():
            # Lấy danh sách ID quân nhân trong bảo vệ an ninh
            bao_ve_ids = self.db.get_bao_ve_an_ninh()
            if not bao_ve_ids:
                return []
            
            all_personnel = self.db.get_all()
            # Lọc chỉ những quân nhân trong bảo vệ an ninh
            bao_ve_personnel = [p for p in all_personnel if p.id in bao_ve_ids]
            
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            bao_ve_personnel = self._sort_personnel_by_cap_bac(bao_ve_personnel)
            
            result = []
            for idx, p in enumerate(bao_ve_personnel, 1):
                # Lấy thông tin thời gian vào/ra
                bao_ve_info = self.db.get_bao_ve_an_ninh_info(p.id)
                thoi_gian_vao = bao_ve_info.get('thoiGianVao', '') or ''
                thoi_gian_ra = bao_ve_info.get('thoiGianRa', '') or ''
                
                # Lấy thông tin người thân
                gia_dinh_info = []
                try:
                    nguoi_than_list = self.db.get_nguoi_than_by_personnel(p.id)
                    
                    bo_de = []
                    me_de = []
                    vo = []
                    
                    for nguoi_than in nguoi_than_list:
                        moi_quan_he = (nguoi_than.moiQuanHe or '').lower().strip()
                        ho_ten = (nguoi_than.hoTen or '').strip()
                        ngay_sinh = (nguoi_than.ngaySinh or '').strip()
                        noi_dung = (nguoi_than.noiDung or '').strip()
                        
                        if not ho_ten:
                            continue
                        
                        # Lấy năm sinh
                        nam_sinh = ""
                        if ngay_sinh:
                            try:
                                if '/' in ngay_sinh:
                                    parts = ngay_sinh.split('/')
                                    nam_sinh = parts[-1] if len(parts) >= 3 else (parts[1] if len(parts) == 2 else parts[0][:4])
                                else:
                                    nam_sinh = ngay_sinh[:4] if len(ngay_sinh) >= 4 else ngay_sinh
                            except:
                                nam_sinh = ""
                        
                        nghe = noi_dung if noi_dung else "làm nông"
                        
                        if nam_sinh:
                            info_str = f"{ho_ten} ({nam_sinh}-{nghe})"
                        else:
                            info_str = f"{ho_ten} ({nghe})"
                        
                        # Phân loại
                        if ('bố đẻ' in moi_quan_he or 'cha đẻ' in moi_quan_he or 
                            (('bố' in moi_quan_he or 'cha' in moi_quan_he) and 
                             'vợ' not in moi_quan_he and 'vo' not in moi_quan_he)):
                            bo_de.append(info_str)
                        elif ('mẹ đẻ' in moi_quan_he or 'me đẻ' in moi_quan_he or 
                              (('mẹ' in moi_quan_he or 'me' in moi_quan_he) and 
                               'vợ' not in moi_quan_he and 'vo' not in moi_quan_he)):
                            me_de.append(info_str)
                        elif 'vợ' in moi_quan_he or 'vo' in moi_quan_he:
                            vo.append(info_str)
                        elif 'bố' in moi_quan_he or 'cha' in moi_quan_he:
                            if 'vợ' not in moi_quan_he and 'vo' not in moi_quan_he:
                                bo_de.append(info_str)
                        elif 'mẹ' in moi_quan_he or 'me' in moi_quan_he:
                            if 'vợ' not in moi_quan_he and 'vo' not in moi_quan_he:
                                me_de.append(info_str)
                    
                    gia_dinh_info.extend(bo_de)
                    gia_dinh_info.extend(me_de)
                    gia_dinh_info.extend(vo)
                    
                except Exception:
                    # Fallback: sử dụng các field cũ
                    if p.hoTenCha:
                        nam_sinh_cha = ""
                        if p.ngaySinhCha:
                            try:
                                if '/' in p.ngaySinhCha:
                                    parts = p.ngaySinhCha.split('/')
                                    nam_sinh_cha = parts[-1] if len(parts) >= 3 else (parts[1] if len(parts) == 2 else parts[0][:4])
                                else:
                                    nam_sinh_cha = p.ngaySinhCha[:4] if len(p.ngaySinhCha) >= 4 else p.ngaySinhCha
                            except:
                                pass
                        gia_dinh_info.append(f"{p.hoTenCha} ({nam_sinh_cha}-làm nông)" if nam_sinh_cha else f"{p.hoTenCha} (làm nông)")
                    if p.hoTenMe:
                        nam_sinh_me = ""
                        if p.ngaySinhMe:
                            try:
                                if '/' in p.ngaySinhMe:
                                    parts = p.ngaySinhMe.split('/')
                                    nam_sinh_me = parts[-1] if len(parts) >= 3 else (parts[1] if len(parts) == 2 else parts[0][:4])
                                else:
                                    nam_sinh_me = p.ngaySinhMe[:4] if len(p.ngaySinhMe) >= 4 else p.ngaySinhMe
                            except:
                                pass
                        gia_dinh_info.append(f"{p.hoTenMe} ({nam_sinh_me}-làm nông)" if nam_sinh_me else f"{p.hoTenMe} (làm nông)")
                    if p.hoTenVo:
                        nam_sinh_vo = ""
                        if p.ngaySinhVo:
                            try:
                                if '/' in p.ngaySinhVo:
                                    parts = p.ngaySinhVo.split('/')
                                    nam_sinh_vo = parts[-1] if len(parts) >= 3 else (parts[1] if len(parts) == 2 else parts[0][:4])
                                else:
                                    nam_sinh_vo = p.ngaySinhVo[:4] if len(p.ngaySinhVo) >= 4 else p.ngaySinhVo
                            except:
                                pass
                        gia_dinh_info.append(f"{p.hoTenVo} ({nam_sinh_vo}-GV)" if nam_sinh_vo else f"{p.hoTenVo} (GV)")
                
                thong_tin_nguoi_than = " / ".join(gia_dinh_info) if gia_dinh_info else ''
                
                result.append({
                    'id': p.id,
                    'values': (
                        idx,
                        p.hoTen or '',
                        p.capBac or '',
                        p.chucVu or '',
                        p.donVi or '',
                        thong_tin_nguoi_than,
                        thoi_gian_vao,
                        thoi_gian_ra
                    )
                })
            return result
        
        # Toolbar
        toolbar = tk.Frame(parent, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        title_label = tk.Label(
            toolbar,
            text="BÍ THƯ CẤP UỶ, CHI BỘ PHỤ TRÁCH CÔNG TÁC BVAN VÀ CHIẾN SỸ BẢO VỆ",
            font=('Segoe UI', 12, 'bold'),
            bg=self.bg_color,
            fg='#388E3C'
        )
        title_label.pack(side=tk.LEFT, padx=10)
        
        # Thêm search toolbar - tạo tree_ref list để có thể cập nhật sau
        tree_ref = [None]
        get_filtered_data = self._add_search_toolbar(parent, get_data, tree_ref, None)
        
        # Buttons toolbar
        btn_container = tk.Frame(toolbar, bg=self.bg_color)
        btn_container.pack(side=tk.RIGHT, padx=5)
        
        # Treeview - tạo trước để có thể dùng trong các nút
        tree_frame = tk.Frame(parent, bg=self.bg_color)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL)
        scrollbar_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', 
                           yscrollcommand=scrollbar_y.set,
                           xscrollcommand=scrollbar_x.set)
        
        # Cập nhật tree_ref
        tree_ref[0] = tree
        
        scrollbar_y.config(command=tree.yview)
        scrollbar_x.config(command=tree.xview)
        
        # Configure columns
        for col in columns:
            tree.heading(col, text=col)
            if col == 'STT':
                tree.column(col, width=50, anchor=tk.CENTER)
            elif col == 'Họ và Tên':
                tree.column(col, width=200, anchor=tk.W)
            elif col == 'Thông tin người thân':
                tree.column(col, width=300, anchor=tk.W)
            elif col == 'Thời gian vào' or col == 'Thời gian ra':
                tree.column(col, width=120, anchor=tk.CENTER)
            else:
                tree.column(col, width=120, anchor=tk.W)
        
        # Thêm border cho các hàng
        self._add_treeview_border(tree)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Nút Làm Mới
        def refresh_list():
            tree = tree_ref[0] if tree_ref else None
            if tree:
                for item in tree.get_children():
                    tree.delete(item)
                data = get_filtered_data()
                for item_data in data:
                    tree.insert('', tk.END, values=item_data['values'], tags=(item_data.get('id'),))
        
        # Nút Chọn quân nhân
        def choose_personnel():
            self.choose_bao_ve_an_ninh_personnel(parent)
        
        tk.Button(
            btn_container,
            text="👥 Chọn Quân Nhân",
            command=choose_personnel,
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Sửa Thời Gian
        def edit_time():
            tree = tree_ref[0] if tree_ref else None
            if not tree:
                return
            
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn quân nhân cần sửa!")
                return
            
            # Lấy personnel_id từ tags
            item = selected[0]
            tags = tree.item(item, 'tags')
            personnel_id = tags[0] if tags else None
            if not personnel_id:
                messagebox.showerror("Lỗi", "Không tìm thấy ID quân nhân!")
                return
            
            # Lấy thông tin quân nhân
            values = tree.item(item, 'values')
            ho_ten = values[1] if len(values) > 1 else ''
            
            # Lấy thông tin hiện tại
            bao_ve_info = self.db.get_bao_ve_an_ninh_info(personnel_id)
            thoi_gian_vao = bao_ve_info.get('thoiGianVao', '') or ''
            thoi_gian_ra = bao_ve_info.get('thoiGianRa', '') or ''
            
            # Mở dialog chỉnh sửa
            self.edit_bao_ve_an_ninh_time(parent, personnel_id, ho_ten, thoi_gian_vao, thoi_gian_ra, refresh_list)
        
        tk.Button(
            btn_container,
            text="✏️ Sửa Thời Gian",
            command=edit_time,
            font=('Segoe UI', 10),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xóa khỏi danh sách
        def remove_from_list():
            tree = tree_ref[0] if tree_ref else None
            if not tree:
                return
            
            selected = tree.selection()
            if not selected:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn quân nhân cần xóa!")
                return
            
            # Lấy personnel_id từ tags
            item = selected[0]
            tags = tree.item(item, 'tags')
            personnel_id = tags[0] if tags else None
            
            if not personnel_id:
                messagebox.showerror("Lỗi", "Không tìm thấy ID quân nhân!")
                return
            
            if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa quân nhân này khỏi danh sách?"):
                if self.db.remove_bao_ve_an_ninh(personnel_id):
                    messagebox.showinfo("Thành công", "Đã xóa quân nhân khỏi danh sách!")
                    refresh_list()
                else:
                    messagebox.showerror("Lỗi", "Không thể xóa quân nhân!")
        
        tk.Button(
            btn_container,
            text="❌ Xóa Khỏi Danh Sách",
            command=remove_from_list,
            font=('Segoe UI', 10),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        tk.Button(
            btn_container,
            text="🔄 Làm Mới",
            command=refresh_list,
            font=('Segoe UI', 10),
            bg='#388E3C',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Xuất Word
        def export_word():
            self.export_bao_ve_an_ninh_word(parent, get_filtered_data)
        
        tk.Button(
            btn_container,
            text="📄 Xuất Word",
            command=export_word,
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Bind double-click để chỉnh sửa thời gian vào/ra
        def on_double_click(event):
            item = tree.selection()[0] if tree.selection() else None
            if not item:
                return
            
            # Lấy personnel_id từ tags
            tags = tree.item(item, 'tags')
            personnel_id = tags[0] if tags else None
            if not personnel_id:
                return
            
            # Lấy thông tin quân nhân
            values = tree.item(item, 'values')
            ho_ten = values[1] if len(values) > 1 else ''
            
            # Lấy thông tin hiện tại
            bao_ve_info = self.db.get_bao_ve_an_ninh_info(personnel_id)
            thoi_gian_vao = bao_ve_info.get('thoiGianVao', '') or ''
            thoi_gian_ra = bao_ve_info.get('thoiGianRa', '') or ''
            
            # Mở dialog chỉnh sửa
            self.edit_bao_ve_an_ninh_time(parent, personnel_id, ho_ten, thoi_gian_vao, thoi_gian_ra, refresh_list)
        
        tree.bind('<Double-1>', on_double_click)
        
        # Load data
        refresh_list()
    
    def choose_bao_ve_an_ninh_personnel(self, parent):
        """Dialog chọn quân nhân vào bảo vệ an ninh"""
        dialog = tk.Toplevel(parent)
        dialog.title("Chọn Quân Nhân Bảo Vệ An Ninh")
        dialog.geometry("1100x750")
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(True, True)
        
        # Dùng grid để control layout tốt hơn
        dialog.grid_rowconfigure(0, weight=1)  # Row 0 (list_frame) có thể expand
        dialog.grid_rowconfigure(1, weight=0)  # Row 1 (time_container) không expand
        dialog.grid_rowconfigure(2, weight=0)  # Row 2 (btn_frame) không expand
        dialog.grid_columnconfigure(0, weight=1)
        
        # Frame chứa danh sách
        list_frame = tk.Frame(dialog, bg='#FAFAFA')
        list_frame.grid(row=0, column=0, padx=10, pady=10, sticky=tk.NSEW)
        
        # Label
        label = tk.Label(
            list_frame,
            text="Chọn quân nhân để thêm vào danh sách bảo vệ an ninh:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        )
        label.pack(anchor=tk.W, pady=5)
        
        # Toolbar với tìm kiếm và chọn tất cả
        toolbar_frame = tk.Frame(list_frame, bg='#FAFAFA')
        toolbar_frame.pack(fill=tk.X, pady=5)
        
        # Tìm kiếm
        tk.Label(toolbar_frame, text="Tìm kiếm:", font=('Segoe UI', 9), bg='#FAFAFA').pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(toolbar_frame, textvariable=search_var, width=30, font=('Segoe UI', 9))
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Nút chọn tất cả / Bỏ chọn tất cả
        select_all_var = tk.BooleanVar(value=False)
        def toggle_select_all():
            select_all_var.set(not select_all_var.get())
            for item in tree.get_children():
                item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
                if item_id:
                    if select_all_var.get():
                        tree.item(item, text='✓')
                        selected_ids.add(item_id)
                        if item_id not in time_data:
                            time_data[item_id] = {'vao': '', 'ra': ''}
                    else:
                        tree.item(item, text='')
                        selected_ids.discard(item_id)
            update_time_frame()
        
        select_all_btn = tk.Button(
            toolbar_frame,
            text="☑ Chọn Tất Cả",
            command=toggle_select_all,
            font=('Segoe UI', 9),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor='hand2'
        )
        select_all_btn.pack(side=tk.LEFT, padx=5)
        
        # Treeview với checkbox
        columns = ('Họ và Tên', 'Ngày Sinh', 'Cấp Bậc', 'Chức Vụ', 'Đơn Vị', 'Thông tin người thân')
        tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=15)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Configure columns
        tree.heading('#0', text='Chọn')
        tree.column('#0', width=50, anchor=tk.CENTER)
        for col in columns:
            tree.heading(col, text=col)
            if col == 'Họ và Tên':
                tree.column(col, width=180)
            elif col == 'Thông tin người thân':
                tree.column(col, width=300)
            else:
                tree.column(col, width=100)
        
        # Load data - tất cả quân nhân
        all_personnel = self.db.get_all()
        bao_ve_ids = set(self.db.get_bao_ve_an_ninh())
        
        selected_ids = set()
        
        def get_gia_dinh_info(p):
            """Lấy thông tin người thân từ database"""
            gia_dinh_info = []
            try:
                nguoi_than_list = self.db.get_nguoi_than_by_personnel(p.id)
                
                bo_de = []
                me_de = []
                vo = []
                
                for nguoi_than in nguoi_than_list:
                    moi_quan_he = (nguoi_than.moiQuanHe or '').lower().strip()
                    ho_ten = (nguoi_than.hoTen or '').strip()
                    ngay_sinh = (nguoi_than.ngaySinh or '').strip()
                    noi_dung = (nguoi_than.noiDung or '').strip()
                    
                    if not ho_ten:
                        continue
                    
                    # Lấy năm sinh
                    nam_sinh = ""
                    if ngay_sinh:
                        try:
                            if '/' in ngay_sinh:
                                parts = ngay_sinh.split('/')
                                nam_sinh = parts[-1] if len(parts) >= 3 else (parts[1] if len(parts) == 2 else parts[0][:4])
                            else:
                                nam_sinh = ngay_sinh[:4] if len(ngay_sinh) >= 4 else ngay_sinh
                        except:
                            nam_sinh = ""
                    
                    nghe = noi_dung if noi_dung else "làm nông"
                    
                    if nam_sinh:
                        info_str = f"{ho_ten} ({nam_sinh}-{nghe})"
                    else:
                        info_str = f"{ho_ten} ({nghe})"
                    
                    # Phân loại
                    if ('bố đẻ' in moi_quan_he or 'cha đẻ' in moi_quan_he or 
                        (('bố' in moi_quan_he or 'cha' in moi_quan_he) and 
                         'vợ' not in moi_quan_he and 'vo' not in moi_quan_he)):
                        bo_de.append(info_str)
                    elif ('mẹ đẻ' in moi_quan_he or 'me đẻ' in moi_quan_he or 
                          (('mẹ' in moi_quan_he or 'me' in moi_quan_he) and 
                           'vợ' not in moi_quan_he and 'vo' not in moi_quan_he)):
                        me_de.append(info_str)
                    elif 'vợ' in moi_quan_he or 'vo' in moi_quan_he:
                        vo.append(info_str)
                    elif 'bố' in moi_quan_he or 'cha' in moi_quan_he:
                        if 'vợ' not in moi_quan_he and 'vo' not in moi_quan_he:
                            bo_de.append(info_str)
                    elif 'mẹ' in moi_quan_he or 'me' in moi_quan_he:
                        if 'vợ' not in moi_quan_he and 'vo' not in moi_quan_he:
                            me_de.append(info_str)
                
                gia_dinh_info.extend(bo_de)
                gia_dinh_info.extend(me_de)
                gia_dinh_info.extend(vo)
                
            except Exception:
                # Fallback: sử dụng các field cũ
                if p.hoTenCha:
                    nam_sinh_cha = ""
                    if p.ngaySinhCha:
                        try:
                            if '/' in p.ngaySinhCha:
                                parts = p.ngaySinhCha.split('/')
                                nam_sinh_cha = parts[-1] if len(parts) >= 3 else (parts[1] if len(parts) == 2 else parts[0][:4])
                            else:
                                nam_sinh_cha = p.ngaySinhCha[:4] if len(p.ngaySinhCha) >= 4 else p.ngaySinhCha
                        except:
                            pass
                    gia_dinh_info.append(f"{p.hoTenCha} ({nam_sinh_cha}-làm nông)" if nam_sinh_cha else f"{p.hoTenCha} (làm nông)")
                if p.hoTenMe:
                    nam_sinh_me = ""
                    if p.ngaySinhMe:
                        try:
                            if '/' in p.ngaySinhMe:
                                parts = p.ngaySinhMe.split('/')
                                nam_sinh_me = parts[-1] if len(parts) >= 3 else (parts[1] if len(parts) == 2 else parts[0][:4])
                            else:
                                nam_sinh_me = p.ngaySinhMe[:4] if len(p.ngaySinhMe) >= 4 else p.ngaySinhMe
                        except:
                            pass
                    gia_dinh_info.append(f"{p.hoTenMe} ({nam_sinh_me}-làm nông)" if nam_sinh_me else f"{p.hoTenMe} (làm nông)")
                if p.hoTenVo:
                    nam_sinh_vo = ""
                    if p.ngaySinhVo:
                        try:
                            if '/' in p.ngaySinhVo:
                                parts = p.ngaySinhVo.split('/')
                                nam_sinh_vo = parts[-1] if len(parts) >= 3 else (parts[1] if len(parts) == 2 else parts[0][:4])
                            else:
                                nam_sinh_vo = p.ngaySinhVo[:4] if len(p.ngaySinhVo) >= 4 else p.ngaySinhVo
                        except:
                            pass
                    gia_dinh_info.append(f"{p.hoTenVo} ({nam_sinh_vo}-GV)" if nam_sinh_vo else f"{p.hoTenVo} (GV)")
            
            return " / ".join(gia_dinh_info) if gia_dinh_info else ''
        
        def filter_tree():
            """Lọc danh sách theo từ khóa tìm kiếm"""
            search_text = search_var.get().lower().strip()
            
            # Xóa tất cả items
            for item in tree.get_children():
                tree.delete(item)
            
            # Thêm lại items đã lọc
            for p in all_personnel:
                # Lọc theo từ khóa
                if search_text:
                    searchable_text = f"{p.hoTen or ''} {p.ngaySinh or ''} {p.capBac or ''} {p.chucVu or ''} {p.donVi or ''}".lower()
                    if search_text not in searchable_text:
                        continue
                
                is_selected = p.id in bao_ve_ids
                if is_selected:
                    selected_ids.add(p.id)
                
                gia_dinh_info = get_gia_dinh_info(p)
                
                tree.insert('', tk.END, 
                           text='✓' if is_selected else '',
                           values=(
                               p.hoTen or '',
                               p.ngaySinh or '',
                               p.capBac or '',
                               p.chucVu or '',
                               p.donVi or '',
                               gia_dinh_info
                           ),
                           tags=(p.id,))
        
        # Bind tìm kiếm
        search_var.trace('w', lambda *args: filter_tree())
        
        # Load dữ liệu ban đầu
        filter_tree()
        
        # Dictionary để lưu thời gian vào/ra cho từng quân nhân
        time_data = {}  # {personnel_id: {'vao': '', 'ra': ''}}
        
        # Load thời gian hiện tại cho các quân nhân đã có
        for p in all_personnel:
            if p.id in bao_ve_ids:
                bao_ve_info = self.db.get_bao_ve_an_ninh_info(p.id)
                time_data[p.id] = {
                    'vao': bao_ve_info.get('thoiGianVao', '') or '',
                    'ra': bao_ve_info.get('thoiGianRa', '') or ''
                }
        
        # Bind click để toggle
        def toggle_selection(event):
            item = tree.selection()[0] if tree.selection() else None
            if not item:
                return
            
            item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
            if not item_id:
                return
            
            current_text = tree.item(item, 'text')
            if current_text == '✓':
                tree.item(item, text='')
                selected_ids.discard(item_id)
            else:
                tree.item(item, text='✓')
                selected_ids.add(item_id)
                # Khởi tạo thời gian nếu chưa có
                if item_id not in time_data:
                    time_data[item_id] = {'vao': '', 'ra': ''}
            
            # Cập nhật frame nhập thời gian
            update_time_frame()
        
        def update_time_frame():
            # Xóa frame cũ
            for widget in time_container.winfo_children():
                widget.destroy()
            
            # Hiển thị frame nhập thời gian cho quân nhân đã chọn
            if selected_ids:
                tk.Label(
                    time_container,
                    text="Nhập thời gian vào/ra cho từng quân nhân:",
                    font=('Segoe UI', 10, 'bold'),
                    bg='#FAFAFA'
                ).pack(anchor=tk.W, pady=5)
                
                # Tạo frame scrollable cho danh sách quân nhân đã chọn
                scroll_frame = tk.Frame(time_container, bg='#FAFAFA')
                scroll_frame.pack(fill=tk.BOTH, expand=True, pady=5)
                
                # Canvas và scrollbar cho danh sách
                canvas = tk.Canvas(scroll_frame, bg='#FAFAFA', height=150)
                scrollbar_time = ttk.Scrollbar(scroll_frame, orient=tk.VERTICAL, command=canvas.yview)
                scrollable_frame = tk.Frame(canvas, bg='#FAFAFA')
                
                scrollable_frame.bind(
                    "<Configure>",
                    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
                )
                
                canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar_time.set)
                
                # Thêm Entry cho mỗi quân nhân đã chọn
                for personnel_id in selected_ids:
                    p = next((p for p in all_personnel if p.id == personnel_id), None)
                    if not p:
                        continue
                    
                    # Khởi tạo nếu chưa có
                    if personnel_id not in time_data:
                        time_data[personnel_id] = {'vao': '', 'ra': ''}
                    
                    person_frame = tk.Frame(scrollable_frame, bg='#FAFAFA')
                    person_frame.pack(fill=tk.X, pady=2)
                    
                    tk.Label(
                        person_frame,
                        text=f"{p.hoTen or 'N/A'}:",
                        font=('Segoe UI', 9),
                        bg='#FAFAFA',
                        width=20,
                        anchor=tk.W
                    ).pack(side=tk.LEFT, padx=5)
                    
                    tk.Label(
                        person_frame,
                        text="Vào (MM/YYYY):",
                        font=('Segoe UI', 9),
                        bg='#FAFAFA'
                    ).pack(side=tk.LEFT, padx=2)
                    
                    vao_var = tk.StringVar(value=time_data[personnel_id]['vao'])
                    vao_entry = tk.Entry(person_frame, textvariable=vao_var, width=15, font=('Segoe UI', 9))
                    vao_entry.pack(side=tk.LEFT, padx=2)
                    if not time_data[personnel_id]['vao']:
                        vao_entry.insert(0, "MM/YYYY")
                        vao_entry.config(fg='gray')
                        def on_vao_focus_in(e):
                            if vao_entry.get() == "MM/YYYY":
                                vao_entry.delete(0, tk.END)
                                vao_entry.config(fg='black')
                        def on_vao_focus_out(e):
                            if not vao_entry.get():
                                vao_entry.insert(0, "MM/YYYY")
                                vao_entry.config(fg='gray')
                        vao_entry.bind('<FocusIn>', on_vao_focus_in)
                        vao_entry.bind('<FocusOut>', on_vao_focus_out)
                    vao_entry.bind('<KeyRelease>', lambda e, pid=personnel_id, var=vao_var: time_data.update({pid: {**time_data.get(pid, {}), 'vao': var.get() if var.get() != "MM/YYYY" else ''}}))
                    
                    tk.Label(
                        person_frame,
                        text="Ra (MM/YYYY):",
                        font=('Segoe UI', 9),
                        bg='#FAFAFA'
                    ).pack(side=tk.LEFT, padx=2)
                    
                    ra_var = tk.StringVar(value=time_data[personnel_id]['ra'])
                    ra_entry = tk.Entry(person_frame, textvariable=ra_var, width=15, font=('Segoe UI', 9))
                    ra_entry.pack(side=tk.LEFT, padx=2)
                    if not time_data[personnel_id]['ra']:
                        ra_entry.insert(0, "MM/YYYY")
                        ra_entry.config(fg='gray')
                        def on_ra_focus_in(e):
                            if ra_entry.get() == "MM/YYYY":
                                ra_entry.delete(0, tk.END)
                                ra_entry.config(fg='black')
                        def on_ra_focus_out(e):
                            if not ra_entry.get():
                                ra_entry.insert(0, "MM/YYYY")
                                ra_entry.config(fg='gray')
                        ra_entry.bind('<FocusIn>', on_ra_focus_in)
                        ra_entry.bind('<FocusOut>', on_ra_focus_out)
                    ra_entry.bind('<KeyRelease>', lambda e, pid=personnel_id, var=ra_var: time_data.update({pid: {**time_data.get(pid, {}), 'ra': var.get() if var.get() != "MM/YYYY" else ''}}))
                
                canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                scrollbar_time.pack(side=tk.RIGHT, fill=tk.Y)
        
        tree.bind('<Button-1>', toggle_selection)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Frame container cho thời gian vào/ra - Row 1
        time_container = tk.Frame(dialog, bg='#FAFAFA', height=200)
        time_container.grid(row=1, column=0, padx=10, pady=5, sticky=tk.EW)
        time_container.grid_propagate(False)
        
        # Cập nhật frame thời gian ban đầu
        update_time_frame()
        
        # Buttons - Row 2, LUÔN HIỂN THỊ
        btn_frame = tk.Frame(dialog, bg='#FAFAFA', height=70)
        btn_frame.grid(row=2, column=0, padx=10, pady=10, sticky=tk.EW)
        btn_frame.grid_propagate(False)
        btn_frame.grid_columnconfigure(1, weight=1)
        
        def save_selection():
            # Lưu tất cả quân nhân đã chọn với thời gian riêng của từng người
            success_count = 0
            for personnel_id in selected_ids:
                thoi_gian_vao = time_data.get(personnel_id, {}).get('vao', '').strip()
                thoi_gian_ra = time_data.get(personnel_id, {}).get('ra', '').strip()
                if self.db.add_bao_ve_an_ninh(personnel_id, thoi_gian_vao, thoi_gian_ra):
                    success_count += 1
            
            # Xóa những quân nhân không được chọn
            all_bao_ve_ids = set(self.db.get_bao_ve_an_ninh())
            to_remove = all_bao_ve_ids - selected_ids
            for personnel_id in to_remove:
                self.db.remove_bao_ve_an_ninh(personnel_id)
            
            messagebox.showinfo("Thành công", f"Đã cập nhật {success_count} quân nhân vào danh sách!")
            dialog.destroy()
            # Refresh lại tab
            self.create_bao_ve_an_ninh_tab(parent)
        
        # Nút Hủy
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor='hand2',
            width=10
        ).grid(row=0, column=0, padx=5, sticky=tk.W)
        
        # Spacer
        tk.Frame(btn_frame, bg='#FAFAFA').grid(row=0, column=1, sticky=tk.EW)
        
        # Nút XONG
        tk.Button(
            btn_frame,
            text="✅ XONG",
            command=save_selection,
            font=('Segoe UI', 11, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.RAISED,
            padx=25,
            pady=8,
            cursor='hand2',
            width=12,
            bd=2
        ).grid(row=0, column=2, padx=5, sticky=tk.E)
    
    def edit_bao_ve_an_ninh_time(self, parent, personnel_id, ho_ten, thoi_gian_vao, thoi_gian_ra, refresh_callback):
        """Dialog chỉnh sửa thời gian vào/ra cho quân nhân"""
        dialog = tk.Toplevel(parent)
        dialog.title("Chỉnh Sửa Thời Gian")
        dialog.geometry("450x250")
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Frame chứa form
        form_frame = tk.Frame(dialog, bg='#FAFAFA', padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tiêu đề
        title_label = tk.Label(
            form_frame,
            text=f"Chỉnh sửa thời gian cho: {ho_ten}",
            font=('Segoe UI', 11, 'bold'),
            bg='#FAFAFA',
            fg='#388E3C'
        )
        title_label.pack(pady=(0, 20))
        
        # Thời gian vào
        time_frame1 = tk.Frame(form_frame, bg='#FAFAFA')
        time_frame1.pack(fill=tk.X, pady=10)
        
        tk.Label(
            time_frame1,
            text="Thời gian vào (MM/YYYY):",
            font=('Segoe UI', 10),
            bg='#FAFAFA',
            width=20
        ).pack(side=tk.LEFT, padx=5)
        
        thoi_gian_vao_var = tk.StringVar(value=thoi_gian_vao)
        thoi_gian_vao_entry = tk.Entry(time_frame1, textvariable=thoi_gian_vao_var, width=15, font=('Segoe UI', 10))
        thoi_gian_vao_entry.pack(side=tk.LEFT, padx=5)
        thoi_gian_vao_entry.focus()
        
        # Thời gian ra
        time_frame2 = tk.Frame(form_frame, bg='#FAFAFA')
        time_frame2.pack(fill=tk.X, pady=10)
        
        tk.Label(
            time_frame2,
            text="Thời gian ra (MM/YYYY):",
            font=('Segoe UI', 10),
            bg='#FAFAFA',
            width=20
        ).pack(side=tk.LEFT, padx=5)
        
        thoi_gian_ra_var = tk.StringVar(value=thoi_gian_ra)
        thoi_gian_ra_entry = tk.Entry(time_frame2, textvariable=thoi_gian_ra_var, width=15, font=('Segoe UI', 10))
        thoi_gian_ra_entry.pack(side=tk.LEFT, padx=5)
        
        # Buttons
        btn_frame = tk.Frame(form_frame, bg='#FAFAFA')
        btn_frame.pack(fill=tk.X, pady=(30, 0), side=tk.BOTTOM)
        
        def save_time():
            thoi_gian_vao_new = thoi_gian_vao_var.get().strip()
            thoi_gian_ra_new = thoi_gian_ra_var.get().strip()
            
            if self.db.add_bao_ve_an_ninh(personnel_id, thoi_gian_vao_new, thoi_gian_ra_new):
                messagebox.showinfo("Thành công", "Đã cập nhật thời gian!")
                dialog.destroy()
                if refresh_callback:
                    refresh_callback()
            else:
                messagebox.showerror("Lỗi", "Không thể cập nhật thời gian!")
        
        # Nút Lưu - màu xanh lá, nổi bật
        save_btn = tk.Button(
            btn_frame,
            text="💾 Lưu",
            command=save_time,
            font=('Segoe UI', 11, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=30,
            pady=8,
            cursor='hand2',
            width=10
        )
        save_btn.pack(side=tk.RIGHT, padx=10)
        
        # Nút Hủy
        cancel_btn = tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 11),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=30,
            pady=8,
            cursor='hand2',
            width=10
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        # Bind Enter key để lưu
        thoi_gian_vao_entry.bind('<Return>', lambda e: save_time())
        thoi_gian_ra_entry.bind('<Return>', lambda e: save_time())
        
        # Focus vào ô đầu tiên
        thoi_gian_vao_entry.focus_set()
    
    def export_bao_ve_an_ninh_word(self, parent, get_data_func):
        """Dialog xuất Word cho Bảo Vệ An Ninh"""
        dialog = tk.Toplevel(parent)
        dialog.title("Xuất File Word - Bảo Vệ An Ninh")
        dialog.geometry("550x650")
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Frame chứa form với scrollbar
        canvas = tk.Canvas(dialog, bg='#FAFAFA')
        scrollbar = ttk.Scrollbar(dialog, orient=tk.VERTICAL, command=canvas.yview)
        form_frame = tk.Frame(canvas, bg='#FAFAFA', padx=20, pady=20)
        
        form_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=form_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Tiêu đề
        title_label = tk.Label(
            form_frame,
            text="Thiết lập thông tin xuất file",
            font=('Segoe UI', 12, 'bold'),
            bg='#FAFAFA',
            fg='#388E3C'
        )
        title_label.pack(pady=(0, 20))
        
        # Đơn vị
        tk.Label(
            form_frame,
            text="Đơn vị:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        don_vi_var = tk.StringVar(value="Đại đội 3")
        don_vi_entry = tk.Entry(form_frame, textvariable=don_vi_var, width=40, font=('Segoe UI', 10))
        don_vi_entry.pack(fill=tk.X, pady=5)
        
        # Tiểu đoàn
        tk.Label(
            form_frame,
            text="Tiểu đoàn:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
        tieu_doan_entry = tk.Entry(form_frame, textvariable=tieu_doan_var, width=40, font=('Segoe UI', 10))
        tieu_doan_entry.pack(fill=tk.X, pady=5)
        
        # Địa điểm
        tk.Label(
            form_frame,
            text="Địa điểm:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        dia_diem_var = tk.StringVar(value="Đắk Lắk")
        dia_diem_entry = tk.Entry(form_frame, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10))
        dia_diem_entry.pack(fill=tk.X, pady=5)
        
        # Năm
        tk.Label(
            form_frame,
            text="Năm:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        nam_var = tk.StringVar(value="2025")
        nam_entry = tk.Entry(form_frame, textvariable=nam_var, width=40, font=('Segoe UI', 10))
        nam_entry.pack(fill=tk.X, pady=5)
        
        # Ngày bổ sung
        tk.Label(
            form_frame,
            text="Ngày bổ sung:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        ngay_bo_sung_var = tk.StringVar(value="01")
        ngay_bo_sung_entry = tk.Entry(form_frame, textvariable=ngay_bo_sung_var, width=40, font=('Segoe UI', 10))
        ngay_bo_sung_entry.pack(fill=tk.X, pady=5)
        
        # Tháng bổ sung
        tk.Label(
            form_frame,
            text="Tháng bổ sung:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        thang_bo_sung_var = tk.StringVar(value="7")
        thang_bo_sung_entry = tk.Entry(form_frame, textvariable=thang_bo_sung_var, width=40, font=('Segoe UI', 10))
        thang_bo_sung_entry.pack(fill=tk.X, pady=5)
        
        # Năm bổ sung
        tk.Label(
            form_frame,
            text="Năm bổ sung:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        nam_bo_sung_var = tk.StringVar(value="2025")
        nam_bo_sung_entry = tk.Entry(form_frame, textvariable=nam_bo_sung_var, width=40, font=('Segoe UI', 10))
        nam_bo_sung_entry.pack(fill=tk.X, pady=5)
        
        # Chính trị viên
        tk.Label(
            form_frame,
            text="Chính trị viên:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        ).pack(anchor=tk.W, pady=5)
        
        chinh_tri_vien_var = tk.StringVar(value="Đại úy Triệu Văn Dũng")
        chinh_tri_vien_entry = tk.Entry(form_frame, textvariable=chinh_tri_vien_var, width=40, font=('Segoe UI', 10))
        chinh_tri_vien_entry.pack(fill=tk.X, pady=5)
        
        # Buttons - đặt ở cuối form_frame
        btn_frame = tk.Frame(form_frame, bg='#FAFAFA')
        btn_frame.pack(fill=tk.X, pady=(30, 10))
        
        def save_and_export():
            # Lấy dữ liệu
            don_vi = don_vi_var.get().strip()
            tieu_doan = tieu_doan_var.get().strip()
            dia_diem = dia_diem_var.get().strip()
            nam = nam_var.get().strip()
            ngay_bo_sung = ngay_bo_sung_var.get().strip()
            thang_bo_sung = thang_bo_sung_var.get().strip()
            nam_bo_sung = nam_bo_sung_var.get().strip()
            chinh_tri_vien = chinh_tri_vien_var.get().strip()
            
            if not don_vi:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập đơn vị!")
                return
            
            # Lấy danh sách quân nhân
            data = get_data_func()
            if not data:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            # Lấy danh sách Personnel từ IDs
            bao_ve_ids = self.db.get_bao_ve_an_ninh()
            all_personnel = self.db.get_all()
            personnel_list = [p for p in all_personnel if p.id in bao_ve_ids]
            
            if not personnel_list:
                messagebox.showwarning("Cảnh báo", "Không có quân nhân trong danh sách!")
                return
            
            # Chọn file để lưu
            filename = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                initialfile=f"Bao_Ve_An_Ninh_{don_vi.replace(' ', '_')}.docx"
            )
            
            if not filename:
                return
            
            try:
                # Import và gọi hàm xuất
                from services.export_bao_ve_an_ninh import to_word_docx_bao_ve_an_ninh
                
                word_content = to_word_docx_bao_ve_an_ninh(
                    personnel_list=personnel_list,
                    don_vi=don_vi,
                    tieu_doan=tieu_doan,
                    dia_diem=dia_diem,
                    nam=nam,
                    ngay_bo_sung=ngay_bo_sung,
                    thang_bo_sung=thang_bo_sung,
                    nam_bo_sung=nam_bo_sung,
                    chinh_tri_vien=chinh_tri_vien,
                    db_service=self.db
                )
                
                # Lưu file
                with open(filename, 'wb') as f:
                    f.write(word_content)
                
                messagebox.showinfo("Thành công", f"Đã xuất file Word thành công!\n{filename}")
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        
        # Nút Xuất File - màu xanh lá, nổi bật
        export_btn = tk.Button(
            btn_frame,
            text="📄 Xuất File",
            command=save_and_export,
            font=('Segoe UI', 11, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=30,
            pady=8,
            cursor='hand2',
            width=12
        )
        export_btn.pack(side=tk.RIGHT, padx=10)
        
        # Nút Hủy
        cancel_btn = tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 11),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=30,
            pady=8,
            cursor='hand2',
            width=12
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
    
    def choose_nguoi_than_che_do_cu_personnel(self, parent):
        """Dialog chọn quân nhân có người thân tham gia chế độ cũ"""
        dialog = tk.Toplevel(parent)
        dialog.title("Chọn Quân Nhân Có Người Thân Tham Gia Chế Độ Cũ")
        dialog.geometry("1100x700")
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(True, True)
        
        # Frame chứa danh sách
        list_frame = tk.Frame(dialog, bg='#FAFAFA')
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Label
        label = tk.Label(
            list_frame,
            text="Chọn quân nhân có người thân tham gia chế độ cũ:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        )
        label.pack(anchor=tk.W, pady=5)
        
        # Toolbar với tìm kiếm và chọn tất cả
        toolbar_frame = tk.Frame(list_frame, bg='#FAFAFA')
        toolbar_frame.pack(fill=tk.X, pady=5)
        
        # Tìm kiếm
        tk.Label(toolbar_frame, text="Tìm kiếm:", font=('Segoe UI', 9), bg='#FAFAFA').pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(toolbar_frame, textvariable=search_var, width=30, font=('Segoe UI', 9))
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Nút chọn tất cả / Bỏ chọn tất cả
        select_all_var = tk.BooleanVar(value=False)
        selected_ids = set(self.db.get_nguoi_than_che_do_cu())
        
        def toggle_select_all():
            select_all_var.set(not select_all_var.get())
            for item in tree.get_children():
                item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
                if item_id:
                    if select_all_var.get():
                        tree.item(item, text='✓')
                        selected_ids.add(item_id)
                    else:
                        tree.item(item, text='')
                        selected_ids.discard(item_id)
        
        select_all_btn = tk.Button(
            toolbar_frame,
            text="☑ Chọn Tất Cả",
            command=toggle_select_all,
            font=('Segoe UI', 9),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor='hand2'
        )
        select_all_btn.pack(side=tk.LEFT, padx=5)
        
        # Treeview với checkbox
        columns = ('Họ và Tên', 'Ngày Sinh', 'Cấp Bậc', 'Chức Vụ', 'Đơn Vị', 'Có người thân')
        tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=20)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Configure columns
        tree.heading('#0', text='Chọn')
        tree.column('#0', width=50, anchor=tk.CENTER)
        for col in columns:
            tree.heading(col, text=col)
            if col == 'Họ và Tên':
                tree.column(col, width=200)
            elif col == 'Có người thân':
                tree.column(col, width=150)
            else:
                tree.column(col, width=120)
        
        # Load data - chỉ hiển thị quân nhân có người thân
        all_personnel = self.db.get_all()
        
        def has_nguoi_than(p):
            """Kiểm tra quân nhân có người thân không"""
            try:
                nguoi_than_list = self.db.get_nguoi_than_by_personnel(p.id)
                return len(nguoi_than_list) > 0
            except:
                return False
        
        # Chỉ hiển thị quân nhân có người thân
        filtered_personnel = [p for p in all_personnel if has_nguoi_than(p)]
        
        def load_tree_data():
            """Load dữ liệu vào tree"""
            # Xóa dữ liệu cũ
            for item in tree.get_children():
                tree.delete(item)
            
            # Lọc theo tìm kiếm
            search_text = search_var.get().lower()
            display_personnel = filtered_personnel
            if search_text:
                display_personnel = [p for p in filtered_personnel 
                                  if search_text in (p.hoTen or '').lower()]
            
            for person in display_personnel:
                is_selected = person.id in selected_ids
                item_text = '✓' if is_selected else ''
                
                # Đếm số người thân
                try:
                    nguoi_than_list = self.db.get_nguoi_than_by_personnel(person.id)
                    nguoi_than_count = len(nguoi_than_list)
                except:
                    nguoi_than_count = 0
                
                item = tree.insert('', 'end', 
                                  text=item_text,
                                  tags=(person.id,),
                                  values=(
                                      person.hoTen or '',
                                      person.ngaySinh or '',
                                      person.capBac or '',
                                      person.chucVu or '',
                                      person.donVi or '',
                                      f"{nguoi_than_count} người"
                                  ))
        
        def on_item_click(event):
            """Xử lý click vào item"""
            item = tree.identify_row(event.y)
            if item:
                # Set selection để highlight row
                tree.selection_set(item)
                item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
                if item_id:
                    current_text = tree.item(item, 'text')
                    if current_text == '✓':
                        tree.item(item, text='')
                        selected_ids.discard(item_id)
                    else:
                        tree.item(item, text='✓')
                        selected_ids.add(item_id)
        
        tree.bind('<Button-1>', on_item_click)
        search_var.trace('w', lambda *args: load_tree_data())
        
        # Pack tree và scrollbar
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Load dữ liệu ban đầu
        load_tree_data()
        
        # Buttons - Row 1, LUÔN HIỂN THỊ
        btn_frame = tk.Frame(dialog, bg='#FAFAFA', height=70)
        btn_frame.grid(row=1, column=0, padx=10, pady=10, sticky=tk.EW)
        btn_frame.grid_propagate(False)
        btn_frame.grid_columnconfigure(1, weight=1)
        
        def save_selection():
            """Lưu danh sách đã chọn"""
            # Kiểm tra lại selected_ids từ tree để đảm bảo đồng bộ
            current_selected = set()
            for item in tree.get_children():
                item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
                if item_id and tree.item(item, 'text') == '✓':
                    current_selected.add(item_id)
            
            # Cập nhật selected_ids với dữ liệu từ tree
            selected_ids.clear()
            selected_ids.update(current_selected)
            
            if not selected_ids:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất một quân nhân!")
                return
            
            success_count = 0
            for personnel_id in selected_ids:
                if self.db.add_nguoi_than_che_do_cu(personnel_id):
                    success_count += 1
            
            # Xóa những quân nhân không được chọn
            all_selected_ids = set(self.db.get_nguoi_than_che_do_cu())
            to_remove = all_selected_ids - selected_ids
            for personnel_id in to_remove:
                self.db.remove_nguoi_than_che_do_cu(personnel_id)
            
            messagebox.showinfo("Thành công", f"Đã cập nhật {success_count} quân nhân vào danh sách!")
            dialog.destroy()
            # Refresh lại tab
            self.create_to_3_nguoi_tab(parent)
        
        # Nút Hủy
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor='hand2',
            width=10
        ).grid(row=0, column=0, padx=5, sticky=tk.W)
        
        # Spacer
        tk.Frame(btn_frame, bg='#FAFAFA').grid(row=0, column=1, sticky=tk.EW)
        
        # Nút XONG
        tk.Button(
            btn_frame,
            text="✅ XONG",
            command=save_selection,
            font=('Segoe UI', 11, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.RAISED,
            padx=25,
            pady=8,
            cursor='hand2',
            width=12,
            bd=2
        ).grid(row=0, column=2, padx=5, sticky=tk.E)
    
    def export_nguoi_than_che_do_cu_word(self, get_data_func):
        """Xuất danh sách quân nhân có người thân tham gia chế độ cũ ra Word"""
        try:
            from services.export_nguoi_than_che_do_cu import to_word_docx_nguoi_than_che_do_cu
            
            # Lấy dữ liệu
            data_list = get_data_func()
            if not data_list:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            # Lấy danh sách quân nhân từ IDs
            personnel_ids = [item['id'] for item in data_list]
            personnel_list = [self.db.get_by_id(pid) for pid in personnel_ids if self.db.get_by_id(pid)]
            
            # Mở dialog nhập thông tin
            dialog = tk.Toplevel(self)
            dialog.title("Xuất File Word")
            dialog.geometry("500x300")
            dialog.transient(self)
            dialog.grab_set()
            
            form_frame = tk.Frame(dialog, bg='#FAFAFA', padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH, expand=True)
            
            # Tiểu đoàn
            tk.Label(form_frame, text="Tiểu đoàn:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
            tk.Entry(form_frame, textvariable=tieu_doan_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Đại đội
            tk.Label(form_frame, text="Đại đội:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dai_doi_var = tk.StringVar(value="ĐẠI ĐỘI 3")
            tk.Entry(form_frame, textvariable=dai_doi_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Địa điểm
            tk.Label(form_frame, text="Địa điểm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dia_diem_var = tk.StringVar(value="Đắk Lắk")
            tk.Entry(form_frame, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Chính trị viên
            tk.Label(form_frame, text="Chính trị viên:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            chinh_tri_vien_var = tk.StringVar(value="Đại úy Triệu Văn Dũng")
            tk.Entry(form_frame, textvariable=chinh_tri_vien_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            def save_and_export():
                word_bytes = to_word_docx_nguoi_than_che_do_cu(
                    personnel_list,
                    tieu_doan=tieu_doan_var.get(),
                    dai_doi=dai_doi_var.get(),
                    dia_diem=dia_diem_var.get(),
                    chinh_tri_vien=chinh_tri_vien_var.get(),
                    db_service=self.db
                )
                
                filename = filedialog.asksaveasfilename(
                    defaultextension=".docx",
                    filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                    title="Lưu file Word"
                )
                
                if filename:
                    with open(filename, 'wb') as f:
                        f.write(word_bytes)
                    messagebox.showinfo("Thành công", f"Đã xuất file Word:\n{filename}")
                    dialog.destroy()
            
            btn_frame = tk.Frame(dialog, bg='#FAFAFA', height=60)
            btn_frame.pack(fill=tk.X, padx=10, pady=10)
            btn_frame.pack_propagate(False)
            
            tk.Button(
                btn_frame,
                text="📄 Xuất File",
                command=save_and_export,
                font=('Segoe UI', 11, 'bold'),
                bg='#4CAF50',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=10)
            
            tk.Button(
                btn_frame,
                text="❌ Hủy",
                command=dialog.destroy,
                font=('Segoe UI', 11),
                bg='#757575',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=5)
            
        except ImportError as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
    
    def export_dang_phai_phan_dong_word(self, get_data_func):
        """Xuất danh sách quân nhân có người thân tham gia đảng phái phản động ra Word"""
        try:
            from services.export_dang_phai_phan_dong import to_word_docx_dang_phai_phan_dong
            
            # Lấy dữ liệu
            data_list = get_data_func()
            if not data_list:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            # Lấy danh sách quân nhân từ IDs
            personnel_ids = [item['id'] for item in data_list]
            personnel_list = [self.db.get_by_id(pid) for pid in personnel_ids if self.db.get_by_id(pid)]
            
            # Mở dialog nhập thông tin
            dialog = tk.Toplevel(self)
            dialog.title("Xuất File Word - Người Thân Đảng Phái Phản Động")
            dialog.geometry("500x400")
            dialog.transient(self)
            dialog.grab_set()
            
            main_container = tk.Frame(dialog, bg='#FAFAFA')
            main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            # Tiểu đoàn
            tk.Label(main_container, text="Tiểu đoàn:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
            tk.Entry(main_container, textvariable=tieu_doan_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Đại đội
            tk.Label(main_container, text="Đại đội:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dai_doi_var = tk.StringVar(value="ĐẠI ĐỘI 3")
            tk.Entry(main_container, textvariable=dai_doi_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Địa điểm
            tk.Label(main_container, text="Địa điểm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dia_diem_var = tk.StringVar(value="Đắk Lắk")
            tk.Entry(main_container, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Năm
            tk.Label(main_container, text="Năm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            nam_var = tk.StringVar(value="2025")
            tk.Entry(main_container, textvariable=nam_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Chính trị viên
            tk.Label(main_container, text="Chính trị viên:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            chinh_tri_vien_var = tk.StringVar(value="Đại úy Triệu Văn Dũng")
            tk.Entry(main_container, textvariable=chinh_tri_vien_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            def save_and_export():
                try:
                    word_bytes = to_word_docx_dang_phai_phan_dong(
                        personnel_list=personnel_list,
                        tieu_doan=tieu_doan_var.get(),
                        dai_doi=dai_doi_var.get(),
                        dia_diem=dia_diem_var.get(),
                        nam=nam_var.get(),
                        chinh_tri_vien=chinh_tri_vien_var.get(),
                        db_service=self.db
                    )
                    
                    filename = filedialog.asksaveasfilename(
                        defaultextension=".docx",
                        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                        title="Lưu file Word"
                    )
                    
                    if filename:
                        with open(filename, 'wb') as f:
                            f.write(word_bytes)
                        messagebox.showinfo("Thành công", f"Đã xuất file Word:\n{filename}")
                        dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
            
            btn_frame = tk.Frame(main_container, bg='#FAFAFA', height=70)
            btn_frame.pack(fill=tk.X, padx=10, pady=10, side=tk.BOTTOM)
            btn_frame.pack_propagate(False)
            
            tk.Button(
                btn_frame,
                text="📄 Xuất File",
                command=save_and_export,
                font=('Segoe UI', 11, 'bold'),
                bg='#4CAF50',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=10)
            
            tk.Button(
                btn_frame,
                text="❌ Hủy",
                command=dialog.destroy,
                font=('Segoe UI', 11),
                bg='#757575',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=5)
            
        except ImportError as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
    
    def export_to_dan_van_word(self, get_data_func):
        """Xuất danh sách Tổ công tác dân vận ra Word"""
        try:
            from services.export_to_dan_van import to_word_docx_to_dan_van
            from tkinter import filedialog
            
            # Lấy dữ liệu
            data_list = get_data_func()
            if not data_list:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            # Lấy danh sách quân nhân từ IDs
            personnel_ids = [item['id'] for item in data_list]
            personnel_list = [self.db.get_by_id(pid) for pid in personnel_ids if self.db.get_by_id(pid)]
            
            # Mở dialog nhập thông tin
            dialog = tk.Toplevel(self)
            dialog.title("Xuất File Word - Tổ Công Tác Dân Vận")
            dialog.geometry("500x400")
            dialog.transient(self)
            dialog.grab_set()
            dialog.resizable(False, False)
            
            # Main container để chứa form và buttons
            main_container = tk.Frame(dialog, bg='#FAFAFA')
            main_container.pack(fill=tk.BOTH, expand=True)
            
            # Frame chứa form
            form_frame = tk.Frame(main_container, bg='#FAFAFA', padx=20, pady=20)
            form_frame.pack(fill=tk.BOTH, expand=True)
            
            # Tiểu đoàn
            tk.Label(form_frame, text="Tiểu đoàn:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
            tk.Entry(form_frame, textvariable=tieu_doan_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Đại đội
            tk.Label(form_frame, text="Đại đội:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dai_doi_var = tk.StringVar(value="ĐẠI ĐỘI 3")
            tk.Entry(form_frame, textvariable=dai_doi_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Địa điểm
            tk.Label(form_frame, text="Địa điểm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dia_diem_var = tk.StringVar(value="Đắk Lắk")
            tk.Entry(form_frame, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Ngày tháng năm
            tk.Label(form_frame, text="Ngày tháng năm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            ngay_thang_nam_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
            tk.Entry(form_frame, textvariable=ngay_thang_nam_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Chính trị viên
            tk.Label(form_frame, text="Chính trị viên:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            chinh_tri_vien_var = tk.StringVar(value="Đại úy Triệu Văn Dũng")
            tk.Entry(form_frame, textvariable=chinh_tri_vien_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            def save_and_export():
                word_bytes = to_word_docx_to_dan_van(
                    personnel_list,
                    tieu_doan=tieu_doan_var.get(),
                    don_vi=dai_doi_var.get(),  # Sửa dai_doi thành don_vi
                    dia_diem=dia_diem_var.get(),
                    ngay_thang_nam=ngay_thang_nam_var.get(),
                    chinh_tri_vien=chinh_tri_vien_var.get(),
                    db_service=self.db
                )
                
                filename = filedialog.asksaveasfilename(
                    defaultextension=".docx",
                    filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                    title="Lưu file Word"
                )
                
                if filename:
                    with open(filename, 'wb') as f:
                        f.write(word_bytes)
                    messagebox.showinfo("Thành công", f"Đã xuất file Word:\n{filename}")
                    dialog.destroy()
            
            btn_frame = tk.Frame(main_container, bg='#FAFAFA', height=70)
            btn_frame.pack(fill=tk.X, padx=10, pady=10, side=tk.BOTTOM)
            btn_frame.pack_propagate(False)
            
            tk.Button(
                btn_frame,
                text="📄 Xuất File",
                command=save_and_export,
                font=('Segoe UI', 11, 'bold'),
                bg='#4CAF50',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=10)
            
            tk.Button(
                btn_frame,
                text="❌ Hủy",
                command=dialog.destroy,
                font=('Segoe UI', 11),
                bg='#757575',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=5)
            
        except ImportError as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
    
    def export_dang_phai_phan_dong_word(self, get_data_func):
        """Xuất danh sách quân nhân có người thân tham gia đảng phái phản động ra Word"""
        try:
            from services.export_dang_phai_phan_dong import to_word_docx_dang_phai_phan_dong
            
            # Lấy dữ liệu
            data_list = get_data_func()
            if not data_list:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            # Lấy danh sách quân nhân từ IDs
            personnel_ids = [item['id'] for item in data_list]
            personnel_list = [self.db.get_by_id(pid) for pid in personnel_ids if self.db.get_by_id(pid)]
            
            # Mở dialog nhập thông tin
            dialog = tk.Toplevel(self)
            dialog.title("Xuất File Word - Người Thân Đảng Phái Phản Động")
            dialog.geometry("500x400")
            dialog.transient(self)
            dialog.grab_set()
            
            main_container = tk.Frame(dialog, bg='#FAFAFA')
            main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            # Tiểu đoàn
            tk.Label(main_container, text="Tiểu đoàn:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
            tk.Entry(main_container, textvariable=tieu_doan_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Đại đội
            tk.Label(main_container, text="Đại đội:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dai_doi_var = tk.StringVar(value="ĐẠI ĐỘI 3")
            tk.Entry(main_container, textvariable=dai_doi_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Địa điểm
            tk.Label(main_container, text="Địa điểm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dia_diem_var = tk.StringVar(value="Đắk Lắk")
            tk.Entry(main_container, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Năm
            tk.Label(main_container, text="Năm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            nam_var = tk.StringVar(value="2025")
            tk.Entry(main_container, textvariable=nam_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Chính trị viên
            tk.Label(main_container, text="Chính trị viên:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            chinh_tri_vien_var = tk.StringVar(value="Đại úy Triệu Văn Dũng")
            tk.Entry(main_container, textvariable=chinh_tri_vien_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            def save_and_export():
                try:
                    word_bytes = to_word_docx_dang_phai_phan_dong(
                        personnel_list=personnel_list,
                        tieu_doan=tieu_doan_var.get(),
                        dai_doi=dai_doi_var.get(),
                        dia_diem=dia_diem_var.get(),
                        nam=nam_var.get(),
                        chinh_tri_vien=chinh_tri_vien_var.get(),
                        db_service=self.db
                    )
                    
                    filename = filedialog.asksaveasfilename(
                        defaultextension=".docx",
                        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                        title="Lưu file Word"
                    )
                    
                    if filename:
                        with open(filename, 'wb') as f:
                            f.write(word_bytes)
                        messagebox.showinfo("Thành công", f"Đã xuất file Word:\n{filename}")
                        dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
            
            btn_frame = tk.Frame(main_container, bg='#FAFAFA', height=70)
            btn_frame.pack(fill=tk.X, padx=10, pady=10, side=tk.BOTTOM)
            btn_frame.pack_propagate(False)
            
            tk.Button(
                btn_frame,
                text="📄 Xuất File",
                command=save_and_export,
                font=('Segoe UI', 11, 'bold'),
                bg='#4CAF50',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=10)
            
            tk.Button(
                btn_frame,
                text="❌ Hủy",
                command=dialog.destroy,
                font=('Segoe UI', 11),
                bg='#757575',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=5)
            
        except ImportError as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
    
    def choose_to_dan_van_personnel(self, parent):
        """Dialog chọn quân nhân vào tổ công tác dân vận"""
        dialog = tk.Toplevel(parent)
        dialog.title("Chọn Quân Nhân Vào Tổ Công Tác Dân Vận")
        dialog.geometry("1100x700")
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(True, True)
        
        # Dùng grid để control layout tốt hơn
        dialog.grid_rowconfigure(0, weight=1)  # Row 0 (list_frame) có thể expand
        dialog.grid_rowconfigure(1, weight=0)  # Row 1 (btn_frame) không expand
        dialog.grid_columnconfigure(0, weight=1)
        
        # Frame chứa danh sách
        list_frame = tk.Frame(dialog, bg='#FAFAFA')
        list_frame.grid(row=0, column=0, padx=10, pady=10, sticky=tk.NSEW)
        
        # Label
        label = tk.Label(
            list_frame,
            text="Chọn quân nhân vào tổ công tác dân vận:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        )
        label.pack(anchor=tk.W, pady=5)
        
        # Toolbar với tìm kiếm và chọn tất cả
        toolbar_frame = tk.Frame(list_frame, bg='#FAFAFA')
        toolbar_frame.pack(fill=tk.X, pady=5)
        
        # Tìm kiếm
        tk.Label(toolbar_frame, text="Tìm kiếm:", font=('Segoe UI', 9), bg='#FAFAFA').pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(toolbar_frame, textvariable=search_var, width=30, font=('Segoe UI', 9))
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Nút chọn tất cả / Bỏ chọn tất cả
        select_all_var = tk.BooleanVar(value=False)
        selected_ids = set(self.db.get_to_dan_van())
        
        def toggle_select_all():
            select_all_var.set(not select_all_var.get())
            for item in tree.get_children():
                item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
                if item_id:
                    if select_all_var.get():
                        tree.item(item, text='✓')
                        selected_ids.add(item_id)
                    else:
                        tree.item(item, text='')
                        selected_ids.discard(item_id)
        
        select_all_btn = tk.Button(
            toolbar_frame,
            text="☑ Chọn Tất Cả",
            command=toggle_select_all,
            font=('Segoe UI', 9),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor='hand2'
        )
        select_all_btn.pack(side=tk.LEFT, padx=5)
        
        # Treeview với checkbox
        columns = ('Họ và Tên', 'Ngày Sinh', 'Cấp Bậc', 'Chức Vụ', 'Đơn Vị', 'Dân Tộc')
        tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=20)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Configure columns
        tree.heading('#0', text='Chọn')
        tree.column('#0', width=50, anchor=tk.CENTER)
        for col in columns:
            tree.heading(col, text=col)
            if col == 'Họ và Tên':
                tree.column(col, width=200)
            else:
                tree.column(col, width=120)
        
        # Load data - tất cả quân nhân
        all_personnel = self.db.get_all()
        
        def load_tree_data():
            """Load dữ liệu vào tree"""
            # Xóa dữ liệu cũ
            for item in tree.get_children():
                tree.delete(item)
            
            # Lọc theo tìm kiếm
            search_text = search_var.get().lower()
            display_personnel = all_personnel
            if search_text:
                display_personnel = [p for p in all_personnel 
                                  if search_text in (p.hoTen or '').lower()]
            
            for person in display_personnel:
                is_selected = person.id in selected_ids
                item_text = '✓' if is_selected else ''
                
                item = tree.insert('', 'end', 
                                  text=item_text,
                                  tags=(person.id,),
                                  values=(
                                      person.hoTen or '',
                                      person.ngaySinh or '',
                                      person.capBac or '',
                                      person.chucVu or '',
                                      person.donVi or '',
                                      person.danToc or ''
                                  ))
        
        def on_item_click(event):
            """Xử lý click vào item"""
            item = tree.identify_row(event.y)
            if item:
                # Set selection để highlight row
                tree.selection_set(item)
                item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
                if item_id:
                    current_text = tree.item(item, 'text')
                    if current_text == '✓':
                        tree.item(item, text='')
                        selected_ids.discard(item_id)
                    else:
                        tree.item(item, text='✓')
                        selected_ids.add(item_id)
        
        tree.bind('<Button-1>', on_item_click)
        search_var.trace('w', lambda *args: load_tree_data())
        
        # Pack tree và scrollbar
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Load dữ liệu ban đầu
        load_tree_data()
        
        # Buttons - Row 1, LUÔN HIỂN THỊ
        btn_frame = tk.Frame(dialog, bg='#FAFAFA', height=70)
        btn_frame.grid(row=1, column=0, padx=10, pady=10, sticky=tk.EW)
        btn_frame.grid_propagate(False)
        btn_frame.grid_columnconfigure(1, weight=1)
        
        def save_selection():
            """Lưu danh sách đã chọn"""
            # Kiểm tra lại selected_ids từ tree để đảm bảo đồng bộ
            current_selected = set()
            for item in tree.get_children():
                item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
                if item_id and tree.item(item, 'text') == '✓':
                    current_selected.add(item_id)
            
            # Cập nhật selected_ids với dữ liệu từ tree
            selected_ids.clear()
            selected_ids.update(current_selected)
            
            success_count = 0
            for personnel_id in selected_ids:
                if self.db.add_to_dan_van(personnel_id):
                    success_count += 1
            
            # Xóa những quân nhân không được chọn
            all_selected_ids = set(self.db.get_to_dan_van())
            to_remove = all_selected_ids - selected_ids
            for personnel_id in to_remove:
                self.db.remove_to_dan_van(personnel_id)
            
            messagebox.showinfo("Thành công", f"Đã cập nhật {success_count} quân nhân vào danh sách!")
            dialog.destroy()
            # Refresh lại tab
            self.create_to_dan_van_tab(parent)
        
        # Nút Hủy
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor='hand2',
            width=10
        ).grid(row=0, column=0, padx=5, sticky=tk.W)
        
        # Spacer
        tk.Frame(btn_frame, bg='#FAFAFA').grid(row=0, column=1, sticky=tk.EW)
        
        # Nút XONG
        tk.Button(
            btn_frame,
            text="✅ XONG",
            command=save_selection,
            font=('Segoe UI', 11, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.RAISED,
            padx=25,
            pady=8,
            cursor='hand2',
            width=12,
            bd=2
        ).grid(row=0, column=2, padx=5, sticky=tk.E)
    
    def choose_dang_vien_dien_tap_personnel(self, parent):
        """Dialog chọn quân nhân vào danh sách đảng viên diễn tập"""
        dialog = tk.Toplevel(parent)
        dialog.title("Chọn Quân Nhân Đảng Viên Diễn Tập")
        dialog.geometry("1100x700")
        dialog.transient(parent)
        dialog.grab_set()
        dialog.resizable(True, True)
        
        # Dùng grid để control layout tốt hơn
        dialog.grid_rowconfigure(0, weight=1)  # Row 0 (list_frame) có thể expand
        dialog.grid_rowconfigure(1, weight=0)  # Row 1 (btn_frame) không expand
        dialog.grid_columnconfigure(0, weight=1)
        
        # Frame chứa danh sách
        list_frame = tk.Frame(dialog, bg='#FAFAFA')
        list_frame.grid(row=0, column=0, padx=10, pady=10, sticky=tk.NSEW)
        
        # Label
        label = tk.Label(
            list_frame,
            text="Chọn quân nhân đảng viên vào danh sách diễn tập:",
            font=('Segoe UI', 10),
            bg='#FAFAFA'
        )
        label.pack(anchor=tk.W, pady=5)
        
        # Toolbar với tìm kiếm và chọn tất cả
        toolbar_frame = tk.Frame(list_frame, bg='#FAFAFA')
        toolbar_frame.pack(fill=tk.X, pady=5)
        
        # Tìm kiếm
        tk.Label(toolbar_frame, text="Tìm kiếm:", font=('Segoe UI', 9), bg='#FAFAFA').pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = tk.Entry(toolbar_frame, textvariable=search_var, width=30, font=('Segoe UI', 9))
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Nút chọn tất cả / Bỏ chọn tất cả
        select_all_var = tk.BooleanVar(value=False)
        selected_ids = set(self.db.get_dang_vien_dien_tap())
        
        def toggle_select_all():
            select_all_var.set(not select_all_var.get())
            for item in tree.get_children():
                item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
                if item_id:
                    if select_all_var.get():
                        tree.item(item, text='✓')
                        selected_ids.add(item_id)
                    else:
                        tree.item(item, text='')
                        selected_ids.discard(item_id)
        
        select_all_btn = tk.Button(
            toolbar_frame,
            text="☑ Chọn Tất Cả",
            command=toggle_select_all,
            font=('Segoe UI', 9),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor='hand2'
        )
        select_all_btn.pack(side=tk.LEFT, padx=5)
        
        # Treeview với checkbox
        columns = ('Họ và Tên', 'Ngày Sinh', 'Cấp Bậc', 'Chức Vụ', 'Đơn Vị', 'Dân Tộc')
        tree = ttk.Treeview(list_frame, columns=columns, show='tree headings', height=20)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Configure columns
        tree.heading('#0', text='Chọn')
        tree.column('#0', width=50, anchor=tk.CENTER)
        for col in columns:
            tree.heading(col, text=col)
            if col == 'Họ và Tên':
                tree.column(col, width=200)
            else:
                tree.column(col, width=120)
        
        # Load data - chỉ lấy đảng viên
        all_personnel = self.db.get_all()
        dang_vien = [p for p in all_personnel 
                     if p.thongTinKhac.dang.ngayVao or p.thongTinKhac.dang.ngayChinhThuc]
        
        def load_tree_data():
            """Load dữ liệu vào tree"""
            # Xóa dữ liệu cũ
            for item in tree.get_children():
                tree.delete(item)
            
            # Lọc theo tìm kiếm
            search_text = search_var.get().lower()
            display_personnel = dang_vien
            if search_text:
                display_personnel = [p for p in dang_vien 
                                  if search_text in (p.hoTen or '').lower()]
            
            for person in display_personnel:
                is_selected = person.id in selected_ids
                item_text = '✓' if is_selected else ''
                
                item = tree.insert('', 'end', 
                                  text=item_text,
                                  tags=(person.id,),
                                  values=(
                                      person.hoTen or '',
                                      person.ngaySinh or '',
                                      person.capBac or '',
                                      person.chucVu or '',
                                      person.donVi or '',
                                      person.danToc or ''
                                  ))
        
        def on_item_click(event):
            """Xử lý click vào item"""
            item = tree.identify_row(event.y)
            if item:
                # Set selection để highlight row
                tree.selection_set(item)
                item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
                if item_id:
                    current_text = tree.item(item, 'text')
                    if current_text == '✓':
                        tree.item(item, text='')
                        selected_ids.discard(item_id)
                    else:
                        tree.item(item, text='✓')
                        selected_ids.add(item_id)
        
        tree.bind('<Button-1>', on_item_click)
        search_var.trace('w', lambda *args: load_tree_data())
        
        # Pack tree và scrollbar
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Load dữ liệu ban đầu
        load_tree_data()
        
        # Buttons - Row 1, LUÔN HIỂN THỊ
        btn_frame = tk.Frame(dialog, bg='#FAFAFA', height=70)
        btn_frame.grid(row=1, column=0, padx=10, pady=10, sticky=tk.EW)
        btn_frame.grid_propagate(False)
        btn_frame.grid_columnconfigure(1, weight=1)
        
        def save_selection():
            """Lưu danh sách đã chọn"""
            # Kiểm tra lại selected_ids từ tree để đảm bảo đồng bộ
            current_selected = set()
            for item in tree.get_children():
                item_id = tree.item(item, 'tags')[0] if tree.item(item, 'tags') else None
                if item_id and tree.item(item, 'text') == '✓':
                    current_selected.add(item_id)
            
            # Cập nhật selected_ids với dữ liệu từ tree
            selected_ids.clear()
            selected_ids.update(current_selected)
            
            success_count = 0
            for personnel_id in selected_ids:
                if self.db.add_dang_vien_dien_tap(personnel_id):
                    success_count += 1
            
            # Xóa những quân nhân không được chọn
            all_selected_ids = set(self.db.get_dang_vien_dien_tap())
            to_remove = all_selected_ids - selected_ids
            for personnel_id in to_remove:
                self.db.remove_dang_vien_dien_tap(personnel_id)
            
            messagebox.showinfo("Thành công", f"Đã cập nhật {success_count} quân nhân vào danh sách!")
            dialog.destroy()
            # Refresh lại tab
            self.create_dang_vien_dien_tap_tab(parent)
        
        # Nút Hủy
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor='hand2',
            width=10
        ).grid(row=0, column=0, padx=5, sticky=tk.W)
        
        # Spacer
        tk.Frame(btn_frame, bg='#FAFAFA').grid(row=0, column=1, sticky=tk.EW)
        
        # Nút XONG
        tk.Button(
            btn_frame,
            text="✅ XONG",
            command=save_selection,
            font=('Segoe UI', 11, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.RAISED,
            padx=25,
            pady=8,
            cursor='hand2',
            width=12,
            bd=2
        ).grid(row=0, column=2, padx=5, sticky=tk.E)
    
    def export_dang_vien_dien_tap_word(self, get_data_func):
        """Xuất danh sách đảng viên diễn tập ra Word"""
        try:
            from tkinter import messagebox, filedialog
            from services.export_dang_vien_dien_tap import to_word_docx_dang_vien_dien_tap
            
            # Lấy dữ liệu
            data = get_data_func()
            if not data:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            # Lấy danh sách Personnel từ IDs
            personnel_list = []
            for item in data:
                personnel = self.db.get_by_id(item['id'])
                if personnel:
                    personnel_list.append(personnel)
            
            if not personnel_list:
                messagebox.showwarning("Cảnh báo", "Không tìm thấy dữ liệu quân nhân!")
                return
            
            # Dialog nhập thông tin
            dialog = tk.Toplevel(self)
            dialog.title("Xuất File Word - Đảng Viên Diễn Tập")
            dialog.geometry("500x400")
            dialog.transient(self)
            dialog.grab_set()
            dialog.resizable(False, False)
            
            main_container = tk.Frame(dialog, bg='#FAFAFA')
            main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            # Form fields
            tk.Label(main_container, text="ĐẢNG BỘ:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
            tk.Entry(main_container, textvariable=tieu_doan_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=5)
            
            tk.Label(main_container, text="CHI BỘ:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dai_doi_var = tk.StringVar(value="ĐẠI ĐỘI 3")
            tk.Entry(main_container, textvariable=dai_doi_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=5)
            
            tk.Label(main_container, text="Địa điểm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dia_diem_var = tk.StringVar(value="Đăk Lăk")
            tk.Entry(main_container, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=5)
            
            tk.Label(main_container, text="Năm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            nam_var = tk.StringVar(value="2025")
            tk.Entry(main_container, textvariable=nam_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=5)
            
            def save_and_export():
                try:
                    word_bytes = to_word_docx_dang_vien_dien_tap(
                        personnel_list=personnel_list,
                        tieu_doan=tieu_doan_var.get(),
                        dai_doi=dai_doi_var.get(),
                        dia_diem=dia_diem_var.get(),
                        nam=nam_var.get(),
                        db_service=self.db
                    )
                    
                    filename = filedialog.asksaveasfilename(
                        defaultextension=".docx",
                        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                        title="Lưu file Word"
                    )
                    
                    if filename:
                        with open(filename, 'wb') as f:
                            f.write(word_bytes)
                        messagebox.showinfo("Thành công", f"Đã xuất file Word:\n{filename}")
                        dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
            
            btn_frame = tk.Frame(main_container, bg='#FAFAFA', height=70)
            btn_frame.pack(fill=tk.X, padx=10, pady=10, side=tk.BOTTOM)
            btn_frame.pack_propagate(False)
            
            tk.Button(
                btn_frame,
                text="📄 Xuất File",
                command=save_and_export,
                font=('Segoe UI', 11, 'bold'),
                bg='#4CAF50',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=10)
            
            tk.Button(
                btn_frame,
                text="❌ Hủy",
                command=dialog.destroy,
                font=('Segoe UI', 11),
                bg='#757575',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=5)
            
        except ImportError as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
    
    def export_dang_phai_phan_dong_word(self, get_data_func):
        """Xuất danh sách quân nhân có người thân tham gia đảng phái phản động ra Word"""
        try:
            from services.export_dang_phai_phan_dong import to_word_docx_dang_phai_phan_dong
            
            # Lấy dữ liệu
            data_list = get_data_func()
            if not data_list:
                messagebox.showwarning("Cảnh báo", "Không có dữ liệu để xuất!")
                return
            
            # Lấy danh sách quân nhân từ IDs
            personnel_ids = [item['id'] for item in data_list]
            personnel_list = [self.db.get_by_id(pid) for pid in personnel_ids if self.db.get_by_id(pid)]
            
            # Mở dialog nhập thông tin
            dialog = tk.Toplevel(self)
            dialog.title("Xuất File Word - Người Thân Đảng Phái Phản Động")
            dialog.geometry("500x400")
            dialog.transient(self)
            dialog.grab_set()
            
            main_container = tk.Frame(dialog, bg='#FAFAFA')
            main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            # Tiểu đoàn
            tk.Label(main_container, text="Tiểu đoàn:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
            tk.Entry(main_container, textvariable=tieu_doan_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Đại đội
            tk.Label(main_container, text="Đại đội:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dai_doi_var = tk.StringVar(value="ĐẠI ĐỘI 3")
            tk.Entry(main_container, textvariable=dai_doi_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Địa điểm
            tk.Label(main_container, text="Địa điểm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            dia_diem_var = tk.StringVar(value="Đắk Lắk")
            tk.Entry(main_container, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Năm
            tk.Label(main_container, text="Năm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            nam_var = tk.StringVar(value="2025")
            tk.Entry(main_container, textvariable=nam_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            # Chính trị viên
            tk.Label(main_container, text="Chính trị viên:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
            chinh_tri_vien_var = tk.StringVar(value="Đại úy Triệu Văn Dũng")
            tk.Entry(main_container, textvariable=chinh_tri_vien_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
            
            def save_and_export():
                try:
                    word_bytes = to_word_docx_dang_phai_phan_dong(
                        personnel_list=personnel_list,
                        tieu_doan=tieu_doan_var.get(),
                        dai_doi=dai_doi_var.get(),
                        dia_diem=dia_diem_var.get(),
                        nam=nam_var.get(),
                        chinh_tri_vien=chinh_tri_vien_var.get(),
                        db_service=self.db
                    )
                    
                    filename = filedialog.asksaveasfilename(
                        defaultextension=".docx",
                        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                        title="Lưu file Word"
                    )
                    
                    if filename:
                        with open(filename, 'wb') as f:
                            f.write(word_bytes)
                        messagebox.showinfo("Thành công", f"Đã xuất file Word:\n{filename}")
                        dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
            
            btn_frame = tk.Frame(main_container, bg='#FAFAFA', height=70)
            btn_frame.pack(fill=tk.X, padx=10, pady=10, side=tk.BOTTOM)
            btn_frame.pack_propagate(False)
            
            tk.Button(
                btn_frame,
                text="📄 Xuất File",
                command=save_and_export,
                font=('Segoe UI', 11, 'bold'),
                bg='#4CAF50',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=10)
            
            tk.Button(
                btn_frame,
                text="❌ Hủy",
                command=dialog.destroy,
                font=('Segoe UI', 11),
                bg='#757575',
                fg='white',
                relief=tk.FLAT,
                padx=30,
                pady=8,
                cursor='hand2',
                width=12
            ).pack(side=tk.RIGHT, padx=5)
            
        except ImportError as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file Word:\n{str(e)}")