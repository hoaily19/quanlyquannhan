"""
Frame danh sách quân nhân
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
from gui.tooltip import create_tooltip


class PersonnelListFrame(tk.Frame):
    """Frame hiển thị danh sách quân nhân"""
    
    def __init__(self, parent, db: DatabaseService):
        """
        Args:
            parent: Parent widget
            db: DatabaseService instance
        """
        super().__init__(parent)
        self.db = db
        self.personnel_list = []
        self.selected_id = None
        self.setup_ui()
        self.load_data()
    
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
            text="📋 DANH SÁCH QUÂN NHÂN",
            font=('Arial', 16, 'bold'),
            bg=MILITARY_COLORS['primary'],
            fg=MILITARY_COLORS['text_light']
        )
        title_label.pack(expand=True)
        
        # Toolbar với nền trắng và border đẹp
        toolbar = tk.Frame(self, bg='#FFFFFF', relief=tk.RAISED, bd=1)
        toolbar.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        # Tìm kiếm
        search_frame = tk.Frame(toolbar, bg='#FFFFFF')
        search_frame.pack(side=tk.LEFT, padx=10, pady=8)
        
        tk.Label(
            search_frame,
            text="🔍 Tìm kiếm:",
            font=('Arial', 10, 'bold'),
            bg='#FFFFFF',
            fg=MILITARY_COLORS['text_dark']
        ).pack(side=tk.LEFT, padx=5)
        
        self.search_var = tk.StringVar()
        self.search_var.trace('w', lambda *args: self.filter_data())
        search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            width=30,
            font=('Arial', 10),
            relief=tk.SOLID,
            bd=1,
            highlightthickness=1,
            highlightcolor=MILITARY_COLORS['primary'],
            highlightbackground='#CCCCCC'
        )
        search_entry.pack(side=tk.LEFT, padx=5)
        
        # Filters
        filter_frame = tk.Frame(toolbar, bg='#FFFFFF')
        filter_frame.pack(side=tk.LEFT, padx=15)
        
        # Đơn vị
        tk.Label(
            filter_frame,
            text="Đơn vị:",
            font=('Arial', 10, 'bold'),
            bg='#FFFFFF',
            fg=MILITARY_COLORS['text_dark']
        ).pack(side=tk.LEFT, padx=5)
        
        self.unit_var = tk.StringVar()
        unit_combo = ttk.Combobox(filter_frame, textvariable=self.unit_var, width=15, state='readonly')
        unit_combo['values'] = [''] + self.db.get_unique_values('donVi')
        unit_combo.pack(side=tk.LEFT, padx=5)
        unit_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_data())
        
        # Cấp bậc
        tk.Label(
            filter_frame,
            text="Cấp bậc:",
            font=('Arial', 10, 'bold'),
            bg='#FFFFFF',
            fg=MILITARY_COLORS['text_dark']
        ).pack(side=tk.LEFT, padx=(10, 5))
        
        self.rank_var = tk.StringVar()
        rank_combo = ttk.Combobox(filter_frame, textvariable=self.rank_var, width=15, state='readonly')
        rank_combo['values'] = [''] + self.db.get_unique_values('capBac')
        rank_combo.pack(side=tk.LEFT, padx=5)
        rank_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_data())
        
        # Dân tộc
        tk.Label(
            filter_frame,
            text="Dân tộc:",
            font=('Arial', 10, 'bold'),
            bg='#FFFFFF',
            fg=MILITARY_COLORS['text_dark']
        ).pack(side=tk.LEFT, padx=(10, 5))
        
        self.ethnic_var = tk.StringVar()
        ethnic_combo = ttk.Combobox(filter_frame, textvariable=self.ethnic_var, width=15, state='readonly')
        ethnic_combo['values'] = [''] + self.db.get_unique_values('danToc')
        ethnic_combo.pack(side=tk.LEFT, padx=5)
        ethnic_combo.bind('<<ComboboxSelected>>', lambda e: self.filter_data())
        
        # Buttons - Hiển thị icon + text, kích thước đồng đều, dễ bấm
        btn_frame = tk.Frame(toolbar, bg='#FFFFFF')
        btn_frame.pack(side=tk.RIGHT, padx=10, pady=5)
        
        common_btn_opts = {
            "width": 11,
            "height": 1,
        }
        
        # Nút xem chi tiết
        view_btn_toolbar = tk.Button(
            btn_frame,
            text="👁️ Chi Tiết",
            command=lambda: self.view_selected(),
            **get_button_style('info'),
            **common_btn_opts,
        )
        view_btn_toolbar.pack(side=tk.LEFT, padx=4)
        
        # Nút sửa
        edit_btn_toolbar = tk.Button(
            btn_frame,
            text="✏️ Sửa",
            command=lambda: self.edit_selected(),
            **get_button_style('secondary'),
            **common_btn_opts,
        )
        edit_btn_toolbar.pack(side=tk.LEFT, padx=4)
        
        # Nút xóa
        delete_btn_toolbar = tk.Button(
            btn_frame,
            text="🗑️ Xóa",
            command=lambda: self.delete_selected(),
            **get_button_style('danger'),
            **common_btn_opts,
        )
        delete_btn_toolbar.pack(side=tk.LEFT, padx=4)
        
        # Separator giữa nhóm thao tác và nhóm thêm/xuất
        separator = tk.Frame(btn_frame, bg='#CCCCCC', width=1)
        separator.pack(side=tk.LEFT, padx=6, fill=tk.Y, pady=4)
        
        add_btn = tk.Button(
            btn_frame,
            text="➕ Thêm Mới",
            command=self.add_new,
            **get_button_style('success'),
            **common_btn_opts,
        )
        add_btn.pack(side=tk.LEFT, padx=4)
        
        export_btn = tk.Button(
            btn_frame,
            text="📥 Xuất CSV",
            command=self.export_csv,
            **get_button_style('secondary'),
            **common_btn_opts,
        )
        export_btn.pack(side=tk.LEFT, padx=4)
        
        export_word_btn = tk.Button(
            btn_frame,
            text="📄 Xuất Word",
            command=self.export_word,
            **get_button_style('secondary'),
            **common_btn_opts,
        )
        export_word_btn.pack(side=tk.LEFT, padx=4)
        
        # Treeview với border đẹp hơn
        tree_frame = tk.Frame(self, bg='#FFFFFF', relief=tk.SOLID, bd=1)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        # Định nghĩa columns (bỏ cột actions vì không thể đặt widget)
        columns = ('stt', 'hoTen', 'capBac', 'chucVu', 'donVi', 'danToc')
        
        self.tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
            height=18
        )
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Lưu trạng thái sắp xếp để toggle A-Z / Z-A
        self.sort_state = {}
        
        # Columns headings và width - căn chữ thẳng với header, dễ đọc
        self.tree.heading('stt', text='STT', anchor=tk.CENTER)
        self.tree.heading('hoTen', text='Họ và Tên', anchor=tk.W,
                          command=lambda: self.sort_by_column('hoTen'))
        self.tree.heading('capBac', text='Cấp Bậc', anchor=tk.W,
                          command=lambda: self.sort_by_column('capBac'))
        self.tree.heading('chucVu', text='Chức Vụ', anchor=tk.W,
                          command=lambda: self.sort_by_column('chucVu'))
        self.tree.heading('donVi', text='Đơn Vị', anchor=tk.W,
                          command=lambda: self.sort_by_column('donVi'))
        self.tree.heading('danToc', text='Dân Tộc', anchor=tk.W,
                          command=lambda: self.sort_by_column('danToc'))
        
        self.tree.column('stt', width=60, anchor=tk.CENTER, minwidth=50)
        self.tree.column('hoTen', width=260, anchor=tk.W, minwidth=220)
        self.tree.column('capBac', width=120, anchor=tk.W, minwidth=100)
        self.tree.column('chucVu', width=160, anchor=tk.W, minwidth=130)
        self.tree.column('donVi', width=220, anchor=tk.W, minwidth=170)
        self.tree.column('danToc', width=120, anchor=tk.W, minwidth=100)
        
        # Style cho treeview - đẹp hơn, có cảm giác kẻ hàng
        style = ttk.Style()
        style.configure("Treeview", 
                       rowheight=40,  # Tăng chiều cao hàng
                       font=('Arial', 10),
                       background='#FFFFFF',
                       fieldbackground='#FFFFFF',
                       borderwidth=1,
                       relief=tk.SOLID)
        style.configure("Treeview.Heading", 
                       font=('Arial', 11, 'bold'), 
                       background=MILITARY_COLORS['primary'],
                       foreground=MILITARY_COLORS['text_light'],
                       relief=tk.FLAT)
        style.map("Treeview.Heading",
                 background=[('active', MILITARY_COLORS['primary_dark'])])
        style.map("Treeview",
                 background=[('selected', '#E3F2FD')],  # Xanh nhạt khi chọn
                 foreground=[('selected', MILITARY_COLORS['text_dark'])])
        
        # Kẻ hàng xen kẽ cho dễ đọc
        self.tree.tag_configure('evenrow', background='#FFFFFF')
        self.tree.tag_configure('oddrow', background='#F5F5F5')
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Bind events
        self.tree.bind('<Double-1>', self.on_double_click)
        self.tree.bind('<Button-3>', self.on_right_click)
        self.tree.bind('<Button-1>', self.on_single_click)
        
        # Thêm nút actions bên dưới treeview với style đẹp hơn và nổi bật hơn
        action_frame = tk.Frame(self, bg='#FFFFFF', relief=tk.RAISED, bd=2)
        action_frame.pack(fill=tk.X, padx=10, pady=(5, 10))
        
        # Label hướng dẫn và số lượng
        left_info = tk.Frame(action_frame, bg='#FFFFFF')
        left_info.pack(side=tk.LEFT, padx=15, pady=10)
        
        hint_label = tk.Label(
            left_info,
            text="💡 Chọn một hàng để thao tác:",
            font=('Arial', 10, 'bold'),
            bg='#FFFFFF',
            fg=MILITARY_COLORS['primary_dark']
        )
        hint_label.pack(side=tk.LEFT, padx=5)
        
        # Buttons với spacing đều và kích thước lớn hơn - Chỉ hiển thị icon, text trong tooltip
        btn_container = tk.Frame(action_frame, bg='#FFFFFF')
        btn_container.pack(side=tk.RIGHT, padx=15, pady=8)
        
        # Nút Chi Tiết - màu xanh dương nổi bật - Icon + text
        view_btn = tk.Button(
            btn_container,
            text="👁️ Chi Tiết",
            command=lambda: self.view_selected(),
            **get_button_style('info'),
            width=12,
            height=2
        )
        view_btn.pack(side=tk.LEFT, padx=5)
        
        # Nút Sửa - màu xanh lá - Icon + text
        edit_btn = tk.Button(
            btn_container,
            text="✏️ Sửa",
            command=lambda: self.edit_selected(),
            **get_button_style('secondary'),
            width=12,
            height=2
        )
        edit_btn.pack(side=tk.LEFT, padx=5)
        
        # Nút Xóa - màu đỏ - Icon + text
        delete_btn = tk.Button(
            btn_container,
            text="🗑️ Xóa",
            command=lambda: self.delete_selected(),
            **get_button_style('danger'),
            width=12,
            height=2
        )
        delete_btn.pack(side=tk.LEFT, padx=5)
    
    def on_single_click(self, event):
        """Xử lý single click để lưu selection"""
        item = self.tree.identify_row(event.y)
        if item:
            self.selected_id = item
            # Highlight row
            self.tree.selection_set(item)
    
    def sort_by_column(self, column_key: str):
        """Sắp xếp theo cột được chọn, toggle A-Z / Z-A"""
        if not self.personnel_list:
            return
        
        # Map tên cột treeview -> thuộc tính Personnel
        attr_map = {
            'hoTen': 'hoTen',
            'capBac': 'capBac',
            'chucVu': 'chucVu',
            'donVi': 'donVi',
            'danToc': 'danToc',
        }
        if column_key not in attr_map:
            return
        
        # Toggle trạng thái: lần đầu A-Z, lần sau Z-A
        reverse = self.sort_state.get(column_key, False)
        reverse = not reverse
        self.sort_state[column_key] = reverse
        
        attr_name = attr_map[column_key]
        
        def sort_key(person):
            value = getattr(person, attr_name, '') or ''
            return value.lower()
        
        try:
            self.personnel_list.sort(key=sort_key, reverse=reverse)
            self.refresh_tree()
        except Exception:
            # Nếu có lỗi khi sort, bỏ qua để không làm bể giao diện
            pass
    
    def load_data(self):
        """Load dữ liệu - Xử lý lỗi an toàn"""
        try:
            self.personnel_list = self.db.get_all()
            self.refresh_tree()
        except Exception as e:
            # Xử lý lỗi khi load data - không để giao diện bị nát
            import traceback
            traceback.print_exc()
            try:
                from tkinter import messagebox
                messagebox.showerror("Lỗi", f"Không thể tải dữ liệu:\n{str(e)}")
            except:
                print(f"Lỗi khi load data: {str(e)}")
            # Đảm bảo personnel_list luôn là list
            self.personnel_list = []
            # Refresh tree với list rỗng
            try:
                self.refresh_tree()
            except:
                pass
    
    def filter_data(self):
        """Lọc dữ liệu"""
        search_query = self.search_var.get()
        filters = {}
        
        if self.unit_var.get():
            filters['donVi'] = self.unit_var.get()
        if self.rank_var.get():
            filters['capBac'] = self.rank_var.get()
        if hasattr(self, 'ethnic_var') and self.ethnic_var.get():
            filters['danToc'] = self.ethnic_var.get()
        
        self.personnel_list = self.db.search(search_query, filters if filters else None)
        self.refresh_tree()
    
    def refresh_tree(self):
        """Refresh treeview - Sửa lỗi Item already exists"""
        try:
            # Xóa tất cả items một cách an toàn
            for item in self.tree.get_children():
                try:
                    self.tree.delete(item)
                except:
                    pass
            
            # Đảm bảo treeview đã được clear hoàn toàn
            self.tree.delete(*self.tree.get_children())
            
            # Cache đơn vị để tránh query nhiều lần
            units_cache = {}
            try:
                all_units = self.db.get_all_units()
                for unit in all_units:
                    units_cache[unit.id] = unit.ten
            except:
                pass
            
            # Set để track các ID đã insert (tránh duplicate)
            inserted_ids = set()
            
            # Thêm items mới với STT
            for idx, person in enumerate(self.personnel_list, 1):
                # Kiểm tra ID hợp lệ và chưa được insert
                if not person.id:
                    # Nếu không có ID, tạo ID tạm thời
                    temp_id = f"temp_{idx}"
                    while temp_id in inserted_ids:
                        temp_id = f"temp_{idx}_{len(inserted_ids)}"
                    person_id = temp_id
                else:
                    person_id = person.id
                
                # Bỏ qua nếu ID đã được insert
                if person_id in inserted_ids:
                    continue
                
                try:
                    # Lấy tên đơn vị từ unitId nếu có
                    don_vi_display = person.donVi or ''
                    if person.unitId and person.unitId in units_cache:
                        don_vi_display = units_cache[person.unitId]
                    
                    # Kiểm tra xem item đã tồn tại chưa
                    if person_id in self.tree.get_children():
                        # Nếu đã tồn tại, xóa và insert lại
                        self.tree.delete(person_id)
                    
                    row_tag = 'oddrow' if idx % 2 else 'evenrow'
                    self.tree.insert(
                        '',
                        'end',
                        iid=person_id,
                        values=(
                            idx,
                            person.hoTen or 'Chưa có tên',
                            person.capBac or '',
                            person.chucVu or '',
                            don_vi_display,
                            person.danToc or '',
                        ),
                        tags=(row_tag,),
                    )
                    inserted_ids.add(person_id)
                except tk.TclError as e:
                    # Nếu lỗi "Item already exists", bỏ qua và tiếp tục
                    if "already exists" in str(e):
                        continue
                    else:
                        # Log lỗi khác nhưng không dừng
                        import traceback
                        traceback.print_exc()
                        continue
        except Exception as e:
            # Xử lý lỗi tổng quát - không để giao diện bị nát
            import traceback
            traceback.print_exc()
            # Không hiển thị messagebox ở đây vì có thể gây loop
            print(f"Lỗi khi refresh tree: {str(e)}")
        
        # Hiển thị thông báo số lượng trong action_frame
        count_label_text = f"📊 Tổng số: {len(self.personnel_list)} quân nhân"
        if hasattr(self, 'count_label'):
            self.count_label.config(text=count_label_text)
        else:
            self.count_label = tk.Label(
                self,
                text=count_label_text,
                font=('Arial', 10, 'bold'),
                bg=MILITARY_COLORS['bg_light'],
                fg=MILITARY_COLORS['primary_dark']
            )
            self.count_label.pack(pady=5)
    
    def on_double_click(self, event):
        """Xử lý double click"""
        selection = self.tree.selection()
        if selection:
            personnel_id = selection[0]
            self.edit_personnel(personnel_id)
    
    def on_right_click(self, event):
        """Xử lý right click - context menu"""
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.selected_id = item
            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label="✏️ Sửa", command=lambda: self.edit_personnel(item))
            menu.add_command(label="🗑️ Xóa", command=lambda: self.delete_personnel(item))
            menu.add_command(label="👁️ Xem Chi Tiết", command=lambda: self.view_detail(item))
            menu.post(event.x_root, event.y_root)
    
    def add_new(self):
        """Thêm mới - mở form trong cửa sổ modal (Toplevel)"""
        try:
            from gui.personnel_form_frame import PersonnelFormFrame
            
            # Lấy root window
            root_window = self.winfo_toplevel()
            
            # Tạo cửa sổ modal
            dialog = tk.Toplevel(root_window)
            dialog.title("Thêm Quân Nhân Mới")
            try:
                dialog.configure(bg=MILITARY_COLORS['bg_light'])
            except:
                pass
            
            # Kích thước: gần full cao, thu gọn ngang để dễ xem
            try:
                root_window.update_idletasks()
                screen_width = root_window.winfo_screenwidth()
                screen_height = root_window.winfo_screenheight()
                
                width = int(screen_width * 0.9)
                height = int(screen_height * 0.9)
                x = int((screen_width - width) / 2)
                y = int((screen_height - height) / 2)
                dialog.geometry(f"{width}x{height}+{x}+{y}")
            except:
                pass
            
            # Đặt modal: luôn trên root và khóa focus
            dialog.transient(root_window)
            dialog.grab_set()
            
            # Khi đóng bằng nút X
            def on_close():
                try:
                    dialog.grab_release()
                except:
                    pass
                try:
                    dialog.destroy()
                except:
                    pass
            
            dialog.protocol("WM_DELETE_WINDOW", on_close)
            
            # Nhúng form vào dialog
            form_frame = PersonnelFormFrame(dialog, self.db, is_new=True)
            form_frame.pack(fill=tk.BOTH, expand=True)
            
            try:
                dialog.focus_set()
            except:
                pass
            
            # Chờ đến khi dialog đóng rồi mới tiếp tục
            root_window.wait_window(dialog)
            
            # Sau khi thêm xong, reload lại danh sách
            try:
                self.load_data()
            except:
                pass
        except Exception as e:
            # Nếu có lỗi với modal, fallback về cơ chế cũ để không chặn người dùng
            try:
                if hasattr(self.master, 'master') and hasattr(self.master.master, 'show_frame'):
                    self.master.master.show_frame('add')
                else:
                    from gui.personnel_form_frame import PersonnelFormFrame
                    self.destroy()
                    form_frame = PersonnelFormFrame(self.master, self.db, is_new=True)
                    form_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            except:
                import traceback
                traceback.print_exc()
    
    def edit_selected(self):
        """Sửa quân nhân đã chọn"""
        if not self.selected_id:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn quân nhân cần sửa")
            return
        self.edit_personnel(self.selected_id)
    
    def delete_selected(self):
        """Xóa quân nhân đã chọn"""
        if not self.selected_id:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn quân nhân cần xóa")
            return
        self.delete_personnel(self.selected_id)
    
    def view_selected(self):
        """Xem chi tiết quân nhân đã chọn"""
        if not self.selected_id:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn quân nhân cần xem")
            return
        self.view_detail(self.selected_id)
    
    def edit_personnel(self, personnel_id: str):
        """Sửa quân nhân - mở form trong cửa sổ modal (Toplevel)"""
        try:
            from gui.personnel_form_frame import PersonnelFormFrame
            
            # Lấy root window
            root_window = self.winfo_toplevel()
            
            # Tạo cửa sổ modal
            dialog = tk.Toplevel(root_window)
            dialog.title("Sửa Thông Tin Quân Nhân")
            try:
                dialog.configure(bg=MILITARY_COLORS['bg_light'])
            except:
                pass
            
            # Kích thước: giống cửa sổ thêm mới
            try:
                root_window.update_idletasks()
                screen_width = root_window.winfo_screenwidth()
                screen_height = root_window.winfo_screenheight()
                
                width = int(screen_width * 0.9)
                height = int(screen_height * 0.9)
                x = int((screen_width - width) / 2)
                y = int((screen_height - height) / 2)
                dialog.geometry(f"{width}x{height}+{x}+{y}")
            except:
                pass
            
            # Đặt modal
            dialog.transient(root_window)
            dialog.grab_set()
            
            # Khi đóng bằng nút X
            def on_close():
                try:
                    dialog.grab_release()
                except:
                    pass
                try:
                    dialog.destroy()
                except:
                    pass
            
            dialog.protocol("WM_DELETE_WINDOW", on_close)
            
            # Nhúng form sửa vào dialog
            form_frame = PersonnelFormFrame(dialog, self.db, personnel_id=personnel_id)
            form_frame.pack(fill=tk.BOTH, expand=True)
            
            try:
                dialog.focus_set()
            except:
                pass
            
            # Chờ đến khi dialog đóng
            root_window.wait_window(dialog)
            
            # Sau khi sửa xong, reload danh sách
            try:
                self.load_data()
                # Cố gắng giữ lại selection quân nhân vừa sửa
                if personnel_id in self.tree.get_children():
                    self.tree.selection_set(personnel_id)
                    self.tree.see(personnel_id)
                    self.selected_id = personnel_id
            except:
                pass
        except Exception as e:
            # Nếu có lỗi với modal, fallback về cơ chế cũ
            try:
                if hasattr(self.master, 'master') and hasattr(self.master.master, 'edit_personnel_id'):
                    self.master.master.edit_personnel_id = personnel_id
                    self.master.master.show_frame('edit')
                else:
                    from gui.personnel_form_frame import PersonnelFormFrame
                    self.destroy()
                    form_frame = PersonnelFormFrame(self.master, self.db, personnel_id=personnel_id)
                    form_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            except:
                import traceback
                traceback.print_exc()
    
    def view_detail(self, personnel_id: str):
        """Xem chi tiết - Giao diện sơ yếu lí lịch đẹp mắt"""
        try:
            person = self.db.get_by_id(personnel_id)
            if not person:
                messagebox.showwarning("Cảnh báo", "Không tìm thấy thông tin quân nhân")
                return
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lấy thông tin quân nhân: {str(e)}")
            import traceback
            traceback.print_exc()
            return
        
        # Lấy tên đơn vị từ unitId
        ten_don_vi = person.donVi or ''
        if person.unitId:
            try:
                unit = self.db.get_unit_by_id(person.unitId)
                if unit:
                    ten_don_vi = unit.ten
            except:
                pass
        
        # Lấy danh sách người thân
        nguoi_than_list = []
        try:
            nguoi_than_list = self.db.get_nguoi_than_by_personnel(personnel_id)
        except:
            pass
        
        # Tạo window chi tiết - Đảm bảo parent window đúng
        try:
            # Lấy root window làm parent
            root_window = self.winfo_toplevel()
            detail_window = tk.Toplevel(root_window)
            detail_window.title(f"Sơ Yếu Lí Lịch: {person.hoTen or 'Chưa có tên'}")
            
            # Full width - Lấy kích thước màn hình
            screen_width = detail_window.winfo_screenwidth()
            screen_height = detail_window.winfo_screenheight()
            # Chiếm 95% chiều rộng và 90% chiều cao màn hình
            window_width = int(screen_width * 0.95)
            window_height = int(screen_height * 0.90)
            detail_window.geometry(f"{window_width}x{window_height}")
            
            # Căn giữa cửa sổ
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            detail_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
            
            detail_window.configure(bg='#F5F5F5')
            # Đảm bảo window hiển thị trên cùng
            detail_window.transient(root_window)
            detail_window.grab_set()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tạo cửa sổ chi tiết: {str(e)}")
            import traceback
            traceback.print_exc()
            return
        
        # Header với màu đẹp
        header_frame = tk.Frame(detail_window, bg=MILITARY_COLORS['primary'], height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        header_content = tk.Frame(header_frame, bg=MILITARY_COLORS['primary'])
        header_content.pack(expand=True, fill=tk.BOTH, padx=30, pady=15)
        
        tk.Label(
            header_content,
            text=" SƠ YẾU LÍ LỊCH",
            font=('Arial', 18, 'bold'),
            bg=MILITARY_COLORS['primary'],
            fg=MILITARY_COLORS['text_light']
        ).pack()
        
        tk.Label(
            header_content,
            text=person.hoTen or 'Chưa có tên',
            font=('Arial', 14),
            bg=MILITARY_COLORS['primary'],
            fg=MILITARY_COLORS['text_light']
        ).pack(pady=(5, 0))
        
        # Scrollable content - Tối ưu cho full width, không có thanh cuộn ngang
        canvas = tk.Canvas(detail_window, bg='#F5F5F5', highlightthickness=0)
        # Chỉ có vertical scrollbar, không có horizontal scrollbar
        scrollbar = ttk.Scrollbar(detail_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#F5F5F5')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        # Chỉ cấu hình vertical scroll, không có horizontal scroll
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Biến để lưu wraplength động
        current_wraplength = [1000]  # Dùng list để có thể thay đổi trong closure
        
        # Cập nhật width của scrollable_frame khi canvas resize để tận dụng full width
        # Đảm bảo không có thanh cuộn ngang
        def update_scrollable_width(event):
            canvas_width = event.width
            if canvas_width > 1:  # Đảm bảo canvas đã được render
                # Đặt width của scrollable_frame bằng canvas width để không có horizontal scroll
                canvas.itemconfig(canvas_window, width=canvas_width)
                # Cập nhật wraplength cho các label (trừ đi padding và label width ~350px)
                current_wraplength[0] = max(400, canvas_width - 350)
                # Cập nhật tất cả value labels trong scrollable_frame
                update_all_wraplengths(scrollable_frame)
                # Cập nhật scrollregion sau khi resize
                canvas.configure(scrollregion=canvas.bbox("all"))
        
        # Helper function để cập nhật wraplength cho tất cả labels
        def update_all_wraplengths(widget):
            """Cập nhật wraplength cho tất cả Label widgets (value labels)"""
            if isinstance(widget, tk.Label):
                try:
                    # Chỉ cập nhật nếu là value label (không phải label có width=30)
                    if widget.cget('width') != '30':
                        widget.config(wraplength=current_wraplength[0])
                except:
                    pass
            # Đệ quy cho các widget con
            for child in widget.winfo_children():
                update_all_wraplengths(child)
        
        canvas.bind('<Configure>', update_scrollable_width)
        
        # Helper function để tạo section - Tối ưu cho full width
        def create_section(parent, title, bg_color='#FFFFFF'):
            """Tạo một section với tiêu đề"""
            section_frame = tk.Frame(parent, bg='#F5F5F5')
            section_frame.pack(fill=tk.X, padx=30, pady=12)  # Tăng padding cho màn hình rộng
            
            # Tiêu đề section
            title_frame = tk.Frame(section_frame, bg=bg_color, relief=tk.RAISED, bd=1)
            title_frame.pack(fill=tk.X)
            
            tk.Label(
                title_frame,
                text=title,
                font=('Arial', 13, 'bold'),  # Tăng font size một chút
                bg=bg_color,
                fg=MILITARY_COLORS['primary_dark'],
                anchor=tk.W,
                padx=20,  # Tăng padding
                pady=12
            ).pack(fill=tk.X)
            
            # Content frame
            content_frame = tk.Frame(section_frame, bg=bg_color, relief=tk.SUNKEN, bd=1)
            content_frame.pack(fill=tk.X)
            
            return content_frame
        
        # Helper function để tạo dòng thông tin - Tối ưu cho full width, không có thanh cuộn ngang
        def create_info_row(parent, label, value, row_num):
            """Tạo một dòng thông tin"""
            row_bg = '#FFFFFF' if row_num % 2 == 0 else '#FAFAFA'
            row_frame = tk.Frame(parent, bg=row_bg)
            row_frame.pack(fill=tk.X, padx=0, pady=1)
            
            label_widget = tk.Label(
                row_frame,
                text=f"{label}:",
                font=('Arial', 10, 'bold'),
                bg=row_bg,
                fg=MILITARY_COLORS['text_dark'],
                anchor=tk.W,
                width=30  # Tăng width để phù hợp với màn hình rộng
            )
            label_widget.pack(side=tk.LEFT, padx=20, pady=8)
            
            # Tính wraplength dựa trên canvas width hiện tại
            try:
                canvas_width = canvas.winfo_width()
                if canvas_width > 1:
                    wraplength_value = max(400, canvas_width - 350)
                else:
                    wraplength_value = current_wraplength[0]
            except:
                wraplength_value = current_wraplength[0]
            
            value_widget = tk.Label(
                row_frame,
                text=value or 'Chưa có thông tin',
                font=('Arial', 10),
                bg=row_bg,
                fg=MILITARY_COLORS['text_dark'] if value else MILITARY_COLORS['text_gray'],
                anchor=tk.W,
                wraplength=wraplength_value,  # Dùng wraplength để tự động wrap, tránh thanh cuộn ngang
                justify=tk.LEFT
            )
            value_widget.pack(side=tk.LEFT, padx=10, pady=8, fill=tk.X, expand=True)
            
            return row_frame
        
        # Section 1: Thông tin cá nhân
        section1 = create_section(scrollable_frame, "1. THÔNG TIN CÁ NHÂN")
        row = 0
        create_info_row(section1, "Họ và tên khai sinh", person.hoTen, row); row += 1
        if person.hoTenThuongDung:
            create_info_row(section1, "Họ và tên thường dùng", person.hoTenThuongDung, row); row += 1
        create_info_row(section1, "Ngày tháng năm sinh", person.ngaySinh, row); row += 1
        create_info_row(section1, "Cấp bậc", person.capBac, row); row += 1
        if person.ngayNhanCapBac:
            create_info_row(section1, "Ngày nhận cấp bậc", person.ngayNhanCapBac, row); row += 1
        create_info_row(section1, "Chức vụ", person.chucVu, row); row += 1
        if person.ngayNhanChucVu:
            create_info_row(section1, "Ngày nhận chức vụ", person.ngayNhanChucVu, row); row += 1
        create_info_row(section1, "Đơn vị đang làm nhiệm vụ", ten_don_vi, row); row += 1
        create_info_row(section1, "Nhập ngũ", person.nhapNgu, row); row += 1
        if person.xuatNgu:
            create_info_row(section1, "Xuất ngũ", person.xuatNgu, row); row += 1
        
        # Section 2: Thông tin quê quán, trú quán
        section2 = create_section(scrollable_frame, "2. QUÊ QUÁN, TRÚ QUÁN")
        row = 0
        create_info_row(section2, "Quê quán", person.queQuan, row); row += 1
        create_info_row(section2, "Trú quán", person.truQuan, row); row += 1
        create_info_row(section2, "Dân tộc", person.danToc, row); row += 1
        create_info_row(section2, "Tôn giáo", person.tonGiao, row); row += 1
        create_info_row(section2, "Trình độ văn hóa", person.trinhDoVanHoa, row); row += 1
        if person.thanhPhanGiaDinh:
            create_info_row(section2, "Thành phần gia đình", person.thanhPhanGiaDinh, row); row += 1
        
        # Section 3: Thông tin đào tạo
        if person.quaTruong or person.nganhHoc or person.capHoc:
            section3 = create_section(scrollable_frame, "3. THÔNG TIN ĐÀO TẠO")
            row = 0
            if person.quaTruong:
                create_info_row(section3, "Qua trường", person.quaTruong, row); row += 1
            if person.nganhHoc:
                create_info_row(section3, "Ngành học", person.nganhHoc, row); row += 1
            if person.capHoc:
                create_info_row(section3, "Cấp học", person.capHoc, row); row += 1
            if person.thoiGianDaoTao:
                create_info_row(section3, "Thời gian đào tạo", person.thoiGianDaoTao, row); row += 1
        
        # Section 4: Thông tin liên hệ
        if person.lienHeKhiCan or person.soDienThoaiLienHe:
            section4 = create_section(scrollable_frame, "4. THÔNG TIN LIÊN HỆ")
            row = 0
            if person.lienHeKhiCan:
                create_info_row(section4, "Khi cần báo tin cho ai", person.lienHeKhiCan, row); row += 1
            if person.soDienThoaiLienHe:
                create_info_row(section4, "Số điện thoại liên hệ", person.soDienThoaiLienHe, row); row += 1
        
        # Section 5: Thông tin gia đình
        if person.hoTenCha or person.hoTenMe or person.hoTenVo:
            section5 = create_section(scrollable_frame, "5. THÔNG TIN GIA ĐÌNH")
            row = 0
            if person.hoTenCha:
                create_info_row(section5, "Họ tên cha", person.hoTenCha, row); row += 1
            if person.hoTenMe:
                create_info_row(section5, "Họ tên mẹ", person.hoTenMe, row); row += 1
            if person.hoTenVo:
                create_info_row(section5, "Họ tên vợ", person.hoTenVo, row); row += 1
        
        # Section 6: Thông tin Đảng
        if person.thongTinKhac.dang.ngayVao or person.thongTinKhac.dang.ngayChinhThuc or person.thongTinKhac.dang.chucVuDang:
            section6 = create_section(scrollable_frame, "6. THÔNG TIN ĐẢNG")
            row = 0
            if person.thongTinKhac.dang.ngayVao:
                create_info_row(section6, "Ngày vào Đảng", person.thongTinKhac.dang.ngayVao, row); row += 1
            if person.thongTinKhac.dang.ngayChinhThuc:
                create_info_row(section6, "Ngày chính thức", person.thongTinKhac.dang.ngayChinhThuc, row); row += 1
            if person.thongTinKhac.dang.chucVuDang:
                create_info_row(section6, "Chức vụ Đảng", person.thongTinKhac.dang.chucVuDang, row); row += 1
        
        # Section 7: Thông tin Đoàn
        if person.thongTinKhac.doan.ngayVao or person.thongTinKhac.doan.chucVuDoan:
            section7 = create_section(scrollable_frame, "7. THÔNG TIN ĐOÀN")
            row = 0
            if person.thongTinKhac.doan.ngayVao:
                create_info_row(section7, "Ngày vào Đoàn", person.thongTinKhac.doan.ngayVao, row); row += 1
            if person.thongTinKhac.doan.chucVuDoan:
                create_info_row(section7, "Chức vụ Đoàn", person.thongTinKhac.doan.chucVuDoan, row); row += 1
        
        # Section 8: Thông tin người thân
        if nguoi_than_list:
            section8 = create_section(scrollable_frame, "8. THÔNG TIN NGƯỜI THÂN")
            row_counter = 0
            for idx, nt in enumerate(nguoi_than_list):
                info_text = f"{nt.hoTen or ''}"
                if nt.ngaySinh:
                    info_text += f" (Sinh: {nt.ngaySinh})"
                if nt.moiQuanHe:
                    info_text += f" - {nt.moiQuanHe}"
                create_info_row(section8, f"Người thân {idx + 1}", info_text, row_counter)
                row_counter += 1
                if nt.diaChi:
                    create_info_row(section8, f"  → Địa chỉ", nt.diaChi, row_counter)
                    row_counter += 1
                if nt.soDienThoai:
                    create_info_row(section8, f"  → SĐT", nt.soDienThoai, row_counter)
                    row_counter += 1
                if nt.noiDung:
                    create_info_row(section8, f"  → Nội dung", nt.noiDung, row_counter)
                    row_counter += 1
        
        # Section 9: Thông tin khác
        if person.thongTinKhac.cdCu or person.thongTinKhac.yeuToNN or person.ghiChu:
            section9 = create_section(scrollable_frame, "9. THÔNG TIN KHÁC")
            row = 0
            create_info_row(section9, "Chế độ cũ", 'Có' if person.thongTinKhac.cdCu else 'Không', row); row += 1
            create_info_row(section9, "Yếu tố nước ngoài", 'Có' if person.thongTinKhac.yeuToNN else 'Không', row); row += 1
            if person.ghiChu:
                create_info_row(section9, "Ghi chú", person.ghiChu, row); row += 1
        
        # Pack canvas và scrollbar trực tiếp vào detail_window - Đảm bảo nội dung hiển thị
        # Phải pack canvas TRƯỚC footer để nội dung hiển thị đúng
        canvas.pack(side="left", fill="both", expand=True, padx=0, pady=0)
        scrollbar.pack(side="right", fill="y")
        
        # Footer với nút đóng - Pack sau canvas
        footer_frame = tk.Frame(detail_window, bg='#FFFFFF', relief=tk.RAISED, bd=1)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        close_btn = tk.Button(
            footer_frame,
            text="✖️ Đóng",
            command=detail_window.destroy,
            **get_button_style('secondary'),
            width=15
        )
        close_btn.pack(pady=10)
        
        # Update để đảm bảo scrollbar hoạt động đúng và nội dung hiển thị
        detail_window.update_idletasks()
        
        # Đảm bảo scrollregion được cập nhật sau khi tất cả widgets được tạo
        def update_scrollregion():
            try:
                canvas.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))
            except:
                pass
        
        # Cập nhật scrollregion sau một chút delay để đảm bảo tất cả widgets đã được render
        detail_window.after(100, update_scrollregion)
        detail_window.after(200, update_scrollregion)
        
        # Bind mouse wheel để scroll - Cuộn bằng chuột
        def on_mousewheel(event):
            """Cuộn bằng chuột wheel"""
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        # Bind mouse wheel cho canvas và scrollable_frame
        canvas.bind("<MouseWheel>", on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", on_mousewheel)
        
        # Bind cho toàn bộ window
        def on_window_mousewheel(event):
            """Cuộn khi di chuột vào window"""
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        detail_window.bind("<MouseWheel>", on_window_mousewheel)
    
    def delete_personnel(self, personnel_id: str):
        """Xóa quân nhân"""
        person = self.db.get_by_id(personnel_id)
        if not person:
            return
        
        if messagebox.askyesno("Xác nhận", f"Bạn có chắc muốn xóa {person.hoTen}?"):
            if self.db.delete(personnel_id):
                messagebox.showinfo("Thành công", "Đã xóa quân nhân")
                self.load_data()
                self.selected_id = None
            else:
                messagebox.showerror("Lỗi", "Không thể xóa quân nhân")
    
    def export_csv(self):
        """Xuất CSV"""
        if not self.personnel_list:
            messagebox.showinfo("Thông báo", "Chưa có dữ liệu để xuất")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Xuất file CSV"
        )
        
        if file_path:
            try:
                csv_data = ExportService.to_csv(self.personnel_list)
                with open(file_path, 'w', encoding='utf-8-sig') as f:
                    f.write(csv_data)
                messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xuất file:\n{str(e)}")
    
    def export_word(self):
        """Xuất Word với bản xem trước"""
        # Lọc theo dân tộc thiểu số
        filtered_list = ExportService.filter_ethnic_minority(self.personnel_list)
        
        if not filtered_list:
            messagebox.showinfo("Thông báo", "Không có quân nhân nào là người đồng bào dân tộc thiểu số")
            return
        
        # Dialog nhập thông tin
        dialog = tk.Toplevel(self)
        dialog.title("Xuất file Word - Danh sách dân tộc thiểu số")
        dialog.geometry("500x400")
        dialog.configure(bg='#FAFAFA')
        
        tk.Label(
            dialog,
            text="📄 XUẤT FILE WORD",
            font=('Arial', 14, 'bold'),
            bg='#FAFAFA',
            fg='#388E3C'
        ).pack(pady=10)
        
        # Form nhập thông tin
        form_frame = tk.Frame(dialog, bg='#FAFAFA')
        form_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        tk.Label(form_frame, text="Tiểu đoàn:", bg='#FAFAFA', font=('Arial', 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        tieu_doan_var = tk.StringVar(value="TIỂU ĐOÀN 38")
        tk.Entry(form_frame, textvariable=tieu_doan_var, width=30, font=('Arial', 10)).grid(row=0, column=1, pady=5, padx=10)
        
        tk.Label(form_frame, text="Đại đội:", bg='#FAFAFA', font=('Arial', 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        dai_doi_var = tk.StringVar(value="ĐẠI ĐỘI 3")
        tk.Entry(form_frame, textvariable=dai_doi_var, width=30, font=('Arial', 10)).grid(row=1, column=1, pady=5, padx=10)
        
        tk.Label(form_frame, text="Địa điểm:", bg='#FAFAFA', font=('Arial', 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        dia_diem_var = tk.StringVar(value="Đắk Lắk")
        tk.Entry(form_frame, textvariable=dia_diem_var, width=30, font=('Arial', 10)).grid(row=2, column=1, pady=5, padx=10)
        
        tk.Label(form_frame, text="Chính trị viên:", bg='#FAFAFA', font=('Arial', 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
        chinh_tri_vien_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=chinh_tri_vien_var, width=30, font=('Arial', 10)).grid(row=3, column=1, pady=5, padx=10)
        
        # Thông tin xem trước
        preview_frame = tk.Frame(dialog, bg='#FFFFFF', relief=tk.SOLID, bd=1)
        preview_frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        tk.Label(
            preview_frame,
            text=f"📊 Xem trước:\nSố lượng quân nhân: {len(filtered_list)}",
            bg='#FFFFFF',
            font=('Arial', 10),
            justify=tk.LEFT
        ).pack(padx=10, pady=10, anchor=tk.W)
        
        # Danh sách dân tộc (lấy từ database, loại trừ Kinh)
        ethnic_list = {}
        for person in filtered_list:
            dan_toc = (person.danToc or '').strip()
            if dan_toc and dan_toc.lower() not in ['kinh', 'việt', 'việt nam']:
                ethnic_list[dan_toc] = ethnic_list.get(dan_toc, 0) + 1
        
        ethnic_text = "Dân tộc trong danh sách:\n"
        if ethnic_list:
            for ethnic, count in sorted(ethnic_list.items()):
                ethnic_text += f"  • {ethnic}: {count}\n"
        else:
            ethnic_text += "  (Chưa có dữ liệu)"
        
        tk.Label(
            preview_frame,
            text=ethnic_text,
            bg='#FFFFFF',
            font=('Arial', 9),
            justify=tk.LEFT
        ).pack(padx=10, pady=5, anchor=tk.W)
        
        def do_export():
            try:
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".docx",
                    filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
                    title="Lưu file Word",
                    initialfile=f"Danh_sach_dan_toc_thieu_so_{datetime.now().strftime('%Y%m%d')}.docx"
                )
                
                if file_path:
                    word_data = ExportService.to_word_docx(
                        filtered_list,
                        tieu_doan_var.get(),
                        dai_doi_var.get(),
                        dia_diem_var.get(),
                        chinh_tri_vien_var.get()
                    )
                    
                    with open(file_path, 'wb') as f:
                        f.write(word_data)
                    
                    messagebox.showinfo("Thành công", f"Đã xuất file:\n{file_path}\n\nSố lượng: {len(filtered_list)} quân nhân")
                    dialog.destroy()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xuất file:\n{str(e)}")
        
        # Nút xuất
        btn_frame = tk.Frame(dialog, bg='#FAFAFA')
        btn_frame.pack(pady=10)
        
        tk.Button(
            btn_frame,
            text="📄 Xuất File",
            command=do_export,
            bg='#4CAF50',
            fg='white',
            font=('Arial', 11, 'bold'),
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="Hủy",
            command=dialog.destroy,
            bg='#CCCCCC',
            fg='black',
            font=('Arial', 11),
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)