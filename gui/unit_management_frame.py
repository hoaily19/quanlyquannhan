"""
Frame quản lý đơn vị (Đại đội, Trung đội, Xe...)
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from models.unit import Unit
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
        
        # Toolbar - TẤT CẢ NÚT XẾP NGANG 1 HÀNG
        toolbar = tk.Frame(self, bg=self.bg_color, pady=10)
        toolbar.pack(fill=tk.X, padx=10)
        
        # Nhóm 1: Tạo đơn vị
        tk.Button(
            toolbar,
            text="➕ Tạo Đại Đội",
            command=lambda: self.create_unit('dai_doi'),
            font=('Segoe UI', 10),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        tk.Button(
            toolbar,
            text="➕ Tạo Trung Đội",
            command=lambda: self.create_unit('trung_doi'),
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        tk.Button(
            toolbar,
            text="➕ Tạo Xe",
            command=lambda: self.create_unit('xe'),
            font=('Segoe UI', 10),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        tk.Button(
            toolbar,
            text="➕ Tạo Tổ",
            command=lambda: self.create_unit('to'),
            font=('Segoe UI', 10),
            bg='#9C27B0',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Separator (khoảng trống)
        tk.Frame(toolbar, width=10, bg=self.bg_color).pack(side=tk.LEFT)
        
        # Nhóm 2: Thao tác đơn vị
        tk.Button(
            toolbar,
            text="✏️ Sửa",
            command=self.edit_unit,
            font=('Segoe UI', 10, 'bold'),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        tk.Button(
            toolbar,
            text="🗑️ Xóa",
            command=self.delete_unit,
            font=('Segoe UI', 10, 'bold'),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Nút Thêm Đơn Vị Con với menu
        add_child_menu = tk.Menubutton(
            toolbar,
            text="➕ Thêm Đơn Vị Con",
            font=('Segoe UI', 10),
            bg='#9C27B0',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2',
            direction='below'
        )
        add_child_menu.pack(side=tk.LEFT, padx=3)
        
        add_child_dropdown = tk.Menu(add_child_menu, tearoff=0)
        add_child_menu.config(menu=add_child_dropdown)
        
        add_child_dropdown.add_command(
            label="➕ Thêm Tổ",
            command=self.add_to_to_unit
        )
        add_child_dropdown.add_command(
            label="➕ Thêm Xe",
            command=lambda: self.add_child_unit('xe')
        )
        add_child_dropdown.add_command(
            label="➕ Thêm Trung Đội",
            command=lambda: self.add_child_unit('trung_doi')
        )
        
        # Separator
        tk.Frame(toolbar, width=10, bg=self.bg_color).pack(side=tk.LEFT)
        
        # Nhóm 3: Quản lý quân nhân
        tk.Button(
            toolbar,
            text="👥 Quản Lý Quân Nhân",
            command=self.manage_personnel,
            font=('Segoe UI', 10),
            bg='#388E3C',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        tk.Button(
            toolbar,
            text="🔄 Chuyển Tổ",
            command=self.move_personnel_to_other_unit,
            font=('Segoe UI', 10),
            bg='#FF5722',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Separator
        tk.Frame(toolbar, width=10, bg=self.bg_color).pack(side=tk.LEFT)
        
        # Nhóm 4: Xuất file
        tk.Button(
            toolbar,
            text="📄 Xuất Word",
            command=self.export_word,
            font=('Segoe UI', 10),
            bg='#2196F3',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
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
        tree = ttk.Treeview(tree_frame, columns=columns, show='tree headings', height=20)
        
        # Cột tree (#0) để hiển thị cây phân cấp
        tree.heading('#0', text='')
        tree.column('#0', width=20, stretch=False)  # Đủ rộng để hiển thị icon expand/collapse
        
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
        
        # Toolbar cho panel quân nhân - ĐẶT TRƯỚC để hiển thị ở trên cùng
        personnel_toolbar = tk.Frame(right_frame, bg=self.bg_color, pady=5)
        personnel_toolbar.pack(fill=tk.X, padx=5, pady=5, side=tk.TOP)
        
        # Nút Chọn Quân Nhân
        self.add_personnel_btn = tk.Button(
            personnel_toolbar,
            text="➕ Chọn Quân Nhân",
            command=self.manage_personnel,
            font=('Segoe UI', 10, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2',
            state=tk.DISABLED  # Disabled cho đến khi chọn đơn vị
        )
        self.add_personnel_btn.pack(side=tk.LEFT, padx=5)
        
        # Nút Xóa Quân Nhân
        self.remove_personnel_btn = tk.Button(
            personnel_toolbar,
            text="➖ Xóa Khỏi Đơn Vị",
            command=self.remove_personnel_from_unit,
            font=('Segoe UI', 10),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2',
            state=tk.DISABLED  # Disabled cho đến khi chọn đơn vị
        )
        self.remove_personnel_btn.pack(side=tk.LEFT, padx=5)
        
        # Treeview frame cho danh sách quân nhân - ĐẶT SAU toolbar
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
        
        # Label thông báo (hiển thị trong frame thay vì dialog)
        self.status_label = tk.Label(
            self,
            text="",
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg='#4CAF50',
            wraplength=800
        )
        self.status_label.pack(pady=5, padx=10)
        
    
    def _count_personnel_recursive(self, unit_id, child_map):
        """Đếm số quân nhân trong đơn vị và tất cả đơn vị con (đệ quy)"""
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            return 0
        
        count = len(unit.personnelIds) if unit.personnelIds else 0
        
        # Cộng dồn từ các đơn vị con
        if unit_id in child_map:
            for child_unit in child_map[unit_id]:
                count += self._count_personnel_recursive(child_unit.id, child_map)
        
        return count
    
    def load_units(self):
        """Load danh sách đơn vị với cây phân cấp (đại đội/trung đội -> tổ, xe, trung đội...)"""
        # Xóa dữ liệu cũ
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Load từ database
        try:
            all_units = self.db.get_all_units()
            
            # Tách đơn vị cha (đại đội, trung đội) và tất cả đơn vị con (tổ, xe, trung đội...)
            parent_units = [u for u in all_units if u.loai in ['dai_doi', 'trung_doi'] and not u.parentId]
            # Load TẤT CẢ đơn vị con (không chỉ tổ, mà cả xe, trung đội con...)
            child_units = [u for u in all_units if u.parentId]  # Tất cả đơn vị có parentId
            
            # Tạo dictionary để map parentId -> danh sách đơn vị con (tất cả loại)
            child_map = {}
            for child in child_units:
                if child.parentId not in child_map:
                    child_map[child.parentId] = []
                child_map[child.parentId].append(child)
            
            # Sắp xếp đơn vị con theo loại: trung_doi -> xe -> to
            def sort_key(unit):
                order = {'trung_doi': 0, 'xe': 1, 'to': 2}
                return order.get(unit.loai, 99)
            
            # Load đơn vị cha và tất cả đơn vị con
            stt = 1
            for parent_unit in parent_units:
                # Tính tổng số quân nhân (bao gồm cả đơn vị con)
                total_personnel = self._count_personnel_recursive(parent_unit.id, child_map)
                
                # Thêm đơn vị cha
                parent_item = self.tree.insert('', tk.END, iid=parent_unit.id, 
                    text='',  # Không hiển thị text trong cột tree, chỉ icon
                    values=(
                        stt,
                        parent_unit.ten,
                        self._get_loai_name(parent_unit.loai),
                        total_personnel,  # Sử dụng tổng số quân nhân (cộng dồn)
                        parent_unit.ghiChu or ''
                    ))
                stt += 1
                
                # Thêm tất cả đơn vị con nếu có (tổ, xe, trung đội...)
                if parent_unit.id in child_map:
                    # Sắp xếp đơn vị con theo loại
                    sorted_children = sorted(child_map[parent_unit.id], key=sort_key)
                    for child_unit in sorted_children:
                        # Thêm đơn vị con dưới đơn vị cha
                        self.tree.insert(parent_item, tk.END, iid=child_unit.id,
                            text='',  # Không hiển thị text trong cột tree
                            values=(
                                '',  # STT để trống cho đơn vị con
                                child_unit.ten,
                                self._get_loai_name(child_unit.loai),
                                len(child_unit.personnelIds),
                                child_unit.ghiChu or ''
                            ))
                    # Tự động mở rộng (expand) đơn vị cha để hiển thị đơn vị con
                    self.tree.item(parent_item, open=True)
            
            # Load các đơn vị không có parent và không phải đại đội/trung đội gốc
            other_units = [u for u in all_units 
                          if u.loai not in ['dai_doi', 'trung_doi'] and not u.parentId]
            for unit in other_units:
                self.tree.insert('', tk.END, iid=unit.id, 
                    text='',
                    values=(
                        stt,
                        unit.ten,
                        self._get_loai_name(unit.loai),
                        len(unit.personnelIds),
                        unit.ghiChu or ''
                    ))
                stt += 1
            
            # Load các đơn vị con không có parent hợp lệ (orphan)
            parent_ids = [pu.id for pu in parent_units]
            orphan_children = [u for u in child_units if u.parentId not in parent_ids]
            for unit in orphan_children:
                self.tree.insert('', tk.END, iid=unit.id,
                    text='',
                    values=(
                        stt,
                        unit.ten,
                        self._get_loai_name(unit.loai),
                        len(unit.personnelIds),
                        unit.ghiChu or ''
                    ))
                stt += 1
                
        except Exception as e:
            # Nếu chưa có hàm get_all_units, hiển thị thông báo
            messagebox.showinfo("Thông báo", f"Lỗi khi load đơn vị: {str(e)}")
    
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
        
        # Enable các nút
        self.add_personnel_btn.config(state=tk.NORMAL)
        self.remove_personnel_btn.config(state=tk.NORMAL)
        
        # Xóa danh sách cũ
        for item in self.personnel_tree.get_children():
            self.personnel_tree.delete(item)
        
        # Load quân nhân trong đơn vị
        try:
            personnel_list = []
            
            # Nếu là đại đội, lấy quân nhân từ đại đội và tất cả đơn vị con
            if unit.loai == 'dai_doi':
                # Lấy quân nhân trực tiếp từ đại đội
                personnel_list.extend(self.db.get_personnel_by_unit(unit_id))
                
                # Lấy tất cả đơn vị con (trung đội, xe, tổ...)
                child_units = self.db.get_units_by_parent_id(unit_id)
                for child_unit in child_units:
                    # Lấy quân nhân từ đơn vị con
                    child_personnel = self.db.get_personnel_by_unit(child_unit.id)
                    personnel_list.extend(child_personnel)
                    
                    # Nếu đơn vị con có đơn vị con nữa (ví dụ: trung đội có tổ)
                    if child_unit.loai in ['trung_doi', 'xe']:
                        grandchild_units = self.db.get_units_by_parent_id(child_unit.id)
                        for grandchild_unit in grandchild_units:
                            grandchild_personnel = self.db.get_personnel_by_unit(grandchild_unit.id)
                            personnel_list.extend(grandchild_personnel)
            else:
                # Nếu không phải đại đội, chỉ lấy quân nhân trực tiếp
                personnel_list = self.db.get_personnel_by_unit(unit_id)
            
            # Loại bỏ trùng lặp (nếu có)
            seen_ids = set()
            unique_personnel = []
            for person in personnel_list:
                if person.id not in seen_ids:
                    seen_ids.add(person.id)
                    unique_personnel.append(person)
            personnel_list = unique_personnel
            
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
        
        # Disable các nút
        self.add_personnel_btn.config(state=tk.DISABLED)
        self.remove_personnel_btn.config(state=tk.DISABLED)
        
        self.personnel_info_label.config(
            text="👉 Chọn một đơn vị để xem danh sách quân nhân"
        )
        self.personnel_info_label.pack(pady=20)
    
    def remove_personnel_from_unit(self):
        """Xóa quân nhân khỏi đơn vị"""
        selection = self.tree.selection()
        if not selection:
            self.status_label.config(text="⚠️ Vui lòng chọn đơn vị", fg='#FF9800')
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            self.status_label.config(text="❌ Không tìm thấy đơn vị", fg='#F44336')
            return
        
        # Lấy quân nhân đã chọn trong danh sách
        personnel_selection = self.personnel_tree.selection()
        if not personnel_selection:
            self.status_label.config(text="⚠️ Vui lòng chọn quân nhân cần xóa", fg='#FF9800')
            return
        
        # Xác nhận
        from tkinter import messagebox
        person_ids = list(personnel_selection)
        if len(person_ids) == 1:
            person = self.db.get_by_id(person_ids[0])
            person_name = person.hoTen if person else "quân nhân này"
            confirm = messagebox.askyesno(
                "Xác nhận",
                f"Bạn có chắc muốn xóa {person_name} khỏi đơn vị '{unit.ten}'?"
            )
        else:
            confirm = messagebox.askyesno(
                "Xác nhận",
                f"Bạn có chắc muốn xóa {len(person_ids)} quân nhân khỏi đơn vị '{unit.ten}'?"
            )
        
        if not confirm:
            return
        
        try:
            # Cập nhật unitId của quân nhân về None
            for person_id in person_ids:
                person = self.db.get_by_id(person_id)
                if person:
                    person.unitId = None
                    self.db.update(person)
            
            # Cập nhật unit.personnelIds (nếu có)
            if hasattr(unit, 'personnelIds') and unit.personnelIds:
                unit.personnelIds = [pid for pid in unit.personnelIds if pid not in person_ids]
                self.db.update_unit(unit)
            
            self.status_label.config(
                text=f"✅ Đã xóa {len(person_ids)} quân nhân khỏi đơn vị '{unit.ten}'",
                fg='#4CAF50'
            )
            
            # Reload danh sách
            self.load_units()
            self.on_unit_select(None)
            
        except Exception as e:
            self.status_label.config(text=f"❌ Lỗi: {str(e)}", fg='#F44336')
    
    def edit_personnel_from_list(self, event=None):
        """Sửa quân nhân từ danh sách trong đơn vị"""
        selection = self.personnel_tree.selection()
        if not selection:
            self.status_label.config(text="⚠️ Vui lòng chọn quân nhân cần sửa", fg='#FF9800')
            return
        
        person_id = selection[0]
        person = self.db.get_by_id(person_id)
        if not person:
            self.status_label.config(text="❌ Không tìm thấy quân nhân", fg='#F44336')
            return
        
        # Lấy đơn vị hiện tại để reload sau khi sửa
        unit_selection = self.tree.selection()
        current_unit_id = unit_selection[0] if unit_selection else None
        
        # Mở dialog sửa quân nhân
        try:
            from gui.personnel_form_frame import PersonnelFormFrame
            
            edit_window = tk.Toplevel(self)
            edit_window.title(f"Sửa Quân Nhân - {person.hoTen or ''}")
            edit_window.geometry("800x700")
            edit_window.transient(self)
            edit_window.grab_set()
            
            edit_frame = PersonnelFormFrame(edit_window, self.db, personnel_id=person_id)
            edit_frame.pack(fill=tk.BOTH, expand=True)
            
            def on_edit_close():
                # Reload danh sách quân nhân trong đơn vị
                if current_unit_id:
                    # Xóa danh sách cũ
                    for item in self.personnel_tree.get_children():
                        self.personnel_tree.delete(item)
                    
                    # Load lại quân nhân
                    try:
                        personnel_list = self.db.get_personnel_by_unit(current_unit_id)
                        
                        if not personnel_list:
                            self.personnel_info_label.config(
                                text=f"Đơn vị '{self.db.get_unit_by_id(current_unit_id).ten if self.db.get_unit_by_id(current_unit_id) else ''}' chưa có quân nhân nào"
                            )
                            self.personnel_info_label.pack(pady=20)
                        else:
                            # Ẩn label thông báo
                            self.personnel_info_label.pack_forget()
                            
                            # Hiển thị danh sách
                            for idx, person in enumerate(personnel_list, 1):
                                self.personnel_tree.insert('', tk.END, iid=person.id, values=(
                                    idx,
                                    person.hoTen or '',
                                    person.capBac or '',
                                    person.chucVu or '',
                                    person.ngaySinh or ''
                                ))
                        
                        # Cập nhật số quân nhân trong danh sách đơn vị
                        self.load_units()
                        
                        # Chọn lại đơn vị
                        if current_unit_id in self.tree.get_children() or any(
                            current_unit_id in self.tree.get_children(item) 
                            for item in self.tree.get_children()
                        ):
                            self.tree.selection_set(current_unit_id)
                            self.tree.see(current_unit_id)
                        
                        self.status_label.config(
                            text=f"✅ Đã cập nhật thông tin quân nhân: {person.hoTen or ''}",
                            fg='#4CAF50'
                        )
                    except Exception as e:
                        self.status_label.config(text=f"❌ Lỗi khi reload: {str(e)}", fg='#F44336')
                
                edit_window.destroy()
            
            edit_window.protocol("WM_DELETE_WINDOW", on_edit_close)
            
        except ImportError:
            self.status_label.config(
                text="❌ Không thể mở form sửa quân nhân. Vui lòng kiểm tra lại.",
                fg='#F44336'
            )
        except Exception as e:
            self.status_label.config(text=f"❌ Lỗi: {str(e)}", fg='#F44336')
    
    def move_personnel_to_other_unit(self):
        """Chuyển quân nhân từ tổ này sang tổ khác"""
        # Kiểm tra quân nhân được chọn
        personnel_selection = self.personnel_tree.selection()
        if not personnel_selection:
            self.status_label.config(text="⚠️ Vui lòng chọn quân nhân cần chuyển", fg='#FF9800')
            return
        
        # Kiểm tra đơn vị hiện tại (tổ hiện tại)
        unit_selection = self.tree.selection()
        if not unit_selection:
            self.status_label.config(text="⚠️ Vui lòng chọn đơn vị (tổ) hiện tại", fg='#FF9800')
            return
        
        current_unit_id = unit_selection[0]
        current_unit = self.db.get_unit_by_id(current_unit_id)
        if not current_unit:
            self.status_label.config(text="❌ Không tìm thấy đơn vị hiện tại", fg='#F44336')
            return
        
        # Kiểm tra xem đơn vị hiện tại có phải là tổ không
        if current_unit.loai != 'to':
            self.status_label.config(text="⚠️ Chỉ có thể chuyển quân nhân giữa các tổ", fg='#FF9800')
            return
        
        # Lấy thông tin quân nhân
        person_ids = list(personnel_selection)
        personnel_list = []
        for person_id in person_ids:
            person = self.db.get_by_id(person_id)
            if person:
                personnel_list.append(person)
        
        if not personnel_list:
            self.status_label.config(text="❌ Không tìm thấy quân nhân", fg='#F44336')
            return
        
        # Mở dialog chọn tổ mới
        dialog = tk.Toplevel(self)
        dialog.title("Chuyển Quân Nhân Sang Tổ Khác")
        dialog.geometry("500x400")
        dialog.configure(bg=self.bg_color)
        dialog.transient(self)
        dialog.grab_set()
        
        # Hiển thị thông tin quân nhân sẽ chuyển
        info_frame = tk.Frame(dialog, bg=self.bg_color)
        info_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(
            info_frame,
            text="Quân nhân sẽ chuyển:",
            font=('Segoe UI', 11, 'bold'),
            bg=self.bg_color
        ).pack(anchor=tk.W)
        
        personnel_names = ", ".join([p.hoTen or f"ID: {p.id}" for p in personnel_list])
        tk.Label(
            info_frame,
            text=personnel_names,
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg='#388E3C',
            wraplength=450
        ).pack(anchor=tk.W, pady=5)
        
        tk.Label(
            info_frame,
            text=f"Từ tổ: {current_unit.ten}",
            font=('Segoe UI', 10),
            bg=self.bg_color
        ).pack(anchor=tk.W, pady=5)
        
        # Chọn đại đội mới
        tk.Label(
            dialog,
            text="Chọn đại đội:",
            font=('Segoe UI', 11, 'bold'),
            bg=self.bg_color
        ).pack(pady=(10, 5))
        
        dai_doi_var = tk.StringVar()
        dai_doi_combo = ttk.Combobox(
            dialog,
            textvariable=dai_doi_var,
            font=('Segoe UI', 11),
            width=40,
            state='readonly'
        )
        dai_doi_combo.pack(pady=5, padx=10)
        
        # Load danh sách đại đội và trung đội
        all_units = self.db.get_all_units()
        dai_doi_list = []
        for unit in all_units:
            if unit.loai in ['dai_doi', 'trung_doi']:
                dai_doi_list.append(f"{unit.id}|{unit.ten}")
        
        if not dai_doi_list:
            tk.Label(
                dialog,
                text="⚠️ Không có đại đội/trung đội nào",
                font=('Segoe UI', 10),
                bg=self.bg_color,
                fg='#FF9800'
            ).pack(pady=10)
            
            tk.Button(
                dialog,
                text="Đóng",
                command=dialog.destroy,
                font=('Segoe UI', 10),
                bg='#757575',
                fg='white',
                relief=tk.FLAT,
                padx=20,
                pady=5
            ).pack(pady=10)
            return
        
        dai_doi_combo['values'] = [opt.split('|')[1] for opt in dai_doi_list]
        if dai_doi_list:
            dai_doi_combo.current(0)
        
        # Chọn tổ mới
        tk.Label(
            dialog,
            text="Chọn tổ:",
            font=('Segoe UI', 11, 'bold'),
            bg=self.bg_color
        ).pack(pady=(10, 5))
        
        target_unit_var = tk.StringVar()
        target_combo = ttk.Combobox(
            dialog,
            textvariable=target_unit_var,
            font=('Segoe UI', 11),
            width=40,
            state='readonly'
        )
        target_combo.pack(pady=5, padx=10)
        
        # Hàm cập nhật danh sách tổ khi chọn đại đội
        def update_to_list(event=None):
            selected_dai_doi_text = dai_doi_var.get()
            if not selected_dai_doi_text:
                target_combo['values'] = []
                target_unit_var.set('')
                return
            
            # Tìm đại đội ID
            selected_dai_doi_id = None
            for opt in dai_doi_list:
                if opt.split('|')[1] == selected_dai_doi_text:
                    selected_dai_doi_id = opt.split('|')[0]
                    break
            
            if not selected_dai_doi_id:
                target_combo['values'] = []
                target_unit_var.set('')
                return
            
            # Load các tổ trong đại đội này (loại trừ tổ hiện tại)
            target_units = []
            for unit in all_units:
                if unit.loai == 'to' and unit.parentId == selected_dai_doi_id and unit.id != current_unit_id:
                    target_units.append(f"{unit.id}|{unit.ten}")
            
            if target_units:
                target_combo['values'] = [opt.split('|')[1] for opt in target_units]
                target_combo.current(0)
            else:
                target_combo['values'] = []
                target_unit_var.set('')
        
        dai_doi_combo.bind('<<ComboboxSelected>>', update_to_list)
        
        # Load tổ ban đầu cho đại đội đầu tiên
        update_to_list()
        
        # Label thông báo
        status_label_dialog = tk.Label(
            dialog,
            text="",
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg='#4CAF50',
            wraplength=450
        )
        status_label_dialog.pack(pady=5)
        
        def save():
            selected_dai_doi_text = dai_doi_var.get()
            if not selected_dai_doi_text:
                status_label_dialog.config(text="⚠️ Vui lòng chọn đại đội", fg='#FF9800')
                return
            
            selected_to_text = target_unit_var.get()
            if not selected_to_text:
                status_label_dialog.config(text="⚠️ Vui lòng chọn tổ", fg='#FF9800')
                return
            
            # Tìm đại đội ID
            selected_dai_doi_id = None
            for opt in dai_doi_list:
                if opt.split('|')[1] == selected_dai_doi_text:
                    selected_dai_doi_id = opt.split('|')[0]
                    break
            
            if not selected_dai_doi_id:
                status_label_dialog.config(text="❌ Không tìm thấy đại đội", fg='#F44336')
                return
            
            # Tìm tổ trong đại đội đã chọn
            target_units = []
            for unit in all_units:
                if unit.loai == 'to' and unit.parentId == selected_dai_doi_id and unit.id != current_unit_id:
                    target_units.append(f"{unit.id}|{unit.ten}")
            
            # Tìm target_unit_id
            target_unit_id = None
            for opt in target_units:
                if opt.split('|')[1] == selected_to_text:
                    target_unit_id = opt.split('|')[0]
                    break
            
            if not target_unit_id:
                status_label_dialog.config(text="❌ Không tìm thấy tổ mới", fg='#F44336')
                return
            
            target_unit = self.db.get_unit_by_id(target_unit_id)
            if not target_unit:
                status_label_dialog.config(text="❌ Không tìm thấy tổ mới", fg='#F44336')
                return
            
            try:
                # Xóa quân nhân khỏi tổ cũ
                current_unit.personnelIds = [pid for pid in current_unit.personnelIds if pid not in person_ids]
                self.db.update_unit(current_unit)
                
                # Thêm quân nhân vào tổ mới
                target_unit.personnelIds = list(set(target_unit.personnelIds + person_ids))
                self.db.update_unit(target_unit)
                
                # Cập nhật unitId cho quân nhân
                for person_id in person_ids:
                    person = self.db.get_by_id(person_id)
                    if person:
                        person.unitId = target_unit_id
                        self.db.update(person)
                
                # Reload danh sách
                self.load_units()
                
                # Chọn lại tổ cũ để xem danh sách đã cập nhật
                if current_unit_id in self.tree.get_children() or any(
                    current_unit_id in self.tree.get_children(item) 
                    for item in self.tree.get_children()
                ):
                    self.tree.selection_set(current_unit_id)
                    self.on_unit_select(None)
                
                personnel_names_str = ", ".join([p.hoTen or f"ID: {p.id}" for p in personnel_list])
                self.status_label.config(
                    text=f"✅ Đã chuyển {len(personnel_list)} quân nhân ({personnel_names_str}) từ '{current_unit.ten}' sang '{target_unit.ten}'",
                    fg='#4CAF50'
                )
                
                dialog.destroy()
            except Exception as e:
                status_label_dialog.config(text=f"❌ Lỗi: {str(e)}", fg='#F44336')
        
        # Buttons
        btn_frame = tk.Frame(dialog, bg=self.bg_color)
        btn_frame.pack(pady=10)
        
        tk.Button(
            btn_frame,
            text="💾 Chuyển",
            command=save,
            font=('Segoe UI', 10, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
    
    def _get_loai_name(self, loai: str) -> str:
        """Chuyển loại sang tên hiển thị"""
        mapping = {
            'dai_doi': 'Đại Đội',
            'trung_doi': 'Trung Đội',
            'xe': 'Xe',
            'to': 'Tổ'
        }
        return mapping.get(loai, loai)
    
    def create_unit(self, loai: str, parent_id: str = None):
        """Tạo đơn vị mới"""
        dialog = tk.Toplevel(self)
        dialog.title(f"Tạo {self._get_loai_name(loai)}")
        dialog.geometry("400x350" if loai == 'to' else "400x250")
        dialog.configure(bg=self.bg_color)
        dialog.transient(self)  # Đảm bảo dialog không bị ẩn khi click vào parent
        dialog.grab_set()  # Modal dialog
        
        tk.Label(
            dialog,
            text=f"Tên {self._get_loai_name(loai)}:",
            font=('Segoe UI', 11),
            bg=self.bg_color
        ).pack(pady=10)
        
        name_entry = tk.Entry(dialog, font=('Segoe UI', 11), width=30)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        parent_var = None
        parent_combo = None
        parent_options = []
        
        # Nếu tạo đơn vị con (tổ, xe, trung đội), thêm combobox chọn đơn vị cha
        if loai in ['to', 'xe', 'trung_doi']:
            tk.Label(
                dialog,
                text="Thuộc đơn vị:",
                font=('Segoe UI', 11),
                bg=self.bg_color
            ).pack(pady=(10, 5))
            
            parent_var = tk.StringVar()
            parent_combo = ttk.Combobox(
                dialog,
                textvariable=parent_var,
                font=('Segoe UI', 11),
                width=28,
                state='readonly'
            )
            parent_combo.pack(pady=5)
            
            # Load danh sách đơn vị cha phù hợp
            all_units = self.db.get_all_units()
            for unit in all_units:
                # Tổ có thể thuộc đại đội, trung đội, hoặc xe
                if loai == 'to' and unit.loai in ['dai_doi', 'trung_doi', 'xe']:
                    parent_options.append(f"{unit.id}|{unit.ten}")
                # Xe và trung đội chỉ có thể thuộc đại đội
                elif loai in ['xe', 'trung_doi'] and unit.loai == 'dai_doi':
                    parent_options.append(f"{unit.id}|{unit.ten}")
            
            if parent_id:
                # Nếu có parent_id, tìm và set
                for opt in parent_options:
                    if opt.startswith(parent_id + "|"):
                        parent_var.set(opt.split('|')[1])
                        break
            else:
                # Nếu không có parent_id, set option đầu tiên nếu có
                if parent_options:
                    parent_var.set(parent_options[0].split('|')[1])
            
            parent_combo['values'] = [opt.split('|')[1] for opt in parent_options]
            if parent_options:
                parent_combo.current(0)
        
        # Label thông báo (sẽ được cập nhật khi lưu thành công)
        status_label = tk.Label(
            dialog,
            text="",
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg='#4CAF50',
            wraplength=350
        )
        status_label.pack(pady=5)
        
        def save():
            ten = name_entry.get().strip()
            if not ten:
                status_label.config(text="⚠️ Vui lòng nhập tên đơn vị", fg='#FF9800')
                return
            
            # Lấy parent_id nếu là đơn vị con
            selected_parent_id = None
            if loai in ['to', 'xe', 'trung_doi'] and parent_var:
                selected_text = parent_var.get()
                if selected_text:
                    # Tìm parent_id từ selected_text
                    for opt in parent_options:
                        if opt.split('|')[1] == selected_text:
                            selected_parent_id = opt.split('|')[0]
                            break
                    if not selected_parent_id:
                        status_label.config(text="⚠️ Vui lòng chọn đơn vị cha", fg='#FF9800')
                        return
            
            try:
                unit = Unit(ten=ten, loai=loai, parentId=selected_parent_id)
                created_unit_id = self.db.create_unit(unit)
                
                # Debug: In ra thông tin để kiểm tra
                print(f"DEBUG: Đã tạo unit - ID: {created_unit_id}, Tên: {ten}, Loại: {loai}, ParentID: {selected_parent_id}")
                
                # Reload danh sách để hiển thị tổ con nếu có
                self.load_units()
                
                # Nếu là tổ và có parent, tự động expand đơn vị cha trong tree
                if loai == 'to' and selected_parent_id:
                    try:
                        # Tìm và expand đơn vị cha
                        if selected_parent_id in self.tree.get_children():
                            self.tree.item(selected_parent_id, open=True)
                            # Chọn đơn vị cha để highlight
                            self.tree.selection_set(selected_parent_id)
                            self.tree.see(selected_parent_id)
                    except Exception as e:
                        print(f"DEBUG: Lỗi khi expand parent: {e}")
                
                # Hiển thị thông báo khác tùy loại đơn vị
                if loai in ['dai_doi', 'trung_doi']:
                    # Kiểm tra xem đơn vị vừa tạo có tổ con không
                    child_units = self.db.get_units_by_parent_id(created_unit_id)
                    if child_units:
                        # Có tổ con, hiển thị thông báo với tên các tổ
                        to_names = ", ".join([to.ten for to in child_units])
                        status_label.config(
                            text=f"✅ Đã tạo {self._get_loai_name(loai)}: {ten}\n📋 Các tổ trong đơn vị: {to_names}",
                            fg='#4CAF50'
                        )
                    else:
                        # Chưa có tổ, gợi ý thêm tổ
                        status_label.config(
                            text=f"✅ Đã tạo {self._get_loai_name(loai)}: {ten}\n💡 Bạn có thể thêm tổ bằng nút '➕ Thêm Tổ'",
                            fg='#4CAF50'
                        )
                elif loai == 'to' and selected_parent_id:
                    # Tạo tổ, hiển thị tên đơn vị cha
                    parent_unit = self.db.get_unit_by_id(selected_parent_id)
                    parent_name = parent_unit.ten if parent_unit else "đơn vị cha"
                    status_label.config(
                        text=f"✅ Đã tạo {self._get_loai_name(loai)}: {ten}\n📁 Thuộc: {parent_name}",
                        fg='#4CAF50'
                    )
                else:
                    # Tạo đơn vị khác
                    status_label.config(
                        text=f"✅ Đã tạo {self._get_loai_name(loai)}: {ten}",
                        fg='#4CAF50'
                    )
                
                # Xóa nội dung input để có thể tạo tiếp
                name_entry.delete(0, tk.END)
                name_entry.focus()
                
            except Exception as e:
                status_label.config(text=f"❌ Lỗi: {str(e)}", fg='#F44336')
        
        # Frame chứa các nút
        btn_frame = tk.Frame(dialog, bg=self.bg_color)
        btn_frame.pack(pady=10)
        
        tk.Button(
            btn_frame,
            text="💾 Lưu",
            command=save,
            font=('Segoe UI', 10, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="❌ Đóng",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        name_entry.bind('<Return>', lambda e: save())
    
    def edit_unit(self):
        """Sửa đơn vị"""
        selection = self.tree.selection()
        if not selection:
            self.status_label.config(text="⚠️ Vui lòng chọn đơn vị cần sửa", fg='#FF9800')
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            self.status_label.config(text="❌ Không tìm thấy đơn vị", fg='#F44336')
            return
        
        # Mở dialog sửa đơn vị
        dialog = tk.Toplevel(self)
        dialog.title(f"Sửa {self._get_loai_name(unit.loai)}")
        dialog.geometry("400x400" if unit.parentId else "400x300")
        dialog.configure(bg=self.bg_color)
        dialog.transient(self)
        dialog.grab_set()
        
        tk.Label(
            dialog,
            text=f"Tên {self._get_loai_name(unit.loai)}:",
            font=('Segoe UI', 11),
            bg=self.bg_color
        ).pack(pady=10)
        
        name_entry = tk.Entry(dialog, font=('Segoe UI', 11), width=30)
        name_entry.insert(0, unit.ten)
        name_entry.pack(pady=5)
        name_entry.focus()
        name_entry.select_range(0, tk.END)
        
        # Nếu là đơn vị con, cho phép thay đổi đơn vị cha
        parent_var = None
        parent_combo = None
        parent_options = []
        if unit.parentId:
            tk.Label(
                dialog,
                text="Thuộc đơn vị:",
                font=('Segoe UI', 11),
                bg=self.bg_color
            ).pack(pady=(10, 5))
            
            parent_var = tk.StringVar()
            parent_combo = ttk.Combobox(
                dialog,
                textvariable=parent_var,
                font=('Segoe UI', 11),
                width=28,
                state='readonly'
            )
            parent_combo.pack(pady=5)
            
            # Load danh sách đơn vị cha phù hợp
            all_units = self.db.get_all_units()
            for u in all_units:
                # Tổ có thể thuộc đại đội, trung đội, hoặc xe
                if unit.loai == 'to' and u.loai in ['dai_doi', 'trung_doi', 'xe']:
                    parent_options.append(f"{u.id}|{u.ten}")
                # Xe và trung đội chỉ có thể thuộc đại đội
                elif unit.loai in ['xe', 'trung_doi'] and u.loai == 'dai_doi':
                    parent_options.append(f"{u.id}|{u.ten}")
            
            # Set giá trị hiện tại
            if unit.parentId:
                for opt in parent_options:
                    if opt.startswith(unit.parentId + "|"):
                        parent_var.set(opt.split('|')[1])
                        break
            
            parent_combo['values'] = [opt.split('|')[1] for opt in parent_options]
            if parent_options and not parent_var.get():
                parent_combo.current(0)
        
        # Ghi chú
        tk.Label(
            dialog,
            text="Ghi chú:",
            font=('Segoe UI', 11),
            bg=self.bg_color
        ).pack(pady=(10, 5))
        
        ghi_chu_entry = tk.Text(dialog, font=('Segoe UI', 10), width=30, height=3)
        ghi_chu_entry.insert('1.0', unit.ghiChu or '')
        ghi_chu_entry.pack(pady=5)
        
        # Label thông báo trong dialog
        status_label_dialog = tk.Label(
            dialog,
            text="",
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg='#4CAF50',
            wraplength=350
        )
        status_label_dialog.pack(pady=5)
        
        def save():
            ten = name_entry.get().strip()
            if not ten:
                status_label_dialog.config(text="⚠️ Vui lòng nhập tên đơn vị", fg='#FF9800')
                return
            
            # Lấy parent_id mới nếu có thay đổi
            new_parent_id = unit.parentId
            if parent_var and parent_combo:
                selected_text = parent_var.get()
                if selected_text:
                    for opt in parent_options:
                        if opt.split('|')[1] == selected_text:
                            new_parent_id = opt.split('|')[0]
                            break
            
            try:
                unit.ten = ten
                unit.parentId = new_parent_id
                unit.ghiChu = ghi_chu_entry.get('1.0', tk.END).strip()
                self.db.update_unit(unit)
                self.status_label.config(
                    text=f"✅ Đã sửa {self._get_loai_name(unit.loai)}: {ten}",
                    fg='#4CAF50'
                )
                self.load_units()
                # Chọn lại đơn vị đã sửa
                self.tree.selection_set(unit_id)
                self.tree.see(unit_id)
                dialog.destroy()
            except Exception as e:
                status_label_dialog.config(text=f"❌ Lỗi: {str(e)}", fg='#F44336')
        
        btn_frame = tk.Frame(dialog, bg=self.bg_color)
        btn_frame.pack(pady=10)
        
        tk.Button(
            btn_frame,
            text="💾 Lưu",
            command=save,
            font=('Segoe UI', 10, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=dialog.destroy,
            font=('Segoe UI', 10),
            bg='#757575',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        name_entry.bind('<Return>', lambda e: save())
    
    def delete_unit(self):
        """Xóa đơn vị (xóa trực tiếp không cần xác nhận)"""
        selection = self.tree.selection()
        if not selection:
            self.status_label.config(text="⚠️ Vui lòng chọn đơn vị cần xóa", fg='#FF9800')
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            self.status_label.config(text="❌ Không tìm thấy đơn vị", fg='#F44336')
            return
        
        # Kiểm tra xem đơn vị có tổ con không
        child_units = self.db.get_units_by_parent_id(unit_id)
        if child_units:
            self.status_label.config(
                text=f"⚠️ Không thể xóa đơn vị '{unit.ten}' vì có {len(child_units)} tổ con. Vui lòng xóa các tổ con trước.",
                fg='#FF9800'
            )
            return
        
        # Xóa trực tiếp
        try:
            self.db.delete_unit(unit_id)
            self.status_label.config(
                text=f"✅ Đã xóa đơn vị '{unit.ten}'",
                fg='#4CAF50'
            )
            self.load_units()
        except Exception as e:
            self.status_label.config(text=f"❌ Lỗi: {str(e)}", fg='#F44336')
    
    def manage_personnel(self):
        """Quản lý quân nhân trong đơn vị"""
        selection = self.tree.selection()
        if not selection:
            self.status_label.config(text="⚠️ Vui lòng chọn đơn vị", fg='#FF9800')
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            self.status_label.config(text="❌ Không tìm thấy đơn vị", fg='#F44336')
            return
        
        # Mở dialog chọn quân nhân
        dialog = tk.Toplevel(self)
        dialog.title(f"Quản Lý Quân Nhân - {unit.ten}")
        dialog.geometry("1100x700")
        dialog.configure(bg='#FAFAFA')
        dialog.resizable(True, True)
        dialog.transient(self)
        dialog.grab_set()
        
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
            text=f"Chọn quân nhân cho: {unit.ten}",
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
        
        # Treeview với checkbox
        tree_frame = tk.Frame(list_frame, bg='#FAFAFA')
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
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
        
        # Load tất cả quân nhân
        all_personnel = self.db.get_all()
        selected_ids = set(unit.personnelIds)
        
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
                    else:
                        tree.item(item, text='')
                        selected_ids.discard(item_id)
        
        select_all_btn = tk.Button(
            toolbar_frame,
            text="✔ Chọn Tất Cả",
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
        
        def load_tree_data():
            """Load dữ liệu vào tree với filter và sắp xếp theo cấp bậc"""
            # Xóa dữ liệu cũ
            for item in tree.get_children():
                tree.delete(item)
            
            # Lọc theo tìm kiếm
            search_text = search_var.get().lower()
            display_personnel = all_personnel
            if search_text:
                display_personnel = [p for p in all_personnel 
                                  if search_text in (p.hoTen or '').lower() or
                                     search_text in (p.capBac or '').lower() or
                                     search_text in (p.chucVu or '').lower()]
            
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            def _parse_cap_bac_rank(cap_bac: str) -> int:
                """Parse cấp bậc thành số để so sánh"""
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
            
            def sort_key(person):
                cap_bac_rank = _parse_cap_bac_rank(person.capBac or '')
                ho_ten = (person.hoTen or '').lower()
                return (-cap_bac_rank, ho_ten)
            
            display_personnel = sorted(display_personnel, key=sort_key)
            
            # Load vào tree
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
        tree.bind('<Double-1>', lambda e: edit_selected_personnel())  # Double click để sửa
        search_var.trace('w', lambda *args: load_tree_data())
        
        # Pack tree và scrollbar
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Load dữ liệu ban đầu
        load_tree_data()
        
        # Hàm sửa quân nhân
        def edit_selected_personnel():
            selection = tree.selection()
            if not selection:
                return
            
            person_id = selection[0]
            person = self.db.get_by_id(person_id)
            if not person:
                return
            
            # Mở dialog sửa quân nhân
            from gui.personnel_form_frame import PersonnelFormFrame
            edit_window = tk.Toplevel(dialog)
            edit_window.title(f"Sửa Quân Nhân - {person.hoTen or ''}")
            edit_window.geometry("800x700")
            edit_frame = PersonnelFormFrame(edit_window, self.db, personnel_id=person_id)
            edit_frame.pack(fill=tk.BOTH, expand=True)
            
            def on_edit_close():
                # Reload danh sách quân nhân trong dialog
                nonlocal all_personnel, selected_ids
                all_personnel = self.db.get_all()
                selected_ids = set(unit.personnelIds)
                filter_text = search_var.get()
                load_tree_data(filter_text)
            
            edit_window.protocol("WM_DELETE_WINDOW", lambda: [on_edit_close(), edit_window.destroy()])
        
        def save():
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
            
            # Cập nhật unit
            unit.personnelIds = list(selected_ids)
            try:
                self.db.update_unit(unit)
                
                # Cập nhật unitId cho quân nhân
                all_personnel = self.db.get_all()
                for person in all_personnel:
                    if person.id in selected_ids:
                        person.unitId = unit.id
                    elif person.unitId == unit.id:
                        person.unitId = None
                    self.db.update(person)
                
                self.status_label.config(
                    text=f"✅ Đã cập nhật {len(selected_ids)} quân nhân vào đơn vị '{unit.ten}'",
                    fg='#4CAF50'
                )
                dialog.destroy()
                self.load_units()
                self.on_unit_select(None)  # Refresh danh sách quân nhân
            except Exception as e:
                self.status_label.config(text=f"❌ Lỗi: {str(e)}", fg='#F44336')
        
        # Buttons - Row 1, LUÔN HIỂN THỊ
        btn_frame = tk.Frame(dialog, bg='#FAFAFA', height=70)
        btn_frame.grid(row=1, column=0, padx=10, pady=10, sticky=tk.EW)
        btn_frame.grid_propagate(False)
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
            command=save,
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
        
        # Đảm bảo dialog có đủ không gian
        dialog.update_idletasks()
        dialog.minsize(1100, 700)
    
    def add_to_to_unit(self):
        """Thêm tổ vào đơn vị đã chọn"""
        selection = self.tree.selection()
        if not selection:
            self.status_label.config(text="⚠️ Vui lòng chọn đơn vị cha", fg='#FF9800')
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            self.status_label.config(text="❌ Không tìm thấy đơn vị", fg='#F44336')
            return
        
        # Cho phép thêm tổ vào đại đội, trung đội, hoặc xe
        if unit.loai not in ['dai_doi', 'trung_doi', 'xe']:
            self.status_label.config(
                text="⚠️ Chỉ có thể thêm tổ vào Đại đội, Trung đội hoặc Xe",
                fg='#FF9800'
            )
            return
        
        # Tạo tổ với parent_id là unit_id
        self.create_unit('to', parent_id=unit_id)
    
    def add_child_unit(self, loai: str):
        """Thêm đơn vị con (xe, trung đội) vào đơn vị đã chọn"""
        selection = self.tree.selection()
        if not selection:
            self.status_label.config(text="⚠️ Vui lòng chọn đơn vị cha", fg='#FF9800')
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            self.status_label.config(text="❌ Không tìm thấy đơn vị", fg='#F44336')
            return
        
        # Chỉ cho phép thêm đơn vị con vào đại đội
        if unit.loai != 'dai_doi':
            self.status_label.config(
                text="⚠️ Chỉ có thể thêm đơn vị con vào Đại đội",
                fg='#FF9800'
            )
            return
        
        # Tạo đơn vị con với parent_id là unit_id
        self.create_unit(loai, parent_id=unit_id)
    
    def export_word(self):
        """Xuất file Word cho đơn vị đã chọn"""
        from tkinter import filedialog
        from services.export import ExportService
        from datetime import datetime
        
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn đơn vị cần xuất")
            return
        
        unit_id = selection[0]
        unit = self.db.get_unit_by_id(unit_id)
        if not unit:
            messagebox.showerror("Lỗi", "Không tìm thấy đơn vị")
            return
        
        # Chỉ cho phép xuất đại đội hoặc trung đội
        if unit.loai not in ['dai_doi', 'trung_doi']:
            messagebox.showwarning("Cảnh báo", "Chỉ có thể xuất file cho Đại đội hoặc Trung đội")
            return
        
        # Lấy tất cả tổ con
        child_units = self.db.get_units_by_parent_id(unit_id)
        
        if not child_units:
            messagebox.showwarning("Cảnh báo", f"Đơn vị '{unit.ten}' chưa có tổ nào")
            return
        
        # Thu thập tất cả quân nhân từ các tổ và sắp xếp theo cấp bậc
        def _parse_cap_bac_rank(cap_bac: str) -> int:
            """Parse cấp bậc thành số để so sánh"""
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
        
        def _sort_personnel_by_cap_bac(personnel_list):
            """Sắp xếp danh sách quân nhân theo cấp bậc (từ cao xuống thấp)"""
            def sort_key(personnel):
                cap_bac_rank = _parse_cap_bac_rank(personnel.capBac or '')
                ho_ten = (personnel.hoTen or '').lower()
                return (-cap_bac_rank, ho_ten)
            return sorted(personnel_list, key=sort_key)
        
        all_personnel_data = []
        for child_unit in child_units:
            personnel_list = self.db.get_personnel_by_unit(child_unit.id)
            # Sắp xếp theo cấp bậc (từ cao xuống thấp)
            personnel_list = _sort_personnel_by_cap_bac(personnel_list)
            all_personnel_data.append({
                'to': child_unit,
                'personnel': personnel_list
            })
        
        # Mở dialog nhập thông tin
        dialog = tk.Toplevel(self)
        dialog.title("Xuất File Word - Quản Lý Đơn Vị")
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
        dai_doi_default = unit.ten if unit.loai == 'dai_doi' else ""
        dai_doi_var = tk.StringVar(value=dai_doi_default)
        tk.Entry(main_container, textvariable=dai_doi_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
        
        # Địa điểm
        tk.Label(main_container, text="Địa điểm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
        dia_diem_var = tk.StringVar(value="Đắk Lăk")
        tk.Entry(main_container, textvariable=dia_diem_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
        
        # Năm
        tk.Label(main_container, text="Năm:", font=('Segoe UI', 10), bg='#FAFAFA').pack(anchor=tk.W, pady=5)
        nam_var = tk.StringVar(value=str(datetime.now().year))
        tk.Entry(main_container, textvariable=nam_var, width=40, font=('Segoe UI', 10)).pack(anchor=tk.W, pady=2)
        
        def save_and_export():
            try:
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".docx",
                    filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
                    title="Lưu file Word"
                )
                
                if not file_path:
                    return
                
                # Gọi hàm xuất Word mới - sử dụng export_trich_ngang với units_data
                from services.export_trich_ngang import to_word_docx_trich_ngang
                
                # Thu thập tất cả quân nhân từ các tổ
                all_personnel = []
                for unit_group in all_personnel_data:
                    all_personnel.extend(unit_group.get('personnel', []))
                
                word_bytes = to_word_docx_trich_ngang(
                    personnel_list=all_personnel,
                    tieu_doan=tieu_doan_var.get(),
                    dai_doi=dai_doi_var.get(),
                    dia_diem=dia_diem_var.get(),
                    nam=nam_var.get(),
                    db_service=self.db,
                    units_data=all_personnel_data  # Truyền units_data để nhóm theo đơn vị
                )
                
                with open(file_path, 'wb') as f:
                    f.write(word_bytes)
                
                messagebox.showinfo("Thành công", f"Đã xuất file Word:\n{file_path}")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xuất file:\n{str(e)}")
        
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