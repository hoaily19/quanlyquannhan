"""
Frame form thêm/sửa quân nhân - Giao diện dễ chịu, thoải mái
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path
import logging

sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from models.personnel import Personnel, ThongTinDang, ThongTinDoan, ThongTinKhac
from models.nguoi_than import NguoiThan
from gui.date_picker import DatePicker
from gui.theme import MILITARY_COLORS, get_button_style, get_label_style

# Setup logging
logger = logging.getLogger(__name__)


class PersonnelFormFrame(tk.Frame):
    """Frame form thêm/sửa quân nhân - Giao diện dễ chịu"""
    
    def __init__(self, parent, db: DatabaseService, is_new: bool = False, personnel_id: str = None):
        """
        Args:
            parent: Parent widget
            db: DatabaseService instance
            is_new: True nếu là thêm mới
            personnel_id: ID quân nhân nếu là sửa
        """
        super().__init__(parent)
        self.db = db
        self.is_new = is_new
        self.personnel_id = personnel_id
        self.personnel = Personnel() if is_new else db.get_by_id(personnel_id)
        
        if not self.personnel:
            self.personnel = Personnel()
            self.is_new = True
        
        # Đảm bảo frame có background đúng
        self.configure(bg=MILITARY_COLORS['bg_light'])
        
        # Màu sắc nhẹ nhàng hơn
        self.bg_color = '#FAFAFA'  # Trắng nhẹ
        self.section_bg = '#FFFFFF'  # Trắng tinh
        self.border_color = '#E0E0E0'  # Xám nhẹ
        self.text_color = '#424242'  # Xám đậm nhẹ
        self.title_color = '#388E3C'  # Xanh lá nhẹ hơn
        
        self.setup_ui()
        if not is_new:
            self.load_data()
    
    def setup_ui(self):
        """Thiết lập giao diện - dễ chịu, thoải mái"""
        # Configure frame background - màu nhẹ nhàng
        self.configure(bg=self.bg_color)
        
        # Title bar - mềm mại hơn
        title = "➕ Thêm Quân Nhân Mới" if self.is_new else f"✏️ Sửa: {self.personnel.hoTen or ''}"
        title_frame = tk.Frame(self, bg=self.title_color, height=70)
        title_frame.pack(fill=tk.X, pady=(0, 0))
        title_frame.pack_propagate(False)
        
        tk.Label(
            title_frame,
            text=title,
            font=('Segoe UI', 18, 'bold'),
            bg=self.title_color,
            fg='white'
        ).pack(expand=True, pady=20)
        
        # Scrollable frame - màu nền nhẹ
        canvas = tk.Canvas(self, bg=self.bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.bg_color)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Tạo window với full width
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # Bind để resize scrollable_frame theo canvas width
        def configure_scroll_region(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        canvas.bind('<Configure>', configure_scroll_region)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Form fields với layout 2 cột - spacing rộng rãi
        self.create_form_fields(scrollable_frame)
        
        # Buttons - mềm mại hơn
        btn_frame = tk.Frame(scrollable_frame, bg=self.bg_color, pady=30)
        btn_frame.pack(fill=tk.X)
        
        # Container cho buttons căn giữa
        btn_container = tk.Frame(btn_frame, bg=self.bg_color)
        btn_container.pack()
        
        save_btn = tk.Button(
            btn_container,
            text="💾 Lưu",
            command=self.save,
            **get_button_style('success')
        )
        save_btn.config(
            font=('Segoe UI', 11, 'bold'),
            padx=35,
            pady=12,
            width=14,
            relief=tk.FLAT,
            bd=0
        )
        save_btn.pack(side=tk.LEFT, padx=12)
        
        cancel_btn = tk.Button(
            btn_container,
            text="❌ Hủy",
            command=self.cancel,
            **get_button_style('danger')
        )
        cancel_btn.config(
            font=('Segoe UI', 11, 'bold'),
            padx=35,
            pady=12,
            width=14,
            relief=tk.FLAT,
            bd=0
        )
        cancel_btn.pack(side=tk.LEFT, padx=12)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel - xử lý cuộn cho toàn bộ form nhưng không gây lỗi khi canvas bị destroy
        def _on_mousewheel(event):
            """Cuộn nội dung theo bánh xe chuột"""
            try:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                # Tránh crash nếu canvas đã bị destroy
                pass
        
        # Lưu handler để có thể unbind khi destroy
        self._mousewheel_handler = _on_mousewheel
        
        # Bind cho nhiều vùng để người dùng đặt chuột ở đâu cũng cuộn được
        canvas.bind("<MouseWheel>", _on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        self.bind("<MouseWheel>", _on_mousewheel)
        try:
            # Hỗ trợ cả trường hợp focus không nằm trên canvas
            self.bind_all("<MouseWheel>", _on_mousewheel)
        except Exception:
            pass

    def destroy(self):
        """Hủy frame - đảm bảo unbind mouse wheel để tránh side-effect toàn app"""
        try:
            if hasattr(self, "_mousewheel_handler"):
                try:
                    self.unbind("<MouseWheel>")
                except Exception:
                    pass
                try:
                    self.unbind_all("<MouseWheel>")
                except Exception:
                    pass
        except Exception:
            pass
        # Gọi destroy gốc
        super().destroy()
    
    def create_form_fields(self, parent):
        """Tạo các trường form với layout 2 cột - spacing rộng rãi"""
        # Thông tin cơ bản
        self.create_section(parent, "📋 Thông Tin Cơ Bản")
        
        # Container 2 cột - padding rộng rãi
        basic_container = tk.Frame(parent, bg=self.bg_color)
        basic_container.pack(fill=tk.X, padx=25, pady=12)
        
        # Cột trái - expand full
        left_col = tk.Frame(basic_container, bg=self.bg_color)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        # Cột phải - expand full
        right_col = tk.Frame(basic_container, bg=self.bg_color)
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(15, 0))
        
        # Cột trái
        self.ho_ten_var = self.create_field(left_col, "Họ và Tên *", self.personnel.hoTen, required=True)
        self.ngay_sinh_picker = self.create_date_field(left_col, "Ngày Sinh", self.personnel.ngaySinh)
        self.cap_bac_var = self.create_field(left_col, "Cấp Bậc", self.personnel.capBac)
        self.chuc_vu_var = self.create_field(left_col, "Chức Vụ", self.personnel.chucVu)
        self.don_vi_var = self.create_field(left_col, "Đơn Vị", self.personnel.donVi)
        self.nhap_ngu_picker = self.create_date_field(left_col, "Nhập Ngũ", self.personnel.nhapNgu)
        
        # Cột phải
        self.que_quan_var = self.create_field(right_col, "Quê Quán", self.personnel.queQuan)
        self.tru_quan_var = self.create_field(right_col, "Trú Quán", self.personnel.truQuan)
        self.dan_toc_var = self.create_field(right_col, "Dân Tộc", self.personnel.danToc)
        self.ton_giao_var = self.create_field(right_col, "Tôn Giáo", self.personnel.tonGiao)
        self.trinh_do_var = self.create_field(right_col, "Trình Độ Văn Hóa", self.personnel.trinhDoVanHoa)
        self.ngoai_ngu_var = self.create_field(right_col, "Ngoại Ngữ", self.personnel.ngoaiNgu)
        self.tieng_dtts_var = self.create_field(right_col, "Tiếng DTTS", self.personnel.tiengDTTS)
        
        # Thông tin học vấn
        self.create_section(parent, "🎓 Thông Tin Học Vấn")
        
        hoc_van_container = tk.Frame(parent, bg=self.bg_color)
        hoc_van_container.pack(fill=tk.X, padx=25, pady=12)
        
        hoc_van_left = tk.Frame(hoc_van_container, bg=self.bg_color)
        hoc_van_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        hoc_van_right = tk.Frame(hoc_van_container, bg=self.bg_color)
        hoc_van_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(15, 0))
        
        self.qua_truong_var = self.create_field(hoc_van_left, "Qua Trường", self.personnel.quaTruong)
        self.nganh_hoc_var = self.create_field(hoc_van_left, "Ngành Học", self.personnel.nganhHoc)
        self.cap_hoc_var = self.create_field(hoc_van_right, "Cấp Học", self.personnel.capHoc)
        self.thoi_gian_dao_tao_var = self.create_field(hoc_van_right, "Thời Gian Đào Tạo", self.personnel.thoiGianDaoTao)
        self.ket_qua_dao_tao_var = self.create_field(hoc_van_right, "Kết Quả Đào Tạo", self.personnel.ketQuaDaoTao)
        
        # Thông tin chức vụ và thời gian
        self.create_section(parent, "⚔️ Thông Tin Chức Vụ Chiến Đấu")
        
        chuc_vu_container = tk.Frame(parent, bg=self.bg_color)
        chuc_vu_container.pack(fill=tk.X, padx=25, pady=12)
        
        chuc_vu_left = tk.Frame(chuc_vu_container, bg=self.bg_color)
        chuc_vu_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        chuc_vu_right = tk.Frame(chuc_vu_container, bg=self.bg_color)
        chuc_vu_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(15, 0))
        
        self.chuc_vu_chien_dau_var = self.create_field(chuc_vu_left, "Chức Vụ Chiến Đấu", self.personnel.chucVuChienDau)
        self.thoi_gian_chuc_vu_chien_dau_var = self.create_field(chuc_vu_left, "Thời Gian Chức Vụ Chiến Đấu", self.personnel.thoiGianChucVuChienDau)
        self.chuc_vu_da_qua_var = self.create_field(chuc_vu_right, "Chức Vụ Đã Qua", self.personnel.chucVuDaQua)
        self.thoi_gian_chuc_vu_da_qua_var = self.create_field(chuc_vu_right, "Thời Gian Chức Vụ Đã Qua", self.personnel.thoiGianChucVuDaQua)
        
        # Thông tin CM Quân và ngày nhận
        self.create_section(parent, "📅 Thông Tin Ngày Nhận")
        
        ngay_nhan_container = tk.Frame(parent, bg=self.bg_color)
        ngay_nhan_container.pack(fill=tk.X, padx=25, pady=12)
        
        ngay_nhan_left = tk.Frame(ngay_nhan_container, bg=self.bg_color)
        ngay_nhan_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        ngay_nhan_right = tk.Frame(ngay_nhan_container, bg=self.bg_color)
        ngay_nhan_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(15, 0))
        
        self.ngay_nhan_cap_bac_picker = self.create_date_field(ngay_nhan_left, "Ngày Nhận Cấp Bậc", self.personnel.ngayNhanCapBac)
        self.ngay_nhan_chuc_vu_picker = self.create_date_field(ngay_nhan_left, "Ngày Nhận Chức Vụ", self.personnel.ngayNhanChucVu)
        self.cm_quan_picker = self.create_date_field(ngay_nhan_right, "CM Quân (Tháng năm)", self.personnel.cmQuan or self.personnel.nhapNgu)
        
        # Thông tin đảng
        self.create_section(parent, "🏛️ Thông Tin Đảng")
        
        dang_container = tk.Frame(parent, bg=self.bg_color)
        dang_container.pack(fill=tk.X, padx=25, pady=12)
        
        dang_left = tk.Frame(dang_container, bg=self.bg_color)
        dang_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        dang_right = tk.Frame(dang_container, bg=self.bg_color)
        dang_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(15, 0))
        
        self.dang_ngay_vao_picker = self.create_date_field(dang_left, "Ngày Vào Đảng", self.personnel.thongTinKhac.dang.ngayVao)
        self.dang_ngay_chinh_thuc_picker = self.create_date_field(dang_left, "Ngày Chính Thức", self.personnel.thongTinKhac.dang.ngayChinhThuc)
        self.dang_chuc_vu_var = self.create_field(dang_right, "Chức Vụ Đảng", self.personnel.thongTinKhac.dang.chucVuDang)
        
        # Thông tin đoàn
        self.create_section(parent, "👥 Thông Tin Đoàn")
        
        doan_container = tk.Frame(parent, bg=self.bg_color)
        doan_container.pack(fill=tk.X, padx=25, pady=12)
        
        doan_left = tk.Frame(doan_container, bg=self.bg_color)
        doan_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        doan_right = tk.Frame(doan_container, bg=self.bg_color)
        doan_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(15, 0))
        
        self.doan_ngay_vao_picker = self.create_date_field(doan_left, "Ngày Vào Đoàn", self.personnel.thongTinKhac.doan.ngayVao)
        self.doan_chuc_vu_var = self.create_field(doan_right, "Chức Vụ Đoàn", self.personnel.thongTinKhac.doan.chucVuDoan)
        
        # Thông tin khác
        self.create_section(parent, "ℹ️ Thông Tin Khác")
        
        other_frame = tk.Frame(parent, bg=self.section_bg, relief=tk.FLAT, bd=0)
        other_frame.pack(fill=tk.X, padx=25, pady=12)
        
        # Thêm border nhẹ bằng cách dùng Frame bên ngoài
        border_frame = tk.Frame(other_frame, bg=self.border_color, height=1)
        border_frame.pack(fill=tk.X, padx=0, pady=0)
        
        inner_frame = tk.Frame(other_frame, bg=self.section_bg)
        inner_frame.pack(fill=tk.X, padx=20, pady=15)
        
        self.cd_cu_var = tk.BooleanVar(value=self.personnel.thongTinKhac.cdCu)
        cd_cu_check = tk.Checkbutton(
            inner_frame,
            text="Có người thân tham gia chế độ cũ",
            variable=self.cd_cu_var,
            font=('Segoe UI', 10),
            bg=self.section_bg,
            fg=self.text_color,
            activebackground=self.section_bg,
            activeforeground=self.text_color,
            selectcolor='white'
        )
        cd_cu_check.pack(anchor=tk.W, pady=8)
        
        self.yeu_to_nn_var = tk.BooleanVar(value=self.personnel.thongTinKhac.yeuToNN)
        
        # Frame chứa checkbox và nút
        yeu_to_nn_frame = tk.Frame(inner_frame, bg=self.section_bg)
        yeu_to_nn_frame.pack(fill=tk.X, pady=8)
        
        yeu_to_nn_check = tk.Checkbutton(
            yeu_to_nn_frame,
            text="Có yếu tố nước ngoài",
            variable=self.yeu_to_nn_var,
            font=('Segoe UI', 10),
            bg=self.section_bg,
            fg=self.text_color,
            activebackground=self.section_bg,
            activeforeground=self.text_color,
            selectcolor='white',
            command=self.on_yeu_to_nn_changed
        )
        yeu_to_nn_check.pack(side=tk.LEFT)
        
        # Nút nhập thông tin yếu tố nước ngoài - chỉ hiện khi đã tick
        self.yeu_to_nn_btn = tk.Button(
            yeu_to_nn_frame,
            text="📝 Nhập Thông Tin Yếu Tố Nước Ngoài",
            command=self.open_yeu_to_nn_form,
            font=('Segoe UI', 9),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=10,
            pady=3,
            cursor='hand2'
        )
        
        # Ẩn/hiện nút dựa trên giá trị checkbox ban đầu
        if self.yeu_to_nn_var.get():
            self.yeu_to_nn_btn.pack(side=tk.LEFT, padx=(15, 0))
        else:
            self.yeu_to_nn_btn.pack(side=tk.LEFT, padx=(15, 0))
            self.yeu_to_nn_btn.pack_forget()
        
        # Thông tin THAM GIA chế độ cũ
        self.create_section(parent, "📋 Thông Tin THAM GIA Chế Độ Cũ")
        
        tham_gia_container = tk.Frame(parent, bg=self.bg_color)
        tham_gia_container.pack(fill=tk.X, padx=25, pady=12)
        
        tham_gia_left = tk.Frame(tham_gia_container, bg=self.bg_color)
        tham_gia_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        tham_gia_right = tk.Frame(tham_gia_container, bg=self.bg_color)
        tham_gia_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(15, 0))
        
        # Checkbox Ngụy quân
        self.tham_gia_nguy_quan_var = tk.BooleanVar(value=bool(self.personnel.thamGiaNguyQuan))
        nguy_quan_check = tk.Checkbutton(
            tham_gia_left,
            text="Ngụy quân",
            variable=self.tham_gia_nguy_quan_var,
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg=self.text_color,
            activebackground=self.bg_color,
            activeforeground=self.text_color,
            selectcolor='white'
        )
        nguy_quan_check.pack(anchor=tk.W, pady=8)
        
        # Checkbox Ngụy quyền
        self.tham_gia_nguy_quyen_var = tk.BooleanVar(value=bool(self.personnel.thamGiaNguyQuyen))
        nguy_quyen_check = tk.Checkbutton(
            tham_gia_left,
            text="Ngụy quyền",
            variable=self.tham_gia_nguy_quyen_var,
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg=self.text_color,
            activebackground=self.bg_color,
            activeforeground=self.text_color,
            selectcolor='white'
        )
        nguy_quyen_check.pack(anchor=tk.W, pady=8)
        
        # Select Nợ máu/không nợ máu
        self.tham_gia_no_mau_var = tk.StringVar(value=self.personnel.thamGiaNoMau or '')
        no_mau_frame = tk.Frame(tham_gia_right, bg=self.bg_color)
        no_mau_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(
            no_mau_frame,
            text="Nợ máu/không nợ máu",
            font=('Segoe UI', 10),
            width=20,
            anchor=tk.W,
            bg=self.bg_color,
            fg=self.text_color
        ).pack(side=tk.LEFT, padx=(0, 15))
        
        no_mau_combo = ttk.Combobox(
            no_mau_frame,
            textvariable=self.tham_gia_no_mau_var,
            values=['', 'Nợ máu', 'Không nợ máu'],
            font=('Segoe UI', 10),
            state='readonly',
            width=20
        )
        no_mau_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Select Đã cải tạo/chưa cải tạo
        self.da_cai_tao_var = tk.StringVar(value=self.personnel.daCaiTao or '')
        cai_tao_frame = tk.Frame(tham_gia_right, bg=self.bg_color)
        cai_tao_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(
            cai_tao_frame,
            text="Đã cải tạo/chưa cải tạo",
            font=('Segoe UI', 10),
            width=20,
            anchor=tk.W,
            bg=self.bg_color,
            fg=self.text_color
        ).pack(side=tk.LEFT, padx=(0, 15))
        
        cai_tao_combo = ttk.Combobox(
            cai_tao_frame,
            textvariable=self.da_cai_tao_var,
            values=['', 'Đã cải tạo', 'Chưa cải tạo'],
            font=('Segoe UI', 10),
            state='readonly',
            width=20
        )
        cai_tao_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Thông tin người thân
        self.create_section(parent, "👥 Thông Tin Người Thân")
        
        # Checkbox "Tham gia đảng phái phản động"
        dang_phai_frame = tk.Frame(parent, bg=self.bg_color)
        dang_phai_frame.pack(fill=tk.X, padx=25, pady=10)
        
        self.dang_phai_phan_dong_var = tk.BooleanVar(value=self.personnel.thongTinKhac.dangPhaiPhanDong)
        dang_phai_checkbox = tk.Checkbutton(
            dang_phai_frame,
            text="Tham gia đảng phái phản động",
            variable=self.dang_phai_phan_dong_var,
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg=self.text_color,
            activebackground=self.bg_color,
            activeforeground=self.text_color,
            selectcolor=self.bg_color
        )
        dang_phai_checkbox.pack(side=tk.LEFT)
        
        # Toolbar với nút thêm người thân
        nguoi_than_toolbar = tk.Frame(parent, bg=self.bg_color)
        nguoi_than_toolbar.pack(fill=tk.X, padx=25, pady=5)
        
        tk.Button(
            nguoi_than_toolbar,
            text="➕ Thêm Người Thân Mới",
            command=self.add_nguoi_than,
            font=('Segoe UI', 10, 'bold'),
            bg='#4CAF50',
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT)
        
        # Danh sách người thân - mở rộng để hiển thị đầy đủ
        self.nguoi_than_frame = tk.Frame(parent, bg=self.bg_color)
        self.nguoi_than_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=15)
        
        # Load danh sách người thân hiện có
        self.load_nguoi_than_list()
    
    def create_section(self, parent, title):
        """Tạo section header - mềm mại, dễ nhìn"""
        # Separator line - mềm mại
        separator_frame = tk.Frame(parent, bg=self.bg_color, height=2)
        separator_frame.pack(fill=tk.X, padx=0, pady=(30, 10))
        
        # Line màu nhẹ - mềm mại
        separator = tk.Frame(separator_frame, bg=self.border_color, height=1)
        separator.pack(fill=tk.X, padx=25)
        
        # Section title - font nhẹ nhàng
        title_frame = tk.Frame(parent, bg=self.bg_color)
        title_frame.pack(fill=tk.X, padx=25, pady=(0, 15))
        
        tk.Label(
            title_frame,
            text=title,
            font=('Segoe UI', 13, 'bold'),
            bg=self.bg_color,
            fg=self.title_color
        ).pack(anchor=tk.W)
    
    def create_field(self, parent, label, default_value="", required=False, is_textarea=False):
        """Tạo một trường input - mềm mại, dễ nhìn"""
        field_frame = tk.Frame(parent, bg=self.bg_color)
        field_frame.pack(fill=tk.X, pady=10)
        
        # Label - font nhẹ nhàng
        label_widget = tk.Label(
            field_frame,
            text=label,
            font=('Segoe UI', 10),
            width=20,
            anchor=tk.W,
            bg=self.bg_color,
            fg='#E53935' if required else self.text_color
        )
        label_widget.pack(side=tk.LEFT, padx=(0, 15))
        
        if is_textarea:
            # Textarea cho nội dung dài
            var = tk.StringVar(value=default_value or "")
            text_widget = tk.Text(
                field_frame,
                font=('Segoe UI', 10),
                relief=tk.FLAT,
                bd=1,
                bg=self.section_bg,
                fg=self.text_color,
                insertbackground=self.title_color,
                highlightthickness=1,
                highlightcolor=self.title_color,
                highlightbackground=self.border_color,
                wrap=tk.WORD,
                height=4
            )
            text_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
            text_widget.insert('1.0', default_value or "")
            
            # Bind để cập nhật var
            def update_var(event=None):
                var.set(text_widget.get('1.0', tk.END).strip())
            text_widget.bind('<KeyRelease>', update_var)
            text_widget.bind('<FocusOut>', update_var)
            
            return var
        
        # Entry - border mềm mại
        var = tk.StringVar(value=default_value or "")
        entry = tk.Entry(
            field_frame,
            textvariable=var,
            font=('Segoe UI', 10),
            relief=tk.FLAT,
            bd=1,
            bg=self.section_bg,
            fg=self.text_color,
            insertbackground=self.title_color,
            highlightthickness=1,
            highlightcolor=self.title_color,
            highlightbackground=self.border_color
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        
        # Highlight border khi focus - mềm mại
        def on_focus_in(e):
            entry.config(highlightbackground=self.title_color, highlightthickness=2)
        
        def on_focus_out(e):
            entry.config(highlightbackground=self.border_color, highlightthickness=1)
        
        entry.bind('<FocusIn>', on_focus_in)
        entry.bind('<FocusOut>', on_focus_out)
        
        return var
    
    def create_field(self, parent, label, default_value="", required=False):
        """Tạo một trường input - mềm mại, dễ nhìn"""
        field_frame = tk.Frame(parent, bg=self.bg_color)
        field_frame.pack(fill=tk.X, pady=10)
        
        # Label - font nhẹ nhàng
        label_widget = tk.Label(
            field_frame,
            text=label,
            font=('Segoe UI', 10),
            width=20,
            anchor=tk.W,
            bg=self.bg_color,
            fg='#E53935' if required else self.text_color
        )
        label_widget.pack(side=tk.LEFT, padx=(0, 15))
        
        # Entry - border mềm mại
        var = tk.StringVar(value=default_value or "")
        entry = tk.Entry(
            field_frame,
            textvariable=var,
            font=('Segoe UI', 10),
            relief=tk.FLAT,
            bd=1,
            bg=self.section_bg,
            fg=self.text_color,
            insertbackground=self.title_color,
            highlightthickness=1,
            highlightcolor=self.title_color,
            highlightbackground=self.border_color
        )
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        
        # Highlight border khi focus - mềm mại
        def on_focus_in(e):
            entry.config(highlightbackground=self.title_color, highlightthickness=2)
        
        def on_focus_out(e):
            entry.config(highlightbackground=self.border_color, highlightthickness=1)
        
        entry.bind('<FocusIn>', on_focus_in)
        entry.bind('<FocusOut>', on_focus_out)
        
        return var
    
    def create_date_field(self, parent, label, default_value=""):
        """Tạo một trường date picker - mềm mại, dễ nhìn"""
        field_frame = tk.Frame(parent, bg=self.bg_color)
        field_frame.pack(fill=tk.X, pady=10)
        
        # Label - font nhẹ nhàng
        tk.Label(
            field_frame,
            text=label,
            font=('Segoe UI', 10),
            width=20,
            anchor=tk.W,
            bg=self.bg_color,
            fg=self.text_color
        ).pack(side=tk.LEFT, padx=(0, 15))
        
        # Date picker - full width
        date_picker = DatePicker(field_frame, default_value)
        date_picker.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        return date_picker
    
    def load_data(self):
        """Load dữ liệu vào form"""
        # Đã load trong __init__, không cần làm gì thêm
        pass
    
    def save(self):
        """Lưu dữ liệu"""
        # Validate
        if not self.ho_ten_var.get().strip():
            messagebox.showerror("Lỗi", "Vui lòng nhập Họ và Tên (trường bắt buộc)")
            return
        
        # Cập nhật personnel object
        self.personnel.hoTen = self.ho_ten_var.get().strip()
        self.personnel.ngaySinh = self.ngay_sinh_picker.get_date()
        self.personnel.capBac = self.cap_bac_var.get().strip()
        self.personnel.chucVu = self.chuc_vu_var.get().strip()
        self.personnel.donVi = self.don_vi_var.get().strip()
        self.personnel.nhapNgu = self.nhap_ngu_picker.get_date()
        self.personnel.queQuan = self.que_quan_var.get().strip()
        self.personnel.truQuan = self.tru_quan_var.get().strip()
        self.personnel.danToc = self.dan_toc_var.get().strip()
        self.personnel.tonGiao = self.ton_giao_var.get().strip()
        self.personnel.trinhDoVanHoa = self.trinh_do_var.get().strip()
        self.personnel.ngoaiNgu = self.ngoai_ngu_var.get().strip()
        self.personnel.tiengDTTS = self.tieng_dtts_var.get().strip()
        # Thông tin học vấn
        self.personnel.quaTruong = self.qua_truong_var.get().strip()
        self.personnel.nganhHoc = self.nganh_hoc_var.get().strip()
        self.personnel.capHoc = self.cap_hoc_var.get().strip()
        self.personnel.thoiGianDaoTao = self.thoi_gian_dao_tao_var.get().strip()
        self.personnel.ketQuaDaoTao = self.ket_qua_dao_tao_var.get().strip()
        
        # Thông tin chức vụ chiến đấu
        self.personnel.chucVuChienDau = self.chuc_vu_chien_dau_var.get().strip()
        self.personnel.thoiGianChucVuChienDau = self.thoi_gian_chuc_vu_chien_dau_var.get().strip()
        self.personnel.chucVuDaQua = self.chuc_vu_da_qua_var.get().strip()
        self.personnel.thoiGianChucVuDaQua = self.thoi_gian_chuc_vu_da_qua_var.get().strip()
        
        # Thông tin ngày nhận
        self.personnel.ngayNhanCapBac = self.ngay_nhan_cap_bac_picker.get_date()
        self.personnel.ngayNhanChucVu = self.ngay_nhan_chuc_vu_picker.get_date()
        self.personnel.cmQuan = self.cm_quan_picker.get_date()
        
        self.personnel.thongTinKhac.dang.ngayVao = self.dang_ngay_vao_picker.get_date()
        self.personnel.thongTinKhac.dang.ngayChinhThuc = self.dang_ngay_chinh_thuc_picker.get_date()
        self.personnel.thongTinKhac.dang.chucVuDang = self.dang_chuc_vu_var.get().strip()
        
        self.personnel.thongTinKhac.doan.ngayVao = self.doan_ngay_vao_picker.get_date()
        self.personnel.thongTinKhac.doan.chucVuDoan = self.doan_chuc_vu_var.get().strip()
        
        self.personnel.thongTinKhac.cdCu = self.cd_cu_var.get()
        self.personnel.thongTinKhac.yeuToNN = self.yeu_to_nn_var.get()
        self.personnel.thongTinKhac.dangPhaiPhanDong = self.dang_phai_phan_dong_var.get()
        
        # Thông tin THAM GIA
        self.personnel.thamGiaNguyQuan = 'X' if self.tham_gia_nguy_quan_var.get() else ''
        self.personnel.thamGiaNguyQuyen = 'X' if self.tham_gia_nguy_quyen_var.get() else ''
        self.personnel.thamGiaNoMau = self.tham_gia_no_mau_var.get().strip()
        self.personnel.daCaiTao = self.da_cai_tao_var.get().strip()
        
        # Thông tin người thân - không cần lưu vào personnel nữa vì đã có bảng riêng
        
        # Lưu vào database
        try:
            cd_cu_value = self.cd_cu_var.get()
            dang_phai_phan_dong_value = self.dang_phai_phan_dong_var.get()
            if self.is_new:
                self.db.create(self.personnel)
                messagebox.showinfo("Thành công", f"Đã thêm quân nhân: {self.personnel.hoTen}")
                # Lưu personnel_id để load người thân
                self.personnel_id = self.personnel.id
            else:
                self.personnel.id = self.personnel_id
                if self.db.update(self.personnel):
                    messagebox.showinfo("Thành công", f"Đã cập nhật quân nhân: {self.personnel.hoTen}")
            
            # Tự động thêm/xóa khỏi danh sách "Quân nhân có người thân tham gia chế độ cũ"
            # dựa trên checkbox "Có người thân tham gia chế độ cũ"
            if self.personnel.id:
                if cd_cu_value:
                    # Nếu checkbox được đánh dấu, thêm vào danh sách
                    self.db.add_nguoi_than_che_do_cu(self.personnel.id)
                else:
                    # Nếu checkbox không được đánh dấu, xóa khỏi danh sách
                    self.db.remove_nguoi_than_che_do_cu(self.personnel.id)
                
                # Tự động thêm/xóa khỏi danh sách "Người thân đảng phái phản động"
                # dựa trên checkbox "Tham gia đảng phái phản động"
                if dang_phai_phan_dong_value:
                    # Nếu checkbox được đánh dấu, thêm vào danh sách
                    self.db.add_nguoi_than_dang_phai_phan_dong(self.personnel.id)
                else:
                    # Nếu checkbox không được đánh dấu, xóa khỏi danh sách
                    self.db.remove_nguoi_than_dang_phai_phan_dong(self.personnel.id)

            # Xử lý đóng form sau khi lưu
            parent = self.master

            # Nếu form đang nằm trong cửa sổ modal (Toplevel) -> chỉ đóng đúng modal
            if isinstance(parent, tk.Toplevel):
                try:
                    parent.grab_release()
                except Exception:
                    pass
                try:
                    parent.destroy()
                except Exception:
                    pass
                # Reload danh sách sẽ được xử lý ở nơi mở dialog (PersonnelListFrame)
                return

            # Ngược lại: form đang chạy trực tiếp trên root -> giữ hành vi cũ (quay về danh sách)
            try:
                def close_all_toplevels(widget):
                    """Đệ quy đóng tất cả Toplevel windows"""
                    toplevels = []

                    def find_toplevels(w):
                        if isinstance(w, tk.Toplevel):
                            toplevels.append(w)
                        try:
                            for child in w.winfo_children():
                                find_toplevels(child)
                        except Exception:
                            pass

                    find_toplevels(widget)

                    for toplevel in toplevels:
                        try:
                            toplevel.grab_release()
                        except Exception:
                            pass
                        try:
                            toplevel.destroy()
                        except Exception:
                            pass

                # Đóng tất cả Toplevel trong frame
                close_all_toplevels(self)

                # Đóng tất cả Toplevel trong root window
                root = self.master
                while hasattr(root, 'master') and root.master:
                    root = root.master
                close_all_toplevels(root)

                # Update UI sau khi đóng popup
                try:
                    self.update_idletasks()
                    self.master.update_idletasks()
                    root.update_idletasks()
                except Exception:
                    pass
            except Exception:
                pass

            # Quay lại danh sách - đóng form hiện tại
            self.cancel()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lưu:\n{str(e)}")
            import traceback
            traceback.print_exc()
    
    def load_nguoi_than_list(self):
        """Load danh sách người thân"""
        # Xóa các widget cũ
        for widget in self.nguoi_than_frame.winfo_children():
            widget.destroy()
        
        if not self.personnel_id:
            tk.Label(
                self.nguoi_than_frame,
                text="Lưu quân nhân trước để thêm người thân",
                font=('Segoe UI', 10, 'italic'),
                bg=self.bg_color,
                fg='#666666'
            ).pack(pady=20)
            return
        
        # Load từ database
        try:
            nguoi_than_list = self.db.get_nguoi_than_by_personnel(self.personnel_id)
            
            if not nguoi_than_list:
                tk.Label(
                    self.nguoi_than_frame,
                    text="Chưa có người thân nào. Click '➕ Thêm Người Thân Mới' để thêm.",
                    font=('Segoe UI', 10, 'italic'),
                    bg=self.bg_color,
                    fg='#666666'
                ).pack(pady=20)
                return
            
            # Hiển thị danh sách
            for idx, nguoi_than in enumerate(nguoi_than_list, 1):
                self.create_nguoi_than_item(nguoi_than, idx)
        except Exception as e:
            # Nếu chưa có hàm get_nguoi_than_by_personnel, hiển thị thông báo
            tk.Label(
                self.nguoi_than_frame,
                text="Chưa có người thân nào. Click '➕ Thêm Người Thân Mới' để thêm.",
                font=('Segoe UI', 10, 'italic'),
                bg=self.bg_color,
                fg='#666666'
            ).pack(pady=20)
    
    def create_nguoi_than_item(self, nguoi_than: NguoiThan, stt: int):
        """Tạo item hiển thị người thân - mở rộng để hiển thị đầy đủ"""
        item_frame = tk.Frame(self.nguoi_than_frame, bg=self.section_bg, relief=tk.FLAT, bd=1)
        item_frame.pack(fill=tk.X, pady=8, padx=5)
        
        # Header - tăng chiều cao
        header_frame = tk.Frame(item_frame, bg='#388E3C', height=50)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Title với màu trắng trên nền xanh
        title_label = tk.Label(
            header_frame,
            text=f"{stt}. {nguoi_than.hoTen or 'Chưa có tên'} - {nguoi_than.moiQuanHe or ''}",
            font=('Segoe UI', 12, 'bold'),
            bg='#388E3C',
            fg='white'
        )
        title_label.pack(side=tk.LEFT, padx=15, pady=12)
        
        # Buttons
        btn_container = tk.Frame(header_frame, bg='#388E3C')
        btn_container.pack(side=tk.RIGHT, padx=10)
        
        tk.Button(
            btn_container,
            text="✏️ Sửa",
            command=lambda: self.edit_nguoi_than(nguoi_than),
            font=('Segoe UI', 9),
            bg='#FF9800',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        tk.Button(
            btn_container,
            text="🗑️ Xóa",
            command=lambda: self.delete_nguoi_than(nguoi_than.id),
            font=('Segoe UI', 9),
            bg='#F44336',
            fg='white',
            relief=tk.FLAT,
            padx=12,
            pady=5,
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=3)
        
        # Content - hiển thị từng dòng riêng biệt
        content_frame = tk.Frame(item_frame, bg='#C8E6C9')
        content_frame.pack(fill=tk.X, padx=0, pady=0)
        
        # Tạo từng dòng thông tin riêng biệt
        info_items = []
        
        if nguoi_than.ngaySinh:
            info_items.append(("Ngày sinh:", nguoi_than.ngaySinh))
        if nguoi_than.diaChi:
            info_items.append(("Địa chỉ:", nguoi_than.diaChi))
        if nguoi_than.soDienThoai:
            info_items.append(("Số điện thoại:", nguoi_than.soDienThoai))
        if nguoi_than.noiDung:
            info_items.append(("Nội dung:", nguoi_than.noiDung))
        if nguoi_than.ghiChu:
            info_items.append(("Ghi chú:", nguoi_than.ghiChu))
        
        # Hiển thị từng dòng
        for idx, (label, value) in enumerate(info_items):
            row_frame = tk.Frame(content_frame, bg='#C8E6C9')
            row_frame.pack(fill=tk.X, padx=15, pady=6)
            
            # Label
            tk.Label(
                row_frame,
                text=label,
                font=('Segoe UI', 10, 'bold'),
                bg='#C8E6C9',
                fg='#2E7D32',
                width=15,
                anchor=tk.W
            ).pack(side=tk.LEFT, padx=(0, 10))
            
            # Value - cho phép wrap text
            value_label = tk.Label(
                row_frame,
                text=value,
                font=('Segoe UI', 10),
                bg='#C8E6C9',
                fg='#424242',
                anchor=tk.W,
                justify=tk.LEFT,
                wraplength=800  # Tăng wraplength để hiển thị nhiều hơn
            )
            value_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Nếu không có thông tin nào
        if not info_items:
            tk.Label(
                content_frame,
                text="Chưa có thông tin chi tiết",
                font=('Segoe UI', 9, 'italic'),
                bg='#C8E6C9',
                fg='#666666'
            ).pack(padx=15, pady=10)
    
    def on_yeu_to_nn_changed(self):
        """Xử lý khi checkbox yếu tố nước ngoài thay đổi"""
        if self.yeu_to_nn_var.get():
            # Hiện nút nhập thông tin
            self.yeu_to_nn_btn.pack(side=tk.LEFT, padx=(15, 0))
        else:
            # Ẩn nút
            self.yeu_to_nn_btn.pack_forget()
    
    def open_yeu_to_nn_form(self):
        """Mở form nhập thông tin yếu tố nước ngoài"""
        if not self.yeu_to_nn_var.get():
            messagebox.showwarning("Cảnh báo", "Vui lòng đánh dấu 'Có yếu tố nước ngoài' trước")
            return
        
        # Kiểm tra personnel có ID chưa (phải đã lưu)
        if not self.personnel.id:
            messagebox.showwarning("Cảnh báo", "Vui lòng lưu quân nhân trước khi nhập thông tin yếu tố nước ngoài")
            return
        
        # Lưu tạm các thay đổi hiện tại trước khi mở form
        try:
            # Cập nhật personnel với dữ liệu hiện tại (chưa lưu vào DB)
            self.personnel.thongTinKhac.yeuToNN = True
        except:
            pass
        
        # Mở form yếu tố nước ngoài
        try:
            from gui.yeu_to_nuoc_ngoai_form import YeuToNuocNgoaiFormDialog
            # Lấy personnel mới nhất từ database
            current_personnel = self.db.get_by_id(self.personnel.id)
            if not current_personnel:
                messagebox.showerror("Lỗi", "Không tìm thấy quân nhân trong database")
                return
            
            dialog = YeuToNuocNgoaiFormDialog(self.master, self.db, current_personnel)
            # Đảm bảo focus về parent window sau khi đóng dialog
            self.master.wait_window(dialog.dialog)
            
            # Đảm bảo grab được release
            try:
                self.master.focus_set()
            except:
                pass
            
            # Sau khi đóng form, reload dữ liệu nếu đã lưu
            if dialog.result:
                # Reload personnel từ database để có dữ liệu mới nhất
                updated_personnel = self.db.get_by_id(self.personnel.id)
                if updated_personnel:
                    self.personnel = updated_personnel
                    # Cập nhật lại checkbox
                    self.yeu_to_nn_var.set(self.personnel.thongTinKhac.yeuToNN)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể mở form: {str(e)}")
            import traceback
            traceback.print_exc()
            # Đảm bảo focus được trả về
            try:
                self.master.focus_set()
            except:
                pass
    
    def add_nguoi_than(self):
        """Thêm người thân mới"""
        if not self.personnel_id:
            messagebox.showwarning("Cảnh báo", "Vui lòng lưu quân nhân trước khi thêm người thân")
            return
        
        from gui.nguoi_than_form import NguoiThanFormDialog
        
        dialog_obj = NguoiThanFormDialog(self, self.db, self.personnel_id)
        dialog_obj.show()
        # Đợi dialog đóng rồi reload danh sách
        self.wait_window(dialog_obj.dialog)
        # Reload danh sách ngay sau khi dialog đóng
        self.load_nguoi_than_list()
    
    def edit_nguoi_than(self, nguoi_than: NguoiThan):
        """Sửa người thân"""
        from gui.nguoi_than_form import NguoiThanFormDialog
        
        dialog_obj = NguoiThanFormDialog(self, self.db, self.personnel_id, nguoi_than.id)
        dialog_obj.show()
        # Đợi dialog đóng rồi reload danh sách
        self.wait_window(dialog_obj.dialog)
        # Reload danh sách ngay sau khi dialog đóng
        self.load_nguoi_than_list()
    
    def delete_nguoi_than(self, nguoi_than_id: str):
        """Xóa người thân"""
        if messagebox.askyesno("Xác nhận", "Bạn có chắc muốn xóa người thân này?"):
            try:
                self.db.delete_nguoi_than(nguoi_than_id)
                messagebox.showinfo("Thành công", "Đã xóa người thân")
                self.load_nguoi_than_list()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xóa:\n{str(e)}")
    
    def cancel(self):
        """Hủy và quay lại - Đóng form và quay về danh sách"""
        try:
            logger.debug("Bắt đầu cancel form")
        except:
            pass
        
        # Nếu form đang nằm trong cửa sổ modal (Toplevel) thì chỉ đóng đúng modal
        if isinstance(self.master, tk.Toplevel):
            parent = self.master
            try:
                parent.grab_release()
            except Exception:
                pass
            try:
                parent.destroy()
            except Exception:
                pass
            return

        # Với form gắn trực tiếp vào root: chỉ cần nhờ MainWindow.show_frame('list')
        # để xử lý toàn bộ cleanup giao diện.
        try:
            if hasattr(self.master, 'master') and hasattr(self.master.master, 'show_frame'):
                try:
                    self.master.master.show_frame('list')
                    logger.debug("Đã gọi show_frame('list') thành công")
                except Exception as e:
                    logger.error(f"Lỗi khi gọi show_frame('list'): {e}", exc_info=True)
                    raise
            else:
                # Fallback: cleanup và tạo list frame trực tiếp
                try:
                    self.pack_forget()
                    root.update_idletasks()
                except:
                    pass
                try:
                    self.destroy()
                except:
                    pass
                from gui.personnel_list_frame import PersonnelListFrame
                list_frame = PersonnelListFrame(root, self.db)
                list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                root.update_idletasks()
        except Exception as e:
            # Nếu có lỗi, log và cleanup
            logger.error(f"Lỗi trong cancel(): {e}", exc_info=True)
            import traceback
            traceback.print_exc()
            try:
                root.grab_release()
                root.update_idletasks()
            except:
                pass
