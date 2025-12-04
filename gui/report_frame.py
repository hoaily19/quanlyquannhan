"""
Frame báo cáo và thống kê
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from services.export import ExportService
from gui.theme import MILITARY_COLORS, get_button_style, get_label_style

# Import matplotlib cho biểu đồ
try:
    import matplotlib
    matplotlib.use('TkAgg')  # Sử dụng TkAgg backend
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    print("Matplotlib không có sẵn. Biểu đồ sẽ không được hiển thị.")


class ReportFrame(tk.Frame):
    """Frame báo cáo thống kê"""
    
    def __init__(self, parent, db: DatabaseService):
        """
        Args:
            parent: Parent widget
            db: DatabaseService instance
        """
        super().__init__(parent)
        self.db = db
        self.setup_ui()
        self.update_stats()
    
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
            text=" 📊 BÁO CÁO TỔNG HỢP",
            font=('Arial', 16, 'bold'),
            bg=MILITARY_COLORS['primary'],
            fg=MILITARY_COLORS['text_light']
        )
        title_label.pack(expand=True)
        
        # Tổng quan
        overview_frame = tk.LabelFrame(
            self,
            text=" Tổng Quan",
            font=('Arial', 12, 'bold'),
            padx=10,
            pady=10,
            bg=MILITARY_COLORS['bg_light'],
            fg=MILITARY_COLORS['primary_dark']
        )
        overview_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.total_label = tk.Label(
            overview_frame,
            text="Tổng Số: 0",
            font=('Arial', 11, 'bold'),
            bg=MILITARY_COLORS['bg_light'],
            fg=MILITARY_COLORS['primary_dark']
        )
        self.total_label.pack(side=tk.LEFT, padx=20)
        
        self.dang_vien_label = tk.Label(
            overview_frame,
            text="Đảng Viên: 0",
            font=('Arial', 11, 'bold'),
            bg=MILITARY_COLORS['bg_light'],
            fg=MILITARY_COLORS['primary_dark']
        )
        self.dang_vien_label.pack(side=tk.LEFT, padx=20)
        
        self.doan_vien_label = tk.Label(
            overview_frame,
            text="Đoàn Viên: 0",
            font=('Arial', 11, 'bold'),
            bg=MILITARY_COLORS['bg_light'],
            fg=MILITARY_COLORS['primary_dark']
        )
        self.doan_vien_label.pack(side=tk.LEFT, padx=20)
        
        self.cd_cu_label = tk.Label(
            overview_frame,
            text="Có Chế Độ Cũ: 0",
            font=('Arial', 11, 'bold'),
            bg=MILITARY_COLORS['bg_light'],
            fg=MILITARY_COLORS['primary_dark']
        )
        self.cd_cu_label.pack(side=tk.LEFT, padx=20)
        
        # Chọn tiêu chí thống kê
        criteria_frame = tk.LabelFrame(
            self,
            text="Chọn Tiêu Chí Thống Kê",
            font=('Arial', 12, 'bold'),
            padx=10,
            pady=10,
            bg=MILITARY_COLORS['bg_light'],
            fg=MILITARY_COLORS['primary_dark']
        )
        criteria_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(
            criteria_frame,
            text="Tiêu chí:",
            font=('Arial', 10, 'bold'),
            bg=MILITARY_COLORS['bg_light'],
            fg=MILITARY_COLORS['text_dark']
        ).pack(side=tk.LEFT, padx=5)
        
        self.criteria_var = tk.StringVar(value="Dân Tộc")
        criteria_combo = ttk.Combobox(
            criteria_frame,
            textvariable=self.criteria_var,
            values=[
                "Dân Tộc",
                "Tôn Giáo",
                "Cấp Bậc",
                "Chức Vụ",
                "Đơn Vị",
                "Đảng Viên",
                "Đoàn Viên"
            ],
            state='readonly',
            width=20
        )
        criteria_combo.pack(side=tk.LEFT, padx=5)
        criteria_combo.bind('<<ComboboxSelected>>', lambda e: self.update_stats())
        
        # Nút xuất CSV
        export_btn = tk.Button(
            criteria_frame,
            text="📥 Xuất CSV",
            command=self.export_csv,
            **get_button_style('success')
        )
        export_btn.pack(side=tk.LEFT, padx=10)
        
        # Kết quả thống kê - Chia làm 2 phần: Biểu đồ và Bảng
        result_frame = tk.LabelFrame(
            self,
            text="Kết Quả Thống Kê",
            font=('Arial', 12, 'bold'),
            padx=10,
            pady=10,
            bg=MILITARY_COLORS['bg_light'],
            fg=MILITARY_COLORS['primary_dark']
        )
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Frame chứa biểu đồ và bảng
        content_frame = tk.Frame(result_frame, bg=MILITARY_COLORS['bg_light'])
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Phần biểu đồ (bên trái)
        chart_frame = tk.Frame(content_frame, bg='white', relief=tk.RAISED, bd=2)
        chart_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        if MATPLOTLIB_AVAILABLE:
            # Tạo figure cho matplotlib
            self.fig = Figure(figsize=(8, 6), dpi=100, facecolor='white')
            self.ax_bar = self.fig.add_subplot(211)  # Biểu đồ cột ở trên
            self.ax_pie = self.fig.add_subplot(212)  # Biểu đồ tròn ở dưới
            
            # Canvas để hiển thị biểu đồ
            self.canvas = FigureCanvasTkAgg(self.fig, chart_frame)
            self.canvas.draw()
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        else:
            # Nếu không có matplotlib, hiển thị thông báo
            no_chart_label = tk.Label(
                chart_frame,
                text="Cần cài đặt matplotlib để hiển thị biểu đồ\npip install matplotlib",
                font=('Arial', 11),
                bg='white',
                fg='red'
            )
            no_chart_label.pack(expand=True)
        
        # Phần bảng (bên phải)
        table_frame = tk.Frame(content_frame, bg='white', relief=tk.RAISED, bd=2)
        table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False, padx=(5, 0))
        table_frame.config(width=350)
        table_frame.pack_propagate(False)
        
        # Treeview để hiển thị kết quả
        columns = ('Tiêu Chí', 'Số Lượng')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor=tk.CENTER)
        
        # Style cho treeview
        style = ttk.Style()
        style.configure("Treeview", 
                       rowheight=35, 
                       font=('Arial', 10),
                       background='white',
                       fieldbackground='white')
        style.configure("Treeview.Heading", 
                       font=('Arial', 11, 'bold'), 
                       background=MILITARY_COLORS['primary'],
                       foreground=MILITARY_COLORS['text_light'])
        style.map("Treeview.Heading",
                 background=[('active', MILITARY_COLORS['primary_dark'])])
        style.map("Treeview",
                 background=[('selected', MILITARY_COLORS['primary_light'])],
                 foreground=[('selected', MILITARY_COLORS['text_dark'])])
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def update_stats(self):
        """Cập nhật thống kê"""
        all_personnel = self.db.get_all()
        
        if not all_personnel:
            # Xóa tree
            for item in self.tree.get_children():
                self.tree.delete(item)
            # Hiển thị thông báo
            self.tree.insert('', 'end', values=('Chưa có dữ liệu', '0'))
            # Xóa biểu đồ nếu có
            if MATPLOTLIB_AVAILABLE:
                try:
                    self.ax_bar.clear()
                    self.ax_pie.clear()
                    self.ax_bar.text(0.5, 0.5, 'Chưa có dữ liệu', 
                                   ha='center', va='center', fontsize=14, 
                                   transform=self.ax_bar.transAxes)
                    self.ax_pie.text(0.5, 0.5, 'Chưa có dữ liệu', 
                                    ha='center', va='center', fontsize=14, 
                                    transform=self.ax_pie.transAxes)
                    self.canvas.draw()
                except:
                    pass
            return
        
        # Tổng quan
        self.total_label.config(text=f"Tổng Số: {len(all_personnel)}")
        
        dang_vien = sum(1 for p in all_personnel 
                       if p.thongTinKhac.dang.ngayVao or p.thongTinKhac.dang.ngayChinhThuc)
        self.dang_vien_label.config(text=f"Đảng Viên: {dang_vien}")
        
        doan_vien = sum(1 for p in all_personnel if p.thongTinKhac.doan.ngayVao)
        self.doan_vien_label.config(text=f"Đoàn Viên: {doan_vien}")
        
        co_cd_cu = sum(1 for p in all_personnel if p.thongTinKhac.cdCu)
        self.cd_cu_label.config(text=f"Có Chế Độ Cũ: {co_cd_cu}")
        
        # Tính toán thống kê theo tiêu chí
        criteria = self.criteria_var.get()
        stats = {}
        
        if criteria == "Dân Tộc":
            for person in all_personnel:
                key = person.danToc or "Chưa xác định"
                stats[key] = stats.get(key, 0) + 1
        
        elif criteria == "Tôn Giáo":
            for person in all_personnel:
                key = person.tonGiao or "Không"
                stats[key] = stats.get(key, 0) + 1
        
        elif criteria == "Cấp Bậc":
            for person in all_personnel:
                key = person.capBac or "Chưa xác định"
                stats[key] = stats.get(key, 0) + 1
        
        elif criteria == "Chức Vụ":
            for person in all_personnel:
                key = person.chucVu or "Chưa xác định"
                stats[key] = stats.get(key, 0) + 1
        
        elif criteria == "Đơn Vị":
            for person in all_personnel:
                key = person.donVi or "Chưa xác định"
                stats[key] = stats.get(key, 0) + 1
        
        elif criteria == "Đảng Viên":
            dang_vien_count = sum(1 for p in all_personnel 
                                 if p.thongTinKhac.dang.ngayVao or p.thongTinKhac.dang.ngayChinhThuc)
            stats["Đảng viên"] = dang_vien_count
            stats["Không phải đảng viên"] = len(all_personnel) - dang_vien_count
        
        elif criteria == "Đoàn Viên":
            doan_vien_count = sum(1 for p in all_personnel if p.thongTinKhac.doan.ngayVao)
            stats["Đoàn viên"] = doan_vien_count
            stats["Không phải đoàn viên"] = len(all_personnel) - doan_vien_count
        
        # Hiển thị kết quả trong bảng
        self.tree.delete(*self.tree.get_children())
        
        # Sắp xếp theo số lượng giảm dần
        sorted_stats = sorted(stats.items(), key=lambda x: x[1], reverse=True)
        
        for key, value in sorted_stats:
            self.tree.insert('', tk.END, values=(key, value))
        
        # Cập nhật biểu đồ
        if MATPLOTLIB_AVAILABLE and sorted_stats:
            self.update_charts(sorted_stats, criteria)
    
    def update_charts(self, sorted_stats, criteria):
        """Cập nhật biểu đồ"""
        if not MATPLOTLIB_AVAILABLE or not sorted_stats:
            return
        
        try:
            # Lấy dữ liệu
            labels = [item[0] for item in sorted_stats]
            values = [item[1] for item in sorted_stats]
            
            # Xóa biểu đồ cũ
            self.ax_bar.clear()
            self.ax_pie.clear()
            
            # Màu sắc cho biểu đồ
            colors = plt.cm.Set3(range(len(labels)))
            
            # Biểu đồ cột (Bar Chart)
            self.ax_bar.bar(labels, values, color=colors, edgecolor='black', linewidth=1.2)
            self.ax_bar.set_title(f'Thống Kê Theo {criteria}', fontsize=12, fontweight='bold', pad=10)
            self.ax_bar.set_xlabel('Tiêu Chí', fontsize=10)
            self.ax_bar.set_ylabel('Số Lượng', fontsize=10)
            self.ax_bar.tick_params(axis='x', rotation=45, labelsize=9)
            self.ax_bar.grid(axis='y', alpha=0.3, linestyle='--')
            
            # Thêm số liệu trên cột
            for i, (label, value) in enumerate(sorted_stats):
                self.ax_bar.text(i, value + 0.05 * max(values), str(value), 
                               ha='center', va='bottom', fontsize=9, fontweight='bold')
            
            # Biểu đồ tròn (Pie Chart)
            self.ax_pie.pie(values, labels=labels, autopct='%1.1f%%', 
                           colors=colors, startangle=90, textprops={'fontsize': 9})
            self.ax_pie.set_title('Tỷ Lệ Phần Trăm', fontsize=12, fontweight='bold', pad=10)
            
            # Cập nhật canvas
            self.fig.tight_layout()
            self.canvas.draw()
            
        except Exception as e:
            print(f"Lỗi khi vẽ biểu đồ: {e}")
            import traceback
            traceback.print_exc()
    
    def export_csv(self):
        """Xuất CSV"""
        all_personnel = self.db.get_all()
        if not all_personnel:
            messagebox.showinfo("Thông báo", "Chưa có dữ liệu để xuất")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                csv_data = ExportService.to_csv(all_personnel)
                with open(file_path, 'w', encoding='utf-8-sig') as f:
                    f.write(csv_data)
                messagebox.showinfo("Thành công", f"Đã xuất file: {file_path}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xuất file: {str(e)}")