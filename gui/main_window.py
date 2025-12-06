"""
Cửa sổ chính của ứng dụng
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from PIL import Image, ImageTk
import sys
import logging

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from services.export import ExportService
from services.discord_bot import get_discord_bot
from gui.personnel_list_frame import PersonnelListFrame
from gui.personnel_form_frame import PersonnelFormFrame
from gui.report_frame import ReportFrame
from gui.import_frame import ImportFrame
from gui.reports_list_frame import ReportsListFrame
from gui.theme import MILITARY_COLORS, get_button_style
from gui.tooltip import create_tooltip

# Setup logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app_focus.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class MainWindow:
    """Cửa sổ chính"""
    
    def __init__(self, root):
        """
        Args:
            root: Tkinter root window
        """
        self.root = root
        self.db = DatabaseService()
        self.edit_personnel_id = None  # Lưu ID khi edit
        self.current_username = ""  # Username từ đăng nhập
        self.setup_menu()
        self.current_frame = None
        
        # Khởi động Discord bot
        try:
            self.discord_bot = get_discord_bot()
            
            # Thiết lập callbacks để điều khiển ứng dụng từ Discord
            self.discord_bot.set_shutdown_callback(self._shutdown_app)
            self.discord_bot.set_restart_callback(self._restart_app)
            
            self.discord_bot.start()
            logger.info("Đã khởi động Discord bot")
            
            # Gửi thông báo khi ứng dụng khởi động (sau một chút để bot kết nối)
            self.root.after(3000, lambda: self._notify_app_started())
        except Exception as e:
            logger.error(f"Lỗi khi khởi động Discord bot: {str(e)}")
            self.discord_bot = None
    
    def _shutdown_app(self):
        """Tắt ứng dụng (được gọi từ Discord bot)"""
        logger.warning("⚠️ Nhận lệnh tắt ứng dụng từ Discord")
        try:
            if self.discord_bot:
                self.discord_bot.send_notification(
                    "🛑 Ứng Dụng Đã Tắt",
                    "Ứng dụng đã được tắt từ xa qua Discord",
                    color=0xF44336
                )
            # Đợi một chút để gửi thông báo
            self.root.after(1000, self.root.quit)
        except Exception as e:
            logger.error(f"Lỗi khi tắt ứng dụng: {str(e)}")
            self.root.quit()
    
    def _restart_app(self):
        """Khởi động lại ứng dụng (được gọi từ Discord bot)"""
        logger.warning("⚠️ Nhận lệnh khởi động lại ứng dụng từ Discord")
        try:
            if self.discord_bot:
                self.discord_bot.send_notification(
                    "🔄 Đang Khởi Động Lại",
                    "Ứng dụng đang được khởi động lại...",
                    color=0xFF9800
                )
            # Đợi một chút để gửi thông báo, sau đó restart
            self.root.after(2000, lambda: self.root.quit())
            # Note: Để restart thực sự, cần có script wrapper hoặc system call
        except Exception as e:
            logger.error(f"Lỗi khi khởi động lại ứng dụng: {str(e)}")
    
    def _notify_app_started(self):
        """Gửi thông báo khi ứng dụng khởi động"""
        try:
            if self.discord_bot:
                # Sử dụng username từ đăng nhập hoặc lấy từ hệ thống
                username = self.current_username
                if not username:
                    import os
                    username = os.getenv('USERNAME') or os.getenv('USER') or ''
                self.discord_bot.notify_app_started(username)
        except Exception as e:
            logger.error(f"Lỗi khi gửi thông báo khởi động: {str(e)}")
        
    def setup_menu(self):
        """Thiết lập menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Menu File
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Xuất CSV...", command=self.export_csv)
        file_menu.add_command(label="Xuất PDF...", command=self.export_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="Thoát", command=self.root.quit)
        
        # Menu Quản Lý
        manage_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Quản Lý", menu=manage_menu)
        manage_menu.add_command(label="Danh Sách Quân Nhân", command=lambda: self.show_frame('list'))
        manage_menu.add_command(label="Thêm Quân Nhân", command=lambda: self.show_frame('add'))
        manage_menu.add_command(label="Nhập Dữ Liệu", command=lambda: self.show_frame('import'))
        
        # Menu Báo Cáo
        report_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Báo Cáo", menu=report_menu)
        report_menu.add_command(label="Thống Kê Tổng Hợp", command=lambda: self.show_frame('report'))
        report_menu.add_command(label="Danh Sách Báo Cáo", command=lambda: self.show_frame('reports_list'))
        
        # Menu Hệ Thống
        system_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Hệ Thống", menu=system_menu)
        system_menu.add_command(label="Test Discord Bot", command=self.test_discord_bot)
        system_menu.add_separator()
        system_menu.add_command(label="Đăng Xuất", command=self.logout)
        system_menu.add_command(label="Thoát", command=self.root.quit)
    
    def show_frame(self, frame_name: str):
        """Hiển thị frame tương ứng - Đảm bảo tính đồng nhất và xử lý lỗi tốt hơn"""
        try:
            logger.debug(f"Bắt đầu chuyển sang frame: {frame_name}")
        except:
            pass
        
        # Bước 1: Release tất cả grab trước
        try:
            # Release grab từ root window
            self.root.grab_release()
            logger.debug("Đã release grab từ root window")
        except Exception as e:
            logger.error(f"Lỗi khi release grab từ root: {e}", exc_info=True)
        
        # Bước 2: Đóng tất cả dialog và DatePicker popup đang mở trước
        try:
            # Tìm tất cả Toplevel windows (bao gồm cả dialog con và DatePicker popup)
            all_toplevels = []
            def find_toplevels(widget):
                if isinstance(widget, tk.Toplevel):
                    all_toplevels.append(widget)
                try:
                    for child in widget.winfo_children():
                        find_toplevels(child)
                except:
                    pass
            
            find_toplevels(self.root)
            
            # Release grab và destroy tất cả - đảm bảo đóng theo thứ tự đúng
            for toplevel in all_toplevels:
                try:
                    # Release grab trước
                    toplevel.grab_release()
                except:
                    pass
                try:
                    # Update để đảm bảo release được xử lý
                    toplevel.update_idletasks()
                except:
                    pass
            
            # Destroy tất cả sau khi release grab
            for toplevel in all_toplevels:
                try:
                    toplevel.destroy()
                except:
                    pass
            
            # Update nhiều lần để đảm bảo UI được render lại đúng
            try:
                self.root.update_idletasks()
                self.root.update()
                self.root.update_idletasks()
            except:
                pass
        except:
            pass
        
        # Bước 3: Xóa frame hiện tại một cách an toàn
        if self.current_frame:
            # Lưu reference frame cũ để destroy sau
            old_frame = self.current_frame
            self.current_frame = None  # Set None trước để tránh conflict
            
            try:
                # Đóng tất cả dialog con trong frame trước
                def close_all_dialogs_in_frame(widget):
                    """Đệ quy đóng tất cả dialog trong frame"""
                    if isinstance(widget, tk.Toplevel):
                        try:
                            widget.grab_release()
                            widget.destroy()
                        except:
                            pass
                    try:
                        for child in widget.winfo_children():
                            close_all_dialogs_in_frame(child)
                    except:
                        pass
                
                close_all_dialogs_in_frame(old_frame)
            except:
                pass
            
            try:
                # Unpack trước
                old_frame.pack_forget()
            except:
                pass
            
            # Update để đảm bảo unpack được xử lý
            try:
                self.root.update_idletasks()
                self.root.update()
            except:
                pass
            
            # Destroy frame sau khi đã unpack
            try:
                old_frame.destroy()
            except:
                pass
            
            # Update lại để đảm bảo cleanup hoàn toàn
            try:
                self.root.update_idletasks()
                self.root.update()
                self.root.update_idletasks()
            except:
                pass
        
        # Bước 4: Đảm bảo focus về root window
        try:
            self.root.focus_set()
            self.root.update_idletasks()
            logger.debug("Đã set focus về root window")
        except Exception as e:
            logger.error(f"Lỗi khi set focus về root window: {e}", exc_info=True)
        
        # Bước 5: Đảm bảo root window có background đúng
        try:
            self.root.configure(bg=MILITARY_COLORS['bg_light'])
        except:
            pass
        
        # Bước 6: Tạo frame mới với xử lý lỗi tốt hơn
        try:
            # Đảm bảo root window sẵn sàng
            self.root.update_idletasks()
            
            if frame_name == 'list':
                self.current_frame = PersonnelListFrame(self.root, self.db)
            elif frame_name == 'add':
                self.current_frame = PersonnelFormFrame(self.root, self.db, is_new=True)
            elif frame_name == 'edit':
                # Edit với personnel_id đã lưu
                if self.edit_personnel_id:
                    self.current_frame = PersonnelFormFrame(self.root, self.db, personnel_id=self.edit_personnel_id)
                    self.edit_personnel_id = None  # Reset
                else:
                    messagebox.showwarning("Cảnh báo", "Không có quân nhân được chọn để sửa")
                    return
            elif frame_name == 'report':
                self.current_frame = ReportFrame(self.root, self.db)
            elif frame_name == 'reports_list':
                self.current_frame = ReportsListFrame(self.root, self.db)
            elif frame_name == 'import':
                self.current_frame = ImportFrame(self.root, self.db)
            else:
                messagebox.showwarning("Cảnh báo", f"Chức năng '{frame_name}' không tồn tại")
                return  # Frame name không hợp lệ
            
            if self.current_frame:
                # Đảm bảo frame có background đúng
                try:
                    self.current_frame.configure(bg=MILITARY_COLORS['bg_light'])
                except:
                    pass
                
                # Form frame không có padding để tràn viền
                if frame_name in ['add', 'edit']:
                    self.current_frame.pack(fill=tk.BOTH, expand=True)
                else:
                    self.current_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                # Update nhiều lần để đảm bảo UI được render đúng
                try:
                    self.root.update_idletasks()
                    self.root.update()
                    self.current_frame.update_idletasks()
                    self.root.update_idletasks()
                    self.root.update()
                except:
                    pass
                
                # Đảm bảo root window có focus và được hiển thị đúng
                try:
                    self.root.focus_set()
                    self.root.lift()
                    logger.debug(f"Đã set focus và lift root window sau khi tạo frame {frame_name}")
                except Exception as e:
                    logger.error(f"Lỗi khi set focus/lift root window: {e}", exc_info=True)
        except Exception as e:
            error_msg = f"Không thể hiển thị trang '{frame_name}': {str(e)}"
            logger.error(f"Lỗi khi hiển thị frame '{frame_name}': {e}", exc_info=True)
            try:
                messagebox.showerror("Lỗi", error_msg)
            except:
                pass
            import traceback
            print(f"Lỗi khi hiển thị frame '{frame_name}':")
            traceback.print_exc()
            # Đảm bảo luôn có frame hiển thị (fallback về list)
            if frame_name != 'list':
                try:
                    self.show_frame('list')
                except:
                    pass
    
    def show(self):
        """Hiển thị cửa sổ chính"""
        # Configure root background
        self.root.configure(bg=MILITARY_COLORS['bg_light'])
        
        # Tạo toolbar với màu quân đội
        toolbar = tk.Frame(
            self.root,
            bg=MILITARY_COLORS['primary_dark'],
            height=60,
            relief=tk.RAISED,
            bd=3
        )
        toolbar.pack(fill=tk.X, side=tk.TOP)
        toolbar.pack_propagate(False)
        
        # Logo/Title trên toolbar
        title_label = tk.Label(
            toolbar,
            text="🪖 HỆ THỐNG QUẢN LÝ HỒ SƠ QUÂN NHÂN",
            font=('Arial', 12, 'bold'),
            bg=MILITARY_COLORS['primary_dark'],
            fg=MILITARY_COLORS['text_light']
        )
        title_label.pack(side=tk.LEFT, padx=15)
        
        # Buttons trên toolbar với style quân đội - Hiển thị icon + text
        btn_frame = tk.Frame(toolbar, bg=MILITARY_COLORS['primary_dark'])
        btn_frame.pack(side=tk.RIGHT, padx=10)
        
        btn_list = tk.Button(
            btn_frame,
            text="📋 Danh Sách",
            command=lambda: self.show_frame('list'),
            **get_button_style('primary')
        )
        btn_list.pack(side=tk.LEFT, padx=3)
        
        btn_add = tk.Button(
            btn_frame,
            text="➕ Thêm Mới",
            command=lambda: self.show_frame('add'),
            **get_button_style('success')
        )
        btn_add.pack(side=tk.LEFT, padx=3)
        
        btn_report = tk.Button(
            btn_frame,
            text="📊 Báo Cáo",
            command=lambda: self.show_frame('report'),
            **get_button_style('secondary')
        )
        btn_report.pack(side=tk.LEFT, padx=3)
        
        btn_reports_list = tk.Button(
            btn_frame,
            text="📑 Danh Sách Báo Cáo",
            command=lambda: self.show_frame('reports_list'),
            **get_button_style('info')
        )
        btn_reports_list.pack(side=tk.LEFT, padx=3)
        
        btn_import = tk.Button(
            btn_frame,
            text="📥 Nhập Dữ Liệu",
            command=lambda: self.show_frame('import'),
            **get_button_style('accent')
        )
        btn_import.pack(side=tk.LEFT, padx=3)
        
        # Hiển thị frame mặc định
        self.show_frame('list')
    
    def export_csv(self):
        """Xuất CSV"""
        from tkinter import filedialog
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
    
    def export_pdf(self):
        """Xuất PDF"""
        from tkinter import filedialog
        all_personnel = self.db.get_all()
        if not all_personnel:
            messagebox.showinfo("Thông báo", "Chưa có dữ liệu để xuất")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                pdf_data = ExportService.to_pdf(all_personnel)
                with open(file_path, 'wb') as f:
                    f.write(pdf_data)
                messagebox.showinfo("Thành công", f"Đã xuất file: {file_path}")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể xuất file: {str(e)}")
    
    def test_discord_bot(self):
        """Test kết nối Discord bot"""
        try:
            if not self.discord_bot:
                messagebox.showwarning("Cảnh báo", "Discord bot chưa được khởi động")
                return
            
            result = self.discord_bot.test_connection()
            if result:
                messagebox.showinfo("Thành công", 
                    "✅ Bot đã kết nối thành công!\n"
                    "Đã gửi thông báo test lên Discord.\n"
                    "Vui lòng kiểm tra channel trên Discord.")
            else:
                messagebox.showwarning("Cảnh báo", 
                    "❌ Bot chưa kết nối hoặc chưa có channel.\n"
                    "Vui lòng kiểm tra:\n"
                    "1. Bot đã được mời vào server chưa?\n"
                    "2. Channel ID có đúng không?\n"
                    "3. Bot có quyền gửi tin nhắn không?\n"
                    "4. Xem log để biết chi tiết lỗi.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể test bot: {str(e)}")
            logger.error(f"Lỗi khi test Discord bot: {str(e)}", exc_info=True)
    
    def logout(self):
        """Đăng xuất"""
        from services.auth import AuthService
        auth = AuthService()
        auth.logout()
        
        # Quay lại màn hình login
        from gui.login_window import LoginWindow
        for widget in self.root.winfo_children():
            widget.destroy()
        login_window = LoginWindow(self.root, self.on_login_success)
        login_window.show()
    
    def on_login_success(self):
        """Callback khi đăng nhập lại thành công"""
        for widget in self.root.winfo_children():
            widget.destroy()
        self.show()