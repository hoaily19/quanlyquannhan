"""
Ứng dụng Quản Lý Hồ Sơ Quân Nhân - Desktop App
Entry point chính
"""

import tkinter as tk
from tkinter import ttk, messagebox
import sys
from pathlib import Path

# Thêm thư mục hiện tại vào path
sys.path.insert(0, str(Path(__file__).parent))

from gui.main_window import MainWindow
from gui.splash_screen import SplashScreen
from services.auth import AuthService


class App:
    """Lớp chính của ứng dụng"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.auth_service = AuthService()
        self.setup_window()
        
    def setup_window(self):
        """Thiết lập cửa sổ chính"""
        self.root.title("🪖 Quản Lý Hồ Sơ Quân Nhân")
        # Tăng chiều cao để hiển thị nhiều dữ liệu hơn
        self.root.geometry("1400x1100")
        self.root.minsize(1200, 800)
        
        # Set background color
        self.root.configure(bg='#ecf0f1')
        
        # Set icon (nếu có)
        try:
            from PIL import Image, ImageTk
            
            # Ưu tiên 1: icon.png trong thư mục icons
            icon_png_path = Path(__file__).parent / "icons" / "icon.png"
            if icon_png_path.exists():
                img = Image.open(icon_png_path)
                photo = ImageTk.PhotoImage(img)
                self.root.iconphoto(True, photo)
            else:
                # Ưu tiên 2: logo.ico trong thư mục icons
                ico_path = Path(__file__).parent / "icons" / "logo.ico"
                if ico_path.exists():
                    self.root.iconbitmap(str(ico_path))
                else:
                    # Fallback: logo.jpg ở root
                    logo_path = Path(__file__).parent / "logo.jpg"
                    if logo_path.exists():
                        img = Image.open(logo_path)
                        photo = ImageTk.PhotoImage(img)
                        self.root.iconphoto(True, photo)
        except Exception as e:
            print(f"Không thể load icon: {e}")
        
        # Center window
        self.center_window()
        
    def center_window(self):
        """Căn giữa cửa sổ trên màn hình"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def show_login(self):
        """Hiển thị màn hình đăng nhập"""
        # Xóa tất cả widgets hiện tại một cách an toàn
        try:
            for widget in list(self.root.winfo_children()):
                try:
                    widget.destroy()
                except:
                    pass
            self.root.update_idletasks()
        except:
            pass
        
        from gui.login_window import LoginWindow
        login_window = LoginWindow(self.root, self.on_login_success)
        login_window.show()
    
    def on_login_success(self):
        """Callback khi đăng nhập thành công"""
        # Sử dụng after với delay nhỏ để đảm bảo destroy được thực hiện sau khi callback hoàn thành
        self.root.after(100, self._cleanup_and_show_main)
    
    def _cleanup_and_show_main(self):
        """Xóa login window và hiển thị main window một cách an toàn"""
        try:
            # Xóa tất cả widgets hiện tại một cách an toàn
            for widget in list(self.root.winfo_children()):
                try:
                    widget.destroy()
                except:
                    pass
            
            # Update để đảm bảo destroy được thực hiện
            self.root.update_idletasks()
            
            # Hiển thị main window
            main_window = MainWindow(self.root)
            main_window.show()
        except Exception as e:
            print(f"Lỗi khi chuyển sang main window: {e}")
            import traceback
            traceback.print_exc()
    
    def run(self):
        """Chạy ứng dụng"""
        # Hiển thị splash screen trước
        splash = SplashScreen(self.root, on_close_callback=self.show_login, duration=2000)
        splash.show()
        
        # Chạy main loop
        self.root.mainloop()


def main():
    """Hàm main"""
    app = App()
    app.run()


if __name__ == "__main__":
    main()

