"""
Màn hình đăng nhập
"""

import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from PIL import Image, ImageTk
from gui.theme import MILITARY_COLORS, get_button_style, get_label_style


class LoginWindow:
    """Window đăng nhập"""
    
    def __init__(self, root, on_success_callback):
        """
        Args:
            root: Tkinter root window
            on_success_callback: Callback khi đăng nhập thành công
        """
        self.root = root
        self.on_success = on_success_callback
        self.frame = None
        
    def show(self):
        """Hiển thị màn hình đăng nhập"""
        # Frame chính với gradient effect (màu quân đội)
        self.frame = tk.Frame(self.root, bg=MILITARY_COLORS['primary_dark'])
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        # Container chính với padding
        main_container = tk.Frame(self.frame, bg=MILITARY_COLORS['primary_dark'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Container căn giữa
        center_frame = tk.Frame(main_container, bg=MILITARY_COLORS['primary_dark'])
        center_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # Logo/Icon với shadow effect
        logo_container = tk.Frame(center_frame, bg=MILITARY_COLORS['primary_dark'])
        logo_container.pack(pady=(0, 40))
        
        try:
            # Ưu tiên icon.png
            icon_path = Path(__file__).parent.parent / "icons" / "icon.png"
            if not icon_path.exists():
                icon_path = Path(__file__).parent.parent / "logo.jpg"
            
            if icon_path.exists():
                img = Image.open(icon_path)
                # Resize với kích thước lớn hơn và đẹp hơn
                img = img.resize((180, 180), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                logo_label = tk.Label(
                    logo_container,
                    image=photo,
                    bg=MILITARY_COLORS['primary_dark']
                )
                logo_label.image = photo  # Giữ reference
                logo_label.pack()
        except Exception as e:
            print(f"Không thể load logo: {e}")
            # Fallback: hiển thị emoji
            logo_label = tk.Label(
                logo_container,
                text="🪖",
                font=('Arial', 100),
                bg=MILITARY_COLORS['primary_dark'],
                fg=MILITARY_COLORS['gold_light']
            )
            logo_label.pack()
        
        # Title section với spacing tốt hơn
        title_frame = tk.Frame(center_frame, bg=MILITARY_COLORS['primary_dark'])
        title_frame.pack(pady=(0, 35))
        
        # Title chính
        title_label = tk.Label(
            title_frame,
            text="QUẢN LÝ HỒ SƠ QUÂN NHÂN",
            font=('Arial', 24, 'bold'),
            bg=MILITARY_COLORS['primary_dark'],
            fg=MILITARY_COLORS['gold_light']
        )
        title_label.pack(pady=(0, 12))
        
        # Subtitle
        subtitle_label = tk.Label(
            title_frame,
            text="HỆ THỐNG QUẢN LÝ NỘI BỘ",
            font=('Arial', 13),
            bg=MILITARY_COLORS['primary_dark'],
            fg=MILITARY_COLORS['text_light']
        )
        subtitle_label.pack(pady=(0, 8))
        
        # Instruction text
        instruction_label = tk.Label(
            title_frame,
            text="Đăng nhập để tiếp tục",
            font=('Arial', 10),
            bg=MILITARY_COLORS['primary_dark'],
            fg=MILITARY_COLORS['text_gray']
        )
        instruction_label.pack()
        
        # Form đăng nhập với style đẹp
        form_frame = tk.Frame(
            center_frame,
            bg='#2a3a2e',
            padx=45,
            pady=35,
            relief=tk.FLAT,
            bd=0
        )
        form_frame.pack()
        
        # Username section
        username_frame = tk.Frame(form_frame, bg='#2a3a2e')
        username_frame.pack(fill=tk.X, pady=(0, 18))
        
        username_label = tk.Label(
            username_frame,
            text="👤 Tên đăng nhập",
            font=('Arial', 10, 'bold'),
            bg='#2a3a2e',
            fg='#e0e0e0',
            anchor=tk.W
        )
        username_label.pack(fill=tk.X, pady=(0, 8))
        
        self.username_entry = tk.Entry(
            username_frame,
            font=('Arial', 12),
            width=28,
            bg='#ffffff',
            fg='#212121',
            relief=tk.FLAT,
            bd=0,
            insertbackground='#212121',
            highlightthickness=2,
            highlightbackground='#4CAF50',
            highlightcolor='#4CAF50'
        )
        self.username_entry.pack(fill=tk.X, ipady=10)
        self.username_entry.focus()
        
        # Password section
        password_frame = tk.Frame(form_frame, bg='#2a3a2e')
        password_frame.pack(fill=tk.X, pady=(0, 25))
        
        password_label = tk.Label(
            password_frame,
            text="Mật khẩu",
            font=('Arial', 10, 'bold'),
            bg='#2a3a2e',
            fg='#e0e0e0',
            anchor=tk.W
        )
        password_label.pack(fill=tk.X, pady=(0, 8))
        
        self.password_entry = tk.Entry(
            password_frame,
            font=('Arial', 12),
            width=28,
            show='●',
            bg='#ffffff',
            fg='#212121',
            relief=tk.FLAT,
            bd=0,
            insertbackground='#212121',
            highlightthickness=2,
            highlightbackground='#4CAF50',
            highlightcolor='#4CAF50'
        )
        self.password_entry.pack(fill=tk.X, ipady=10)
        
        # Bind Enter key
        self.username_entry.bind('<Return>', lambda e: self.password_entry.focus())
        self.password_entry.bind('<Return>', lambda e: self.handle_login())
        
        # Nút đăng nhập với style đẹp hơn
        login_btn_frame = tk.Frame(form_frame, bg='#2a3a2e')
        login_btn_frame.pack(fill=tk.X)
        
        login_btn = tk.Button(
            login_btn_frame,
            text="ĐĂNG NHẬP",
            command=self.handle_login,
            font=('Arial', 12, 'bold'),
            bg='#4CAF50',
            fg='#ffffff',
            activebackground='#45a049',
            activeforeground='#ffffff',
            relief=tk.FLAT,
            bd=0,
            cursor='hand2',
            padx=20,
            pady=12
        )
        login_btn.pack(fill=tk.X, ipady=8)
        
        # Hover effect cho button
        def on_enter(e):
            login_btn.config(bg='#45a049')
        
        def on_leave(e):
            login_btn.config(bg='#4CAF50')
        
        login_btn.bind('<Enter>', on_enter)
        login_btn.bind('<Leave>', on_leave)
        
        # Thông tin mặc định với style đẹp hơn
        info_frame = tk.Frame(center_frame, bg=MILITARY_COLORS['primary_dark'])
        info_frame.pack(pady=(25, 0))
        
        info_label = tk.Label(
            info_frame,
            text="Mặc định: admin / admin123",
            font=('Arial', 9),
            bg=MILITARY_COLORS['primary_dark'],
            fg='#9e9e9e'
        )
        info_label.pack()
    
    def handle_login(self):
        """Xử lý đăng nhập"""
        try:
            username = self.username_entry.get().strip()
            password = self.password_entry.get()
        except Exception as e:
            print(f"Lỗi khi lấy thông tin đăng nhập: {e}")
            messagebox.showerror("Lỗi", "Có lỗi xảy ra. Vui lòng thử lại.")
            return
        
        if not username or not password:
            messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin")
            return
        
        from services.auth import AuthService
        auth = AuthService()
        
        if auth.login(username, password):
            # Sử dụng after với delay nhỏ để đảm bảo callback được gọi sau khi xử lý hoàn tất
            self.root.after(50, self.on_success)
        else:
            messagebox.showerror("Lỗi", "Tên đăng nhập hoặc mật khẩu không đúng")

