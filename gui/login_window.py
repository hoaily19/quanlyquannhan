"""
Màn hình đăng nhập - Phong cách Garena với màu xanh quân đội
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
        self.remember_password = tk.BooleanVar(value=False)
        
    def show(self):
        """Hiển thị màn hình đăng nhập"""
        # Frame chính
        self.frame = tk.Frame(self.root, bg='#0a1a0f')
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        # Container chính với layout 2 cột
        main_container = tk.Frame(self.frame, bg='#0a1a0f')
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # === CỘT TRÁI: Background với logo và title ===
        left_panel = tk.Frame(main_container, bg='#0d2e1a')
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Container nội dung bên trái
        left_content = tk.Frame(left_panel, bg='#0d2e1a')
        left_content.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # Logo
        logo_container = tk.Frame(left_content, bg='#0d2e1a')
        logo_container.pack(pady=(0, 30))
        
        try:
            # Ưu tiên icon.png
            icon_path = Path(__file__).parent.parent / "icons" / "icon.png"
            if not icon_path.exists():
                icon_path = Path(__file__).parent.parent / "logo.jpg"
            
            if icon_path.exists():
                img = Image.open(icon_path)
                # Resize với kích thước lớn hơn
                img = img.resize((200, 200), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                logo_label = tk.Label(
                    logo_container,
                    image=photo,
                    bg='#0d2e1a'
                )
                logo_label.image = photo  # Giữ reference
                logo_label.pack()
        except Exception as e:
            print(f"Không thể load logo: {e}")
            # Fallback: hiển thị emoji
            logo_label = tk.Label(
                logo_container,
                text="🪖",
                font=('Arial', 120),
                bg='#0d2e1a',
                fg=MILITARY_COLORS['gold_light']
            )
            logo_label.pack()
        
        # Title section
        title_frame = tk.Frame(left_content, bg='#0d2e1a')
        title_frame.pack()
        
        # "Welcome to the" text
        welcome_label = tk.Label(
            title_frame,
            text="Welcome to the",
            font=('Arial', 14),
            bg='#0d2e1a',
            fg='#81C784'
        )
        welcome_label.pack()
        
        # "New" text
        new_label = tk.Label(
            title_frame,
            text="New",
            font=('Arial', 20, 'bold'),
            bg='#0d2e1a',
            fg='#4CAF50'
        )
        new_label.pack()
        
        # Main title
        title_label = tk.Label(
            title_frame,
            text="QUẢN LÝ HỒ SƠ QUÂN NHÂN",
            font=('Arial', 28, 'bold'),
            bg='#0d2e1a',
            fg='#4CAF50'
        )
        title_label.pack(pady=(5, 0))
        
        # Subtitle
        subtitle_label = tk.Label(
            title_frame,
            text="HỆ THỐNG QUẢN LÝ NỘI BỘ",
            font=('Arial', 12),
            bg='#0d2e1a',
            fg='#A5D6A7'
        )
        subtitle_label.pack(pady=(10, 0))
        
        # === CỘT PHẢI: Form đăng nhập ===
        right_panel = tk.Frame(main_container, bg='#1B5E20', width=500)
        right_panel.pack(side=tk.RIGHT, fill=tk.Y, expand=False, ipadx=60, ipady=40)
        
        # Container form - dùng pack để căn giữa
        form_container = tk.Frame(right_panel, bg='#1B5E20')
        form_container.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        # Logo nhỏ ở trên cùng form
        logo_small_frame = tk.Frame(form_container, bg='#1B5E20')
        logo_small_frame.pack(pady=(0, 30))
        
        try:
            icon_path = Path(__file__).parent.parent / "icons" / "icon.png"
            if not icon_path.exists():
                icon_path = Path(__file__).parent.parent / "logo.jpg"
            
            if icon_path.exists():
                img = Image.open(icon_path)
                img = img.resize((80, 80), Image.Resampling.LANCZOS)
                photo_small = ImageTk.PhotoImage(img)
                logo_small = tk.Label(
                    logo_small_frame,
                    image=photo_small,
                    bg='#1B5E20'
                )
                logo_small.image = photo_small
                logo_small.pack()
        except:
            pass
        
        # Title "QUẢN LÝ HỒ SƠ QUÂN NHÂN" trên form
        form_title = tk.Label(
            form_container,
            text="QUẢN LÝ HỒ SƠ QUÂN NHÂN",
            font=('Arial', 14, 'bold'),
            bg='#1B5E20',
            fg='#FFFFFF',
            wraplength=300
        )
        form_title.pack(pady=(0, 40))
        
        # Username field với dropdown style
        username_frame = tk.Frame(form_container, bg='#1B5E20')
        username_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Username entry với style đẹp
        self.username_entry = tk.Entry(
            username_frame,
            font=('Arial', 11),
            bg='#2E7D32',
            fg='#FFFFFF',
            relief=tk.FLAT,
            bd=0,
            insertbackground='#FFFFFF',
            highlightthickness=1,
            highlightbackground='#4CAF50',
            highlightcolor='#66BB6A'
        )
        self.username_entry.pack(fill=tk.X, ipady=12)
        self.username_entry.insert(0, "Username")  # Placeholder
        self.username_entry.config(fg='#B0BEC5')  # Màu placeholder
        self.username_entry.focus()
        
        # Bind events cho username
        def on_username_focus_in(e):
            if self.username_entry.get() == "Username":
                self.username_entry.delete(0, tk.END)
                self.username_entry.config(fg='#FFFFFF')
        
        def on_username_focus_out(e):
            if not self.username_entry.get():
                self.username_entry.insert(0, "Username")
                self.username_entry.config(fg='#B0BEC5')
        
        self.username_entry.bind('<FocusIn>', on_username_focus_in)
        self.username_entry.bind('<FocusOut>', on_username_focus_out)
        
        # Password field
        password_frame = tk.Frame(form_container, bg='#1B5E20')
        password_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.password_entry = tk.Entry(
            password_frame,
            font=('Arial', 11),
            show='●',
            bg='#2E7D32',
            fg='#FFFFFF',
            relief=tk.FLAT,
            bd=0,
            insertbackground='#FFFFFF',
            highlightthickness=1,
            highlightbackground='#4CAF50',
            highlightcolor='#66BB6A'
        )
        self.password_entry.pack(fill=tk.X, ipady=12)
        
        # Remember Password checkbox
        remember_frame = tk.Frame(form_container, bg='#1B5E20')
        remember_frame.pack(fill=tk.X, pady=(0, 20))
        
        remember_check = tk.Checkbutton(
            remember_frame,
            text="Remember Password",
            variable=self.remember_password,
            font=('Arial', 10),
            bg='#1B5E20',
            fg='#FFFFFF',
            activebackground='#1B5E20',
            activeforeground='#FFFFFF',
            selectcolor='#2E7D32',
            cursor='hand2'
        )
        remember_check.pack(side=tk.LEFT)
        
        # Login button
        login_btn = tk.Button(
            form_container,
            text="LOGIN",
            command=self.handle_login,
            font=('Arial', 12, 'bold'),
            bg='#4CAF50',
            fg='#FFFFFF',
            activebackground='#45a049',
            activeforeground='#FFFFFF',
            relief=tk.FLAT,
            bd=0,
            cursor='hand2',
            padx=20,
            pady=15
        )
        login_btn.pack(fill=tk.X, pady=(0, 15))
        
        # Hover effect cho button
        def on_enter(e):
            login_btn.config(bg='#66BB6A')
        
        def on_leave(e):
            login_btn.config(bg='#4CAF50')
        
        login_btn.bind('<Enter>', on_enter)
        login_btn.bind('<Leave>', on_leave)
        
        # Create Account button
        create_account_btn = tk.Button(
            form_container,
            text="CREATE ACCOUNT",
            command=lambda: messagebox.showinfo("Thông báo", "Tính năng đang phát triển"),
            font=('Arial', 10),
            bg='#2C2C2C',
            fg='#FFFFFF',
            activebackground='#3C3C3C',
            activeforeground='#FFFFFF',
            relief=tk.FLAT,
            bd=0,
            cursor='hand2',
            padx=20,
            pady=12
        )
        create_account_btn.pack(fill=tk.X, pady=(0, 15))
        
        # Forgot Password link
        forgot_frame = tk.Frame(form_container, bg='#1B5E20')
        forgot_frame.pack(fill=tk.X)
        
        forgot_link = tk.Label(
            forgot_frame,
            text="Forgot Password?",
            font=('Arial', 9),
            bg='#1B5E20',
            fg='#81C784',
            cursor='hand2'
        )
        forgot_link.pack(side=tk.RIGHT)
        
        def on_forgot_click(e):
            messagebox.showinfo("Quên mật khẩu", "Mặc định: admin / admin123")
        
        forgot_link.bind('<Button-1>', on_forgot_click)
        
        # Default credentials info
        info_frame = tk.Frame(form_container, bg='#1B5E20')
        info_frame.pack(pady=(20, 0))
        
        info_label = tk.Label(
            info_frame,
            text="Mặc định: admin / admin123",
            font=('Arial', 9),
            bg='#1B5E20',
            fg='#A5D6A7'
        )
        info_label.pack()
        
        # Bind Enter key
        self.username_entry.bind('<Return>', lambda e: self.password_entry.focus())
        self.password_entry.bind('<Return>', lambda e: self.handle_login())
    
    def handle_login(self):
        """Xử lý đăng nhập"""
        try:
            username = self.username_entry.get().strip()
            password = self.password_entry.get()
            
            # Xử lý placeholder text
            if username == "Username":
                username = ""
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
            # Truyền username vào callback
            self.root.after(50, lambda: self.on_success(username))
        else:
            messagebox.showerror("Lỗi", "Tên đăng nhập hoặc mật khẩu không đúng")

