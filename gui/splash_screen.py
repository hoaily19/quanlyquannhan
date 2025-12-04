"""
Cửa sổ loading/splash screen khi khởi động app
"""

import tkinter as tk
from pathlib import Path
from PIL import Image, ImageTk
import math


class SplashScreen:
    """Cửa sổ loading hiển thị khi khởi động app"""
    
    def __init__(self, root, on_close_callback=None, duration=2000):
        """
        Args:
            root: Tkinter root window
            on_close_callback: Callback khi splash screen đóng
            duration: Thời gian hiển thị (ms)
        """
        self.root = root
        self.on_close = on_close_callback
        self.duration = duration
        self.splash = None
        self.image_photo = None
        self.progress_value = 0
        self.loading_dots = 0
        self.animation_running = False
        self.icon_label = None
        self.progress_canvas = None
        
    def show(self):
        """Hiển thị splash screen"""
        # Tạo cửa sổ splash
        self.splash = tk.Toplevel(self.root)
        self.splash.title("")
        self.splash.overrideredirect(True)  # Ẩn thanh tiêu đề
        
        # Kích thước cửa sổ
        splash_width = 500
        splash_height = 500
        
        # Căn giữa màn hình
        screen_width = self.splash.winfo_screenwidth()
        screen_height = self.splash.winfo_screenheight()
        x = (screen_width // 2) - (splash_width // 2)
        y = (screen_height // 2) - (splash_height // 2)
        self.splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")
        
        # Background màu quân đội
        self.splash.configure(bg='#1a4d2e')
        
        # Container chính
        main_frame = tk.Frame(self.splash, bg='#1a4d2e')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Load và hiển thị icon
        try:
            icon_path = Path(__file__).parent.parent / "icons" / "icon.png"
            if icon_path.exists():
                img = Image.open(icon_path)
                # Resize để vừa với cửa sổ (giữ tỷ lệ)
                img_width, img_height = img.size
                max_size = 400
                
                if img_width > img_height:
                    new_width = max_size
                    new_height = int(img_height * (max_size / img_width))
                else:
                    new_height = max_size
                    new_width = int(img_width * (max_size / img_height))
                
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                self.image_photo = ImageTk.PhotoImage(img)
                
                self.icon_label = tk.Label(
                    main_frame,
                    image=self.image_photo,
                    bg='#1a4d2e'
                )
                self.icon_label.image = self.image_photo  # Giữ reference
                self.icon_label.pack(expand=True, pady=20)
            else:
                # Nếu không có icon, hiển thị text
                text_label = tk.Label(
                    main_frame,
                    text="🪖",
                    font=('Arial', 120),
                    bg='#1a4d2e',
                    fg='#ffd700'
                )
                text_label.pack(expand=True, pady=20)
        except Exception as e:
            print(f"Không thể load icon: {e}")
            # Fallback: hiển thị text
            text_label = tk.Label(
                main_frame,
                text="🪖",
                font=('Arial', 120),
                bg='#1a4d2e',
                fg='#ffd700'
            )
            text_label.pack(expand=True, pady=20)
        
        # Text loading
        title_label = tk.Label(
            main_frame,
            text="QUẢN LÝ HỒ SƠ QUÂN NHÂN",
            font=('Arial', 18, 'bold'),
            bg='#1a4d2e',
            fg='#ffd700'
        )
        title_label.pack(pady=(0, 10))
        
        # Loading text với animation dots
        self.loading_label = tk.Label(
            main_frame,
            text="Đang tải",
            font=('Arial', 12),
            bg='#1a4d2e',
            fg='#ffffff'
        )
        self.loading_label.pack(pady=(0, 20))
        
        # Progress bar với animation
        progress_container = tk.Frame(main_frame, bg='#1a4d2e')
        progress_container.pack(pady=(0, 30), padx=50, fill=tk.X)
        
        # Canvas cho progress bar custom
        self.progress_canvas = tk.Canvas(
            progress_container,
            height=6,
            bg='#0d2e1a',
            highlightthickness=0,
            relief=tk.FLAT
        )
        self.progress_canvas.pack(fill=tk.X)
        
        # Ẩn cửa sổ chính tạm thời
        self.root.withdraw()
        
        # Fade in effect
        self.splash.attributes('-alpha', 0.0)
        self.splash.update()
        
        # Cập nhật để hiển thị splash
        self.splash.update()
        
        # Bắt đầu animations
        self.start_animations()
        
        # Tự động đóng sau duration
        self.splash.after(self.duration, self.close)
    
    def start_animations(self):
        """Bắt đầu các hiệu ứng animation"""
        self.animation_running = True
        
        # Fade in effect
        self.fade_in()
        
        # Bắt đầu progress bar animation
        self.animate_progress()
        
        # Bắt đầu loading dots animation
        self.animate_loading_dots()
        
        # Bắt đầu pulse effect cho icon
        if self.icon_label:
            self.animate_pulse()
    
    def fade_in(self, alpha=0.0):
        """Fade in effect cho cửa sổ"""
        if not self.animation_running or not self.splash:
            return
        
        alpha += 0.05
        if alpha >= 1.0:
            alpha = 1.0
            self.splash.attributes('-alpha', alpha)
        else:
            self.splash.attributes('-alpha', alpha)
            self.splash.after(20, lambda: self.fade_in(alpha))
    
    def animate_progress(self, progress=0):
        """Animation cho progress bar"""
        if not self.animation_running or not self.progress_canvas:
            return
        
        # Xóa progress bar cũ
        self.progress_canvas.delete("progress")
        
        # Tính toán vị trí progress
        canvas_width = self.progress_canvas.winfo_width()
        if canvas_width <= 1:
            canvas_width = 400  # Default width
        
        # Sử dụng sine wave để tạo hiệu ứng mượt
        progress_percent = (math.sin(progress * 0.02) + 1) / 2  # 0-1 range
        progress_width = int(canvas_width * progress_percent)
        
        # Vẽ progress bar với gradient effect
        if progress_width > 0:
            # Gradient colors từ vàng đến xanh lá
            colors = ['#ffd700', '#ffed4e', '#c9e265', '#8bc34a', '#4caf50']
            num_segments = len(colors)
            segment_width = progress_width / num_segments
            
            for i, color in enumerate(colors):
                x1 = i * segment_width
                x2 = (i + 1) * segment_width
                if x2 > progress_width:
                    x2 = progress_width
                if x1 < progress_width:
                    self.progress_canvas.create_rectangle(
                        x1, 0, x2, 6,
                        fill=color,
                        outline=color,
                        tags="progress"
                    )
        
        # Tiếp tục animation
        self.splash.after(30, lambda: self.animate_progress(progress + 1))
    
    def animate_loading_dots(self):
        """Animation cho loading dots"""
        if not self.animation_running or not self.loading_label:
            return
        
        dots = "." * (self.loading_dots % 4)
        self.loading_label.config(text=f"Đang tải{dots}")
        self.loading_dots += 1
        
        # Tiếp tục animation
        self.splash.after(500, self.animate_loading_dots)
    
    def animate_pulse(self, scale=1.0, direction=1):
        """Pulse effect cho icon"""
        if not self.animation_running or not self.icon_label:
            return
        
        # Tạo hiệu ứng pulse nhẹ bằng cách thay đổi opacity
        # (Tkinter không hỗ trợ scale trực tiếp, nên dùng cách khác)
        # Thay vào đó, ta có thể thay đổi màu nền hoặc thêm border
        
        # Tiếp tục animation với scale nhẹ
        scale += direction * 0.02
        if scale >= 1.1:
            direction = -1
        elif scale <= 0.95:
            direction = 1
        
        # Có thể thêm hiệu ứng khác ở đây
        self.splash.after(50, lambda: self.animate_pulse(scale, direction))
    
    def close(self):
        """Đóng splash screen với fade out effect"""
        self.animation_running = False
        
        if self.splash:
            # Fade out effect
            self.fade_out()
        else:
            self.finish_close()
    
    def fade_out(self, alpha=1.0):
        """Fade out effect cho cửa sổ"""
        if not self.splash:
            self.finish_close()
            return
        
        alpha -= 0.1
        if alpha <= 0.0:
            alpha = 0.0
            self.splash.attributes('-alpha', alpha)
            self.finish_close()
        else:
            self.splash.attributes('-alpha', alpha)
            self.splash.after(30, lambda: self.fade_out(alpha))
    
    def finish_close(self):
        """Hoàn tất việc đóng splash screen"""
        if self.splash:
            try:
                self.splash.destroy()
            except:
                pass
        
        # Hiển thị lại cửa sổ chính
        self.root.deiconify()
        
        # Gọi callback nếu có
        if self.on_close:
            self.on_close()

import tkinter as tk
from pathlib import Path
from PIL import Image, ImageTk
import math


class SplashScreen:
    """Cửa sổ loading hiển thị khi khởi động app"""
    
    def __init__(self, root, on_close_callback=None, duration=2000):
        """
        Args:
            root: Tkinter root window
            on_close_callback: Callback khi splash screen đóng
            duration: Thời gian hiển thị (ms)
        """
        self.root = root
        self.on_close = on_close_callback
        self.duration = duration
        self.splash = None
        self.image_photo = None
        self.progress_value = 0
        self.loading_dots = 0
        self.animation_running = False
        self.icon_label = None
        self.progress_canvas = None
        
    def show(self):
        """Hiển thị splash screen"""
        # Tạo cửa sổ splash
        self.splash = tk.Toplevel(self.root)
        self.splash.title("")
        self.splash.overrideredirect(True)  # Ẩn thanh tiêu đề
        
        # Kích thước cửa sổ
        splash_width = 500
        splash_height = 500
        
        # Căn giữa màn hình
        screen_width = self.splash.winfo_screenwidth()
        screen_height = self.splash.winfo_screenheight()
        x = (screen_width // 2) - (splash_width // 2)
        y = (screen_height // 2) - (splash_height // 2)
        self.splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")
        
        # Background màu quân đội
        self.splash.configure(bg='#1a4d2e')
        
        # Container chính
        main_frame = tk.Frame(self.splash, bg='#1a4d2e')
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Load và hiển thị icon
        try:
            icon_path = Path(__file__).parent.parent / "icons" / "icon.png"
            if icon_path.exists():
                img = Image.open(icon_path)
                # Resize để vừa với cửa sổ (giữ tỷ lệ)
                img_width, img_height = img.size
                max_size = 400
                
                if img_width > img_height:
                    new_width = max_size
                    new_height = int(img_height * (max_size / img_width))
                else:
                    new_height = max_size
                    new_width = int(img_width * (max_size / img_height))
                
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                self.image_photo = ImageTk.PhotoImage(img)
                
                self.icon_label = tk.Label(
                    main_frame,
                    image=self.image_photo,
                    bg='#1a4d2e'
                )
                self.icon_label.image = self.image_photo  # Giữ reference
                self.icon_label.pack(expand=True, pady=20)
            else:
                # Nếu không có icon, hiển thị text
                text_label = tk.Label(
                    main_frame,
                    text="🪖",
                    font=('Arial', 120),
                    bg='#1a4d2e',
                    fg='#ffd700'
                )
                text_label.pack(expand=True, pady=20)
        except Exception as e:
            print(f"Không thể load icon: {e}")
            # Fallback: hiển thị text
            text_label = tk.Label(
                main_frame,
                text="🪖",
                font=('Arial', 120),
                bg='#1a4d2e',
                fg='#ffd700'
            )
            text_label.pack(expand=True, pady=20)
        
        # Text loading
        title_label = tk.Label(
            main_frame,
            text="QUẢN LÝ HỒ SƠ QUÂN NHÂN",
            font=('Arial', 18, 'bold'),
            bg='#1a4d2e',
            fg='#ffd700'
        )
        title_label.pack(pady=(0, 10))
        
        # Loading text với animation dots
        self.loading_label = tk.Label(
            main_frame,
            text="Đang tải",
            font=('Arial', 12),
            bg='#1a4d2e',
            fg='#ffffff'
        )
        self.loading_label.pack(pady=(0, 20))
        
        # Progress bar với animation
        progress_container = tk.Frame(main_frame, bg='#1a4d2e')
        progress_container.pack(pady=(0, 30), padx=50, fill=tk.X)
        
        # Canvas cho progress bar custom
        self.progress_canvas = tk.Canvas(
            progress_container,
            height=6,
            bg='#0d2e1a',
            highlightthickness=0,
            relief=tk.FLAT
        )
        self.progress_canvas.pack(fill=tk.X)
        
        # Ẩn cửa sổ chính tạm thời
        self.root.withdraw()
        
        # Fade in effect
        self.splash.attributes('-alpha', 0.0)
        self.splash.update()
        
        # Cập nhật để hiển thị splash
        self.splash.update()
        
        # Bắt đầu animations
        self.start_animations()
        
        # Tự động đóng sau duration
        self.splash.after(self.duration, self.close)
    
    def start_animations(self):
        """Bắt đầu các hiệu ứng animation"""
        self.animation_running = True
        
        # Fade in effect
        self.fade_in()
        
        # Bắt đầu progress bar animation
        self.animate_progress()
        
        # Bắt đầu loading dots animation
        self.animate_loading_dots()
        
        # Bắt đầu pulse effect cho icon
        if self.icon_label:
            self.animate_pulse()
    
    def fade_in(self, alpha=0.0):
        """Fade in effect cho cửa sổ"""
        if not self.animation_running or not self.splash:
            return
        
        alpha += 0.05
        if alpha >= 1.0:
            alpha = 1.0
            self.splash.attributes('-alpha', alpha)
        else:
            self.splash.attributes('-alpha', alpha)
            self.splash.after(20, lambda: self.fade_in(alpha))
    
    def animate_progress(self, progress=0):
        """Animation cho progress bar"""
        if not self.animation_running or not self.progress_canvas:
            return
        
        # Xóa progress bar cũ
        self.progress_canvas.delete("progress")
        
        # Tính toán vị trí progress
        canvas_width = self.progress_canvas.winfo_width()
        if canvas_width <= 1:
            canvas_width = 400  # Default width
        
        # Sử dụng sine wave để tạo hiệu ứng mượt
        progress_percent = (math.sin(progress * 0.02) + 1) / 2  # 0-1 range
        progress_width = int(canvas_width * progress_percent)
        
        # Vẽ progress bar với gradient effect
        if progress_width > 0:
            # Gradient colors từ vàng đến xanh lá
            colors = ['#ffd700', '#ffed4e', '#c9e265', '#8bc34a', '#4caf50']
            num_segments = len(colors)
            segment_width = progress_width / num_segments
            
            for i, color in enumerate(colors):
                x1 = i * segment_width
                x2 = (i + 1) * segment_width
                if x2 > progress_width:
                    x2 = progress_width
                if x1 < progress_width:
                    self.progress_canvas.create_rectangle(
                        x1, 0, x2, 6,
                        fill=color,
                        outline=color,
                        tags="progress"
                    )
        
        # Tiếp tục animation
        self.splash.after(30, lambda: self.animate_progress(progress + 1))
    
    def animate_loading_dots(self):
        """Animation cho loading dots"""
        if not self.animation_running or not self.loading_label:
            return
        
        dots = "." * (self.loading_dots % 4)
        self.loading_label.config(text=f"Đang tải{dots}")
        self.loading_dots += 1
        
        # Tiếp tục animation
        self.splash.after(500, self.animate_loading_dots)
    
    def animate_pulse(self, scale=1.0, direction=1):
        """Pulse effect cho icon"""
        if not self.animation_running or not self.icon_label:
            return
        
        # Tạo hiệu ứng pulse nhẹ bằng cách thay đổi opacity
        # (Tkinter không hỗ trợ scale trực tiếp, nên dùng cách khác)
        # Thay vào đó, ta có thể thay đổi màu nền hoặc thêm border
        
        # Tiếp tục animation với scale nhẹ
        scale += direction * 0.02
        if scale >= 1.1:
            direction = -1
        elif scale <= 0.95:
            direction = 1
        
        # Có thể thêm hiệu ứng khác ở đây
        self.splash.after(50, lambda: self.animate_pulse(scale, direction))
    
    def close(self):
        """Đóng splash screen với fade out effect"""
        self.animation_running = False
        
        if self.splash:
            # Fade out effect
            self.fade_out()
        else:
            self.finish_close()
    
    def fade_out(self, alpha=1.0):
        """Fade out effect cho cửa sổ"""
        if not self.splash:
            self.finish_close()
            return
        
        alpha -= 0.1
        if alpha <= 0.0:
            alpha = 0.0
            self.splash.attributes('-alpha', alpha)
            self.finish_close()
        else:
            self.splash.attributes('-alpha', alpha)
            self.splash.after(30, lambda: self.fade_out(alpha))
    
    def finish_close(self):
        """Hoàn tất việc đóng splash screen"""
        if self.splash:
            try:
                self.splash.destroy()
            except:
                pass
        
        # Hiển thị lại cửa sổ chính
        self.root.deiconify()
        
        # Gọi callback nếu có
        if self.on_close:
            self.on_close()

