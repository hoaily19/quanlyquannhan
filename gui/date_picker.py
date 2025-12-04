"""
Date picker component cho Tkinter - Có chọn năm
"""

import tkinter as tk
from tkinter import ttk
from datetime import datetime
from calendar import monthcalendar, month_name


class DatePicker:
    """Date picker widget với chọn năm"""
    
    def __init__(self, parent, default_date=None):
        """
        Args:
            parent: Parent widget
            default_date: Date string in DD/MM/YYYY format
        """
        self.parent = parent
        self.selected_date = None
        self.result = None
        
        # Parse default date
        if default_date:
            try:
                self.selected_date = datetime.strptime(default_date, "%d/%m/%Y")
            except:
                self.selected_date = datetime.now()
        else:
            self.selected_date = datetime.now()
        
        self.create_widget()
    
    def create_widget(self):
        """Tạo date picker widget - mềm mại, dễ nhìn"""
        # Frame chính - full width
        self.frame = tk.Frame(self.parent)
        self.frame.pack(fill=tk.X, expand=True)
        
        # Entry để hiển thị - style mềm mại
        self.entry_var = tk.StringVar()
        self.entry_var.set(self.selected_date.strftime("%d/%m/%Y"))
        
        self.entry = tk.Entry(
            self.frame,
            textvariable=self.entry_var,
            font=('Segoe UI', 10),
            justify=tk.CENTER,
            relief=tk.FLAT,
            bd=1,
            bg='white',
            fg='#424242',
            insertbackground='#388E3C',
            highlightthickness=1,
            highlightcolor='#388E3C',
            highlightbackground='#E0E0E0'
        )
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8), ipady=8)
        
        # Nút calendar - mềm mại hơn
        self.cal_btn = tk.Button(
            self.frame,
            text="📅",
            command=self.show_calendar,
            font=('Segoe UI', 11),
            width=5,
            height=1,
            cursor='hand2',
            bg='#388E3C',
            fg='white',
            activebackground='#2E7D32',
            activeforeground='white',
            relief=tk.FLAT,
            bd=0
        )
        self.cal_btn.pack(side=tk.LEFT)
        
        # Bind để validate format
        self.entry_var.trace('w', self.validate_date)
    
    def validate_date(self, *args):
        """Validate date format"""
        value = self.entry_var.get()
        try:
            if value:
                datetime.strptime(value, "%d/%m/%Y")
        except:
            pass  # Invalid format, user đang nhập
    
    def show_calendar(self):
        """Hiển thị calendar popup với chọn năm"""
        # Tạo popup window với style đẹp
        popup = tk.Toplevel(self.parent)
        popup.title("📅 Chọn Ngày")
        popup.geometry("340x380")
        popup.transient(self.parent)
        popup.grab_set()
        popup.configure(bg='#FAFAFA')
        
        # Lưu reference để có thể đóng đúng cách
        self.current_popup = popup
        
        # Center window
        popup.update_idletasks()
        x = (popup.winfo_screenwidth() // 2) - (340 // 2)
        y = (popup.winfo_screenheight() // 2) - (380 // 2)
        popup.geometry(f"340x380+{x}+{y}")
        
        # Frame chứa calendar
        cal_frame = tk.Frame(popup, padx=15, pady=15, bg='#FAFAFA')
        cal_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header với tháng/năm - style đẹp
        header_frame = tk.Frame(cal_frame, bg='#388E3C', relief=tk.FLAT, bd=0)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.current_date = self.selected_date
        
        self.month_var = tk.StringVar(value=month_name[self.current_date.month])
        self.year_var = tk.IntVar(value=self.current_date.year)
        
        # Nút prev tháng
        prev_month_btn = tk.Button(
            header_frame,
            text="◀",
            command=lambda: self.change_month(-1, cal_frame),
            width=3,
            bg='#2E7D32',
            fg='white',
            font=('Segoe UI', 11, 'bold'),
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        prev_month_btn.pack(side=tk.LEFT, padx=5, pady=8)
        
        # Label tháng - có thể click để chọn tháng (tùy chọn)
        month_label = tk.Label(
            header_frame,
            textvariable=self.month_var,
            font=('Segoe UI', 12, 'bold'),
            width=12,
            bg='#388E3C',
            fg='white'
        )
        month_label.pack(side=tk.LEFT, padx=5)
        
        # Nút prev năm
        prev_year_btn = tk.Button(
            header_frame,
            text="◁",
            command=lambda: self.change_year(-1, cal_frame),
            width=3,
            bg='#2E7D32',
            fg='white',
            font=('Segoe UI', 11, 'bold'),
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        prev_year_btn.pack(side=tk.LEFT, padx=2)
        
        # Label năm - click để chọn năm
        year_label = tk.Label(
            header_frame,
            textvariable=self.year_var,
            font=('Segoe UI', 12, 'bold'),
            width=6,
            bg='#388E3C',
            fg='white',
            cursor='hand2'
        )
        year_label.pack(side=tk.LEFT, padx=2)
        year_label.bind('<Button-1>', lambda e: self.show_year_picker(cal_frame))
        
        # Nút next năm
        next_year_btn = tk.Button(
            header_frame,
            text="▷",
            command=lambda: self.change_year(1, cal_frame),
            width=3,
            bg='#2E7D32',
            fg='white',
            font=('Segoe UI', 11, 'bold'),
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        next_year_btn.pack(side=tk.LEFT, padx=2)
        
        # Nút next tháng
        next_month_btn = tk.Button(
            header_frame,
            text="▶",
            command=lambda: self.change_month(1, cal_frame),
            width=3,
            bg='#2E7D32',
            fg='white',
            font=('Segoe UI', 11, 'bold'),
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        next_month_btn.pack(side=tk.LEFT, padx=5, pady=8)
        
        # Calendar grid với background
        self.cal_grid = tk.Frame(cal_frame, bg='white', relief=tk.FLAT, bd=1, highlightbackground='#E0E0E0', highlightthickness=1)
        self.cal_grid.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.update_calendar(cal_frame)
        
        # Buttons với style đẹp
        btn_frame = tk.Frame(cal_frame, bg='#FAFAFA')
        btn_frame.pack(fill=tk.X, pady=5)
        
        def close_popup_safely():
            """Đóng popup an toàn và update parent"""
            try:
                popup.grab_release()
            except:
                pass
            try:
                popup.destroy()
            except:
                pass
            # Update parent để đảm bảo UI được render lại
            try:
                self.parent.update_idletasks()
                if hasattr(self.parent, 'master'):
                    self.parent.master.update_idletasks()
            except:
                pass
            # Clear reference
            if hasattr(self, 'current_popup'):
                self.current_popup = None
        
        today_btn = tk.Button(
            btn_frame,
            text="📅 Hôm Nay",
            command=lambda: self.select_today(popup, close_popup_safely),
            bg='#388E3C',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        today_btn.pack(side=tk.LEFT, padx=5)
        
        def close_popup_safely():
            """Đóng popup an toàn và update parent"""
            try:
                popup.grab_release()
            except:
                pass
            try:
                popup.destroy()
            except:
                pass
            # Update parent để đảm bảo UI được render lại
            try:
                self.parent.update_idletasks()
                if hasattr(self.parent, 'master'):
                    self.parent.master.update_idletasks()
            except:
                pass
            # Clear reference
            if hasattr(self, 'current_popup'):
                self.current_popup = None
        
        cancel_btn = tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=close_popup_safely,
            bg='#E53935',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
    
    def change_month(self, delta, cal_frame):
        """Thay đổi tháng"""
        if delta > 0:
            if self.current_date.month == 12:
                self.current_date = self.current_date.replace(year=self.current_date.year + 1, month=1)
            else:
                self.current_date = self.current_date.replace(month=self.current_date.month + 1)
        else:
            if self.current_date.month == 1:
                self.current_date = self.current_date.replace(year=self.current_date.year - 1, month=12)
            else:
                self.current_date = self.current_date.replace(month=self.current_date.month - 1)
        
        self.month_var.set(month_name[self.current_date.month])
        self.year_var.set(self.current_date.year)
        self.update_calendar(cal_frame)
    
    def change_year(self, delta, cal_frame):
        """Thay đổi năm"""
        self.current_date = self.current_date.replace(year=self.current_date.year + delta)
        self.year_var.set(self.current_date.year)
        self.update_calendar(cal_frame)
    
    def show_year_picker(self, cal_frame):
        """Hiển thị dialog chọn năm"""
        year_popup = tk.Toplevel(cal_frame)
        year_popup.title("Chọn Năm")
        year_popup.geometry("300x400")
        year_popup.transient(cal_frame)
        year_popup.grab_set()
        year_popup.configure(bg='#FAFAFA')
        
        # Center window
        year_popup.update_idletasks()
        x = (year_popup.winfo_screenwidth() // 2) - (300 // 2)
        y = (year_popup.winfo_screenheight() // 2) - (400 // 2)
        year_popup.geometry(f"300x400+{x}+{y}")
        
        # Title
        title_label = tk.Label(
            year_popup,
            text="Chọn Năm",
            font=('Segoe UI', 14, 'bold'),
            bg='#388E3C',
            fg='white',
            pady=15
        )
        title_label.pack(fill=tk.X)
        
        # Frame chứa danh sách năm
        year_frame = tk.Frame(year_popup, bg='#FAFAFA')
        year_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollable frame cho năm
        canvas = tk.Canvas(year_frame, bg='white', highlightthickness=0)
        scrollbar = ttk.Scrollbar(year_frame, orient="vertical", command=canvas.yview)
        scrollable_years = tk.Frame(canvas, bg='white')
        
        scrollable_years.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_years, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Tạo danh sách năm (từ năm hiện tại - 100 đến + 10)
        current_year = datetime.now().year
        years = list(range(current_year - 100, current_year + 11))
        years.reverse()  # Năm mới nhất ở trên
        
        for year in years:
            year_btn = tk.Button(
                scrollable_years,
                text=str(year),
                font=('Segoe UI', 11),
                bg='white' if year != self.current_date.year else '#388E3C',
                fg='#424242' if year != self.current_date.year else 'white',
                activebackground='#E8F5E9',
                activeforeground='#424242',
                relief=tk.FLAT,
                bd=0,
                cursor='hand2',
                pady=8,
                command=lambda y=year: self.select_year(y, year_popup, cal_frame)
            )
            year_btn.pack(fill=tk.X, padx=5, pady=2)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
    
    def select_year(self, year, year_popup, cal_frame):
        """Chọn năm"""
        self.current_date = self.current_date.replace(year=year)
        self.year_var.set(year)
        self.update_calendar(cal_frame)
        # Đóng year popup đúng cách
        try:
            year_popup.grab_release()
        except:
            pass
        try:
            year_popup.destroy()
        except:
            pass
        # Update để đảm bảo UI được render lại
        try:
            cal_frame.update_idletasks()
        except:
            pass
    
    def update_calendar(self, cal_frame):
        """Cập nhật calendar grid với style đẹp"""
        # Xóa grid cũ
        for widget in self.cal_grid.winfo_children():
            widget.destroy()
        
        # Header ngày trong tuần với style
        days = ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN']
        for i, day in enumerate(days):
            label = tk.Label(
                self.cal_grid,
                text=day,
                font=('Segoe UI', 10, 'bold'),
                width=5,
                bg='#388E3C',
                fg='white',
                relief=tk.FLAT,
                bd=0
            )
            label.grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
        
        # Lấy calendar của tháng
        cal = monthcalendar(self.current_date.year, self.current_date.month)
        
        # Tạo buttons cho mỗi ngày với style đẹp
        for week_num, week in enumerate(cal, 1):
            for day_num, day in enumerate(week):
                if day == 0:
                    # Empty cell
                    empty = tk.Label(self.cal_grid, text='', width=5, bg='white')
                    empty.grid(row=week_num, column=day_num, padx=1, pady=1)
                    continue
                
                btn = tk.Button(
                    self.cal_grid,
                    text=str(day),
                    width=5,
                    height=2,
                    command=lambda d=day: self.select_date(d, cal_frame.master),
                    font=('Segoe UI', 10),
                    bg='white',
                    fg='#424242',
                    activebackground='#E8F5E9',
                    activeforeground='#424242',
                    relief=tk.FLAT,
                    bd=0,
                    cursor='hand2'
                )
                
                # Highlight ngày hiện tại
                if (day == datetime.now().day and 
                    self.current_date.month == datetime.now().month and
                    self.current_date.year == datetime.now().year):
                    btn.config(bg='#4CAF50', fg='white', font=('Segoe UI', 10, 'bold'))
                
                # Highlight ngày đã chọn
                if (day == self.selected_date.day and
                    self.current_date.month == self.selected_date.month and
                    self.current_date.year == self.selected_date.year):
                    btn.config(bg='#2196F3', fg='white')
                
                btn.grid(row=week_num, column=day_num, padx=1, pady=1, sticky='nsew')
        
        # Configure grid weights
        for i in range(7):
            self.cal_grid.grid_columnconfigure(i, weight=1)
    
    def select_date(self, day, popup):
        """Chọn ngày"""
        selected = datetime(self.current_date.year, self.current_date.month, day)
        self.selected_date = selected
        self.entry_var.set(selected.strftime("%d/%m/%Y"))
        # Đóng popup đúng cách
        try:
            popup.grab_release()
        except:
            pass
        try:
            popup.destroy()
        except:
            pass
        # Update parent để đảm bảo UI được render lại
        try:
            self.parent.update_idletasks()
            if hasattr(self.parent, 'master'):
                self.parent.master.update_idletasks()
        except:
            pass
        # Clear reference
        if hasattr(self, 'current_popup'):
            self.current_popup = None
    
    def select_today(self, popup, close_callback=None):
        """Chọn hôm nay"""
        self.selected_date = datetime.now()
        self.current_date = self.selected_date
        self.month_var.set(month_name[self.selected_date.month])
        self.year_var.set(self.selected_date.year)
        self.entry_var.set(self.selected_date.strftime("%d/%m/%Y"))
        # Đóng popup đúng cách
        if close_callback:
            close_callback()
        else:
            try:
                popup.grab_release()
            except:
                pass
            try:
                popup.destroy()
            except:
                pass
            # Update parent để đảm bảo UI được render lại
            try:
                self.parent.update_idletasks()
                if hasattr(self.parent, 'master'):
                    self.parent.master.update_idletasks()
            except:
                pass
            # Clear reference
            if hasattr(self, 'current_popup'):
                self.current_popup = None
    
    def get_date(self):
        """Lấy ngày đã chọn (DD/MM/YYYY)"""
        return self.entry_var.get()
    
    def pack(self, **kwargs):
        """Pack frame"""
        self.frame.pack(**kwargs)
    
    def grid(self, **kwargs):
        """Grid frame"""
        self.frame.grid(**kwargs)

import tkinter as tk
from tkinter import ttk
from datetime import datetime
from calendar import monthcalendar, month_name


class DatePicker:
    """Date picker widget với chọn năm"""
    
    def __init__(self, parent, default_date=None):
        """
        Args:
            parent: Parent widget
            default_date: Date string in DD/MM/YYYY format
        """
        self.parent = parent
        self.selected_date = None
        self.result = None
        
        # Parse default date
        if default_date:
            try:
                self.selected_date = datetime.strptime(default_date, "%d/%m/%Y")
            except:
                self.selected_date = datetime.now()
        else:
            self.selected_date = datetime.now()
        
        self.create_widget()
    
    def create_widget(self):
        """Tạo date picker widget - mềm mại, dễ nhìn"""
        # Frame chính - full width
        self.frame = tk.Frame(self.parent)
        self.frame.pack(fill=tk.X, expand=True)
        
        # Entry để hiển thị - style mềm mại
        self.entry_var = tk.StringVar()
        self.entry_var.set(self.selected_date.strftime("%d/%m/%Y"))
        
        self.entry = tk.Entry(
            self.frame,
            textvariable=self.entry_var,
            font=('Segoe UI', 10),
            justify=tk.CENTER,
            relief=tk.FLAT,
            bd=1,
            bg='white',
            fg='#424242',
            insertbackground='#388E3C',
            highlightthickness=1,
            highlightcolor='#388E3C',
            highlightbackground='#E0E0E0'
        )
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8), ipady=8)
        
        # Nút calendar - mềm mại hơn
        self.cal_btn = tk.Button(
            self.frame,
            text="📅",
            command=self.show_calendar,
            font=('Segoe UI', 11),
            width=5,
            height=1,
            cursor='hand2',
            bg='#388E3C',
            fg='white',
            activebackground='#2E7D32',
            activeforeground='white',
            relief=tk.FLAT,
            bd=0
        )
        self.cal_btn.pack(side=tk.LEFT)
        
        # Bind để validate format
        self.entry_var.trace('w', self.validate_date)
    
    def validate_date(self, *args):
        """Validate date format"""
        value = self.entry_var.get()
        try:
            if value:
                datetime.strptime(value, "%d/%m/%Y")
        except:
            pass  # Invalid format, user đang nhập
    
    def show_calendar(self):
        """Hiển thị calendar popup với chọn năm"""
        # Tạo popup window với style đẹp
        popup = tk.Toplevel(self.parent)
        popup.title("📅 Chọn Ngày")
        popup.geometry("340x380")
        popup.transient(self.parent)
        popup.grab_set()
        popup.configure(bg='#FAFAFA')
        
        # Lưu reference để có thể đóng đúng cách
        self.current_popup = popup
        
        # Center window
        popup.update_idletasks()
        x = (popup.winfo_screenwidth() // 2) - (340 // 2)
        y = (popup.winfo_screenheight() // 2) - (380 // 2)
        popup.geometry(f"340x380+{x}+{y}")
        
        # Frame chứa calendar
        cal_frame = tk.Frame(popup, padx=15, pady=15, bg='#FAFAFA')
        cal_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header với tháng/năm - style đẹp
        header_frame = tk.Frame(cal_frame, bg='#388E3C', relief=tk.FLAT, bd=0)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.current_date = self.selected_date
        
        self.month_var = tk.StringVar(value=month_name[self.current_date.month])
        self.year_var = tk.IntVar(value=self.current_date.year)
        
        # Nút prev tháng
        prev_month_btn = tk.Button(
            header_frame,
            text="◀",
            command=lambda: self.change_month(-1, cal_frame),
            width=3,
            bg='#2E7D32',
            fg='white',
            font=('Segoe UI', 11, 'bold'),
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        prev_month_btn.pack(side=tk.LEFT, padx=5, pady=8)
        
        # Label tháng - có thể click để chọn tháng (tùy chọn)
        month_label = tk.Label(
            header_frame,
            textvariable=self.month_var,
            font=('Segoe UI', 12, 'bold'),
            width=12,
            bg='#388E3C',
            fg='white'
        )
        month_label.pack(side=tk.LEFT, padx=5)
        
        # Nút prev năm
        prev_year_btn = tk.Button(
            header_frame,
            text="◁",
            command=lambda: self.change_year(-1, cal_frame),
            width=3,
            bg='#2E7D32',
            fg='white',
            font=('Segoe UI', 11, 'bold'),
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        prev_year_btn.pack(side=tk.LEFT, padx=2)
        
        # Label năm - click để chọn năm
        year_label = tk.Label(
            header_frame,
            textvariable=self.year_var,
            font=('Segoe UI', 12, 'bold'),
            width=6,
            bg='#388E3C',
            fg='white',
            cursor='hand2'
        )
        year_label.pack(side=tk.LEFT, padx=2)
        year_label.bind('<Button-1>', lambda e: self.show_year_picker(cal_frame))
        
        # Nút next năm
        next_year_btn = tk.Button(
            header_frame,
            text="▷",
            command=lambda: self.change_year(1, cal_frame),
            width=3,
            bg='#2E7D32',
            fg='white',
            font=('Segoe UI', 11, 'bold'),
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        next_year_btn.pack(side=tk.LEFT, padx=2)
        
        # Nút next tháng
        next_month_btn = tk.Button(
            header_frame,
            text="▶",
            command=lambda: self.change_month(1, cal_frame),
            width=3,
            bg='#2E7D32',
            fg='white',
            font=('Segoe UI', 11, 'bold'),
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        next_month_btn.pack(side=tk.LEFT, padx=5, pady=8)
        
        # Calendar grid với background
        self.cal_grid = tk.Frame(cal_frame, bg='white', relief=tk.FLAT, bd=1, highlightbackground='#E0E0E0', highlightthickness=1)
        self.cal_grid.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.update_calendar(cal_frame)
        
        # Buttons với style đẹp
        btn_frame = tk.Frame(cal_frame, bg='#FAFAFA')
        btn_frame.pack(fill=tk.X, pady=5)
        
        def close_popup_safely():
            """Đóng popup an toàn và update parent"""
            try:
                popup.grab_release()
            except:
                pass
            try:
                popup.destroy()
            except:
                pass
            # Update parent để đảm bảo UI được render lại
            try:
                self.parent.update_idletasks()
                if hasattr(self.parent, 'master'):
                    self.parent.master.update_idletasks()
            except:
                pass
            # Clear reference
            if hasattr(self, 'current_popup'):
                self.current_popup = None
        
        today_btn = tk.Button(
            btn_frame,
            text="📅 Hôm Nay",
            command=lambda: self.select_today(popup, close_popup_safely),
            bg='#388E3C',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        today_btn.pack(side=tk.LEFT, padx=5)
        
        def close_popup_safely():
            """Đóng popup an toàn và update parent"""
            try:
                popup.grab_release()
            except:
                pass
            try:
                popup.destroy()
            except:
                pass
            # Update parent để đảm bảo UI được render lại
            try:
                self.parent.update_idletasks()
                if hasattr(self.parent, 'master'):
                    self.parent.master.update_idletasks()
            except:
                pass
            # Clear reference
            if hasattr(self, 'current_popup'):
                self.current_popup = None
        
        cancel_btn = tk.Button(
            btn_frame,
            text="❌ Hủy",
            command=close_popup_safely,
            bg='#E53935',
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            padx=15,
            pady=8,
            cursor='hand2',
            relief=tk.FLAT,
            bd=0
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
    
    def change_month(self, delta, cal_frame):
        """Thay đổi tháng"""
        if delta > 0:
            if self.current_date.month == 12:
                self.current_date = self.current_date.replace(year=self.current_date.year + 1, month=1)
            else:
                self.current_date = self.current_date.replace(month=self.current_date.month + 1)
        else:
            if self.current_date.month == 1:
                self.current_date = self.current_date.replace(year=self.current_date.year - 1, month=12)
            else:
                self.current_date = self.current_date.replace(month=self.current_date.month - 1)
        
        self.month_var.set(month_name[self.current_date.month])
        self.year_var.set(self.current_date.year)
        self.update_calendar(cal_frame)
    
    def change_year(self, delta, cal_frame):
        """Thay đổi năm"""
        self.current_date = self.current_date.replace(year=self.current_date.year + delta)
        self.year_var.set(self.current_date.year)
        self.update_calendar(cal_frame)
    
    def show_year_picker(self, cal_frame):
        """Hiển thị dialog chọn năm"""
        year_popup = tk.Toplevel(cal_frame)
        year_popup.title("Chọn Năm")
        year_popup.geometry("300x400")
        year_popup.transient(cal_frame)
        year_popup.grab_set()
        year_popup.configure(bg='#FAFAFA')
        
        # Center window
        year_popup.update_idletasks()
        x = (year_popup.winfo_screenwidth() // 2) - (300 // 2)
        y = (year_popup.winfo_screenheight() // 2) - (400 // 2)
        year_popup.geometry(f"300x400+{x}+{y}")
        
        # Title
        title_label = tk.Label(
            year_popup,
            text="Chọn Năm",
            font=('Segoe UI', 14, 'bold'),
            bg='#388E3C',
            fg='white',
            pady=15
        )
        title_label.pack(fill=tk.X)
        
        # Frame chứa danh sách năm
        year_frame = tk.Frame(year_popup, bg='#FAFAFA')
        year_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Scrollable frame cho năm
        canvas = tk.Canvas(year_frame, bg='white', highlightthickness=0)
        scrollbar = ttk.Scrollbar(year_frame, orient="vertical", command=canvas.yview)
        scrollable_years = tk.Frame(canvas, bg='white')
        
        scrollable_years.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_years, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Tạo danh sách năm (từ năm hiện tại - 100 đến + 10)
        current_year = datetime.now().year
        years = list(range(current_year - 100, current_year + 11))
        years.reverse()  # Năm mới nhất ở trên
        
        for year in years:
            year_btn = tk.Button(
                scrollable_years,
                text=str(year),
                font=('Segoe UI', 11),
                bg='white' if year != self.current_date.year else '#388E3C',
                fg='#424242' if year != self.current_date.year else 'white',
                activebackground='#E8F5E9',
                activeforeground='#424242',
                relief=tk.FLAT,
                bd=0,
                cursor='hand2',
                pady=8,
                command=lambda y=year: self.select_year(y, year_popup, cal_frame)
            )
            year_btn.pack(fill=tk.X, padx=5, pady=2)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
    
    def select_year(self, year, year_popup, cal_frame):
        """Chọn năm"""
        self.current_date = self.current_date.replace(year=year)
        self.year_var.set(year)
        self.update_calendar(cal_frame)
        # Đóng year popup đúng cách
        try:
            year_popup.grab_release()
        except:
            pass
        try:
            year_popup.destroy()
        except:
            pass
        # Update để đảm bảo UI được render lại
        try:
            cal_frame.update_idletasks()
        except:
            pass
    
    def update_calendar(self, cal_frame):
        """Cập nhật calendar grid với style đẹp"""
        # Xóa grid cũ
        for widget in self.cal_grid.winfo_children():
            widget.destroy()
        
        # Header ngày trong tuần với style
        days = ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN']
        for i, day in enumerate(days):
            label = tk.Label(
                self.cal_grid,
                text=day,
                font=('Segoe UI', 10, 'bold'),
                width=5,
                bg='#388E3C',
                fg='white',
                relief=tk.FLAT,
                bd=0
            )
            label.grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
        
        # Lấy calendar của tháng
        cal = monthcalendar(self.current_date.year, self.current_date.month)
        
        # Tạo buttons cho mỗi ngày với style đẹp
        for week_num, week in enumerate(cal, 1):
            for day_num, day in enumerate(week):
                if day == 0:
                    # Empty cell
                    empty = tk.Label(self.cal_grid, text='', width=5, bg='white')
                    empty.grid(row=week_num, column=day_num, padx=1, pady=1)
                    continue
                
                btn = tk.Button(
                    self.cal_grid,
                    text=str(day),
                    width=5,
                    height=2,
                    command=lambda d=day: self.select_date(d, cal_frame.master),
                    font=('Segoe UI', 10),
                    bg='white',
                    fg='#424242',
                    activebackground='#E8F5E9',
                    activeforeground='#424242',
                    relief=tk.FLAT,
                    bd=0,
                    cursor='hand2'
                )
                
                # Highlight ngày hiện tại
                if (day == datetime.now().day and 
                    self.current_date.month == datetime.now().month and
                    self.current_date.year == datetime.now().year):
                    btn.config(bg='#4CAF50', fg='white', font=('Segoe UI', 10, 'bold'))
                
                # Highlight ngày đã chọn
                if (day == self.selected_date.day and
                    self.current_date.month == self.selected_date.month and
                    self.current_date.year == self.selected_date.year):
                    btn.config(bg='#2196F3', fg='white')
                
                btn.grid(row=week_num, column=day_num, padx=1, pady=1, sticky='nsew')
        
        # Configure grid weights
        for i in range(7):
            self.cal_grid.grid_columnconfigure(i, weight=1)
    
    def select_date(self, day, popup):
        """Chọn ngày"""
        selected = datetime(self.current_date.year, self.current_date.month, day)
        self.selected_date = selected
        self.entry_var.set(selected.strftime("%d/%m/%Y"))
        # Đóng popup đúng cách
        try:
            popup.grab_release()
        except:
            pass
        try:
            popup.destroy()
        except:
            pass
        # Update parent để đảm bảo UI được render lại
        try:
            self.parent.update_idletasks()
            if hasattr(self.parent, 'master'):
                self.parent.master.update_idletasks()
        except:
            pass
        # Clear reference
        if hasattr(self, 'current_popup'):
            self.current_popup = None
    
    def select_today(self, popup, close_callback=None):
        """Chọn hôm nay"""
        self.selected_date = datetime.now()
        self.current_date = self.selected_date
        self.month_var.set(month_name[self.selected_date.month])
        self.year_var.set(self.selected_date.year)
        self.entry_var.set(self.selected_date.strftime("%d/%m/%Y"))
        # Đóng popup đúng cách
        if close_callback:
            close_callback()
        else:
            try:
                popup.grab_release()
            except:
                pass
            try:
                popup.destroy()
            except:
                pass
            # Update parent để đảm bảo UI được render lại
            try:
                self.parent.update_idletasks()
                if hasattr(self.parent, 'master'):
                    self.parent.master.update_idletasks()
            except:
                pass
            # Clear reference
            if hasattr(self, 'current_popup'):
                self.current_popup = None
    
    def get_date(self):
        """Lấy ngày đã chọn (DD/MM/YYYY)"""
        return self.entry_var.get()
    
    def pack(self, **kwargs):
        """Pack frame"""
        self.frame.pack(**kwargs)
    
    def grid(self, **kwargs):
        """Grid frame"""
        self.frame.grid(**kwargs)