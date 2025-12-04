"""
Tooltip utility cho Tkinter widgets
"""

import tkinter as tk


class ToolTip:
    """Tooltip widget cho Tkinter"""
    
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0
        
    def showtip(self):
        """Hiển thị tooltip"""
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 25
        y = y + cy + self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                        background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                        font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)
        
    def hidetip(self):
        """Ẩn tooltip"""
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()


def create_tooltip(widget, text):
    """
    Tạo tooltip cho widget
    
    Args:
        widget: Tkinter widget cần tooltip
        text: Text hiển thị trong tooltip
    """
    toolTip = ToolTip(widget, text)
    
    def enter(event):
        toolTip.showtip()
        
    def leave(event):
        toolTip.hidetip()
        
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)
    
    return toolTip

