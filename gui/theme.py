"""
Theme và màu sắc phong cách quân đội
"""

# Màu sắc quân đội
MILITARY_COLORS = {
    # Màu chính
    'primary': '#2E7D32',      # Xanh lá đậm (quân đội)
    'primary_light': '#4CAF50', # Xanh lá sáng
    'primary_dark': '#1B5E20',  # Xanh lá rất đậm
    
    # Màu phụ
    'secondary': '#1565C0',     # Xanh dương (biển)
    'secondary_light': '#2196F3',
    'secondary_dark': '#0D47A1',
    
    # Màu accent
    'accent': '#FF6F00',        # Cam (cảnh báo)
    'accent_light': '#FF9800',
    'accent_dark': '#E65100',
    
    # Màu vàng (huy chương)
    'gold': '#F57F17',
    'gold_light': '#FFC107',
    
    # Màu nền
    'bg_dark': '#1A1A1A',        # Nền tối
    'bg_medium': '#2C2C2C',      # Nền trung bình
    'bg_light': '#F5F5F5',       # Nền sáng
    'bg_cream': '#FFF8E1',        # Nền kem
    
    # Màu chữ
    'text_light': '#FFFFFF',     # Chữ trắng
    'text_dark': '#212121',      # Chữ đen
    'text_gray': '#757575',      # Chữ xám
    
    # Màu border
    'border': '#BDBDBD',
    'border_dark': '#616161',
    
    # Màu trạng thái
    'success': '#4CAF50',
    'warning': '#FF9800',
    'error': '#F44336',
    'info': '#2196F3',
}

# Style cho buttons
def get_button_style(style_type='primary'):
    """Lấy style cho button"""
    styles = {
        'primary': {
            'bg': MILITARY_COLORS['primary'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': MILITARY_COLORS['primary_dark'],
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        },
        'secondary': {
            'bg': MILITARY_COLORS['secondary'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': MILITARY_COLORS['secondary_dark'],
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        },
        'accent': {
            'bg': MILITARY_COLORS['accent'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': MILITARY_COLORS['accent_dark'],
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        },
        'success': {
            'bg': MILITARY_COLORS['success'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': '#45a049',
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        },
        'danger': {
            'bg': MILITARY_COLORS['error'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': '#d32f2f',
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        },
        'info': {
            'bg': MILITARY_COLORS['info'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': '#1976D2',
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        }
    }
    return styles.get(style_type, styles['primary'])

# Style cho labels
def get_label_style(style_type='normal'):
    """Lấy style cho label"""
    styles = {
        'title': {
            'font': ('Arial', 18, 'bold'),
            'fg': MILITARY_COLORS['primary_dark'],
            'bg': MILITARY_COLORS['bg_light'],
        },
        'subtitle': {
            'font': ('Arial', 12, 'bold'),
            'fg': MILITARY_COLORS['text_dark'],
            'bg': MILITARY_COLORS['bg_light'],
        },
        'normal': {
            'font': ('Arial', 10),
            'fg': MILITARY_COLORS['text_dark'],
            'bg': MILITARY_COLORS['bg_light'],
        },
        'dark': {
            'font': ('Arial', 10),
            'fg': MILITARY_COLORS['text_light'],
            'bg': MILITARY_COLORS['bg_medium'],
        }
    }
    return styles.get(style_type, styles['normal'])

# Màu sắc quân đội
MILITARY_COLORS = {
    # Màu chính
    'primary': '#2E7D32',      # Xanh lá đậm (quân đội)
    'primary_light': '#4CAF50', # Xanh lá sáng
    'primary_dark': '#1B5E20',  # Xanh lá rất đậm
    
    # Màu phụ
    'secondary': '#1565C0',     # Xanh dương (biển)
    'secondary_light': '#2196F3',
    'secondary_dark': '#0D47A1',
    
    # Màu accent
    'accent': '#FF6F00',        # Cam (cảnh báo)
    'accent_light': '#FF9800',
    'accent_dark': '#E65100',
    
    # Màu vàng (huy chương)
    'gold': '#F57F17',
    'gold_light': '#FFC107',
    
    # Màu nền
    'bg_dark': '#1A1A1A',        # Nền tối
    'bg_medium': '#2C2C2C',      # Nền trung bình
    'bg_light': '#F5F5F5',       # Nền sáng
    'bg_cream': '#FFF8E1',        # Nền kem
    
    # Màu chữ
    'text_light': '#FFFFFF',     # Chữ trắng
    'text_dark': '#212121',      # Chữ đen
    'text_gray': '#757575',      # Chữ xám
    
    # Màu border
    'border': '#BDBDBD',
    'border_dark': '#616161',
    
    # Màu trạng thái
    'success': '#4CAF50',
    'warning': '#FF9800',
    'error': '#F44336',
    'info': '#2196F3',
}

# Style cho buttons
def get_button_style(style_type='primary'):
    """Lấy style cho button"""
    styles = {
        'primary': {
            'bg': MILITARY_COLORS['primary'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': MILITARY_COLORS['primary_dark'],
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        },
        'secondary': {
            'bg': MILITARY_COLORS['secondary'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': MILITARY_COLORS['secondary_dark'],
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        },
        'accent': {
            'bg': MILITARY_COLORS['accent'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': MILITARY_COLORS['accent_dark'],
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        },
        'success': {
            'bg': MILITARY_COLORS['success'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': '#45a049',
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        },
        'danger': {
            'bg': MILITARY_COLORS['error'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': '#d32f2f',
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        },
        'info': {
            'bg': MILITARY_COLORS['info'],
            'fg': MILITARY_COLORS['text_light'],
            'activebackground': '#1976D2',
            'activeforeground': MILITARY_COLORS['text_light'],
            'font': ('Arial', 10, 'bold'),
            'relief': 'raised',
            'bd': 2,
            'cursor': 'hand2',
        }
    }
    return styles.get(style_type, styles['primary'])

# Style cho labels
def get_label_style(style_type='normal'):
    """Lấy style cho label"""
    styles = {
        'title': {
            'font': ('Arial', 18, 'bold'),
            'fg': MILITARY_COLORS['primary_dark'],
            'bg': MILITARY_COLORS['bg_light'],
        },
        'subtitle': {
            'font': ('Arial', 12, 'bold'),
            'fg': MILITARY_COLORS['text_dark'],
            'bg': MILITARY_COLORS['bg_light'],
        },
        'normal': {
            'font': ('Arial', 10),
            'fg': MILITARY_COLORS['text_dark'],
            'bg': MILITARY_COLORS['bg_light'],
        },
        'dark': {
            'font': ('Arial', 10),
            'fg': MILITARY_COLORS['text_light'],
            'bg': MILITARY_COLORS['bg_medium'],
        }
    }
    return styles.get(style_type, styles['normal'])