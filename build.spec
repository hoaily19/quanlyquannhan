# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file cho ứng dụng Quản Lý Quân Nhân
Đảm bảo tất cả modules được đóng gói đầy đủ
"""

from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

# Tự động collect tất cả submodules
gui_modules = collect_submodules('gui')
services_modules = collect_submodules('services')
models_modules = collect_submodules('models')
utils_modules = collect_submodules('utils')

# In ra để debug
print(f"GUI modules found: {len(gui_modules)}")
print(f"Services modules found: {len(services_modules)}")
print(f"Models modules found: {len(models_modules)}")
print(f"Utils modules found: {len(utils_modules)}")

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('icons', 'icons'),  # Include icons folder
        ('logo.jpg', '.'),   # Include logo.jpg in root
        ('data', 'data'),    # Include data folder (database, users.json)
    ],
    hiddenimports=[
        'PIL._tkinter_finder',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'tkinter.scrolledtext',
        'pandas',
        'openpyxl',
        'xlrd',
        'docx',  # python-docx imports as 'docx'
        'matplotlib',
        'matplotlib.backends.backend_tkagg',
        'reportlab',
        'cryptography',
        'bcrypt',
        # Explicit GUI modules - đảm bảo không thiếu
        'gui',
        'gui.__init__',
        'gui.tooltip',
        'gui.main_window',
        'gui.login_window',
        'gui.splash_screen',
        'gui.personnel_list_frame',
        'gui.personnel_form_frame',
        'gui.report_frame',
        'gui.import_frame',
        'gui.reports_list_frame',
        'gui.date_picker',
        'gui.theme',
        'gui.nguoi_than_form',
        'gui.yeu_to_nuoc_ngoai_form',
        'gui.unit_management_frame',
        # Explicit Services
        'services',
        'services.__init__',
        'services.database',
        'services.auth',
        'services.encryption',
        'services.export',
        # Explicit Models
        'models',
        'models.__init__',
        'models.personnel',
        'models.nguoi_than',
        'models.unit',
        # Explicit Utils
        'utils',
        'utils.__init__',
        'utils.file_reader',
        # Tự động collect modules (backup - thêm vào cuối)
    ] + list(gui_modules) + list(services_modules) + list(models_modules) + list(utils_modules),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['streamlit', 'app', 'pages'],  # Exclude streamlit app và pages
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='QuanLyQuanNhan',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Không hiện console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icons/logo.ico',  # Sử dụng icon từ thư mục icons
)
