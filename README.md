# Ứng Dụng Quản Lý Hồ Sơ Quân Nhân

Ứng dụng desktop Python (Tkinter) để quản lý, xem, tìm kiếm và tổng hợp dữ liệu hồ sơ quân nhân.

## Tính Năng

- ✅ Quản lý hồ sơ quân nhân (CRUD)
- ✅ Tìm kiếm và lọc hồ sơ
- ✅ Xem thống kê và báo cáo tổng hợp
- ✅ Xuất dữ liệu ra CSV/PDF
- ✅ Nhập dữ liệu từ file Word/Excel
- ✅ **Bảo mật**: Authentication, Password hashing (bcrypt), Data encryption
- ✅ **Build thành EXE**: Có thể build thành file exe để phân phối

## Cài Đặt

### 1. Cài đặt Python dependencies

```bash
pip install -r requirements.txt
```

### 2. Chạy ứng dụng

```bash
python main.py
```

Hoặc dùng script:
```powershell
.\run.ps1
```

## Build Thành EXE

Xem file `BUILD_GUIDE.md` để biết chi tiết.

**Nhanh:**
```powershell
.\build.ps1
```

File exe sẽ được tạo tại: `dist/QuanLyQuanNhan.exe`

## Cấu Trúc Dự Án

```
Quan ly quan nhan/
├── main.py                 # Entry point
├── app.py                  # Streamlit (cũ, có thể xóa)
├── requirements.txt        # Dependencies
├── build.ps1              # Script build exe
├── build_exe.py          # Python script build
├── build.spec            # PyInstaller spec
├── logo.jpg              # Logo/Icon
├── models/               # Data models
├── services/              # Business logic
│   ├── auth.py          # Authentication & Security
│   ├── encryption.py    # Data encryption
│   ├── database.py     # Database service
│   └── export.py        # Export service
├── gui/                  # GUI components
│   ├── login_window.py  # Màn hình đăng nhập
│   ├── main_window.py  # Cửa sổ chính
│   └── ...
└── utils/                # Utilities
```

## Bảo Mật

### Authentication
- Username/Password authentication
- Password hashing với bcrypt (an toàn)
- Session management

### Data Encryption
- Encryption key tự động tạo tại `data/encryption.key`
- Dữ liệu nhạy cảm có thể được mã hóa
- Database được bảo vệ

### Mặc Định
- Username: `admin`
- Password: `admin123`

**⚠️ Lưu ý:** Đổi mật khẩu ngay sau lần đăng nhập đầu tiên!

## Database

Ứng dụng sử dụng SQLite, file database được lưu tại `data/personnel.db`

## Nhập Dữ Liệu

1. Vào menu "Quản Lý" > "Nhập Dữ Liệu"
2. Chọn thư mục chứa file Word/Excel
3. Click "Đọc File"
4. Xem preview và click "Import Tất Cả"

## Xuất Dữ Liệu

- **CSV**: Menu "File" > "Xuất CSV..."
- **PDF**: Menu "File" > "Xuất PDF..."

## Yêu Cầu

- Python 3.8+
- Windows (cho exe)
- Xem `requirements.txt` để biết đầy đủ dependencies

## Lưu Ý

- File Word/Excel cần có cấu trúc nhất định (bảng hoặc cột)
- Cần điều chỉnh `utils/file_reader.py` để map đúng với cấu trúc file thực tế
- Database SQLite tự động tạo khi chạy lần đầu
- Encryption key tự động tạo tại `data/encryption.key`

## Hỗ Trợ

Nếu gặp vấn đề, xem:
- `BUILD_GUIDE.md` - Hướng dẫn build exe
- `QUICK_START.md` - Hướng dẫn nhanh
