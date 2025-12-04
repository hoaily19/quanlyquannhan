# Hướng Dẫn Build Ứng Dụng Thành EXE

## Yêu Cầu

- Python 3.8 trở lên
- Windows (để build file .exe)
- Đã cài đặt các dependencies từ `requirements.txt`

## Cách Build

### Cách 1: Sử dụng PowerShell Script (Khuyên dùng)

```powershell
.\build.ps1
```

Script này sẽ tự động:
- Kiểm tra Python và pip
- Cài đặt PyInstaller nếu chưa có
- Cài đặt các dependencies
- Build ứng dụng thành EXE

### Cách 2: Sử dụng Python Script

```bash
python build_exe.py
```

### Cách 3: Sử dụng PyInstaller trực tiếp

```bash
pyinstaller build.spec --clean --noconfirm
```

## Kết Quả

Sau khi build thành công, file EXE sẽ được tạo tại:
```
dist/QuanLyQuanNhan.exe
```

## Icon

Ứng dụng sẽ sử dụng icon từ `icons/logo.ico`. 
Nếu file này không tồn tại, app sẽ không có icon.

## Phân Phối

File EXE có thể chạy độc lập trên bất kỳ máy Windows nào mà không cần cài đặt Python.

**Lưu ý:**
- File EXE sẽ khá lớn (khoảng 100-200MB) vì đã bao gồm Python runtime và tất cả dependencies
- Lần đầu chạy có thể hơi chậm do Windows Defender quét file
- Nếu gặp lỗi "Windows protected your PC", click "More info" > "Run anyway"

## Troubleshooting

### Lỗi: "PyInstaller not found"
```bash
pip install pyinstaller
```

### Lỗi: "Module not found" khi chạy EXE
- Kiểm tra file `build.spec` đã include đầy đủ modules trong `hiddenimports`
- Thử build lại với `--clean` flag

### EXE không có icon
- Kiểm tra file `icons/logo.ico` có tồn tại không
- Đảm bảo đường dẫn icon trong `build.spec` đúng

### EXE quá lớn
- Có thể loại bỏ một số dependencies không cần thiết
- Sử dụng UPX compression (đã bật trong spec)

