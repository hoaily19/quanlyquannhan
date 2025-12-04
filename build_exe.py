"""
Script build ứng dụng thành file EXE
Sử dụng PyInstaller
"""

import subprocess
import sys
import os
from pathlib import Path

def main():
    """Build ứng dụng thành EXE"""
    print("=" * 60)
    print("🔨 Bắt đầu build ứng dụng Quản Lý Quân Nhân")
    print("=" * 60)
    
    # Kiểm tra PyInstaller đã được cài đặt chưa
    try:
        import PyInstaller
        print("✅ PyInstaller đã được cài đặt")
    except ImportError:
        print("❌ PyInstaller chưa được cài đặt!")
        print("Đang cài đặt PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✅ Đã cài đặt PyInstaller")
    
    # Kiểm tra file spec
    spec_file = Path(__file__).parent / "build.spec"
    if not spec_file.exists():
        print("❌ Không tìm thấy file build.spec!")
        return 1
    
    # Kiểm tra icon
    icon_file = Path(__file__).parent / "icons" / "logo.ico"
    if not icon_file.exists():
        print("⚠️  Cảnh báo: Không tìm thấy icons/logo.ico")
        print("   App sẽ không có icon")
    
    # Xóa thư mục build và dist cũ (nếu có)
    build_dir = Path(__file__).parent / "build"
    dist_dir = Path(__file__).parent / "dist"
    
    if build_dir.exists():
        print(f"🗑️  Xóa thư mục build cũ...")
        import shutil
        shutil.rmtree(build_dir)
    
    # Build với PyInstaller
    print("\n🔨 Đang build...")
    print(f"   Sử dụng file: {spec_file}")
    
    try:
        cmd = [
            sys.executable, "-m", "PyInstaller",
            str(spec_file),
            "--clean",
            "--noconfirm"
        ]
        
        result = subprocess.run(cmd, check=True, cwd=Path(__file__).parent)
        
        print("\n" + "=" * 60)
        print("✅ Build thành công!")
        print("=" * 60)
        print(f"📦 File EXE được tạo tại: {dist_dir / 'QuanLyQuanNhan.exe'}")
        print("\n💡 Bạn có thể chạy file EXE này trên bất kỳ máy Windows nào")
        print("   (không cần cài Python)")
        
        return 0
        
    except subprocess.CalledProcessError as e:
        print("\n" + "=" * 60)
        print("❌ Build thất bại!")
        print("=" * 60)
        print(f"Lỗi: {e}")
        return 1
    except Exception as e:
        print("\n" + "=" * 60)
        print("❌ Có lỗi xảy ra!")
        print("=" * 60)
        print(f"Lỗi: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())

