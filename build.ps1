# PowerShell script để build ứng dụng thành EXE
# Chạy: .\build.ps1

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "🔨 Build Ứng Dụng Quản Lý Quân Nhân" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Kiểm tra Python
try {
    $pythonVersion = python --version 2>&1
    Write-Host "✅ Python: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Không tìm thấy Python! Vui lòng cài đặt Python trước." -ForegroundColor Red
    exit 1
}

# Kiểm tra pip
try {
    $pipVersion = pip --version 2>&1
    Write-Host "✅ pip: $pipVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Không tìm thấy pip!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "📦 Kiểm tra dependencies..." -ForegroundColor Yellow

# Cài đặt dependencies nếu chưa có
if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
    Write-Host "   Đang cài đặt PyInstaller..." -ForegroundColor Yellow
    pip install pyinstaller
}

# Cài đặt các dependencies khác
Write-Host "   Đang cài đặt các dependencies từ requirements.txt..." -ForegroundColor Yellow
pip install -r requirements.txt

Write-Host ""
Write-Host "🔨 Bắt đầu build..." -ForegroundColor Yellow
Write-Host ""

# Chạy build script
python build_exe.py

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "✅ Build hoàn tất!" -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "📦 File EXE: dist\QuanLyQuanNhan.exe" -ForegroundColor Cyan
    Write-Host ""
} else {
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host "❌ Build thất bại!" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Red
    exit 1
}

