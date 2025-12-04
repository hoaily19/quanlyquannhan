"""
Service xác thực và bảo mật
"""

import json
import sys
from pathlib import Path
from typing import Optional
from datetime import datetime, timedelta

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

# Bảo mật - dùng bcrypt thay vì sha256
try:
    import bcrypt
    USE_BCRYPT = True
except ImportError:
    import hashlib
    USE_BCRYPT = False
    print("⚠️  bcrypt chưa được cài, dùng SHA256 (kém bảo mật hơn)")


class AuthService:
    """Service quản lý authentication và bảo mật"""
    
    def __init__(self, users_file: str = "data/users.json"):
        """
        Khởi tạo AuthService
        Args:
            users_file: Đường dẫn file lưu thông tin users
        """
        self.users_file = Path(users_file)
        self.users_file.parent.mkdir(parents=True, exist_ok=True)
        self.current_user = None
        self.load_users()
    
    def load_users(self):
        """Load danh sách users từ file"""
        if self.users_file.exists():
            try:
                with open(self.users_file, 'r', encoding='utf-8') as f:
                    self.users = json.load(f)
            except:
                self.users = {}
        else:
            # Tạo user mặc định
            self.users = {
                'admin': {
                    'password_hash': self._hash_password('admin123'),
                    'role': 'admin',
                    'created_at': datetime.now().isoformat()
                }
            }
            self.save_users()
    
    def save_users(self):
        """Lưu danh sách users vào file"""
        with open(self.users_file, 'w', encoding='utf-8') as f:
            json.dump(self.users, f, ensure_ascii=False, indent=2)
    
    def _hash_password(self, password: str) -> str:
        """
        Hash password với bcrypt (an toàn hơn)
        Args:
            password: Mật khẩu gốc
        Returns:
            Hash của password
        """
        if USE_BCRYPT:
            # Dùng bcrypt với salt tự động
            salt = bcrypt.gensalt()
            return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')
        else:
            # Fallback về SHA256 (kém bảo mật)
            return hashlib.sha256(password.encode()).hexdigest()
    
    def _verify_password(self, password: str, password_hash: str) -> bool:
        """
        Verify password
        Args:
            password: Mật khẩu cần kiểm tra
            password_hash: Hash đã lưu
        Returns:
            True nếu password đúng
        """
        if USE_BCRYPT:
            try:
                return bcrypt.checkpw(password.encode('utf-8'), password_hash.encode('utf-8'))
            except:
                return False
        else:
            # Fallback về SHA256
            return hashlib.sha256(password.encode()).hexdigest() == password_hash
    
    def register(self, username: str, password: str, role: str = 'user') -> bool:
        """
        Đăng ký user mới
        Args:
            username: Tên đăng nhập
            password: Mật khẩu
            role: Vai trò (admin/user)
        Returns:
            True nếu thành công
        """
        if username in self.users:
            return False
        
        self.users[username] = {
            'password_hash': self._hash_password(password),
            'role': role,
            'created_at': datetime.now().isoformat()
        }
        self.save_users()
        return True
    
    def login(self, username: str, password: str) -> bool:
        """
        Đăng nhập
        Args:
            username: Tên đăng nhập
            password: Mật khẩu
        Returns:
            True nếu đăng nhập thành công
        """
        if username not in self.users:
            return False
        
        stored_hash = self.users[username]['password_hash']
        if self._verify_password(password, stored_hash):
            self.current_user = {
                'username': username,
                'role': self.users[username]['role']
            }
            return True
        
        return False
    
    def logout(self):
        """Đăng xuất"""
        self.current_user = None
    
    def is_authenticated(self) -> bool:
        """Kiểm tra xem đã đăng nhập chưa"""
        return self.current_user is not None
    
    def is_admin(self) -> bool:
        """Kiểm tra xem user hiện tại có phải admin không"""
        if not self.current_user:
            return False
        return self.current_user.get('role') == 'admin'
    
    def get_current_user(self) -> Optional[dict]:
        """Lấy thông tin user hiện tại"""
        return self.current_user


        if username not in self.users:
            return False
        
        stored_hash = self.users[username]['password_hash']
        if self._verify_password(password, stored_hash):
            self.current_user = {
                'username': username,
                'role': self.users[username]['role']
            }
            return True
        
        return False
    
    def logout(self):
        """Đăng xuất"""
        self.current_user = None
    
    def is_authenticated(self) -> bool:
        """Kiểm tra xem đã đăng nhập chưa"""
        return self.current_user is not None
    
    def is_admin(self) -> bool:
        """Kiểm tra xem user hiện tại có phải admin không"""
        if not self.current_user:
            return False
        return self.current_user.get('role') == 'admin'
    
    def get_current_user(self) -> Optional[dict]:
        """Lấy thông tin user hiện tại"""
        return self.current_user
