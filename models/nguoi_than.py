"""
Model dữ liệu người thân
"""

from dataclasses import dataclass
from typing import Optional, Dict, Any
from datetime import datetime


@dataclass
class NguoiThan:
    """Model người thân của quân nhân"""
    id: Optional[str] = None
    personnelId: Optional[str] = None  # ID quân nhân
    hoTen: str = ""  # Họ và tên
    ngaySinh: str = ""  # Ngày sinh
    diaChi: str = ""  # Địa chỉ
    soDienThoai: str = ""  # Số điện thoại
    moiQuanHe: str = ""  # Mối quan hệ (Bố, Mẹ, Anh, Em...)
    noiDung: str = ""  # Nội dung
    ghiChu: str = ""  # Ghi chú
    createdAt: Optional[datetime] = None
    updatedAt: Optional[datetime] = None

    def to_dict(self) -> Dict[str, Any]:
        """Chuyển đổi sang dictionary"""
        return {
            'id': self.id,
            'personnelId': self.personnelId,
            'hoTen': self.hoTen,
            'ngaySinh': self.ngaySinh,
            'diaChi': self.diaChi,
            'soDienThoai': self.soDienThoai,
            'moiQuanHe': self.moiQuanHe,
            'noiDung': self.noiDung,
            'ghiChu': self.ghiChu,
            'createdAt': self.createdAt.isoformat() if self.createdAt else None,
            'updatedAt': self.updatedAt.isoformat() if self.updatedAt else None,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'NguoiThan':
        """Tạo từ dictionary"""
        # Parse datetime nếu có
        createdAt = None
        if data.get('createdAt'):
            try:
                if isinstance(data['createdAt'], str):
                    createdAt = datetime.fromisoformat(data['createdAt'])
                elif isinstance(data['createdAt'], datetime):
                    createdAt = data['createdAt']
            except:
                pass
        
        updatedAt = None
        if data.get('updatedAt'):
            try:
                if isinstance(data['updatedAt'], str):
                    updatedAt = datetime.fromisoformat(data['updatedAt'])
                elif isinstance(data['updatedAt'], datetime):
                    updatedAt = data['updatedAt']
            except:
                pass
        
        return cls(
            id=data.get('id'),
            personnelId=data.get('personnelId'),
            hoTen=data.get('hoTen', ''),
            ngaySinh=data.get('ngaySinh', ''),
            diaChi=data.get('diaChi', ''),
            soDienThoai=data.get('soDienThoai', ''),
            moiQuanHe=data.get('moiQuanHe', ''),
            noiDung=data.get('noiDung', ''),
            ghiChu=data.get('ghiChu', ''),
            createdAt=createdAt,
            updatedAt=updatedAt,
        )

