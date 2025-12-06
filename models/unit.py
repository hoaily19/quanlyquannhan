"""
Model dữ liệu đơn vị (Đại đội, Trung đội, Xe, Tổ...)
"""

from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any


@dataclass
class Unit:
    """Model đơn vị"""
    id: Optional[str] = None
    ten: str = ""  # Tên đơn vị
    loai: str = ""  # Loại: dai_doi, trung_doi, xe, to
    parentId: Optional[str] = None  # ID đơn vị cha (nếu có)
    personnelIds: List[str] = field(default_factory=list)  # Danh sách ID quân nhân
    ghiChu: Optional[str] = None  # Ghi chú
    createdAt: Optional[str] = None
    updatedAt: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """Chuyển đổi sang dictionary"""
        return {
            'id': self.id or '',
            'ten': self.ten,
            'loai': self.loai,
            'parentId': self.parentId or '',
            'personnelIds': self.personnelIds or [],
            'ghiChu': self.ghiChu or '',
            'createdAt': self.createdAt or '',
            'updatedAt': self.updatedAt or ''
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Unit':
        """Tạo Unit từ dictionary"""
        import json
        
        # Xử lý personnelIds nếu là string JSON
        personnel_ids = data.get('personnelIds', [])
        if isinstance(personnel_ids, str):
            try:
                personnel_ids = json.loads(personnel_ids) if personnel_ids else []
            except:
                personnel_ids = []
        
        return cls(
            id=data.get('id'),
            ten=data.get('ten', ''),
            loai=data.get('loai', ''),
            parentId=data.get('parentId'),
            personnelIds=personnel_ids,
            ghiChu=data.get('ghiChu'),
            createdAt=data.get('createdAt'),
            updatedAt=data.get('updatedAt')
        )


