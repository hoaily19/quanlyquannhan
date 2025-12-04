"""
Models cho ứng dụng
"""

# Import trực tiếp từ module con
from .personnel import Personnel, ThongTinDang, ThongTinDoan, ThongTinKhac
from .nguoi_than import NguoiThan

__all__ = ['Personnel', 'ThongTinDang', 'ThongTinDoan', 'ThongTinKhac', 'NguoiThan']
