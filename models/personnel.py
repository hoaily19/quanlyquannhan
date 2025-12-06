"""
Model dữ liệu quân nhân
"""

from dataclasses import dataclass, field
from typing import Optional, Dict, Any
from datetime import datetime


@dataclass
class ThongTinDang:
    """Thông tin đảng viên"""
    ngayVao: str = ""
    ngayChinhThuc: str = ""
    chucVuDang: str = ""


@dataclass
class ThongTinDoan:
    """Thông tin đoàn viên"""
    ngayVao: str = ""
    chucVuDoan: str = ""


@dataclass
class ThongTinKhac:
    """Thông tin khác"""
    dang: ThongTinDang = field(default_factory=ThongTinDang)
    doan: ThongTinDoan = field(default_factory=ThongTinDoan)
    cdCu: bool = False  # Chế độ cũ
    yeuToNN: bool = False  # Yếu tố nước ngoài
    noiDungYeuToNN: str = ""  # Nội dung yếu tố nước ngoài
    moiQuanHeYeuToNN: str = ""  # Mối quan hệ (Bố, Mẹ, Anh, Em...)
    tenNuoc: str = ""  # Tên nước
    dangPhaiPhanDong: bool = False  # Tham gia đảng phái phản động


@dataclass
class Personnel:
    """Model quân nhân"""
    id: Optional[str] = None
    hoTen: str = ""  # Họ và tên khai sinh
    hoTenThuongDung: str = ""  # Họ và tên thường dùng
    ngaySinh: str = ""
    capBac: str = ""
    ngayNhanCapBac: str = ""  # Ngày nhận cấp bậc
    chucVu: str = ""
    ngayNhanChucVu: str = ""  # Ngày nhận chức vụ
    donVi: str = ""
    unitId: Optional[str] = None  # ID đơn vị (đại đội, trung đội, xe...)
    nhapNgu: str = ""
    xuatNgu: str = ""  # Xuất ngũ
    queQuan: str = ""
    truQuan: str = ""
    danToc: str = ""
    tonGiao: str = ""
    trinhDoVanHoa: str = ""
    thanhPhanGiaDinh: str = ""  # Thành phần gia đình (Bần nông, Công nhân...)
    # Thông tin đào tạo
    quaTruong: str = ""  # Qua trường
    nganhHoc: str = ""  # Ngành học
    capHoc: str = ""  # Cấp học
    thoiGianDaoTao: str = ""  # Thời gian đào tạo
    ketQuaDaoTao: str = ""  # Kết quả đào tạo
    # Thông tin chức vụ chiến đấu
    chucVuChienDau: str = ""  # Chức vụ chiến đấu
    thoiGianChucVuChienDau: str = ""  # Thời gian chức vụ chiến đấu
    chucVuDaQua: str = ""  # Chức vụ đã qua
    thoiGianChucVuDaQua: str = ""  # Thời gian chức vụ đã qua
    # CM Quân
    cmQuan: str = ""  # CM Quân (Tháng năm) - có thể dùng nhapNgu nhưng thêm field riêng
    # Thông tin liên hệ
    lienHeKhiCan: str = ""  # Khi cần báo tin cho ai
    soDienThoaiLienHe: str = ""  # SĐT liên hệ
    # Thông tin gia đình
    hoTenCha: str = ""
    hoTenMe: str = ""
    hoTenVo: str = ""
    ghiChu: str = ""
    ngoaiNgu: str = ""  # Ngoại ngữ
    tiengDTTS: str = ""  # Tiếng dân tộc thiểu số
    # Thông tin người thân tham gia đảng phái phản động
    hoTenNguoiThan: str = ""  # Họ tên người thân
    moiQuanHe: str = ""  # Mối quan hệ (Bố, Mẹ, Anh, Em...)
    noiDungNguoiThan: str = ""  # Nội dung người thân tham gia
    # Thông tin THAM GIA chế độ cũ
    thamGiaNguyQuan: str = ""  # Ngụy quân (có thể là "X" hoặc "")
    thamGiaNguyQuyen: str = ""  # Ngụy quyền (có thể là "X" hoặc "")
    thamGiaNoMau: str = ""  # Nợ máu/không nợ máu
    daCaiTao: str = ""  # Đã cải tạo/chưa cải tạo
    thongTinKhac: ThongTinKhac = field(default_factory=ThongTinKhac)
    createdAt: Optional[datetime] = None
    updatedAt: Optional[datetime] = None

    def to_dict(self) -> Dict[str, Any]:
        """Chuyển đổi sang dictionary"""
        return {
            'id': self.id,
            'hoTen': self.hoTen,
            'hoTenThuongDung': self.hoTenThuongDung,
            'ngaySinh': self.ngaySinh,
            'capBac': self.capBac,
            'ngayNhanCapBac': self.ngayNhanCapBac,
            'chucVu': self.chucVu,
            'ngayNhanChucVu': self.ngayNhanChucVu,
            'donVi': self.donVi,
            'unitId': self.unitId,
            'nhapNgu': self.nhapNgu,
            'xuatNgu': self.xuatNgu,
            'queQuan': self.queQuan,
            'truQuan': self.truQuan,
            'danToc': self.danToc,
            'tonGiao': self.tonGiao,
            'trinhDoVanHoa': self.trinhDoVanHoa,
            'thanhPhanGiaDinh': self.thanhPhanGiaDinh,
            'quaTruong': self.quaTruong,
            'nganhHoc': self.nganhHoc,
            'capHoc': self.capHoc,
            'thoiGianDaoTao': self.thoiGianDaoTao,
            'ketQuaDaoTao': self.ketQuaDaoTao,
            'chucVuChienDau': self.chucVuChienDau,
            'thoiGianChucVuChienDau': self.thoiGianChucVuChienDau,
            'chucVuDaQua': self.chucVuDaQua,
            'thoiGianChucVuDaQua': self.thoiGianChucVuDaQua,
            'cmQuan': self.cmQuan,
            'lienHeKhiCan': self.lienHeKhiCan,
            'soDienThoaiLienHe': self.soDienThoaiLienHe,
            'hoTenCha': self.hoTenCha,
            'hoTenMe': self.hoTenMe,
            'hoTenVo': self.hoTenVo,
            'hoTenNguoiThan': self.hoTenNguoiThan,
            'moiQuanHe': self.moiQuanHe,
            'noiDungNguoiThan': self.noiDungNguoiThan,
            'thamGiaNguyQuan': self.thamGiaNguyQuan,
            'thamGiaNguyQuyen': self.thamGiaNguyQuyen,
            'thamGiaNoMau': self.thamGiaNoMau,
            'daCaiTao': self.daCaiTao,
            'ghiChu': self.ghiChu,
            'ngoaiNgu': self.ngoaiNgu,
            'tiengDTTS': self.tiengDTTS,
            'thongTinKhac': {
                'dang': {
                    'ngayVao': self.thongTinKhac.dang.ngayVao,
                    'ngayChinhThuc': self.thongTinKhac.dang.ngayChinhThuc,
                    'chucVuDang': self.thongTinKhac.dang.chucVuDang,
                },
                'doan': {
                    'ngayVao': self.thongTinKhac.doan.ngayVao,
                    'chucVuDoan': self.thongTinKhac.doan.chucVuDoan,
                },
                'cdCu': self.thongTinKhac.cdCu,
                'yeuToNN': self.thongTinKhac.yeuToNN,
                'noiDungYeuToNN': self.thongTinKhac.noiDungYeuToNN,
                'moiQuanHeYeuToNN': self.thongTinKhac.moiQuanHeYeuToNN,
                'tenNuoc': self.thongTinKhac.tenNuoc,
                'dangPhaiPhanDong': self.thongTinKhac.dangPhaiPhanDong,
            },
            'createdAt': self.createdAt.isoformat() if self.createdAt else None,
            'updatedAt': self.updatedAt.isoformat() if self.updatedAt else None,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Personnel':
        """Tạo từ dictionary"""
        thong_tin_khac = ThongTinKhac()
        if 'thongTinKhac' in data:
            tt = data['thongTinKhac']
            thong_tin_khac.dang = ThongTinDang(
                ngayVao=tt.get('dang', {}).get('ngayVao', ''),
                ngayChinhThuc=tt.get('dang', {}).get('ngayChinhThuc', ''),
                chucVuDang=tt.get('dang', {}).get('chucVuDang', ''),
            )
            thong_tin_khac.doan = ThongTinDoan(
                ngayVao=tt.get('doan', {}).get('ngayVao', ''),
                chucVuDoan=tt.get('doan', {}).get('chucVuDoan', ''),
            )
            thong_tin_khac.cdCu = tt.get('cdCu', False)
            thong_tin_khac.yeuToNN = tt.get('yeuToNN', False)
            thong_tin_khac.noiDungYeuToNN = tt.get('noiDungYeuToNN', '')
            thong_tin_khac.moiQuanHeYeuToNN = tt.get('moiQuanHeYeuToNN', '')
            thong_tin_khac.tenNuoc = tt.get('tenNuoc', '')
            thong_tin_khac.dangPhaiPhanDong = tt.get('dangPhaiPhanDong', False)

        return cls(
            id=data.get('id'),
            hoTen=data.get('hoTen', ''),
            hoTenThuongDung=data.get('hoTenThuongDung', ''),
            ngaySinh=data.get('ngaySinh', ''),
            capBac=data.get('capBac', ''),
            ngayNhanCapBac=data.get('ngayNhanCapBac', ''),
            chucVu=data.get('chucVu', ''),
            ngayNhanChucVu=data.get('ngayNhanChucVu', ''),
            donVi=data.get('donVi', ''),
            unitId=data.get('unitId'),
            nhapNgu=data.get('nhapNgu', ''),
            xuatNgu=data.get('xuatNgu', ''),
            queQuan=data.get('queQuan', ''),
            truQuan=data.get('truQuan', ''),
            danToc=data.get('danToc', ''),
            tonGiao=data.get('tonGiao', ''),
            trinhDoVanHoa=data.get('trinhDoVanHoa', ''),
            thanhPhanGiaDinh=data.get('thanhPhanGiaDinh', ''),
            quaTruong=data.get('quaTruong', ''),
            nganhHoc=data.get('nganhHoc', ''),
            capHoc=data.get('capHoc', ''),
            thoiGianDaoTao=data.get('thoiGianDaoTao', ''),
            ketQuaDaoTao=data.get('ketQuaDaoTao', ''),
            chucVuChienDau=data.get('chucVuChienDau', ''),
            thoiGianChucVuChienDau=data.get('thoiGianChucVuChienDau', ''),
            chucVuDaQua=data.get('chucVuDaQua', ''),
            thoiGianChucVuDaQua=data.get('thoiGianChucVuDaQua', ''),
            cmQuan=data.get('cmQuan', ''),
            lienHeKhiCan=data.get('lienHeKhiCan', ''),
            soDienThoaiLienHe=data.get('soDienThoaiLienHe', ''),
            hoTenCha=data.get('hoTenCha', ''),
            hoTenMe=data.get('hoTenMe', ''),
            hoTenVo=data.get('hoTenVo', ''),
            hoTenNguoiThan=data.get('hoTenNguoiThan', ''),
            moiQuanHe=data.get('moiQuanHe', ''),
            noiDungNguoiThan=data.get('noiDungNguoiThan', ''),
            thamGiaNguyQuan=data.get('thamGiaNguyQuan', ''),
            thamGiaNguyQuyen=data.get('thamGiaNguyQuyen', ''),
            thamGiaNoMau=data.get('thamGiaNoMau', ''),
            daCaiTao=data.get('daCaiTao', ''),
            ghiChu=data.get('ghiChu', ''),
            ngoaiNgu=data.get('ngoaiNgu', ''),
            tiengDTTS=data.get('tiengDTTS', ''),
            thongTinKhac=thong_tin_khac,
        )