"""
Service quản lý database (SQLite hoặc Firebase)
"""

import sqlite3
import json
import sys
from typing import List, Optional, Dict, Any
from datetime import datetime
from pathlib import Path

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from models.personnel import Personnel


class DatabaseService:
    """Service quản lý database"""
    
    def __init__(self, db_path: str = "data/personnel.db", use_encryption: bool = True):
        """
        Khởi tạo database service
        Args:
            db_path: Đường dẫn file SQLite
            use_encryption: Có dùng encryption không
        """
        self.db_path = db_path
        self._init_database()
    
    def _init_database(self):
        """Khởi tạo database và tạo bảng nếu chưa có"""
        # Tạo thư mục nếu chưa có
        Path(self.db_path).parent.mkdir(parents=True, exist_ok=True)
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Tạo bảng personnel với các cột cơ bản
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS personnel (
                id TEXT PRIMARY KEY,
                hoTen TEXT NOT NULL,
                ngaySinh TEXT,
                capBac TEXT,
                chucVu TEXT,
                donVi TEXT,
                nhapNgu TEXT,
                queQuan TEXT,
                truQuan TEXT,
                danToc TEXT,
                tonGiao TEXT,
                trinhDoVanHoa TEXT,
                thongTinKhac TEXT,
                createdAt TEXT,
                updatedAt TEXT
            )
        """)
        
        # Thêm các cột mới nếu chưa có (migration)
        self._migrate_personnel_table(cursor)
        conn.commit()  # Commit migration ngay lập tức
        
        # Tạo bảng units
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS units (
                id TEXT PRIMARY KEY,
                ten TEXT NOT NULL,
                loai TEXT,
                parentId TEXT,
                personnelIds TEXT,
                ghiChu TEXT,
                createdAt TEXT,
                updatedAt TEXT
            )
        """)
        
        # Tạo bảng nguoi_than
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS nguoi_than (
                id TEXT PRIMARY KEY,
                personnelId TEXT,
                hoTen TEXT NOT NULL,
                ngaySinh TEXT,
                diaChi TEXT,
                soDienThoai TEXT,
                moiQuanHe TEXT,
                noiDung TEXT,
                ghiChu TEXT,
                createdAt TEXT,
                updatedAt TEXT
            )
        """)
        
        # Tạo bảng ban_chap_hanh_chi_doan
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS ban_chap_hanh_chi_doan (
                id TEXT PRIMARY KEY,
                personnelId TEXT NOT NULL,
                chucVuDoan TEXT,
                createdAt TEXT,
                updatedAt TEXT,
                UNIQUE(personnelId)
            )
        """)
        
        # Tạo bảng bao_ve_an_ninh
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS bao_ve_an_ninh (
                id TEXT PRIMARY KEY,
                personnelId TEXT NOT NULL,
                thoiGianVao TEXT,
                thoiGianRa TEXT,
                createdAt TEXT,
                updatedAt TEXT,
                UNIQUE(personnelId)
            )
        """)
        
        # Tạo bảng nguoi_than_che_do_cu (quân nhân có người thân tham gia chế độ cũ)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS nguoi_than_che_do_cu (
                id TEXT PRIMARY KEY,
                personnelId TEXT NOT NULL,
                createdAt TEXT,
                updatedAt TEXT,
                UNIQUE(personnelId)
            )
        """)
        
        # Tạo bảng to_dan_van (quân nhân trong tổ công tác dân vận)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS to_dan_van (
                id TEXT PRIMARY KEY,
                personnelId TEXT NOT NULL UNIQUE,
                ghiChu TEXT,
                createdAt TEXT,
                updatedAt TEXT
            )
        """)
        
        # Tạo bảng dang_vien_dien_tap (đảng viên tham gia diễn tập)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS dang_vien_dien_tap (
                id TEXT PRIMARY KEY,
                personnelId TEXT NOT NULL UNIQUE,
                ghiChu TEXT,
                createdAt TEXT,
                updatedAt TEXT
            )
        """)
        
        # Tạo bảng nguoi_than_dang_phai_phan_dong (quân nhân có người thân tham gia đảng phái phản động)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS nguoi_than_dang_phai_phan_dong (
                id TEXT PRIMARY KEY,
                personnelId TEXT NOT NULL UNIQUE,
                createdAt TEXT,
                updatedAt TEXT
            )
        """)
        
        # Migration: thêm cột ghiChu vào các bảng nếu chưa có
        try:
            cursor.execute("ALTER TABLE to_dan_van ADD COLUMN ghiChu TEXT")
        except sqlite3.OperationalError:
            pass  # Cột đã tồn tại
        
        try:
            cursor.execute("ALTER TABLE dang_vien_dien_tap ADD COLUMN ghiChu TEXT")
        except sqlite3.OperationalError:
            pass  # Cột đã tồn tại
        
        try:
            cursor.execute("ALTER TABLE nguoi_than_che_do_cu ADD COLUMN ghiChu TEXT")
        except sqlite3.OperationalError:
            pass  # Cột đã tồn tại
        
        conn.commit()
        conn.close()
    
    def _migrate_personnel_table(self, cursor):
        """Migration: thêm các cột mới vào bảng personnel"""
        # Lấy danh sách cột hiện có
        cursor.execute("PRAGMA table_info(personnel)")
        existing_columns = [row[1] for row in cursor.fetchall()]
        
        # Danh sách cột mới cần thêm
        new_columns = {
            'hoTenThuongDung': 'TEXT',
            'ngayNhanCapBac': 'TEXT',
            'ngayNhanChucVu': 'TEXT',
            'unitId': 'TEXT',
            'xuatNgu': 'TEXT',
            'thanhPhanGiaDinh': 'TEXT',
            'quaTruong': 'TEXT',
            'nganhHoc': 'TEXT',
            'capHoc': 'TEXT',
            'thoiGianDaoTao': 'TEXT',
            'ketQuaDaoTao': 'TEXT',
            'chucVuChienDau': 'TEXT',
            'thoiGianChucVuChienDau': 'TEXT',
            'chucVuDaQua': 'TEXT',
            'thoiGianChucVuDaQua': 'TEXT',
            'cmQuan': 'TEXT',
            'lienHeKhiCan': 'TEXT',
            'soDienThoaiLienHe': 'TEXT',
            'hoTenCha': 'TEXT',
            'hoTenMe': 'TEXT',
            'hoTenVo': 'TEXT',
            'hoTenNguoiThan': 'TEXT',
            'moiQuanHe': 'TEXT',
            'noiDungNguoiThan': 'TEXT',
            'ghiChu': 'TEXT',
            'thamGiaNguyQuan': 'TEXT',
            'thamGiaNguyQuyen': 'TEXT',
            'thamGiaNoMau': 'TEXT',
            'daCaiTao': 'TEXT',
            'ngoaiNgu': 'TEXT',
            'tiengDTTS': 'TEXT',
        }
        
        # Thêm các cột chưa có
        for col_name, col_type in new_columns.items():
            if col_name not in existing_columns:
                try:
                    cursor.execute(f"ALTER TABLE personnel ADD COLUMN {col_name} {col_type}")
                    print(f"Đã thêm cột {col_name} vào bảng personnel")
                except sqlite3.OperationalError as e:
                    print(f"Lỗi khi thêm cột {col_name}: {e}")
                    pass  # Cột đã tồn tại hoặc lỗi khác
    
    def get_all(self) -> List[Personnel]:
        """Lấy tất cả quân nhân"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM personnel ORDER BY hoTen")
        rows = cursor.fetchall()
        conn.close()
        
        result = []
        for row in rows:
            data = dict(row)
            if data.get('thongTinKhac'):
                data['thongTinKhac'] = json.loads(data['thongTinKhac'])
            result.append(Personnel.from_dict(data))
        
        return result
    
    def get_by_id(self, personnel_id: str) -> Optional[Personnel]:
        """Lấy quân nhân theo ID"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM personnel WHERE id = ?", (personnel_id,))
        row = cursor.fetchone()
        conn.close()
        
        if not row:
            return None
        
        data = dict(row)
        if data.get('thongTinKhac'):
            data['thongTinKhac'] = json.loads(data['thongTinKhac'])
        
        return Personnel.from_dict(data)
    
    def create(self, personnel: Personnel) -> str:
        """Tạo quân nhân mới"""
        import uuid
        
        if not personnel.id:
            personnel.id = str(uuid.uuid4())
        
        personnel.createdAt = datetime.now()
        personnel.updatedAt = datetime.now()
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        data = personnel.to_dict()

        # Danh sách cột và values tương ứng - đảm bảo số lượng cột khớp hoàn toàn với số values
        columns = [
            'id', 'hoTen', 'hoTenThuongDung', 'ngaySinh', 'capBac', 'ngayNhanCapBac',
            'chucVu', 'ngayNhanChucVu', 'donVi', 'unitId', 'nhapNgu', 'xuatNgu',
            'queQuan', 'truQuan', 'danToc', 'tonGiao', 'trinhDoVanHoa', 'thanhPhanGiaDinh',
            'quaTruong', 'nganhHoc', 'capHoc', 'thoiGianDaoTao', 'ketQuaDaoTao',
            'chucVuChienDau', 'thoiGianChucVuChienDau', 'chucVuDaQua', 'thoiGianChucVuDaQua', 'cmQuan',
            'lienHeKhiCan', 'soDienThoaiLienHe',
            'hoTenCha', 'hoTenMe', 'hoTenVo',
            'hoTenNguoiThan', 'moiQuanHe', 'noiDungNguoiThan',
            'thamGiaNguyQuan', 'thamGiaNguyQuyen', 'thamGiaNoMau', 'daCaiTao',
            'ghiChu', 'ngoaiNgu', 'tiengDTTS', 'thongTinKhac', 'createdAt', 'updatedAt',
        ]

        placeholders = ", ".join(["?"] * len(columns))

        values = (
            data['id'],
            data['hoTen'],
            data.get('hoTenThuongDung', ''),
            data['ngaySinh'],
            data['capBac'],
            data.get('ngayNhanCapBac', ''),
            data['chucVu'],
            data.get('ngayNhanChucVu', ''),
            data['donVi'],
            data.get('unitId'),
            data['nhapNgu'],
            data.get('xuatNgu', ''),
            data['queQuan'],
            data['truQuan'],
            data['danToc'],
            data['tonGiao'],
            data['trinhDoVanHoa'],
            data.get('thanhPhanGiaDinh', ''),
            data.get('quaTruong', ''),
            data.get('nganhHoc', ''),
            data.get('capHoc', ''),
            data.get('thoiGianDaoTao', ''),
            data.get('ketQuaDaoTao', ''),
            data.get('chucVuChienDau', ''),
            data.get('thoiGianChucVuChienDau', ''),
            data.get('chucVuDaQua', ''),
            data.get('thoiGianChucVuDaQua', ''),
            data.get('cmQuan', ''),
            data.get('lienHeKhiCan', ''),
            data.get('soDienThoaiLienHe', ''),
            data.get('hoTenCha', ''),
            data.get('hoTenMe', ''),
            data.get('hoTenVo', ''),
            data.get('hoTenNguoiThan', ''),
            data.get('moiQuanHe', ''),
            data.get('noiDungNguoiThan', ''),
            data.get('thamGiaNguyQuan', ''),
            data.get('thamGiaNguyQuyen', ''),
            data.get('thamGiaNoMau', ''),
            data.get('daCaiTao', ''),
            data.get('ghiChu', ''),
            data.get('ngoaiNgu', ''),
            data.get('tiengDTTS', ''),
            json.dumps(data['thongTinKhac']),
            data['createdAt'],
            data['updatedAt'],
        )

        cursor.execute(
            f"INSERT INTO personnel ({', '.join(columns)}) VALUES ({placeholders})",
            values,
        )
        
        conn.commit()
        conn.close()
        
        return personnel.id
    
    def update(self, personnel: Personnel) -> bool:
        """Cập nhật quân nhân"""
        personnel.updatedAt = datetime.now()
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        data = personnel.to_dict()
        cursor.execute("""
            UPDATE personnel SET
                hoTen = ?, hoTenThuongDung = ?, ngaySinh = ?, capBac = ?, ngayNhanCapBac = ?,
                chucVu = ?, ngayNhanChucVu = ?, donVi = ?, unitId = ?,
                nhapNgu = ?, xuatNgu = ?, queQuan = ?, truQuan = ?,
                danToc = ?, tonGiao = ?, trinhDoVanHoa = ?, thanhPhanGiaDinh = ?,
                quaTruong = ?, nganhHoc = ?, capHoc = ?, thoiGianDaoTao = ?, ketQuaDaoTao = ?,
                chucVuChienDau = ?, thoiGianChucVuChienDau = ?, chucVuDaQua = ?, thoiGianChucVuDaQua = ?, cmQuan = ?,
                lienHeKhiCan = ?, soDienThoaiLienHe = ?,
                hoTenCha = ?, hoTenMe = ?, hoTenVo = ?,
                hoTenNguoiThan = ?, moiQuanHe = ?, noiDungNguoiThan = ?,
                thamGiaNguyQuan = ?, thamGiaNguyQuyen = ?, thamGiaNoMau = ?, daCaiTao = ?,
                ghiChu = ?, ngoaiNgu = ?, tiengDTTS = ?, thongTinKhac = ?, updatedAt = ?
            WHERE id = ?
        """, (
            data['hoTen'],
            data.get('hoTenThuongDung', ''),
            data['ngaySinh'],
            data['capBac'],
            data.get('ngayNhanCapBac', ''),
            data['chucVu'],
            data.get('ngayNhanChucVu', ''),
            data['donVi'],
            data.get('unitId'),
            data['nhapNgu'],
            data.get('xuatNgu', ''),
            data['queQuan'],
            data['truQuan'],
            data['danToc'],
            data['tonGiao'],
            data['trinhDoVanHoa'],
            data.get('thanhPhanGiaDinh', ''),
            data.get('quaTruong', ''),
            data.get('nganhHoc', ''),
            data.get('capHoc', ''),
            data.get('thoiGianDaoTao', ''),
            data.get('ketQuaDaoTao', ''),
            data.get('chucVuChienDau', ''),
            data.get('thoiGianChucVuChienDau', ''),
            data.get('chucVuDaQua', ''),
            data.get('thoiGianChucVuDaQua', ''),
            data.get('cmQuan', ''),
            data.get('lienHeKhiCan', ''),
            data.get('soDienThoaiLienHe', ''),
            data.get('hoTenCha', ''),
            data.get('hoTenMe', ''),
            data.get('hoTenVo', ''),
            data.get('hoTenNguoiThan', ''),
            data.get('moiQuanHe', ''),
            data.get('noiDungNguoiThan', ''),
            data.get('thamGiaNguyQuan', ''),
            data.get('thamGiaNguyQuyen', ''),
            data.get('thamGiaNoMau', ''),
            data.get('daCaiTao', ''),
            data.get('ghiChu', ''),
            data.get('ngoaiNgu', ''),
            data.get('tiengDTTS', ''),
            json.dumps(data['thongTinKhac']),
            data['updatedAt'],
            data['id'],
        ))
        
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        
        return success
    
    def delete(self, personnel_id: str) -> bool:
        """Xóa quân nhân"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("DELETE FROM personnel WHERE id = ?", (personnel_id,))
        success = cursor.rowcount > 0
        
        conn.commit()
        conn.close()
        
        return success
    
    def search(self, query: str, filters: Optional[Dict[str, Any]] = None) -> List[Personnel]:
        """Tìm kiếm quân nhân"""
        all_personnel = self.get_all()
        results = []
        
        query_lower = query.lower() if query else ""
        
        for person in all_personnel:
            # Tìm kiếm theo tên
            if query_lower and query_lower not in (person.hoTen or "").lower():
                continue
            
            # Lọc theo các tiêu chí
            if filters:
                if filters.get('donVi') and person.donVi != filters['donVi']:
                    continue
                if filters.get('capBac') and person.capBac != filters['capBac']:
                    continue
                if filters.get('chucVu') and person.chucVu != filters['chucVu']:
                    continue
                if filters.get('danToc') and person.danToc != filters['danToc']:
                    continue
                if filters.get('tonGiao') and person.tonGiao != filters['tonGiao']:
                    continue
            
            results.append(person)
        
        return results
    
    def get_unique_values(self, field: str) -> List[str]:
        """Lấy danh sách giá trị unique của một trường"""
        all_personnel = self.get_all()
        values = set()
        
        for person in all_personnel:
            value = getattr(person, field, None)
            if value:
                values.add(value)
        
        return sorted(list(values))
    
    # ========== Unit Management ==========
    
    def get_all_units(self):
        """Lấy tất cả đơn vị"""
        try:
            from models.unit import Unit
            
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM units ORDER BY loai, ten")
            rows = cursor.fetchall()
            conn.close()
            
            result = []
            for row in rows:
                data = dict(row)
                if data.get('personnelIds'):
                    data['personnelIds'] = json.loads(data['personnelIds'])
                else:
                    data['personnelIds'] = []
                result.append(Unit.from_dict(data))
            
            return result
        except sqlite3.OperationalError:
            # Bảng chưa tồn tại
            return []
    
    def get_unit_by_id(self, unit_id: str):
        """Lấy đơn vị theo ID"""
        try:
            from models.unit import Unit
            
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM units WHERE id = ?", (unit_id,))
            row = cursor.fetchone()
            conn.close()
            
            if not row:
                return None
            
            data = dict(row)
            if data.get('personnelIds'):
                data['personnelIds'] = json.loads(data['personnelIds'])
            else:
                data['personnelIds'] = []
            
            return Unit.from_dict(data)
        except sqlite3.OperationalError:
            return None
    
    def get_units_by_parent_id(self, parent_id: str):
        """Lấy danh sách đơn vị con theo parent ID"""
        try:
            from models.unit import Unit
            
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM units WHERE parentId = ? ORDER BY ten", (parent_id,))
            rows = cursor.fetchall()
            conn.close()
            
            result = []
            for row in rows:
                data = dict(row)
                if data.get('personnelIds'):
                    data['personnelIds'] = json.loads(data['personnelIds'])
                else:
                    data['personnelIds'] = []
                result.append(Unit.from_dict(data))
            
            return result
        except sqlite3.OperationalError:
            return []
    
    def create_unit(self, unit) -> str:
        """Tạo đơn vị mới"""
        import uuid
        
        if not unit.id:
            unit.id = str(uuid.uuid4())
        
        unit.createdAt = datetime.now()
        unit.updatedAt = datetime.now()
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        data = unit.to_dict()
        cursor.execute("""
            INSERT INTO units (
                id, ten, loai, parentId, personnelIds, ghiChu, createdAt, updatedAt
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            data['id'],
            data['ten'],
            data['loai'],
            data['parentId'],
            json.dumps(data['personnelIds']),
            data['ghiChu'],
            data['createdAt'],
            data['updatedAt'],
        ))
        
        conn.commit()
        conn.close()
        
        return unit.id
    
    def update_unit(self, unit) -> bool:
        """Cập nhật đơn vị"""
        unit.updatedAt = datetime.now()
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        data = unit.to_dict()
        cursor.execute("""
            UPDATE units SET
                ten = ?, loai = ?, parentId = ?, personnelIds = ?,
                ghiChu = ?, updatedAt = ?
            WHERE id = ?
        """, (
            data['ten'],
            data['loai'],
            data['parentId'],
            json.dumps(data['personnelIds']),
            data['ghiChu'],
            data['updatedAt'],
            data['id'],
        ))
        
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        
        return success
    
    def delete_unit(self, unit_id: str) -> bool:
        """Xóa đơn vị"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("DELETE FROM units WHERE id = ?", (unit_id,))
        success = cursor.rowcount > 0
        
        conn.commit()
        conn.close()
        
        return success
    
    def get_personnel_by_unit(self, unit_id: str) -> List[Personnel]:
        """Lấy danh sách quân nhân trong đơn vị"""
        unit = self.get_unit_by_id(unit_id)
        if not unit:
            return []
        
        all_personnel = self.get_all()
        result = []
        for person in all_personnel:
            if person.id in unit.personnelIds:
                result.append(person)
        
        return result
    
    # ========== NguoiThan Management ==========
    
    def get_nguoi_than_by_personnel(self, personnel_id: str) -> List:
        """Lấy danh sách người thân của quân nhân"""
        try:
            from models.nguoi_than import NguoiThan
            
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM nguoi_than WHERE personnelId = ? ORDER BY hoTen", (personnel_id,))
            rows = cursor.fetchall()
            conn.close()
            
            result = []
            for row in rows:
                data = dict(row)
                result.append(NguoiThan.from_dict(data))
            
            return result
        except sqlite3.OperationalError:
            return []
        except Exception:
            return []
    
    def get_nguoi_than_by_id(self, nguoi_than_id: str):
        """Lấy người thân theo ID"""
        try:
            from models.nguoi_than import NguoiThan
            
            conn = sqlite3.connect(self.db_path)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            cursor.execute("SELECT * FROM nguoi_than WHERE id = ?", (nguoi_than_id,))
            row = cursor.fetchone()
            conn.close()
            
            if not row:
                return None
            
            return NguoiThan.from_dict(dict(row))
        except Exception:
            return None
    
    def create_nguoi_than(self, nguoi_than) -> str:
        """Tạo người thân mới"""
        import uuid
        from models.nguoi_than import NguoiThan
        
        if not nguoi_than.id:
            nguoi_than.id = str(uuid.uuid4())
        
        nguoi_than.createdAt = datetime.now()
        nguoi_than.updatedAt = datetime.now()
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        data = nguoi_than.to_dict()
        cursor.execute("""
            INSERT INTO nguoi_than (
                id, personnelId, hoTen, ngaySinh, diaChi, soDienThoai,
                moiQuanHe, noiDung, ghiChu, createdAt, updatedAt
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            data['id'],
            data['personnelId'],
            data['hoTen'],
            data['ngaySinh'],
            data['diaChi'],
            data['soDienThoai'],
            data['moiQuanHe'],
            data['noiDung'],
            data['ghiChu'],
            data['createdAt'],
            data['updatedAt'],
        ))
        
        conn.commit()
        conn.close()
        
        return nguoi_than.id
    
    def update_nguoi_than(self, nguoi_than) -> bool:
        """Cập nhật người thân"""
        from models.nguoi_than import NguoiThan
        
        nguoi_than.updatedAt = datetime.now()
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        data = nguoi_than.to_dict()
        cursor.execute("""
            UPDATE nguoi_than SET
                hoTen = ?, ngaySinh = ?, diaChi = ?, soDienThoai = ?,
                moiQuanHe = ?, noiDung = ?, ghiChu = ?, updatedAt = ?
            WHERE id = ?
        """, (
            data['hoTen'],
            data['ngaySinh'],
            data['diaChi'],
            data['soDienThoai'],
            data['moiQuanHe'],
            data['noiDung'],
            data['ghiChu'],
            data['updatedAt'],
            data['id'],
        ))
        
        success = cursor.rowcount > 0
        conn.commit()
        conn.close()
        
        return success
    
    def delete_nguoi_than(self, nguoi_than_id: str) -> bool:
        """Xóa người thân"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("DELETE FROM nguoi_than WHERE id = ?", (nguoi_than_id,))
        success = cursor.rowcount > 0
        
        conn.commit()
        conn.close()
        
        return success
    
    def get_ban_chap_hanh_chi_doan(self) -> List[str]:
        """Lấy danh sách ID quân nhân trong ban chấp hành chi đoàn"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT personnelId FROM ban_chap_hanh_chi_doan")
            rows = cursor.fetchall()
            conn.close()
            return [row[0] for row in rows]
        except Exception:
            return []
    
    def add_ban_chap_hanh_chi_doan(self, personnel_id: str, chuc_vu_doan: str = "") -> bool:
        """Thêm quân nhân vào ban chấp hành chi đoàn"""
        try:
            import uuid
            from datetime import datetime
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Kiểm tra xem đã có chưa
            cursor.execute("SELECT id FROM ban_chap_hanh_chi_doan WHERE personnelId = ?", (personnel_id,))
            existing = cursor.fetchone()
            
            if existing:
                # Cập nhật
                cursor.execute("""
                    UPDATE ban_chap_hanh_chi_doan 
                    SET chucVuDoan = ?, updatedAt = ?
                    WHERE personnelId = ?
                """, (chuc_vu_doan, datetime.now().isoformat(), personnel_id))
            else:
                # Thêm mới
                ban_chap_hanh_id = str(uuid.uuid4())
                cursor.execute("""
                    INSERT INTO ban_chap_hanh_chi_doan (id, personnelId, chucVuDoan, createdAt, updatedAt)
                    VALUES (?, ?, ?, ?, ?)
                """, (ban_chap_hanh_id, personnel_id, chuc_vu_doan, 
                      datetime.now().isoformat(), datetime.now().isoformat()))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error adding ban chap hanh: {e}")
            return False
    
    def remove_ban_chap_hanh_chi_doan(self, personnel_id: str) -> bool:
        """Xóa quân nhân khỏi ban chấp hành chi đoàn"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM ban_chap_hanh_chi_doan WHERE personnelId = ?", (personnel_id,))
            conn.commit()
            conn.close()
            return True
        except Exception:
            return False
    
    def get_chuc_vu_doan(self, personnel_id: str) -> str:
        """Lấy chức vụ đoàn của quân nhân trong ban chấp hành"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT chucVuDoan FROM ban_chap_hanh_chi_doan WHERE personnelId = ?", (personnel_id,))
            row = cursor.fetchone()
            conn.close()
            return row[0] if row and row[0] else ""
        except Exception:
            return ""
    
    def get_bao_ve_an_ninh(self) -> List[str]:
        """Lấy danh sách ID quân nhân trong bảo vệ an ninh"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT personnelId FROM bao_ve_an_ninh")
            rows = cursor.fetchall()
            conn.close()
            return [row[0] for row in rows]
        except Exception:
            return []
    
    def add_bao_ve_an_ninh(self, personnel_id: str, thoi_gian_vao: str = "", thoi_gian_ra: str = "") -> bool:
        """Thêm quân nhân vào bảo vệ an ninh"""
        try:
            import uuid
            from datetime import datetime
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Kiểm tra xem đã có chưa
            cursor.execute("SELECT id FROM bao_ve_an_ninh WHERE personnelId = ?", (personnel_id,))
            existing = cursor.fetchone()
            
            if existing:
                # Cập nhật
                cursor.execute("""
                    UPDATE bao_ve_an_ninh 
                    SET thoiGianVao = ?, thoiGianRa = ?, updatedAt = ?
                    WHERE personnelId = ?
                """, (thoi_gian_vao, thoi_gian_ra, datetime.now().isoformat(), personnel_id))
            else:
                # Thêm mới
                bao_ve_id = str(uuid.uuid4())
                cursor.execute("""
                    INSERT INTO bao_ve_an_ninh (id, personnelId, thoiGianVao, thoiGianRa, createdAt, updatedAt)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (bao_ve_id, personnel_id, thoi_gian_vao, thoi_gian_ra, 
                      datetime.now().isoformat(), datetime.now().isoformat()))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error adding bao ve an ninh: {e}")
            return False
    
    def remove_bao_ve_an_ninh(self, personnel_id: str) -> bool:
        """Xóa quân nhân khỏi bảo vệ an ninh"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM bao_ve_an_ninh WHERE personnelId = ?", (personnel_id,))
            conn.commit()
            conn.close()
            return True
        except Exception:
            return False
    
    # ========== Nguoi Than Che Do Cu Management ==========
    
    def get_nguoi_than_che_do_cu(self) -> List[str]:
        """Lấy danh sách ID quân nhân có người thân tham gia chế độ cũ"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT personnelId FROM nguoi_than_che_do_cu")
            result = [row[0] for row in cursor.fetchall()]
            conn.close()
            return result
        except Exception:
            return []
    
    def add_nguoi_than_che_do_cu(self, personnel_id: str) -> bool:
        """Thêm quân nhân vào danh sách có người thân tham gia chế độ cũ"""
        try:
            import uuid
            from datetime import datetime
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Kiểm tra xem đã có chưa
            cursor.execute("SELECT id FROM nguoi_than_che_do_cu WHERE personnelId = ?", (personnel_id,))
            existing = cursor.fetchone()
            
            if existing:
                # Cập nhật
                cursor.execute("""
                    UPDATE nguoi_than_che_do_cu 
                    SET updatedAt = ?
                    WHERE personnelId = ?
                """, (datetime.now().isoformat(), personnel_id))
            else:
                # Thêm mới
                record_id = str(uuid.uuid4())
                cursor.execute("""
                    INSERT INTO nguoi_than_che_do_cu (id, personnelId, createdAt, updatedAt)
                    VALUES (?, ?, ?, ?)
                """, (record_id, personnel_id, 
                      datetime.now().isoformat(), datetime.now().isoformat()))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error adding nguoi than che do cu: {e}")
            return False
    
    def remove_nguoi_than_che_do_cu(self, personnel_id: str) -> bool:
        """Xóa quân nhân khỏi danh sách có người thân tham gia chế độ cũ"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM nguoi_than_che_do_cu WHERE personnelId = ?", (personnel_id,))
            conn.commit()
            conn.close()
            return True
        except Exception:
            return False
    
    def get_nguoi_than_che_do_cu_ghi_chu(self, personnel_id: str) -> str:
        """Lấy ghi chú riêng của quân nhân trong người thân chế độ cũ"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT ghiChu FROM nguoi_than_che_do_cu WHERE personnelId = ?", (personnel_id,))
            row = cursor.fetchone()
            conn.close()
            return row[0] if row and row[0] else ''
        except Exception:
            return ''
    
    def update_nguoi_than_che_do_cu_ghi_chu(self, personnel_id: str, ghi_chu: str) -> bool:
        """Cập nhật ghi chú riêng của quân nhân trong người thân chế độ cũ"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE nguoi_than_che_do_cu 
                SET ghiChu = ?, updatedAt = ?
                WHERE personnelId = ?
            """, (ghi_chu, datetime.now().isoformat(), personnel_id))
            conn.commit()
            conn.close()
            return True
        except Exception:
            return False
    
    # ========== To Dan Van Management ==========
    
    def get_to_dan_van(self) -> List[str]:
        """Lấy danh sách ID quân nhân trong tổ công tác dân vận"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT personnelId FROM to_dan_van")
            result = [row[0] for row in cursor.fetchall()]
            conn.close()
            return result
        except Exception:
            return []
    
    def get_to_dan_van_ghi_chu(self, personnel_id: str) -> str:
        """Lấy ghi chú riêng của quân nhân trong tổ công tác dân vận"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT ghiChu FROM to_dan_van WHERE personnelId = ?", (personnel_id,))
            row = cursor.fetchone()
            conn.close()
            return row[0] if row and row[0] else ''
        except Exception:
            return ''
    
    def update_to_dan_van_ghi_chu(self, personnel_id: str, ghi_chu: str) -> bool:
        """Cập nhật ghi chú riêng của quân nhân trong tổ công tác dân vận"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE to_dan_van 
                SET ghiChu = ?, updatedAt = ?
                WHERE personnelId = ?
            """, (ghi_chu, datetime.now().isoformat(), personnel_id))
            conn.commit()
            conn.close()
            return True
        except Exception:
            return False
    
    def add_to_dan_van(self, personnel_id: str) -> bool:
        """Thêm quân nhân vào tổ công tác dân vận"""
        try:
            import uuid
            from datetime import datetime
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Kiểm tra xem đã có chưa
            cursor.execute("SELECT id FROM to_dan_van WHERE personnelId = ?", (personnel_id,))
            existing = cursor.fetchone()
            
            if existing:
                # Cập nhật
                cursor.execute("""
                    UPDATE to_dan_van 
                    SET updatedAt = ?
                    WHERE personnelId = ?
                """, (datetime.now().isoformat(), personnel_id))
            else:
                # Thêm mới
                record_id = str(uuid.uuid4())
                cursor.execute("""
                    INSERT INTO to_dan_van (id, personnelId, createdAt, updatedAt)
                    VALUES (?, ?, ?, ?)
                """, (record_id, personnel_id, 
                      datetime.now().isoformat(), datetime.now().isoformat()))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error adding to dan van: {e}")
            return False
    
    def remove_to_dan_van(self, personnel_id: str) -> bool:
        """Xóa quân nhân khỏi tổ công tác dân vận"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM to_dan_van WHERE personnelId = ?", (personnel_id,))
            conn.commit()
            conn.close()
            return True
        except Exception:
            return False
    
    # ========== Dang Vien Dien Tap Management ==========
    
    def get_dang_vien_dien_tap(self) -> List[str]:
        """Lấy danh sách ID quân nhân đảng viên tham gia diễn tập"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT personnelId FROM dang_vien_dien_tap")
            result = [row[0] for row in cursor.fetchall()]
            conn.close()
            return result
        except Exception:
            return []
    
    def get_dang_vien_dien_tap_ghi_chu(self, personnel_id: str) -> str:
        """Lấy ghi chú riêng của quân nhân trong đảng viên diễn tập"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT ghiChu FROM dang_vien_dien_tap WHERE personnelId = ?", (personnel_id,))
            row = cursor.fetchone()
            conn.close()
            return row[0] if row and row[0] else ''
        except Exception:
            return ''
    
    def update_dang_vien_dien_tap_ghi_chu(self, personnel_id: str, ghi_chu: str) -> bool:
        """Cập nhật ghi chú riêng của quân nhân trong đảng viên diễn tập"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE dang_vien_dien_tap 
                SET ghiChu = ?, updatedAt = ?
                WHERE personnelId = ?
            """, (ghi_chu, datetime.now().isoformat(), personnel_id))
            conn.commit()
            conn.close()
            return True
        except Exception:
            return False
    
    def add_dang_vien_dien_tap(self, personnel_id: str) -> bool:
        """Thêm quân nhân vào danh sách đảng viên tham gia diễn tập"""
        try:
            import uuid
            from datetime import datetime
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Kiểm tra xem đã có chưa
            cursor.execute("SELECT id FROM dang_vien_dien_tap WHERE personnelId = ?", (personnel_id,))
            existing = cursor.fetchone()
            
            if existing:
                # Cập nhật
                cursor.execute("""
                    UPDATE dang_vien_dien_tap 
                    SET updatedAt = ?
                    WHERE personnelId = ?
                """, (datetime.now().isoformat(), personnel_id))
            else:
                # Thêm mới
                record_id = str(uuid.uuid4())
                cursor.execute("""
                    INSERT INTO dang_vien_dien_tap (id, personnelId, createdAt, updatedAt)
                    VALUES (?, ?, ?, ?)
                """, (record_id, personnel_id, 
                      datetime.now().isoformat(), datetime.now().isoformat()))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error adding dang vien dien tap: {e}")
            return False
    
    def remove_dang_vien_dien_tap(self, personnel_id: str) -> bool:
        """Xóa quân nhân khỏi danh sách đảng viên tham gia diễn tập"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM dang_vien_dien_tap WHERE personnelId = ?", (personnel_id,))
            conn.commit()
            conn.close()
            return True
        except Exception:
            return False
    
    # ========== Nguoi Than Dang Phai Phan Dong Management ==========
    
    def get_nguoi_than_dang_phai_phan_dong(self) -> List[str]:
        """Lấy danh sách ID quân nhân có người thân tham gia đảng phái phản động"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT personnelId FROM nguoi_than_dang_phai_phan_dong")
            result = [row[0] for row in cursor.fetchall()]
            conn.close()
            return result
        except Exception:
            return []
    
    def add_nguoi_than_dang_phai_phan_dong(self, personnel_id: str) -> bool:
        """Thêm quân nhân vào danh sách có người thân tham gia đảng phái phản động"""
        try:
            import uuid
            from datetime import datetime
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Kiểm tra xem đã có chưa
            cursor.execute("SELECT id FROM nguoi_than_dang_phai_phan_dong WHERE personnelId = ?", (personnel_id,))
            existing = cursor.fetchone()
            
            if existing:
                # Cập nhật
                cursor.execute("""
                    UPDATE nguoi_than_dang_phai_phan_dong 
                    SET updatedAt = ?
                    WHERE personnelId = ?
                """, (datetime.now().isoformat(), personnel_id))
            else:
                # Thêm mới
                record_id = str(uuid.uuid4())
                cursor.execute("""
                    INSERT INTO nguoi_than_dang_phai_phan_dong (id, personnelId, createdAt, updatedAt)
                    VALUES (?, ?, ?, ?)
                """, (record_id, personnel_id, 
                      datetime.now().isoformat(), datetime.now().isoformat()))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"Error adding nguoi than dang phai phan dong: {e}")
            return False
    
    def remove_nguoi_than_dang_phai_phan_dong(self, personnel_id: str) -> bool:
        """Xóa quân nhân khỏi danh sách có người thân tham gia đảng phái phản động"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM nguoi_than_dang_phai_phan_dong WHERE personnelId = ?", (personnel_id,))
            conn.commit()
            conn.close()
            return True
        except Exception:
            return False
    
    def get_bao_ve_an_ninh_info(self, personnel_id: str) -> dict:
        """Lấy thông tin bảo vệ an ninh của quân nhân"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("SELECT thoiGianVao, thoiGianRa FROM bao_ve_an_ninh WHERE personnelId = ?", (personnel_id,))
            row = cursor.fetchone()
            conn.close()
            if row:
                return {'thoiGianVao': row[0] or '', 'thoiGianRa': row[1] or ''}
            return {'thoiGianVao': '', 'thoiGianRa': ''}
        except Exception:
            return {'thoiGianVao': '', 'thoiGianRa': ''}
