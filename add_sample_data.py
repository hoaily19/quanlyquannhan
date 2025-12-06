"""
Script thêm dữ liệu mẫu vào database để test
"""

import sys
import io
from pathlib import Path
import uuid
from datetime import datetime

# Fix encoding cho Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

sys.path.insert(0, str(Path(__file__).parent))

from services.database import DatabaseService
from models.personnel import Personnel, ThongTinDang, ThongTinDoan, ThongTinKhac

def add_sample_personnel():
    """Thêm 10 quân nhân mẫu với các dân tộc khác nhau"""
    
    db = DatabaseService()
    
    # Danh sách quân nhân mẫu với các dân tộc thiểu số
    sample_data = [
        {
            "hoTen": "Triệu Văn Dũng",
            "ngaySinh": "19/11/1991",
            "capBac": "4",
            "chucVu": "CTV",
            "donVi": "C3",
            "danToc": "Tày",
            "queQuan": "T. Tát Dài, Xã Chợ Rã, tỉnh Thái Nguyên",
            "truQuan": "T. Tam Trung, xã Tam Giang, tỉnh Đắk Lắk",
            "nhapNgu": "2010",
            "trinhDoVanHoa": "12/12"
        },
        {
            "hoTen": "Nguyễn Văn A",
            "ngaySinh": "15/03/1995",
            "capBac": "H1",
            "chucVu": "CS",
            "donVi": "C3",
            "danToc": "Nùng",
            "queQuan": "Xã A, huyện B, tỉnh Lạng Sơn",
            "truQuan": "Xã C, huyện D, tỉnh Đắk Lắk",
            "nhapNgu": "2013",
            "trinhDoVanHoa": "12/12"
        },
        {
            "hoTen": "Y Thị B",
            "ngaySinh": "23/09/2000",
            "capBac": "H1",
            "chucVu": "CS",
            "donVi": "C3",
            "danToc": "Ede",
            "queQuan": "Love Bocas 1. Sala, Hrung, tỉnh Gia Lai",
            "truQuan": "Xã E, huyện F, tỉnh Đắk Lắk",
            "nhapNgu": "2018",
            "trinhDoVanHoa": "12/12"
        },
        {
            "hoTen": "Vàng Văn C",
            "ngaySinh": "10/05/1998",
            "capBac": "H2",
            "chucVu": "LX",
            "donVi": "C3",
            "danToc": "Mông",
            "queQuan": "Xã G, huyện H, tỉnh Hà Giang",
            "truQuan": "Xã I, huyện J, tỉnh Đắk Lắk",
            "nhapNgu": "2016",
            "trinhDoVanHoa": "12/12"
        },
        {
            "hoTen": "Ksor Thạch",
            "ngaySinh": "22/11/2004",
            "capBac": "H3",
            "chucVu": "PT",
            "donVi": "C3",
            "danToc": "Ja Rai",
            "queQuan": "Thêm Linh ô, Xã lahrú, tỉnh Gia Lai",
            "truQuan": "Xã K, huyện L, tỉnh Đắk Lắk",
            "nhapNgu": "2022",
            "trinhDoVanHoa": "12/12"
        },
        {
            "hoTen": "Lý Văn D",
            "ngaySinh": "08/07/1996",
            "capBac": "3",
            "chucVu": "ND",
            "donVi": "C3",
            "danToc": "Sán chay",
            "queQuan": "Xã M, huyện N, tỉnh Bắc Giang",
            "truQuan": "Xã O, huyện P, tỉnh Đắk Lắk",
            "nhapNgu": "2014",
            "trinhDoVanHoa": "12/12"
        },
        {
            "hoTen": "Lò Văn E",
            "ngaySinh": "12/04/1999",
            "capBac": "H1",
            "chucVu": "ND",
            "donVi": "C3",
            "danToc": "Thái",
            "queQuan": "Xã Q, huyện R, tỉnh Sơn La",
            "truQuan": "Xã S, huyện T, tỉnh Đắk Lắk",
            "nhapNgu": "2017",
            "trinhDoVanHoa": "12/12"
        },
        {
            "hoTen": "Nông Văn F",
            "ngaySinh": "25/08/1997",
            "capBac": "2",
            "chucVu": "TV",
            "donVi": "C3",
            "danToc": "Nùng",
            "queQuan": "Xã U, huyện V, tỉnh Cao Bằng",
            "truQuan": "Xã W, huyện X, tỉnh Đắk Lắk",
            "nhapNgu": "2015",
            "trinhDoVanHoa": "12/12"
        },
        {
            "hoTen": "Hà Văn G",
            "ngaySinh": "30/12/2001",
            "capBac": "H2",
            "chucVu": "LX",
            "donVi": "C3",
            "danToc": "Tày",
            "queQuan": "Xã Y, huyện Z, tỉnh Tuyên Quang",
            "truQuan": "Xã AA, huyện BB, tỉnh Đắk Lắk",
            "nhapNgu": "2019",
            "trinhDoVanHoa": "12/12"
        },
        {
            "hoTen": "Đình Anh",
            "ngaySinh": "23/02/2004",
            "capBac": "H1",
            "chucVu": "ND",
            "donVi": "C3",
            "danToc": "Ba na",
            "queQuan": "Làng Hà Giao, xã Canh Liên, tỉnh Gia Lai",
            "truQuan": "Xã CC, huyện DD, tỉnh Đắk Lắk",
            "nhapNgu": "2022",
            "trinhDoVanHoa": "12/12"
        }
    ]
    
    print("=" * 60)
    print("DANG THEM DU LIEU MAU VAO DATABASE")
    print("=" * 60)
    
    added_count = 0
    
    for idx, data in enumerate(sample_data, 1):
        try:
            # Tạo quân nhân mới
            personnel = Personnel()
            personnel.id = str(uuid.uuid4())
            personnel.hoTen = data["hoTen"]
            personnel.ngaySinh = data["ngaySinh"]
            personnel.capBac = data["capBac"]
            personnel.chucVu = data["chucVu"]
            personnel.donVi = data["donVi"]
            personnel.danToc = data["danToc"]
            personnel.queQuan = data["queQuan"]
            personnel.truQuan = data["truQuan"]
            personnel.nhapNgu = data["nhapNgu"]
            personnel.trinhDoVanHoa = data["trinhDoVanHoa"]
            
            # Tạo thông tin khác
            personnel.thongTinKhac = ThongTinKhac()
            
            # Lưu vào database
            db.create(personnel)
            added_count += 1
            
            print(f"[OK] {idx}. Da them: {data['hoTen']} - Dan toc: {data['danToc']}")
            
        except Exception as e:
            print(f"[ERROR] {idx}. Loi khi them {data['hoTen']}: {str(e)}")
    
    print("\n" + "=" * 60)
    print(f"Hoan thanh! Da them {added_count}/10 quan nhan vao database")
    print("=" * 60)
    
    # Hiển thị thống kê
    print("\nTHONG KE DAN TOC:")
    all_personnel = db.get_all()
    ethnic_count = {}
    for person in all_personnel:
        if person.danToc:
            ethnic_count[person.danToc] = ethnic_count.get(person.danToc, 0) + 1
    
    for ethnic, count in sorted(ethnic_count.items()):
        print(f"  • {ethnic}: {count} người")
    
    print(f"\nTong so quan nhan trong database: {len(all_personnel)}")

if __name__ == "__main__":
    add_sample_personnel()
