"""
Utility đọc file Word và Excel
"""

from pathlib import Path
from typing import List
import re

import sys
from pathlib import Path

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from models.personnel import Personnel, ThongTinDang, ThongTinDoan, ThongTinKhac


def read_docx(file_path: Path) -> List[Personnel]:
    """Đọc file .docx"""
    try:
        from docx import Document
        
        doc = Document(file_path)
        personnel_list = []
        
        # Đọc từ bảng trong document
        for table in doc.tables:
            # Giả sử hàng đầu là header
            headers = [cell.text.strip() for cell in table.rows[0].cells]
            
            # Đọc các hàng dữ liệu
            for row in table.rows[1:]:
                cells = [cell.text.strip() for cell in row.cells]
                if len(cells) < 2:  # Bỏ qua hàng rỗng
                    continue
                
                # Map dữ liệu (cần điều chỉnh theo cấu trúc file thực tế)
                person = Personnel()
                person.hoTen = cells[0] if len(cells) > 0 else ""
                # Thêm mapping các trường khác tùy theo cấu trúc file
                
                if person.hoTen:
                    personnel_list.append(person)
        
        return personnel_list
        
    except ImportError:
        raise ImportError("Cần cài đặt python-docx: pip install python-docx")
    except Exception as e:
        print(f"Lỗi đọc file {file_path}: {e}")
        return []


def read_xls(file_path: Path) -> List[Personnel]:
    """Đọc file .xls hoặc .xlsx"""
    try:
        import pandas as pd
        
        # Thử đọc với openpyxl trước (cho .xlsx)
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
        except:
            # Fallback về xlrd cho .xls
            df = pd.read_excel(file_path, engine='xlrd')
        
        personnel_list = []
        
        for _, row in df.iterrows():
            person = Personnel()
            
            # Map các cột (cần điều chỉnh theo tên cột thực tế)
            # Giả sử các cột có tên tiếng Việt
            column_mapping = {
                'Họ và Tên': 'hoTen',
                'Ngày Sinh': 'ngaySinh',
                'Cấp Bậc': 'capBac',
                'Chức Vụ': 'chucVu',
                'Đơn Vị': 'donVi',
                'Nhập Ngũ': 'nhapNgu',
                'Quê Quán': 'queQuan',
                'Trú Quán': 'truQuan',
                'Dân Tộc': 'danToc',
                'Tôn Giáo': 'tonGiao',
                'Trình Độ Văn Hóa': 'trinhDoVanHoa',
            }
            
            for col_name, attr_name in column_mapping.items():
                if col_name in df.columns:
                    value = str(row[col_name]) if pd.notna(row[col_name]) else ""
                    setattr(person, attr_name, value)
            
            if person.hoTen and person.hoTen.strip():
                personnel_list.append(person)
        
        return personnel_list
        
    except ImportError:
        raise ImportError("Cần cài đặt pandas và openpyxl: pip install pandas openpyxl")
    except Exception as e:
        print(f"Lỗi đọc file {file_path}: {e}")
        return []


def read_office_files(folder_path: str) -> List[Personnel]:
    """
    Đọc tất cả file Word và Excel trong thư mục
    Args:
        folder_path: Đường dẫn thư mục
    Returns:
        Danh sách quân nhân
    """
    folder = Path(folder_path)
    if not folder.exists():
        raise FileNotFoundError(f"Thư mục không tồn tại: {folder_path}")
    
    all_personnel = []
    
    # Đọc file .docx
    for docx_file in folder.glob("*.docx"):
        print(f"Đang đọc: {docx_file.name}")
        personnel_list = read_docx(docx_file)
        all_personnel.extend(personnel_list)
    
    # Đọc file .doc (cần convert hoặc dùng thư viện khác)
    # Tạm thời bỏ qua .doc vì phức tạp hơn
    
    # Đọc file .xls và .xlsx
    for xls_file in folder.glob("*.xls*"):
        print(f"Đang đọc: {xls_file.name}")
        personnel_list = read_xls(xls_file)
        all_personnel.extend(personnel_list)
    
    return all_personnel
