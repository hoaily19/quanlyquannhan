"""
Trang nhập dữ liệu từ file
"""

import streamlit as st
from pathlib import Path
import sys

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from utils.file_reader import read_office_files


def show():
    """Hiển thị trang nhập dữ liệu"""
    st.title(" Nhập Dữ Liệu Từ File")
    st.markdown("---")
    
    st.markdown("""
    ### Hướng Dẫn
    
    1. Chọn thư mục chứa các file Word/Excel (thư mục `noidung`)
    2. Hệ thống sẽ tự động đọc và import dữ liệu
    3. Kiểm tra dữ liệu trước khi lưu
    """)
    
    # Chọn thư mục
    folder_path = st.text_input(
        "Đường dẫn thư mục chứa file",
        value="../noidung",
        help="Nhập đường dẫn đến thư mục chứa các file .doc, .docx, .xls, .xlsx"
    )
    
    if st.button(" Đọc File"):
        if not Path(folder_path).exists():
            st.error(f"Thư mục không tồn tại: {folder_path}")
            return
        
        with st.spinner("Đang đọc file..."):
            try:
                personnel_list = read_office_files(folder_path)
                
                if personnel_list:
                    st.success(f"Đã đọc được {len(personnel_list)} hồ sơ")
                    
                    # Hiển thị preview
                    st.subheader("Preview Dữ Liệu")
                    for idx, person in enumerate(personnel_list[:5], 1):  # Chỉ hiển thị 5 đầu
                        with st.expander(f"{idx}. {person.hoTen or 'Chưa có tên'}"):
                            st.json(person.to_dict())
                    
                    if len(personnel_list) > 5:
                        st.info(f"... và {len(personnel_list) - 5} hồ sơ khác")
                    
                    # Nút import
                    if st.button("💾 Import Tất Cả Vào Database", type="primary"):
                        db = DatabaseService()
                        imported = 0
                        skipped = 0
                        
                        for person in personnel_list:
                            # Kiểm tra xem đã tồn tại chưa (theo tên)
                            existing = db.search(person.hoTen)
                            if existing and any(p.hoTen == person.hoTen for p in existing):
                                skipped += 1
                                continue
                            
                            db.create(person)
                            imported += 1
                        
                        st.success(f"Đã import {imported} hồ sơ. Bỏ qua {skipped} hồ sơ trùng lặp.")
                        st.rerun()
                else:
                    st.warning("Không tìm thấy dữ liệu trong các file")
                    
            except Exception as e:
                st.error(f"Lỗi khi đọc file: {str(e)}")
                st.exception(e)




"""

import streamlit as st
from pathlib import Path
import sys

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from utils.file_reader import read_office_files


def show():
    """Hiển thị trang nhập dữ liệu"""
    st.title(" Nhập Dữ Liệu Từ File")
    st.markdown("---")
    
    st.markdown("""
    ### Hướng Dẫn
    
    1. Chọn thư mục chứa các file Word/Excel (thư mục `noidung`)
    2. Hệ thống sẽ tự động đọc và import dữ liệu
    3. Kiểm tra dữ liệu trước khi lưu
    """)
    
    # Chọn thư mục
    folder_path = st.text_input(
        "Đường dẫn thư mục chứa file",
        value="../noidung",
        help="Nhập đường dẫn đến thư mục chứa các file .doc, .docx, .xls, .xlsx"
    )
    
    if st.button(" Đọc File"):
        if not Path(folder_path).exists():
            st.error(f"Thư mục không tồn tại: {folder_path}")
            return
        
        with st.spinner("Đang đọc file..."):
            try:
                personnel_list = read_office_files(folder_path)
                
                if personnel_list:
                    st.success(f"Đã đọc được {len(personnel_list)} hồ sơ")
                    
                    # Hiển thị preview
                    st.subheader("Preview Dữ Liệu")
                    for idx, person in enumerate(personnel_list[:5], 1):  # Chỉ hiển thị 5 đầu
                        with st.expander(f"{idx}. {person.hoTen or 'Chưa có tên'}"):
                            st.json(person.to_dict())
                    
                    if len(personnel_list) > 5:
                        st.info(f"... và {len(personnel_list) - 5} hồ sơ khác")
                    
                    # Nút import
                    if st.button("💾 Import Tất Cả Vào Database", type="primary"):
                        db = DatabaseService()
                        imported = 0
                        skipped = 0
                        
                        for person in personnel_list:
                            # Kiểm tra xem đã tồn tại chưa (theo tên)
                            existing = db.search(person.hoTen)
                            if existing and any(p.hoTen == person.hoTen for p in existing):
                                skipped += 1
                                continue
                            
                            db.create(person)
                            imported += 1
                        
                        st.success(f"Đã import {imported} hồ sơ. Bỏ qua {skipped} hồ sơ trùng lặp.")
                        st.rerun()
                else:
                    st.warning("Không tìm thấy dữ liệu trong các file")
                    
            except Exception as e:
                st.error(f"Lỗi khi đọc file: {str(e)}")
                st.exception(e)