"""
Trang danh sách quân nhân
"""

import streamlit as st
import sys
from pathlib import Path

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from services.export import ExportService


def show():
    """Hiển thị danh sách quân nhân"""
    st.title("📋 Danh Sách Quân Nhân")
    st.markdown("---")
    
    db = DatabaseService()
    
    # Tìm kiếm và lọc
    col1, col2 = st.columns([3, 1])
    
    with col1:
        search_query = st.text_input("🔍 Tìm kiếm theo tên", "")
    
    with col2:
        if st.button("➕ Thêm Mới"):
            st.session_state['edit_personnel_id'] = None
            st.rerun()
    
    # Filters
    with st.expander("🔽 Bộ Lọc", expanded=False):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            don_vi_filter = st.selectbox(
                "Đơn Vị",
                [""] + db.get_unique_values('donVi')
            )
        
        with col2:
            cap_bac_filter = st.selectbox(
                "Cấp Bậc",
                [""] + db.get_unique_values('capBac')
            )
        
        with col3:
            chuc_vu_filter = st.selectbox(
                "Chức Vụ",
                [""] + db.get_unique_values('chucVu')
            )
    
    # Tìm kiếm
    filters = {}
    if don_vi_filter:
        filters['donVi'] = don_vi_filter
    if cap_bac_filter:
        filters['capBac'] = cap_bac_filter
    if chuc_vu_filter:
        filters['chucVu'] = chuc_vu_filter
    
    personnel_list = db.search(search_query, filters if filters else None)
    
    # Hiển thị kết quả
    st.markdown(f"**Tìm thấy: {len(personnel_list)} quân nhân**")
    
    if personnel_list:
        # Nút xuất file
        col1, col2 = st.columns([1, 4])
        with col1:
            csv_data = ExportService.to_csv(personnel_list)
            st.download_button(
                label="📥 Xuất CSV",
                data=csv_data,
                file_name=f"danh-sach-quan-nhan-{st.session_state.get('export_count', 0)}.csv",
                mime="text/csv"
            )
        
        # Bảng dữ liệu
        for idx, person in enumerate(personnel_list):
            with st.container():
                col1, col2, col3 = st.columns([4, 1, 1])
                
                with col1:
                    st.markdown(f"### {person.hoTen or 'Chưa có tên'}")
                    if person.capBac:
                        st.markdown(f"**Cấp bậc:** {person.capBac}")
                    if person.donVi:
                        st.markdown(f"**Đơn vị:** {person.donVi}")
                    if person.chucVu:
                        st.markdown(f"**Chức vụ:** {person.chucVu}")
                
                with col2:
                    if st.button("✏️ Sửa", key=f"edit_{person.id}"):
                        st.session_state['edit_personnel_id'] = person.id
                        st.rerun()
                
                with col3:
                    if st.button("🗑️ Xóa", key=f"delete_{person.id}"):
                        if db.delete(person.id):
                            st.success(f"Đã xóa {person.hoTen}")
                            st.rerun()
                
                st.markdown("---")
    else:
        st.info("Không tìm thấy quân nhân nào.")




"""

import streamlit as st
import sys
from pathlib import Path

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from services.export import ExportService


def show():
    """Hiển thị danh sách quân nhân"""
    st.title("📋 Danh Sách Quân Nhân")
    st.markdown("---")
    
    db = DatabaseService()
    
    # Tìm kiếm và lọc
    col1, col2 = st.columns([3, 1])
    
    with col1:
        search_query = st.text_input("🔍 Tìm kiếm theo tên", "")
    
    with col2:
        if st.button("➕ Thêm Mới"):
            st.session_state['edit_personnel_id'] = None
            st.rerun()
    
    # Filters
    with st.expander("🔽 Bộ Lọc", expanded=False):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            don_vi_filter = st.selectbox(
                "Đơn Vị",
                [""] + db.get_unique_values('donVi')
            )
        
        with col2:
            cap_bac_filter = st.selectbox(
                "Cấp Bậc",
                [""] + db.get_unique_values('capBac')
            )
        
        with col3:
            chuc_vu_filter = st.selectbox(
                "Chức Vụ",
                [""] + db.get_unique_values('chucVu')
            )
    
    # Tìm kiếm
    filters = {}
    if don_vi_filter:
        filters['donVi'] = don_vi_filter
    if cap_bac_filter:
        filters['capBac'] = cap_bac_filter
    if chuc_vu_filter:
        filters['chucVu'] = chuc_vu_filter
    
    personnel_list = db.search(search_query, filters if filters else None)
    
    # Hiển thị kết quả
    st.markdown(f"**Tìm thấy: {len(personnel_list)} quân nhân**")
    
    if personnel_list:
        # Nút xuất file
        col1, col2 = st.columns([1, 4])
        with col1:
            csv_data = ExportService.to_csv(personnel_list)
            st.download_button(
                label="📥 Xuất CSV",
                data=csv_data,
                file_name=f"danh-sach-quan-nhan-{st.session_state.get('export_count', 0)}.csv",
                mime="text/csv"
            )
        
        # Bảng dữ liệu
        for idx, person in enumerate(personnel_list):
            with st.container():
                col1, col2, col3 = st.columns([4, 1, 1])
                
                with col1:
                    st.markdown(f"### {person.hoTen or 'Chưa có tên'}")
                    if person.capBac:
                        st.markdown(f"**Cấp bậc:** {person.capBac}")
                    if person.donVi:
                        st.markdown(f"**Đơn vị:** {person.donVi}")
                    if person.chucVu:
                        st.markdown(f"**Chức vụ:** {person.chucVu}")
                
                with col2:
                    if st.button("✏️ Sửa", key=f"edit_{person.id}"):
                        st.session_state['edit_personnel_id'] = person.id
                        st.rerun()
                
                with col3:
                    if st.button("🗑️ Xóa", key=f"delete_{person.id}"):
                        if db.delete(person.id):
                            st.success(f"Đã xóa {person.hoTen}")
                            st.rerun()
                
                st.markdown("---")
    else:
        st.info("Không tìm thấy quân nhân nào.")