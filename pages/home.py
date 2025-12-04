"""
Trang chủ
"""

import streamlit as st


def show():
    """Hiển thị trang chủ"""
    st.title("Quản Lý Hồ Sơ Quân Nhân")
    st.markdown("---")
    
    st.markdown("""
    ### Chào mừng đến với hệ thống Quản Lý Hồ Sơ Quân Nhân
    
    Ứng dụng giúp bạn:
    - ✅ Quản lý hồ sơ quân nhân (Thêm, Sửa, Xóa, Xem)
    - ✅ Tìm kiếm và lọc hồ sơ
    - ✅ Xem thống kê và báo cáo tổng hợp
    - ✅ Xuất dữ liệu ra CSV/PDF
    - ✅ Nhập dữ liệu từ file Word/Excel
    
    **Sử dụng menu bên trái để điều hướng.**
    """)
    
    # Thống kê nhanh
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from services.database import DatabaseService
    
    db = DatabaseService()
    all_personnel = db.get_all()
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Tổng Số Quân Nhân", len(all_personnel))
    
    with col2:
        unique_units = len(db.get_unique_values('donVi'))
        st.metric("Số Đơn Vị", unique_units)
    
    with col3:
        unique_ranks = len(db.get_unique_values('capBac'))
        st.metric("Số Cấp Bậc", unique_ranks)
    
    with col4:
        # Đếm đảng viên
        dang_vien = sum(1 for p in all_personnel 
                       if p.thongTinKhac.dang.ngayVao or p.thongTinKhac.dang.ngayChinhThuc)
        st.metric("Đảng Viên", dang_vien)




"""

import streamlit as st


def show():
    """Hiển thị trang chủ"""
    st.title("Quản Lý Hồ Sơ Quân Nhân")
    st.markdown("---")
    
    st.markdown("""
    ### Chào mừng đến với hệ thống Quản Lý Hồ Sơ Quân Nhân
    
    Ứng dụng giúp bạn:
    - ✅ Quản lý hồ sơ quân nhân (Thêm, Sửa, Xóa, Xem)
    - ✅ Tìm kiếm và lọc hồ sơ
    - ✅ Xem thống kê và báo cáo tổng hợp
    - ✅ Xuất dữ liệu ra CSV/PDF
    - ✅ Nhập dữ liệu từ file Word/Excel
    
    **Sử dụng menu bên trái để điều hướng.**
    """)
    
    # Thống kê nhanh
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from services.database import DatabaseService
    
    db = DatabaseService()
    all_personnel = db.get_all()
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Tổng Số Quân Nhân", len(all_personnel))
    
    with col2:
        unique_units = len(db.get_unique_values('donVi'))
        st.metric("Số Đơn Vị", unique_units)
    
    with col3:
        unique_ranks = len(db.get_unique_values('capBac'))
        st.metric("Số Cấp Bậc", unique_ranks)
    
    with col4:
        # Đếm đảng viên
        dang_vien = sum(1 for p in all_personnel 
                       if p.thongTinKhac.dang.ngayVao or p.thongTinKhac.dang.ngayChinhThuc)
        st.metric("Đảng Viên", dang_vien)