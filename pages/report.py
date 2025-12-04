"""
Trang báo cáo và thống kê
"""

import streamlit as st
import pandas as pd
import sys
from pathlib import Path

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from services.export import ExportService


def show():
    """Hiển thị báo cáo thống kê"""
    st.title("📊 Báo Cáo Tổng Hợp")
    st.markdown("---")
    
    db = DatabaseService()
    all_personnel = db.get_all()
    
    if not all_personnel:
        st.info("Chưa có dữ liệu để thống kê.")
        return
    
    # Tổng quan
    st.subheader("📈 Tổng Quan")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Tổng Số", len(all_personnel))
    
    with col2:
        dang_vien = sum(1 for p in all_personnel 
                       if p.thongTinKhac.dang.ngayVao or p.thongTinKhac.dang.ngayChinhThuc)
        st.metric("Đảng Viên", dang_vien)
    
    with col3:
        doan_vien = sum(1 for p in all_personnel if p.thongTinKhac.doan.ngayVao)
        st.metric("Đoàn Viên", doan_vien)
    
    with col4:
        co_cd_cu = sum(1 for p in all_personnel if p.thongTinKhac.cdCu)
        st.metric("Có Chế Độ Cũ", co_cd_cu)
    
    st.markdown("---")
    
    # Chọn tiêu chí thống kê
    criteria = st.selectbox(
        "Chọn Tiêu Chí Thống Kê",
        [
            "Dân Tộc",
            "Tôn Giáo",
            "Cấp Bậc",
            "Chức Vụ",
            "Đơn Vị",
            "Đảng Viên",
            "Đoàn Viên"
        ]
    )
    
    # Tính toán thống kê
    stats = {}
    
    if criteria == "Dân Tộc":
        for person in all_personnel:
            key = person.danToc or "Chưa xác định"
            stats[key] = stats.get(key, 0) + 1
    
    elif criteria == "Tôn Giáo":
        for person in all_personnel:
            key = person.tonGiao or "Không"
            stats[key] = stats.get(key, 0) + 1
    
    elif criteria == "Cấp Bậc":
        for person in all_personnel:
            key = person.capBac or "Chưa xác định"
            stats[key] = stats.get(key, 0) + 1
    
    elif criteria == "Chức Vụ":
        for person in all_personnel:
            key = person.chucVu or "Chưa xác định"
            stats[key] = stats.get(key, 0) + 1
    
    elif criteria == "Đơn Vị":
        for person in all_personnel:
            key = person.donVi or "Chưa xác định"
            stats[key] = stats.get(key, 0) + 1
    
    elif criteria == "Đảng Viên":
        dang_vien_count = sum(1 for p in all_personnel 
                             if p.thongTinKhac.dang.ngayVao or p.thongTinKhac.dang.ngayChinhThuc)
        stats["Đảng viên"] = dang_vien_count
        stats["Không phải đảng viên"] = len(all_personnel) - dang_vien_count
    
    elif criteria == "Đoàn Viên":
        doan_vien_count = sum(1 for p in all_personnel if p.thongTinKhac.doan.ngayVao)
        stats["Đoàn viên"] = doan_vien_count
        stats["Không phải đoàn viên"] = len(all_personnel) - doan_vien_count
    
    # Hiển thị biểu đồ
    if stats:
        df = pd.DataFrame({
            'Tiêu Chí': list(stats.keys()),
            'Số Lượng': list(stats.values())
        })
        df = df.sort_values('Số Lượng', ascending=False)
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.bar_chart(df.set_index('Tiêu Chí'))
        
        with col2:
            st.dataframe(df, use_container_width=True)
            
            # Xuất CSV
            csv_data = ExportService.to_csv(all_personnel)
            st.download_button(
                label="📥 Xuất CSV",
                data=csv_data,
                file_name=f"thong-ke-{criteria.lower().replace(' ', '-')}.csv",
                mime="text/csv"
            )




"""

import streamlit as st
import pandas as pd
import sys
from pathlib import Path

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from services.export import ExportService


def show():
    """Hiển thị báo cáo thống kê"""
    st.title("📊 Báo Cáo Tổng Hợp")
    st.markdown("---")
    
    db = DatabaseService()
    all_personnel = db.get_all()
    
    if not all_personnel:
        st.info("Chưa có dữ liệu để thống kê.")
        return
    
    # Tổng quan
    st.subheader("📈 Tổng Quan")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Tổng Số", len(all_personnel))
    
    with col2:
        dang_vien = sum(1 for p in all_personnel 
                       if p.thongTinKhac.dang.ngayVao or p.thongTinKhac.dang.ngayChinhThuc)
        st.metric("Đảng Viên", dang_vien)
    
    with col3:
        doan_vien = sum(1 for p in all_personnel if p.thongTinKhac.doan.ngayVao)
        st.metric("Đoàn Viên", doan_vien)
    
    with col4:
        co_cd_cu = sum(1 for p in all_personnel if p.thongTinKhac.cdCu)
        st.metric("Có Chế Độ Cũ", co_cd_cu)
    
    st.markdown("---")
    
    # Chọn tiêu chí thống kê
    criteria = st.selectbox(
        "Chọn Tiêu Chí Thống Kê",
        [
            "Dân Tộc",
            "Tôn Giáo",
            "Cấp Bậc",
            "Chức Vụ",
            "Đơn Vị",
            "Đảng Viên",
            "Đoàn Viên"
        ]
    )
    
    # Tính toán thống kê
    stats = {}
    
    if criteria == "Dân Tộc":
        for person in all_personnel:
            key = person.danToc or "Chưa xác định"
            stats[key] = stats.get(key, 0) + 1
    
    elif criteria == "Tôn Giáo":
        for person in all_personnel:
            key = person.tonGiao or "Không"
            stats[key] = stats.get(key, 0) + 1
    
    elif criteria == "Cấp Bậc":
        for person in all_personnel:
            key = person.capBac or "Chưa xác định"
            stats[key] = stats.get(key, 0) + 1
    
    elif criteria == "Chức Vụ":
        for person in all_personnel:
            key = person.chucVu or "Chưa xác định"
            stats[key] = stats.get(key, 0) + 1
    
    elif criteria == "Đơn Vị":
        for person in all_personnel:
            key = person.donVi or "Chưa xác định"
            stats[key] = stats.get(key, 0) + 1
    
    elif criteria == "Đảng Viên":
        dang_vien_count = sum(1 for p in all_personnel 
                             if p.thongTinKhac.dang.ngayVao or p.thongTinKhac.dang.ngayChinhThuc)
        stats["Đảng viên"] = dang_vien_count
        stats["Không phải đảng viên"] = len(all_personnel) - dang_vien_count
    
    elif criteria == "Đoàn Viên":
        doan_vien_count = sum(1 for p in all_personnel if p.thongTinKhac.doan.ngayVao)
        stats["Đoàn viên"] = doan_vien_count
        stats["Không phải đoàn viên"] = len(all_personnel) - doan_vien_count
    
    # Hiển thị biểu đồ
    if stats:
        df = pd.DataFrame({
            'Tiêu Chí': list(stats.keys()),
            'Số Lượng': list(stats.values())
        })
        df = df.sort_values('Số Lượng', ascending=False)
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.bar_chart(df.set_index('Tiêu Chí'))
        
        with col2:
            st.dataframe(df, use_container_width=True)
            
            # Xuất CSV
            csv_data = ExportService.to_csv(all_personnel)
            st.download_button(
                label="📥 Xuất CSV",
                data=csv_data,
                file_name=f"thong-ke-{criteria.lower().replace(' ', '-')}.csv",
                mime="text/csv"
            )