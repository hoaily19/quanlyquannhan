"""
Ứng dụng Quản Lý Hồ Sơ Quân Nhân - Streamlit
Entry point chính của ứng dụng
"""

import streamlit as st
import sys
from pathlib import Path

# Thêm thư mục hiện tại vào path
sys.path.insert(0, str(Path(__file__).parent))

from pages import home, personnel_list, personnel_detail, report, import_data

# Cấu hình trang
st.set_page_config(
    page_title="Quản Lý Quân Nhân",
    page_icon="🪖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Sidebar navigation
st.sidebar.title("🪖 Quản Lý Quân Nhân")
st.sidebar.markdown("---")

# Menu điều hướng
page = st.sidebar.selectbox(
    "Chọn chức năng",
    [
        "🏠 Trang Chủ",
        "📋 Danh Sách Quân Nhân",
        "➕ Thêm Quân Nhân",
        "📊 Báo Cáo Tổng Hợp",
        "📥 Nhập Dữ Liệu"
    ]
)

# Điều hướng đến trang tương ứng
if page == "🏠 Trang Chủ":
    home.show()
elif page == "📋 Danh Sách Quân Nhân":
    personnel_list.show()
elif page == "➕ Thêm Quân Nhân":
    personnel_detail.show(is_new=True)
elif page == "📊 Báo Cáo Tổng Hợp":
    report.show()
elif page == "📥 Nhập Dữ Liệu":
    import_data.show()
