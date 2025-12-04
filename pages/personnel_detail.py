"""
Trang chi tiết/quản lý quân nhân
"""

import streamlit as st
import sys
from pathlib import Path

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from models.personnel import Personnel, ThongTinDang, ThongTinDoan, ThongTinKhac


def show(is_new: bool = False):
    """Hiển thị form thêm/sửa quân nhân"""
    db = DatabaseService()
    
    # Lấy ID từ session hoặc parameter
    personnel_id = st.session_state.get('edit_personnel_id') if not is_new else None
    
    if personnel_id:
        personnel = db.get_by_id(personnel_id)
        if not personnel:
            st.error("Không tìm thấy quân nhân")
            return
        title = "✏️ Chỉnh Sửa Quân Nhân"
    else:
        personnel = Personnel()
        title = "➕ Thêm Quân Nhân Mới"
    
    st.title(title)
    st.markdown("---")
    
    # Form
    with st.form("personnel_form"):
        # Thông tin cơ bản
        st.subheader("Thông Tin Cơ Bản")
        
        col1, col2 = st.columns(2)
        with col1:
            ho_ten = st.text_input("Họ và Tên *", value=personnel.hoTen, key="hoTen")
            ngay_sinh = st.text_input("Ngày Sinh (DD/MM/YYYY)", value=personnel.ngaySinh, key="ngaySinh")
            cap_bac = st.text_input("Cấp Bậc", value=personnel.capBac, key="capBac")
            chuc_vu = st.text_input("Chức Vụ", value=personnel.chucVu, key="chucVu")
            don_vi = st.text_input("Đơn Vị", value=personnel.donVi, key="donVi")
        
        with col2:
            nhap_ngu = st.text_input("Nhập Ngũ", value=personnel.nhapNgu, key="nhapNgu")
            que_quan = st.text_input("Quê Quán", value=personnel.queQuan, key="queQuan")
            tru_quan = st.text_input("Trú Quán", value=personnel.truQuan, key="truQuan")
            dan_toc = st.text_input("Dân Tộc", value=personnel.danToc, key="danToc")
            ton_giao = st.text_input("Tôn Giáo", value=personnel.tonGiao, key="tonGiao")
        
        trinh_do_van_hoa = st.text_input("Trình Độ Văn Hóa", value=personnel.trinhDoVanHoa, key="trinhDoVanHoa")
        
        # Thông tin đảng
        st.subheader("Thông Tin Đảng")
        col1, col2, col3 = st.columns(3)
        with col1:
            dang_ngay_vao = st.text_input("Ngày Vào Đảng", value=personnel.thongTinKhac.dang.ngayVao, key="dang_ngay_vao")
        with col2:
            dang_ngay_chinh_thuc = st.text_input("Ngày Chính Thức", value=personnel.thongTinKhac.dang.ngayChinhThuc, key="dang_ngay_chinh_thuc")
        with col3:
            dang_chuc_vu = st.text_input("Chức Vụ Đảng", value=personnel.thongTinKhac.dang.chucVuDang, key="dang_chuc_vu")
        
        # Thông tin đoàn
        st.subheader("Thông Tin Đoàn")
        col1, col2 = st.columns(2)
        with col1:
            doan_ngay_vao = st.text_input("Ngày Vào Đoàn", value=personnel.thongTinKhac.doan.ngayVao, key="doan_ngay_vao")
        with col2:
            doan_chuc_vu = st.text_input("Chức Vụ Đoàn", value=personnel.thongTinKhac.doan.chucVuDoan, key="doan_chuc_vu")
        
        # Thông tin khác
        st.subheader("Thông Tin Khác")
        col1, col2 = st.columns(2)
        with col1:
            cd_cu = st.checkbox("Có người thân tham gia chế độ cũ", value=personnel.thongTinKhac.cdCu, key="cdCu")
        with col2:
            yeu_to_nn = st.checkbox("Có yếu tố nước ngoài", value=personnel.thongTinKhac.yeuToNN, key="yeuToNN")
        
        # Nút submit
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            submitted = st.form_submit_button("💾 Lưu", type="primary")
        with col2:
            if st.form_submit_button("❌ Hủy"):
                if 'edit_personnel_id' in st.session_state:
                    del st.session_state['edit_personnel_id']
                st.rerun()
        
        if submitted:
            # Validate
            if not ho_ten.strip():
                st.error("Vui lòng nhập Họ và Tên")
                return
            
            # Cập nhật dữ liệu
            personnel.hoTen = ho_ten
            personnel.ngaySinh = ngay_sinh
            personnel.capBac = cap_bac
            personnel.chucVu = chuc_vu
            personnel.donVi = don_vi
            personnel.nhapNgu = nhap_ngu
            personnel.queQuan = que_quan
            personnel.truQuan = tru_quan
            personnel.danToc = dan_toc
            personnel.tonGiao = ton_giao
            personnel.trinhDoVanHoa = trinh_do_van_hoa
            
            personnel.thongTinKhac.dang.ngayVao = dang_ngay_vao
            personnel.thongTinKhac.dang.ngayChinhThuc = dang_ngay_chinh_thuc
            personnel.thongTinKhac.dang.chucVuDang = dang_chuc_vu
            
            personnel.thongTinKhac.doan.ngayVao = doan_ngay_vao
            personnel.thongTinKhac.doan.chucVuDoan = doan_chuc_vu
            
            personnel.thongTinKhac.cdCu = cd_cu
            personnel.thongTinKhac.yeuToNN = yeu_to_nn
            
            # Lưu
            if personnel_id:
                personnel.id = personnel_id
                if db.update(personnel):
                    st.success(f"Đã cập nhật {ho_ten}")
                    if 'edit_personnel_id' in st.session_state:
                        del st.session_state['edit_personnel_id']
                    st.rerun()
            else:
                db.create(personnel)
                st.success(f"Đã thêm {ho_ten}")
                st.rerun()




"""

import streamlit as st
import sys
from pathlib import Path

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from services.database import DatabaseService
from models.personnel import Personnel, ThongTinDang, ThongTinDoan, ThongTinKhac


def show(is_new: bool = False):
    """Hiển thị form thêm/sửa quân nhân"""
    db = DatabaseService()
    
    # Lấy ID từ session hoặc parameter
    personnel_id = st.session_state.get('edit_personnel_id') if not is_new else None
    
    if personnel_id:
        personnel = db.get_by_id(personnel_id)
        if not personnel:
            st.error("Không tìm thấy quân nhân")
            return
        title = "✏️ Chỉnh Sửa Quân Nhân"
    else:
        personnel = Personnel()
        title = "➕ Thêm Quân Nhân Mới"
    
    st.title(title)
    st.markdown("---")
    
    # Form
    with st.form("personnel_form"):
        # Thông tin cơ bản
        st.subheader("Thông Tin Cơ Bản")
        
        col1, col2 = st.columns(2)
        with col1:
            ho_ten = st.text_input("Họ và Tên *", value=personnel.hoTen, key="hoTen")
            ngay_sinh = st.text_input("Ngày Sinh (DD/MM/YYYY)", value=personnel.ngaySinh, key="ngaySinh")
            cap_bac = st.text_input("Cấp Bậc", value=personnel.capBac, key="capBac")
            chuc_vu = st.text_input("Chức Vụ", value=personnel.chucVu, key="chucVu")
            don_vi = st.text_input("Đơn Vị", value=personnel.donVi, key="donVi")
        
        with col2:
            nhap_ngu = st.text_input("Nhập Ngũ", value=personnel.nhapNgu, key="nhapNgu")
            que_quan = st.text_input("Quê Quán", value=personnel.queQuan, key="queQuan")
            tru_quan = st.text_input("Trú Quán", value=personnel.truQuan, key="truQuan")
            dan_toc = st.text_input("Dân Tộc", value=personnel.danToc, key="danToc")
            ton_giao = st.text_input("Tôn Giáo", value=personnel.tonGiao, key="tonGiao")
        
        trinh_do_van_hoa = st.text_input("Trình Độ Văn Hóa", value=personnel.trinhDoVanHoa, key="trinhDoVanHoa")
        
        # Thông tin đảng
        st.subheader("Thông Tin Đảng")
        col1, col2, col3 = st.columns(3)
        with col1:
            dang_ngay_vao = st.text_input("Ngày Vào Đảng", value=personnel.thongTinKhac.dang.ngayVao, key="dang_ngay_vao")
        with col2:
            dang_ngay_chinh_thuc = st.text_input("Ngày Chính Thức", value=personnel.thongTinKhac.dang.ngayChinhThuc, key="dang_ngay_chinh_thuc")
        with col3:
            dang_chuc_vu = st.text_input("Chức Vụ Đảng", value=personnel.thongTinKhac.dang.chucVuDang, key="dang_chuc_vu")
        
        # Thông tin đoàn
        st.subheader("Thông Tin Đoàn")
        col1, col2 = st.columns(2)
        with col1:
            doan_ngay_vao = st.text_input("Ngày Vào Đoàn", value=personnel.thongTinKhac.doan.ngayVao, key="doan_ngay_vao")
        with col2:
            doan_chuc_vu = st.text_input("Chức Vụ Đoàn", value=personnel.thongTinKhac.doan.chucVuDoan, key="doan_chuc_vu")
        
        # Thông tin khác
        st.subheader("Thông Tin Khác")
        col1, col2 = st.columns(2)
        with col1:
            cd_cu = st.checkbox("Có người thân tham gia chế độ cũ", value=personnel.thongTinKhac.cdCu, key="cdCu")
        with col2:
            yeu_to_nn = st.checkbox("Có yếu tố nước ngoài", value=personnel.thongTinKhac.yeuToNN, key="yeuToNN")
        
        # Nút submit
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            submitted = st.form_submit_button("💾 Lưu", type="primary")
        with col2:
            if st.form_submit_button("❌ Hủy"):
                if 'edit_personnel_id' in st.session_state:
                    del st.session_state['edit_personnel_id']
                st.rerun()
        
        if submitted:
            # Validate
            if not ho_ten.strip():
                st.error("Vui lòng nhập Họ và Tên")
                return
            
            # Cập nhật dữ liệu
            personnel.hoTen = ho_ten
            personnel.ngaySinh = ngay_sinh
            personnel.capBac = cap_bac
            personnel.chucVu = chuc_vu
            personnel.donVi = don_vi
            personnel.nhapNgu = nhap_ngu
            personnel.queQuan = que_quan
            personnel.truQuan = tru_quan
            personnel.danToc = dan_toc
            personnel.tonGiao = ton_giao
            personnel.trinhDoVanHoa = trinh_do_van_hoa
            
            personnel.thongTinKhac.dang.ngayVao = dang_ngay_vao
            personnel.thongTinKhac.dang.ngayChinhThuc = dang_ngay_chinh_thuc
            personnel.thongTinKhac.dang.chucVuDang = dang_chuc_vu
            
            personnel.thongTinKhac.doan.ngayVao = doan_ngay_vao
            personnel.thongTinKhac.doan.chucVuDoan = doan_chuc_vu
            
            personnel.thongTinKhac.cdCu = cd_cu
            personnel.thongTinKhac.yeuToNN = yeu_to_nn
            
            # Lưu
            if personnel_id:
                personnel.id = personnel_id
                if db.update(personnel):
                    st.success(f"Đã cập nhật {ho_ten}")
                    if 'edit_personnel_id' in st.session_state:
                        del st.session_state['edit_personnel_id']
                    st.rerun()
            else:
                db.create(personnel)
                st.success(f"Đã thêm {ho_ten}")
                st.rerun()