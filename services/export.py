"""
Service xuất dữ liệu ra CSV và PDF
"""

import csv
import io
import sys
from typing import List
from datetime import datetime
from pathlib import Path

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from models.personnel import Personnel


class ExportService:
    """Service xuất file"""
    
    @staticmethod
    def to_csv(personnel_list: List[Personnel]) -> str:
        """
        Chuyển đổi danh sách quân nhân sang CSV
        Returns:
            Chuỗi CSV
        """
        if not personnel_list:
            return ""
        
        output = io.StringIO()
        writer = csv.writer(output)
        
        # Header - Đầy đủ các trường với chính tả đúng
        headers = [
            'Họ và Tên', 'Họ Tên Thường Dùng', 'Ngày Sinh', 'Cấp Bậc', 
            'Ngày Nhận Cấp Bậc', 'Chức Vụ', 'Ngày Nhận Chức Vụ', 'Đơn Vị',
            'Nhập Ngũ', 'Xuất Ngũ', 'Quê Quán', 'Nơi Cư Trú', 'Dân Tộc', 'Tôn Giáo',
            'Trình Độ Văn Hóa', 'Thành Phần Gia Đình', 'Qua Trường', 'Ngành Học',
            'Cấp Học', 'Thời Gian Đào Tạo', 'Liên Hệ Khi Cần', 'Số Điện Thoại Liên Hệ',
            'Họ Tên Cha', 'Họ Tên Mẹ', 'Họ Tên Vợ', 'Ghi Chú',
            'Ngày Vào Đảng', 'Ngày Chính Thức Đảng', 'Chức Vụ Đảng',
            'Ngày Vào Đoàn', 'Chức Vụ Đoàn',
            'Chế Độ Cũ', 'Yếu Tố Nước Ngoài'
        ]
        writer.writerow(headers)
        
        # Dữ liệu - Đầy đủ các trường
        for person in personnel_list:
            row = [
                person.hoTen or '',
                person.hoTenThuongDung or '',
                person.ngaySinh or '',
                person.capBac or '',
                person.ngayNhanCapBac or '',
                person.chucVu or '',
                person.ngayNhanChucVu or '',
                person.donVi or '',
                person.nhapNgu or '',
                person.xuatNgu or '',
                person.queQuan or '',
                person.truQuan or '',
                person.danToc or '',
                person.tonGiao or '',
                person.trinhDoVanHoa or '',
                person.thanhPhanGiaDinh or '',
                person.quaTruong or '',
                person.nganhHoc or '',
                person.capHoc or '',
                person.thoiGianDaoTao or '',
                person.lienHeKhiCan or '',
                person.soDienThoaiLienHe or '',
                person.hoTenCha or '',
                person.hoTenMe or '',
                person.hoTenVo or '',
                person.ghiChu or '',
                person.thongTinKhac.dang.ngayVao or '',
                person.thongTinKhac.dang.ngayChinhThuc or '',
                person.thongTinKhac.dang.chucVuDang or '',
                person.thongTinKhac.doan.ngayVao or '',
                person.thongTinKhac.doan.chucVuDoan or '',
                'Có' if person.thongTinKhac.cdCu else 'Không',
                'Có' if person.thongTinKhac.yeuToNN else 'Không',
            ]
            writer.writerow(row)
        
        return output.getvalue()
    
    @staticmethod
    def to_pdf(personnel_list: List[Personnel], title: str = "Danh Sách Quân Nhân") -> bytes:
        """
        Chuyển đổi danh sách quân nhân sang PDF
        Returns:
            Bytes của file PDF
        """
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import cm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib import colors
            
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            story = []
            styles = getSampleStyleSheet()
            
            # Tiêu đề
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                textColor=colors.HexColor('#4CAF50'),
                spaceAfter=30,
                alignment=1  # Center
            )
            story.append(Paragraph(title, title_style))
            story.append(Paragraph(f"Tổng số: {len(personnel_list)} quân nhân", styles['Normal']))
            story.append(Spacer(1, 0.5*cm))
            
            # Bảng dữ liệu
            if personnel_list:
                data = [['STT', 'Họ và Tên', 'Cấp Bậc', 'Chức Vụ', 'Đơn Vị', 'Dân Tộc']]
                
                for idx, person in enumerate(personnel_list, 1):
                    data.append([
                        str(idx),
                        person.hoTen or '',
                        person.capBac or '',
                        person.chucVu or '',
                        person.donVi or '',
                        person.danToc or '',
                    ])
                
                table = Table(data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4CAF50')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                ]))
                
                story.append(table)
            
            # Footer
            story.append(Spacer(1, 1*cm))
            story.append(Paragraph(
                f"Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                styles['Normal']
            ))
            
            doc.build(story)
            buffer.seek(0)
            return buffer.getvalue()
            
        except ImportError:
            # Fallback nếu không có reportlab
            raise ImportError("Cần cài đặt reportlab: pip install reportlab")

import csv
import io
import sys
from typing import List
from datetime import datetime
from pathlib import Path

# Thêm thư mục gốc vào path
sys.path.insert(0, str(Path(__file__).parent.parent))

from models.personnel import Personnel


class ExportService:
    """Service xuất file"""
    
    @staticmethod
    def to_csv(personnel_list: List[Personnel]) -> str:
        """
        Chuyển đổi danh sách quân nhân sang CSV
        Returns:
            Chuỗi CSV
        """
        if not personnel_list:
            return ""
        
        output = io.StringIO()
        writer = csv.writer(output)
        
        # Header - Đầy đủ các trường với chính tả đúng
        headers = [
            'Họ và Tên', 'Họ Tên Thường Dùng', 'Ngày Sinh', 'Cấp Bậc', 
            'Ngày Nhận Cấp Bậc', 'Chức Vụ', 'Ngày Nhận Chức Vụ', 'Đơn Vị',
            'Nhập Ngũ', 'Xuất Ngũ', 'Quê Quán', 'Nơi Cư Trú', 'Dân Tộc', 'Tôn Giáo',
            'Trình Độ Văn Hóa', 'Thành Phần Gia Đình', 'Qua Trường', 'Ngành Học',
            'Cấp Học', 'Thời Gian Đào Tạo', 'Liên Hệ Khi Cần', 'Số Điện Thoại Liên Hệ',
            'Họ Tên Cha', 'Họ Tên Mẹ', 'Họ Tên Vợ', 'Ghi Chú',
            'Ngày Vào Đảng', 'Ngày Chính Thức Đảng', 'Chức Vụ Đảng',
            'Ngày Vào Đoàn', 'Chức Vụ Đoàn',
            'Chế Độ Cũ', 'Yếu Tố Nước Ngoài'
        ]
        writer.writerow(headers)
        
        # Dữ liệu - Đầy đủ các trường
        for person in personnel_list:
            row = [
                person.hoTen or '',
                person.hoTenThuongDung or '',
                person.ngaySinh or '',
                person.capBac or '',
                person.ngayNhanCapBac or '',
                person.chucVu or '',
                person.ngayNhanChucVu or '',
                person.donVi or '',
                person.nhapNgu or '',
                person.xuatNgu or '',
                person.queQuan or '',
                person.truQuan or '',
                person.danToc or '',
                person.tonGiao or '',
                person.trinhDoVanHoa or '',
                person.thanhPhanGiaDinh or '',
                person.quaTruong or '',
                person.nganhHoc or '',
                person.capHoc or '',
                person.thoiGianDaoTao or '',
                person.lienHeKhiCan or '',
                person.soDienThoaiLienHe or '',
                person.hoTenCha or '',
                person.hoTenMe or '',
                person.hoTenVo or '',
                person.ghiChu or '',
                person.thongTinKhac.dang.ngayVao or '',
                person.thongTinKhac.dang.ngayChinhThuc or '',
                person.thongTinKhac.dang.chucVuDang or '',
                person.thongTinKhac.doan.ngayVao or '',
                person.thongTinKhac.doan.chucVuDoan or '',
                'Có' if person.thongTinKhac.cdCu else 'Không',
                'Có' if person.thongTinKhac.yeuToNN else 'Không',
            ]
            writer.writerow(row)
        
        return output.getvalue()
    
    @staticmethod
    def to_pdf(personnel_list: List[Personnel], title: str = "Danh Sách Quân Nhân") -> bytes:
        """
        Chuyển đổi danh sách quân nhân sang PDF
        Returns:
            Bytes của file PDF
        """
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import cm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib import colors
            
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            story = []
            styles = getSampleStyleSheet()
            
            # Tiêu đề
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                textColor=colors.HexColor('#4CAF50'),
                spaceAfter=30,
                alignment=1  # Center
            )
            story.append(Paragraph(title, title_style))
            story.append(Paragraph(f"Tổng số: {len(personnel_list)} quân nhân", styles['Normal']))
            story.append(Spacer(1, 0.5*cm))
            
            # Bảng dữ liệu
            if personnel_list:
                data = [['STT', 'Họ và Tên', 'Cấp Bậc', 'Chức Vụ', 'Đơn Vị', 'Dân Tộc']]
                
                for idx, person in enumerate(personnel_list, 1):
                    data.append([
                        str(idx),
                        person.hoTen or '',
                        person.capBac or '',
                        person.chucVu or '',
                        person.donVi or '',
                        person.danToc or '',
                    ])
                
                table = Table(data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4CAF50')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                ]))
                
                story.append(table)
            
            # Footer
            story.append(Spacer(1, 1*cm))
            story.append(Paragraph(
                f"Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                styles['Normal']
            ))
            
            doc.build(story)
            buffer.seek(0)
            return buffer.getvalue()
            
        except ImportError:
            # Fallback nếu không có reportlab
            raise ImportError("Cần cài đặt reportlab: pip install reportlab")