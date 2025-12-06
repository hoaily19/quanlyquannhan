"""
Script đọc và xem dữ liệu từ SQLite database
"""

import sqlite3
import json
from pathlib import Path

# Đường dẫn database
db_path = "data/personnel.db"

def read_database():
    """Đọc và hiển thị dữ liệu từ database"""
    
    if not Path(db_path).exists():
        print(f"❌ Không tìm thấy file: {db_path}")
        return
    
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row  # Trả về dạng dictionary
        cursor = conn.cursor()
        
        print("=" * 60)
        print("📊 DỮ LIỆU TRONG DATABASE")
        print("=" * 60)
        
        # 1. Xem các bảng
        print("\n📋 CÁC BẢNG TRONG DATABASE:")
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = cursor.fetchall()
        for table in tables:
            print(f"  • {table['name']}")
        
        # 2. Đếm số lượng quân nhân
        print("\n👥 THỐNG KÊ QUÂN NHÂN:")
        cursor.execute("SELECT COUNT(*) as total FROM personnel")
        total = cursor.fetchone()['total']
        print(f"  Tổng số quân nhân: {total}")
        
        # 3. Danh sách dân tộc
        print("\n🌍 DANH SÁCH DÂN TỘC:")
        cursor.execute("SELECT DISTINCT danToc, COUNT(*) as count FROM personnel WHERE danToc IS NOT NULL AND danToc != '' GROUP BY danToc ORDER BY count DESC")
        ethnic_groups = cursor.fetchall()
        for ethnic in ethnic_groups:
            print(f"  • {ethnic['danToc']}: {ethnic['count']} người")
        
        # 4. Danh sách đơn vị
        print("\n🏛️ DANH SÁCH ĐƠN VỊ:")
        cursor.execute("SELECT DISTINCT donVi, COUNT(*) as count FROM personnel WHERE donVi IS NOT NULL AND donVi != '' GROUP BY donVi ORDER BY count DESC")
        units = cursor.fetchall()
        for unit in units:
            print(f"  • {unit['donVi']}: {unit['count']} người")
        
        # 5. Xem một vài quân nhân mẫu
        print("\n📝 MẪU DỮ LIỆU QUÂN NHÂN (5 người đầu tiên):")
        cursor.execute("SELECT id, hoTen, capBac, chucVu, donVi, danToc FROM personnel LIMIT 5")
        personnel = cursor.fetchall()
        
        for idx, person in enumerate(personnel, 1):
            print(f"\n  {idx}. {person['hoTen'] or 'N/A'}")
            print(f"     - Cấp bậc: {person['capBac'] or 'N/A'}")
            print(f"     - Chức vụ: {person['chucVu'] or 'N/A'}")
            print(f"     - Đơn vị: {person['donVi'] or 'N/A'}")
            print(f"     - Dân tộc: {person['danToc'] or 'N/A'}")
        
        # 6. Xem cấu trúc bảng personnel
        print("\n🔧 CẤU TRÚC BẢNG PERSONNEL:")
        cursor.execute("PRAGMA table_info(personnel)")
        columns = cursor.fetchall()
        print("  Các cột:")
        for col in columns:
            print(f"    • {col['name']} ({col['type']})")
        
        # 7. Xem cấu trúc bảng units (nếu có)
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='units'")
        if cursor.fetchone():
            print("\n🔧 CẤU TRÚC BẢNG UNITS:")
            cursor.execute("PRAGMA table_info(units)")
            columns = cursor.fetchall()
            print("  Các cột:")
            for col in columns:
                print(f"    • {col['name']} ({col['type']})")
            
            # Đếm số đơn vị
            cursor.execute("SELECT COUNT(*) as total FROM units")
            total_units = cursor.fetchone()['total']
            print(f"\n  Tổng số đơn vị: {total_units}")
        
        conn.close()
        print("\n" + "=" * 60)
        print("✅ Đọc database thành công!")
        
    except Exception as e:
        print(f"❌ Lỗi khi đọc database: {str(e)}")

def export_to_json():
    """Xuất toàn bộ dữ liệu ra file JSON"""
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # Lấy tất cả quân nhân
        cursor.execute("SELECT * FROM personnel")
        personnel = cursor.fetchall()
        
        # Chuyển đổi sang dictionary
        data = []
        for person in personnel:
            data.append(dict(person))
        
        # Lưu ra file JSON
        output_file = "data/personnel_export.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=str)
        
        print(f"\n💾 Đã xuất {len(data)} quân nhân ra file: {output_file}")
        
        conn.close()
        
    except Exception as e:
        print(f"❌ Lỗi khi xuất JSON: {str(e)}")

if __name__ == "__main__":
    print("🔍 Đang đọc database...\n")
    read_database()
    
    # Hỏi có muốn xuất ra JSON không
    print("\n" + "=" * 60)
    choice = input("Bạn có muốn xuất dữ liệu ra file JSON? (y/n): ")
    if choice.lower() == 'y':
        export_to_json()


