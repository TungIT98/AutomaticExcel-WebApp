from excel_reader import ExcelReader

# Khởi tạo reader
reader = ExcelReader('0. Hoso.xlsx')

# Lấy context cho hồ sơ
context = reader.create_complete_context('2505382-CNSP.7/ECS')

print('=== DỮ LIỆU THỰC TỪ EXCEL ===')
print(f'so_ma_hoa: {context.get("so_ma_hoa", "N/A")}')
print(f'ten_kh: {context.get("ten_kh", "N/A")}')
print(f'Dia_chi_KH_DKCN: {context.get("Dia_chi_KH_DKCN", "N/A")}')
print(f'BM25_Time_cap_GCN: {context.get("BM25_Time_cap_GCN", "N/A")}')
print(f'Ma_so_thue_DN: {context.get("Ma_so_thue_DN", "N/A")}')
print(f'Nguoi_dai_dien: {context.get("Nguoi_dai_dien", "N/A")}')
print(f'Contact: {context.get("Contact", "N/A")}')
print(f'Invoice: {context.get("Invoice", "N/A")}')
print(f'TKHQ: {context.get("TKHQ", "N/A")}')
print(f'CĐĐG: {context.get("CĐĐG", "N/A")}')
print(f'CGKT: {context.get("CGKT", "N/A")}')
print(f'BM16_Ten_PTN: {context.get("BM16_Ten_PTN", "N/A")}')
print(f'BM16_ĐC_PTN: {context.get("BM16_ĐC_PTN", "N/A")}')
print(f'QD_Thanh_lap_doan_DG: {context.get("QD_Thanh_lap_doan_DG", "N/A")}')
print(f'BM24_so_QĐ_cap_GCN: {context.get("BM24_so_QĐ_cap_GCN", "N/A")}')
print(f'BM24_time_QĐ_cấp_GCN: {context.get("BM24_time_QĐ_cấp_GCN", "N/A")}')
print(f'BM25_Time_cap_GCN: {context.get("BM25_Time_cap_GCN", "N/A")}')
print(f'Tinh_trang_cap_GCN: {context.get("Tinh_trang_cap_GCN", "N/A")}')
print(f'Tinh_trang_ĐG: {context.get("Tinh_trang_ĐG", "N/A")}')

print('\n=== BẢNG HÀNG HÓA ===')
if 'tbl_hanghoa' in context:
    for i, item in enumerate(context['tbl_hanghoa']):
        print(f'Hàng hóa {i+1}:')
        print(f'  - Tên: {item.get(" hh.Ten_hang_hoa ", "N/A")}')
        print(f'  - Đặc tính: {item.get(" hh.Dac_tinh_KT ", "N/A")}')
        print(f'  - Khối lượng: {item.get(" hh.Khoi_luong ", "N/A")}')
        print(f'  - Xuất xứ: {item.get(" hh.Xuat_xu_Nha_SX ", "N/A")}')
        print()

