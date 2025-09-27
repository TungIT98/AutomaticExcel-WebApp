import pandas as pd
import logging
from pathlib import Path
from datetime import datetime
import numpy as np

class ExcelReader:
    """Lớp chuyên dụng để đọc file Excel và tạo context phù hợp cho template Word"""
    
    def __init__(self, excel_file_path):
        self.excel_file_path = Path(excel_file_path)
        self.hoso_df = None
        self.hanghoa_df = None
        self.load_data()
    
    def load_data(self):
        """Đọc dữ liệu từ file Excel"""
        try:
            logging.info(f"Đang đọc dữ liệu từ: {self.excel_file_path}")
            excel_data = pd.read_excel(self.excel_file_path, sheet_name=['hoso', 'hanghoa'])
            
            self.hoso_df = excel_data['hoso']
            self.hanghoa_df = excel_data['hanghoa']
            
            # Xử lý dữ liệu ngày tháng
            self._process_dates()
            
            logging.info(f"Đọc thành công {len(self.hoso_df)} hồ sơ và {len(self.hanghoa_df)} hàng hóa")
            
        except Exception as e:
            logging.error(f"Lỗi khi đọc file Excel: {e}")
            raise
    
    def _process_dates(self):
        """Xử lý các cột ngày tháng trong DataFrame"""
        date_columns = [
            'Ngay_dang_ky', 'Ngay_ ĐG', 'Ngay_dang_ky.1', 'Time_contact', 
            'Time_invoice', 'Time_packinglist', 'Time_bill', 'Time_TKHQ',
            'Time_XX_HSCN', 'Time_TL_doan_DG', 'BM09_Time_lap_KHDG',
            'Time _ĐG_du_kien', 'BM10_Time_xac_nhan_KHDG', 'BM14_Time_lap_BB_KT_lay_mau',
            'BM16_time_gui_mau_TN', 'Time_cap_KQTN', 'BM17_Time_bao_cao_DG_MTN',
            'BM20_Time_BC_KQTS_HSCN', 'BM24_time_QĐ_cấp_GCN', 'BM25_Time_cap_GCN'
        ]
        
        for col in date_columns:
            if col in self.hoso_df.columns:
                self.hoso_df[col] = pd.to_datetime(self.hoso_df[col], errors='coerce')
    
    def get_hoso_data(self, ma_ho_so=None):
        """Lấy dữ liệu hồ sơ theo mã hồ sơ hoặc tất cả"""
        if ma_ho_so:
            return self.hoso_df[self.hoso_df['so_ma_hoa'] == ma_ho_so]
        return self.hoso_df
    
    def get_hanghoa_data(self, ma_ho_so=None):
        """Lấy dữ liệu hàng hóa theo mã hồ sơ hoặc tất cả"""
        if ma_ho_so:
            return self.hanghoa_df[self.hanghoa_df['so_ma_hoa'] == ma_ho_so]
        return self.hanghoa_df
    
    def create_base_context(self, hoso_row):
        """Tạo context cơ sở từ dữ liệu hồ sơ"""
        context = {}
        
        # Thông tin cơ bản
        context.update({
            'so_ma_hoa': str(hoso_row.get('so_ma_hoa', '')),
            'ten_kh': str(hoso_row.get('ten_kh', '')),
            'Dia_chi_KH_DKCN': str(hoso_row.get('Dia_chi_KH_DKCN', '')),
            'BM25_Time_cap_GCN': self._format_date(hoso_row.get('BM25_Time_cap_GCN')),
        })
        
        # Thông tin doanh nghiệp
        context.update({
            'Ma_so_thue_DN': str(hoso_row.get('Ma_so_thue_DN', '')),
            'Nguoi_dai_dien': str(hoso_row.get('Nguoi_dai_dien', '')),
            'Chuc_vu': str(hoso_row.get('Chuc_vu', '')),
            'Sdt': str(hoso_row.get('Sdt', '')),
            'Nguoi_tiep_nhan': str(hoso_row.get('Nguoi_tiep_nhan', '')),
        })
        
        # Thông tin địa lý
        context.update({
            'Tinh_DKKD': str(hoso_row.get('Tinh_DKKD', '')),
            'Tinh_lay_mau': str(hoso_row.get('Tinh_lay_mau', '')),
            'Cua_nhap_khau': str(hoso_row.get('Cua_nhap_khau', '')),
        })
        
        # Thông tin liên hệ và tài liệu
        context.update({
            'Contact': str(hoso_row.get('Contact', '')),
            'Time_contact': self._format_date(hoso_row.get('Time_contact')),
            'Invoice': str(hoso_row.get('Invoice', '')),
            'Time_invoice': self._format_date(hoso_row.get('Time_invoice')),
            'Packinglist': str(hoso_row.get('Packinglist', '')),
            'Time_packinglist': self._format_date(hoso_row.get('Time_packinglist')),
            'Bill_off_loading': str(hoso_row.get('Bill_off_loading', '')),
            'Time_bill': self._format_date(hoso_row.get('Time_bill')),
            'TKHQ': str(hoso_row.get('TKHQ', '')),
            'Time_TKHQ': self._format_date(hoso_row.get('Time_TKHQ')),
        })
        
        # Thông tin đánh giá
        context.update({
            'CĐĐG': str(hoso_row.get('CĐĐG', '')),
            'CGĐG_TS': str(hoso_row.get('CGĐG_TS', '')),
            'CGKT': str(hoso_row.get('CGKT', '')),
            'Nguoi_lay_mau': str(hoso_row.get('Nguoi_lay_mau', '')),
            'BPKT': str(hoso_row.get('BPKT', '')),
            'Ban_tham_xet': str(hoso_row.get('Ban_tham_xet', '')),
            'xx_hscn': str(hoso_row.get('xx_hscn', '')),
        })
        
        # Thông tin quyết định và kế hoạch
        context.update({
            'QD_Thanh_lap_doan_DG': str(hoso_row.get('QD_Thanh_lap_doan_DG', '')),
            'Time_TL_doan_DG': self._format_date(hoso_row.get('Time_TL_doan_DG')),
            'BM09_Time_lap_KHDG': self._format_date(hoso_row.get('BM09_Time_lap_KHDG')),
            'Time_ĐG_du_kien': self._format_date(hoso_row.get('Time _ĐG_du_kien')),
            'BM10_Time_xac_nhan_KHDG': self._format_date(hoso_row.get('BM10_Time_xac_nhan_KHDG')),
            'Ngay_ĐG': self._format_date(hoso_row.get('Ngay_ ĐG')),
            'Ngay_ ĐG': self._format_date(hoso_row.get('Ngay_ ĐG')),  # Giữ nguyên để tương thích với template
        })
        
        # Thông tin phòng thí nghiệm
        context.update({
            'BM16_Ten_PTN': str(hoso_row.get('BM16_Ten_PTN', '')),
            'BM16_ĐC_PTN': str(hoso_row.get('BM16_ĐC_PTN', '')),
            'BM16_nguoi_giao_mau': str(hoso_row.get('BM16_nguoi_giao_mau', '')),
            'BM16_nguoi_nhan_mau': str(hoso_row.get('BM16_nguoi_nhan_mau', '')),
            'BM16_time_gui_mau_TN': self._format_date(hoso_row.get('BM16_time_gui_mau_TN')),
            'Time_cap_KQTN': self._format_date(hoso_row.get('Time_cap_KQTN')),
        })
        
        # Thông tin báo cáo và quyết định
        context.update({
            'BM17_Time_bao_cao_DG_MTN': self._format_date(hoso_row.get('BM17_Time_bao_cao_DG_MTN')),
            'BM20_Time_BC_KQTS_HSCN': self._format_date(hoso_row.get('BM20_Time_BC_KQTS_HSCN')),
            'BM24_so_QĐ_cap_GCN': str(hoso_row.get('BM24_so_QĐ_cap_GCN', '')),
            'BM24_time_QĐ_cấp_GCN': self._format_date(hoso_row.get('BM24_time_QĐ_cấp_GCN')),
        })
        
        # Thông tin trạng thái
        context.update({
            'Tinh_trang_cap_GCN': str(hoso_row.get('Tinh_trang_cap_GCN', '')),
            'Tinh_trang_ĐG': str(hoso_row.get('Tinh_trang_ĐG', '')),
            'Ngay_ ĐG': self._format_date(hoso_row.get('Ngay_ ĐG')),
        })
        
        # Thông tin số phiếu và tick
        context.update({
            'SPH_HĐ': str(hoso_row.get('SPH_HĐ', '')),
            'Tick_HĐ': str(hoso_row.get('Tick_HĐ', '')),
            'SPH_invoice': str(hoso_row.get('SPH_invoice', '')),
            'Tick_invoice': str(hoso_row.get('Tick_invoice', '')),
            'SPH_PKL': str(hoso_row.get('SPH_PKL', '')),
            'Tick_PKL': str(hoso_row.get('Tick_PKL', '')),
            'SPH_bill': str(hoso_row.get('SPH_bill', '')),
            'Tick_bill': str(hoso_row.get('Tick_bill', '')),
            'SPH_TKHQ': str(hoso_row.get('SPH_TKHQ', '')),
            'Tick_TKHQ': str(hoso_row.get('Tick_TKHQ', '')),
        })
        
        return context
    
    def create_hanghoa_context(self, ma_ho_so):
        """Tạo context cho bảng hàng hóa"""
        hanghoa_data = self.get_hanghoa_data(ma_ho_so)
        
        if hanghoa_data.empty:
            return {'tbl_hanghoa': []}
        
        tbl_hanghoa = []
        for idx, (_, row) in enumerate(hanghoa_data.iterrows()):
            tbl_hanghoa.append({
                ' hh.STT ': str(idx + 1),
                ' hh.Ten_hang_hoa ': str(row.get('ten_hang_hoa', '')),
                ' hh.Dac_tinh_KT ': str(row.get('dac_tinh_kt', '')),
                ' hh.Khoi_luong ': str(row.get('khoi_luong', '')),
                ' hh.Xuat_xu_Nha_SX ': str(row.get('xuatxu_nsx', '')),
            })
        
        return {'tbl_hanghoa': tbl_hanghoa}
    
    def create_complete_context(self, ma_ho_so):
        """Tạo context hoàn chỉnh cho một hồ sơ"""
        hoso_data = self.get_hoso_data(ma_ho_so)
        
        if hoso_data.empty:
            logging.warning(f"Không tìm thấy hồ sơ với mã: {ma_ho_so}")
            return {}
        
        hoso_row = hoso_data.iloc[0]
        
        # Tạo context cơ sở
        context = self.create_base_context(hoso_row)
        
        # Thêm context hàng hóa
        hanghoa_context = self.create_hanghoa_context(ma_ho_so)
        context.update(hanghoa_context)
        
        return context
    
    def _format_date(self, date_value):
        """Format ngày tháng thành chuỗi"""
        if pd.isna(date_value) or date_value is None:
            return ""
        
        if isinstance(date_value, (pd.Timestamp, datetime)):
            return date_value.strftime('%d/%m/%Y')
        
        return str(date_value)
    
    def get_all_hoso_codes(self):
        """Lấy danh sách tất cả mã hồ sơ"""
        return self.hoso_df['so_ma_hoa'].unique().tolist()
    
    def validate_data(self):
        """Kiểm tra tính hợp lệ của dữ liệu"""
        issues = []
        
        # Kiểm tra dữ liệu hồ sơ
        if self.hoso_df.empty:
            issues.append("Không có dữ liệu hồ sơ")
        
        # Kiểm tra dữ liệu hàng hóa
        if self.hanghoa_df.empty:
            issues.append("Không có dữ liệu hàng hóa")
        
        # Kiểm tra mã hồ sơ trùng lặp
        duplicate_codes = self.hoso_df['so_ma_hoa'].duplicated().sum()
        if duplicate_codes > 0:
            issues.append(f"Có {duplicate_codes} mã hồ sơ trùng lặp")
        
        return issues

def main():
    """Hàm test để kiểm tra ExcelReader"""
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    try:
        # Khởi tạo reader
        reader = ExcelReader("0. Hoso.xlsx")
        
        # Kiểm tra dữ liệu
        issues = reader.validate_data()
        if issues:
            print("⚠️ Các vấn đề phát hiện:")
            for issue in issues:
                print(f"   - {issue}")
        else:
            print("✅ Dữ liệu hợp lệ")
        
        # Lấy danh sách mã hồ sơ
        hoso_codes = reader.get_all_hoso_codes()
        print(f"\n📋 Tìm thấy {len(hoso_codes)} hồ sơ:")
        for code in hoso_codes:
            print(f"   - {code}")
        
        # Test tạo context cho hồ sơ đầu tiên
        if hoso_codes:
            test_code = hoso_codes[0]
            print(f"\n🔍 Test tạo context cho hồ sơ: {test_code}")
            
            context = reader.create_complete_context(test_code)
            
            print(f"📊 Context có {len(context)} trường:")
            for key, value in context.items():
                if key != 'tbl_hanghoa':  # Bỏ qua bảng hàng hóa để tránh spam
                    print(f"   - {key}: {value}")
            
            if 'tbl_hanghoa' in context:
                print(f"   - tbl_hanghoa: {len(context['tbl_hanghoa'])} hàng hóa")
        
    except Exception as e:
        logging.error(f"Lỗi: {e}")

if __name__ == "__main__":
    main()
