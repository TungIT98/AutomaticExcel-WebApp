import pandas as pd
from docxtpl import DocxTemplate
from pathlib import Path
import logging
from excel_reader import ExcelReader

# --- Cấu hình ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
EXCEL_FILE_PATH = Path("0. Hoso.xlsx")
TEMPLATE_FOLDER = Path("templates")
OUTPUT_FOLDER = Path("output")

# ==============================================================================
# NƠI TẬP HỢP VÀ QUẢN LÝ TẤT CẢ CÁC PLACEHOLDER ĐẶC THÙ
# BẠN CHỈ CẦN SỬA VÀ THÊM THÔNG TIN TRONG NÀY
# ==============================================================================
TEMPLATE_CONFIG = {
    # Tất cả dữ liệu sẽ được lấy từ Excel thông qua ExcelReader
    # Chỉ thêm các placeholder đặc thù nếu cần ghi đè dữ liệu từ Excel
}
# ==============================================================================


def main():
    try:
        logging.info(f"Đảm bảo thư mục output tồn tại tại: {OUTPUT_FOLDER}")
        OUTPUT_FOLDER.mkdir(exist_ok=True)

        # Khởi tạo ExcelReader
        logging.info(f"Bắt đầu đọc dữ liệu từ tệp Excel: {EXCEL_FILE_PATH}")
        excel_reader = ExcelReader(EXCEL_FILE_PATH)
        
        # Kiểm tra dữ liệu
        issues = excel_reader.validate_data()
        if issues:
            logging.warning("Phát hiện các vấn đề trong dữ liệu:")
            for issue in issues:
                logging.warning(f"   - {issue}")

        logging.info(f"Đang tìm kiếm các tệp mẫu trong thư mục: {TEMPLATE_FOLDER}")
        templates = list(TEMPLATE_FOLDER.glob("*.docx"))
        
        # Lọc bỏ các file tạm và file lỗi
        templates = [t for t in templates if not t.name.startswith('~$') and t.name != '9. Phieu YCTN TNN.docx']
        
        # Thêm xử lý đặc biệt cho template có placeholder với khoảng trắng
        logging.info("Lưu ý: Template '7. Ke hoach danh gia.docx' có placeholder 'Ngay_ ĐG' với khoảng trắng")
        
        if not templates:
            logging.error(f"Không tìm thấy tệp mẫu .docx nào trong '{TEMPLATE_FOLDER}'.")
            return
        logging.info(f"Tìm thấy {len(templates)} tệp mẫu hợp lệ.")

        # Lấy danh sách mã hồ sơ
        hoso_codes = excel_reader.get_all_hoso_codes()
        logging.info(f"Tìm thấy {len(hoso_codes)} hồ sơ cần xử lý")

        logging.info("Bắt đầu xử lý từng hồ sơ...")
        
        for ma_ho_so in hoso_codes:
            logging.info(f"--- Đang xử lý hồ sơ: {ma_ho_so} ---")
            
            # Tạo context hoàn chỉnh từ ExcelReader
            context = excel_reader.create_complete_context(ma_ho_so)
            
            for template_path in templates:
                try:
                    template_filename = template_path.name
                    
                    # 2. TẠO CONTEXT CUỐI CÙNG (FINAL CONTEXT)
                    final_context = context.copy() 

                    # 3. GỘP THÊM CONTEXT ĐẶC THÙ (nếu có)
                    if template_filename in TEMPLATE_CONFIG:
                        logging.info(f"    + Áp dụng cấu hình đặc thù cho mẫu: {template_filename}")
                        specific_context = TEMPLATE_CONFIG[template_filename]
                        final_context.update(specific_context)

                    doc = DocxTemplate(template_path)
                    doc.render(final_context) # 4. SỬ DỤNG CONTEXT CUỐI CÙNG
                    
                    # Tạo tên file output (loại bỏ ký tự đặc biệt)
                    ten_kh = context.get('ten_kh', 'Unknown').replace(' ', '_').replace('/', '_').replace('\\', '_')
                    ma_ho_so_clean = ma_ho_so.replace('/', '_').replace('\\', '_')
                    template_name_clean = template_path.stem.replace(' ', '_').replace('/', '_').replace('\\', '_')
                    output_filename = f"{template_name_clean}_{ma_ho_so_clean}_{ten_kh}.docx"
                    output_path = OUTPUT_FOLDER / output_filename
                    
                    doc.save(output_path)
                    logging.info(f"    -> Đã tạo thành công tệp: {output_path}")

                except Exception as e:
                    logging.error(f"    -> LỖI khi xử lý mẫu '{template_filename}' cho hồ sơ '{ma_ho_so}': {e}")
        
        logging.info("=== HOÀN TẤT TOÀN BỘ QUÁ TRÌNH! ===")

    except FileNotFoundError:
        logging.error(f"LỖI: Không tìm thấy tệp Excel tại '{EXCEL_FILE_PATH}'.")
    except KeyError as e:
        logging.error(f"LỖI: Tên cột hoặc tên sheet trong file Excel không đúng: {e}")
    except Exception as e:
        logging.error(f"LỖI không xác định đã xảy ra: {e}")

if __name__ == "__main__":
    main()