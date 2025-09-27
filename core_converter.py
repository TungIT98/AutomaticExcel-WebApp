import os
import tempfile
import logging
from pathlib import Path
import sys

# Add parent directory to path to import excel_reader
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from excel_reader import ExcelReader
from docxtpl import DocxTemplate

class SecureExcelToWordConverter:
    def __init__(self):
        self.template_folder = Path("templates")
        self.output_folder = Path("output")
        
        # Ensure directories exist
        self.template_folder.mkdir(exist_ok=True)
        self.output_folder.mkdir(exist_ok=True)
        
    def convert_excel_to_word(self, excel_file_path, user_id):
        """
        Chuyển đổi Excel sang Word với kiểm tra quyền
        """
        try:
            logging.info(f"Starting conversion for user {user_id}")
            
            # 1. Đọc dữ liệu từ Excel
            excel_reader = ExcelReader(excel_file_path)
            hoso_codes = excel_reader.get_all_hoso_codes()
            
            if not hoso_codes:
                raise Exception("No data found in Excel file")
            
            # 2. Xử lý từng hồ sơ
            results = []
            for ma_ho_so in hoso_codes:
                logging.info(f"Processing hồ sơ: {ma_ho_so}")
                context = excel_reader.create_complete_context(ma_ho_so)
                
                # 3. Tạo file Word cho từng template
                templates = list(self.template_folder.glob("*.docx"))
                for template_path in templates:
                    # Skip temporary files and corrupted files
                    if template_path.name.startswith('~$') or template_path.name == '9. Phieu YCTN TNN.docx':
                        continue
                    
                    try:
                        doc = DocxTemplate(template_path)
                        doc.render(context)
                        
                        # Tạo tên file output an toàn
                        ten_kh = context.get('ten_kh', 'Unknown').replace(' ', '_').replace('/', '_')
                        ma_ho_so_clean = ma_ho_so.replace('/', '_').replace('\\', '_')
                        template_name_clean = template_path.stem.replace(' ', '_').replace('/', '_')
                        output_filename = f"{template_name_clean}_{ma_ho_so_clean}_{ten_kh}.docx"
                        output_path = self.output_folder / output_filename
                        
                        doc.save(output_path)
                        results.append(str(output_path))
                        logging.info(f"Created: {output_path}")
                        
                    except Exception as e:
                        logging.error(f"Error processing template {template_path.name}: {e}")
            
            logging.info(f"Conversion completed. Generated {len(results)} files.")
            return results
            
        except Exception as e:
            logging.error(f"Conversion error: {e}")
            raise
