# HƯỚNG DẪN SETUP VÀ CHẠY ỨNG DỤNG

## 📁 CẤU TRÚC THƯ MỤC

```
AutomaticExcel/
├── 📁 CORE (Chạy local - Script gốc)
│   ├── run.py                    # Script chính
│   ├── excel_reader.py          # Đọc Excel
│   ├── extract_placeholders.py   # Trích xuất placeholders
│   ├── fix_template.py          # Sửa template
│   ├── 0. Hoso.xlsx             # File Excel mẫu
│   ├── templates/               # 13 file Word templates
│   ├── output/                  # Kết quả xuất ra
│   └── requirements.txt          # Dependencies cho core
│
├── 📁 WEBAPP (Web application)
│   ├── app.py                   # Flask app chính
│   ├── templates/               # HTML templates
│   │   ├── index.html          # Giao diện user
│   │   └── admin.html          # Giao diện admin
│   ├── requirements.txt         # Dependencies cho web
│   └── instance/               # Database SQLite
│
└── 📁 SHARED
    ├── templates/              # Templates dùng chung
    └── uploads/               # File upload tạm
```

## 🚀 CÁCH CHẠY

### 1. CHẠY SCRIPT GỐC (Local)
```bash
# Vào thư mục gốc
cd C:\AutomaticExcel

# Tạo môi trường ảo
python -m venv .venv

# Kích hoạt môi trường ảo
.venv\Scripts\activate

# Cài đặt dependencies
pip install -r requirements.txt

# Chạy script
python run.py
```

### 2. CHẠY WEB APP
```bash
# Vào thư mục webapp
cd C:\AutomaticExcel\webapp

# Kích hoạt môi trường ảo (cùng với core)
..\.venv\Scripts\activate

# Cài đặt dependencies
pip install -r requirements.txt

# Chạy web app
python app.py
```

## 🔧 MÔI TRƯỜNG

### Core Dependencies:
- pandas>=1.3.0
- openpyxl>=3.0.0  
- python-docx>=0.8.11
- docxtpl>=0.20.0
- jinja2>=3.0.0
- lxml>=4.6.0

### Web Dependencies:
- Flask==2.3.3
- Flask-JWT-Extended==4.5.3
- Flask-SQLAlchemy==3.0.5
- Werkzeug==2.3.7
- gunicorn==21.2.0

## 📝 GHI CHÚ

1. **Môi trường ảo**: Dùng chung `.venv` cho cả core và webapp
2. **Templates**: Dùng chung từ thư mục gốc
3. **Database**: SQLite trong `webapp/instance/`
4. **Output**: Core xuất vào `output/`, Web xuất vào `webapp/output/`

