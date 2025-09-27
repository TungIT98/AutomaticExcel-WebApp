@echo off
echo ===============================================
echo    EXCEL-TO-WORD WEB APP - CLIENT INSTALLER
echo ===============================================

echo Dang cai dat Python dependencies...
pip install Flask==2.3.3
pip install Flask-SQLAlchemy==3.0.5
pip install Flask-JWT-Extended==4.5.3
pip install pandas
pip install openpyxl
pip install python-docx
pip install docxtpl
pip install jinja2
pip install lxml

echo.
echo ===============================================
echo    CAI DAT THANH CONG!
echo ===============================================
echo.
echo De chay chuong trinh:
echo 1. Chay: run_webapp.bat
echo 2. Truy cap: http://localhost:5000
echo 3. Admin: admin / admin123
echo.
echo ===============================================

pause
