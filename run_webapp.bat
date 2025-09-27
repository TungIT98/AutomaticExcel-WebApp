@echo off
echo ===============================================
echo    EXCEL-TO-WORD WEB APP
echo ===============================================

echo Dang cai dat dependencies...
pip install -r requirements.txt

echo.
echo Dang khoi dong Web App...
echo Truy cap: http://localhost:5000
echo Admin Panel: http://localhost:5000/admin
echo.
echo Default Admin:
echo Username: admin
echo Password: admin123
echo.

python app.py

pause
