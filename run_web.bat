@echo off
echo ========================================
echo    RUNNING WEB APPLICATION
echo ========================================

call .venv\Scripts\activate.bat
cd webapp
python app.py

echo.
echo Press any key to exit...
pause >nul

