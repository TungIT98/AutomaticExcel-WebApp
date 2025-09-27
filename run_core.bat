@echo off
echo ========================================
echo    RUNNING CORE EXCEL TO WORD SCRIPT
echo ========================================

call .venv\Scripts\activate.bat
python run.py

echo.
echo Press any key to exit...
pause >nul

