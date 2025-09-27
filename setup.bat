@echo off
echo ========================================
echo    SETUP AUTOMATIC EXCEL TO WORD
echo ========================================

echo.
echo [1/4] Creating virtual environment...
python -m venv .venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment
    pause
    exit /b 1
)

echo.
echo [2/4] Activating virtual environment...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment
    pause
    exit /b 1
)

echo.
echo [3/4] Installing core dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install core dependencies
    pause
    exit /b 1
)

echo.
echo [4/4] Installing web dependencies...
cd webapp
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install web dependencies
    pause
    exit /b 1
)
cd ..

echo.
echo ========================================
echo    SETUP COMPLETED SUCCESSFULLY!
echo ========================================
echo.
echo To run CORE script:
echo   .venv\Scripts\activate
echo   python run.py
echo.
echo To run WEB APP:
echo   .venv\Scripts\activate
echo   cd webapp
echo   python app.py
echo.
echo Then open: http://localhost:5000
echo.
pause

