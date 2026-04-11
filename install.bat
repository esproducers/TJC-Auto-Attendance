@echo off
echo ==========================================
echo   Auto-Attendance System Setup (Windows)
echo ==========================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found! Please install Python from python.org and 
    echo ensure 'Add to PATH' is checked during installation.
    pause
    exit /b
)

echo [1/3] Creating virtual environment...
python -m venv venv

echo [2/3] Installing dependencies (this may take a few minutes)...
call venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt

echo [3/3] Creating data folders...
if not exist "registered_faces" mkdir registered_faces
if not exist "database" mkdir database
if not exist "cache" mkdir cache
if not exist "records" mkdir records
if not exist "reports" mkdir reports

echo.
echo ==========================================
echo   Setup Complete! 
echo   Use 'run.bat' to start the application.
echo ==========================================
pause
