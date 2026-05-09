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

echo [2/3] Installing dependencies...
call venv\Scripts\activate

:: Upgrade pip first
python -m pip install --upgrade pip

:: [LIGHTWEIGHT FIX] Try installing pre-compiled insightface for Python 3.10
:: This avoids needing 5GB of Visual Studio Build Tools
python --version | findstr "3.10" >nul
if %errorlevel% equ 0 (
    echo [!] Python 3.10 detected. Installing lightweight insightface binary...
    python -m pip install https://github.com/Gourieff/Assets/raw/main/Insightface/insightface-0.7.3-cp310-cp310-win_amd64.whl
) else (
    echo [!] Non-3.10 version or wheel failed. Attempting standard installation...
    python -m pip install insightface
)

:: Install all other requirements (ensuring numpy 1.x)
pip install "numpy<2"
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
