@echo off
if not exist "venv" (
    echo [!] Virtual environment not found. Running setup first...
    call install.bat
)
echo Starting Auto-Attendance System...
call venv\Scripts\activate
python app.py
pause
