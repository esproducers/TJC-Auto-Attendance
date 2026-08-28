@echo off
cd /d "%~dp0"

if not exist "venv" (
    echo [!] Virtual environment not found. Running setup first...
    call install.bat
)
echo Starting Auto-Attendance System...
call venv\Scripts\activate
start "" pythonw.exe app.py
exit
