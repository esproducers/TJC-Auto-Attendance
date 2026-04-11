#!/bin/bash
if [ ! -d "venv" ]; then
    echo "[!] Virtual environment not found. Running setup first..."
    ./install.sh
fi
echo "Starting Auto-Attendance System..."
source venv/bin/activate
python3 app.py
