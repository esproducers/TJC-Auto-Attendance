#!/bin/bash
cd "$(dirname "$0")" || exit 1

if [ ! -d "venv" ]; then
    echo "[!] Virtual environment not found. Running setup first..."
    ./install.sh
fi
echo "Starting Auto-Attendance System..."
source venv/bin/activate
nohup python3 app.py >/dev/null 2>&1 &
osascript -e 'tell application "Terminal" to close first window' &
exit
