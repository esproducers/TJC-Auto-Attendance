#!/bin/bash
cd "$(dirname "$0")" || exit 1

echo "=========================================="
echo "  Auto-Attendance System Setup (Mac/Linux)"
echo "=========================================="
echo

# Check if Python 3 is installed
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python 3 not found! Please install Python from python.org"
    exit 1
fi

# Check Python version (fail if >= 3.13)
python3 -c 'import sys; sys.exit(0 if sys.version_info < (3,13) else 1)'
if [ $? -ne 0 ]; then
    echo "[ERROR] Python 3.13+ detected! 'onnxruntime' does not support Python 3.13 yet."
    echo "        Please install Python 3.12, 3.11, or 3.10 from python.org."
    exit 1
fi

echo "[1/3] Creating virtual environment..."
python3 -m venv venv

echo "[2/3] Installing dependencies..."
source venv/bin/activate
python3 -m pip install --upgrade pip
python3 -m pip install "numpy<2"
python3 -m pip install -r requirements.txt

echo "[3/3] Creating data folders..."
mkdir -p registered_faces database cache records reports

echo ""
echo "=========================================="
echo "  Setup Complete!"
echo "  Use './run.sh' to start the application."
echo "=========================================="
chmod +x run.sh
