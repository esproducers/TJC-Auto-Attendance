# Installation Guide: Auto-Attendance System

This guide explains how to set up the Auto-Attendance tracker on a **New Windows PC** or a **Mac (macOS)**.

## 📋 Prerequisites
- **Python 3.10 or 3.11** (recommended). [Download here](https://www.python.org/downloads/)
- A **Webcam** (USB or built-in).

---

## 💻 1. Windows Installation (Easy)

### Step 1: Install Python
When installing Python from the official website, make sure to check the box: **"Add Python to PATH"**.

### Step 2: Download the Code
Copy the `Auto-Attendance` folder to your new computer (e.g., to your Documents folder).

### Step 3: Install Dependencies
Open **PowerShell** or **Command Prompt** (CMD), navigate to the folder, and run:
```powershell
pip install -r requirements.txt
```

### Step 4: Run the App
```powershell
python app.py
```

---

## 🍎 2. Mac (macOS) Installation

### Step 1: Install Python & Xcode Tools
Open the **Terminal** application and run these commands to set up the environment:
```bash
xcode-select --install
```
(Follow the prompts to install the Apple Developer tools). Then download Python 3.11 from python.org.

### Step 2: Create a Virtual Environment (Recommended on Mac)
```bash
cd /Path/To/Your/Auto-Attendance
python3 -m venv venv
source venv/bin/activate
```

### Step 3: Install Dependencies
```bash
pip install -r requirements.txt
```
> [!NOTE]
> If you have an **M1, M2, or M3 Mac**, you might need to run `pip install onnxruntime-silicon` instead of standard `onnxruntime` if you encounter processing errors.

### Step 4: Run the App
```bash
python3 app.py
```
> [!IMPORTANT]
> macOS will ask for permission to access the **Camera**. You must click **"Allow"** for the facial recognition to work.

---

## 📂 3. Transferring Your Data
If you want to move your existing members and photos to the new computer:
1.  Copy the **`database/attendance.db`** file.
2.  Copy the **`registered_faces/`** folder (contains member photos).
3.  Copy the **`records/`** folder (contains history).

---

## 🚀 4. "One-Click" Portable Version (Optional)
If you want to run the app as a single `.exe` file on Windows without installing anything:
1. Install PyInstaller: `pip install pyinstaller`
2. Run: `pyinstaller --onefile --windowed app.py`
This will create a `dist/app.exe` that you can just double-click!
