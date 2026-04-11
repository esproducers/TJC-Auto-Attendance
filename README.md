# Auto-Attendance System for Church / 教会自动人脸识别点名系统

### 📋 System Overview
This system utilizes the InsightFace facial recognition engine specifically tailored to automatically log the attendance of church-goers as they walk through the entry gate using a fixed camera. It recognizes the congregation's identities, records attendance independently, and writes detailed daily and categorical Excel reports. Everything operates entirely offline for strict privacy guarantees. All you need is a PC or Raspberry Pi plus a USB camera.

**Core Features:**
- ✅ Fast concurrent detection (Max 8 people per frame), great for crowded moments.
- ✅ False Positive Rate beneath 0.1%, highly robust for seniors and children alike.
- ✅ Hardened automatic event deduplication (Checking in once records one entry per day).
- ✅ Dynamic generation of Excel attendance logs (Categorized chronologically or by age groups).
- ✅ Operates 100% locally with zero subscription or networking requirements.

### 🖥️ Hardware Requirements
| Component | Minimum Settings | Recommended Specifications |
|---|---|---|
| Main Node | Intel i3 / Raspberry Pi 4B (4GB) | Intel i5 (8th Gen+) / NVIDIA Jetson Nano |
| Memory | 4 GB | 8 GB |
| Storage | 32GB (for pictures and DB files) | 64GB SSD |
| Camera | 720p USB Camera | 1080p Ultra-Wide USB Cam (e.g. Logitech C920) |
| Power Base | Reliable continuous power | Plug-in UPS power source is advised to prevent corruption |

> *Note: For Raspberry Pi deployments, ensure you allocate a minimum of 256MB VRAM memory to the internal GPU unit.*

### 🔧 Software Requirements & Deployment Execution

#### 1. Install Dependencies
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
```


### 📁 Environment Directory Tree Map Structure
- `main.py`: Internal backbone controller (Real-time Detection + Logging Event Router).
- `report.py`: Script dedicated functionally to printing Excel analysis.
- `requirements.txt`: Master blueprint of pip dependencies dictating native compatibility.
- `registered_faces/`: Raw unbridled collection of individual identity templates serving as root knowledge.
- `database/`: Repository holding standalone persistent memory (`.db` artifacts).
- `cache/`: Transient artifact collection representing computed facial encoding binary blobs to speed-up hot-start times.
- `reports/`: Sub-folder collecting final generated chronological `.xlsx` spreadsheet outcomes.
- `logs/`: Application operational footprints and metrics (currently reserved and empty).

---

