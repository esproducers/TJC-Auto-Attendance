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

## 💻 1. Windows Installation

### Step 1: Install Python
When installing Python from the official website, make sure to check the box: **"Add Python to PATH"**.

### Step 2: Download the Code
Downlaod and extract the `Auto-Attendance` folder to your computer (e.g., C:\Program Files).

### Step 3: Run Setup
Open the folder, and Double-click "install" to auto install:
```powershell
install.bat
```

### Step 4: Run the App
*Wait for it to finish.* Then, whenever you want to start the app, just run:
```bash
run.bat
```

---

**## 🍎 2. Mac (macOS) Installation**
9+-
### Step 1: Install Python
When installing Python from the official website

### Step 2: Install JDK-JAVA
When installing JDK-JAVA from the Oracle JDK download page 
(e.g., JDK 21 -x64 for Intel, aarch64 for Apple Silicon)

### Step 3: Open Terminal and Navigate correctly
1.  Open **Terminal**.
2.  Type `cd ` (type cd followed by a **space**, do not press Enter yet).
3.  Find your project folder on your Desktop.
4.  **Drag and Drop** that folder directly into the Terminal window. It will automatically type the correct path for you.
5.  Press **Enter**.

### Step 3: Run the Setup
Once you are successfully inside the folder (you should see the folder name in your Terminal prompt), copy and paste this **single line** and press Enter:
```bash
chmod +x *.sh && ./install.sh

### Step 4: Start the App
After the installation finishes, just type:
```bash
./run.sh
```

> [!IMPORTANT]
> macOS will ask for permission to access the **Camera**. You must click **"Allow"** for the facial recognition to work.

## ✨ Pro Tip: Create a Desktop Shortcut
To make it even easier for others, you can create a shortcut on your desktop:

### Windows
1. Open the project folder.
2. **Right-click** `run.bat` > **Send to** > **Desktop (create shortcut)**.
3. Rename the shortcut on your desktop to **"Auto-Attendance"**.

### Mac (macOS)
1. Open the project folder.
2. **Right-click** `run.sh` > **Make Alias**.
3. Drag the **Alias** to your desktop and rename it.

---

## 📖 User Guide & Administration

### 👥 Accounts & Access
- **Default Account**: Used for general attendance capture. Restricted from reports and settings.
- **Admin Account**: 
  - **User**: `admin`
  - **Password**: `admin123`
- **Security Word**: Set this in **Settings**. It is required to reset your password if forgotten.

### 🔐 Security & Passwords
- **Network Security**: The "🛡️ Firewall Active" badge in the sidebar indicates that all data is kept 100% offline. You can toggle this if remote sync is ever required.
- **Resetting Password**: 
  - If forgotten: Use the **"Forgot Password"** link on the login screen (requires Security Word).
  - While logged in: Go to **Settings** > **Update Admin Password**.

### 📸 Camera Management
- The system supports multiple cameras (e.g., Built-in + USB).
- Select your preferred camera from the **Sidebar Dropdown**.
- Use the **🔄 Refresh** button to detect new USB cameras without restarting.

### 📅 Seminar Reporting (Annual Report)
- When clicking **"Start"** on the Dashboard, you must select the **Seminar Type**:
  - `Normal`
  - `Friday Seminar`
  - `Saturday Seminar`
- **Crucial**: The **Annually Report** only calculates data from sessions marked as `Friday Seminar` or `Saturday Seminar`. Selecting these correctly ensures your Weekly, Monthly, and Yearly summaries are accurate.

---

## 📂 3. Transferring Your Data
If you want to move your existing members and photos to the new computer:
1.  Copy the **`database/attendance.db`** file.
2.  Copy the **`registered_faces/`** folder (contains member photos).
3.  Copy the **`records/`** folder (contains history).

---

## 📁4. Environment Directory Tree Map Structure
- `main.py`: Internal backbone controller (Real-time Detection + Logging Event Router).
- `app.py`: The graphical user interface (GUI).
- `report.py`: Script dedicated functionally to printing Excel analysis.
- `requirements.txt`: Master blueprint of pip dependencies dictating native compatibility.
- `registered_faces/`: Raw unbridled collection of individual identity templates serving as root knowledge.
- `database/`: Repository holding standalone persistent memory (`.db` artifacts).
- `cache/`: Transient artifact collection representing computed facial encoding binary blobs to speed-up hot-start times.
- `reports/`: Sub-folder collecting final generated chronological `.xlsx` spreadsheet outcomes.
- `logs/`: Application operational footprints and metrics (currently reserved and empty).
