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
- **Python 3.11** (recommended). [Download here](https://www.python.org/downloads/)
- A **Webcam** (USB or built-in).

---

## 💻 1. Windows Installation

### Step 1: Install Python 3.11.0
When installing Python from the official website, make sure to check the box: **"Add Python to PATH"**.

### Step 2: Download the Code
bash
git clone https://github.com/esproducers/TJC-Auto-Attendance.git
🔑 Notes
This requires Git installed on your system.

On Windows: install from git-scm.com.

On macOS/Linux: Git is usually preinstalled, or you can install via package manager (brew install git or sudo apt install git).

The command will create a local folder named TJC-Auto-Attendance in your current directory. 
Notes: You can pull the latest updates anytime in settings. 


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
### Step 1: Install Python 3.11.0
When installing Python from the official website

### Step 2: Install JDK-JAVA
When installing JDK-JAVA from the Oracle JDK download page 
(e.g., JDK 21 -x64 for Intel, aarch64 for Apple Silicon)

### Step 3: Download the Code
Open Terminal and run the following command: 
```bash
bash
git clone https://github.com/esproducers/TJC-Auto-Attendance.git
```
🔑 Notes
This requires Git installed on your system.

On Windows: install from git-scm.com.

On macOS/Linux: Git is usually preinstalled, or you can install via package manager (brew install git or sudo apt install git).

The command will create a local folder named TJC-Auto-Attendance in your current directory. 
Notes: You can pull the latest updates anytime in settings. 


### Step 4: Open Terminal and Navigate correctly
1.  In **Terminal**.
2.  Type `cd ` (type cd followed by a **space**, do not press Enter yet).
3.  Find your project folder on your Desktop.
4.  **Drag and Drop** that folder directly into the Terminal window. It will automatically type the correct path for you.
5.  Press **Enter**.

### Step 5: Run the Setup
Once you are successfully inside the folder (you should see the folder name in your Terminal prompt), copy and paste this **single line** and press Enter:
```bash
chmod +x *.sh && ./install.sh
```

# Note: If you see "Building wheel for opencv-python (pyproject.toml) ..." and it stuck for a long time, it's normal, just wait for it to finish. (Might take 1-3 hours on Mac)

### Step 6: Start the App
After the installation finishes, just Double-click:
```bash
Mac_Start.command
```

> [!WARNING]
> **Troubleshooting: "could not be executed because you do not have appropriate access privileges"**
> If you get this error when double-clicking `Mac_Start.command`, open your Terminal, type:
```bash
chmod +x Mac_Start.command
```
Press Enter. This grants your Mac permission to run the file.


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
2. **Right-click** `Mac_Start.command` > **Make Alias**.
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
- The system supports multiple cameras (e.g., Built-in + USB + WiFi).
- Select your preferred camera from the **Sidebar Dropdown**.
- Use the **🔄 Refresh** button to detect new USB cameras without restarting.

#### 📡 WiFi Camera Setup & Troubleshooting
The **Auto Search WiFi** feature relies on standard network broadcasting protocols (ONVIF/SSDP). However, it may fail to find your camera automatically due to:
1. **Mac Firewall:** Macs have strict built-in firewalls that often block the incoming network broadcast responses.
2. **Camera Security:** If you are using a **TP-Link Tapo** camera, TP-Link disables the video stream by default. You must open the Tapo app on your phone, go to **Camera Settings -> Advanced Settings -> Camera Account**, and create a username/password to activate the stream.
3. **Router Isolation:** Sometimes Mesh routers (like TP-Link Deco) isolate devices so they can't talk to each other directly.

**The Fix (Manual Add):**
Since the automatic search can be blocked by security settings, it is much faster and more reliable to use the red **`+ Add URL`** button right underneath it!
1. Find the camera's IP address in your router app (e.g., `192.168.68.105`).
2. Click the red **`+ Add URL`** button in the Auto-Attendance app.
3. Type in the direct video link. If it's a Tapo camera, it will look exactly like this: 
   `rtsp://your_username:your_password@192.168.68.105:554/stream1`
   *(Replace `your_username` and `your_password` with the camera account details you set up in the Tapo app, and replace the IP address with your real one).*

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
