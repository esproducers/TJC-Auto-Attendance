# Guide: Building a Portable "Zero-Install" App

This guide will help you create a single "Auto-Attendance" folder that runs on any PC without needing Python installed and without needing an internet connection.

## 🛠️ Phase 1: Fueling the "Portable Brain" (Offline Models)

InsightFace usually downloads its AI models to a hidden folder on your Windows User account. We need to copy these into the project folder so they are included in the standalone version.

1.  **Find the models**: On your current PC, go to your "User" folder:
    - Path: `C:\Users\[YOUR_USERNAME]\.insightface\models\buffalo_l`
2.  **Copy the folder**: Copy the entire **`buffalo_l`** folder.
3.  **Paste into Project**: Go to your `Auto-Attendance` project folder and create a new directory structure:
    - Target Path: `Auto-Attendance\models\insightface\models\buffalo_l`
    - (Paste the files inside the last `buffalo_l` folder).

---

## 🏗️ Phase 2: Create the Standalone Folder

Once the models are copied, we will use **PyInstaller** to build the app.

1.  **Open PowerShell** in your `Auto-Attendance` folder.
2.  **Install PyInstaller**:
    ```powershell
    pip install pyinstaller
    ```
3.  **Run the Build Command**:
    ```powershell
    pyinstaller --noconfirm --onedir --windowed `
    --add-data "database;database" `
    --add-data "registered_faces;registered_faces" `
    --add-data "models;models" `
    --add-data "records;records" `
    app.py
    ```

---

## 🚀 Phase 3: Deployment

1.  Go to the new **`dist/app`** folder.
2.  **Copy** the entire `app` folder to any other computer.
3.  **Run it**: Double-click **`app.exe`**.

Everything is now inside that folder: 
- ✅ The Python engine
- ✅ The InsightFace AI models (no Internet needed)
- ✅ Your member database and photos 

---
> [!NOTE]
> If you add more people or check-ins on the new PC, those changes stay inside the `app` folder's `database/attendance.db` file. You can simply copy this entire folder back and forth to keep your attendance data in sync!

---
> [!IMPORTANT]
> **Cross-Platform Building**: This guide is specifically for **Windows**. PyInstaller creates apps for the system it is running on. 
> - To make a **Mac app**, you must follow these same steps on a **Mac computer**. 
> - A `.exe` file generated on Windows will **not** run natively on macOS.


---

### 🍎 How to Build for Mac?
To create a portable version for macOS, follow these steps on a **Mac computer**:

1.  **Install Python**: Download and install the latest version from [python.org](https://www.python.org/).
2.  **Install All Libraries**: Open the Terminal in your project folder and install the project's dependencies:
    ```bash
    pip3 install flask opencv-python insightface pyinstaller
    ```
3.  **Run the Build**: Use the same `pyinstaller` command from Phase 2.
    *   This will generate a **`.app`** bundle (the Mac equivalent of a `.exe`).