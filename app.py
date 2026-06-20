import os
import sys
import socket
import ipaddress
import shutil
import uuid
from tkcalendar import DateEntry

# --- CONFIDENTIAL DATA FIREWALL (BLOCKS ALL NETWORK TRAFFIC EXCEPT LOCAL LAN) ---
# This block ensures data privacy by preventing unauthorized external connections.
FIREWALL_BYPASS = False

_orig_connect = socket.socket.connect
_orig_getaddrinfo = socket.getaddrinfo
_orig_bind = socket.socket.bind
_orig_sendto = socket.socket.sendto

def _is_private_ip(ip_str):
    if not ip_str:
        return False
    if ip_str in ('localhost', '::1'):
        return True
    try:
        if '%' in ip_str:
            ip_str = ip_str.split('%')[0]
        ip = ipaddress.ip_address(ip_str)
        return ip.is_loopback or ip.is_private or ip.is_link_local
    except ValueError:
        return False

def _is_local_host_or_ip(host):
    if not host:
        return False
    if host in ('localhost', '::1'):
        return True
    if _is_private_ip(host):
        return True
    try:
        results = _orig_getaddrinfo(host, None)
        for r in results:
            sockaddr = r[4]
            ip = sockaddr[0]
            if not _is_private_ip(ip):
                return False
        return True
    except Exception:
        return False

def _is_local(address):
    if isinstance(address, tuple): host = address[0]
    else: host = str(address)
    return _is_local_host_or_ip(host)

def secure_connect(self, address):
    if FIREWALL_BYPASS or _is_local(address): return _orig_connect(self, address)
    raise ConnectionRefusedError(f"Firewall Blocked: Outbound connection to {address} is prohibited.")

def secure_getaddrinfo(host, port, *args, **kwargs):
    if FIREWALL_BYPASS or _is_local(host): return _orig_getaddrinfo(host, port, *args, **kwargs)
    raise ConnectionRefusedError(f"Firewall Blocked: DNS lookup for {host} is prohibited.")

def secure_bind(self, address):
    if FIREWALL_BYPASS or _is_local(address): return _orig_bind(self, address)
    raise ConnectionRefusedError(f"Firewall Blocked: Binding on {address} is prohibited.")

def secure_sendto(self, data, address):
    if FIREWALL_BYPASS or _is_local(address): return _orig_sendto(self, data, address)
    raise ConnectionRefusedError(f"Firewall Blocked: UDP packet to {address} is prohibited.")

socket.socket.connect = secure_connect
socket.getaddrinfo = secure_getaddrinfo
socket.socket.bind = secure_bind
socket.socket.sendto = secure_sendto

print("[SECURITY] Confidential Data Firewall is ACTIVE.")
# ----------------------------------------------------------------

import sqlite3
import json
import os
os.environ["OPENCV_LOG_LEVEL"] = "SILENT"
import cv2
import customtkinter as ctk
import pandas as pd
import threading
import queue
import time
import calendar
import uuid
import concurrent.futures
from PIL import Image
from datetime import datetime, date, timedelta
from main import InsightFaceAttendance
from report import ReportGenerator
import tkinter as tk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


# ── Helpers ────────────────────────────────────────────────────────────────────

def _type_color(t):
    t = (t or "").lower()
    if "truth" in t: return "#17A2B8"
    if "unknown" in t: return "#DC3545"
    if "other area" in t: return "#6366F1"
    if "area member" in t: return "#10B981"
    return "#10B981"

def _type_label(t):
    t = (t or "").lower()
    if "truth" in t: return "Truth Seeker"
    if "unknown" in t: return "?"
    if "other area" in t: return "Other Area Member"
    if "area member" in t: return "Area Member"
    return "Area Member"


# ── Widgets ────────────────────────────────────────────────────────────────────

class CheckInCard(ctk.CTkFrame):
    def __init__(self, master, att_id, name, age, img_path, m_type,
                 member_code=None, on_click=None, on_identify=None):
        super().__init__(master, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=10, height=60)
        self.pack_propagate(False) # Preserve height
        self.member_code = member_code
        self._search_data = f"{name} {member_code}".lower()

        # [1] Profile Photo
        img_f = ctk.CTkFrame(self, width=50, height=50, fg_color="transparent")
        img_f.pack(side="left", padx=10, pady=5)
        img_f.pack_propagate(False)
        
        img_lbl = ctk.CTkLabel(img_f, text="📷", font=("Arial", 16))
        img_lbl.pack(expand=True)
        
        # Optimization: Use global image cache provided by app
        app = master.master.master.master.master if hasattr(master.master.master.master.master, "image_cache") else None
        ci = None
        if app and img_path in app.image_cache:
            ci = app.image_cache[img_path]
        elif img_path and os.path.exists(img_path):
            try:
                pil = Image.open(img_path).resize((40, 40))
                ci = ctk.CTkImage(light_image=pil, size=(40, 40))
                if app: app.image_cache[img_path] = ci
            except: pass

        if ci:
            img_lbl.configure(image=ci, text="")

        # [2] Name and Code Info
        txt_f = ctk.CTkFrame(self, fg_color="transparent")
        txt_f.pack(side="left", fill="both", expand=True, padx=5)
        
        ctk.CTkLabel(txt_f, text=(name or "Unknown").upper(), font=("Arial", 12, "bold"), anchor="w").pack(pady=(8, 0), fill="x")
        ctk.CTkLabel(txt_f, text=f"ID: {member_code or '?'}", font=("Arial", 10), text_color="#6B7280", anchor="w").pack(fill="x")

        # [3] Status badge
        right_f = ctk.CTkFrame(self, fg_color="transparent")
        right_f.pack(side="right", padx=15)

        color = _type_color(m_type)
        ctk.CTkLabel(right_f, text=_type_label(m_type).upper(), font=("Arial", 8, "bold"), fg_color=color, text_color="white", corner_radius=4, width=80).pack(pady=8)

        # Actions
        if (m_type or "").lower() == "unknown" and on_identify and att_id:
            def _click_unk(_e=None, aid=att_id, ip=img_path): on_identify(aid, ip)
            self.bind("<Button-1>", _click_unk)
            for w in self.winfo_children(): w.bind("<Button-1>", _click_unk)
        elif on_click and member_code:
            def _click(_e=None, mc=member_code): on_click(mc)
            self.bind("<Button-1>", _click)
            for w in self.winfo_children(): w.bind("<Button-1>", _click)


# ── Tooltip Helper ─────────────────────────────────────────────────────────────

class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.id = None
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hide()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(300, self.show)

    def unschedule(self):
        id_ = self.id
        self.id = None
        if id_:
            self.widget.after_cancel(id_)

    def show(self, event=None):
        x = y = 0
        x, y, cx, cy = self.widget.bbox("insert") if hasattr(self.widget, 'bbox') else (0,0,0,0)
        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + 25
        
        # Create a Toplevel window (standard tkinter tooltips)
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(tw, text=self.text, justify='left',
                         background="#333333", foreground="white",
                         relief='flat', borderwidth=0,
                         padx=8, pady=4,
                         font=("Arial", "9"))
        label.pack(ipadx=1)

    def hide(self):
        tw = self.tooltip_window
        self.tooltip_window = None
        if tw:
            tw.destroy()


# ── Main App ───────────────────────────────────────────────────────────────────

class CustomCalendar(ctk.CTkFrame):
    def __init__(self, parent, on_select, initial_val="", **kwargs):
        super().__init__(parent, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=12, **kwargs)
        self.on_select = on_select
        
        # Parse initial value (DD-MM-YYYY)
        now = datetime.now()
        self.cur_month = now.month
        self.cur_year = now.year
        
        if initial_val:
            try:
                d, m, y = map(int, initial_val.split("-"))
                self.cur_month, self.cur_year = m, y
            except: pass
            
        self.setup_ui()
        self.render_month(self.cur_month, self.cur_year)

    def setup_ui(self):
        # Header
        header = ctk.CTkFrame(self, fg_color="#007BFF", height=45, corner_radius=8)
        header.pack(fill="x", padx=5, pady=5)
        header.pack_propagate(False)
        
        ctk.CTkButton(header, text="<", width=30, fg_color="transparent", hover_color="#0069D9", font=("Arial", 14, "bold"), command=self.prev_month).pack(side="left", padx=5)
        self.month_lbl = ctk.CTkLabel(header, text="Month Year", font=("Arial", 13, "bold"), text_color="white")
        self.month_lbl.pack(side="left", expand=True)
        ctk.CTkButton(header, text=">", width=30, fg_color="transparent", hover_color="#0069D9", font=("Arial", 14, "bold"), command=self.next_month).pack(side="right", padx=5)
        
        # Weekdays Header
        days_f = ctk.CTkFrame(self, fg_color="transparent")
        days_f.pack(fill="x", padx=10, pady=2)
        for i, d in enumerate(["S", "M", "T", "W", "T", "F", "S"]):
            days_f.grid_columnconfigure(i, weight=1)
            ctk.CTkLabel(days_f, text=d, font=("Arial", 10, "bold"), text_color="#9CA3AF").grid(row=0, column=i)
            
        # Grid Container
        self.grid_f = ctk.CTkFrame(self, fg_color="transparent")
        self.grid_f.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        for i in range(7): self.grid_f.grid_columnconfigure(i, weight=1)
        for i in range(6): self.grid_f.grid_rowconfigure(i, weight=1)

    def render_month(self, month, year):
        for w in self.grid_f.winfo_children(): w.destroy()
        
        m_name = calendar.month_name[month]
        self.month_lbl.configure(text=f"{m_name} {year}")
        
        # Get month matrix (0 = empty)
        cal = calendar.Calendar(firstweekday=6) # Sunday start
        month_days = cal.monthdayscalendar(year, month)
        
        now = datetime.now()
        
        for r, week in enumerate(month_days):
            for c, day in enumerate(week):
                if day == 0: continue
                
                is_today = (day == now.day and month == now.month and year == now.year)
                bg = "#EBF5FF" if is_today else "transparent"
                txt = "#007BFF" if is_today else "#1F2937"
                
                btn = ctk.CTkButton(self.grid_f, text=str(day), width=32, height=32, 
                                    fg_color=bg, text_color=txt, hover_color="#F3F4F6", 
                                    font=("Arial", 11, "bold" if is_today else "normal"),
                                    command=lambda d=day: self.select_day(d))
                btn.grid(row=r, column=c, padx=2, pady=4)

    def select_day(self, day):
        val = f"{day:02d}-{self.cur_month:02d}-{self.cur_year}"
        self.on_select(val)

    def prev_month(self):
        self.cur_month -= 1
        if self.cur_month < 1: self.cur_month = 12; self.cur_year -= 1
        self.render_month(self.cur_month, self.cur_year)

    def next_month(self):
        self.cur_month += 1
        if self.cur_month > 12: self.cur_month = 1; self.cur_year += 1
        self.render_month(self.cur_month, self.cur_year)

class WheelDatePicker(ctk.CTkToplevel):
    def __init__(self, parent, title, initial_val="", on_ok=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x520")
        self.attributes("-topmost", True)
        self.grab_set()
        self.resizable(False, False)
        self.configure(fg_color="#FFFFFF")
        
        self.on_ok = on_ok
        self.MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        
        # Parse initial (DD-MM-YYYY)
        now = datetime.now()
        d_init, m_init, y_init = now.day, self.MONTHS[now.month-1], now.year
        if initial_val:
            try:
                parts = initial_val.split("-")
                if len(parts) == 3:
                    if len(parts[0]) == 2: # DD-MM-YYYY
                        d_init = int(parts[0])
                        m_init = self.MONTHS[int(parts[1])-1]
                        y_init = int(parts[2])
                    else: # YYYY-MM-DD
                        y_init = int(parts[0])
                        m_init = self.MONTHS[int(parts[1])-1]
                        d_init = int(parts[2])
            except: pass
            
        self.sel_d = ctk.IntVar(value=d_init)
        self.sel_m = ctk.StringVar(value=m_init)
        self.sel_y = ctk.IntVar(value=y_init)
        
        # Header
        ctk.CTkLabel(self, text=title, font=("Arial", 24, "bold"), text_color="#1F2937").pack(pady=(25, 20))
        
        # Container for wheels
        wheel_frame = ctk.CTkFrame(self, fg_color="transparent")
        wheel_frame.pack(fill="both", expand=True, padx=30)
        
        # Selection Highlight (The blue bar in the middle)
        highlight = ctk.CTkFrame(wheel_frame, fg_color="#F3F4F6", height=45, corner_radius=10)
        highlight.place(relx=0.5, rely=0.5, anchor="center", relwidth=1.0)
        
        self.cols = {}
        
        # Day Column
        self._create_column(wheel_frame, "Day", [f"{i:02d}" for i in range(1, 32)], self.sel_d, 0)
        # Month Column
        self._create_column(wheel_frame, "Month", self.MONTHS, self.sel_m, 1)
        # Year Column
        self._create_column(wheel_frame, "Year", [str(i) for i in range(1900, 2101)], self.sel_y, 2)
        
        # Buttons
        btn_f = ctk.CTkFrame(self, fg_color="transparent")
        btn_f.pack(fill="x", pady=25, padx=30)
        
        ctk.CTkButton(btn_f, text="Cancel", fg_color="#F3F4F6", text_color="#374151", hover_color="#E5E7EB", 
                      height=45, corner_radius=22, command=self.destroy).pack(side="left", expand=True, padx=(0, 10), fill="x")
        
        ctk.CTkButton(btn_f, text="OK", fg_color="#007BFF", hover_color="#0069D9", 
                      height=45, corner_radius=22, command=self._on_ok_click).pack(side="left", expand=True, fill="x")

    def _create_column(self, parent, label, values, var, col_idx):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(side="left", expand=True, fill="both")
        
        scroll = ctk.CTkScrollableFrame(frame, fg_color="transparent", height=250, width=80)
        scroll.pack(fill="both", expand=True)
        
        # Store buttons to update their color
        btns = []
        
        def select(val, btn):
            try:
                var.set(val)
            except:
                var.set(str(val))
            # Update colors
            for b in btns:
                b.configure(text_color="#6B7280", font=("Arial", 14))
            btn.configure(text_color="#007BFF", font=("Arial", 18, "bold"))
            # Scroll to make it central (best effort)
            # scroll._parent_canvas.yview_moveto(...) is internal, skipping for now
            
        for val in values:
            is_sel = (str(val) == str(var.get()) or (isinstance(val, int) and val == var.get()))
            color = "#007BFF" if is_sel else "#6B7280"
            font = ("Arial", 18, "bold") if is_sel else ("Arial", 14)
            
            btn = ctk.CTkButton(scroll, text=str(val), fg_color="transparent", text_color=color, 
                                hover_color="#F3F4F6", font=font, height=40,
                                command=lambda v=val, b=None: select(v, b))
            btn.configure(command=lambda v=val, b=btn: select(v, b))
            btn.pack(fill="x")
            btns.append(btn)
            
            if is_sel:
                # Try to scroll to it after a short delay
                self.after(100, lambda b=btn: scroll._parent_canvas.yview_moveto(max(0, (btns.index(b)-2)/len(btns))))

    def _on_ok_click(self):
        d = self.sel_d.get()
        m = self.sel_m.get()
        y = self.sel_y.get()
        m_idx = self.MONTHS.index(m) + 1
        res = f"{d:02d}-{m_idx:02d}-{y}"
        if self.on_ok:
            self.on_ok(res)
        self.destroy()

class AutoAttendanceApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Auto-Attendance System")
        self.geometry("1280x820")
        self.minsize(1100, 700)

        self.load_settings()
        last_cam = self.settings.get("last_camera_id", 0)
        self.backend  = InsightFaceAttendance(camera_id=last_cam)
        self.backend.init_database() # Ensure DB is migrated before UI queries it
        self.reporter = ReportGenerator()
        
        self.after(500, self.initialize_face_engine)

        self.is_marking       = False
        self.is_paused        = False
        self.session_title    = ""
        self.session_deadline = None
        
        # Performance/Threading states
        self.result_queue = queue.Queue()
        self.is_processing = False
        self.last_results = []
        self.last_stats_count = 0 
        self.last_waiting_count = -1
        self.gui_queue = queue.Queue()
        self.process_gui_queue()
        
        self.capture_feedback = {"msg": "", "expiry": 0, "color": (0,255,0)}
        self.image_cache = {} # Map img_path -> CTKImage
        self.is_admin = False
        if "admin_pass" not in self.settings:
            self.settings["admin_pass"] = "admin123"
        if "admin_hint" not in self.settings:
            self.settings["admin_hint"] = "skudai"
        self.save_settings()

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.init_sidebar()

        self.main_area = ctk.CTkFrame(self, fg_color="#F0F2F5", corner_radius=0)
        self.main_area.grid(row=0, column=1, sticky="nsew")
        self.main_area.grid_columnconfigure(0, weight=1)
        self.main_area.grid_rowconfigure(1, weight=1)

        self.init_header()

        self.container = ctk.CTkFrame(self.main_area, fg_color="transparent")
        self.container.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(0, weight=1)

        self.frames = {}
        self.init_dashboard()
        self.init_members_page()
        self.init_logs_page()
        self.init_reports_page()
        self.init_org_chart_page()
        self.init_sql_page()
        self.init_settings_page()
        self.init_annually_report_page()
        self.check_auto_periodical_reports()


        self.show_frame("dashboard")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.update_camera()

        # Update README with admin info
        self._update_readme_admin()

    def initialize_face_engine(self):
        """Attempts to load AI models, asking for temporary internet access if they are missing."""
        try:
            # First attempt: With firewall active
            self.backend.prepare()
            print("[INFO] AI Engine initialized successfully (Offline).")
        except Exception as e:
            print(f"[WARN] AI Engine failed to initialize: {e}")
            # If it failed, it might be due to missing models needing download
            if messagebox.askyesno("AI Models Missing", 
                                   "The required Face Recognition models are missing from your system.\n\n"
                                   "Would you like to TEMPORARILY enable internet access to download them? (approx. 300MB)\n\n"
                                   "Data privacy firewall will be paused only during the download."):
                global FIREWALL_BYPASS
                FIREWALL_BYPASS = True
                
                # Show a progress notice
                notice = ctk.CTkToplevel(self)
                notice.title("Downloading Models...")
                notice.geometry("300x150")
                notice.attributes("-topmost", True)
                ctk.CTkLabel(notice, text="Downloading AI Models...\nPlease wait, this may take a few minutes.", pady=20).pack()
                self.update()
                
                try:
                    self.backend.prepare()
                    messagebox.showinfo("Success", "Models downloaded and initialized successfully!\nFirewall is now RE-LOCKED.")
                except Exception as e2:
                    messagebox.showerror("Download Error", f"Failed to download models: {e2}")
                finally:
                    FIREWALL_BYPASS = False
                    notice.destroy()
            else:
                messagebox.showwarning("Initialization Incomplete", 
                                       "Face recognition will not work without models. "
                                       "You can manually place the models in the './models' folder to stay 100% offline.")

    def _update_readme_admin(self):
        if os.path.exists("README.md"):
            try:
                with open("README.md", "r") as f:
                    content = f.read()
                if "ADMIN ACCESS" not in content:
                    with open("README.md", "a") as f:
                        f.write("\n\n## ADMIN ACCESS\n- **Username**: admin\n- **Default Password**: admin123\n- **Features**: Full data visibility, Exports, Reports, Organization Chart, and Settings.\n")
            except: pass

    def on_closing(self):
        self.is_marking = False
        if hasattr(self, 'backend') and self.backend and self.backend.camera:
            self.backend.camera.release()
        self.destroy()

    # ── Settings ──────────────────────────────────────────────────────────────

    def load_settings(self):
        try:
            with open("settings.json") as f:
                self.settings = json.load(f)
        except Exception:
            self.settings = {"logo_path": "", "church_name": "True Jesus Church",
                             "default_area": "", "address": ""}

    def save_settings(self):
        with open("settings.json", "w") as f:
            json.dump(self.settings, f, indent=4)

    def check_auto_periodical_reports(self):
        """Automatically generates Excel reports for the PREVIOUS period on milestone days."""
        from datetime import timedelta
        today = date.today()
        last_check = self.settings.get("last_auto_report_check", "")
        if last_check == str(today):
            return
            
        yesterday = today - timedelta(days=1)
        is_monday = today.weekday() == 0
        is_1st_month = today.day == 1
        is_1st_year = today.month == 1 and today.day == 1
        
        if is_monday or is_1st_month or is_1st_year:
            try:
                from report import ReportGenerator
                rg = ReportGenerator()
                def_area = self.settings.get("default_area", "")
                
                if is_1st_year:
                    prev_year = yesterday.strftime('%Y')
                    rg.generate_periodical_excel('yearly', 'All Sessions', def_area, custom_date=prev_year)
                if is_1st_month:
                    prev_month = yesterday.strftime('%Y-%m')
                    rg.generate_periodical_excel('monthly', 'All Sessions', def_area, custom_date=prev_month)
                if is_monday:
                    prev_week = yesterday.strftime('%Y-W%W')
                    rg.generate_periodical_excel('weekly', 'All Sessions', def_area, custom_date=prev_week)
                
                print(f"[AUTO-REPORT] Periodical reports generated for previous period of {today}")
            except Exception as e:
                print(f"[AUTO-REPORT] Auto-Report Error: {e}")

        self.settings["last_auto_report_check"] = str(today)
        self.save_settings()

    # ── Sidebar ───────────────────────────────────────────────────────────────

    def init_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#FFFFFF")
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.pack_propagate(False)

        top = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        top.pack(pady=25, padx=20, fill="x")
        self._display_logo(top)
        self.title_label = ctk.CTkLabel(top, text=self.settings.get("church_name", "True Jesus Church"),
                                        font=("Arial", 16, "bold"))
        self.title_label.pack(pady=(8, 0))

        self.area_label = ctk.CTkLabel(top, text=self.settings.get("default_area", ""),
                                       font=("Arial", 13))
        self.area_label.pack(pady=(0, 8))

        # Bottom Login Frame (packed first so it remains fixed at the bottom)
        self.sidebar_bottom = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.sidebar_bottom.pack(side="bottom", fill="x", pady=20)
        
        self.fw_badge = ctk.CTkLabel(self.sidebar_bottom, text="🛡️ Firewall Active", font=("Arial", 11, "bold"),
                                     fg_color="#D1FAE5", text_color="#059669", corner_radius=6, height=28)
        self.fw_badge.pack(padx=12, pady=(0, 10), fill="x")

        self.login_btn = ctk.CTkButton(self.sidebar_bottom, text="🔐 Admin Login", font=("Arial", 12, "bold"),
                                       height=38, fg_color="#F3F4F6", text_color="#374151",
                                       hover_color="#E5E7EB", command=self.on_login_click)
        self.login_btn.pack(padx=12, fill="x")

        # Scrollable Navigation Container (prevents clipping on smaller screen heights)
        self.nav_scroll = ctk.CTkScrollableFrame(self.sidebar, fg_color="transparent")
        self.nav_scroll.pack(fill="both", expand=True, padx=5, pady=5)

        nav = [("🏠  Dashboard", "dashboard"), ("👥  Members", "members"),
               ("📜  Attendance Logs", "logs"), ("📊  Reports", "reports"),
               ("📅  Annually Report", "annually_report"),
               ("📊  Organization chart", "org_chart"),
               ("⚙  Settings", "settings"),
               ("🗄  SQL Data", "sql")]
        self.nav_buttons = {}
        
        # Pack Dashboard button
        text, key = nav[0]
        btn = ctk.CTkButton(self.nav_scroll, text=text, font=("Arial", 13), height=44,
                            anchor="w", fg_color="transparent", text_color="#333",
                            hover_color="#F0F2F5",
                            command=lambda k=key: self.show_frame(k))
        btn.pack(pady=3, padx=12, fill="x")
        self.nav_buttons[key] = btn


        
        # Pack the rest of the navigation buttons
        for text, key in nav[1:]:
            btn = ctk.CTkButton(self.nav_scroll, text=text, font=("Arial", 13), height=44,
                                anchor="w", fg_color="transparent", text_color="#333",
                                hover_color="#F0F2F5",
                                command=lambda k=key: self.show_frame(k))
            btn.pack(pady=3, padx=12, fill="x")
            self.nav_buttons[key] = btn
            
        # Initialize camera list
        self.after(1000, self.refresh_camera_list)
        
        self.refresh_sidebar_visibility()

    def refresh_sidebar_visibility(self):
        # Hide restricted buttons for public
        restricted = ["reports", "annually_report", "org_chart", "settings", "sql"]
        for key in restricted:
            if key in self.nav_buttons:
                if self.is_admin:
                    self.nav_buttons[key].pack(pady=3, padx=12, fill="x")
                else:
                    self.nav_buttons[key].pack_forget()

    def refresh_camera_list(self):
        """Detect available cameras by trying indices 0-4 and load WiFi cameras."""
        available = []
        current_id = getattr(self.backend, 'current_camera_id', 0)
        
        # Add local USB/Built-in cameras
        for i in range(5):
            if i == current_id:
                available.append(f"Camera {i}")
                continue
            cap = cv2.VideoCapture(i)
            if cap.isOpened():
                available.append(f"Camera {i}")
                cap.release()
        
        if not available and not isinstance(current_id, str):
            available = [f"Camera {current_id}"]
        
        # Sort USB cameras by index
        available.sort(key=lambda x: int(x.split()[-1]))
        
        # Add WiFi cameras from settings
        wifi_cams = self.settings.get("wifi_cameras", [])
        for url in wifi_cams:
            available.append(f"WiFi Camera: {url}")
            
        # Ensure current camera is in list if it's a string URL
        if isinstance(current_id, str) and current_id.startswith(("rtsp://", "http://", "https://")):
            display_name = f"WiFi Camera: {current_id}"
            if display_name not in available:
                available.append(display_name)
        
        self.cam_menu.configure(values=available)
        
        # Keep dropdown in sync
        if isinstance(current_id, str):
            self.cam_var.set(f"WiFi Camera: {current_id}")
        else:
            self.cam_var.set(f"Camera {current_id}")

    def on_camera_change(self, choice):
        try:
            if choice.startswith("WiFi Camera: "):
                cam_id = choice.replace("WiFi Camera: ", "").strip()
            elif choice.startswith("Camera "):
                cam_id = int(choice.split()[-1])
            else:
                cam_id = choice
            
            success = self.backend.switch_camera(cam_id)
            if not success:
                messagebox.showerror("Error", f"Failed to open {choice}")
            else:
                print(f"[INFO] Switched to {choice}")
                self.settings["last_camera_id"] = cam_id
                self.save_settings()
        except Exception as e:
            messagebox.showerror("Error", f"Camera switch failed: {e}")

    def add_wifi_camera_dialog(self):
        dialog = ctk.CTkInputDialog(
            text="Enter WiFi/IP Camera URL (RTSP/HTTP):\nExample:\nrtsp://192.168.1.100:554/live\nhttp://192.168.1.100:8080/video", 
            title="Add WiFi Camera"
        )
        url = dialog.get_input()
        if not url:
            return
        url = url.strip()
        if not url:
            return
            
        print(f"[INFO] Testing connection to WiFi Camera: {url}")
        
        # Create a brief visual feedback or directly test (with timeout or quickly)
        cap = cv2.VideoCapture(url)
        if not cap.isOpened():
            messagebox.showerror("Connection Failed", "Could not connect to the WiFi Camera URL.\nPlease check the URL and make sure the camera is connected to the same network.")
            cap.release()
            return
        cap.release()
        
        # Save to settings
        wifi_cams = self.settings.get("wifi_cameras", [])
        if url not in wifi_cams:
            wifi_cams.append(url)
            self.settings["wifi_cameras"] = wifi_cams
            self.save_settings()
            
        self.refresh_camera_list()
        
        # Select and switch to the new camera
        display_name = f"WiFi Camera: {url}"
        self.cam_var.set(display_name)
        self.on_camera_change(display_name)

    def start_auto_search(self):
        popup = ctk.CTkToplevel(self)
        popup.title("WiFi Camera Scanner")
        popup.geometry("400x180")
        popup.resizable(False, False)
        popup.transient(self)
        popup.attributes("-topmost", True)
        
        popup.update_idletasks()
        w = popup.winfo_width()
        h = popup.winfo_height()
        x = self.winfo_x() + (self.winfo_width() // 2) - (w // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (h // 2)
        popup.geometry(f"+{x}+{y}")
        popup.grab_set()
        
        lbl = ctk.CTkLabel(popup, text="🔍 Scanning network for WiFi cameras...", font=("Arial", 13, "bold"))
        lbl.pack(pady=(25, 10))
        
        progress = ctk.CTkProgressBar(popup, width=300)
        progress.pack(pady=10)
        progress.start()
        
        status_lbl = ctk.CTkLabel(popup, text="Sending discovery probes (ONVIF/SSDP)...", font=("Arial", 11), text_color="#6B7280")
        status_lbl.pack(pady=(0, 15))
        
        def run_scan():
            try:
                ips = discover_cameras()
                if not ips:
                    popup.after(0, lambda: finish_scan([]))
                    return
                    
                popup.after(0, lambda: status_lbl.configure(text=f"Found {len(ips)} IP addresses. Testing streams..."))
                
                candidates = []
                for ip in ips:
                    for p in COMMON_PATHS:
                        if p.startswith(":8080"):
                            candidates.append(f"http://{ip}{p}")
                        else:
                            candidates.append(f"rtsp://{ip}:554{p}")
                            candidates.append(f"rtsp://{ip}{p}")
                            
                working = []
                with concurrent.futures.ThreadPoolExecutor(max_workers=20) as executor:
                    res_map = executor.map(test_camera_url, candidates)
                    for r in res_map:
                        if r:
                            working.append(r)
                            
                popup.after(0, lambda: finish_scan(working))
            except Exception as e:
                popup.after(0, lambda e_err=e: messagebox.showerror("Scan Error", str(e_err)))
                popup.after(0, popup.destroy)
                
        def finish_scan(working_urls):
            progress.stop()
            popup.destroy()
            if not working_urls:
                messagebox.showinfo("Scan Complete", "No WiFi/IP cameras with active video streams were found on your local network.\n\nMake sure the camera is powered on and connected to the same network.")
                return
                
            show_scan_results_dialog(working_urls)
            
        def show_scan_results_dialog(urls):
            res_dialog = ctk.CTkToplevel(self)
            res_dialog.title("Discovered WiFi Cameras")
            res_dialog.geometry("450x220")
            res_dialog.resizable(False, False)
            res_dialog.transient(self)
            res_dialog.attributes("-topmost", True)
            
            res_dialog.update_idletasks()
            rw = res_dialog.winfo_width()
            rh = res_dialog.winfo_height()
            rx = self.winfo_x() + (self.winfo_width() // 2) - (rw // 2)
            ry = self.winfo_y() + (self.winfo_height() // 2) - (rh // 2)
            res_dialog.geometry(f"+{rx}+{ry}")
            res_dialog.grab_set()
            
            ctk.CTkLabel(res_dialog, text="🎉 Discovered WiFi Cameras!", font=("Arial", 14, "bold")).pack(pady=(15, 10))
            ctk.CTkLabel(res_dialog, text="Select a camera stream to add to your list:", font=("Arial", 12), text_color="#4B5563").pack(pady=(0, 10))
            
            selected_url = ctk.StringVar(value=urls[0])
            combo = ctk.CTkComboBox(res_dialog, values=urls, variable=selected_url, width=380, height=35)
            combo.pack(pady=10)
            
            btn_f = ctk.CTkFrame(res_dialog, fg_color="transparent")
            btn_f.pack(fill="x", pady=(15, 0), padx=20)
            
            def add_selected():
                url = selected_url.get().strip()
                if url:
                    wifi_cams = self.settings.get("wifi_cameras", [])
                    if url not in wifi_cams:
                        wifi_cams.append(url)
                        self.settings["wifi_cameras"] = wifi_cams
                        self.save_settings()
                    self.refresh_camera_list()
                    display_name = f"WiFi Camera: {url}"
                    self.cam_var.set(display_name)
                    self.on_camera_change(display_name)
                    res_dialog.destroy()
                    messagebox.showinfo("Success", f"Connected to: {url}")
            
            ctk.CTkButton(btn_f, text="Cancel", width=100, fg_color="#F3F4F6", text_color="#374151", hover_color="#E5E7EB", command=res_dialog.destroy).pack(side="left", padx=10)
            ctk.CTkButton(btn_f, text="Connect & Add", width=140, fg_color="#007BFF", hover_color="#0069D9", text_color="white", command=add_selected).pack(side="right", padx=10)
            
        threading.Thread(target=run_scan, daemon=True).start()

    def delete_wifi_camera(self):
        choice = self.cam_var.get()
        if not choice.startswith("WiFi Camera: "):
            messagebox.showwarning("Invalid Selection", "Please select a WiFi Camera from the dropdown first to remove it.")
            return
            
        url = choice.replace("WiFi Camera: ", "").strip()
        wifi_cams = self.settings.get("wifi_cameras", [])
        if url in wifi_cams:
            if messagebox.askyesno("Confirm Removal", f"Remove this WiFi Camera URL from settings?\n\n{url}"):
                wifi_cams.remove(url)
                self.settings["wifi_cameras"] = wifi_cams
                self.save_settings()
                
                # Switch back to Camera 0
                self.cam_var.set("Camera 0")
                self.on_camera_change("Camera 0")
                self.refresh_camera_list()
        
        # Update login button text
        if self.is_admin:
            self.login_btn.configure(text="🔓 Logout", fg_color="#FEE2E2", text_color="#EF4444", hover_color="#FECACA")
        else:
            self.login_btn.configure(text="🔐 Admin Login", fg_color="#F3F4F6", text_color="#374151", hover_color="#E5E7EB")

    def refresh_members_ui_visibility(self):
        # Hide restricted buttons for public on Members page
        # Check both sync_f and actions_f exist and are valid widgets
        has_sync = hasattr(self, "sync_f") and self.sync_f.winfo_exists()
        has_actions = hasattr(self, "actions_f") and self.actions_f.winfo_exists()
        
        if self.is_admin:
            if has_sync: self.sync_f.pack(side="right", anchor="ne", pady=10)
            if has_actions: self.actions_f.pack(side="right", padx=15, pady=10)
        else:
            if has_sync: self.sync_f.pack_forget()
            if has_actions: self.actions_f.pack_forget()

    def on_login_click(self):
        if self.is_admin:
            if messagebox.askyesno("Logout", "Are you sure you want to logout?"):
                self.is_admin = False
                self.refresh_sidebar_visibility()
                self.refresh_members_ui_visibility()
                self.show_frame("dashboard")
                self.refresh_member_table()
        else:
            self.show_login_popup()

    def show_login_popup(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Admin Login")
        dialog.geometry("340x420")
        dialog.attributes("-topmost", True)
        dialog.grab_set()
        
        ctk.CTkLabel(dialog, text="🔒", font=("Arial", 50)).pack(pady=(30, 10))
        ctk.CTkLabel(dialog, text="Admin Access Required", font=("Arial", 16, "bold")).pack(pady=5)
        
        ctk.CTkLabel(dialog, text="Username", font=("Arial", 12)).pack(anchor="w", padx=40, pady=(20, 0))
        user_e = ctk.CTkEntry(dialog, width=260, placeholder_text="Enter username")
        user_e.insert(0, "admin")
        user_e.pack(pady=5)
        
        ctk.CTkLabel(dialog, text="Password", font=("Arial", 12)).pack(anchor="w", padx=40, pady=(10, 0))
        pass_e = ctk.CTkEntry(dialog, width=260, show="*", placeholder_text="Enter password")
        pass_e.pack(pady=5)
        
        def do_login():
            u = user_e.get().strip()
            p = pass_e.get().strip()
            if u == "admin" and p == self.settings.get("admin_pass", "admin123"):
                self.is_admin = True
                self.refresh_sidebar_visibility()
                self.refresh_members_ui_visibility()
                self.refresh_member_table()
                dialog.destroy()
                messagebox.showinfo("Success", "Welcome back, Admin!")
            else:
                messagebox.showerror("Error", "Invalid credentials.")
        
        def forget_pass():
            dialog.destroy()
            self.show_verify_hint_popup()

        ctk.CTkButton(dialog, text="Login", width=260, height=40, font=("Arial", 13, "bold"), command=do_login).pack(pady=(30, 10))
        ctk.CTkButton(dialog, text="Forget Password?", width=200, height=28, fg_color="transparent", text_color="gray", hover_color="#F0F2F5", command=forget_pass).pack()

    def show_verify_hint_popup(self):
        v_dialog = ctk.CTkToplevel(self)
        v_dialog.title("Password Recovery")
        v_dialog.geometry("340x300")
        v_dialog.attributes("-topmost", True)
        v_dialog.grab_set()
        
        ctk.CTkLabel(v_dialog, text="🔑", font=("Arial", 40)).pack(pady=(20, 10))
        ctk.CTkLabel(v_dialog, text="Enter Security Word", font=("Arial", 14, "bold")).pack()
        ctk.CTkLabel(v_dialog, text="(Set in Admin Settings)", font=("Arial", 10), text_color="gray").pack()
        
        hint_e = ctk.CTkEntry(v_dialog, width=260, placeholder_text="Enter hint word")
        hint_e.pack(pady=20)
        
        def verify():
            h = hint_e.get().strip().lower()
            correct = self.settings.get("admin_hint", "skudai").lower()
            if h == correct:
                v_dialog.destroy()
                self.show_reset_password_dialog()
            else:
                messagebox.showerror("Error", "Incorrect Security Word.", parent=v_dialog)
                
        ctk.CTkButton(v_dialog, text="Verify", width=260, height=35, command=verify).pack()

    def show_reset_password_dialog(self):
        r_dialog = ctk.CTkToplevel(self)
        r_dialog.title("Reset Password")
        r_dialog.geometry("340x300")
        r_dialog.attributes("-topmost", True)
        r_dialog.grab_set()
        
        ctk.CTkLabel(r_dialog, text="🔐", font=("Arial", 40)).pack(pady=(20, 10))
        ctk.CTkLabel(r_dialog, text="Set New Password", font=("Arial", 14, "bold")).pack()
        
        new_p = ctk.CTkEntry(r_dialog, width=260, show="*", placeholder_text="Enter new password")
        new_p.pack(pady=20)
        
        def reset():
            p = new_p.get().strip()
            if p:
                self.settings["admin_pass"] = p
                self.save_settings()
                messagebox.showinfo("Success", "Password reset successful! Please login with your new password.")
                r_dialog.destroy()
                self.show_login_popup()
            else:
                messagebox.showwarning("Warning", "Password cannot be empty.", parent=r_dialog)
                
        ctk.CTkButton(r_dialog, text="Reset & Login", width=260, height=35, command=reset).pack()

    def show_frame(self, page_id):
        # Restriction check
        restricted = ["reports", "annually_report", "org_chart", "settings", "sql"]
        if page_id in restricted and not self.is_admin:
            messagebox.showwarning("Restricted", "Admin login required to access this page.")
            return

        frame = self.frames.get(page_id)
        if frame:
            frame.grid(row=0, column=0, sticky="nsew")
            frame.tkraise()
            self.page_title.configure(text=page_id.replace("_", " ").title())
            
            # Update nav button styles
            for key, btn in self.nav_buttons.items():
                if key == page_id:
                    btn.configure(fg_color="#F0F2F5", text_color="#007BFF")
                else:
                    btn.configure(fg_color="transparent", text_color="#333")

    def _display_logo(self, master):
        lp = self.settings.get("logo_path", "")
        if lp and not os.path.isabs(lp):
            lp = os.path.join(os.path.dirname(os.path.abspath(__file__)), lp)
        if lp and os.path.exists(lp):
            try:
                ci = ctk.CTkImage(light_image=Image.open(lp), size=(60, 60))
                ctk.CTkLabel(master, image=ci, text="").pack()
                return
            except Exception:
                pass
        ctk.CTkLabel(master, text="💒", font=("Arial", 40)).pack()

    # ── Header ────────────────────────────────────────────────────────────────

    def init_header(self):
        hdr = ctk.CTkFrame(self.main_area, height=65, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=30, pady=(15, 5))
        hdr.grid_columnconfigure(1, weight=1)

        self.page_title = ctk.CTkLabel(hdr, text="Dashboard", font=("Arial", 22, "bold"))
        self.page_title.grid(row=0, column=0, sticky="w")

        self.session_info_lbl = ctk.CTkLabel(hdr, text="● No Active Session",
                                              font=("Arial", 13, "bold"), text_color="#999")
        self.session_info_lbl.grid(row=0, column=1, padx=20)

        self.date_label = ctk.CTkLabel(hdr, text=date.today().strftime("%A, %d %B %Y"),
                                        font=("Arial", 12), text_color="gray")
        self.date_label.grid(row=0, column=2, sticky="e")

    # ── Dashboard ─────────────────────────────────────────────────────────────

    def init_dashboard(self):
        f = ctk.CTkFrame(self.container, fg_color="transparent")
        self.frames["dashboard"] = f
        f.grid_rowconfigure(1, weight=1)
        f.grid_columnconfigure(0, weight=1)

        # Stats row
        stats_row = ctk.CTkFrame(f, fg_color="transparent")
        stats_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        cards_cfg = [("Present Today",       "0",  "#28A745"),
                     ("Area Member",         "0",  "#007BFF"),
                     ("Other Area Member",   "0",  "#6610f2"),
                     ("Truth Seeker",        "0",  "#17A2B8"),
                     ("Waiting Recognition", "0",  "#DC3545"),
                     ("Area Rate %",         "0%", "#6F42C1"),
                     ("Overall Rate %",      "0%", "#FFC107")]
        self.cards = {}
        for i, (title, val, color) in enumerate(cards_cfg):
            stats_row.grid_columnconfigure(i, weight=1)
            card = ctk.CTkFrame(stats_row, fg_color="#FFFFFF", corner_radius=10)
            card.grid(row=0, column=i, padx=4, sticky="nsew")
            
            lbl_title = ctk.CTkLabel(card, text=title, font=("Arial", 9), text_color="gray")
            lbl_title.pack(pady=(10, 0))
            
            lbl_val = ctk.CTkLabel(card, text=val, font=("Arial", 20, "bold"), text_color=color)
            lbl_val.pack(pady=(0, 10))
            self.cards[title] = lbl_val

            if title == "Area Rate %":
                desc = "Area Rate % = (Total Area Members Present / Total Area Members in DB) * 100"
                Tooltip(card, desc)
                Tooltip(lbl_title, desc)
                Tooltip(lbl_val, desc)
            elif title == "Overall Rate %":
                desc = "Overall Rate % = (Total Present Today / Total Area Members in DB) * 100"
                Tooltip(card, desc)
                Tooltip(lbl_title, desc)
                Tooltip(lbl_val, desc)

        # Body: camera | side panel
        body = ctk.CTkFrame(f, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        # Camera panel
        cam_panel = ctk.CTkFrame(body, fg_color="#FFFFFF", corner_radius=10)
        cam_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        top_cam_f = ctk.CTkFrame(cam_panel, fg_color="transparent")
        top_cam_f.pack(padx=15, pady=(15, 5), fill="x")

        self.cam_label = ctk.CTkLabel(top_cam_f, text="📷  Camera Offline",
                                       width=600, height=380, fg_color="#1a1a2e",
                                       text_color="white", font=("Arial", 15))
        self.cam_label.pack(side="left")

        # Camera Selection Section (moved from sidebar)
        self.cam_section = ctk.CTkFrame(top_cam_f, fg_color="#F8F9FA", corner_radius=8)
        self.cam_section.pack(side="left", padx=(15, 0), fill="both", expand=True)

        cam_inner = ctk.CTkFrame(self.cam_section, fg_color="transparent")
        cam_inner.pack(padx=15, pady=20, fill="both", expand=True)
        
        ctk.CTkLabel(cam_inner, text="📷 CAMERA SELECTION", font=("Arial", 11, "bold"), text_color="#9CA3AF").pack(anchor="w", pady=(0, 10))
        
        self.cam_select_f = ctk.CTkFrame(cam_inner, fg_color="transparent")
        self.cam_select_f.pack(fill="x")
        
        last_cam = self.settings.get("last_camera_id", 0)
        if isinstance(last_cam, str):
            init_val = f"WiFi Camera: {last_cam}"
        else:
            init_val = f"Camera {last_cam}"

        self.cam_var = ctk.StringVar(value=init_val)
        self.cam_menu = ctk.CTkComboBox(self.cam_select_f, values=[init_val], variable=self.cam_var, command=self.on_camera_change, height=35)
        
        self.cam_refresh_btn = ctk.CTkButton(self.cam_select_f, text="🔄", font=("Arial", 16), width=40, height=35, command=self.refresh_camera_list)
        self.cam_refresh_btn.pack(side="right")
        Tooltip(self.cam_refresh_btn, "Refresh to search active camera")
        
        self.cam_menu.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        self.search_wifi_cam_btn = ctk.CTkButton(cam_inner, text="🔍 Auto Search WiFi 📷", font=("Arial", 11, "bold"), height=30, fg_color="#3B82F6", text_color="white", hover_color="#2563EB", command=self.start_auto_search)
        self.search_wifi_cam_btn.pack(fill="x", pady=(10, 0))

        self.cam_btns_f = ctk.CTkFrame(cam_inner, fg_color="transparent")
        self.cam_btns_f.pack(fill="x", pady=(10, 0))
        
        self.add_wifi_cam_btn = ctk.CTkButton(self.cam_btns_f, text="➕ Add URL", font=("Arial", 11), height=28, fg_color="#E5E7EB", text_color="#374151", hover_color="#D1D5DB", command=self.add_wifi_camera_dialog)
        
        self.del_wifi_cam_btn = ctk.CTkButton(self.cam_btns_f, text="❌", font=("Arial", 11), width=40, height=28, fg_color="#EF4444", text_color="white", hover_color="#DC2626", command=self.delete_wifi_camera)
        self.del_wifi_cam_btn.pack(side="right")
        Tooltip(self.del_wifi_cam_btn, "Remove WiFi Camera")
        
        self.add_wifi_cam_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        ctrl = ctk.CTkFrame(cam_panel, fg_color="transparent")
        ctrl.pack(pady=6)

        self.start_btn = ctk.CTkButton(ctrl, text="▶  Start", width=130, height=42,
                                        font=("Arial", 13, "bold"),
                                        fg_color="#28A745", hover_color="#218838",
                                        command=self.on_start_click)
        self.start_btn.pack(side="left", padx=8)

        self.pause_btn = ctk.CTkButton(ctrl, text="⏸  Pause", width=130, height=42,
                                        font=("Arial", 13, "bold"), fg_color="#FFC107",
                                        text_color="black", hover_color="#E0A800",
                                        state="disabled", command=self.on_pause_click)
        self.pause_btn.pack(side="left", padx=8)

        self.end_btn = ctk.CTkButton(ctrl, text="⏹  End", width=110, height=42,
                                      font=("Arial", 13, "bold"),
                                      fg_color="#DC3545", hover_color="#C82333",
                                      state="disabled", command=self.on_end_click)
        self.end_btn.pack(side="left", padx=5)

        self.resume_btn = ctk.CTkButton(ctrl, text="🔄  Resume", width=110, height=42,
                                         font=("Arial", 13, "bold"),
                                         fg_color="#6F42C1", hover_color="#5A32A3",
                                         command=self.on_resume_click)
        self.resume_btn.pack(side="left", padx=5)

        # Search and Manual Add/Remove Row
        m_ctrl = ctk.CTkFrame(cam_panel, fg_color="transparent")
        m_ctrl.pack(fill="x", padx=15, pady=(10, 0))

        ctk.CTkLabel(m_ctrl, text="CAPTURED ATTENDEES", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 10))

        self.dash_search = ctk.CTkEntry(m_ctrl, placeholder_text="Search name or code...", height=32)
        self.dash_search.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.dash_search.bind("<KeyRelease>", lambda e: self.filter_captured_list())

        self.add_man_btn = ctk.CTkButton(m_ctrl, text="+", width=40, height=32, font=("Arial", 16, "bold"), fg_color="#28A745", hover_color="#218838", command=self.manual_add_popup)
        self.add_man_btn.pack(side="left", padx=2)

        self.rem_man_btn = ctk.CTkButton(m_ctrl, text="-", width=40, height=32, font=("Arial", 16, "bold"), fg_color="#DC3545", hover_color="#C82333", command=self.manual_remove_attendee)
        self.rem_man_btn.pack(side="left", padx=2)

        self.checkin_scroll = ctk.CTkScrollableFrame(cam_panel, orientation="vertical",
                                                      height=280, fg_color="#F8F9FA")
        self.checkin_scroll.pack(fill="both", expand=True, padx=15, pady=(5, 12))

        # Side panel
        side = ctk.CTkFrame(body, fg_color="#FFFFFF", corner_radius=10)
        side.grid(row=0, column=1, sticky="nsew")
        side.grid_rowconfigure(1, weight=1)
        side.grid_rowconfigure(4, weight=1)
        side.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(side, text="Recent Activity", font=("Arial", 13, "bold")).grid(
            row=0, column=0, padx=12, pady=(12, 0), sticky="w")
        self.activity_log = ctk.CTkTextbox(side, height=170, font=("Arial", 11),
                                            fg_color="transparent")
        self.activity_log.grid(row=1, column=0, sticky="nsew", padx=10, pady=(4, 0))

        ctk.CTkFrame(side, height=2, fg_color="#E9ECEF").grid(
            row=2, column=0, sticky="ew", padx=10, pady=8)

        ctk.CTkLabel(side, text="⚠  Waiting Recognition",
                     font=("Arial", 12, "bold"), text_color="#DC3545").grid(
            row=3, column=0, padx=12, pady=(0, 4), sticky="w")

        self.waiting_scroll = ctk.CTkScrollableFrame(side, fg_color="transparent", height=200)
        self.waiting_scroll.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 12))

    # ── Members ───────────────────────────────────────────────────────────────

    def init_members_page(self):
        f = ctk.CTkFrame(self.container, fg_color="#F8F9FA", corner_radius=10)
        self.frames["members"] = f

        header_frame = ctk.CTkFrame(f, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))

        # Title/Subtitle container
        ts_f = ctk.CTkFrame(header_frame, fg_color="transparent")
        ts_f.pack(side="left")

        title_lbl = ctk.CTkLabel(ts_f, text="Member Management", font=("Arial", 28, "bold"), text_color="#1F2937")
        title_lbl.pack(anchor="w")
        sub_lbl = ctk.CTkLabel(ts_f, text="Manage church members, register new ones, and sync data between locations", font=("Arial", 14), text_color="#6B7280")
        sub_lbl.pack(anchor="w")

        # Sync Buttons (Top Right)
        self.sync_f = ctk.CTkFrame(header_frame, fg_color="transparent")
        self.sync_f.pack(side="right", anchor="ne", pady=10)
        self.sync_out_btn = ctk.CTkButton(self.sync_f, text="⬇ Sync Out", width=120, height=36, fg_color="#6366F1", hover_color="#4F46E5", font=("Arial", 11, "bold"), command=self.on_bulk_sync_output)
        self.sync_out_btn.pack(side="left", padx=5)
        self.sync_in_btn = ctk.CTkButton(self.sync_f, text="⬆ Sync In", width=120, height=36, fg_color="#8B5CF6", hover_color="#7C3AED", font=("Arial", 11, "bold"), command=self.on_bulk_sync_input)
        self.sync_in_btn.pack(side="left", padx=5)
        
        # self.refresh_members_ui_visibility() - Moved to end of method
        
        # --- NEW SUMMARY STATS BAR ---
        self.member_stats_frame = ctk.CTkFrame(f, fg_color="transparent")
        self.member_stats_frame.pack(fill="x", padx=20, pady=(15, 5))
        
        self.member_stats_labels = {}
        stat_configs = [
            ("Total DB Count", "#3B82F6"),
            ("Area Member", "#10B981"),
            ("Other Area Member", "#6366F1"),
            ("Truth Seeker", "#14B8A6")
        ]
        
        for i, (title, color) in enumerate(stat_configs):
            card = ctk.CTkFrame(self.member_stats_frame, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=8, height=80)
            card.pack(side="left", fill="both", expand=True, padx=(0 if i==0 else 10, 0))
            card.pack_propagate(False)
            
            ctk.CTkLabel(card, text=title.upper(), font=("Arial", 9, "bold"), text_color="#6B7280").pack(pady=(12, 0))
            lbl = ctk.CTkLabel(card, text="0", font=("Arial", 22, "bold"), text_color=color)
            lbl.pack(pady=(2, 10))
            self.member_stats_labels[title] = lbl
        # -----------------------------

        # Advanced Toolbar (Similar to Reports)
        toolbar = ctk.CTkFrame(f, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=8)
        toolbar.pack(fill="x", padx=20, pady=(10, 15))

        filter_f = ctk.CTkFrame(toolbar, fg_color="transparent")
        filter_f.pack(side="left", fill="x", padx=15, pady=10)

        # Search Name / ID
        s_i_f = ctk.CTkFrame(filter_f, fg_color="transparent")
        s_i_f.pack(side="left", padx=10)
        ctk.CTkLabel(s_i_f, text="SEARCH NAME / ID", font=("Arial", 10, "bold"), text_color="#9CA3AF").pack(anchor="w")
        self.member_search = ctk.CTkEntry(s_i_f, width=180, placeholder_text="e.g. John or SK-001")
        self.member_search.pack()
        self.member_search.bind("<Return>", lambda _: self.refresh_member_table())

        # Type Dropdown
        s_t_f = ctk.CTkFrame(filter_f, fg_color="transparent")
        s_t_f.pack(side="left", padx=10)
        ctk.CTkLabel(s_t_f, text="TYPE", font=("Arial", 10, "bold"), text_color="#9CA3AF").pack(anchor="w")
        self.member_type_filter = ctk.CTkComboBox(s_t_f, values=["All", "Area Member", "Other Area Member", "Truth Seeker"], width=130, command=lambda _: self.refresh_member_table())
        self.member_type_filter.pack()

        # Area Filter
        s_a_f = ctk.CTkFrame(filter_f, fg_color="transparent")
        s_a_f.pack(side="left", padx=10)
        ctk.CTkLabel(s_a_f, text="AREA", font=("Arial", 10, "bold"), text_color="#9CA3AF").pack(anchor="w")
        self.member_area_filter = ctk.CTkEntry(s_a_f, width=130, placeholder_text="e.g. Skudai")
        self.member_area_filter.pack()
        self.member_area_filter.bind("<Return>", lambda _: self.refresh_member_table())

        # Search Button
        ctk.CTkButton(filter_f, text="🔍  Search", width=100, height=36, fg_color="#007BFF", hover_color="#0069D9", font=("Arial", 12, "bold"), command=self.refresh_member_table).pack(side="left", padx=10, pady=(15, 0))

        # Add Member Button
        ctk.CTkButton(filter_f, text="+ Add Member", width=120, height=36, fg_color="#10B981", hover_color="#059669", font=("Arial", 12, "bold"), command=self.add_member_popup).pack(side="left", padx=10, pady=(15, 0))

        # Global Actions (Export)
        self.actions_f = ctk.CTkFrame(toolbar, fg_color="transparent")
        self.actions_f.pack(side="right", padx=15, pady=10)
        
        self.bulk_pdf_btn = ctk.CTkButton(self.actions_f, text="📕 PDF Export", width=160, height=36, fg_color="#FEE2E2", text_color="#EF4444", hover_color="#FCA5A5", font=("Arial", 11, "bold"), command=lambda: self.on_bulk_member_export("pdf"))
        self.bulk_pdf_btn.pack(side="top", pady=2)
        self.bulk_excel_btn = ctk.CTkButton(self.actions_f, text="📗 Excel Export", width=160, height=36, fg_color="#D1FAE5", text_color="#10B981", hover_color="#A7F3D0", font=("Arial", 11, "bold"), command=lambda: self.on_bulk_member_export("excel"))
        self.bulk_excel_btn.pack(side="top", pady=2)
        self.bulk_excel_in_btn = ctk.CTkButton(self.actions_f, text="📗 Excel Import", width=160, height=36, fg_color="#DBEAFE", text_color="#3B82F6", hover_color="#BFDBFE", font=("Arial", 11, "bold"), command=self.on_bulk_excel_import)
        self.bulk_excel_in_btn.pack(side="top", pady=2)
        
        self.refresh_members_ui_visibility()

        # Table Section
        table_container = ctk.CTkFrame(f, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=8)
        table_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        th = ctk.CTkFrame(table_container, fg_color="transparent", height=45)
        th.pack(fill="x", padx=10, pady=(10, 0))
        th.grid_columnconfigure(0, minsize=40)
        th.grid_columnconfigure(1, minsize=80) # Photo
        th.grid_columnconfigure(2, weight=3)    # Name / ID
        th.grid_columnconfigure(3, minsize=100) # Type
        th.grid_columnconfigure(4, minsize=120) # Area
        th.grid_columnconfigure(5, minsize=140) # Actions

        self.member_select_all_var = tk.BooleanVar(value=False)
        self.member_select_all_cb = ctk.CTkCheckBox(th, text="", variable=self.member_select_all_var, width=20, command=self.toggle_member_select_all)
        self.member_select_all_cb.grid(row=0, column=0, padx=(10, 0))

        headers = [("PHOTO", 1), ("NAME / MEMBER ID", 2), ("TYPE", 3), ("AREA", 4), ("ACTIONS", 5)]
        for name, col in headers:
            ctk.CTkLabel(th, text=name, font=("Arial", 11, "bold"), text_color="#9CA3AF").grid(row=0, column=col, sticky="w", padx=10)

        ctk.CTkFrame(table_container, height=1, fg_color="#E5E7EB").pack(fill="x", pady=5)

        self.member_scroll = ctk.CTkScrollableFrame(table_container, fg_color="transparent")
        self.member_scroll.pack(fill="both", expand=True, padx=5, pady=5)
        self.member_checkboxes = {}
        self.refresh_member_table()

    def toggle_member_select_all(self):
        val = self.member_select_all_var.get()
        for var in self.member_checkboxes.values():
            var.set(val)

    def refresh_member_table(self):
        for w in self.member_scroll.winfo_children():
            w.destroy()
        
        self.member_checkboxes = {}
        
        q_name = self.member_search.get().strip().lower()
        q_type = self.member_type_filter.get()
        q_area = self.member_area_filter.get().strip().lower()

        conn  = sqlite3.connect("database/attendance.db")
        query = "SELECT member_code, name, type, age_category, area, image_path, title FROM members WHERE 1=1"
        params = []
        
        if q_name:
            query += " AND (LOWER(name) LIKE ? OR LOWER(member_code) LIKE ?)"
            params.append(f"%{q_name}%"); params.append(f"%{q_name}%")
        if q_type != "All":
            query += " AND type = ?"
            params.append(q_type)
        if q_area:
            query += " AND LOWER(area) LIKE ?"
            params.append(f"%{q_area}%")
            
        query += " ORDER BY member_code DESC"
        df = pd.read_sql(query, conn, params=params)
        conn.close()

        # --- UPDATE STATS COUNTERS ---
        curr_area = (self.settings.get("default_area", "") or "").lower().strip()
        
        # Ensure area and type are strings and handle NULLs
        df['area'] = df['area'].fillna("").astype(str).str.lower().str.strip()
        df['type'] = df['type'].fillna("Area Member").astype(str)
        
        t_total = len(df)
        t_area_m  = len(df[df['type'].str.lower() == 'area member'])
        t_oth_m   = len(df[df['type'].str.lower() == 'other area member'])
        t_ts      = len(df[df['type'].str.lower() == 'truth seeker'])
        
        self.member_stats_labels["Total DB Count"].configure(text=str(t_total))
        self.member_stats_labels["Area Member"].configure(text=str(t_area_m))
        self.member_stats_labels["Other Area Member"].configure(text=str(t_oth_m))
        self.member_stats_labels["Truth Seeker"].configure(text=str(t_ts))
        # -----------------------------

        if df.empty:
            ctk.CTkLabel(self.member_scroll, text="No members found.", font=("Arial", 13), text_color="gray").pack(pady=40)
            return

        for _, row_data in df.iterrows():
            code, name, m_type, age, area, img_p, title = row_data['member_code'], row_data['name'], row_data['type'], row_data['age_category'], row_data['area'], row_data['image_path'], row_data['title']
            row = ctk.CTkFrame(self.member_scroll, fg_color="transparent", height=85)
            row.pack(fill="x", pady=0)
            row.grid_columnconfigure(0, minsize=40); row.grid_columnconfigure(1, minsize=80)
            row.grid_columnconfigure(2, weight=3); row.grid_columnconfigure(3, minsize=100)
            row.grid_columnconfigure(4, minsize=120); row.grid_columnconfigure(5, minsize=140)

            # Checkbox
            cb_var = tk.BooleanVar(value=False); self.member_checkboxes[code] = cb_var
            ctk.CTkCheckBox(row, text="", variable=cb_var, width=20).grid(row=0, column=0, padx=(10, 0))

            # Photo
            img_lbl = ctk.CTkLabel(row, text="👤", width=60, height=60, fg_color="#F3F4F6", corner_radius=30)
            img_lbl.grid(row=0, column=1, padx=10, pady=10)
            if img_p and os.path.exists(img_p):
                try:
                    pil = Image.open(img_p).resize((60, 60))
                    ci = ctk.CTkImage(light_image=pil, size=(60, 60))
                    img_lbl.configure(image=ci, text="")
                except: pass

            # Name & ID
            info_f = ctk.CTkFrame(row, fg_color="transparent")
            info_f.grid(row=0, column=2, sticky="w", padx=10)
            
            display_name = f"{title} {name}" if title else name
            ctk.CTkLabel(info_f, text=display_name, font=("Arial", 15, "bold"), text_color="#1F2937").pack(anchor="w")
            ctk.CTkLabel(info_f, text=f"ID: {code}  |  Age: {age or '--'}  |  Title: {title or '--'}", font=("Arial", 11), text_color="#6B7280").pack(anchor="w")

            # Type
            badge_color = "#EBF5FF" if m_type == "Member" else "#F0FDFA"
            badge_txt = "#2563EB" if m_type == "Member" else "#059669"
            ctk.CTkLabel(row, text=m_type, font=("Arial", 10, "bold"), fg_color=badge_color, text_color=badge_txt, corner_radius=10, width=90, height=26).grid(row=0, column=3, padx=10)

            # Area
            ctk.CTkLabel(row, text=area or "Unknown", font=("Arial", 12), text_color="#4B5563").grid(row=0, column=4, padx=10)

            # Actions
            act_f = ctk.CTkFrame(row, fg_color="transparent")
            act_f.grid(row=0, column=5, sticky="e", padx=5)
            
            if self.is_admin:
                btn_edit = ctk.CTkButton(act_f, text="✎", width=30, height=30, fg_color="transparent", text_color="#10B981", hover_color="#D1FAE5", font=("Arial", 14), command=lambda c=code: self.on_edit_member(c))
                btn_edit.pack(side="left", padx=1)
                Tooltip(btn_edit, "Edit Member Details")
    
                btn_pdf = ctk.CTkButton(act_f, text="📕", width=30, height=30, fg_color="transparent", text_color="#EF4444", hover_color="#FEE2E2", font=("Arial", 14), command=lambda c=code: self.on_individual_member_export("pdf", c))
                btn_pdf.pack(side="left", padx=1)
                Tooltip(btn_pdf, "Export Member Profile")
    
                btn_del = ctk.CTkButton(act_f, text="🗑", width=30, height=30, fg_color="transparent", text_color="#9CA3AF", hover_color="#F3F4F6", font=("Arial", 14), command=lambda c=code: self.on_delete_member(c))
                btn_del.pack(side="left", padx=1)
                Tooltip(btn_del, "Delete Member")
            else:
                ctk.CTkLabel(act_f, text="[Admin only]", font=("Arial", 10), text_color="gray").pack(side="left", padx=10)

            ctk.CTkFrame(self.member_scroll, height=1, fg_color="#F3F4F6").pack(fill="x", padx=10)

    # ── Attendance Logs ───────────────────────────────────────────────────────

    def init_logs_page(self):
        f = ctk.CTkFrame(self.container, fg_color="transparent")
        self.frames["logs"] = f

        top = ctk.CTkFrame(f, fg_color="transparent")
        top.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(top, text="Attendance History", font=("Arial", 17, "bold")).pack(side="left")
        ctk.CTkButton(top, text="🔄 Refresh", width=100,
                      command=self.refresh_logs_table).pack(side="right")

        self.logs_scroll = ctk.CTkScrollableFrame(f, fg_color="#FFFFFF", corner_radius=10)
        self.logs_scroll.pack(fill="both", expand=True, padx=4, pady=4)
        self.refresh_logs_table()

    def refresh_logs_table(self):
        # Redesigned to show by Task/Session for easier identification
        for w in self.logs_scroll.winfo_children():
            w.destroy()
        
        conn = sqlite3.connect("database/attendance.db")
        # Fetching sessions and checking for counts of 'unknown' status
        sessions = conn.execute("""
            SELECT s.id, s.title, s.date, s.start_time,
                   SUM(CASE WHEN a.status = 'unknown' THEN 1 ELSE 0 END) AS unk_count
            FROM sessions s
            JOIN attendance a ON a.session_id = s.id
            GROUP BY s.id
            ORDER BY s.date DESC, s.start_time DESC
        """).fetchall()
        conn.close()

        if not sessions:
            ctk.CTkLabel(self.logs_scroll, text="No attendance logs found yet.",
                         font=("Arial", 13), text_color="gray").pack(pady=40)
            return

        def add_session_gradually(index=0):
            if index >= len(sessions): return
            sid, title, dt, st, unk_count = sessions[index]
            
            row = ctk.CTkFrame(self.logs_scroll, fg_color="#FFFFFF", height=60, corner_radius=8, border_width=1, border_color="#E5E7EB")
            row.pack(fill="x", pady=4, padx=8)
            row.pack_propagate(False)

            info_f = ctk.CTkFrame(row, fg_color="transparent")
            info_f.pack(side="left", padx=15, pady=10)
            ctk.CTkLabel(info_f, text=f"{dt}   {str(title).upper()}", font=("Arial", 13, "bold"), text_color="#1F2937").pack(anchor="w")
            ctk.CTkLabel(info_f, text=f"Started at {str(st)[11:16]}", font=("Arial", 10), text_color="#6B7280").pack(anchor="w")

            btn_f = ctk.CTkFrame(row, fg_color="transparent")
            btn_f.pack(side="right", padx=15)

            # Red color if unknown attendees exist
            b_color = "#DC3545" if (unk_count and unk_count > 0) else "#007BFF"
            h_color = "#C82333" if (unk_count and unk_count > 0) else "#0056b3"

            ctk.CTkButton(btn_f, text="View Attendees", width=140, height=32, font=("Arial", 11, "bold"), fg_color=b_color, hover_color=h_color,
                          command=lambda s=sid: self.show_session_details_popup(s)).pack(side="left", padx=5)
            
            ctk.CTkButton(btn_f, text="X", width=32, height=32, font=("Arial", 11, "bold"), fg_color="transparent", text_color="#DC3545", border_width=1, border_color="#DC3545", hover_color="#FEE2E2",
                          command=lambda s=sid: self.delete_session_log(s)).pack(side="left")

            self.after(20, lambda: add_session_gradually(index + 1))

        add_session_gradually()

    def delete_session_log(self, session_id):
        if messagebox.askyesno("Delete Session", "Are you sure you want to permanently delete this session and all its records (including photos)?"):
            try:
                conn = sqlite3.connect("database/attendance.db")
                
                # 1. Get all image paths before deleting records
                imgs = conn.execute("SELECT record_image FROM attendance WHERE session_id=?", (session_id,)).fetchall()
                
                # 2. Delete physical files from hard drive
                print(f"[DELETE] Starting cleanup for session {session_id}...")
                deleted_count = 0
                for (img_path,) in imgs:
                    if img_path:
                        # Normalize path for Windows
                        abs_path = os.path.normpath(os.path.join(os.getcwd(), img_path))
                        if os.path.exists(abs_path):
                            try:
                                os.remove(abs_path)
                                print(f"[DELETE] Successfully removed: {abs_path}")
                                deleted_count += 1
                            except Exception as e:
                                print(f"[WARN] Failed to remove {abs_path}: {e}")
                        else:
                            print(f"[WARN] File not found for deletion: {abs_path}")

                # 3. Delete from database
                conn.execute("DELETE FROM attendance WHERE session_id=?", (session_id,))
                conn.execute("DELETE FROM sessions WHERE id=?", (session_id,))
                conn.commit()
                conn.close()
                
                self.refresh_logs_table()
                self.refresh_sessions_summary() # Force refresh the Reports tab list
                self.refresh_stats()
                
                messagebox.showinfo("Deleted", f"Session removed. {deleted_count} image files physically deleted.")
                print(f"[SESSION] Deleted session {session_id} and {deleted_count} images.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete session: {e}")
                print(f"[ERROR] Session deletion failed: {e}")

    # ── Reports ───────────────────────────────────────────────────────────────

    def init_reports_page(self):
        f = ctk.CTkFrame(self.container, fg_color="#F8F9FA", corner_radius=10)
        self.frames["reports"] = f

        header_frame = ctk.CTkFrame(f, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))

        title_lbl = ctk.CTkLabel(header_frame, text="Report Management", font=("Arial", 28, "bold"), text_color="#1F2937")
        title_lbl.pack(anchor="w")
        sub_lbl = ctk.CTkLabel(header_frame, text="Manage attendance records, view session details, and export reports", font=("Arial", 14), text_color="#6B7280")
        sub_lbl.pack(anchor="w")

        # Advanced Search Toolbar
        toolbar = ctk.CTkFrame(f, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=8)
        toolbar.pack(fill="x", padx=20, pady=(10, 15))

        filter_f = ctk.CTkFrame(toolbar, fg_color="transparent")
        filter_f.pack(side="left", fill="x", expand=True, padx=15, pady=10)

        # Search ID / Name (FIRST)
        s_i_f = ctk.CTkFrame(filter_f, fg_color="transparent")
        s_i_f.pack(side="left", padx=10)
        ctk.CTkLabel(s_i_f, text="SEARCH ID / NAME", font=("Arial", 10, "bold"), text_color="#9CA3AF").pack(anchor="w")
        self.report_search = ctk.CTkEntry(s_i_f, width=200, placeholder_text="e.g. session name")
        self.report_search.pack()
        self.report_search.bind("<Return>", lambda _: self.refresh_sessions_summary())

        # From Date (Read-only + Icon)
        f_d_f = ctk.CTkFrame(filter_f, fg_color="transparent")
        f_d_f.pack(side="left", padx=10)
        ctk.CTkLabel(f_d_f, text="FROM DATE", font=("Arial", 10, "bold"), text_color="#9CA3AF").pack(anchor="w")
        fd_row = ctk.CTkFrame(f_d_f, fg_color="transparent")
        fd_row.pack()
        self.report_from = ctk.CTkEntry(fd_row, width=100, placeholder_text="DD-MM-YYYY", state="readonly")
        self.report_from.pack(side="left")
        ctk.CTkButton(fd_row, text="📅", width=30, height=28, fg_color="#F3F4F6", text_color="#374151", hover_color="#E5E7EB", command=lambda: self.open_report_date_picker(self.report_from)).pack(side="left", padx=2)

        # To Date (Read-only + Icon)
        t_d_f = ctk.CTkFrame(filter_f, fg_color="transparent")
        t_d_f.pack(side="left", padx=10)
        ctk.CTkLabel(t_d_f, text="TO DATE", font=("Arial", 10, "bold"), text_color="#9CA3AF").pack(anchor="w")
        td_row = ctk.CTkFrame(t_d_f, fg_color="transparent")
        td_row.pack()
        self.report_to = ctk.CTkEntry(td_row, width=100, placeholder_text="DD-MM-YYYY", state="readonly")
        self.report_to.pack(side="left")
        ctk.CTkButton(td_row, text="📅", width=30, height=28, fg_color="#F3F4F6", text_color="#374151", hover_color="#E5E7EB", command=lambda: self.open_report_date_picker(self.report_to)).pack(side="left", padx=2)

        # Multi-Search Button
        ctk.CTkButton(filter_f, text="🔍  Search", width=120, height=36, fg_color="#007BFF", hover_color="#0069D9", font=("Arial", 13, "bold"), command=self.refresh_sessions_summary).pack(side="left", padx=20, pady=(15, 0))

        # Global Export Buttons
        btn_f = ctk.CTkFrame(toolbar, fg_color="transparent")
        btn_f.pack(side="right", padx=15, pady=10)
        
        btn_excel = ctk.CTkButton(btn_f, text="⬇ Excel", width=90, height=36, fg_color="#D1FAE5", text_color="#10B981", hover_color="#A7F3D0", font=("Arial", 12, "bold"), command=lambda: self._run_export_selected("excel"))
        btn_excel.pack(side="right", padx=5)

        btn_pdf = ctk.CTkButton(btn_f, text="📄 PDF", width=90, height=36, fg_color="#FEE2E2", text_color="#EF4444", hover_color="#FCA5A5", font=("Arial", 12, "bold"), command=lambda: self._run_export_selected("pdf"))
        btn_pdf.pack(side="right", padx=5)

        table_container = ctk.CTkFrame(f, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=8)
        table_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        th = ctk.CTkFrame(table_container, fg_color="transparent", height=45)
        th.pack(fill="x", padx=10, pady=(10, 0))
        th.grid_columnconfigure(0, minsize=40)
        th.grid_columnconfigure(1, weight=3)
        th.grid_columnconfigure(2, minsize=50)
        th.grid_columnconfigure(3, minsize=50)
        th.grid_columnconfigure(4, minsize=50)
        th.grid_columnconfigure(5, minsize=70)
        th.grid_columnconfigure(6, minsize=70)
        th.grid_columnconfigure(7, minsize=70)
        th.grid_columnconfigure(8, minsize=70)
        th.grid_columnconfigure(9, minsize=140)

        self.select_all_var = tk.BooleanVar(value=False)
        self.select_all_cb = ctk.CTkCheckBox(th, text="", variable=self.select_all_var, width=20, command=self.toggle_select_all)
        self.select_all_cb.grid(row=0, column=0, padx=(10, 0))

        headers = [("SESSION TITLE / DATE", 1), ("TOTAL", 2), ("AREA M.", 3), ("OTHER M.", 4), ("T. SEEKER", 5), ("AREA %", 6), ("OVERALL %", 7), ("SPEC. %", 8), ("ACTIONS", 9)]
        for name, col in headers:
            ctk.CTkLabel(th, text=name, font=("Arial", 11, "bold"), text_color="#9CA3AF").grid(row=0, column=col, sticky="w", padx=10)

        ctk.CTkFrame(table_container, height=1, fg_color="#E5E7EB").pack(fill="x", pady=5)

        self.sessions_frame = ctk.CTkScrollableFrame(table_container, fg_color="transparent")
        self.sessions_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.session_checkboxes = {}
        self.refresh_sessions_summary()

    def toggle_select_all(self):
        val = self.select_all_var.get()
        for var in self.session_checkboxes.values():
            var.set(val)

    def open_report_date_picker(self, entry):
        popup = ctk.CTkToplevel(self)
        popup.title("Select Date")
        popup.geometry("360x480")
        popup.attributes("-topmost", True)
        popup.grab_set()

        ctk.CTkLabel(popup, text="Select Search Date", font=("Arial", 15, "bold")).pack(pady=(15, 10))
        
        curr = entry.get()
        
        def on_date_click(val):
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.insert(0, val)
            entry.configure(state="readonly")
            popup.destroy()

        cal = CustomCalendar(popup, on_date_click, initial_val=curr)
        cal.pack(padx=20, pady=5, fill="both", expand=True)
        
        def clear():
            entry.configure(state="normal")
            entry.delete(0, "end")
            entry.configure(state="readonly")
            popup.destroy()

        ctk.CTkButton(popup, text="🗑  Clear Search Date", fg_color="transparent", text_color="#6B7280", font=("Arial", 11), height=24, command=clear).pack(pady=10)

    def refresh_sessions_summary(self):
        for w in self.sessions_frame.winfo_children():
            w.destroy()
        self.session_checkboxes = {}

        # --- Show loading indicator immediately so the page feels instant ---
        loading_lbl = ctk.CTkLabel(self.sessions_frame, text="⏳  Loading sessions...",
                                   font=("Arial", 13), text_color="#6B7280")
        loading_lbl.pack(pady=30)

        # Capture filter values NOW (on UI thread) before handing off
        q_name = self.report_search.get().strip().lower()
        q_from = self.report_from.get().strip()
        q_to   = self.report_to.get().strip()

        def to_sql_date(d_str):
            if not d_str: return None
            try:
                parts = d_str.split("-")
                return f"{parts[2]}-{parts[1]}-{parts[0]}"
            except: return None

        sql_from = to_sql_date(q_from)
        sql_to   = to_sql_date(q_to)

        # ── Background thread: DB fetch only ─────────────────────────────────
        def fetch_data():
            try:
                conn = sqlite3.connect("database/attendance.db", check_same_thread=False)

                query = '''
                    SELECT
                        s.id, s.title, s.date, s.target_count,
                        COUNT(a.id) as total_p,
                        SUM(CASE WHEN LOWER(a.status) = 'area member'       THEN 1 ELSE 0 END) as area_p,
                        SUM(CASE WHEN LOWER(a.status) = 'other area member' THEN 1 ELSE 0 END) as other_p,
                        SUM(CASE WHEN a.status LIKE '%truth%'               THEN 1 ELSE 0 END) as gosp_p,
                        SUM(CASE WHEN a.status = 'unknown'                  THEN 1 ELSE 0 END) as wait_p,
                        SUM(CASE WHEN LOWER(a.status) = 'area member'       THEN 1 ELSE 0 END) as area_present
                    FROM sessions s
                    LEFT JOIN attendance a ON a.session_id = s.id
                    WHERE EXISTS (SELECT 1 FROM attendance a2 WHERE a2.session_id = s.id)
                '''
                params = []
                if q_name:
                    query += " AND (LOWER(s.title) LIKE ? OR s.id LIKE ?)"
                    params += [f"%{q_name}%", f"%{q_name}%"]
                if sql_from:
                    query += " AND s.date >= ?"
                    params.append(sql_from)
                if sql_to:
                    query += " AND s.date <= ?"
                    params.append(sql_to)
                query += " GROUP BY s.id, s.title, s.date, s.target_count"
                query += " ORDER BY s.date DESC, s.start_time DESC LIMIT 150"

                sessions   = conn.execute(query, params).fetchall()
                area_total = conn.execute("SELECT COUNT(*) FROM members WHERE type='Area Member'").fetchone()[0] or 0
                conn.close()

                # Hand results back to UI thread via queue
                self.gui_queue.put(lambda s=sessions, a=area_total: render_sessions(s, a))
            except Exception as e:
                print(f"[REPORT] DB fetch error: {e}")

        # ── UI rendering: runs on main thread after fetch completes ──────────
        def render_sessions(sessions, area_total):
            # Remove loading label (safe – still on main thread)
            try:
                loading_lbl.destroy()
            except Exception:
                pass

            if not sessions:
                ctk.CTkLabel(self.sessions_frame, text="No sessions yet.",
                             font=("Arial", 13), text_color="gray").pack(pady=20)
                return

            def add_batch(index=0):
                if index >= len(sessions):
                    return
                chunk = sessions[index:index + 8]   # 8 rows per batch
                for sid, title, dt, target_count, total_p, area_p, other_p, gosp_p, wait_p, area_present in chunk:
                    row = ctk.CTkFrame(self.sessions_frame, fg_color="transparent", height=75)
                    row.pack(fill="x", pady=0)
                    row.grid_columnconfigure(0, minsize=40); row.grid_columnconfigure(1, weight=3)
                    row.grid_columnconfigure(2, minsize=60); row.grid_columnconfigure(3, minsize=60)
                    row.grid_columnconfigure(4, minsize=60); row.grid_columnconfigure(5, minsize=80)
                    row.grid_columnconfigure(6, minsize=80); row.grid_columnconfigure(7, minsize=80)
                    row.grid_columnconfigure(8, minsize=140)

                    cb_var = tk.BooleanVar(value=False)
                    self.session_checkboxes[sid] = cb_var
                    ctk.CTkCheckBox(row, text="", variable=cb_var, width=20).grid(row=0, column=0, padx=(10, 0))

                    det_f = ctk.CTkFrame(row, fg_color="transparent")
                    det_f.grid(row=0, column=1, sticky="w", padx=10, pady=10)
                    ctk.CTkLabel(det_f, text=str(title), font=("Arial", 14, "bold"), text_color="#1F2937").pack(anchor="w")
                    ctk.CTkLabel(det_f, text=str(dt or "No Date"), font=("Arial", 11), text_color="#6B7280").pack(anchor="w")

                    tp  = total_p    or 0
                    ap  = area_p     or 0
                    op  = other_p    or 0
                    gp  = gosp_p     or 0
                    apr = area_present or 0
                    tc  = target_count or 0

                    area_rate    = (apr / area_total * 100) if area_total > 0 else 0
                    overall_rate = (tp  / area_total * 100) if area_total > 0 else 0
                    spec_rate    = (tp  / tc         * 100) if tc > 0 else 0

                    ctk.CTkLabel(row, text=f"{tp}",              font=("Arial", 12, "bold"), text_color="#4B5563").grid(row=0, column=2, sticky="w", padx=10)
                    ctk.CTkLabel(row, text=f"{ap}",              font=("Arial", 12, "bold"), text_color="#4B5563").grid(row=0, column=3, sticky="w", padx=10)
                    ctk.CTkLabel(row, text=f"{op}",              font=("Arial", 12, "bold"), text_color="#4B5563").grid(row=0, column=4, sticky="w", padx=10)
                    ctk.CTkLabel(row, text=f"{gp}",              font=("Arial", 12, "bold"), text_color="#17A2B8").grid(row=0, column=5, sticky="w", padx=10)
                    ctk.CTkLabel(row, text=f"{area_rate:.1f}%",  font=("Arial", 12, "bold"), text_color="#6F42C1").grid(row=0, column=6, sticky="w", padx=10)
                    ctk.CTkLabel(row, text=f"{overall_rate:.1f}%",font=("Arial", 12, "bold"), text_color="#FFC107").grid(row=0, column=7, sticky="w", padx=10)
                    ctk.CTkLabel(row, text=f"{spec_rate:.1f}%",  font=("Arial", 12, "bold"), text_color="#E11D48").grid(row=0, column=8, sticky="w", padx=10)

                    act_f = ctk.CTkFrame(row, fg_color="transparent")
                    act_f.grid(row=0, column=9, sticky="e", padx=5)
                    ctk.CTkButton(act_f, text="✎ Details", width=55, height=28, fg_color="transparent", text_color="#8B5CF6", hover_color="#EDE9FE", font=("Arial", 10, "bold"), command=lambda s=sid: self.show_session_details_popup(s)).pack(side="left", padx=1)
                    ctk.CTkButton(act_f, text="📗 Exc",    width=40, height=28, fg_color="transparent", text_color="#10B981", hover_color="#D1FAE5", font=("Arial", 10, "bold"), command=lambda s=sid: self._run_export("excel", session_id=s)).pack(side="left", padx=1)
                    ctk.CTkButton(act_f, text="📕 PDF",    width=40, height=28, fg_color="transparent", text_color="#EF4444", hover_color="#FEE2E2", font=("Arial", 10, "bold"), command=lambda s=sid: self._run_export("pdf",   session_id=s)).pack(side="left", padx=1)

                    ctk.CTkFrame(self.sessions_frame, height=1, fg_color="#F3F4F6").pack(fill="x", padx=10)

                self.after(1, lambda: add_batch(index + 8))

            add_batch()

        threading.Thread(target=fetch_data, daemon=True).start()

    def show_session_details_popup(self, session_id):
        conn = sqlite3.connect("database/attendance.db")
        sess = conn.execute("SELECT title, date, target_count FROM sessions WHERE id=?", (session_id,)).fetchone()
        if not sess: 
            conn.close()
            return
        title_val, dt_val, target_val = sess
        
        # Calculate session-specific stats
        total_sys_m = conn.execute("SELECT COUNT(*) FROM members").fetchone()[0]
        # Denominator: Total 'Area Member' DB count
        area_total = conn.execute("SELECT COUNT(*) FROM members WHERE type='Area Member'").fetchone()[0]

        attendees = conn.execute("""
            SELECT a.id, COALESCE(m.name, a.person_name), a.status, a.check_in_time, a.record_image,
                   m.age_category, a.member_code, m.area, m.title
            FROM attendance a
            LEFT JOIN members m ON a.member_code = m.member_code
            WHERE a.session_id = ?
            ORDER BY a.check_in_time ASC
        """, (session_id,)).fetchall()
        conn.close()

        # Grouping
        area_list = [a for a in attendees if a[2].lower() == 'area member']
        other_list = [a for a in attendees if a[2].lower() == 'other area member']
        gospel_list = [a for a in attendees if 'truth' in a[2].lower()]
        unknown_list = [a for a in attendees if a[2].lower() == 'unknown']
        
        p_total = len(area_list) + len(other_list) + len(gospel_list)
        area_present = len(area_list)
            
        area_rate = (area_present / area_total * 100) if area_total > 0 else 0
        overall_rate = (p_total / area_total * 100) if area_total > 0 else 0
        special_rate = (p_total / target_val * 100) if (target_val and target_val > 0) else 0

        popup = ctk.CTkToplevel(self)
        popup.title(f"Attendance Report - Session {session_id}")
        popup.geometry("900x800")
        popup.attributes("-topmost", True)
        popup.grab_set()

        # Header - Report Title & Edit
        hdr = ctk.CTkFrame(popup, fg_color="transparent")
        hdr.pack(fill="x", padx=30, pady=(20, 10))
        
        title_f = ctk.CTkFrame(hdr, fg_color="transparent")
        title_f.pack(side="left")
        ctk.CTkLabel(title_f, text="SESSION ATTENDANCE REPORT", font=("Arial", 22, "bold"), text_color="#111827").pack(anchor="w")
        ctk.CTkLabel(title_f, text=f"Generated on {dt_val}", font=("Arial", 12), text_color="#6B7280").pack(anchor="w")

        edit_f = ctk.CTkFrame(hdr, fg_color="transparent")
        edit_f.pack(side="right")
        ctk.CTkLabel(edit_f, text="Edit Title: ", font=("Arial", 11, "bold")).pack(side="left")
        title_e = ctk.CTkEntry(edit_f, width=160, height=32)
        title_e.insert(0, title_val)
        title_e.pack(side="left", padx=5)

        ctk.CTkLabel(edit_f, text="Target: ", font=("Arial", 11, "bold")).pack(side="left", padx=(10, 0))
        target_e = ctk.CTkEntry(edit_f, width=60, height=32)
        target_e.insert(0, str(target_val or 0))
        target_e.pack(side="left", padx=5)

        def save_session_data():
            nt = title_e.get().strip()
            tc = target_e.get().strip()
            try:
                tc_int = int(tc) if tc else 0
            except:
                tc_int = 0
            
            if nt:
                c = sqlite3.connect("database/attendance.db")
                c.execute("UPDATE sessions SET title=?, target_count=? WHERE id=?", (nt, tc_int, session_id))
                c.commit()
                c.close()
                popup.destroy()
                self.show_session_details_popup(session_id)
                self.refresh_sessions_summary()
                messagebox.showinfo("Success", "Session data updated.")

        ctk.CTkButton(edit_f, text="Update", width=60, height=32, command=save_session_data).pack(side="left", padx=5)

        # Stats Cards
        card_row = ctk.CTkFrame(popup, fg_color="transparent")
        card_row.pack(fill="x", padx=30, pady=10)
        
        stats = [
            ("PRESENT TODAY", str(p_total), "#28A745"),
            ("AREA MEMBER", str(len(area_list)), "#007BFF"),
            ("OTHER AREA MEMBER", str(len(other_list)), "#6610f2"),
            ("TRUTH SEEKER", str(len(gospel_list)), "#17A2B8"),
            ("AREA RATE%", f"{area_rate:.1f}%", "#6F42C1"),
            ("OVERALL RATE%", f"{overall_rate:.1f}%", "#FFC107"),
            ("SPECIAL RATE%", f"{special_rate:.1f}%", "#E11D48")
        ]

        for i, (l, v, c) in enumerate(stats):
            card_row.grid_columnconfigure(i, weight=1)
            cf = ctk.CTkFrame(card_row, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=10)
            cf.grid(row=0, column=i, padx=5, sticky="nsew")
            ctk.CTkLabel(cf, text=l, font=("Arial", 10, "bold"), text_color="#6B7280").pack(pady=(10, 0))
            ctk.CTkLabel(cf, text=v, font=("Arial", 22, "bold"), text_color=c).pack(pady=(0, 10))

        # Main List Content
        sc = ctk.CTkScrollableFrame(popup, fg_color="#F9FAFB", corner_radius=12)
        sc.pack(fill="both", expand=True, padx=30, pady=10)

        def render_section(title, attendees_subset):
            if not attendees_subset: return
            
            ctk.CTkLabel(sc, text=title.upper(), font=("Arial", 14, "bold"), text_color="#374151").pack(anchor="w", padx=10, pady=(20, 5))
            ctk.CTkFrame(sc, height=2, fg_color="#E5E7EB").pack(fill="x", padx=10, pady=(0, 10))

            for aid, name, status, ts, img_p, age, m_code, area, title in attendees_subset:
                row = ctk.CTkFrame(sc, fg_color="#FFFFFF", corner_radius=8, height=80, border_width=1, border_color="#F3F4F6")
                row.pack(fill="x", pady=4, padx=5)
                row.pack_propagate(False)

                # Profile Image
                img_lbl = ctk.CTkLabel(row, text="📷", width=65, height=65, fg_color="#F3F4F6", corner_radius=6)
                img_lbl.pack(side="left", padx=10, pady=7)
                if img_p and os.path.exists(img_p):
                    try:
                        pil = Image.open(img_p).resize((65, 65))
                        ci = ctk.CTkImage(light_image=pil, size=(65, 65))
                        img_lbl.configure(image=ci, text="")
                    except: pass

                # Text Info
                info_f = ctk.CTkFrame(row, fg_color="transparent")
                info_f.pack(side="left", padx=10, fill="y", pady=10)
                ctk.CTkLabel(info_f, text=name, font=("Arial", 15, "bold"), text_color="#111827").pack(anchor="w")
                det_txt = f"Age: {age or '--'}  |  Title: {title or '--'}  |  Area: {area or 'Unknown'}  |  Time: {str(ts)[11:16]}"
                ctk.CTkLabel(info_f, text=det_txt, font=("Arial", 11), text_color="#6B7280").pack(anchor="w")

                # Action Links (Identify if Unknown)
                if status.lower() == 'unknown':
                    ctk.CTkButton(row, text="Identify", width=80, height=30, fg_color="#FFC107", text_color="black", font=("Arial", 11, "bold"), 
                                  command=lambda aid=aid, ip=img_p: [popup.destroy(), self.identify_unknown_popup(aid, ip)]).pack(side="right", padx=15)
                else:
                    ctk.CTkLabel(row, text=status.capitalize(), font=("Arial", 10, "bold"), text_color="#007BFF", fg_color="#EBF5FF", corner_radius=10, width=80, height=26).pack(side="right", padx=15)

        render_section("Area Member Present", area_list)
        render_section("Other Area Member Present", other_list)
        render_section("Truth Seeker Present", gospel_list)
        render_section("Unidentified Individuals (Waiting Identification)", unknown_list)

        # Footer Export
        ftr = ctk.CTkFrame(popup, fg_color="transparent")
        ftr.pack(fill="x", padx=30, pady=(10, 20))
        
        ctk.CTkLabel(ftr, text="Export this session:", font=("Arial", 12, "bold"), text_color="#4B5563").pack(side="left")
        
        btn_excel = ctk.CTkButton(ftr, text="⬇ Download Excel", width=140, height=38, fg_color="#D1FAE5", text_color="#10B981", hover_color="#A7F3D0", font=("Arial", 12, "bold"), 
                                  command=lambda: self._run_export("excel", session_id=session_id, parent=popup))
        btn_excel.pack(side="left", padx=10)

        btn_pdf = ctk.CTkButton(ftr, text="📄 Download PDF", width=140, height=38, fg_color="#FEE2E2", text_color="#EF4444", hover_color="#FCA5A5", font=("Arial", 12, "bold"), 
                                command=lambda: self._run_export("pdf", session_id=session_id, parent=popup))
        btn_pdf.pack(side="left")

    def _run_export_selected(self, kind):
        selected_ids = [sid for sid, var in self.session_checkboxes.items() if var.get()]
        if not selected_ids:
            messagebox.showwarning("Select", "No sessions ticked.")
            return
        self._run_export(kind, session_id=selected_ids, summary=True)

    def on_bulk_sync_output(self):
        """Export selected members to a zip file for syncing to another PC."""
        selected_ids = [code for code, var in self.member_checkboxes.items() if var.get()]
        if not selected_ids:
            messagebox.showwarning("Select", "No members selected for sync.")
            return

        # ── Field Selection Dialog ──
        dialog = ctk.CTkToplevel(self)
        dialog.title("Sync Out - Data Selection")
        dialog.geometry("380x620")
        dialog.attributes("-topmost", True)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text="📦 Sync Out Data Selection", font=("Arial", 16, "bold")).pack(pady=(20, 10))
        ctk.CTkLabel(dialog, text="Tick the info you want to sync to other PC:", font=("Arial", 11), text_color="gray").pack(pady=(0, 15))
        
        scroll = ctk.CTkScrollableFrame(dialog, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=10)

        fields_cfg = [
            ("📸 Profile Photo", "photo"),
            ("👤 Name", "name"),
            ("🏅 Title", "title"),
            ("📊 Age Category", "age_category"),
            ("🏷 Member Type", "type"),
            ("📍 Area", "area"),
            ("📅 Date of Birth", "dob"),
            ("💧 Baptism Date", "baptism_date"),
            ("🏠 Address", "address"),
            ("📧 Email", "email"),
            ("📞 Phone", "phone"),
            ("🕊 Holy Spirit", "has_holy_spirit"),
            ("📝 Remark", "remark")
        ]
        
        field_vars = {}
        for text, key in fields_cfg:
            # Default: photo, name, age_category are pre-ticked
            v = tk.BooleanVar(value=(key in ["photo", "name", "age_category"]))
            cb = ctk.CTkCheckBox(scroll, text=text, variable=v, font=("Arial", 12))
            cb.pack(pady=6, padx=20, anchor="w")
            field_vars[key] = v

        def do_export():
            chosen_fields = [k for k, v in field_vars.items() if v.get()]
            if not chosen_fields:
                messagebox.showwarning("Select", "Please select at least one field to export.", parent=dialog)
                return
            
            dialog.destroy()
            
            from tkinter import filedialog
            path = filedialog.asksaveasfilename(
                defaultextension=".zip",
                filetypes=[("Zip files", "*.zip")],
                initialfile=f"Member_Sync_{date.today().strftime('%Y%m%d')}.zip",
                title="Export Sync File"
            )
            if not path: return
            
            ok = self.backend.bulk_export_archive(selected_ids, path, fields=chosen_fields)
            if ok:
                messagebox.showinfo("Success", f"Sync file created successfully!\nPath: {path}")

        ctk.CTkButton(dialog, text="Proceed to Sync Out", height=40, font=("Arial", 13, "bold"), command=do_export).pack(pady=20, padx=40, fill="x")

    def on_bulk_sync_input(self):
        """Import members from a zip file."""
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            filetypes=[("Zip files", "*.zip")],
            title="Import Sync File"
        )
        if not path: return
        
        ok, msg = self.backend.bulk_import_archive(path)
        if ok:
            messagebox.showinfo("Import Success", msg)
            self.refresh_member_table()
        else:
            messagebox.showerror("Import Failed", msg)

    def on_bulk_excel_import(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")],
            title="Import Excel File"
        )
        if not path: return
        
        prefix = self.settings.get("member_prefix", "TJC")
        ok, msg = self.backend.bulk_import_excel(path, prefix=prefix)
        if ok:
            messagebox.showinfo("Excel Import", msg)
            self.refresh_member_table()
        else:
            messagebox.showerror("Import Failed", msg)

    def on_bulk_member_export(self, kind):
        selected_ids = [code for code, var in self.member_checkboxes.items() if var.get()]
        if not selected_ids:
            messagebox.showwarning("Select", "No members selected.")
            return
        self._run_member_export(kind, member_ids=selected_ids)

    def on_individual_member_export(self, kind, member_id):
        self._run_member_export(kind, member_ids=[member_id])

    def _run_member_export(self, kind, member_ids):
        try:
            ext = ".xlsx" if kind == "excel" else ".pdf"
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Member_Report_{timestamp}{ext}"
            out_p = os.path.join(downloads_path, filename)
            
            if kind == "excel":
                self.reporter.generate_member_excel(member_ids, out_p)
            else:
                self.reporter.generate_member_pdf(member_ids, out_p, self.settings)
                
            messagebox.showinfo("Export Exported", f"Report saved to:\n{out_p}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

    def _run_export(self, kind, session_id=None, parent=None, summary=False):
        try:
            ext = ".xlsx" if kind == "excel" else ".pdf"
            
            # Find the user's Downloads folder
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            if not os.path.exists(downloads_path):
                # Fallback if Downloads folder doesn't exist for some reason
                downloads_path = os.path.abspath("reports")
                os.makedirs(downloads_path, exist_ok=True)

            # Generate filename automatically
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Attendance_Summary_{timestamp}{ext}" if summary else f"Attendance_Detailed_{timestamp}{ext}"
            p = os.path.join(downloads_path, filename)

            def_area = self.settings.get("default_area", "")

            if kind == "excel":
                self.reporter.generate_excel(session_ids=session_id, out_path=p, summary=summary, default_area=def_area)
            else:
                self.reporter.generate_pdf(session_ids=session_id, out_path=p, summary=summary, 
                                           default_area=def_area, settings=self.settings)
                
            messagebox.showinfo("Export Successful", f"File has been downloaded to:\n{p}", parent=parent)
        except Exception as e:
            messagebox.showerror("Export Failed", str(e), parent=parent)

    # ── Organization Chart ────────────────────────────────────────────────────

    def init_annually_report_page(self):
        f = ctk.CTkFrame(self.container, fg_color="#F8F9FA", corner_radius=10)
        self.frames["annually_report"] = f
        
        header = ctk.CTkFrame(f, fg_color="transparent")
        header.pack(fill="x", padx=30, pady=(30, 20))
        
        ctk.CTkLabel(header, text="ANNUALLY & PERIODICAL REPORTS", font=("Arial", 24, "bold"), text_color="#111827").pack(anchor="w")
        ctk.CTkLabel(header, text="Detailed attendance analytics for Friday and Saturday Seminars", font=("Arial", 14), text_color="#6B7280").pack(anchor="w")

        # Top Level: Seminar Type Selection
        type_frame = ctk.CTkFrame(f, fg_color="transparent")
        type_frame.pack(fill="x", padx=30, pady=(0, 10))
        
        self.report_type_var = ctk.StringVar(value="All Sessions")
        type_sel = ctk.CTkSegmentedButton(type_frame, values=["Friday Seminar", "Saturday Seminar", "Fri & Sat", "Other Sessions", "All Sessions"], 
                                         variable=self.report_type_var, height=40, font=("Arial", 13, "bold"),
                                         command=lambda _: self.refresh_annually_report())
        type_sel.pack(side="left", fill="x", expand=True)

        # Export Buttons
        btn_frame = ctk.CTkFrame(type_frame, fg_color="transparent")
        btn_frame.pack(side="right", padx=(20, 0))
        
        ctk.CTkButton(btn_frame, text="📊 Excel", width=100, fg_color="#28A745", hover_color="#218838",
                      command=lambda: self.export_annually_report("excel")).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="📄 PDF", width=100, fg_color="#DC3545", hover_color="#C82333",
                      command=lambda: self.export_annually_report("pdf")).pack(side="left", padx=5)

        # Sub Level: Period Selection & Search
        period_frame = ctk.CTkFrame(f, fg_color="transparent")
        period_frame.pack(fill="x", padx=30, pady=(0, 20))
        
        self.report_period_var = ctk.StringVar(value="Monthly")
        period_sel = ctk.CTkSegmentedButton(period_frame, values=["Weekly", "Monthly", "Yearly"], 
                                           variable=self.report_period_var, height=35,
                                           command=lambda _: self.refresh_annually_report())
        period_sel.pack(side="left")

        # Date Search UI
        search_f = ctk.CTkFrame(period_frame, fg_color="transparent")
        search_f.pack(side="right")
        
        ctk.CTkLabel(search_f, text="From:", font=("Arial", 12, "bold")).pack(side="left", padx=5)
        self.report_date_start = DateEntry(search_f, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd',
                                           weekendbackground='white', weekendforeground='black')
        self.report_date_start.pack(side="left", padx=5)
        
        ctk.CTkLabel(search_f, text="To:", font=("Arial", 12, "bold")).pack(side="left", padx=5)
        self.report_date_end = DateEntry(search_f, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd',
                                         weekendbackground='white', weekendforeground='black')
        self.report_date_end.pack(side="left", padx=5)
        
        ctk.CTkButton(search_f, text="Search", width=60, command=self.refresh_annually_report).pack(side="left", padx=5)
        ctk.CTkButton(search_f, text="Clear", width=60, fg_color="#6B7280", hover_color="#4B5563", 
                      command=self.clear_report_date_search).pack(side="left", padx=5)

        # Content Area
        self.annually_table_container = ctk.CTkFrame(f, fg_color="transparent")
        self.annually_table_container.pack(fill="both", expand=True, padx=30)
        
        # Initial load
        self.clear_report_date_search() # clear will init the proper dates and refresh

    def clear_report_date_search(self):
        # Set to wide range by default or empty? If tkcalendar, empty is hard, so set to current year start/end
        today = date.today()
        self.report_date_start.set_date(date(today.year, 1, 1))
        self.report_date_end.set_date(date(today.year, 12, 31))
        self.refresh_annually_report()

    def refresh_annually_report(self):
        for w in self.annually_table_container.winfo_children(): w.destroy()
        
        p_type = self.report_period_var.get().lower()
        s_type = self.report_type_var.get()
        if s_type == "Other Sessions": s_type = "Other"
        start_date = self.report_date_start.get_date().strftime("%Y-%m-%d")
        end_date = self.report_date_end.get_date().strftime("%Y-%m-%d")
        
        def_area = self.settings.get("default_area", "").strip()
        df = self.backend.get_periodical_stats(p_type, seminar_filter=s_type, default_area=def_area, start_date=start_date, end_date=end_date)
        
        # Table Header
        table_f = ctk.CTkScrollableFrame(self.annually_table_container, fg_color="#FFFFFF", corner_radius=10, border_width=1, border_color="#E5E7EB")
        table_f.pack(fill="both", expand=True)
        
        headers = ["PERIOD", "PRESENT", "BRO / SIS", "AREA / OTHER / T.SEEKER", "AREA RATE %", "OVERALL RATE %", "DOWNLOAD"]
        h_f = ctk.CTkFrame(table_f, fg_color="#F9FAFB", height=45)
        h_f.pack(fill="x", pady=(0, 5))
        relx_map = [0.03, 0.15, 0.26, 0.40, 0.60, 0.72, 0.84]
        for i, h in enumerate(headers):
            ctk.CTkLabel(h_f, text=h, font=("Arial", 10, "bold"), text_color="#4B5563").place(relx=relx_map[i], rely=0.5, anchor="w")
        
        if df.empty:
            ctk.CTkLabel(table_f, text="No records found for the selected filters.", font=("Arial", 13), pady=40).pack()
            return

        # Map display columns
        color = "#2563EB" if s_type == "Friday Seminar" else ("#D97706" if s_type == "Saturday Seminar" else "#059669")

        # Data Rows
        for _, row in df.iterrows():
            if row['present'] == 0: continue
            p_val = str(row['period'])

            r_f = ctk.CTkFrame(table_f, fg_color="transparent", height=45)
            r_f.pack(fill="x")
            
            ctk.CTkLabel(r_f, text=p_val, font=("Arial", 12, "bold"), text_color="#111827").place(relx=0.03, rely=0.5, anchor="w")
            ctk.CTkLabel(r_f, text=str(int(row['present'])), font=("Arial", 12)).place(relx=0.15, rely=0.5, anchor="w")
            bs_txt = f"{int(row['bro'])} / {int(row['sis'])}"
            ctk.CTkLabel(r_f, text=bs_txt, font=("Arial", 12)).place(relx=0.26, rely=0.5, anchor="w")
            aot_txt = f"{int(row.get('area_present',0))} / {int(row.get('other_mbr',0))} / {int(row.get('ts',0))}"
            ctk.CTkLabel(r_f, text=aot_txt, font=("Arial", 12)).place(relx=0.40, rely=0.5, anchor="w")
            ctk.CTkLabel(r_f, text=f"{row['area_rate']:.1f}%", font=("Arial", 12, "bold"), text_color="#2563EB").place(relx=0.60, rely=0.5, anchor="w")
            ctk.CTkLabel(r_f, text=f"{row['overall_rate']:.1f}%", font=("Arial", 12, "bold"), text_color=color).place(relx=0.72, rely=0.5, anchor="w")
            
            # Action Buttons
            act_f = ctk.CTkFrame(r_f, fg_color="transparent")
            act_f.place(relx=0.84, rely=0.5, anchor="w")
            
            ctk.CTkButton(act_f, text="📊", width=30, height=28, fg_color="#28A745", hover_color="#218838",
                          command=lambda p=p_val: self.export_individual_period(p, "excel")).pack(side="left", padx=2)
            ctk.CTkButton(act_f, text="📄", width=30, height=28, fg_color="#DC3545", hover_color="#C82333",
                          command=lambda p=p_val: self.export_individual_period(p, "pdf")).pack(side="left", padx=2)

            ctk.CTkFrame(table_f, fg_color="#F3F4F6", height=1).pack(fill="x", padx=10)

        # Summary Average Row
        avg_f = ctk.CTkFrame(self.annually_table_container, fg_color="#F3F4F6", height=60, corner_radius=10)
        avg_f.pack(fill="x", pady=(10, 0))
        
        if not df.empty:
            a_pres = df['present'].mean()
            a_bro = df['bro'].mean()
            a_sis = df['sis'].mean()
            a_area_p = df['area_present'].mean() if 'area_present' in df else 0
            a_other = df['other_mbr'].mean() if 'other_mbr' in df else 0
            a_ts = df['ts'].mean() if 'ts' in df else 0
            a_area = df['area_rate'].mean()
            a_over = df['overall_rate'].mean()
            
            ctk.CTkLabel(avg_f, text="AVERAGE:", font=("Arial", 12, "bold")).place(relx=0.03, rely=0.5, anchor="w")
            ctk.CTkLabel(avg_f, text=f"{a_pres:.1f}", font=("Arial", 12, "bold")).place(relx=0.15, rely=0.5, anchor="w")
            ctk.CTkLabel(avg_f, text=f"{a_bro:.1f} / {a_sis:.1f}", font=("Arial", 11)).place(relx=0.26, rely=0.5, anchor="w")
            ctk.CTkLabel(avg_f, text=f"{a_area_p:.1f} / {a_other:.1f} / {a_ts:.1f}", font=("Arial", 11)).place(relx=0.40, rely=0.5, anchor="w")
            ctk.CTkLabel(avg_f, text=f"{a_area:.1f}%", font=("Arial", 12, "bold"), text_color="#2563EB").place(relx=0.60, rely=0.5, anchor="w")
            ctk.CTkLabel(avg_f, text=f"{a_over:.1f}%", font=("Arial", 12, "bold"), text_color=color).place(relx=0.72, rely=0.5, anchor="w")

    def export_individual_period(self, period_str, kind):
        """Downloads a detailed report for every session within the clicked period."""
        try:
            p_type = self.report_period_var.get().lower()
            s_type = self.report_type_var.get()
            def_area = self.settings.get("default_area", "")
            
            from report import ReportGenerator
            rg = ReportGenerator()
            path = rg.generate_individual_period_report(period_str, p_type, s_type, def_area, kind)
            
            if path:
                messagebox.showinfo("Download Complete", f"Individual report for {period_str} saved to:\n{os.path.abspath(path)}")
            else:
                messagebox.showwarning("Empty Report", f"No detailed session records found for {period_str}.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def export_annually_report(self, kind):
        try:
            p_type = self.report_period_var.get().lower()
            s_type = self.report_type_var.get()
            if s_type == "Other Sessions": s_type = "Other"
            start_date = self.report_date_start.get_date().strftime("%Y-%m-%d")
            end_date = self.report_date_end.get_date().strftime("%Y-%m-%d")
            
            def_area = self.settings.get("default_area", "")
            from report import ReportGenerator
            rg = ReportGenerator()
            
            if kind == "excel":
                path = rg.generate_periodical_excel(p_type, s_type, def_area, start_date=start_date, end_date=end_date)
            else:
                path = rg.generate_periodical_pdf(p_type, s_type, def_area, start_date=start_date, end_date=end_date)
                
            if path:
                messagebox.showinfo("Export Successful", f"Report saved to:\n{os.path.abspath(path)}")
            else:
                messagebox.showwarning("Export Failed", "No data available for export.")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
        

    def init_org_chart_page(self):
        f = ctk.CTkFrame(self.container, fg_color="#F8F9FA", corner_radius=10)
        self.frames["org_chart"] = f

        header_frame = ctk.CTkFrame(f, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))

        title_lbl = ctk.CTkLabel(header_frame, text="Organization Chart Management", font=("Arial", 28, "bold"), text_color="#1F2937")
        title_lbl.pack(anchor="w")
        sub_lbl = ctk.CTkLabel(header_frame, text="Manage and view the church's organizational hierarchy by year", font=("Arial", 14), text_color="#6B7280")
        sub_lbl.pack(anchor="w")

        # Toolbar
        toolbar = ctk.CTkFrame(f, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=8)
        toolbar.pack(fill="x", padx=20, pady=(10, 15))

        ctk.CTkButton(toolbar, text="+ Create New Chart", width=180, height=40, fg_color="#10B981", hover_color="#059669", font=("Arial", 13, "bold"), command=self.add_org_chart_popup).pack(side="left", padx=15, pady=10)

        # Table Section
        table_container = ctk.CTkFrame(f, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=8)
        table_container.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        th = ctk.CTkFrame(table_container, fg_color="transparent", height=45)
        th.pack(fill="x", padx=10, pady=(10, 0))
        th.grid_columnconfigure(0, minsize=100) # Year
        th.grid_columnconfigure(1, weight=1)    # Title
        th.grid_columnconfigure(2, minsize=180) # Actions

        headers = [("YEAR", 0), ("TITLE", 1), ("ACTIONS", 2)]
        for name, col in headers:
            ctk.CTkLabel(th, text=name, font=("Arial", 11, "bold"), text_color="#9CA3AF").grid(row=0, column=col, sticky="w" if col < 2 else "e", padx=15)

        ctk.CTkFrame(table_container, height=1, fg_color="#E5E7EB").pack(fill="x", pady=5)

        self.org_chart_scroll = ctk.CTkScrollableFrame(table_container, fg_color="transparent")
        self.org_chart_scroll.pack(fill="both", expand=True, padx=5, pady=5)
        self.refresh_org_chart_table()

    def refresh_org_chart_table(self):
        for w in self.org_chart_scroll.winfo_children():
            w.destroy()
        
        conn = sqlite3.connect("database/attendance.db")
        rows = conn.execute("SELECT id, title, year FROM org_charts ORDER BY year DESC, created_at DESC").fetchall()
        conn.close()

        if not rows:
            ctk.CTkLabel(self.org_chart_scroll, text="No organization charts found.", font=("Arial", 13), text_color="gray").pack(pady=40)
            return

        for cid, title, year in rows:
            row = ctk.CTkFrame(self.org_chart_scroll, fg_color="transparent", height=60)
            row.pack(fill="x", pady=0)
            row.grid_columnconfigure(0, minsize=100); row.grid_columnconfigure(1, weight=1); row.grid_columnconfigure(2, minsize=180)

            ctk.CTkLabel(row, text=str(year), font=("Arial", 14, "bold")).grid(row=0, column=0, padx=15, sticky="w")
            ctk.CTkLabel(row, text=title, font=("Arial", 14)).grid(row=0, column=1, padx=15, sticky="w")

            act_f = ctk.CTkFrame(row, fg_color="transparent")
            act_f.grid(row=0, column=2, sticky="e", padx=15)
            
            # View (New)
            btn_view = ctk.CTkButton(act_f, text="👁", width=30, height=30, fg_color="transparent", text_color="#3B82F6", hover_color="#DBEAFE", font=("Arial", 14), command=lambda c=cid: self.view_org_chart_popup(c))
            btn_view.pack(side="left", padx=1)
            Tooltip(btn_view, "View Chart Preview")

            # Edit
            btn_edit = ctk.CTkButton(act_f, text="✎", width=30, height=30, fg_color="transparent", text_color="#10B981", hover_color="#D1FAE5", font=("Arial", 14), command=lambda c=cid: self.add_org_chart_popup(c))
            btn_edit.pack(side="left", padx=1)
            Tooltip(btn_edit, "Edit Chart")

            # PDF
            btn_pdf = ctk.CTkButton(act_f, text="📕", width=30, height=30, fg_color="transparent", text_color="#EF4444", hover_color="#FEE2E2", font=("Arial", 14), command=lambda c=cid: self.on_export_org_chart("pdf", c))
            btn_pdf.pack(side="left", padx=1)
            Tooltip(btn_pdf, "Export to PDF")

            # Excel
            btn_excel = ctk.CTkButton(act_f, text="📗", width=30, height=30, fg_color="transparent", text_color="#059669", hover_color="#D1FAE5", font=("Arial", 14), command=lambda c=cid: self.on_export_org_chart("excel", c))
            btn_excel.pack(side="left", padx=1)
            Tooltip(btn_excel, "Export to Excel")

            # Delete
            btn_del = ctk.CTkButton(act_f, text="🗑", width=30, height=30, fg_color="transparent", text_color="#EF4444", hover_color="#FEE2E2", font=("Arial", 14), command=lambda c=cid: self.delete_org_chart(c))
            btn_del.pack(side="left", padx=1)
            Tooltip(btn_del, "Delete Chart")

            ctk.CTkFrame(self.org_chart_scroll, height=1, fg_color="#F3F4F6").pack(fill="x", padx=10)

    def delete_org_chart(self, chart_id):
        if messagebox.askyesno("Confirm Delete", "Permanently delete this organization chart?"):
            conn = sqlite3.connect("database/attendance.db")
            conn.execute("DELETE FROM org_chart_roles WHERE chart_id=?", (chart_id,))
            conn.execute("DELETE FROM org_charts WHERE id=?", (chart_id,))
            conn.commit()
            conn.close()
            self.refresh_org_chart_table()

    def on_export_org_chart(self, kind, chart_id):
        try:
            ext = ".xlsx" if kind == "excel" else ".pdf"
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Organization_Chart_{timestamp}{ext}"
            out_p = os.path.join(downloads_path, filename)
            
            if kind == "excel":
                self.reporter.generate_org_chart_excel(chart_id, out_p)
            else:
                self.reporter.generate_org_chart_pdf(chart_id, out_p, self.settings)
                
            messagebox.showinfo("Export Successful", f"Chart saved to:\n{out_p}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

    def view_org_chart_popup(self, chart_id):
        popup = ctk.CTkToplevel(self)
        popup.title("Organization Chart Preview")
        popup.geometry("900x700")
        popup.attributes("-topmost", True)

        conn = sqlite3.connect("database/attendance.db")
        chart_info = conn.execute("SELECT title, year FROM org_charts WHERE id=?", (chart_id,)).fetchone()
        roles = conn.execute("""
            SELECT r.id, r.parent_role_id, r.role_name, r.member_code, m.name, m.image_path
            FROM org_chart_roles r
            LEFT JOIN members m ON r.member_code = m.member_code
            WHERE r.chart_id = ?
        """, (chart_id,)).fetchall()
        conn.close()

        if not chart_info: return

        # Header in popup
        hdr = ctk.CTkFrame(popup, fg_color="#F8F9FA", height=60)
        hdr.pack(fill="x")
        ctk.CTkLabel(hdr, text=f"Preview: {chart_info[0]} ({chart_info[1]})", font=("Arial", 16, "bold")).pack(pady=15)

        # Canvas for Drawing
        canvas_f = ctk.CTkFrame(popup, fg_color="white")
        canvas_f.pack(fill="both", expand=True, padx=20, pady=20)
        
        canvas = tk.Canvas(canvas_f, bg="white", highlightthickness=0)
        scroll_h = ctk.CTkScrollbar(canvas_f, orientation="horizontal", command=canvas.xview)
        scroll_v = ctk.CTkScrollbar(canvas_f, orientation="vertical", command=canvas.yview)
        canvas.configure(xscrollcommand=scroll_h.set, yscrollcommand=scroll_v.set)

        scroll_h.pack(side="bottom", fill="x")
        scroll_v.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Build tree structure
        nodes = {}
        for rid, pid, rname, mcode, mname, img in roles:
            nodes[rid] = {"id": rid, "pid": pid, "role": rname, "name": mname or "TBA", "img": img, "children": []}
        
        root_nodes = []
        for rid, node in nodes.items():
            if node["pid"] and node["pid"] in nodes:
                nodes[node["pid"]]["children"].append(node)
            else:
                root_nodes.append(node)

        # Layout constants
        box_w = 150
        photo_h = 130
        text_h = 60
        node_h = photo_h + text_h
        v_gap, h_gap = 80, 50
        COLORS = ["#14B8A6", "#6F42C1", "#D63384", "#FD7E14", "#0D6EFD"]

        def get_tree_width(node):
            if not node["children"]: return box_w
            return sum(get_tree_width(c) for c in node["children"]) + (len(node["children"])-1)*h_gap

        def draw_node(node, x, y, level=0):
            color = COLORS[level % len(COLORS)]
            
            # Draw Box (Card)
            bx1, by1 = x - box_w/2, y
            bx2, by2 = x + box_w/2, y + node_h
            
            canvas.create_rectangle(bx1, by1, bx2, by2, fill="white", outline=color, width=2)
            
            # Photo Area
            if node["img"] and os.path.exists(node["img"]):
                try:
                    img = Image.open(node["img"]).resize((box_w - 4, photo_h - 4))
                    import PIL.ImageTk
                    tk_img = PIL.ImageTk.PhotoImage(img)
                    if not hasattr(popup, "_images"): popup._images = []
                    popup._images.append(tk_img)
                    canvas.create_image(x, y + photo_h/2, image=tk_img)
                except:
                    canvas.create_rectangle(bx1+2, by1+2, bx2-2, by1+photo_h-2, fill="#F3F4F6", outline="")
                    canvas.create_text(x, y + photo_h/2, text="📷", font=("Arial", 24))
            else:
                canvas.create_rectangle(bx1+2, by1+2, bx2-2, by1+photo_h-2, fill="#F3F4F6", outline="")
                canvas.create_text(x, y + photo_h/2, text="👤", font=("Arial", 32), fill="#9CA3AF")

            # Text Area background (Solid color)
            canvas.create_rectangle(bx1+1, by1 + photo_h, bx2-1, by2-1, fill=color, outline=color)

            # Roles & Name (Centered in text area)
            canvas.create_text(x, by1 + photo_h + 20, text=node["role"], font=("Arial", 14, "bold"), fill="white")
            canvas.create_text(x, by1 + photo_h + 42, text=node["name"], font=("Arial", 11), fill="white")

            # Children
            if node["children"]:
                child_y = y + node_h + v_gap
                total_w = get_tree_width(node)
                curr_x = x - total_w/2
                
                # Line down from parent
                mid_y = y + node_h + v_gap/2
                canvas.create_line(x, y + node_h, x, mid_y, fill="#9CA3AF", width=2)

                child_x_coords = []
                for child in node["children"]:
                    cw = get_tree_width(child)
                    cx = curr_x + cw/2
                    child_x_coords.append(cx)
                    
                    # Line up to bridge
                    canvas.create_line(cx, mid_y, cx, child_y, fill="#9CA3AF", width=2)
                    draw_node(child, cx, child_y, level + 1)
                    curr_x += cw + h_gap
                
                # Horizontal bridge
                if len(child_x_coords) > 1:
                    canvas.create_line(child_x_coords[0], mid_y, child_x_coords[-1], mid_y, fill="#9CA3AF", width=2)

        # Build visual tree
        total_w = sum(get_tree_width(r) for r in root_nodes) + (len(root_nodes)-1)*h_gap
        # Start at a reasonable x offset
        start_x_offset = max(500, total_w/2 + 50)
        curr_x = start_x_offset - total_w/2 + (get_tree_width(root_nodes[0])/2 if root_nodes else 0)
        
        for root in root_nodes:
            rw = get_tree_width(root)
            draw_node(root, curr_x, 30, 0)
            curr_x += rw + h_gap

        # Configure scrollregion
        canvas.update_idletasks()
        bbox = canvas.bbox("all")
        if bbox:
            canvas.configure(scrollregion=(bbox[0]-200, bbox[1]-100, bbox[2]+200, bbox[3]+200))

    def add_org_chart_popup(self, chart_id=None):
        popup = ctk.CTkToplevel(self)
        popup.title("Create/Edit Organization Chart")
        popup.geometry("700x800")
        popup.attributes("-topmost", True)
        popup.grab_set()

        # State management
        self.temp_chart_year = tk.StringVar(value=str(datetime.now().year))
        self.temp_chart_title = tk.StringVar(value="Organization Chart")
        self.roles_data = [] # List of {id, parent_id, role_name, member_code, member_name}
        self.next_role_id = 1

        if chart_id:
            conn = sqlite3.connect("database/attendance.db")
            chart = conn.execute("SELECT title, year FROM org_charts WHERE id=?", (chart_id,)).fetchone()
            if chart:
                self.temp_chart_title.set(chart[0])
                self.temp_chart_year.set(str(chart[1]))
                roles = conn.execute("SELECT id, parent_role_id, role_name, member_code FROM org_chart_roles WHERE chart_id=? ORDER BY id", (chart_id,)).fetchall()
                for rid, pid, rname, mcode in roles:
                    mname = ""
                    if mcode:
                        m = conn.execute("SELECT name FROM members WHERE member_code=?", (mcode,)).fetchone()
                        if m: mname = m[0]
                    self.roles_data.append({"id": rid, "parent_id": pid, "role_name": rname, "member_code": mcode, "member_name": mname})
                    if rid >= self.next_role_id: self.next_role_id = rid + 1
            conn.close()
        else:
            # Default structure
            self.roles_data.append({"id": 1, "parent_id": None, "role_name": "Chairman", "member_code": "", "member_name": ""})
            self.next_role_id = 2
            # Add defaults as per mockup
            for r in ["Financial", "Youth", "REU"]:
                self.roles_data.append({"id": self.next_role_id, "parent_id": 1, "role_name": r, "member_code": "", "member_name": ""})
                self.next_role_id += 1

        hdr = ctk.CTkFrame(popup, fg_color="transparent")
        hdr.pack(fill="x", padx=30, pady=20)
        
        ctk.CTkLabel(hdr, text="Organization Chart Title:", font=("Arial", 12, "bold")).pack(side="left")
        title_e = ctk.CTkEntry(hdr, textvariable=self.temp_chart_title, width=250)
        title_e.pack(side="left", padx=10)
        
        ctk.CTkLabel(hdr, text="Year:", font=("Arial", 12, "bold")).pack(side="left", padx=(20, 0))
        year_e = ctk.CTkEntry(hdr, textvariable=self.temp_chart_year, width=80)
        year_e.pack(side="left", padx=10)

        tree_f = ctk.CTkScrollableFrame(popup, fg_color="#F8F9FA", corner_radius=12)
        tree_f.pack(fill="both", expand=True, padx=30, pady=10)

        def render_tree():
            for w in tree_f.winfo_children(): w.destroy()
            
            def draw_node(parent_id, indent=0):
                children = [r for r in self.roles_data if r["parent_id"] == parent_id]
                for r in children:
                    row = ctk.CTkFrame(tree_f, fg_color="white", height=45, corner_radius=6, border_width=1, border_color="#E5E7EB")
                    row.pack(fill="x", pady=2, padx=(indent, 5))
                    row.pack_propagate(False)

                    # Role Title
                    r_e = ctk.CTkEntry(row, width=150, height=28, font=("Arial", 11))
                    r_e.insert(0, r["role_name"])
                    r_e.pack(side="left", padx=10)
                    r_e.bind("<KeyRelease>", lambda e, role=r: role.update({"role_name": e.widget.get()}))

                    # Member Select
                    m_txt = r["member_name"] if r["member_name"] else "Click to select member..."
                    m_btn = ctk.CTkButton(row, text=m_txt, width=200, height=28, fg_color="transparent", text_color="#007BFF", border_width=1, border_color="#007BFF", font=("Arial", 11),
                                          command=lambda role=r: self.pick_member_popup(lambda code, role=role: update_member(role, code)))
                    m_btn.pack(side="left", padx=10)

                    # Actions: Add Child (+)
                    plus = ctk.CTkButton(row, text="+", width=30, height=28, fg_color="#10B981", hover_color="#059669", command=lambda role=r: add_child(role["id"]))
                    plus.pack(side="left", padx=2)
                    Tooltip(plus, "Add Sub-role")

                    # Actions: Remove (x)
                    if r["parent_id"] is not None:
                        rem = ctk.CTkButton(row, text="x", width=30, height=28, fg_color="#DC3545", hover_color="#C82333", command=lambda role=r: remove_role(role["id"]))
                        rem.pack(side="left", padx=2)
                        Tooltip(rem, "Remove Role")

                    draw_node(r["id"], indent + 30)

            draw_node(None)
            ctk.CTkButton(tree_f, text="+ Add Level 1 Role", command=lambda: add_child(None), width=150, height=30, fg_color="transparent", text_color="#10B981", border_width=1, border_color="#10B981").pack(pady=10)

        def update_member(role, code):
            conn = sqlite3.connect("database/attendance.db")
            m = conn.execute("SELECT name FROM members WHERE member_code=?", (code,)).fetchone()
            conn.close()
            if m:
                role["member_code"] = code
                role["member_name"] = m[0]
                render_tree()

        def add_child(pid):
            self.roles_data.append({"id": self.next_role_id, "parent_id": pid, "role_name": "New Role", "member_code": "", "member_name": ""})
            self.next_role_id += 1
            render_tree()

        def remove_role(rid):
            to_remove = [rid]
            def find_children(pid):
                for r in self.roles_data:
                    if r["parent_id"] == pid:
                        to_remove.append(r["id"])
                        find_children(r["id"])
            find_children(rid)
            self.roles_data = [r for r in self.roles_data if r["id"] not in to_remove]
            render_tree()

        render_tree()

        def save_all():
            title = self.temp_chart_title.get().strip()
            year_str = self.temp_chart_year.get().strip()
            try: year = int(year_str)
            except: 
                messagebox.showerror("Error", "Invalid year"); return

            if not title:
                messagebox.showerror("Error", "Title is required"); return

            conn = sqlite3.connect("database/attendance.db")
            if chart_id:
                conn.execute("UPDATE org_charts SET title=?, year=? WHERE id=?", (title, year, chart_id))
                conn.execute("DELETE FROM org_chart_roles WHERE chart_id=?", (chart_id,))
                current_id = chart_id
            else:
                c = conn.cursor()
                c.execute("INSERT INTO org_charts (title, year) VALUES (?, ?)", (title, year))
                current_id = c.lastrowid

            id_map = {} 
            
            def save_node(temp_pid, db_pid):
                children = [r for r in self.roles_data if r["parent_id"] == temp_pid]
                for r in children:
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO org_chart_roles (chart_id, parent_role_id, role_name, member_code) VALUES (?, ?, ?, ?)",
                                   (current_id, db_pid, r["role_name"], r["member_code"]))
                    new_db_id = cursor.lastrowid
                    save_node(r["id"], new_db_id)

            save_node(None, None)
            conn.commit()
            conn.close()
            self.refresh_org_chart_table()
            popup.destroy()
            messagebox.showinfo("Success", "Organization Chart saved!")

        ftr = ctk.CTkFrame(popup, fg_color="transparent")
        ftr.pack(pady=20)
        ctk.CTkButton(ftr, text="💾 Save Organization Chart", width=250, height=45, fg_color="#28A745", hover_color="#218838", font=("Arial", 14, "bold"), command=save_all).pack()

    # ── Settings ──────────────────────────────────────────────────────────────

    def init_settings_page(self):
        main_f = ctk.CTkScrollableFrame(self.container, fg_color="#FFFFFF", corner_radius=10)
        self.frames["settings"] = main_f

        # Main 2-Column Split
        split = ctk.CTkFrame(main_f, fg_color="transparent")
        split.pack(fill="both", expand=True, padx=40, pady=20)
        split.grid_columnconfigure(0, weight=1, pad=40)
        split.grid_columnconfigure(1, weight=1, pad=40)

        # ── Left Side: App Settings ──────────────────────────────────────────
        left = ctk.CTkFrame(split, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew")

        # Logo Preview at the top
        self.settings_logo_preview = ctk.CTkLabel(left, text="", width=120, height=120, fg_color="#F3F4F6", corner_radius=15)
        self.settings_logo_preview.pack(pady=(0, 15), anchor="w")
        self._update_settings_logo_preview()

        ctk.CTkLabel(left, text="App Settings", font=("Arial", 18, "bold"), text_color="#111827").pack(pady=(0, 20), anchor="w")

        ctk.CTkLabel(left, text="Church Name:", font=("Arial", 12, "bold")).pack(anchor="w", pady=(5, 2))
        self.church_entry = ctk.CTkEntry(left, width=320, height=35)
        self.church_entry.insert(0, self.settings.get("church_name", ""))
        self.church_entry.pack(pady=(0, 15), anchor="w")

        ctk.CTkLabel(left, text="Default Area (used for Area Rate % calculation):", font=("Arial", 12, "bold")).pack(anchor="w", pady=(5, 2))
        self.area_entry = ctk.CTkEntry(left, width=320, height=35, placeholder_text="e.g. Kuala Lumpur")
        self.area_entry.insert(0, self.settings.get("default_area", ""))
        self.area_entry.pack(pady=(0, 15), anchor="w")

        ctk.CTkLabel(left, text="Member ID Prefix (e.g. SK for Skudai):", font=("Arial", 12, "bold")).pack(anchor="w", pady=(5, 2))
        self.prefix_entry = ctk.CTkEntry(left, width=120, height=35, placeholder_text="e.g. SK")
        self.prefix_entry.insert(0, self.settings.get("member_prefix", ""))
        self.prefix_entry.pack(pady=(0, 15), anchor="w")

        ctk.CTkLabel(left, text="Church Address:", font=("Arial", 12, "bold")).pack(anchor="w", pady=(5, 2))
        self.address_entry = ctk.CTkEntry(left, width=320, height=35, placeholder_text="e.g. 123 Church St, City")
        self.address_entry.insert(0, self.settings.get("address", ""))
        self.address_entry.pack(pady=(0, 15), anchor="w")

        ctk.CTkLabel(left, text="Church Logo:", font=("Arial", 12, "bold")).pack(anchor="w", pady=(5, 2))
        ctk.CTkButton(left, text="📁 Upload Logo Image", command=self.upload_logo, width=320, height=35, fg_color="#E5E7EB", text_color="#374151", hover_color="#D1D5DB").pack(pady=(0, 20), anchor="w")

        # ── Firewall Management ──
        ctk.CTkLabel(left, text="Network Security", font=("Arial", 18, "bold"), text_color="#111827").pack(pady=(15, 10), anchor="w")
        ctk.CTkLabel(left, text="Control the Confidential Data Firewall:", font=("Arial", 11), text_color="gray").pack(anchor="w", pady=(0, 10))
        
        self.fw_toggle_var = tk.BooleanVar(value=not FIREWALL_BYPASS)
        self.fw_switch = ctk.CTkSwitch(left, text="Data Privacy Firewall (Active)", 
                                       font=("Arial", 12, "bold"),
                                       progress_color="#10B981",
                                       variable=self.fw_toggle_var,
                                       command=self.toggle_firewall_manual)
        self.fw_switch.pack(pady=10, anchor="w")

        ctk.CTkButton(left, text="💾 Save App Settings", fg_color="#28A745", hover_color="#218838", width=320, height=42, font=("Arial", 13, "bold"), command=self.apply_settings).pack(pady=20, anchor="w")

        # ── Right Side: Backup Process ───────────────────────────────────────
        right = ctk.CTkFrame(split, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(right, text="Backup Process", font=("Arial", 18, "bold"), text_color="#111827").pack(pady=(10, 20), anchor="w")

        ctk.CTkLabel(right, text="Regularly backup your attendance database to prevent data loss.", font=("Arial", 11), text_color="gray", wraplength=320, justify="left").pack(anchor="w", pady=(0, 20))

        ctk.CTkButton(right, text="📤  Create New Backup", fg_color="#007BFF", hover_color="#0069D9", width=320, height=45, font=("Arial", 13, "bold"), command=self.perform_backup).pack(pady=10, anchor="w")

        ctk.CTkButton(right, text="📥  Restore from Backup", fg_color="transparent", text_color="#007BFF", border_width=1, border_color="#007BFF", hover_color="#EBF5FF", width=320, height=45, font=("Arial", 13, "bold"), command=self.perform_restore).pack(pady=10, anchor="w")

        # Last Backup Info
        self.backup_info_lbl = ctk.CTkLabel(right, text=f"Last Backup: {self.settings.get('last_backup', 'Never')}", font=("Arial", 11), text_color="#6B7280")
        self.backup_info_lbl.pack(pady=15, anchor="w")

        # ── Admin Security ──
        ctk.CTkLabel(right, text="Admin Security", font=("Arial", 18, "bold"), text_color="#111827").pack(pady=(30, 20), anchor="w")
        
        ctk.CTkLabel(right, text="Update Admin Password:", font=("Arial", 12, "bold")).pack(anchor="w", pady=(5, 2))
        self.new_pass_entry = ctk.CTkEntry(right, width=320, height=35, show="*", placeholder_text="Enter new password")
        self.new_pass_entry.pack(pady=(0, 10), anchor="w")
        
        ctk.CTkButton(right, text="🔐 Update Password", fg_color="#F59E0B", hover_color="#D97706", width=320, height=38, font=("Arial", 12, "bold"), command=self.update_admin_password).pack(anchor="w", pady=(0, 20))

        ctk.CTkLabel(right, text="Update Security Word (for password reset):", font=("Arial", 12, "bold")).pack(anchor="w", pady=(5, 2))
        self.hint_word_entry = ctk.CTkEntry(right, width=320, height=35, placeholder_text="e.g. skudai")
        self.hint_word_entry.insert(0, self.settings.get("admin_hint", "skudai"))
        self.hint_word_entry.pack(pady=(0, 10), anchor="w")
        
        ctk.CTkButton(right, text="🛡 Save Security Word", fg_color="#6366F1", hover_color="#4F46E5", width=320, height=38, font=("Arial", 12, "bold"), command=self.update_admin_hint).pack(anchor="w", pady=(0, 30))

        # ── GitHub Updates ──
        ctk.CTkLabel(right, text="Software Updates", font=("Arial", 18, "bold"), text_color="#111827").pack(pady=(10, 10), anchor="w")
        ctk.CTkLabel(right, text="Sync code with official GitHub repository:", font=("Arial", 11), text_color="gray").pack(anchor="w", pady=(0, 15))
        
        ctk.CTkButton(right, text="☁  Update from GitHub", fg_color="#6366F1", hover_color="#4F46E5", width=320, height=45, font=("Arial", 13, "bold"), command=self.update_from_github).pack(pady=5, anchor="w")
        ctk.CTkLabel(right, text="Note: Requires internet access and Git installed.", font=("Arial", 9), text_color="#9CA3AF").pack(anchor="w")

    def toggle_firewall_manual(self):
        global FIREWALL_BYPASS
        # If switch is ON (True), firewall is active -> bypass is FALSE
        FIREWALL_BYPASS = not self.fw_toggle_var.get()
        
        if FIREWALL_BYPASS:
            self.fw_badge.configure(text="⚠️ Firewall Disabled", fg_color="#FEE2E2", text_color="#EF4444")
            self.fw_switch.configure(text="Data Privacy Firewall (OFF)")
            messagebox.showwarning("Security Warning", "Data Privacy Firewall is now DISABLED.\nExternal network connections are now PERMITTED.\n\nOnly use this temporarily if you need to download models or sync data.")
        else:
            self.fw_badge.configure(text="🛡️ Firewall Active", fg_color="#D1FAE5", text_color="#059669")
            self.fw_switch.configure(text="Data Privacy Firewall (Active)")
            messagebox.showinfo("Security Info", "Data Privacy Firewall is now ACTIVE.\nExternal network connections are now BLOCKED.")

    def update_from_github(self):
        """Fetches and pulls the latest code from the official GitHub repository."""
        global FIREWALL_BYPASS
        
        # 1. Firewall Check
        if not FIREWALL_BYPASS:
            messagebox.showwarning("Firewall Blocked", 
                                   "The Data Privacy Firewall is currently ACTIVE.\n\n"
                                   "To download updates from GitHub, you must first switch the 'Data Privacy Firewall' to OFF in Settings.")
            return

        # 2. Confirmation
        if not messagebox.askyesno("Confirm Update", 
                                   "This will download the latest version from GitHub.\n\n"
                                   "WARNING: Any local code modifications will be replaced by the official version.\n\n"
                                   "Do you want to proceed?"):
            return

        import subprocess
        try:
            # Check if git is installed
            subprocess.run(["git", "--version"], check=True, capture_output=True)
            
            # Check if it's a git repo
            is_git = subprocess.run(["git", "rev-parse", "--is-inside-work-tree"], cwd=os.getcwd(), capture_output=True)
            if is_git.returncode != 0:
                # Convert ZIP download into a git repo automatically
                subprocess.run(["git", "init"], check=True, cwd=os.getcwd(), capture_output=True)
                subprocess.run(["git", "remote", "add", "origin", "https://github.com/esproducers/TJC-Auto-Attendance.git"], cwd=os.getcwd(), capture_output=True)
                
            # 3. Execution
            # We use git reset --hard origin/main to ensure the local code becomes identical to the server
            # This is the most reliable way to update for non-developers.
            subprocess.run(["git", "fetch", "--all"], check=True, cwd=os.getcwd(), capture_output=True)
            subprocess.run(["git", "reset", "--hard", "origin/main"], check=True, cwd=os.getcwd(), capture_output=True)
            
            messagebox.showinfo("Update Success", 
                                "Application updated successfully to the latest GitHub version!\n\n"
                                "Please RESTART the application for changes to take effect.")
        except FileNotFoundError:
            messagebox.showerror("Update Error", "Git is not installed on this system.\n\nPlease install Git from git-scm.com first.")
        except subprocess.CalledProcessError as e:
            err_msg = e.stderr.decode() if e.stderr else str(e)
            messagebox.showerror("Update Error", f"Failed to update from GitHub:\n{err_msg}")
        except Exception as e:
            messagebox.showerror("Update Error", f"An unexpected error occurred:\n{str(e)}")

    # ── SQL Data Page ─────────────────────────────────────────────────────────

    def init_sql_page(self):
        f = ctk.CTkFrame(self.container, fg_color="#F8F9FA", corner_radius=10)
        self.frames["sql"] = f
        
        hdr = ctk.CTkFrame(f, fg_color="transparent")
        hdr.pack(fill="x", padx=20, pady=20)
        
        ctk.CTkLabel(hdr, text="🗄 SQL Database Management", font=("Arial", 24, "bold")).pack(side="left")
        
        ctrl = ctk.CTkFrame(f, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color="#E5E7EB")
        ctrl.pack(fill="x", padx=20, pady=(0, 15))
        
        ctk.CTkLabel(ctrl, text="Table:", font=("Arial", 12, "bold")).pack(side="left", padx=(20, 5), pady=15)
        self.sql_table_var = ctk.StringVar(value="members")
        self.sql_table_cb = ctk.CTkComboBox(ctrl, variable=self.sql_table_var, values=["members", "attendance", "sessions", "org_charts"], command=lambda _: self.refresh_sql_table())
        self.sql_table_cb.pack(side="left", padx=5)
        
        ctk.CTkButton(ctrl, text="🔄 Refresh", width=100, command=self.refresh_sql_table).pack(side="left", padx=20)
        
        # Add Column Tool
        ctk.CTkLabel(ctrl, text="Add New Column:", font=("Arial", 11, "bold")).pack(side="left", padx=(40, 5))
        self.new_col_name = ctk.CTkEntry(ctrl, width=150, placeholder_text="column_name")
        self.new_col_name.pack(side="left", padx=5)
        ctk.CTkButton(ctrl, text="+ Add", width=60, fg_color="#10B981", command=self.on_sql_add_column).pack(side="left", padx=5)

        self.sql_scroll = ctk.CTkScrollableFrame(f, fg_color="#FFFFFF", corner_radius=8, border_width=1, border_color="#E5E7EB")
        self.sql_scroll.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
    def on_sql_add_column(self):
        col = self.new_col_name.get().strip().replace(" ", "_")
        table = self.sql_table_var.get()
        if not col: return
        
        if messagebox.askyesno("Confirm", f"Add column '{col}' to table '{table}'?"):
            try:
                conn = sqlite3.connect("database/attendance.db")
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} TEXT")
                conn.commit()
                conn.close()
                messagebox.showinfo("Success", f"Column '{col}' added successfully.")
                self.new_col_name.delete(0, 'end')
                self.refresh_sql_table()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add column: {e}")

    def refresh_sql_table(self):
        for w in self.sql_scroll.winfo_children(): w.destroy()
        
        table = self.sql_table_var.get()
        conn = sqlite3.connect("database/attendance.db")
        cursor = conn.cursor()
        try:
            cursor.execute(f"SELECT * FROM {table} LIMIT 100")
            rows = cursor.fetchall()
            cols = [description[0] for description in cursor.description]
            conn.close()
        except:
            conn.close()
            return

        # Header
        h_f = ctk.CTkFrame(self.sql_scroll, fg_color="#F3F4F6", height=40)
        h_f.pack(fill="x", pady=(0, 5))
        for i, c in enumerate(cols):
            ctk.CTkLabel(h_f, text=c.upper(), font=("Arial", 10, "bold"), text_color="#374151").place(x=10 + i*150, y=10)

        # Rows
        for r_idx, row_data in enumerate(rows):
            r_f = ctk.CTkFrame(self.sql_scroll, fg_color="transparent", height=35)
            r_f.pack(fill="x")
            for c_idx, val in enumerate(row_data):
                txt = str(val) if val is not None else ""
                if len(txt) > 20: txt = txt[:17] + "..."
                ctk.CTkLabel(r_f, text=txt, font=("Arial", 11)).place(x=10 + c_idx*150, y=5)
            
            ctk.CTkFrame(self.sql_scroll, height=1, fg_color="#E5E7EB").pack(fill="x")

    def update_admin_hint(self):
        nh = self.hint_word_entry.get().strip()
        if not nh:
            messagebox.showwarning("Warning", "Security word cannot be empty.")
            return
        if messagebox.askyesno("Confirm", "Update security word?"):
            self.settings["admin_hint"] = nh
            self.save_settings()
            messagebox.showinfo("Success", "Security word updated.")

    def update_admin_password(self):
        np = self.new_pass_entry.get().strip()
        if not np:
            messagebox.showwarning("Warning", "Password cannot be empty.")
            return
        if messagebox.askyesno("Confirm", "Are you sure you want to change the admin password?"):
            self.settings["admin_pass"] = np
            self.save_settings()
            messagebox.showinfo("Success", "Admin password updated.")
            self.new_pass_entry.delete(0, 'end')

    def perform_backup(self):
        try:
            os.makedirs("backup", exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            src = "database/attendance.db"
            if not os.path.exists(src):
                messagebox.showwarning("Error", "Database file not found!")
                return
                
            dst = os.path.join("backup", f"attendance_backup_{ts}.db")
            import shutil
            shutil.copy2(src, dst)
            
            # Update settings
            now_str = datetime.now().strftime("%d-%b-%Y %H:%M:%S")
            self.settings["last_backup"] = now_str
            self.save_settings()
            
            if hasattr(self, "backup_info_lbl"):
                self.backup_info_lbl.configure(text=f"Last Backup: {now_str}")
                
            messagebox.showinfo("Backup Success", f"Backup created successfully!\n\nLocation: {dst}")
        except Exception as e:
            messagebox.showerror("Backup Failed", str(e))

    def perform_restore(self):
        try:
            os.makedirs("backup", exist_ok=True)
            p = filedialog.askopenfilename(initialdir="backup", title="Select Backup to Restore", filetypes=[("Database", "*.db")])
            if not p: return

            if not messagebox.askyesno("Confirm Restore", "⚠️ RESTORE DATA?\n\nThis will replace all your current attendance data with this backup. Current recordings will be overwritten.\n\nContinue?"):
                return

            # Replace DB
            src = p
            dst = "database/attendance.db"
            
            # Close connection if possible (though sqlite in python usually handles this with re-opening)
            import shutil
            shutil.copy2(src, dst)
            
            messagebox.showinfo("Restore Success", "Data restored successfully! Please restart the app for changes to take effect.")
        except Exception as e:
            messagebox.showerror("Restore Failed", str(e))

    def _update_settings_logo_preview(self):
        p = self.settings.get("logo_path", "")
        if p and not os.path.isabs(p):
            p = os.path.join(os.path.dirname(os.path.abspath(__file__)), p)
        if p and os.path.exists(p):
            try:
                pil = Image.open(p).resize((120, 120))
                ci = ctk.CTkImage(light_image=pil, dark_image=pil, size=(120, 120))
                self.settings_logo_preview.configure(image=ci, text="")
                self.settings_logo_preview.image = ci
            except: 
                self.settings_logo_preview.configure(image=None, text="📷")
        else:
            self.settings_logo_preview.configure(image=None, text="📷")

    def upload_logo(self):
        p = filedialog.askopenfilename(filetypes=[("Image", "*.png *.jpg *.jpeg")])
        if p:
            try:
                # Create a local copy in the database folder
                import shutil
                ext = os.path.splitext(p)[1]
                local_path = os.path.join("database", f"app_logo{ext}")
                shutil.copy2(p, local_path)
                
                self.settings["logo_path"] = local_path
                self.save_settings()
                self._update_settings_logo_preview()
                
                # Update sidebar as well
                for w in self.sidebar.winfo_children():
                    if isinstance(w, ctk.CTkFrame): # Logo container
                        for sub in w.winfo_children():
                            if isinstance(sub, ctk.CTkLabel) and not sub.cget("text"): # Is logo label
                                self._display_logo(w)
                                break
                messagebox.showinfo("Logo Updated", f"Logo successfully saved as default!\n\nInternal Location: {local_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save logo: {e}")

    def apply_settings(self):
        self.settings["church_name"]  = self.church_entry.get()
        self.settings["default_area"] = self.area_entry.get().strip()
        self.settings["member_prefix"] = self.prefix_entry.get().strip()
        self.settings["address"]      = self.address_entry.get().strip()
        self.save_settings()
        self.title_label.configure(text=self.settings["church_name"])
        self.area_label.configure(text=self.settings["default_area"])
        messagebox.showinfo("Settings", "Settings saved!")

    def clear_attendance_history(self):
        msg = "⚠️ CLEAR ALL HISTORY?\n\nThis will permanently delete all attendance logs and records from the database. Registered members will NOT be deleted.\n\nType 'DELETE' to confirm:"
        dialog = ctk.CTkInputDialog(text=msg, title="Security Check")
        if dialog.get_input() == "DELETE":
            try:
                c = sqlite3.connect("database/attendance.db")
                c.execute("DELETE FROM attendance")
                c.execute("DELETE FROM sessions")
                c.commit()
                # Clear folders
                import shutil
                for folder in ["records/attendance", "records/unknown"]:
                    if os.path.exists(folder): shutil.rmtree(folder); os.makedirs(folder)
                c.close()
                self.refresh_stats()
                messagebox.showinfo("Success", "All history cleared.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to clear history: {e}")

    # ── Navigation ────────────────────────────────────────────────────────────

    def show_frame(self, name):
        for f in self.frames.values():
            f.grid_forget()
        self.frames[name].grid(row=0, column=0, sticky="nsew")
        self.page_title.configure(text=name.replace("_", " ").title())
        for k, btn in self.nav_buttons.items():
            btn.configure(fg_color="#007BFF" if k == name else "transparent",
                          text_color="white" if k == name else "#333")
        if   name == "dashboard": self.refresh_stats()
        elif name == "logs":      self.refresh_logs_table()
        elif name == "reports":   self.refresh_sessions_summary()
        elif name == "org_chart": self.refresh_org_chart_table()

    # ── Dashboard Stats ───────────────────────────────────────────────────────

    def refresh_stats(self):
        area = self.settings.get("default_area", "") or None
        s = self.backend.get_summary(default_area=area)
        
        present_count = s["p_total"]
        waiting_count = s["waiting"]
        # Rebuild widgets if members OR unknowns changed
        needs_rebuild = (present_count != self.last_stats_count) or (waiting_count != self.last_waiting_count)
        self.last_stats_count = present_count
        self.last_waiting_count = waiting_count

        self.cards["Present Today"].configure(text=str(s["p_total"]))
        self.cards["Area Member"].configure(text=str(s["p_area_member"]))
        self.cards["Other Area Member"].configure(text=str(s["p_other_member"]))
        self.cards["Truth Seeker"].configure(text=str(s["p_truth"]))
        self.cards["Waiting Recognition"].configure(text=str(s["waiting"]))
        self.cards["Area Rate %"].configure(text=f"{s['area_rate']:.1f}%")
        self.cards["Overall Rate %"].configure(text=f"{s['overall_rate']:.1f}%")

        if not needs_rebuild:
            return

        # Rebuild captured cards
        for w in self.checkin_scroll.winfo_children():
            w.destroy()

        df = s["list"]
        if not df.empty:
            for _, row in df.head(20).iterrows():
                img_path = str(row.get("record_image", "") or "")
                if not os.path.exists(img_path):
                    img_path = ""
                CheckInCard(
                    self.checkin_scroll,
                    att_id      = row.get("id"),
                    name        = str(row.get("name", "?")),
                    age         = row.get("age", ""),
                    img_path    = img_path,
                    m_type      = str(row.get("status", "member")),
                    member_code = row.get("member_code"),
                    on_click    = self.on_view_member,
                    on_identify = self.identify_unknown_popup,
                ).pack(side="top", fill="x", padx=10, pady=4)

        self._refresh_waiting_panel()

    def _refresh_waiting_panel(self):
        for w in self.waiting_scroll.winfo_children():
            w.destroy()
        rows = self.backend.get_waiting_list()
        if not rows:
            ctk.CTkLabel(self.waiting_scroll, text="None",
                         font=("Arial", 11), text_color="gray").pack(pady=8)
            return
        for att_id, img_path, t in rows:
            r = ctk.CTkFrame(self.waiting_scroll, fg_color="#FFF3CD", corner_radius=6, height=48)
            r.pack(fill="x", pady=3)
            r.pack_propagate(False)
            ctk.CTkLabel(r, text=f"  Unknown  {str(t)[:16]}",
                         font=("Arial", 10), anchor="w").pack(side="left", padx=6, fill="x", expand=True)
            ctk.CTkButton(r, text="Identify", width=72, height=28, font=("Arial", 9),
                           fg_color="#FFC107", text_color="black", hover_color="#E0A800",
                           command=lambda aid=att_id, ip=img_path:
                           self.identify_unknown_popup(aid, ip)).pack(side="right", padx=6)

    # ── Session Controls ──────────────────────────────────────────────────────

    def on_start_click(self):
        popup = ctk.CTkToplevel(self)
        popup.title("Start Attendance Session")
        popup.geometry("420x400")
        popup.attributes("-topmost", True)
        popup.grab_set()

        ctk.CTkLabel(popup, text="Start New Session", font=("Arial", 16, "bold")).pack(pady=(20, 15))

        ctk.CTkLabel(popup, text="Seminar / Event Title:", anchor="w").pack(fill="x", padx=30)
        dt = datetime.now()
        default_title = f"{dt.strftime('%A').upper()}-SEMINAR-{dt.strftime('%d%b%Y').upper()}"
        title_e = ctk.CTkEntry(popup, width=340, placeholder_text="e.g. Sunday Service")
        title_e.insert(0, default_title)
        title_e.pack(padx=30, pady=(4, 12))

        ctk.CTkLabel(popup, text="Duration in MINUTES (leave blank = manual stop):", anchor="w").pack(fill="x", padx=30)
        dur_e = ctk.CTkEntry(popup, width=340, placeholder_text="e.g.  60  (for 1 hour)")
        dur_e.pack(padx=30, pady=(4, 12))

        ctk.CTkLabel(popup, text="Seminar Type (for periodical reporting):", anchor="w").pack(fill="x", padx=30)
        sem_var = ctk.StringVar(value="Other")
        if dt.strftime('%A') == 'Friday': sem_var.set("Friday Seminar")
        elif dt.strftime('%A') == 'Saturday': sem_var.set("Saturday Seminar")
        
        sem_menu = ctk.CTkSegmentedButton(popup, values=["Other", "Friday Seminar", "Saturday Seminar"], variable=sem_var)
        sem_menu.pack(padx=30, pady=(4, 20), fill="x")

        def confirm():
            title = title_e.get().strip()
            dur   = dur_e.get().strip()
            stype = sem_var.get()
            if not title:
                messagebox.showwarning("Missing", "Please enter a title.", parent=popup)
                return

            dur_mins = None
            if dur:
                try:
                    dur_mins = int(dur)
                    if dur_mins <= 0:
                        raise ValueError
                except ValueError:
                    messagebox.showwarning("Invalid", "Duration must be a positive whole number (minutes).",
                                           parent=popup)
                    return

            self.backend.start_session(title, dur_mins, seminar_type=stype)
            self.session_title    = title
            self.session_deadline = (datetime.now() + timedelta(minutes=dur_mins)) if dur_mins else None

            self.is_marking = True
            self.is_paused  = False
            self.start_btn.configure(state="disabled")
            self.pause_btn.configure(state="normal", text="⏸  Pause", fg_color="#FFC107", text_color="black")
            self.end_btn.configure(state="normal")
            self.session_info_lbl.configure(text=f"● {title}", text_color="#28A745")

            # Clear scroll lists
            for w in self.checkin_scroll.winfo_children():
                w.destroy()
            for w in self.waiting_scroll.winfo_children():
                w.destroy()
            self.activity_log.delete("1.0", "end")

            popup.destroy()

        ctk.CTkButton(popup, text="▶  Start Session", fg_color="#28A745",
                      hover_color="#218838", width=200, height=40, command=confirm).pack()

    def on_pause_click(self):
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.is_marking = False
            self.pause_btn.configure(text="▶  Resume", fg_color="#28A745", text_color="white")
        else:
            self.is_marking = True
            self.pause_btn.configure(text="⏸  Pause", fg_color="#FFC107", text_color="black")

    def on_end_click(self):
        if messagebox.askyesno("End Session", "End this session and save all attendance?"):
            self.finalize_session(manual=True)

    def finalize_session(self, manual=False):
        self.backend.end_session()
        self.is_marking       = False
        self.is_paused        = False
        self.session_deadline = None

        self.start_btn.configure(state="normal")
        self.pause_btn.configure(state="disabled", text="⏸  Pause",
                                  fg_color="#FFC107", text_color="black")
        self.end_btn.configure(state="disabled")
        self.resume_btn.configure(state="normal")
        self.session_info_lbl.configure(text="● No Active Session", text_color="#999")

        self.refresh_stats()
        self.refresh_logs_table()

    def on_resume_click(self):
        conn = sqlite3.connect("database/attendance.db")
        last_sess = conn.execute("SELECT id, title FROM sessions ORDER BY id DESC LIMIT 1").fetchone()
        conn.close()

        if not last_sess:
            messagebox.showwarning("Resume", "No previous session found to resume from.")
            return

        old_id, old_title = last_sess
        
        # Open start dialog with a hint
        popup = ctk.CTkToplevel(self)
        popup.title("Resume Previous Session")
        popup.geometry("400x420")
        popup.attributes("-topmost", True)
        popup.grab_set()

        ctk.CTkLabel(popup, text="RE-START SESSION", font=("Arial", 16, "bold")).pack(pady=15)
        ctk.CTkLabel(popup, text=f"Cloning attendees from: {old_title}", font=("Arial", 11), text_color="gray").pack()
        
        row1 = ctk.CTkFrame(popup, fg_color="transparent")
        row1.pack(pady=10)
        ctk.CTkLabel(row1, text="Title: ").pack(side="left")
        title_e = ctk.CTkEntry(row1, width=250)
        title_e.insert(0, f"{old_title} (Part 2)")
        title_e.pack(side="left")

        row2 = ctk.CTkFrame(popup, fg_color="transparent")
        row2.pack(pady=10)
        ctk.CTkLabel(row2, text="Duration (mins): ").pack(side="left")
        dur_e = ctk.CTkEntry(row2, width=80)
        dur_e.insert(0, "60")
        dur_e.pack(side="left")

        def confirm_resume():
            title = title_e.get().strip()
            dur_str = dur_e.get().strip()
            if not title: return
            
            try:
                dur_val = int(dur_str) if dur_str else None
            except:
                dur_val = 60

            # 1. Create new session (this also clears backend lists)
            new_id = self.backend.start_session(title, duration_mins=dur_val)
            if not new_id: return

            # 2. Setup Timer Logic
            if dur_val:
                self.session_deadline = datetime.now() + timedelta(minutes=dur_val)

            # 3. Clone attendance records
            conn = sqlite3.connect("database/attendance.db")
            conn.execute("""
                INSERT INTO attendance (person_name, member_code, session_id, record_image, check_in_time, service_date, status)
                SELECT person_name, member_code, ?, record_image, ?, service_date, status
                FROM attendance WHERE session_id = ?
            """, (new_id, datetime.now(), old_id))
            conn.commit()
            
            # Fetch for UI population
            cloned = conn.execute("SELECT person_name, record_image, status, member_code FROM attendance WHERE session_id=?", (new_id,)).fetchall()
            conn.close()

            # 4. CRITICAL: Update backend captured list to STOP double-marking
            for _, _, _, code in cloned:
                if code: self.backend.session_captured_ids.add(code)

            # 5. Update UI
            self.is_marking = True
            self.start_btn.configure(state="disabled")
            self.pause_btn.configure(state="normal")
            self.end_btn.configure(state="normal")
            self.resume_btn.configure(state="disabled")
            self.session_info_lbl.configure(text=f"● {title}", text_color="#28A745")

            # Populate the Captured list immediately
            for w in self.checkin_scroll.winfo_children(): w.destroy()
            for name, img, status, code in cloned:
                self.add_attendee_card(name, img, status, code)

            self.refresh_stats()
            popup.destroy()

        ctk.CTkButton(popup, text="🚀  Resume & Start", fg_color="#6F42C1", hover_color="#5A32A3", width=200, height=40, command=confirm_resume).pack(pady=10)

    # ── Manual Controls ───────────────────────────────────────────────────────

    def filter_captured_list(self, _e=None):
        query = self.dash_search.get().lower().strip()
        for card in self.checkin_scroll.winfo_children():
            if hasattr(card, "_search_data"):
                if not query or query in card._search_data.lower():
                    card.pack(side="top", fill="x", padx=10, pady=4)
                else:
                    card.pack_forget()

    def manual_add_popup(self):
        if not hasattr(self.backend, "active_session_id") or not self.backend.active_session_id:
            messagebox.showwarning("No Session", "Please start a session first.")
            return

        dialog = ctk.CTkToplevel(self)
        dialog.title("Manual Add Attendee")
        dialog.geometry("400x500")
        dialog.attributes("-topmost", True)

        ctk.CTkLabel(dialog, text="Select Member to Add:", font=("Arial", 14, "bold")).pack(pady=10)
        
        s_entry = ctk.CTkEntry(dialog, placeholder_text="Search all members...")
        s_entry.pack(fill="x", padx=20, pady=5)

        list_f = ctk.CTkScrollableFrame(dialog)
        list_f.pack(fill="both", expand=True, padx=20, pady=10)

        def populate_list():
            for w in list_f.winfo_children(): w.destroy()
            q = s_entry.get().lower()
            
            conn = sqlite3.connect("database/attendance.db")
            # Exclude those already present in this session
            members = conn.execute("""
                SELECT member_code, name, type FROM members 
                WHERE member_code NOT IN (SELECT member_code FROM attendance WHERE session_id=?)
                AND (LOWER(name) LIKE ? OR LOWER(member_code) LIKE ?)
            """, (self.backend.active_session_id, f"%{q}%", f"%{q}%")).fetchall()
            conn.close()

            for code, name, mtype in members:
                btn = ctk.CTkButton(list_f, text=f"{name} ({code})", anchor="w", fg_color="transparent", text_color="#333", hover_color="#F0F2F5",
                                    command=lambda c=code, n=name, t=mtype: [self.do_manual_mark(c, n, t), dialog.destroy()])
                btn.pack(fill="x", pady=1)

        s_entry.bind("<KeyRelease>", lambda e: populate_list())
        populate_list()

    def do_manual_mark(self, code, name, mtype):
        # Fetch profile image and title
        conn = sqlite3.connect("database/attendance.db")
        row = conn.execute("SELECT image_path, title FROM members WHERE member_code=?", (code,)).fetchone()
        conn.close()
        p_img = row[0] if row else ""
        m_title = row[1] if row else ""

        # Mark in backend
        if hasattr(self.backend, "manual_mark"):
            ok, path = self.backend.manual_mark(name, code, p_img, mtype)
        else:
            ok, path = self.backend.mark_attendance(name, code, None, mtype)

        if ok:
            self.refresh_stats()
            self.activity_log.insert("1.0", f"[{datetime.now().strftime('%H:%M')}] ✅ {m_title} {name} (MANUAL)\n")
            # Card addition
            self.add_attendee_card(name, p_img, mtype, code, title=m_title)

    # ── Thread-Safe GUI Updates ───────────────────────────────────────────────

    def process_gui_queue(self):
        """Polls the queue for UI updates from background threads."""
        try:
            while True:
                task = self.gui_queue.get_nowait()
                task()
        except queue.Empty:
            pass
        self.after(50, self.process_gui_queue)

    def manual_remove_attendee(self):
        if not hasattr(self.backend, "active_session_id") or not self.backend.active_session_id: return
        
        # This is a bit tricky: we ask for the name/code or use the selection
        dialog = ctk.CTkInputDialog(text="Enter Member Code to Remove:", title="Remove Attendee")
        code = dialog.get_input()
        if code:
            conn = sqlite3.connect("database/attendance.db")
            conn.execute("DELETE FROM attendance WHERE session_id=? AND member_code=?", (self.backend.active_session_id, code))
            conn.commit()
            conn.close()
            
            # Remove from UI list
            for card in self.checkin_scroll.winfo_children():
                if hasattr(card, "member_code") and card.member_code == code:
                    card.destroy()
            
            self.refresh_stats()
            messagebox.showinfo("Success", f"Member {code} removed from session.")

    def add_attendee_card(self, name, img_path, m_type, code, title=""):
        # Prevent duplicates
        for child in self.checkin_scroll.winfo_children():
            if hasattr(child, "member_code") and child.member_code == code:
                return child

        # Horizontal Row Design (Stable & Sleek)
        card = ctk.CTkFrame(self.checkin_scroll, fg_color="#FFFFFF", border_width=1, border_color="#E5E7EB", corner_radius=10, height=60)
        card.pack(side="top", fill="x", padx=10, pady=4)
        card.pack_propagate(False) # Preserve height
        card.member_code = code

        # [1] Profile Photo
        img_f = ctk.CTkFrame(card, width=50, height=50, fg_color="transparent")
        img_f.pack(side="left", padx=10, pady=5)
        img_f.pack_propagate(False)
        
        img_lbl = ctk.CTkLabel(img_f, text="📷", font=("Arial", 16))
        img_lbl.pack(expand=True)
        if img_path and os.path.exists(img_path):
            try:
                pil = Image.open(img_path).resize((40, 40))
                ci = ctk.CTkImage(light_image=pil, size=(40, 40))
                img_lbl.configure(image=ci, text="")
            except: pass

        # [2] Name and Code Info
        txt_f = ctk.CTkFrame(card, fg_color="transparent")
        txt_f.pack(side="left", fill="both", expand=True, padx=5)
        
        ctk.CTkLabel(txt_f, text=name.upper(), font=("Arial", 12, "bold"), anchor="w").pack(pady=(8, 0), fill="x")
        ctk.CTkLabel(txt_f, text=f"ID: {code}  |  {title}", font=("Arial", 10), text_color="#6B7280", anchor="w").pack(fill="x")

        # [3] Status and Time
        right_f = ctk.CTkFrame(card, fg_color="transparent")
        right_f.pack(side="right", padx=15)

        badge_colors = {"area member": "#10B981", "other area member": "#6366F1", "truth seeker": "#17A2B8", "unknown": "#DC3545"}
        b_color = badge_colors.get(m_type.lower(), "#10B981")
        ctk.CTkLabel(right_f, text=m_type.upper(), font=("Arial", 8, "bold"), fg_color=b_color, text_color="white", corner_radius=4, width=80).pack(pady=(8, 2))
        
        now_t = datetime.now().strftime("%I:%M %p")
        ctk.CTkLabel(right_f, text=now_t, font=("Arial", 9), text_color="#9CA3AF").pack()

        # Store metadata for robust searching
        card._search_data = f"{name} {code}".lower()
        return card

    # ── Camera Loop ───────────────────────────────────────────────────────────

    def update_camera(self):
        # Countdown / auto-stop
        if self.session_deadline:
            rem = self.session_deadline - datetime.now()
            if rem.total_seconds() > 0:
                total_sec = int(rem.total_seconds())
                h, r = divmod(total_sec, 3600)
                m, s = divmod(r, 60)
                if h > 0:
                    timer_str = f"{h:02d}h {m:02d}m {s:02d}s"
                else:
                    timer_str = f"{m:02d}m {s:02d}s"
                self.session_info_lbl.configure(
                    text=f"● {self.session_title}  |  ⏱ {timer_str} left",
                    text_color="#DC3545")
            elif self.is_marking:
                self.finalize_session(manual=False)

        ret, frame = self.backend.camera.read()
        if ret:
            self.last_frame = frame.copy()
            if self.is_marking:
                # 1. Dispatch background thread if idle
                if not self.is_processing:
                    self.is_processing = True
                    def worker(f):
                        try:
                            res_frame, res_list = self.backend.process_frame(f)
                            self.result_queue.put((res_frame, res_list))
                        except Exception as e:
                            # Ensure we don't dead-lock if backend crashes
                            print(f"[SYSTEM] Background processing error: {e}")
                            self.result_queue.put((None, []))
                    threading.Thread(target=worker, args=(frame.copy(),), daemon=True).start()

                # 2. Check for results from thread
                try:
                    res_frame, results = self.result_queue.get_nowait()
                    self.is_processing = False
                    self.last_results = results # Keep for boxes
                    
                    if results:
                        self.refresh_stats()
                        for res in results:
                            # Set feedback banner (3 seconds)
                            if res.get('new'):
                                color = (0, 200, 0) if res['name'] != "Unknown" else (0, 140, 255)
                                self.capture_feedback = {
                                    "msg": f"CAPTURED: {res['name'].upper()}",
                                    "expiry": time.time() + 3.0,
                                    "color": color
                                }

                            ts = datetime.now().strftime("%H:%M")
                            if res['name'] != "Unknown":
                                m_title = res.get('title', '')
                                self.activity_log.insert("end", f"[{ts}] ✅ {m_title} {res['name']} ({res['type']})\n")
                                if res.get('new'):
                                    self.add_attendee_card(res['name'], res['img'], res['type'], res['code'], title=m_title)
                            else:
                                self.activity_log.insert("end", f"[{ts}] ❓ Unknown face captured\n")
                                if res.get('new'):
                                    # unknown logic - stays in waiting list
                                    pass
                            self.activity_log.see("end")
                except queue.Empty:
                    pass

                # Draw boxes for active tracking/feedback
                if self.last_results:
                    for r in self.last_results:
                        b     = r['bbox']
                        color = (255, 255, 0) # Cyan (BGR)
                        cv2.rectangle(frame, (b[0], b[1]), (b[2], b[3]), color, 2)
                        cv2.putText(frame, r['name'], (b[0], b[1] - 10),
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)

                # 3. Draw Persistent Feedback Banner (top of screen)
                if time.time() < self.capture_feedback["expiry"]:
                    msg = self.capture_feedback["msg"]
                    color = self.capture_feedback["color"]
                    h, w = frame.shape[:2]
                    # Draw a semi-transparent black strip at the top
                    overlay = frame.copy()
                    cv2.rectangle(overlay, (0, 0), (w, 45), (0,0,0), -1)
                    cv2.addWeighted(overlay, 0.6, frame, 0.4, 0, frame)
                    # Text
                    cv2.putText(frame, msg, (20, 32), cv2.FONT_HERSHEY_DUPLEX, 0.8, color, 2)
                    # Small checkmark icon
                    ico = "OK" if "UNKNOWN" not in msg else "??"
                    cv2.putText(frame, ico, (w - 50, 32), cv2.FONT_HERSHEY_DUPLEX, 0.8, color, 2)

            # Update display
            pil_img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
            ci      = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(600, 380))
            self.cam_label.configure(image=ci, text="")
            self.cam_label._image = ci

        self.after(30, self.update_camera)

    # ── Identify Unknown ──────────────────────────────────────────────────────

    def identify_unknown_popup(self, att_id, img_path):
        popup = ctk.CTkToplevel(self)
        popup.title("Identify Person")
        popup.geometry("480x580")
        popup.attributes("-topmost", True)
        popup.grab_set()

        ctk.CTkLabel(popup, text="Identify Unrecognised Person",
                     font=("Arial", 15, "bold")).pack(pady=(18, 10))

        # Show the captured photo
        try:
            pil = Image.open(img_path).resize((160, 160))
            ci  = ctk.CTkImage(light_image=pil, dark_image=pil, size=(160, 160))
            lbl = ctk.CTkLabel(popup, image=ci, text="")
            lbl.image = ci
            lbl.pack(pady=4)
        except Exception:
            ctk.CTkLabel(popup, text="📷", font=("Arial", 48)).pack(pady=4)

        ctk.CTkLabel(popup, text="Action:", anchor="w").pack(fill="x", padx=30, pady=(10, 2))
        action_var = ctk.StringVar(value="register_new")
        ctk.CTkRadioButton(popup, text="Register as NEW member",
                           variable=action_var, value="register_new").pack(anchor="w", padx=40)
        ctk.CTkRadioButton(popup, text="Link to EXISTING member code",
                           variable=action_var, value="link_existing").pack(anchor="w", padx=40, pady=(4, 12))

        form = ctk.CTkScrollableFrame(popup, fg_color="transparent", height=320)
        form.pack(fill="both", expand=True, padx=20)

        # Form Fields
        ctk.CTkLabel(form, text="Existing member code (if linking):",
                     font=("Arial", 11, "bold")).pack(anchor="w", pady=(10, 0))
        existing_code_e = ctk.CTkEntry(form, width=360, placeholder_text="e.g. 0003")
        existing_code_e.pack(pady=3, fill="x")

        def on_member_selected(code):
            conn = sqlite3.connect("database/attendance.db")
            m = conn.execute("SELECT name, type, dob, area, phone, baptism_date, address, email, has_holy_spirit, title FROM members WHERE member_code=?", (code,)).fetchone()
            conn.close()
            if m:
                name_e.delete(0, "end");   name_e.insert(0, m[0] or "")
                type_cb.set(m[1] or "Member")
                set_dob(m[2] or "")
                area_e.delete(0, "end");   area_e.insert(0, m[3] or "")
                phone_e.delete(0, "end");  phone_e.insert(0, m[4] or "")
                set_bap(m[5] or "")
                address_e.delete(0, "end"); address_e.insert(0, m[6] or "")
                email_e.delete(0, "end"); email_e.insert(0, m[7] or "")
                hs_var.set(bool(m[8]))
                title_cb.set(m[9] or "")
                
                # Switch to link mode
                action_var.set("link_existing")
                existing_code_e.delete(0, "end"); existing_code_e.insert(0, code)

        ctk.CTkButton(form, text="🔍  Search & Select Existing Member", font=("Arial", 11),
                      fg_color="#6C757D", hover_color="#5A6268", height=28,
                      command=lambda: self.pick_member_popup(on_member_selected)).pack(pady=5)

        ctk.CTkLabel(form, text="Name *", font=("Arial", 11, "bold")).pack(anchor="w", pady=(15, 0))
        name_e = ctk.CTkEntry(form, width=360)
        name_e.pack(pady=3, fill="x")

        ctk.CTkLabel(form, text="Title", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        title_var = ctk.StringVar(value="")
        title_cb = ctk.CTkComboBox(form, variable=title_var, values=["", "Brother", "Sister"], width=360)
        title_cb.pack(pady=3, fill="x")

        ctk.CTkLabel(form, text="Type (Area Member / Other Area Member / Truth Seeker)", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        type_var = ctk.StringVar(value="Area Member")
        type_cb = ctk.CTkComboBox(form, variable=type_var, values=["Area Member", "Other Area Member", "Truth Seeker"], width=360)
        type_cb.pack(pady=3, fill="x")

        ctk.CTkLabel(form, text="Area", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        area_e = ctk.CTkEntry(form, width=360, placeholder_text="e.g. Kuala Lumpur")
        area_e.pack(pady=3, fill="x")

        def update_age_cat_ui(dob):
            if not dob or dob == "--": 
                age_cat_var.set("")
                return
            try:
                birth = datetime.strptime(dob, '%d-%m-%Y')
                age = (date.today() - birth.date()).days // 365
                if age <= 12: cat = "Child"
                elif age <= 24: cat = "Youth"
                elif age <= 64: cat = "Adult"
                else: cat = "Elder"
                age_cat_var.set(cat)
            except: age_cat_var.set("")

        get_dob, set_dob = self._date_picker(form, "DOB (DD-MM-YYYY)", on_change=update_age_cat_ui)

        # Age Category Dropdown (Manual select allowed)
        ctk.CTkLabel(form, text="Age Category", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        age_cat_var = ctk.StringVar(value="")
        age_cat_cb = ctk.CTkComboBox(form, variable=age_cat_var, width=360, 
                                      values=["", "Child", "Youth", "Adult", "Elder"])
        age_cat_cb.pack(pady=3, fill="x")

        get_bap, set_bap = self._date_picker(form, "Date of Baptism (DD-MM-YYYY)")

        ctk.CTkLabel(form, text="Address", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        address_e = ctk.CTkEntry(form, width=360)
        address_e.pack(pady=3, fill="x")

        ctk.CTkLabel(form, text="Email", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        email_e = ctk.CTkEntry(form, width=360)
        email_e.pack(pady=3, fill="x")

        ctk.CTkLabel(form, text="Phone", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        phone_e = ctk.CTkEntry(form, width=360)
        phone_e.pack(pady=3, fill="x")

        hs_var = tk.BooleanVar(value=False)
        hs_cb = ctk.CTkCheckBox(form, text="Holy Spirit Received", variable=hs_var, font=("Arial", 12, "bold"))
        hs_cb.pack(anchor="w", pady=10)

        ctk.CTkLabel(form, text="Remark", font=("Arial", 11, "bold")).pack(anchor="w", pady=(8, 0))
        remark_e = ctk.CTkTextbox(form, width=360, height=70) # ~3 lines
        remark_e.pack(pady=3, fill="x")


        def save():
            name = name_e.get().strip()
            if not name:
                messagebox.showwarning("Missing", "Name is required.", parent=popup)
                return

            m_type = type_var.get()

            # --- DUPLICATE CHECK ---
            target_code = existing_code_e.get().strip() if action_var.get() == "link_existing" else None
            if target_code:
                conn = sqlite3.connect("database/attendance.db")
                exists = conn.execute("SELECT 1 FROM attendance WHERE session_id=? AND member_code=?", 
                                     (self.backend.active_session_id, target_code)).fetchone()
                conn.close()
                if exists:
                    messagebox.showerror("Already Present", f"Member {target_code} is already checked in for this session.", parent=popup)
                    return
            # ------------------------

            if action_var.get() == "link_existing":
                code = existing_code_e.get().strip()
                if not code:
                    messagebox.showwarning("Missing", "Member code required.", parent=popup)
                    return
                try:
                    # Update member title/details if needed
                    self.backend.register_member({"name": name, "title": title_cb.get(), "type": m_type}, force_code=code)
                    # Promote attendance row
                    self.backend.identify_unknown(att_id, name, code, m_type)
                except Exception as e:
                    messagebox.showerror("Save Error", f"Could not update member record: {str(e)}", parent=popup)
                    return
            else:
                # DUPLICATE NAME CHECK FOR NEW MEMBER
                conn = sqlite3.connect("database/attendance.db")
                exists_name = conn.execute("SELECT member_code FROM members WHERE name = ? COLLATE NOCASE", (name,)).fetchone()
                conn.close()
                if exists_name:
                    if not messagebox.askyesno("Duplicate Name", f"A member with the name '{name}' already exists (ID: {exists_name[0]}).\n\nDo you want to proceed and create a new duplicate record?", parent=popup):
                        return
                        
                # Register brand-new member
                # Load extra fields schema
                extra_entries = {}
                conn = sqlite3.connect("database/attendance.db")
                cursor = conn.cursor()
                cursor.execute("PRAGMA table_info(members)")
                db_cols = [row[1].lower() for row in cursor.fetchall()]
                conn.close()
                
                standard_fields = ["member_code", "name", "type", "age", "dob", "baptism_date", 
                                   "address", "email", "phone", "has_holy_spirit", "image_path", 
                                   "registration_date", "area", "remark", "age_category", "title"]
                extra_fields = [c for c in db_cols if c not in standard_fields]
                
                # Show extra fields in UI before save (optional, but requested to reflect)
                # For this quick popup, we'll just gather them if they were somehow in the form
                # Actually, the user wants "add here will reflect to default list"
                # So I should probably add them to the register_member_popup too.
                
                data = {
                    "name": name,
                    "type": m_type,
                    "title": title_var.get(),
                    "dob":  get_dob(),
                    "baptism_date": get_bap(),
                    "area": area_e.get().strip(),
                    "address": address_e.get().strip(),
                    "email": email_e.get().strip(),
                    "phone": phone_e.get().strip(),
                    "has_holy_spirit": hs_var.get(),
                    "remark": remark_e.get("1.0", "end").strip(),
                    "age_category": age_cat_var.get(),
                    "image_path": img_path
                }
                prefix = self.settings.get("member_prefix", "")
                code = self.backend.register_member(data, prefix=prefix)
                self.backend.identify_unknown(att_id, name, code, m_type)

            self.refresh_stats()
            self.refresh_logs_table()
            
            # Show in captured list right away
            if code and m_type:
                self.add_attendee_card(name, img_path, m_type, code, title=title_var.get())
                
            popup.destroy()
            messagebox.showinfo("Done", f"Person identified as '{name}' (code: {code}).")

        def dismiss():
            if messagebox.askyesno("Confirm", "Are you sure you want to dismiss this?", parent=popup):
                c = sqlite3.connect("database/attendance.db")
                c.execute("DELETE FROM attendance WHERE id=?", (att_id,))
                c.commit()
                c.close()
                self.refresh_stats()
                self.refresh_logs_table()
                popup.destroy()

        btn_row = ctk.CTkFrame(popup, fg_color="transparent")
        btn_row.pack(pady=14)

        ctk.CTkButton(btn_row, text="✔  Save & Identify", fg_color="#28A745",
                      hover_color="#218838", width=160, height=38, command=save).pack(side="left", padx=5)
        
        ctk.CTkButton(btn_row, text="❌  Dismiss / Ignore", fg_color="#DC3545",
                      hover_color="#C82333", width=160, height=38, command=dismiss).pack(side="left", padx=5)

    def pick_member_popup(self, callback):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Select Member")
        dialog.geometry("500x650")
        dialog.attributes("-topmost", True)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text="Select Existing Member", font=("Arial", 16, "bold")).pack(pady=15)

        search_f = ctk.CTkFrame(dialog, fg_color="transparent")
        search_f.pack(fill="x", padx=20, pady=(0, 10))
        
        search_e = ctk.CTkEntry(search_f, placeholder_text="Search by name or code…", width=440)
        search_e.pack(padx=20, pady=5)

        scroll = ctk.CTkScrollableFrame(dialog, fg_color="#F8F9FA", corner_radius=8, height=450)
        scroll.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        def refresh_list(_e=None):
            for w in scroll.winfo_children():
                w.destroy()
            
            q = search_e.get().strip().lower()
            conn = sqlite3.connect("database/attendance.db")
            if q:
                rows = conn.execute("""
                    SELECT member_code, name, type 
                    FROM members 
                    WHERE name LIKE ? OR member_code LIKE ? 
                    ORDER BY member_code""", (f"%{q}%", f"%{q}%")).fetchall()
            else:
                rows = conn.execute("SELECT member_code, name, type FROM members ORDER BY member_code").fetchall()
            conn.close()

            if not rows:
                ctk.CTkLabel(scroll, text="No members found matching search.", font=("Arial", 12), text_color="gray").pack(pady=20)
                return

            for code, name, mtype in rows:
                row = ctk.CTkFrame(scroll, fg_color="#FFFFFF", height=50, corner_radius=6)
                row.pack(fill="x", pady=3, padx=5)
                row.pack_propagate(False)

                ctk.CTkLabel(row, text=f"{name} ({code})", font=("Arial", 12, "bold")).pack(side="left", padx=12)
                ctk.CTkLabel(row, text=mtype, font=("Arial", 10), text_color="gray").pack(side="left", padx=10)
                
                def select(c=code):
                    callback(c)
                    dialog.destroy()

                ctk.CTkButton(row, text="Select", width=70, height=30, command=lambda c=code: select(c)).pack(side="right", padx=10)

        search_e.bind("<KeyRelease>", refresh_list)
        refresh_list()

    # ── Member CRUD ───────────────────────────────────────────────────────────

    def add_member_popup(self):
        self.member_dialog("Add New Member")

    def on_view_member(self, code):
        self.member_dialog("Member Details", code, readonly=True)

    def on_edit_member(self, code):
        self.member_dialog("Edit Member", code)

    def on_delete_member(self, code):
        if messagebox.askyesno("Delete", f"Delete member {code}?"):
            conn = sqlite3.connect("database/attendance.db")
            conn.execute("DELETE FROM members WHERE member_code=?", (code,))
            conn.commit()
            conn.close()
            self.refresh_member_table()

    def _date_picker(self, parent, label, existing_val="", readonly=False, default_today=False, on_change=None):
        """Wheel-style date picker. Returns a callable get() → 'DD-MM-YYYY' or ''."""
        ctk.CTkLabel(parent, text=label, font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 0))
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", pady=3)

        val_var = ctk.StringVar(value=existing_val)
        
        # If default requested and empty
        if default_today and not existing_val:
            val_var.set(datetime.now().strftime("%d-%m-%Y"))

        entry = ctk.CTkEntry(row, textvariable=val_var, state="readonly", height=38)
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        def open_picker():
            if readonly: return
            
            def on_ok_callback(v):
                val_var.set(v)
                if on_change: on_change(v)
                
            WheelDatePicker(self, f"Select {label}", initial_val=val_var.get(), 
                            on_ok=on_ok_callback)

        btn = ctk.CTkButton(row, text="📅", width=45, height=38, fg_color="#F3F4F6", text_color="#374151", 
                            hover_color="#E5E7EB", command=open_picker)
        btn.pack(side="right")
        
        if readonly:
            btn.configure(state="disabled")

        def get_val():
            return val_var.get()

        def set_val(val):
            val_var.set(val)

        return get_val, set_val

    def delete_custom_field(self, field_name, parent_dialog):
        if messagebox.askyesno("Delete Custom Field", f"Are you sure you want to delete the field '{field_name}'?\n\nWARNING: This will permanently remove this field and ALL its data for ALL members in the database.", parent=parent_dialog):
            try:
                conn = sqlite3.connect("database/attendance.db")
                conn.execute(f"ALTER TABLE members DROP COLUMN {field_name}")
                conn.commit()
                conn.close()
                messagebox.showinfo("Success", f"Field '{field_name}' has been deleted. Please close and reopen the dialog to see changes.", parent=parent_dialog)
                parent_dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete field: {e}", parent=parent_dialog)

    def member_dialog(self, title, code=None, readonly=False):
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("520x780")
        dialog.attributes("-topmost", True)
        dialog.grab_set()

        scroll = ctk.CTkScrollableFrame(dialog, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=20)

        # Load existing
        existing = {}
        db_cols = []
        conn = sqlite3.connect("database/attendance.db")
        try:
            # Get columns dynamically
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(members)")
            db_cols = [row[1].lower() for row in cursor.fetchall()]
            
            if code:
                # Use pandas for easy dict conversion
                df = pd.read_sql("SELECT * FROM members WHERE member_code=?", conn, params=[code])
                if not df.empty:
                    # Normalize keys to lowercase for robust matching
                    existing = {str(k).lower(): v for k, v in df.iloc[0].to_dict().items()}
        except: pass
        finally: conn.close()

        standard_fields = ["member_code", "name", "type", "age", "dob", "baptism_date", 
                           "address", "email", "phone", "has_holy_spirit", "image_path", 
                           "registration_date", "area", "remark", "age_category", "title"]
        extra_fields = [c for c in db_cols if c not in standard_fields]

        # ── Profile Photo ──────────────────────────────────────────────────────
        self.dialog_img_path = existing.get("image_path", "")

        photo_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        photo_frame.pack(pady=(5, 10), fill="x")

        photo_lbl = ctk.CTkLabel(photo_frame, text="📷 No Photo", width=120, height=120, fg_color="#E9ECEF", corner_radius=10)
        photo_lbl.pack(side="left", padx=10)

        def update_photo_preview(path):
            if path and os.path.exists(path):
                try:
                    pil = Image.open(path).resize((120, 120))
                    ci = ctk.CTkImage(light_image=pil, dark_image=pil, size=(120, 120))
                    photo_lbl.configure(image=ci, text="")
                    photo_lbl.image = ci
                    self.dialog_img_path = path
                except: pass

        if self.dialog_img_path:
            update_photo_preview(self.dialog_img_path)

        btn_frame = ctk.CTkFrame(photo_frame, fg_color="transparent")
        btn_frame.pack(side="left", padx=10)

        def browse_photo():
            p = filedialog.askopenfilename(filetypes=[("Image", "*.png *.jpg *.jpeg")])
            if p: update_photo_preview(p)

        def capture_photo():
            if hasattr(self, "last_frame") and self.last_frame is not None:
                os.makedirs(os.path.join("records", "unknown"), exist_ok=True)
                p = os.path.join("records", "unknown", f"temp_cap_{datetime.now().strftime('%H%M%S')}.jpg")
                cv2.imwrite(p, self.last_frame)
                update_photo_preview(p)
            else:
                messagebox.showwarning("Error", "Camera not active. Please ensure the dashboard camera is running.", parent=dialog)

        if not readonly:
            ctk.CTkButton(btn_frame, text="Browse File", width=120, command=browse_photo).pack(pady=5)
            ctk.CTkButton(btn_frame, text="Take Photo", width=120, command=capture_photo).pack(pady=5)

        # ── Name ──────────────────────────────────────────────────────────────
        ctk.CTkLabel(scroll, text="Name", font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 0))
        name_e = ctk.CTkEntry(scroll, width=400)
        name_e.insert(0, existing.get("name", ""))
        if readonly: name_e.configure(state="disabled")
        name_e.pack(pady=4)

        # ── Title selection ───────────────────────────────────────────────────
        ctk.CTkLabel(scroll, text="Title", font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 0))
        title_var = ctk.StringVar(value=existing.get("title", "") or "")
        title_cb = ctk.CTkComboBox(scroll, variable=title_var,
                                    values=["", "Brother", "Sister"],
                                    width=400, state="disabled" if readonly else "normal")
        title_cb.pack(pady=4)

        # ── Type dropdown ─────────────────────────────────────────────────────
        ctk.CTkLabel(scroll, text="Type", font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 0))
        type_var = ctk.StringVar(value=existing.get("type", "Area Member") or "Area Member")
        type_cb  = ctk.CTkComboBox(scroll, variable=type_var,
                                    values=["Area Member", "Other Area Member", "Truth Seeker"],
                                    width=400, state="disabled" if readonly else "normal")
        type_cb.pack(pady=4)

        def update_age_cat_ui(dob):
            if not dob or dob == "--": 
                age_cat_var.set("")
                return
            try:
                birth = datetime.strptime(dob, '%d-%m-%Y')
                age = (date.today() - birth.date()).days // 365
                if age <= 12: cat = "Child"
                elif age <= 24: cat = "Youth"
                elif age <= 64: cat = "Adult"
                else: cat = "Elder"
                age_cat_var.set(cat)
            except: age_cat_var.set("")

        # ── Date of Birth ──────────────────────────────────────────────────────
        get_dob, set_dob = self._date_picker(scroll, "Date of Birth", existing.get("dob", ""), readonly, on_change=update_age_cat_ui)

        # ── Age Category ───────────────────────────────────────────────────────
        ctk.CTkLabel(scroll, text="Age Category", font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 0))
        age_cat_var = ctk.StringVar(value=existing.get("age_category", ""))
        age_cat_cb = ctk.CTkComboBox(scroll, variable=age_cat_var, width=400,
                                      values=["", "Child", "Youth", "Adult", "Elder"],
                                      state="disabled" if readonly else "normal")
        age_cat_cb.pack(pady=4)

        # ── Date of Baptism ──────────────────────────────────────────────────
        if not readonly or self.is_admin:
            get_bap, set_bap = self._date_picker(scroll, "Date of Baptism", existing.get("baptism_date", ""), readonly)

        # ── Area ──────────────────────────────────────────────────────────────
        ctk.CTkLabel(scroll, text="Area", font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 0))
        area_e = ctk.CTkEntry(scroll, width=400,
                               placeholder_text=self.settings.get("default_area", "") or "Enter area")
        area_e.insert(0, existing.get("area", "") or self.settings.get("default_area", ""))
        if readonly: area_e.configure(state="disabled")
        area_e.pack(pady=4)

        # ── Address / Email / Phone ────────────────────────────────────────────
        s_entries = {}
        if not readonly or self.is_admin:
            simple_fields = [("Address", "address"), ("Email", "email"), ("Phone", "phone")]
            for lbl_txt, key in simple_fields:
                ctk.CTkLabel(scroll, text=lbl_txt, font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 0))
                e = ctk.CTkEntry(scroll, width=400)
                e.insert(0, existing.get(key, "") or "")
                if readonly: e.configure(state="disabled")
                e.pack(pady=4)
                s_entries[key] = e

            # ── Holy Spirit ───────────────────────────────────────────────────────
            ctk.CTkLabel(scroll, text="Holy Spirit?", font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 0))
            hs_var   = tk.BooleanVar(value=bool(existing.get("has_holy_spirit", False)))
            hs_check = ctk.CTkCheckBox(scroll, text="Received Holy Spirit", variable=hs_var)
            if readonly: hs_check.configure(state="disabled")
            hs_check.pack(pady=4)

        # ── Remark ────────────────────────────────────────────────────────────
        if not readonly or self.is_admin:
            ctk.CTkLabel(scroll, text="Remark", font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 0))
            remark_e = ctk.CTkTextbox(scroll, width=400, height=70)
            remark_e.insert("1.0", existing.get("remark", "") or "")
            if readonly: remark_e.configure(state="disabled")
            remark_e.pack(pady=4)

        # ── Extra Fields (SQL Custom) ──
        extra_entries = {}
        if extra_fields:
            ctk.CTkFrame(scroll, height=2, fg_color="#E5E7EB").pack(fill="x", pady=20)
            ctk.CTkLabel(scroll, text="🛡️ Extra Information (Custom SQL Fields)", font=("Arial", 13, "bold"), text_color="#6366F1").pack(anchor="w", pady=(0, 10))
            
            for f_name in extra_fields:
                f_hdr = ctk.CTkFrame(scroll, fg_color="transparent")
                f_hdr.pack(fill="x", pady=(8, 0))
                ctk.CTkLabel(f_hdr, text=f_name.replace("_", " ").upper(), font=("Arial", 11, "bold")).pack(side="left")
                if not readonly and self.is_admin:
                    btn_del = ctk.CTkButton(f_hdr, text="Delete Field", fg_color="#EF4444", hover_color="#DC2626", width=80, height=24, font=("Arial", 10, "bold"), command=lambda f=f_name: self.delete_custom_field(f, dialog))
                    btn_del.pack(side="right")
                
                ent = ctk.CTkEntry(scroll, width=400)
                ent.insert(0, str(existing.get(f_name, "")) if existing.get(f_name) is not None else "")
                ent.pack(pady=4)
                extra_entries[f_name] = ent
                if readonly: ent.configure(state="disabled")

        if not readonly:
            def save():
                data = {
                    "name":          name_e.get().strip(),
                    "type":          type_var.get(),
                    "title":         title_var.get(),
                    "dob":           get_dob(),
                    "baptism_date":  get_bap(),
                    "area":          area_e.get().strip(),
                    "address":       s_entries["address"].get(),
                    "email":         s_entries["email"].get(),
                    "phone":         s_entries["phone"].get(),
                    "has_holy_spirit": hs_var.get(),
                    "image_path":    self.dialog_img_path,
                    "remark":        remark_e.get("1.0", "end").strip(),
                    "age_category":  age_cat_var.get(),
                }
                # Collect extra fields
                for f_name, ent in extra_entries.items():
                    data[f_name] = ent.get().strip()
                if not data["name"]:
                    messagebox.showwarning("Missing", "Name is required.", parent=dialog)
                    return
                    
                # DUPLICATE NAME CHECK FOR NEW MEMBER
                if not code:
                    conn = sqlite3.connect("database/attendance.db")
                    exists_name = conn.execute("SELECT member_code FROM members WHERE name = ? COLLATE NOCASE", (data["name"],)).fetchone()
                    conn.close()
                    if exists_name:
                        if not messagebox.askyesno("Duplicate Name", f"A member with the name '{data['name']}' already exists (ID: {exists_name[0]}).\n\nDo you want to proceed and create another record with the same name?", parent=dialog):
                            return
                            
                # Use prefix if brand new member
                prefix = self.settings.get("member_prefix", "") if not code else ""
                try:
                    # Collect title directly from widget to be sure
                    data["title"] = title_cb.get()
                    self.backend.register_member(data, force_code=code, prefix=prefix)
                except Exception as e:
                    messagebox.showerror("Save Error", f"Backend failure: {str(e)}", parent=dialog)
                    return
                
                self.refresh_member_table()
                dialog.destroy()

            ctk.CTkButton(scroll, text="💾  Save", fg_color="#28A745",
                          width=200, height=42, command=save).pack(pady=22)
        else:
            ctk.CTkButton(scroll, text="Close", width=140, command=dialog.destroy).pack(pady=22)


# ── WiFi Camera Network Discovery Helpers ───────────────────────────────────

COMMON_PATHS = [
    "", 
    "/live",
    "/h264Preview_01_main",
    "/onvif-media1",
    "/stream1",
    "/1",
    "/h264",
    "/ch1",
    "/cam/realmonitor?channel=1&subtype=0",
    ":8080/video",
    ":8080/videofeed"
]

def discover_cameras():
    ips = set()
    probe_uuid = uuid.uuid4()
    ws_discovery_soap = f"""<?xml version="1.0" encoding="utf-8"?>
<Envelope xmlns:tds="http://www.onvif.org/ver10/device/wsdl" xmlns="http://www.w3.org/2003/05/soap-envelope" xmlns:wsa="http://schemas.xmlsoap.org/ws/2004/08/addressing">
  <Header>
    <wsa:MessageID>urn:uuid:{probe_uuid}</wsa:MessageID>
    <wsa:To>urn:schemas-xmlsoap-org:ws:2004:08:addressing</wsa:To>
    <wsa:Action>http://schemas.xmlsoap.org/ws/2005/04/discovery/Probe</wsa:Action>
  </Header>
  <Body>
    <Probe xmlns="http://schemas.xmlsoap.org/ws/2005/04/discovery">
      <Types>tds:Device</Types>
    </Probe>
  </Body>
</Envelope>"""

    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
    s.settimeout(1.5)
    s.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_TTL, 4)
    try:
        s.sendto(ws_discovery_soap.encode('utf-8'), ('239.255.255.250', 3702))
        while True:
            try:
                data, addr = s.recvfrom(65535)
                ips.add(addr[0])
            except socket.timeout:
                break
    except Exception:
        pass
    finally:
        s.close()

    ssdp_request = (
        "M-SEARCH * HTTP/1.1\r\n"
        "HOST: 239.255.255.250:1900\r\n"
        "MAN: \"ssdp:discover\"\r\n"
        "MX: 2\r\n"
        "ST: ssdp:all\r\n"
        "\r\n"
    )
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM, socket.IPPROTO_UDP)
    s.settimeout(1.5)
    s.setsockopt(socket.IPPROTO_IP, socket.IP_MULTICAST_TTL, 4)
    try:
        s.sendto(ssdp_request.encode('utf-8'), ('239.255.255.250', 1900))
        while True:
            try:
                data, addr = s.recvfrom(65535)
                data_str = data.decode('utf-8', errors='ignore').lower()
                if any(k in data_str for k in ('camera', 'cam', 'video', 'rtsp', 'onvif', 'ipc')):
                    ips.add(addr[0])
            except socket.timeout:
                break
    except Exception:
        pass
    finally:
        s.close()
        
    return list(ips)

def test_camera_url(url):
    import cv2
    cap = cv2.VideoCapture(url)
    if cap.isOpened():
        cap.release()
        return url
    return None


if __name__ == "__main__":
    for d in ["database", "cache", "records", "reports", "registered_faces", "backup"]:
        os.makedirs(d, exist_ok=True)
    app = AutoAttendanceApp()
    app.mainloop()
