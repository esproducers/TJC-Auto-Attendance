import os
os.environ["OPENCV_LOG_LEVEL"] = "SILENT"
import cv2
import sqlite3
import pickle
import shutil
import zipfile
import json
import numpy as np
import pandas as pd
import time
import uuid
from datetime import datetime, date
import insightface
from PIL import Image

COLOR_CYAN = (255, 255, 0) # BGR

class InsightFaceAttendance:
    def _create_video_capture(self, source):
        """
        Creates a VideoCapture object supporting integer camera IDs (local USB webcams)
        and string URLs (RTSP / HTTP WiFi IP cameras) with optimized network flags.
        """
        if isinstance(source, str) and source.isdigit():
            source = int(source)

        if isinstance(source, str):
            # Enforce TCP transport for RTSP streams to eliminate UDP packet drops/corruption
            os.environ["OPENCV_FFMPEG_CAPTURE_OPTIONS"] = "rtsp_flags;tcp"

            # Attempt 1: Direct capture
            cap = cv2.VideoCapture(source)
            if cap.isOpened():
                return cap

            # Attempt 2: Force FFmpeg backend
            cap = cv2.VideoCapture(source, cv2.CAP_FFMPEG)
            if cap.isOpened():
                return cap

            # Attempt 3: Strip default port :554 if present
            if ":554" in source:
                alt_url = source.replace(":554", "")
                cap = cv2.VideoCapture(alt_url, cv2.CAP_FFMPEG)
                if cap.isOpened():
                    return cap

            return cap
        else:
            return cv2.VideoCapture(source)

    def __init__(self, camera_id=0, face_dir="registered_faces",
                 db_path="database/attendance.db",
                 cache_file="cache/church_faces_insight.pkl",
                 process_width=640,        # 新增：处理宽度（自动缩放至此）
                 flip_mode=0):             # 新增：画面方向修正 0=正常,1=水平镜像,2=垂直翻转,3=180°
        self.face_dir    = face_dir
        self.db_path     = db_path
        self.cache_file  = cache_file
        self.records_dir = os.path.join("records", "attendance")
        self.unknown_dir = os.path.join("records", "unknown")

        for d in [self.face_dir,
                  os.path.dirname(self.db_path),
                  os.path.dirname(self.cache_file),
                  self.records_dir, self.unknown_dir, "reports"]:
            os.makedirs(d, exist_ok=True)

        # ---- Camera Initialization (Supports USB Webcams & RTSP/WiFi Streams) ----
        self.current_camera_id = camera_id
        self.camera = self._create_video_capture(camera_id)
        if not self.camera.isOpened():
            print(f"[WARN] Unable to open initial camera stream: {camera_id}")

        # Only apply MJPEG fourcc to local USB cameras
        if isinstance(camera_id, int) or (isinstance(camera_id, str) and camera_id.isdigit()):
            try:
                fourcc = cv2.VideoWriter_fourcc('M','J','P','G')
                self.camera.set(cv2.CAP_PROP_FOURCC, fourcc)
            except Exception:
                pass

        try:
            w = int(self.camera.get(cv2.CAP_PROP_FRAME_WIDTH))
            h = int(self.camera.get(cv2.CAP_PROP_FRAME_HEIGHT))
            if w <= 0 or h <= 0:
                w, h = 640, 480
        except Exception:
            w, h = 640, 480

        print(f"[CAM] Camera Resolution: {w}x{h}")

        # 设定处理宽度（所有输入图像都会缩放到此宽度再识别）
        self.process_width = process_width
        self.original_width = w
        self.original_height = h

        # 画面方向修正（应对倒装/镜像）
        self.flip_mode = flip_mode

        # InsightFace（加载模型，det_size 设为 320x320 加速）
        self.face_app = None
        self.is_prepared = False
        
        # 基础属性
        self.active_session_id  = None
        self.session_captured_ids = set()
        self.session_captured_names = set()
        self.session_unknown_encodings = []
        self.pending_unknowns = {}
        self.frame_count = 0
        self.process_every_n_frames = 5   # 每5帧处理一次（可根据性能调整）
        
        # 背光补偿模式（默认关闭）
        self.bright_light_mode = False
        self.bright_light_params = {
            "gain": 1.0,
            "contrast": 1.0,
            "saturation": 1.0,
            "white_balance": 0.0
        }
        
        self.known_face_encodings = []
        self.known_face_names = []
        self.known_face_ids = []

    # ---------- 背光补偿 ----------
    def set_bright_light_mode(self, enabled: bool):
        """启用/关闭背光补偿（只走纯软件图像增强，不篡改硬件曝光参数以防止卡死）"""
        self.bright_light_mode = enabled
        status_str = "ENABLED" if enabled else "DISABLED"
        print(f"[CAM] 背光补偿模式: {status_str}")

    def set_bright_light_params(self, params: dict):
        """Update manual adjustment settings (gain, contrast, saturation, white_balance) for Bright Light WDR mode."""
        if hasattr(self, 'bright_light_params') and isinstance(params, dict):
            self.bright_light_params.update(params)

    def apply_smart_bright_light_compensation(self, frame):
        """
        极速通用背光补偿（仅处理亮度通道，速度极快）
        在检测帧对小图调用，不占用主循环性能
        """
        if frame is None or not self.bright_light_mode:
            return frame
        try:
            # 转为 HSV，只调整 V（明度）通道
            hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)
            h, s, v = cv2.split(hsv)
            
            # 如果整体偏暗（均值<80），做 Gamma 提亮
            mean_val = np.mean(v)
            if mean_val < 80:
                v_float = v.astype(np.float32) / 255.0
                gamma = 0.6   # 提亮系数（越小越亮）
                v_corrected = np.power(v_float, gamma) * 255.0
                v = np.clip(v_corrected, 0, 255).astype(np.uint8)
            
            enhanced = cv2.merge((h, s, v))
            return cv2.cvtColor(enhanced, cv2.COLOR_HSV2BGR)
        except Exception:
            return frame

    # ---------- 摄像头切换 ----------
    def switch_camera(self, camera_id):
        """切换摄像头（支持 RTSP 字符串和整数 ID）"""
        if hasattr(self, 'camera') and self.camera:
            self.camera.release()
            
        self.current_camera_id = camera_id
        self.camera = self._create_video_capture(camera_id)
        if not self.camera.isOpened():
            print(f"[CAM ERROR] Failed to open camera stream: {camera_id}")
            return False
            
        # 只有本地摄像头才尝试设置 MJPEG
        if isinstance(camera_id, int) or (isinstance(camera_id, str) and camera_id.isdigit()):
            try:
                fourcc = cv2.VideoWriter_fourcc('M','J','P','G')
                self.camera.set(cv2.CAP_PROP_FOURCC, fourcc)
            except Exception:
                pass
            
        try:
            w = int(self.camera.get(cv2.CAP_PROP_FRAME_WIDTH))
            h = int(self.camera.get(cv2.CAP_PROP_FRAME_HEIGHT))
            if w <= 0 or h <= 0:
                w, h = 640, 480
            print(f"[CAM] 切换摄像头分辨率: {w}x{h}")
            self.original_width = w
            self.original_height = h
        except Exception:
            pass

        if self.bright_light_mode:
            self.set_bright_light_mode(True)
        return True

    # ---------- 准备模型 ----------
    def prepare(self, ctx_id=-1, det_size=(320, 320)):   # det_size 改为 320
        """初始化 InsightFace 模型（使用 320x320 加速）"""
        if self.face_app is None:
            self.face_app = insightface.app.FaceAnalysis(name='buffalo_l', root='./models', providers=['CPUExecutionProvider'])
        self.face_app.prepare(ctx_id=ctx_id, det_size=det_size)  # 不设 max_num，允许多人
        self.is_prepared = True
        
        self.known_face_encodings = []
        self.known_face_names     = []
        self.known_face_ids       = []
        self.init_database()
        self.load_known_faces()

    # ── Face cache ────────────────────────────────────────────────────────────

    def load_known_faces(self, force_rebuild=False):
        """Smart incremental face cache loader.
        Only computes embeddings for new/modified files instead of rebuilding all 120+ faces."""
        if not os.path.exists(self.face_dir):
            os.makedirs(self.face_dir, exist_ok=True)

        current_files = {f: os.path.getmtime(os.path.join(self.face_dir, f)) 
                         for f in os.listdir(self.face_dir) 
                         if f.lower().endswith(('.jpg', '.jpeg', '.png'))}

        active_member_codes = set()
        all_db_member_codes = set()
        if os.path.exists(self.db_path):
            try:
                conn = sqlite3.connect(self.db_path)
                rows = conn.execute("SELECT member_code, is_disabled FROM members").fetchall()
                conn.close()
                for code, dis in rows:
                    if code:
                        all_db_member_codes.add(code)
                        if not dis:
                            active_member_codes.add(code)
            except Exception as e:
                print(f"[CACHE] DB active members query error: {e}")

        # Clean up orphan face photo files for members deleted from database
        for fn in list(current_files.keys()):
            stem = os.path.splitext(fn)[0]
            p_id = stem.split('_', 1)[0] if '_' in stem else stem
            if all_db_member_codes and p_id and p_id not in all_db_member_codes:
                try:
                    os.remove(os.path.join(self.face_dir, fn))
                    print(f"[CACHE] Removed orphan face photo of deleted member: {fn}")
                    del current_files[fn]
                except Exception as e:
                    print(f"[WARN] Failed to delete orphan photo {fn}: {e}")

        cache_map = {}
        if not force_rebuild and os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'rb') as f:
                    data = pickle.load(f)
                if isinstance(data, dict) and 'file_map' in data:
                    cache_map = data['file_map']
                elif isinstance(data, dict) and 'encodings' in data and 'names' in data and 'ids' in data:
                    # Upgrade legacy cache format to file_map format if sizes match
                    legacy_names = data.get('names', [])
                    legacy_encs = data.get('encodings', [])
                    legacy_ids = data.get('ids', [])
                    if len(legacy_names) == len(current_files):
                        disk_fns = sorted(list(current_files.keys()))
                        for i, fn in enumerate(disk_fns):
                            cache_map[fn] = {
                                'enc': legacy_encs[i],
                                'name': legacy_names[i],
                                'id': legacy_ids[i],
                                'mtime': current_files[fn]
                            }
            except Exception as e:
                print(f"[CACHE] Read cache error: {e}")
                cache_map = {}

        updated_map = {}
        added_count = 0
        reused_count = 0

        for fn, mtime in current_files.items():
            # If present in cache and timestamp matches, reuse embedding directly
            if fn in cache_map and abs(cache_map[fn].get('mtime', 0) - mtime) < 1e-3:
                updated_map[fn] = cache_map[fn]
                reused_count += 1
            else:
                # Compute embedding ONLY for this single new or updated image file
                stem = os.path.splitext(fn)[0]
                p_id, p_name = (stem.split('_', 1) if '_' in stem else ('', stem))
                img_path = os.path.join(self.face_dir, fn)
                img = cv2.imread(img_path)
                if img is not None:
                    faces = self.face_app.get(img)
                    if faces:
                        updated_map[fn] = {
                            'enc': faces[0].embedding,
                            'name': p_name,
                            'id': p_id,
                            'mtime': mtime
                        }
                        added_count += 1

        enc, names, ids = [], [], []
        for fn in sorted(updated_map.keys()):
            item = updated_map[fn]
            # Exclude disabled members or deleted members from active face recognition
            if not active_member_codes or item['id'] in active_member_codes:
                enc.append(item['enc'])
                names.append(item['name'])
                ids.append(item['id'])

        self.known_face_encodings = enc
        self.known_face_names     = names
        self.known_face_ids       = ids

        # Matrix Vectorization: pre-normalize encodings into a 2D NumPy array matrix (N, 512)
        if enc:
            mat = np.array(enc, dtype=np.float32)
            norms = np.linalg.norm(mat, axis=1, keepdims=True)
            norms[norms == 0] = 1e-9
            self.known_matrix = mat / norms
        else:
            self.known_matrix = np.empty((0, 512), dtype=np.float32)

        # In-memory Member Metadata Cache to avoid disk SQL queries inside per-frame loop
        self.known_member_meta = {}
        if os.path.exists(self.db_path):
            try:
                conn = sqlite3.connect(self.db_path)
                rows = conn.execute("SELECT member_code, type, title FROM members WHERE (is_disabled = 0 OR is_disabled IS NULL)").fetchall()
                conn.close()
                for code, mtype, title in rows:
                    self.known_member_meta[code] = {
                        'type': (mtype.lower() if mtype else 'area member'),
                        'title': (title if title else '')
                    }
            except Exception as e:
                print(f"[CACHE] Metadata cache error: {e}")

        try:
            with open(self.cache_file, 'wb') as f:
                pickle.dump({
                    'encodings': enc,
                    'names': names,
                    'ids': ids,
                    'file_map': updated_map
                }, f)
        except Exception as e:
            print(f"[CACHE] Write cache error: {e}")

        if added_count > 0:
            print(f"[CACHE] Incremental update: processed {added_count} new/modified face(s), kept {reused_count} cached face(s). Total: {len(enc)}")
        else:
            print(f"[CACHE] Loaded {reused_count} faces from cache (no rebuild required).")

    # ── Database ──────────────────────────────────────────────────────────────

    def init_database(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()

        c.execute('''CREATE TABLE IF NOT EXISTS members (
            member_code TEXT PRIMARY KEY,
            name TEXT, type TEXT, age INTEGER,
            dob TEXT, baptism_date TEXT,
            address TEXT, email TEXT, phone TEXT,
            has_holy_spirit INTEGER,
            image_path TEXT, registration_date DATE)''')

        c.execute('''CREATE TABLE IF NOT EXISTS sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT, date DATE,
            start_time TIMESTAMP, end_time TIMESTAMP,
            duration_mins INTEGER,
            target_count INTEGER DEFAULT 0)''')

        c.execute('''CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            person_name TEXT, member_code TEXT,
            session_id INTEGER, record_image TEXT,
            check_in_time TIMESTAMP, service_date DATE,
            status TEXT DEFAULT "member")''')

        c.execute('''CREATE TABLE IF NOT EXISTS org_charts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            year INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')

        c.execute('''CREATE TABLE IF NOT EXISTS org_chart_roles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chart_id INTEGER,
            parent_role_id INTEGER,
            role_name TEXT,
            member_code TEXT,
            FOREIGN KEY(chart_id) REFERENCES org_charts(id))''')

        # Migrations
        c.execute("PRAGMA table_info(attendance)")
        cols = [r[1] for r in c.fetchall()]
        if 'session_id' not in cols:
            c.execute("ALTER TABLE attendance ADD COLUMN session_id INTEGER")
        if 'status' not in cols:
            c.execute("ALTER TABLE attendance ADD COLUMN status TEXT DEFAULT 'member'")

        c.execute("PRAGMA table_info(members)")
        mcols = [r[1] for r in c.fetchall()]
        if 'area' not in mcols:
            c.execute("ALTER TABLE members ADD COLUMN area TEXT DEFAULT ''")
        if 'remark' not in mcols:
            c.execute("ALTER TABLE members ADD COLUMN remark TEXT DEFAULT ''")
        if 'age_category' not in mcols:
            c.execute("ALTER TABLE members ADD COLUMN age_category TEXT DEFAULT ''")
        if 'title' not in mcols:
            c.execute("ALTER TABLE members ADD COLUMN title TEXT DEFAULT ''")
        if 'reu_class' not in mcols:
            c.execute("ALTER TABLE members ADD COLUMN reu_class TEXT DEFAULT ''")
        if 'is_disabled' not in mcols:
            c.execute("ALTER TABLE members ADD COLUMN is_disabled INTEGER DEFAULT 0")
        if 'disable_remark' not in mcols:
            c.execute("ALTER TABLE members ADD COLUMN disable_remark TEXT DEFAULT ''")
        
        # Master Data table creation & initial seed
        c.execute('''CREATE TABLE IF NOT EXISTS master_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT,
            item_value TEXT,
            item_order INTEGER DEFAULT 0,
            is_default INTEGER DEFAULT 0)''')
        
        c.execute("SELECT COUNT(*) FROM master_data")
        if c.fetchone()[0] == 0:
            initial_seeds = [
                ('title', 'Brother', 1, 1),
                ('title', 'Sister', 2, 0),
                ('title', 'Preacher', 3, 0),
                ('title', 'Preceptor', 4, 0),
                ('title', 'Deacon', 5, 0),
                ('title', 'Deaconess', 6, 0),
                ('type', 'Area Member', 1, 1),
                ('type', 'Other Area Member', 2, 0),
                ('type', 'Truth Seeker', 3, 0),
                ('reu_class', 'N/A', 1, 1),
                ('reu_class', 'Junior Youth (JY)', 2, 0),
                ('reu_class', 'Upper Primary (UP)', 3, 0),
                ('reu_class', 'Lower Primary (LP)', 4, 0),
                ('age_category', 'Adult', 1, 1),
                ('age_category', 'Youth', 2, 0),
                ('age_category', 'Child', 3, 0),
                ('age_category', 'Senior', 4, 0)
            ]
            c.executemany("INSERT INTO master_data (category, item_value, item_order, is_default) VALUES (?, ?, ?, ?)", initial_seeds)
        
        # One-time fix: set age_category to "" if DOB is empty
        c.execute("UPDATE members SET age_category='' WHERE dob IS NULL OR dob='' OR dob='--'")

        # sessions migrations
        c.execute("PRAGMA table_info(sessions)")
        scols = [row[1] for row in c.fetchall()]
        if 'target_count' not in scols:
            c.execute("ALTER TABLE sessions ADD COLUMN target_count INTEGER DEFAULT 0")
        if 'seminar_type' not in scols:
            c.execute("ALTER TABLE sessions ADD COLUMN seminar_type TEXT DEFAULT 'Other'")

        # Optimization Indexes
        c.execute("CREATE INDEX IF NOT EXISTS idx_attendance_sid ON attendance(session_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_attendance_time ON attendance(check_in_time DESC)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_members_name ON members(name)")

        # One-time migration: change old 'Member' / 'member' type to 'Area Member'
        c.execute("UPDATE members SET type = 'Area Member' WHERE type = 'Member' OR type = 'member'")
        c.execute("UPDATE attendance SET status = 'Area Member' WHERE status = 'Member' OR status = 'member'")

        conn.commit()
        conn.close()

    # ── Members ───────────────────────────────────────────────────────────────

    def get_next_member_code(self, prefix=""):
        """Generate next formatted code like SK-0001 or just 0001."""
        conn = sqlite3.connect(self.db_path)
        p_str = f"{prefix}-" if prefix else ""
        
        # If prefix exists, search for codes starting with that prefix
        if prefix:
            res = conn.execute("SELECT member_code FROM members WHERE member_code LIKE ? ORDER BY member_code DESC LIMIT 1", (f"{prefix}-%",)).fetchone()
        else:
            # Fallback to numeric or any max code
            res = conn.execute("SELECT member_code FROM members ORDER BY member_code DESC LIMIT 1").fetchone()
        conn.close()

        if res:
            try:
                last_code = res[0]
                if prefix and "-" in last_code:
                    num_part = last_code.split("-")[-1]
                    return f"{prefix}-{int(num_part) + 1:04d}"
                else:
                    return f"{int(last_code) + 1:04d}"
            except: pass
            
        return f"{prefix}-0001" if prefix else "0001"

    def bulk_export_archive(self, selected_codes, out_path, fields=None):
        """Export members and their face photos into a .zip file. 
           If fields is provided, only those columns (plus member_code) are exported."""
        if not selected_codes: return False
        
        temp_dir = "temp_export"
        os.makedirs(temp_dir, exist_ok=True)
        img_dir = os.path.join(temp_dir, "photos")
        os.makedirs(img_dir, exist_ok=True)
        
        conn = sqlite3.connect(self.db_path)
        placeholders = ",".join(["?"] * len(selected_codes))
        members = pd.read_sql(f"SELECT * FROM members WHERE member_code IN ({placeholders})", conn, params=selected_codes)
        conn.close()

        # Filter fields if requested
        if fields:
            # We must ALWAYS keep member_code for sync identification
            keep = set(fields)
            keep.add('member_code')
            # Remove fields that don't exist in DB (like 'photo' which is handled separately)
            db_cols = [c for c in keep if c in members.columns]
            members = members[db_cols]
        
        # Save JSON data
        members.to_json(os.path.join(temp_dir, "members.json"), orient="records", indent=4)
        
        # Copy photos ONLY if 'photo' is in fields or if fields is None
        if fields is None or 'photo' in fields:
            for _, m in members.iterrows():
                code = m['member_code']
                # Find photo in registered_faces
                for fn in os.listdir(self.face_dir):
                    if fn.startswith(f"{code}_"):
                        shutil.copy(os.path.join(self.face_dir, fn), os.path.join(img_dir, fn))
                        break
        
        # Zip it up
        shutil.make_archive(out_path.replace(".zip", ""), 'zip', temp_dir)
        shutil.rmtree(temp_dir)
        return True

    def bulk_import_archive(self, zip_path):
        """Extract zip, upsert members, and copy photos."""
        temp_dir = "temp_import"
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
        
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            
        json_p = os.path.join(temp_dir, "members.json")
        if not os.path.exists(json_p):
            shutil.rmtree(temp_dir)
            return False, "Invalid migration file: members.json missing."
            
        with open(json_p, 'r') as f:
            members_data = json.load(f)
            
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        
        added, updated = 0, 0
        for m in members_data:
            code = m['member_code']
            exists = c.execute("SELECT 1 FROM members WHERE member_code=?", (code,)).fetchone()
            
            # Prepare data row
            vals = (code, m['name'], m['type'], m['age'], m['dob'], m['baptism_date'],
                    m['address'], m['email'], m['phone'], m['has_holy_spirit'],
                    m['image_path'], m['registration_date'], m['area'], m.get('remark', ''),
                    m.get('title', ''))
            
            if exists:
                c.execute("""UPDATE members SET name=?, type=?, age=?, dob=?, baptism_date=?, 
                           address=?, email=?, phone=?, has_holy_spirit=?, image_path=?, 
                           registration_date=?, area=?, remark=?, title=? WHERE member_code=?""", 
                        (vals[1], vals[2], vals[3], vals[4], vals[5], vals[6], vals[7], 
                         vals[8], vals[9], vals[10], vals[11], vals[12], vals[13], vals[14], code))
                updated += 1
            else:
                c.execute("""INSERT INTO members (member_code, name, type, age, dob, baptism_date, 
                           address, email, phone, has_holy_spirit, image_path, registration_date, area, remark, title)
                           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", vals)
                added += 1
        
        conn.commit()
        conn.close()
        
        # Copy photos
        img_dir = os.path.join(temp_dir, "photos")
        if os.path.exists(img_dir):
            for fn in os.listdir(img_dir):
                shutil.copy(os.path.join(img_dir, fn), os.path.join(self.face_dir, fn))
        
        # Cleanup
        shutil.rmtree(temp_dir)
        self.load_known_faces()
        
        return True, f"Import Finished: {added} added, {updated} updated."

    def bulk_import_excel(self, excel_path, prefix="TJC"):
        """Import members from an Excel file, skipping duplicate names."""
        try:
            df = pd.read_excel(excel_path)
            
            # Map expected columns to DB keys
            col_map = {
                "name": "name", "title": "title", "type": "type", "area": "area",
                "age category": "age_category", "dob": "dob", "date of baptism": "baptism_date",
                "phone": "phone", "email": "email", "address": "address", 
                "holy spirit?": "has_holy_spirit", "remark": "remark"
            }
            
            # Normalize column names in df for easier mapping
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            added = 0
            skipped = 0
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            for _, row in df.iterrows():
                # Extract data based on mapped columns
                data = {}
                for df_col, db_col in col_map.items():
                    if df_col in df.columns:
                        val = row[df_col]
                        if pd.isna(val):
                            val = ""
                        # Handle specific types
                        if db_col == "has_holy_spirit":
                            # Convert Yes/No or True/False to 1/0
                            str_val = str(val).strip().lower()
                            data[db_col] = 1 if str_val in ['1', 'yes', 'y', 'true', 't'] else 0
                        elif isinstance(val, (datetime, pd.Timestamp)):
                            data[db_col] = val.strftime('%d-%m-%Y')
                        else:
                            data[db_col] = str(val).strip()
                
                # Check Name
                name = data.get("name", "")
                if not name:
                    continue # Skip empty names
                
                # Check for duplicates
                exists = cursor.execute("SELECT 1 FROM members WHERE name = ? COLLATE NOCASE", (name,)).fetchone()
                if exists:
                    skipped += 1
                    continue
                
                # Add extra fields (any columns not in the map)
                for df_col in df.columns:
                    if df_col not in col_map and df_col != "member_code" and df_col != "registration_date":
                        val = row[df_col]
                        if pd.isna(val): val = ""
                        data[df_col.replace(' ', '_')] = str(val).strip()

                # Save member using existing function to handle IDs and database insertion
                self.register_member(data, prefix=prefix)
                added += 1
                
            conn.close()
            return True, f"Excel Import Finished: {added} added, {skipped} skipped (duplicate names)."
            
        except Exception as e:
            return False, f"Failed to parse Excel file: {str(e)}"

    def register_member(self, data, force_code=None, prefix=""):
        """Insert or update a member. Returns the member code used."""
        code = force_code if force_code else self.get_next_member_code(prefix=prefix)

        age = 0
        dob_str = data.get('dob', '')
        if dob_str:
            try:
                parts = dob_str.split('-')
                if len(parts) == 3:
                    if len(parts[0]) == 2:          # DD-MM-YYYY
                        birth = datetime.strptime(dob_str, '%d-%m-%Y')
                    else:                            # YYYY-MM-DD
                        birth = datetime.strptime(dob_str, '%Y-%m-%d')
                    age = (date.today() - birth.date()).days // 365
            except Exception:
                pass

        # Determine Age Category
        age_cat = data.get('age_category', '').strip()
        if dob_str and dob_str != "--":
            if age <= 12: age_cat = "Child"
            elif age <= 24: age_cat = "Youth"
            elif age <= 64: age_cat = "Adult"
            else: age_cat = "Elder"

        conn = sqlite3.connect(self.db_path)
        c    = conn.cursor()
        exists = c.execute("SELECT 1 FROM members WHERE member_code=?", (code,)).fetchone() # FOUND IT

        if exists:
            # Fetch existing data to avoid overwriting with defaults
            conn.row_factory = sqlite3.Row
            # Use conn.execute directly or recreate cursor
            old = conn.execute("SELECT * FROM members WHERE member_code=?", (code,)).fetchone()
            old = dict(old) if old else {}
            
            final_data = old.copy()
            # Update only if provided in 'data'
            if 'name' in data: final_data['name'] = data['name']
            if 'type' in data: final_data['type'] = data['type']
            if 'dob' in data: 
                final_data['dob'] = data['dob']
                final_data['age'] = age
            if 'baptism_date' in data: final_data['baptism_date'] = data['baptism_date']
            if 'address' in data: final_data['address'] = data['address']
            if 'email' in data: final_data['email'] = data['email']
            if 'phone' in data: final_data['phone'] = data['phone']
            if 'has_holy_spirit' in data: final_data['has_holy_spirit'] = 1 if data['has_holy_spirit'] else 0
            if 'area' in data: final_data['area'] = data['area'].strip()
            if 'image_path' in data: final_data['image_path'] = data['image_path']
            if 'remark' in data: final_data['remark'] = data['remark'].strip()
            if 'age_category' in data: final_data['age_category'] = data['age_category']
            if 'title' in data: final_data['title'] = data['title'].strip()
            if 'reu_class' in data: final_data['reu_class'] = data['reu_class'].strip()
            
            # Re-calculate age category if DOB changed
            if 'dob' in data:
                final_data['age_category'] = age_cat
                
            # Get all column names for extra fields
            c.execute("PRAGMA table_info(members)")
            db_cols = [row[1] for row in c.fetchall()]
            
            # Merge extra fields
            for k, v in data.items():
                if k in db_cols and k not in final_data and k != 'member_code' and k != 'registration_date':
                    final_data[k] = v
            
            final_data["title"] = data.get('title', '').strip()
            if 'reu_class' in data: final_data["reu_class"] = data['reu_class'].strip()
            
            sets = ", ".join([f"{k}=?" for k in final_data.keys()])
            vals = list(final_data.values())
            vals.append(code)
            c.execute(f"UPDATE members SET {sets} WHERE member_code=?", vals)
        else:
            c.execute("PRAGMA table_info(members)")
            db_cols = [row[1] for row in c.fetchall()]
            
            final_data = {
                "member_code": code, "registration_date": str(date.today()),
                "name": data.get('name', ''), "type": data.get('type', 'Area Member'),
                "age": age, "dob": data.get('dob'), "baptism_date": data.get('baptism_date'),
                "address": data.get('address'), "email": data.get('email'), "phone": data.get('phone'),
                "has_holy_spirit": 1 if data.get('has_holy_spirit') else 0,
                "area": data.get('area', '').strip(), "image_path": data.get('image_path', ''),
                "remark": data.get('remark', '').strip(), "age_category": age_cat,
                "title": data.get('title', '')
            }
            # Merge extra fields
            for k, v in data.items():
                if k in db_cols and k not in final_data:
                    final_data[k] = v
            
            # Ensure all DB columns are present
            for col in db_cols:
                if col not in final_data:
                    final_data[col] = None
            
            final_data["title"] = data.get('title', '').strip()
            
            cols_str = ", ".join(final_data.keys())
            places = ", ".join(["?"] * len(final_data))
            c.execute(f"INSERT INTO members ({cols_str}) VALUES ({places})", list(final_data.values()))

        conn.commit()
        conn.close()


        # Copy face photo and update cache incrementally
        img_path = data.get('image_path', '')
        if img_path and os.path.exists(img_path):
            # Sanitize name for filename safety
            safe_name = "".join([c for c in data.get('name','') if c.isalnum() or c in (' ','_')]).strip()
            ext = os.path.splitext(img_path)[1] or ".jpg"
            new_path = os.path.join(self.face_dir, f"{code}_{safe_name}{ext}")

            # Remove old photos for this code if overwriting/replacing photo
            if os.path.exists(self.face_dir):
                for fn in os.listdir(self.face_dir):
                    if fn.startswith(f"{code}_") and os.path.join(self.face_dir, fn) != new_path:
                        try:
                            os.remove(os.path.join(self.face_dir, fn))
                        except Exception:
                            pass
            try:
                shutil.copy2(img_path, new_path)
                self.load_known_faces()
            except Exception as e:
                print(f"[WARN] Photo copy error: {e}")

        return code

    def delete_member(self, code):
        """Permanently delete member from database and remove face photo. History attendance is kept intact."""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("DELETE FROM members WHERE member_code=?", (code,))
        conn.commit()
        conn.close()

        # Delete photo file(s) from registered_faces
        if os.path.exists(self.face_dir):
            for fn in os.listdir(self.face_dir):
                if fn.startswith(f"{code}_"):
                    try:
                        os.remove(os.path.join(self.face_dir, fn))
                    except Exception as e:
                        print(f"[WARN] Failed to delete face photo {fn}: {e}")

        # Update face cache incrementally so recognition updates immediately without hanging
        self.load_known_faces()
        return True

    def disable_member(self, code, remark=""):
        """Disable a member (e.g. passed away, moved to another country). Excludes from recognition & active counts."""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("UPDATE members SET is_disabled = 1, disable_remark = ? WHERE member_code = ?", (remark.strip(), code))
        conn.commit()
        conn.close()

        # Reload faces incrementally so disabled face is removed from recognition memory
        self.load_known_faces()
        return True

    def enable_member(self, code):
        """Re-enable a disabled member. Restores face recognition & active count."""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("UPDATE members SET is_disabled = 0 WHERE member_code = ?", (code,))
        conn.commit()
        conn.close()

        # Reload faces incrementally so face is included back in recognition memory
        self.load_known_faces()
        return True

    # ── Sessions ──────────────────────────────────────────────────────────────

    def start_session(self, title, duration_mins=None, seminar_type='Other'):
        now  = datetime.now()
        conn = sqlite3.connect(self.db_path)
        c    = conn.cursor()
        c.execute("INSERT INTO sessions (title,date,start_time,duration_mins,seminar_type) VALUES (?,?,?,?,?)",
                  (title, now.date(), now, duration_mins, seminar_type))
        self.active_session_id = c.lastrowid
        conn.commit()
        conn.close()
        self.session_captured_ids = set()
        self.session_unknown_encodings = []
        self.pending_unknowns = {} # map of uid -> {enc, first_seen, last_seen, best_score, best_frame, bbox}
        self.frame_count = 0
        self.session_captured_names = set()
        print(f"[SESSION] Started '{title}' id={self.active_session_id}")
        return self.active_session_id

    def end_session(self):
        if self.active_session_id:
            conn = sqlite3.connect(self.db_path)
            conn.execute("UPDATE sessions SET end_time=? WHERE id=?",
                         (datetime.now(), self.active_session_id))
            conn.commit()
            conn.close()
            print(f"[SESSION] Ended id={self.active_session_id}")
        self.active_session_id    = None
        self.session_captured_ids = set()
        self.session_captured_names = set()
        self.session_unknown_encodings = []

    # ── Attendance marking ────────────────────────────────────────────────────

    def mark_attendance(self, name, member_code, frame, m_type='Area Member', bbox=None):
        """Returns (new_capture: bool, save_path: str)"""
        if not self.active_session_id:
            return False, None
        if member_code and member_code in self.session_captured_ids:
            return False, None
            
        # Case-insensitive name check to prevent "Chin Khim Fung" vs "CHIN KHIM FUNG" duplicates
        name_lower = name.strip().lower()
        if name_lower in self.session_captured_names:
            return False, None

        today   = date.today()
        day_dir = os.path.join(self.records_dir, str(today))
        os.makedirs(day_dir, exist_ok=True)
        filename  = f"S{self.active_session_id}_{member_code}_{datetime.now().strftime('%H%M%S')}.jpg"
        save_path = os.path.join(day_dir, filename)

        # Draw annotation before saving
        if frame is not None:
            to_save = frame.copy()
            if bbox:
                cv2.rectangle(to_save, (bbox[0], bbox[1]), (bbox[2], bbox[3]), COLOR_CYAN, 2)
                cv2.putText(to_save, name, (bbox[0], bbox[1] - 10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.8, COLOR_CYAN, 2)

            ok = cv2.imwrite(save_path, to_save)
            if not ok:
                print(f"[WARN] imwrite failed: {save_path}")
                save_path = os.path.join("records", filename)
                cv2.imwrite(save_path, frame)
        else:
            # For manual entries, copy the profile image if m_type is provided as a path or use existing
            if m_type and os.path.exists(m_type) and m_type.endswith(('.jpg', '.png')):
                import shutil
                shutil.copy(m_type, save_path)
            else:
                save_path = "" # No image for this manual record

        conn = sqlite3.connect(self.db_path)
        conn.execute(
            "INSERT INTO attendance (person_name,member_code,session_id,record_image,"
            "check_in_time,service_date,status) VALUES (?,?,?,?,?,?,?)",
            (name, member_code, self.active_session_id, save_path,
             datetime.now(), today, m_type))
        conn.commit()
        conn.close()

        self.session_captured_ids.add(member_code)
        self.session_captured_names.add(name.strip().lower())
        print(f"[ATT] Marked: {name} ({member_code}) -> {save_path}")
        return True, save_path

    def save_unknown(self, frame, bbox=None):
        """Save unknown face. If bbox is provided, saves a square crop for easier identification."""
        if not self.active_session_id:
            return None

        today   = date.today()
        day_dir = os.path.join(self.unknown_dir, str(today))
        os.makedirs(day_dir, exist_ok=True)
        ts        = datetime.now().strftime('%H%M%S_%f')
        filename  = f"S{self.active_session_id}_unk_{ts}.jpg"
        save_path = os.path.join(day_dir, filename)

        # Better Human Identification: If we have a bbox, try to save a crop instead of full frame
        to_save = frame
        if bbox:
            try:
                x1, y1, x2, y2 = map(int, bbox)
                h, w = frame.shape[:2]
                # Add padding
                cx, cy = (x1+x2)//2, (y1+y2)//2
                side = int(max(x2-x1, y2-y1) * 1.5)
                if side > 5:
                    nx1, ny1 = max(0, cx-side//2), max(0, cy-side//2)
                    nx2, ny2 = min(w, nx1+side), min(h, ny1+side)
                if nx2 > nx1 and ny2 > ny1:
                        to_save = frame[ny1:ny2, nx1:nx2].copy()
            except Exception as e:
                print(f"[ERROR] Crop failed: {e}")
                to_save = frame.copy()

        # Draw box on unknown capture
        if bbox and to_save.shape[0] == frame.shape[0]: # Only if full frame
            cv2.rectangle(to_save, (bbox[0], bbox[1]), (bbox[2], bbox[3]), COLOR_CYAN, 2)
            cv2.putText(to_save, "Unknown", (bbox[0], bbox[1] - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.8, COLOR_CYAN, 2)

        ok = cv2.imwrite(save_path, to_save)
        if not ok:
            print(f"[WARN] imwrite failed: {save_path}")
            return None

        conn = sqlite3.connect(self.db_path)
        conn.execute(
            "INSERT INTO attendance (person_name,member_code,session_id,record_image,"
            "check_in_time,service_date,status) VALUES (?,?,?,?,?,?,?)",
            ('Unknown', None, self.active_session_id, save_path,
             datetime.now(), today, 'unknown'))
        conn.commit()
        conn.close()
        print(f"[UNK] Saved unknown: {save_path}")
        return save_path

    def identify_unknown(self, attendance_id, name, member_code, m_type):
        """Promote an 'unknown' attendance row to a real member."""
        conn = sqlite3.connect(self.db_path)
        conn.execute(
            "UPDATE attendance SET person_name=?,member_code=?,status=? WHERE id=?",
            (name, member_code, m_type, attendance_id))
        conn.commit()
        conn.close()
        
        # Prevent camera from capturing them again in this session
        if member_code:
            self.session_captured_ids.add(member_code)
        if name:
            self.session_captured_names.add(name.strip().lower())

    # ── Strict Unknown Quality Filter ─────────────────────────────────────────

    def is_quality_unknown_face(self, face, frame_w, frame_h, frame=None):
        """
        Strict quality filter for UNKNOWN face capture to prevent blurry, half, cut-off,
        or poor quality faces from cluttering the 'Waiting Recognition' list.
        
        Rules:
        1. High detection confidence: face.det_score >= 0.68
        2. Sufficient face size: width and height >= 70 pixels
        3. Full face (no cut-offs at borders): margins >= 20px from frame edges
        4. Frontal pose check: keypoints symmetry (eyes, nose, mouth present & inside box)
        5. Blur check: Laplacian variance >= 45.0 (if frame provided)
        """
        if face is None:
            return False

        # Rule 1: High detection score (ignore low confidence/blurry/side-profile detections)
        det_score = getattr(face, 'det_score', 0.0)
        if det_score < 0.68:
            return False

        # Rule 2 & 3: Border cut-off and face size check (on full frame resolution)
        bbox = face.bbox.astype(int) # [x1, y1, x2, y2]
        scale = frame_w / self.process_width if hasattr(self, 'process_width') and self.process_width > 0 else 1.0
        x1, y1, x2, y2 = (bbox * scale).astype(int)

        w = x2 - x1
        h = y2 - y1

        # Minimum size requirement for a clear face (reject tiny background faces)
        if w < 70 or h < 70:
            return False

        # Border margin check: Reject half-faces / cut-off faces at edge of camera view
        margin = 20
        if x1 <= margin or y1 <= margin or x2 >= (frame_w - margin) or y2 >= (frame_h - margin):
            return False

        # Rule 4: Facial Keypoints Pose Check (Ensure full frontal face with 2 eyes, nose, mouth visible)
        if hasattr(face, 'kps') and face.kps is not None and len(face.kps) == 5:
            kps = face.kps * scale
            left_eye, right_eye, nose, left_mouth, right_mouth = kps
            
            # Check all 5 keypoints are inside the bounding box
            for pt in [left_eye, right_eye, nose, left_mouth, right_mouth]:
                if not (x1 <= pt[0] <= x2 and y1 <= pt[1] <= y2):
                    return False
            
            # Eye distance check (ensure face is not severely turned sideways)
            eye_dist = np.linalg.norm(right_eye - left_eye)
            if eye_dist < (w * 0.25):
                return False

        # Rule 5: Blur check (Laplacian variance) on face crop
        if frame is not None:
            try:
                cx1 = max(0, x1)
                cy1 = max(0, y1)
                cx2 = min(frame_w, x2)
                cy2 = min(frame_h, y2)
                
                face_crop = frame[cy1:cy2, cx1:cx2]
                if face_crop.size > 0:
                    gray_crop = cv2.cvtColor(face_crop, cv2.COLOR_BGR2GRAY)
                    lap_var = cv2.Laplacian(gray_crop, cv2.CV_64F).var()
                    if lap_var < 45.0: # Below 45 is considered out of focus or blurry
                        return False
            except Exception:
                pass

        return True

    # ── Frame processing (核心改动) ──────────────────────────────────────────

    def process_frame(self, frame):
        """返回 (annotated_frame, list_of_result_dicts)"""
        results = []
        if not self.is_prepared:
            return frame, results

        # ===== 1. 画面方向修正（应对倒装/镜像） =====
        if self.flip_mode == 1:
            frame = cv2.flip(frame, 1)   # 水平镜像
        elif self.flip_mode == 2:
            frame = cv2.flip(frame, 0)   # 垂直翻转
        elif self.flip_mode == 3:
            frame = cv2.rotate(frame, cv2.ROTATE_180)

        # ===== 2. 动态缩放到统一处理宽度 =====
        h, w = frame.shape[:2]
        if w != self.process_width:
            scale = self.process_width / w
            new_w = self.process_width
            new_h = int(h * scale)
            small = cv2.resize(frame, (new_w, new_h))
        else:
            scale = 1.0
            small = frame

        # ===== 3. 每隔 N 帧检测一次 =====
        if self.frame_count % self.process_every_n_frames == 0:
            # ---- 3a. 背光补偿（仅在小图上执行，极速） ----
            if self.bright_light_mode:
                small = self.apply_smart_bright_light_compensation(small)

            # ---- 3b. 人脸检测（基于小图） ----
            faces = self.face_app.get(small)

            for face in faces:
                embedding = face.embedding
                best_dist, match_name, match_code = float('inf'), "Unknown", ""

                # 矩阵向量化匹配（您原有的加速逻辑）
                if hasattr(self, 'known_matrix') and len(self.known_matrix) > 0:
                    emb_norm = embedding / (np.linalg.norm(embedding) + 1e-9)
                    cos_sims = np.dot(self.known_matrix, emb_norm)
                    best_idx = np.argmax(cos_sims)
                    best_sim = cos_sims[best_idx]
                    dist = 1.0 - best_sim
                    if dist < 0.5:   # 阈值放宽至 0.5（通用性更好）
                        best_dist = dist
                        match_name = self.known_face_names[best_idx]
                        match_code = self.known_face_ids[best_idx]
                else:
                    for i, enc in enumerate(self.known_face_encodings):
                        cos_sim = np.dot(embedding, enc) / (
                            np.linalg.norm(embedding) * np.linalg.norm(enc) + 1e-9)
                        dist = 1 - cos_sim
                        if dist < best_dist and dist < 0.5:
                            best_dist = dist
                            match_name = self.known_face_names[i]
                            match_code = self.known_face_ids[i]

                # ---- 3c. 坐标还原到原图尺寸 ----
                bbox = (face.bbox / scale).astype(int).tolist()

                if match_name != "Unknown":
                    meta = getattr(self, 'known_member_meta', {}).get(match_code, {})
                    m_type = meta.get('type', 'area member')
                    m_title = meta.get('title', '')
                    results.append({'name': match_name, 'code': match_code,
                                    'bbox': bbox, 'new': False,
                                    'img': None, 'type': m_type, 'title': m_title})
                    if match_code not in self.session_captured_ids:
                        new, img = self.mark_attendance(match_name, match_code, frame, m_type, bbox=bbox)
                        results[-1]['new'] = new
                        results[-1]['img'] = img
                    # 清理重叠的未知待处理项
                    to_del = []
                    for uid, p in self.pending_unknowns.items():
                        pb = p['bbox']
                        dist_center = ((bbox[0]+bbox[2])/2 - (pb[0]+pb[2])/2)**2 + ((bbox[1]+bbox[3])/2 - (pb[1]+pb[3])/2)**2
                        if dist_center < (max(bbox[2]-bbox[0], 50)**2):
                            to_del.append(uid)
                    for uid in to_del:
                        del self.pending_unknowns[uid]

                elif self.active_session_id:
                    # 严格人脸质量过滤：忽略半脸、边缘切边、模糊脸、低置信度脸
                    if not self.is_quality_unknown_face(face, w, h, frame=frame):
                        continue

                    # 未知人脸处理
                    is_saved = False
                    for u_enc in self.session_unknown_encodings:
                        sim = np.dot(embedding, u_enc) / (np.linalg.norm(embedding) * np.linalg.norm(u_enc) + 1e-9)
                        if (1 - sim) < 0.4:
                            is_saved = True
                            break
                    if is_saved:
                        continue

                    target_uid = None
                    for uid, p in self.pending_unknowns.items():
                        sim = np.dot(embedding, p['enc']) / (np.linalg.norm(embedding) * np.linalg.norm(p['enc']) + 1e-9)
                        if (1 - sim) < 0.4:
                            target_uid = uid
                            break

                    now = time.time()
                    if target_uid:
                        p = self.pending_unknowns[target_uid]
                        p['last_seen'] = now
                        if face.det_score > p['best_score']:
                            p['best_score'] = face.det_score
                            p['best_frame'] = frame.copy()
                            p['bbox'] = bbox
                        if now - p['first_seen'] > 1.2:
                            img = self.save_unknown(p['best_frame'], bbox=p['bbox'])
                            self.session_unknown_encodings.append(p['enc'])
                            results.append({'name': 'Unknown', 'code': '', 'bbox': p['bbox'],
                                            'new': True, 'img': img, 'type': 'unknown'})
                            del self.pending_unknowns[target_uid]
                    else:
                        uid = str(uuid.uuid4())
                        self.pending_unknowns[uid] = {
                            'enc': embedding, 'first_seen': now, 'last_seen': now,
                            'best_score': face.det_score, 'best_frame': frame.copy(),
                            'bbox': bbox
                        }

        # 清理过期的未知待处理项
        now = time.time()
        self.pending_unknowns = {uid: p for uid, p in self.pending_unknowns.items() if now - p['last_seen'] < 2.0}

        # ===== 4. 在原图上绘制识别框 =====
        for r in results:
            b = r['bbox']
            cv2.rectangle(frame, (b[0], b[1]), (b[2], b[3]), COLOR_CYAN, 2)
            cv2.putText(frame, r['name'], (b[0], b[1] - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, COLOR_CYAN, 2)

        self.frame_count += 1
        return frame, results

    # ── Summary & waiting list ────────────────────────────────────────────────

    def get_summary(self, default_area=None):
        conn = sqlite3.connect(self.db_path)
        # Denominator: Total 'Area Member' DB count
        area_total = conn.execute("SELECT COUNT(*) FROM members WHERE type='Area Member' AND (is_disabled = 0 OR is_disabled IS NULL)").fetchone()[0]

        if self.active_session_id:
            sid = self.active_session_id
            df  = pd.read_sql("""
                SELECT a.id,
                       COALESCE(m.name,        a.person_name)  AS name,
                       COALESCE(m.age_category, '')    AS age,
                       COALESCE(m.member_code, a.member_code)  AS member_code,
                       COALESCE(m.type,        a.status)       AS type,
                       COALESCE(m.area,        '')             AS area,
                       a.record_image, a.check_in_time,        a.status
                FROM   attendance a
                LEFT JOIN members m ON a.member_code = m.member_code
                WHERE  a.session_id = ?
                ORDER  BY a.check_in_time DESC
            """, conn, params=[sid])

            p_area_member = int(df['type'].str.lower().isin(['area member']).sum()) if not df.empty else 0
            p_other_member = int(df['type'].str.lower().isin(['other area member']).sum()) if not df.empty else 0
            p_truth = int(df['type'].str.lower().str.contains('truth', na=False).sum()) if not df.empty else 0
            waiting = int((df['type'].str.lower() == 'unknown').sum()) if not df.empty else 0
            present = p_area_member + p_other_member + p_truth

            area_rate    = (p_area_member  / area_total    * 100) if area_total    > 0 else 0
            overall_rate = (present       / area_total    * 100) if area_total    > 0 else 0
        else:
            present = p_area_member = p_other_member = p_truth = waiting = 0
            area_rate = overall_rate = 0.0
            df = pd.DataFrame()

        conn.close()
        return {
            "p_total":        present,
            "p_area_member":  p_area_member,
            "p_other_member": p_other_member,
            "p_truth":        p_truth,
            "waiting":        waiting,
            "area_rate":      area_rate,
            "overall_rate":   overall_rate,
            "list":           df,
        }

    def get_waiting_list(self, session_id=None):
        """Return rows (id, record_image, check_in_time) for unknown attendees."""
        sid  = session_id or self.active_session_id
        if not sid:
            return []
        conn = sqlite3.connect(self.db_path)
        rows = conn.execute("""
            SELECT id, record_image, check_in_time
            FROM   attendance
            WHERE  session_id=? AND status='unknown'
            ORDER  BY check_in_time DESC
        """, (sid,)).fetchall()
        conn.close()
        return rows

    def get_periodical_stats(self, period_type='weekly', seminar_filter='All Sessions', default_area=None, start_date=None, end_date=None, custom_date=None):
        """Aggregate stats for seminars by week, month, or year."""
        import sqlite3
        import pandas as pd
        conn = sqlite3.connect(self.db_path)
        
        area_member_db_count = conn.execute("SELECT COUNT(*) FROM members WHERE type='Area Member' AND (is_disabled = 0 OR is_disabled IS NULL)").fetchone()[0]

        if period_type == 'weekly':
            date_fmt = '%Y-W%W'
        elif period_type == 'monthly':
            date_fmt = '%Y-%m'
        else:
            date_fmt = '%Y'
            
        # Determine filter clause
        if seminar_filter == 'All Sessions':
            where_clause = ""
        elif seminar_filter == 'Fri & Sat':
            where_clause = "WHERE s.seminar_type IN ('Friday Seminar', 'Saturday Seminar')"
        else:
            where_clause = f"WHERE s.seminar_type = '{seminar_filter}'"
            
        if start_date and end_date:
            if where_clause: where_clause += f" AND s.date BETWEEN '{start_date}' AND '{end_date}'"
            else: where_clause = f"WHERE s.date BETWEEN '{start_date}' AND '{end_date}'"
        elif custom_date:
            if where_clause: where_clause += f" AND strftime('{date_fmt}', s.date) LIKE '{custom_date}%'"
            else: where_clause = f"WHERE strftime('{date_fmt}', s.date) LIKE '{custom_date}%'"
            
        query = f"""
            SELECT 
                strftime('{date_fmt}', s.date) as period,
                
                -- Stats
                COUNT(a.id) as present,
                SUM(CASE WHEN m.title = 'Brother' THEN 1 ELSE 0 END) as bro,
                SUM(CASE WHEN m.title = 'Sister' THEN 1 ELSE 0 END) as sis,
                SUM(CASE WHEN m.type = 'Area Member' THEN 1 ELSE 0 END) as area_present,
                SUM(CASE WHEN m.type = 'Other Area Member' THEN 1 ELSE 0 END) as other_mbr,
                SUM(CASE WHEN m.type LIKE '%Truth%' THEN 1 ELSE 0 END) as ts,
                COUNT(DISTINCT s.id) as sess_count

            FROM sessions s
            JOIN attendance a ON a.session_id = s.id
            LEFT JOIN members m ON a.member_code = m.member_code
            {where_clause}
            GROUP BY period
            ORDER BY period DESC
        """
        
        df = pd.read_sql(query, conn)
        
        if not df.empty:
            # Format Period strings for readability
            def format_p(p):
                try:
                    if period_type == 'weekly':
                        y, w = p.split('-W')
                        return f"Week {w}, {y}"
                    elif period_type == 'monthly':
                        dt = datetime.strptime(p, '%Y-%m')
                        return dt.strftime('%b %Y').upper()
                    return p
                except: return p
            df['period'] = df['period'].apply(format_p)

            # Rates
            df['overall_rate'] = (df['present'] / (area_member_db_count * df['sess_count'].replace(0, 1)) * 100).fillna(0)
            df['area_rate'] = (df['area_present'] / (area_member_db_count * df['sess_count'].replace(0, 1)) * 100).fillna(0)
            
        conn.close()
        return df

    def get_detailed_period_sessions(self, period_str, period_type='monthly', seminar_filter='All Sessions', default_area=None):
        """Fetch all individual sessions within a specific period (e.g., 'May 2026')."""
        import sqlite3
        import pandas as pd
        conn = sqlite3.connect(self.db_path)
        
        area_member_db_count = conn.execute("SELECT COUNT(*) FROM members WHERE type='Area Member' AND (is_disabled = 0 OR is_disabled IS NULL)").fetchone()[0]

        # Convert Period String back to SQL friendly match
        if period_type == 'weekly':
            try:
                import re
                match = re.search(r'Week (\d+), (\d+)', period_str)
                w, y = match.groups()
                sql_period = f"{y}-W{int(w):02d}"
                date_fmt = '%Y-W%W'
            except: return pd.DataFrame()
        elif period_type == 'monthly':
            try:
                dt = datetime.strptime(period_str, '%b %Y')
                sql_period = dt.strftime('%Y-%m')
                date_fmt = '%Y-%m'
            except: return pd.DataFrame()
        else:
            sql_period = period_str
            date_fmt = '%Y'

        if seminar_filter == 'All Sessions':
            where_clause = f"WHERE strftime('{date_fmt}', s.date) = ?"
        elif seminar_filter == 'Fri & Sat':
            where_clause = f"WHERE s.seminar_type IN ('Friday Seminar', 'Saturday Seminar') AND strftime('{date_fmt}', s.date) = ?"
        else:
            s_filter = "Other" if seminar_filter == "Other Sessions" else seminar_filter
            where_clause = f"WHERE s.seminar_type = '{s_filter}' AND strftime('{date_fmt}', s.date) = ?"

        query = f"""
            SELECT 
                s.date as Date,
                s.seminar_type as Type,
                COUNT(a.id) as Present,
                SUM(CASE WHEN m.title = 'Brother' THEN 1 ELSE 0 END) as Bro,
                SUM(CASE WHEN m.title = 'Sister' THEN 1 ELSE 0 END) as Sis,
                SUM(CASE WHEN m.type = 'Area Member' THEN 1 ELSE 0 END) as Area_Present,
                SUM(CASE WHEN m.type = 'Other Area Member' THEN 1 ELSE 0 END) as Other_Mbr,
                SUM(CASE WHEN m.type LIKE '%Truth%' THEN 1 ELSE 0 END) as TS
            FROM sessions s
            JOIN attendance a ON a.session_id = s.id
            LEFT JOIN members m ON a.member_code = m.member_code
            {where_clause}
            GROUP BY s.id
            ORDER BY s.date ASC
        """
        
        df = pd.read_sql(query, conn, params=[sql_period])
        
        if not df.empty:
            df['Area_Rate%'] = (df['Area_Present'] / (area_member_db_count if area_member_db_count > 0 else 1) * 100).fillna(0).round(1)
            df['Overall_Rate%'] = (df['Present'] / (area_member_db_count if area_member_db_count > 0 else 1) * 100).fillna(0).round(1)
            # Keep Area_Present, Other_Mbr, TS for PDF generation
            
        conn.close()
        return df

    # ---------- 运行循环 ----------
    def run(self):
        """简单运行循环，按 q 退出"""
        print("按 'q' 退出，按 's' 保存当前画面")
        while True:
            ret, frame = self.camera.read()
            if not ret:
                print("摄像头读取失败")
                break
            annotated, _ = self.process_frame(frame)
            cv2.imshow('Church Attendance', annotated)
            key = cv2.waitKey(1) & 0xFF
            if key == ord('q'):
                break
            elif key == ord('s'):
                cv2.imwrite(f"capture_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg", frame)
                print("画面已保存")
        self.camera.release()
        cv2.destroyAllWindows()


if __name__ == "__main__":
    # 实例化系统（可指定 process_width 和 flip_mode）
    system = InsightFaceAttendance(process_width=640, flip_mode=0)
    system.prepare(det_size=(320, 320))   # 关键：det_size 改为 320
    if len(system.known_face_encodings) == 0:
        print("请将会众照片放入 registered_faces 文件夹后重新运行。")
    else:
        # 如果门口背光严重，可启用背光补偿（根据实际情况决定）
        # system.set_bright_light_mode(True)
        try:
            system.run()
        except KeyboardInterrupt:
            print("\n程序已退出")