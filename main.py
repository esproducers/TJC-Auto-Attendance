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
    def __init__(self, camera_id=0, face_dir="registered_faces",
                 db_path="database/attendance.db",
                 cache_file="cache/church_faces_insight.pkl"):
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

        # Camera
        self.current_camera_id = camera_id
        self.camera = cv2.VideoCapture(self.current_camera_id)
        try:
            self.camera.set(cv2.CAP_PROP_FRAME_WIDTH,  1920)
            self.camera.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
        except Exception:
            pass

        # InsightFace (Portable: root='./models' lets you keep the AI files in the project folder)
        self.face_app = None
        self.is_prepared = False
        
        # Initialize basic attributes to prevent crashes before preparation
        self.active_session_id  = None
        self.session_captured_ids = set()
        self.session_captured_names = set()
        self.session_unknown_encodings = []
        self.pending_unknowns = {}
        self.frame_count = 0
        self.process_every_n_frames = 5
        self.known_face_encodings = []
        self.known_face_names = []
        self.known_face_ids = []

    def switch_camera(self, camera_id):
        """Release current camera and open a new one."""
        if hasattr(self, 'camera') and self.camera:
            self.camera.release()
        self.current_camera_id = camera_id
        self.camera = cv2.VideoCapture(self.current_camera_id)
        try:
            self.camera.set(cv2.CAP_PROP_FRAME_WIDTH,  1920)
            self.camera.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
        except Exception:
            pass
        return self.camera.isOpened()

    def prepare(self, ctx_id=-1, det_size=(640, 640)):
        """Prepare the FaceAnalysis app. Might download models if missing."""
        if self.face_app is None:
            self.face_app = insightface.app.FaceAnalysis(name='buffalo_l', root='./models')
        self.face_app.prepare(ctx_id=ctx_id, det_size=det_size)
        self.is_prepared = True
        
        self.known_face_encodings = []
        self.known_face_names     = []
        self.known_face_ids       = []
        self.load_known_faces()

        self.init_database()

    # ── Face cache ────────────────────────────────────────────────────────────

    def load_known_faces(self):
        # 1. Get current files on disk
        current_files = [f for f in os.listdir(self.face_dir) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
        
        # 2. Try loading cache
        if os.path.exists(self.cache_file):
            try:
                with open(self.cache_file, 'rb') as f:
                    data = pickle.load(f)
                
                # Validation: Does cache size match disk count?
                cached_names = data.get('names', [])
                if len(cached_names) == len(current_files):
                    self.known_face_encodings = data.get('encodings', [])
                    self.known_face_names     = cached_names
                    self.known_face_ids       = data.get('ids', [])
                    print(f"[CACHE] Loaded {len(self.known_face_names)} faces from {self.face_dir}")
                    return
                else:
                    print(f"[CACHE] Sync Mismatch: Disk({len(current_files)}) vs Cache({len(cached_names)})")
            except Exception as e:
                print(f"[CACHE] Load error: {e}")

        # 3. Rebuild cache if valid cache not found or mismatched
        print(f"[CACHE] Rebuilding cache for {len(current_files)} faces in {self.face_dir}...")

        enc, names, ids = [], [], []
        for fn in os.listdir(self.face_dir):
            if not fn.lower().endswith(('.jpg', '.jpeg', '.png')):
                continue
            stem = os.path.splitext(fn)[0]
            p_id, p_name = (stem.split('_', 1) if '_' in stem else ('', stem))
            img = cv2.imread(os.path.join(self.face_dir, fn))
            if img is not None:
                faces = self.face_app.get(img)
                if faces:
                    enc.append(faces[0].embedding)
                    names.append(p_name)
                    ids.append(p_id)

        self.known_face_encodings = enc
        self.known_face_names     = names
        self.known_face_ids       = ids
        if enc:
            with open(self.cache_file, 'wb') as f:
                pickle.dump({'encodings': enc, 'names': names, 'ids': ids}, f)
        print(f"[CACHE] Built cache with {len(enc)} faces")

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
        if os.path.exists(self.cache_file): os.remove(self.cache_file)
        self.load_known_faces()
        
        return True, f"Import Finished: {added} added, {updated} updated."

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


        # Copy face photo and rebuild cache
        img_path = data.get('image_path', '')
        if img_path and os.path.exists(img_path):
            # Sanitize name for filename safety
            safe_name = "".join([c for c in data.get('name','') if c.isalnum() or c in (' ','_')]).strip()
            ext = os.path.splitext(img_path)[1] or ".jpg"
            new_path = os.path.join(self.face_dir, f"{code}_{safe_name}{ext}")
            try:
                shutil.copy2(img_path, new_path)
                if os.path.exists(self.cache_file):
                    os.remove(self.cache_file)
                self.load_known_faces()
            except Exception as e:
                print(f"[WARN] Photo copy error: {e}")

        return code

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

    # ── Frame processing ──────────────────────────────────────────────────────

    def process_frame(self, frame):
        """Returns (annotated_frame, list_of_result_dicts)"""
        results = []
        if not self.is_prepared:
            return frame, results

        if self.frame_count % self.process_every_n_frames == 0:
            small = cv2.resize(frame, (0, 0), fx=0.5, fy=0.5)
            faces = self.face_app.get(small)

            for face in faces:
                embedding = face.embedding
                best_dist, match_name, match_code = float('inf'), "Unknown", ""

                for i, enc in enumerate(self.known_face_encodings):
                    cos_sim = np.dot(embedding, enc) / (
                        np.linalg.norm(embedding) * np.linalg.norm(enc) + 1e-9)
                    dist = 1 - cos_sim
                    if dist < best_dist and dist < 0.4:
                        best_dist  = dist
                        match_name = self.known_face_names[i]
                        match_code = self.known_face_ids[i]

                bbox = (face.bbox.astype(int) * 2).tolist()

                if match_name != "Unknown":
                    # Look up member type and title
                    conn   = sqlite3.connect(self.db_path)
                    row    = conn.execute("SELECT type, title FROM members WHERE member_code=?",
                                         (match_code,)).fetchone()
                    conn.close()
                    m_type  = (row[0].lower() if row else 'area member')
                    m_title = (row[1] if row else '')

                    # Always add to results for visual display
                    results.append({'name': match_name, 'code': match_code,
                                    'bbox': bbox, 'new': False,
                                    'img': None, 'type': m_type, 'title': m_title})

                    # Only perform attendance marking if not yet captured this session
                    if match_code not in self.session_captured_ids:
                        new, img = self.mark_attendance(match_name, match_code, frame, m_type, bbox=bbox)
                        results[-1]['new'] = new
                        results[-1]['img'] = img
                    
                    # Prevent overlapping unknown captures for this recognized face
                    to_del = []
                    for uid, p in self.pending_unknowns.items():
                        pb = p['bbox']
                        # Simple overlap check: center distance < box half-width
                        dist = ((bbox[0]+bbox[2])/2 - (pb[0]+pb[2])/2)**2 + ((bbox[1]+bbox[3])/2 - (pb[1]+pb[3])/2)**2
                        if dist < (max(bbox[2]-bbox[0], 50)**2):
                            to_del.append(uid)
                    for uid in to_del: del self.pending_unknowns[uid]
                elif self.active_session_id:
                    # 1. Check if already saved in this session
                    is_saved = False
                    for u_enc in self.session_unknown_encodings:
                        sim = np.dot(embedding, u_enc) / (np.linalg.norm(embedding) * np.linalg.norm(u_enc) + 1e-9)
                        if (1 - sim) < 0.4:
                            is_saved = True; break
                    if is_saved: continue

                    # 2. Check pending buffer
                    target_uid = None
                    for uid, p in self.pending_unknowns.items():
                        sim = np.dot(embedding, p['enc']) / (np.linalg.norm(embedding) * np.linalg.norm(p['enc']) + 1e-9)
                        if (1 - sim) < 0.4:
                            target_uid = uid; break

                    now = time.time()
                    if target_uid:
                        p = self.pending_unknowns[target_uid]
                        p['last_seen'] = now
                        # Update best frame if current one is clearer (higher detection score)
                        if face.det_score > p['best_score']:
                            p['best_score'] = face.det_score
                            p['best_frame'] = frame.copy()
                            p['bbox'] = bbox
                        
                        # If seen for > 1.2s, trigger final capture
                        if now - p['first_seen'] > 1.2:
                            img = self.save_unknown(p['best_frame'], bbox=p['bbox'])
                            self.session_unknown_encodings.append(p['enc'])
                            results.append({'name': 'Unknown', 'code': '', 'bbox': p['bbox'], 'new': True, 'img': img, 'type': 'unknown'})
                            del self.pending_unknowns[target_uid]
                    else:
                        # New pending entry
                        uid = str(uuid.uuid4())
                        self.pending_unknowns[uid] = {
                            'enc': embedding, 'first_seen': now, 'last_seen': now,
                            'best_score': face.det_score, 'best_frame': frame.copy(), 'bbox': bbox
                        }

        # Clean up stale unknowns (absent for > 2 seconds)
        now = time.time()
        self.pending_unknowns = {uid: p for uid, p in self.pending_unknowns.items() if now - p['last_seen'] < 2.0}

        # Draw bounding boxes
        for r in results:
            b     = r['bbox']
            color = COLOR_CYAN
            cv2.rectangle(frame, (b[0], b[1]), (b[2], b[3]), color, 2)
            cv2.putText(frame, r['name'], (b[0], b[1] - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)

        self.frame_count += 1
        return frame, results

    # ── Summary & waiting list ────────────────────────────────────────────────

    def get_summary(self, default_area=None):
        conn = sqlite3.connect(self.db_path)
        # Denominator: Total 'Area Member' DB count
        area_total = conn.execute("SELECT COUNT(*) FROM members WHERE type='Area Member'").fetchone()[0]

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
        
        area_member_db_count = conn.execute("SELECT COUNT(*) FROM members WHERE type='Area Member'").fetchone()[0]

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
        
        area_member_db_count = conn.execute("SELECT COUNT(*) FROM members WHERE type='Area Member'").fetchone()[0]

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
