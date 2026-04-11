import cv2
import sqlite3
import pickle
import os
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
        self.camera = cv2.VideoCapture(camera_id)
        self.camera.set(cv2.CAP_PROP_FRAME_WIDTH,  1920)
        self.camera.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)

        # InsightFace (Portable: root='./models' lets you keep the AI files in the project folder)
        self.face_app = insightface.app.FaceAnalysis(name='buffalo_l', root='./models')
        self.face_app.prepare(ctx_id=-1, det_size=(640, 640))

        self.known_face_encodings = []
        self.known_face_names     = []
        self.known_face_ids       = []
        self.load_known_faces()

        self.init_database()

        self.active_session_id  = None
        self.session_captured_ids = set()   # member codes captured this session
        self.session_unknown_encodings = [] # encodings of unknown faces this session
        self.frame_count          = 0
        self.process_every_n_frames = 5

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

        # sessions migrations
        c.execute("PRAGMA table_info(sessions)")
        scols = [row[1] for row in c.fetchall()]
        if 'target_count' not in scols:
            c.execute("ALTER TABLE sessions ADD COLUMN target_count INTEGER DEFAULT 0")

        # Optimization Indexes
        c.execute("CREATE INDEX IF NOT EXISTS idx_attendance_sid ON attendance(session_id)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_attendance_time ON attendance(check_in_time DESC)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_members_name ON members(name)")

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

    def bulk_export_archive(self, selected_codes, out_path):
        """Export members and their face photos into a .zip file."""
        if not selected_codes: return False
        
        temp_dir = "temp_export"
        os.makedirs(temp_dir, exist_ok=True)
        img_dir = os.path.join(temp_dir, "photos")
        os.makedirs(img_dir, exist_ok=True)
        
        conn = sqlite3.connect(self.db_path)
        placeholders = ",".join(["?"] * len(selected_codes))
        members = pd.read_sql(f"SELECT * FROM members WHERE member_code IN ({placeholders})", conn, params=selected_codes)
        conn.close()
        
        # Save JSON data
        members.to_json(os.path.join(temp_dir, "members.json"), orient="records", indent=4)
        
        # Copy photos
        for _, m in members.iterrows():
            code, name = m['member_code'], m['name']
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
            
            # Prepare data row (column order matching sqlite init)
            vals = (code, m['name'], m['type'], m['age'], m['dob'], m['baptism_date'],
                    m['address'], m['email'], m['phone'], m['has_holy_spirit'],
                    m['image_path'], m['registration_date'], m['area'])
            
            if exists:
                c.execute("""UPDATE members SET name=?, type=?, age=?, dob=?, baptism_date=?, 
                           address=?, email=?, phone=?, has_holy_spirit=?, image_path=?, 
                           registration_date=?, area=? WHERE member_code=?""", 
                        (vals[1], vals[2], vals[3], vals[4], vals[5], vals[6], vals[7], 
                         vals[8], vals[9], vals[10], vals[11], vals[12], code))
                updated += 1
            else:
                c.execute("""INSERT INTO members (member_code, name, type, age, dob, baptism_date, 
                           address, email, phone, has_holy_spirit, image_path, registration_date, area)
                           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", vals)
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

        conn = sqlite3.connect(self.db_path)
        c    = conn.cursor()
        exists = c.execute("SELECT 1 FROM members WHERE member_code=?", (code,)).fetchone()

        if exists:
            c.execute('''UPDATE members SET name=?,type=?,age=?,dob=?,baptism_date=?,
                         address=?,email=?,phone=?,has_holy_spirit=?,area=?,image_path=? WHERE member_code=?''',
                      (data.get('name',''), data.get('type','Member'), age,
                       data.get('dob'), data.get('baptism_date'),
                       data.get('address'), data.get('email'), data.get('phone'),
                       1 if data.get('has_holy_spirit') else 0,
                       data.get('area', '').strip(), data.get('image_path', ''), code))
        else:
            c.execute('''INSERT INTO members
                (member_code,name,type,age,dob,baptism_date,address,email,phone,
                 has_holy_spirit,image_path,registration_date,area)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (code, data.get('name',''), data.get('type','Member'), age,
                 data.get('dob'), data.get('baptism_date'),
                 data.get('address'), data.get('email'), data.get('phone'),
                 1 if data.get('has_holy_spirit') else 0,
                 data.get('image_path'), date.today(),
                 data.get('area', '').strip()))

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

    def start_session(self, title, duration_mins=None):
        now  = datetime.now()
        conn = sqlite3.connect(self.db_path)
        c    = conn.cursor()
        c.execute("INSERT INTO sessions (title,date,start_time,duration_mins) VALUES (?,?,?,?)",
                  (title, now.date(), now, duration_mins))
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

    def mark_attendance(self, name, member_code, frame, m_type='member', bbox=None):
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
            (name, member_code, m_type.lower(), attendance_id))
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

        if self.frame_count % self.process_every_n_frames == 0:
            small  = cv2.resize(frame, (0, 0), fx=0.5, fy=0.5)
            faces  = self.face_app.get(small)

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
                    # Look up member type
                    conn   = sqlite3.connect(self.db_path)
                    row    = conn.execute("SELECT type FROM members WHERE member_code=?",
                                         (match_code,)).fetchone()
                    conn.close()
                    m_type = (row[0].lower() if row else 'member')

                    # Always add to results for visual display
                    results.append({'name': match_name, 'code': match_code,
                                    'bbox': bbox, 'new': False,
                                    'img': None, 'type': m_type})

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
        total_members = conn.execute("SELECT COUNT(*) FROM members").fetchone()[0]

        # Area-specific member count
        if default_area and default_area.strip():
            da = default_area.strip().lower()
            area_total = conn.execute(
                "SELECT COUNT(*) FROM members WHERE LOWER(TRIM(area))=?",
                (da,)).fetchone()[0]
        else:
            da = None
            area_total = total_members

        if self.active_session_id:
            sid = self.active_session_id
            df  = pd.read_sql("""
                SELECT a.id,
                       COALESCE(m.name,        a.person_name)  AS name,
                       COALESCE(m.age,         0)              AS age,
                       COALESCE(m.member_code, a.member_code)  AS member_code,
                       COALESCE(m.type,        a.status)       AS type,
                       COALESCE(m.area,        '')             AS area,
                       a.record_image, a.check_in_time,        a.status
                FROM   attendance a
                LEFT JOIN members m ON a.member_code = m.member_code
                WHERE  a.session_id = ?
                ORDER  BY a.check_in_time DESC
            """, conn, params=[sid])

            p_members = int((df['status'].str.lower() == 'member').sum())   if not df.empty else 0
            p_gospel  = int(df['status'].str.lower().str.contains('gospel', na=False).sum()) if not df.empty else 0
            waiting   = int((df['status'].str.lower() == 'unknown').sum())  if not df.empty else 0
            present   = p_members + p_gospel

            # Area rate: people from default area present / total area members
            if da and not df.empty:
                area_present = int(df['area'].apply(lambda x: str(x).strip().lower() == da if x else False).sum())
            else:
                area_present = present
            area_rate    = (area_present  / area_total    * 100) if area_total    > 0 else 0
            overall_rate = (present       / total_members * 100) if total_members > 0 else 0
        else:
            present = p_members = p_gospel = waiting = area_present = 0
            area_rate = overall_rate = 0.0
            df = pd.DataFrame()

        conn.close()
        return {
            "p_total":      present,
            "p_members":    p_members,
            "p_gospel":     p_gospel,
            "waiting":      waiting,
            "area_rate":    area_rate,
            "overall_rate": overall_rate,
            "list":         df,
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
