import sqlite3
import pandas as pd
from datetime import date, datetime
from fpdf import FPDF
import os


class ReportGenerator:
    def __init__(self, db_path="database/attendance.db"):
        self.db_path = db_path
        os.makedirs("reports", exist_ok=True)

    def _get_session_stats(self, session_id, default_area=None):
        conn = sqlite3.connect(self.db_path)
        sess = conn.execute("""
            SELECT title, date, start_time, duration_mins, target_count 
            FROM sessions WHERE id=?
        """, (session_id,)).fetchone()
        if not sess:
            conn.close()
            return {}
        
        title, dt, start_ts, duration, target = sess
        
        # System totals
        total_sys_m = conn.execute("SELECT COUNT(*) FROM members").fetchone()[0]
        da = default_area.strip().lower() if default_area and default_area.strip() else None
        
        if da:
            area_total = conn.execute("SELECT COUNT(*) FROM members WHERE LOWER(TRIM(area))=?", (da,)).fetchone()[0]
        else:
            area_total = total_sys_m

        # Attendees
        df = pd.read_sql("""
            SELECT a.status, m.area
            FROM   attendance a
            LEFT JOIN members m ON a.member_code = m.member_code
            WHERE  a.session_id = ?
        """, conn, params=[session_id])
        conn.close()

        p_members = len(df[df['status'].str.lower() == 'member'])
        p_gospel  = len(df[df['status'].str.lower().str.contains('gospel', na=False)])
        p_total   = p_members + p_gospel
        waiting   = len(df[df['status'].str.lower() == 'unknown']) if not df.empty else 0

        # Area rate
        area_present = 0
        if da and not df.empty:
            area_present = len(df[(df['status'].str.lower() == 'member') & 
                                 (df['area'].fillna('').str.strip().str.lower() == da)])
        elif not da:
            area_present = p_members
        
        area_rate    = (area_present / area_total * 100) if area_total > 0 else 0
        overall_rate = (p_total / total_sys_m * 100) if total_sys_m > 0 else 0

        return {
            "Session Title": title,
            "Date": dt,
            "Start Time": str(start_ts)[11:16] if start_ts else "N/A",
            "Duration": f"{duration} mins" if duration else "N/A",
            "Target Total": target or 0,
            "Total Present": p_total,
            "Members": p_members,
            "Gospel Friends": p_gospel,
            "Waiting Recognition": waiting,
            "Area Rate %": f"{area_rate:.1f}%",
            "Overall Rate %": f"{overall_rate:.1f}%",
            "Area Present": area_present,
            "Area Total": area_total,
            "Total Sys Members": total_sys_m
        }

    def generate_excel(self, session_ids=None, out_path=None, summary=False, default_area=None):
        if isinstance(session_ids, int):
            session_ids = [session_ids]
            
        if summary and session_ids:
            rows = [self._get_session_stats(sid, default_area) for sid in session_ids]
            # Rearrange columns to put date second
            ordered_rows = []
            for r in rows:
                ordered_rows.append({
                    "Session Title": r["Session Title"],
                    "Date": r["Date"],
                    "Total Present": r["Total Present"],
                    "Members": r["Members"],
                    "Gospel Friends": r["Gospel Friends"],
                    "Waiting Recognition": r["Waiting Recognition"],
                    "Area Rate %": r["Area Rate %"],
                    "Overall Rate %": r["Overall Rate %"]
                })
            df = pd.DataFrame(ordered_rows)
            title = "Batch_Summary"
        else:
            conn = sqlite3.connect(self.db_path)
            if session_ids:
                ph = ",".join("?" * len(session_ids))
                df = pd.read_sql(f"""
                    SELECT COALESCE(m.member_code, a.member_code) AS ID,
                           COALESCE(m.name, a.person_name)        AS Name,
                           COALESCE(m.age, '')                    AS Age,
                           COALESCE(m.type, a.status)             AS Type,
                           a.check_in_time                        AS CheckIn,
                           a.status                               AS Status,
                           s.title                                AS Session_Title
                    FROM   attendance a
                    LEFT JOIN members m ON a.member_code = m.member_code
                    LEFT JOIN sessions s ON a.session_id = s.id
                    WHERE  a.session_id IN ({ph})
                    ORDER  BY a.check_in_time
                """, conn, params=session_ids)
                title = "Selected_Sessions"
            else:
                df = pd.read_sql("""
                    SELECT COALESCE(m.member_code, a.member_code) AS ID,
                           COALESCE(m.name, a.person_name)        AS Name,
                           COALESCE(m.age, '')                    AS Age,
                           a.check_in_time                        AS CheckIn,
                           a.status                               AS Status,
                           s.title                                AS Session_Title
                    FROM   attendance a
                    LEFT JOIN members m ON a.member_code = m.member_code
                    LEFT JOIN sessions s ON a.session_id = s.id
                    ORDER  BY a.check_in_time DESC
                    LIMIT  200
                """, conn)
                title = "All_Sessions"
            conn.close()

        safe_title = "".join(c for c in title if c.isalnum() or c in " _-")[:40]
        filename   = out_path if out_path else f"reports/Attendance_{safe_title}_{date.today()}.xlsx"
        df.to_excel(filename, index=False)
        return filename

    def _add_standard_header(self, pdf, settings):
        church_name = settings.get("church_name", "Attendance Report")
        address = settings.get("address", "")
        logo_path = settings.get("logo_path", "")
        def_area_label = settings.get("default_area", "")

        # Header (Logo=LEFT, Name/Address=CENTERED)
        if logo_path and os.path.exists(logo_path):
            pdf.image(logo_path, 20, 10, 30)
            
        pdf.set_y(10)
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, church_name, ln=True, align='C')
        pdf.set_font("Arial", size=12)
        pdf.cell(0, 7, f"{def_area_label}", ln=True, align='C')
        
        if address: 
            pdf.set_font("Arial", size=10)
            # Center address using full page width
            pdf.multi_cell(0, 5, address, align='C')
        
        pdf.ln(5)

    def generate_pdf(self, session_ids=None, out_path=None, summary=False, default_area=None, settings=None):
        if isinstance(session_ids, int):
            session_ids = [session_ids]
            
        settings = settings or {}

        if summary and session_ids:
            rows = [self._get_session_stats(sid, default_area) for sid in session_ids]
            title = "Batch Summary Report"
            
            pdf = FPDF()
            pdf.add_page()
            self._add_standard_header(pdf, settings)

            pdf.set_font("Arial", 'B', 18)
            pdf.cell(0, 15, title, ln=True, align='C')
            pdf.ln(5)

            # Table header
            pdf.set_fill_color(220, 230, 241)
            pdf.set_font("Arial", 'B', 9)
            cols = ["Session Title", "Date", "Total", "Member", "Gospel", "Area %", "Overall %"]
            w = [55, 25, 18, 18, 20, 27, 27]
            for i, c in enumerate(cols):
                pdf.cell(w[i], 10, c, 1, 0, 'C', True)
            pdf.ln()

            pdf.set_font("Arial", size=9)
            for r in rows:
                pdf.cell(w[0], 9, str(r["Session Title"])[:30], 1)
                pdf.cell(w[1], 9, str(r["Date"]), 1, 0, 'C')
                pdf.cell(w[2], 9, str(r["Total Present"]), 1, 0, 'C')
                pdf.cell(w[3], 9, str(r["Members"]), 1, 0, 'C')
                pdf.cell(w[4], 9, str(r["Gospel Friends"]), 1, 0, 'C')
                pdf.cell(w[5], 9, str(r["Area Rate %"]), 1, 0, 'C')
                pdf.cell(w[6], 9, str(r["Overall Rate %"]), 1, 0, 'C')
                pdf.ln()

            filename = out_path if out_path else f"reports/Batch_Report_{date.today()}.pdf"
            pdf.output(filename)
            return filename

        # --- Detailed Report (Redesigned based on mockup) ---
        if not session_ids: return None
        sid = session_ids[0] # Detailed report usually for one session
        stats = self._get_session_stats(sid, default_area)
        
        conn = sqlite3.connect(self.db_path)
        df = pd.read_sql("""
            SELECT COALESCE(m.member_code, a.member_code) AS ID,
                   COALESCE(m.name, a.person_name)        AS Name,
                   COALESCE(m.age, '')                    AS Age,
                   COALESCE(m.type, a.status)             AS Type,
                   m.area                                 AS Area,
                   a.check_in_time                        AS CheckIn,
                   a.status                               AS Status
            FROM   attendance a
            LEFT JOIN members m ON a.member_code = m.member_code
            WHERE  a.session_id = ?
            ORDER  BY a.check_in_time
        """, conn, params=[sid])
        conn.close()

        pdf = FPDF()
        pdf.add_page()
        self._add_standard_header(pdf, settings)
        
        # Main Title (Session Name)
        pdf.set_font("Arial", 'B', 18)
        pdf.cell(0, 15, stats["Session Title"], ln=True, align='C')
        pdf.ln(5)

        # 2. Two-Column Stats section
        start_y = pdf.get_y()
        right_x = 125
        
        def left_line(lbl, val):
            pdf.set_font("Arial", 'B', 11)
            pdf.cell(45, 8, lbl)
            pdf.set_font("Arial", size=11)
            pdf.cell(50, 8, f":   {val}", ln=True)

        left_line("Date", str(stats["Date"]))
        left_line("Time", stats["Start Time"])
        left_line("Duration", stats["Duration"])
        left_line("Total Member", stats["Area Total"])
        left_line("Total Register in DB", stats["Total Sys Members"])
        
        # Right Column (positioned using absolute coordinates)
        pdf.set_y(start_y)
        
        def right_cell(label, val):
            pdf.set_x(right_x)
            pdf.set_font("Arial", 'B', 11)
            pdf.cell(50, 8, label)
            pdf.set_font("Arial", size=11)
            pdf.cell(30, 8, f"=  {val}", ln=True)

        right_cell("Total Present", stats["Total Present"])
        right_cell("Member Present", stats["Members"])
        right_cell("Gospel Friend Present", stats["Gospel Friends"])
        
        # Member Rate = Area Present / Area Total
        a_p, a_t = stats["Area Present"], stats["Area Total"]
        m_rate = int(round((a_p / a_t * 100))) if a_t > 0 else 0
        right_cell(f"Member Rate (%) ({a_p}/{a_t})", f"{m_rate}%")
        
        # Overall Rate = Total Present / Area Total (As requested)
        o_p = stats["Total Present"]
        o_rate = int(round((o_p / a_t * 100))) if a_t > 0 else 0
        right_cell(f"Overall Rate (%) ({o_p}/{a_t})", f"{o_rate}%")
        
        # 3. Special Stats (TARGET ONLY)
        target = stats.get("Target Total", 0)
        if target > 0:
            pdf.set_x(right_x)
            def detail_line(lbl, val):
                pdf.set_x(right_x)
                pdf.set_font("Arial", 'B', 10)
                pdf.cell(50, 8, lbl)
                pdf.set_font("Arial", size=10); pdf.cell(30, 8, f"= {val}", ln=True)

            spec_rate = int(round((stats["Total Present"] / target * 100)))
            detail_line("Special Rate", f"{spec_rate}%")
            detail_line("Target Total", target)
        
        pdf.ln(10)

        # 4. Detailed Table
        pdf.set_font("Arial", 'B', 13)
        pdf.cell(0, 10, "Member Present Name List", ln=True)
        pdf.set_fill_color(240, 240, 240)
        pdf.set_font("Arial", 'B', 10)
        
        # Table Header
        pdf.cell(25, 9, "Code", 1, 0, 'C', True)
        pdf.cell(60, 9, "Name", 1, 0, 'C', True)
        pdf.cell(20, 9, "Age",  1, 0, 'C', True)
        pdf.cell(40, 9, "Type", 1, 0, 'C', True)
        pdf.cell(45, 9, "Area", 1, 1, 'C', True)
        
        pdf.set_font("Arial", size=9)
        for _, row in df.iterrows():
            pdf.cell(25, 8, str(row['ID'])[:12], 1)
            pdf.cell(60, 8, str(row['Name'])[:30], 1)
            pdf.cell(20, 8, str(row['Age']), 1, 0, 'C')
            pdf.cell(40, 8, str(row['Type']).capitalize(), 1)
            pdf.cell(45, 8, str(row['Area'] or 'Unknown')[:20], 1)
            pdf.ln()

        filename = out_path if out_path else f"reports/Detailed_Report_{sid}_{date.today()}.pdf"
        pdf.output(filename)
        return filename

    def generate_member_excel(self, member_ids, out_path):
        """Export member directory/list to Excel with all fields."""
        conn = sqlite3.connect(self.db_path)
        ph = ",".join("?" * len(member_ids))
        query = f"""
            SELECT member_code as 'Member ID', name as 'Name', type as 'Type', age as 'Age', 
                   dob as 'Date of Birth', area as 'Area', address as 'Address', 
                   phone as 'Phone', email as 'Email', baptism_date as 'Date of Baptism',
                   CASE WHEN has_holy_spirit = 1 THEN 'Yes' ELSE 'No' END as 'Has Holy Spirit',
                   registration_date as 'Registration Date'
            FROM members 
            WHERE member_code IN ({ph})
        """
        df = pd.read_sql(query, conn, params=member_ids)
        conn.close()
        df.to_excel(out_path, index=False)
        return out_path

    def generate_member_pdf(self, member_ids, out_path, settings):
        """Export member directory/list to PDF with profile cards."""
        conn = sqlite3.connect(self.db_path)
        ph = ",".join("?" * len(member_ids))
        query = f"SELECT member_code, name, type, age, area, address, phone, email, dob, baptism_date, has_holy_spirit, image_path FROM members WHERE member_code IN ({ph})"
        rows = conn.execute(query, member_ids).fetchall()
        conn.close()

        pdf = FPDF()
        
        for r in rows:
            m = {
                "code": r[0], "name": r[1], "type": r[2], "age": r[3], "area": r[4],
                "address": r[5], "phone": r[6], "email": r[7], "dob": r[8],
                "baptism_date": r[9], "has_holy_spirit": r[10], "image": r[11]
            }
            
            pdf.add_page()
            self._add_standard_header(pdf, settings)

            # Title: Personal Info
            pdf.set_font("Arial", 'B', 18)
            pdf.set_text_color(31, 41, 55) # #1F2937
            pdf.cell(0, 15, "Personal Info", ln=True, align='L')
            pdf.ln(5)

            # Draw a subtle separator line
            pdf.set_draw_color(229, 231, 235) # #E5E7EB
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            pdf.ln(10)

            # --- TOP SECTION: PHOTO & KEY INFO ---
            start_y = pdf.get_y()
            
            # 1. Member Photo (Left)
            photo_w, photo_h = 45, 45
            has_photo = False
            if m["image"] and os.path.exists(m["image"]):
                try:
                    pdf.image(m["image"], x=15, y=start_y, w=photo_w)
                    has_photo = True
                except: pass
            
            if not has_photo:
                pdf.set_fill_color(243, 244, 246) # #F3F4F6
                pdf.rect(15, start_y, photo_w, photo_h, 'F')
                pdf.set_xy(15, start_y + 20)
                pdf.set_font("Arial", size=8)
                pdf.cell(photo_w, 5, "NO PHOTO", 0, 0, 'C')

            # 2. Key Info & Details (Right of photo)
            info_x = 70
            pdf.set_xy(info_x, start_y)
            pdf.set_font("Arial", 'B', 16)
            pdf.set_text_color(17, 24, 39) # #111827
            pdf.cell(0, 10, m["name"].upper(), ln=True)
            
            def add_detail_row(label, value, is_bold_val=False):
                pdf.set_x(info_x)
                pdf.set_font("Arial", 'B', 10)
                pdf.set_text_color(107, 114, 128) # #6B7280
                pdf.cell(35, 7, f"{label}:", 0, 0)
                pdf.set_font("Arial", 'B' if is_bold_val else '', 10)
                pdf.set_text_color(31, 41, 55) # #1F2937
                pdf.cell(0, 7, str(value) if value else "--", 0, 1)

            add_detail_row("Member ID", m["code"], True)
            add_detail_row("Type", m["type"])
            add_detail_row("Age", m["age"])
            add_detail_row("Area", m["area"])
            add_detail_row("Date of Birth", m["dob"])
            add_detail_row("Date of Baptism", m["baptism_date"])
            add_detail_row("Has Holy Spirit", "Yes" if m["has_holy_spirit"] else "No")
            add_detail_row("Phone Number", m["phone"])
            add_detail_row("Email Address", m["email"])
            
            # Home Address (aligned with others)
            pdf.set_x(info_x)
            pdf.set_font("Arial", 'B', 10)
            pdf.set_text_color(107, 114, 128)
            pdf.cell(35, 7, "Home Address:", 0, 0)
            pdf.set_font("Arial", '', 10)
            pdf.set_text_color(31, 41, 55)
            pdf.multi_cell(0, 7, m["address"] if m["address"] else "--")

        pdf.output(out_path)
        return out_path

    def generate_org_chart_excel(self, chart_id, out_path):
        conn = sqlite3.connect(self.db_path)
        roles = pd.read_sql("""
            SELECT r.role_name AS Role, m.name AS Member, m.member_code AS 'Member ID', 
                   COALESCE(p.role_name, 'None') AS 'Parent Role'
            FROM org_chart_roles r
            LEFT JOIN members m ON r.member_code = m.member_code
            LEFT JOIN org_chart_roles p ON r.parent_role_id = p.id
            WHERE r.chart_id = ?
            ORDER BY r.id
        """, conn, params=[chart_id])
        conn.close()
        roles.to_excel(out_path, index=False)
        return out_path

    def generate_org_chart_pdf(self, chart_id, out_path, settings):
        conn = sqlite3.connect(self.db_path)
        chart_info = conn.execute("SELECT title, year FROM org_charts WHERE id=?", (chart_id,)).fetchone()
        roles = conn.execute("""
            SELECT r.id, r.parent_role_id, r.role_name, r.member_code, m.name, m.image_path
            FROM org_chart_roles r
            LEFT JOIN members m ON r.member_code = m.member_code
            WHERE r.chart_id = ?
        """, (chart_id,)).fetchall()
        conn.close()

        if not chart_info: return None
        
        # Use Landscape for wide charts
        pdf = FPDF(orientation='L', unit='mm', format='A4')
        pdf.add_page()
        self._add_standard_header(pdf, settings)
        
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, f"{chart_info[0]} - Year {chart_info[1]}", ln=True, align='C')
        pdf.ln(10)

        # Build tree structure for layout
        nodes = {}
        for rid, pid, rname, mcode, mname, img in roles:
            nodes[rid] = {"id": rid, "pid": pid, "role": rname, "name": mname or "TBA", "img": img, "children": []}
        
        root_nodes = []
        for rid, node in nodes.items():
            if node["pid"] and node["pid"] in nodes:
                nodes[node["pid"]]["children"].append(node)
            else:
                root_nodes.append(node)

        # Helper to get max depth
        def get_max_depth(nodes_list, current_depth=1):
            if not nodes_list: return current_depth
            max_d = current_depth
            for n in nodes_list:
                d = get_max_depth(n["children"], current_depth + 1)
                if d > max_d: max_d = d
            return max_d

        def get_width(node):
            if not node["children"]: return 1
            return sum(get_width(c) for c in node["children"])

        # Auto-Scaling Logic
        total_leaf_units = sum(get_width(n) for n in root_nodes) if root_nodes else 1
        max_depth = get_max_depth(root_nodes)
        
        # Base dimensions (if no scaling)
        base_box_w = 45
        base_photo_h = 40
        base_text_h = 22
        base_node_h = base_photo_h + base_text_h
        base_v_space = 30
        
        # Calculate required space vs available space
        avail_w = pdf.w - 20
        avail_h = pdf.h - pdf.get_y() - 20 # Leave some margin at bottom
        
        req_w = total_leaf_units * base_box_w
        req_h = max_depth * base_node_h + (max_depth - 1) * base_v_space
        
        scale_w = avail_w / req_w if req_w > avail_w else 1.0
        scale_h = avail_h / req_h if req_h > avail_h else 1.0
        scale = min(scale_w, scale_h, 1.0) # Only scale down, not up too much if small
        
        # Minimum allowed scale to keep text readable
        scale = max(scale, 0.3)

        # Scaled Layout Configuration
        box_w = base_box_w * scale
        photo_h = base_photo_h * scale
        text_h = base_text_h * scale
        node_h = photo_h + text_h
        v_spacing = base_v_space * scale
        start_y = pdf.get_y() + 10
        
        # Color Palette for levels
        LEVEL_COLORS = [
            (20, 184, 166),  # Teal (Chairman)
            (111, 66, 193),  # Purple
            (214, 51, 132),  # Pink
            (253, 126, 20),  # Orange
            (13, 110, 253),  # Blue
        ]

        def draw_node(node, x_start, y, available_w, level=0):
            center_x = x_start + (available_w / 2)
            node_x = center_x - (box_w / 2)
            
            color = LEVEL_COLORS[level % len(LEVEL_COLORS)]
            
            # --- Draw Lines ---
            if node["children"]:
                child_y = y + node_h + v_spacing
                total_child_w_units = sum(get_width(c) for c in node["children"])
                unit_w = available_w / total_child_w_units
                
                pdf.set_draw_color(180, 180, 180)
                pdf.set_line_width(0.3 * scale)
                # Line down from parent
                pdf.line(center_x, y + node_h, center_x, y + node_h + (v_spacing / 2))
                
                # Horizontal bridge
                first_child_x = x_start + (unit_w * get_width(node["children"][0]) / 2)
                last_child_x = x_start + available_w - (unit_w * get_width(node["children"][-1]) / 2)
                if len(node["children"]) > 1:
                    pdf.line(first_child_x, y + node_h + (v_spacing / 2), last_child_x, y + node_h + (v_spacing / 2))

            # --- Draw Card ---
            # 1. Main Background & Photo Area
            pdf.set_fill_color(245, 245, 245)
            pdf.set_draw_color(*color)
            pdf.rect(node_x, y, box_w, node_h, 'FD')
            
            # 2. Large Photo
            if node["img"] and os.path.exists(node["img"]):
                try:
                    pdf.image(node["img"], node_x, y, box_w, photo_h)
                except: 
                    pdf.set_fill_color(230, 230, 230)
                    pdf.rect(node_x, y, box_w, photo_h, 'F')
            else:
                pdf.set_fill_color(230, 230, 230)
                pdf.rect(node_x, y, box_w, photo_h, 'F')
            
            # 3. Text area (below photo)
            pdf.set_fill_color(*color)
            pdf.rect(node_x, y + photo_h, box_w, text_h, 'F')
            
            # 4. Text
            pdf.set_xy(node_x, y + photo_h + (3 * scale))
            pdf.set_font("Arial", 'B', max(6, 12 * scale)) # Scale font but keep min 6pt
            pdf.set_text_color(255, 255, 255) # White text on colored background
            pdf.cell(box_w, 7 * scale, node["role"][:25], ln=True, align='C')
            
            pdf.set_x(node_x)
            pdf.set_font("Arial", '', max(5, 10 * scale)) 
            pdf.set_text_color(255, 255, 255)
            pdf.cell(box_w, 6 * scale, node["name"][:25], ln=True, align='C')

            # --- Draw Children ---
            if node["children"]:
                current_x = x_start
                unit_w = available_w / sum(get_width(c) for c in node["children"])
                for child in node["children"]:
                    child_w = unit_w * get_width(child)
                    child_center_x = current_x + (child_w / 2)
                    # Line up to bridge
                    pdf.set_draw_color(180, 180, 180)
                    pdf.line(child_center_x, y + node_h + (v_spacing / 2), child_center_x, child_y)
                    draw_node(child, current_x, child_y, child_w, level + 1)
                    current_x += child_w

        # Start Drawing
        if root_nodes:
            unit_w = avail_w / total_leaf_units
            curr_x = 10
            for root in root_nodes:
                root_w = unit_w * get_width(root)
                draw_node(root, curr_x, start_y, root_w)
                curr_x += root_w

        pdf.output(out_path)
        return out_path
