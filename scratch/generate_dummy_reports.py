import sqlite3
import random
from datetime import datetime, timedelta

def generate_dummy_data():
    conn = sqlite3.connect('database/attendance.db')
    cursor = conn.cursor()
    
    # 1. Get some members to work with
    members = cursor.execute("SELECT member_code, name, type, title FROM members").fetchall()
    if not members:
        print("No members found. Please add some members first.")
        return
    
    print(f"Found {len(members)} members. Generating dummy seminar data...")
    
    # 2. Clear previous seminar sessions to avoid mess
    cursor.execute("DELETE FROM sessions WHERE seminar_type != 'Normal'")
    
    # 3. Generate sessions for the past year
    # Every Friday and Saturday
    start_date = datetime.now() - timedelta(days=365)
    
    session_count = 0
    attendance_count = 0
    
    current_date = start_date
    while current_date <= datetime.now():
        day_name = current_date.strftime('%A')
        if day_name in ['Friday', 'Saturday']:
            seminar_type = f"{day_name} Seminar"
            title = f"DUMMY {day_name.upper()} SEMINAR {current_date.strftime('%d-%m-%Y')}"
            
            # Create session
            cursor.execute("INSERT INTO sessions (title, date, start_time, seminar_type) VALUES (?, ?, ?, ?)",
                           (title, current_date.date(), current_date, seminar_type))
            session_id = cursor.lastrowid
            session_count += 1
            
            # Create random attendance
            # Select 40-80% of members randomly
            present_count = random.randint(int(len(members)*0.4), int(len(members)*0.8))
            present_members = random.sample(members, present_count)
            
            for m_code, m_name, m_type, m_title in present_members:
                cursor.execute("""
                    INSERT INTO attendance (person_name, member_code, session_id, check_in_time, service_date, status)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (m_name, m_code, session_id, current_date, current_date.date(), m_type.lower()))
                attendance_count += 1
                
        current_date += timedelta(days=1)
    
    conn.commit()
    conn.close()
    print(f"Finished! Created {session_count} sessions and {attendance_count} attendance records.")

if __name__ == "__main__":
    generate_dummy_data()
