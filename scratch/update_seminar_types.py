import sqlite3
conn = sqlite3.connect('database/attendance.db')
conn.execute("UPDATE sessions SET seminar_type='Other' WHERE seminar_type='Normal'")
conn.commit()
conn.close()
print("Database updated: 'Normal' seminar types converted to 'Other'.")
