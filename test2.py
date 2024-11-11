import datetime
import firebase_admin
from firebase_admin import credentials, db
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Firebaseアプリの初期化（未初期化の場合のみ実行）
if not firebase_admin._apps:
    cred = credentials.Certificate('/tmp/firebase_service_account.json')
    firebase_admin.initialize_app(cred, {
        'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
    })

# Google Sheets API用のスコープを設定
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('/tmp/gcp_service_account.json', scope)
client = gspread.authorize(creds)

def get_data_from_firebase(path):
    ref = db.reference(path)
    return ref.get()

# Step 3: Fetch data from Firebase
students_data = db.reference("Students").get()
courses_data = db.reference("Courses").get()

# Step 4: Helper function to match attendance with class schedule
def match_entry_with_schedule(entry_time_str, student_enrollments, courses_data):
    entry_time = datetime.strptime(entry_time_str, "%Y-%m-%d %H:%M:%S")
    for course_id in student_enrollments:
        if course_id and course_id in courses_data["course_id"]:
            course_info = courses_data["course_id"][course_id]
            schedule = course_info.get("schedule", {})
            # Check if the day and time match
            scheduled_time = schedule.get("time", "")
            if scheduled_time:
                start_time_str, _ = scheduled_time.split('-')
                scheduled_time_dt = datetime.strptime(start_time_str, "%H:%M")
                # Compare entry date with schedule's start time (date ignored)
                if (entry_time.hour == scheduled_time_dt.hour and 
                    entry_time.minute == scheduled_time_dt.minute):
                    return True
    return False

# Step 5: Process each student’s attendance and check if they match the schedule
for student_id, student_attendance in students_data["attendance"]["students_id"].items():
    entry = student_attendance.get("entry1", {})
    entry_time_str = entry.get("read_datetime", "")
    serial_number = entry.get("serial_number", "")

    if entry_time_str:
        # Retrieve student's enrollment
        student_enrollment = students_data["enrollment"]["student_number"].get(student_id, {}).get("class_id", [])
        # Check if entry time matches any class schedule
        if match_entry_with_schedule(entry_time_str, student_enrollment, courses_data):
            # Step 6: Write "〇" in Google Sheets if entry time matches the class schedule
            worksheet.update("B2", "〇")  # Update cell B2 with a circle
            break  # Stop after finding the first match
