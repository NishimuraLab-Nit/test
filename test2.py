import datetime
import firebase_admin
from firebase_admin import credentials, db
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Initialize Firebase app
if not firebase_admin._apps:
    cred = credentials.Certificate('/tmp/firebase_service_account.json')
    firebase_admin.initialize_app(cred, {
        'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
    })

# Set up Google Sheets API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('/tmp/gcp_service_account.json', scope)
client = gspread.authorize(creds)

def get_data_from_firebase(path):
    ref = db.reference(path)
    return ref.get()

def record_attendance(students_data, courses_data):
    attendance_data = students_data.get('attendance', {}).get('students_id', {})
    enrollment_data = students_data.get('enrollment', {}).get('student_number', {})
    item_data = students_data.get('item', {}).get('student_number', {})
    courses_list = courses_data.get('course_id', [])

    for student_id, attendance in attendance_data.items():
        entry_time_str = attendance.get('entry1', {}).get('read_datetime')
        if not entry_time_str:
            continue

        entry_time = datetime.datetime.strptime(entry_time_str, "%Y-%m-%d %H:%M:%S")
        entry_day = entry_time.strftime("%A")

        for student_number, class_ids in enrollment_data.items():
            for class_id in class_ids.get('class_id', []):
                course = next((c for c in courses_list if c and c.get('schedule', {}).get('class_room_id') == class_id), None)
                if not course:
                    continue

                serial_number = attendance.get('entry1', {}).get('serial_number')
                if serial_number == courses_data.get('class_room_id', [])[1].get('serial_number'):
                    sheet_id = item_data.get(student_number, {}).get('sheet_id')
                    if sheet_id:
                        sheet = client.open_by_key(sheet_id).sheet1
                        sheet.append_row([student_number, course['class_name'], "○"])
                    continue

                if course['schedule']['day'] != entry_day:
                    continue

                start_time_str = course['schedule']['time'].split('-')[0]
                start_time = datetime.datetime.strptime(start_time_str, "%H:%M")
                entry_minutes = entry_time.hour * 60 + entry_time.minute
                start_minutes = start_time.hour * 60 + start_time.minute

                if abs(entry_minutes - start_minutes) <= 5:
                    sheet_id = item_data.get(student_number, {}).get('sheet_id')
                    if sheet_id:
                        sheet = client.open_by_key(sheet_id).sheet1
                        sheet.append_row([student_number, course['class_name'], "○"])

students_data = get_data_from_firebase('Students')
courses_data = get_data_from_firebase('Courses')

record_attendance(students_data, courses_data)
