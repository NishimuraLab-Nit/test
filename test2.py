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

# Firebaseからデータを取得する関数
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
            raise ValueError(f"学生 {student_id} の入室データが見つかりません。")

        entry_time = datetime.datetime.strptime(entry_time_str, "%Y-%m-%d %H:%M:%S")
        entry_day = entry_time.strftime("%A")
        entry_minutes = entry_time.hour * 60 + entry_time.minute

        student_info = students_data.get('student_info', {}).get('student_id', {}).get(student_id)
        if not student_info:
            raise ValueError(f"学生 {student_id} の情報が見つかりません。")
        
        student_number = student_info.get('student_number')
        if student_number not in enrollment_data:
            raise ValueError(f"学生番号 {student_number} の登録クラスが見つかりません。")

        class_ids = enrollment_data[student_number]['class_id']

        sheet_id = item_data.get(student_number, {}).get('sheet_id')
        if not sheet_id:
            raise ValueError(f"学生番号 {student_number} に対応するスプレッドシートIDが見つかりません。")
        
        sheet = client.open_by_key(sheet_id).sheet1

        for i, class_id in enumerate(class_ids, start=2):
            course = next((c for c in courses_list if c and c.get('schedule', {}).get('class_room_id') == class_id), None)
            if not course:
                raise ValueError(f"クラスID {class_id} に対応する授業が見つかりません。")
            if course['schedule']['day'] != entry_day:
                raise ValueError(f"学生 {student_number} の授業 {course['class_name']} は、入室日 {entry_day} には実施されません。")

            start_time_str = course['schedule']['time'].split('-')[0]
            start_time = datetime.datetime.strptime(start_time_str, "%H:%M")
            start_minutes = start_time.hour * 60 + start_time.minute

            if abs(entry_minutes - start_minutes) <= 5:
                # Update attendance without checking date format
                sheet.update_cell(i, 2, "○")  # Assuming column 2 is for attendance
                print(f"Marked: ○ for student {student_number} in class {course['class_name']}")
            else:
                raise ValueError(f"学生 {student_number} は授業 {course['class_name']} の出席条件を満たしていません。実行を停止します。")

# Firebaseからデータを取得し、出席を記録
students_data = get_data_from_firebase('Students')
courses_data = get_data_from_firebase('Courses')
record_attendance(students_data, courses_data)
