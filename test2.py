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

# 出席記録をGoogleスプレッドシートに更新する関数
def record_attendance(students_data, courses_data):
    # Firebaseから取得したデータを分解
    attendance_data = students_data.get('attendance', {}).get('students_id', {})
    enrollment_data = students_data.get('enrollment', {}).get('student_number', {})
    item_data = students_data.get('item', {}).get('student_number', {})
    courses_list = courses_data.get('course_id', [])

    for student_id, attendance in attendance_data.items():
        entry_time_str = attendance.get('entry1', {}).get('read_datetime')
        
        # 入室データがない場合は次の学生へ
        if not entry_time_str:
            continue

        # 入室時間の解析
        entry_time = datetime.datetime.strptime(entry_time_str, "%Y-%m-%d %H:%M:%S")
        entry_day = entry_time.strftime("%A")
        entry_minutes = entry_time.hour * 60 + entry_time.minute

        # 登録された学生か確認
        for student_number, class_ids in enrollment_data.items():
            if student_id not in student_number:
                continue

            # 対応するGoogleスプレッドシートを取得
            sheet_id = item_data.get(student_number, {}).get('sheet_id')
            if not sheet_id:
                continue
            sheet = client.open_by_key(sheet_id).sheet1
            date_str = entry_time.strftime("%Y-%m-%d")

            # 日付に対応する列の取得
            try:
                date_col = sheet.row_values(1).index(date_str) + 1
            except ValueError:
                continue  # 日付がスプレッドシートにない場合はスキップ

            # クラスごとの出席記録を更新
            for i, class_id in enumerate(class_ids.get('class_id', []), start=2):
                # クラスIDと入室日が一致する授業を検索
                course = next((c for c in courses_list if c and c.get('schedule', {}).get('class_room_id') == class_id), None)
                if not course or course['schedule']['day'] != entry_day:
                    sheet.update_cell(i, date_col, "×")
                    continue

                # 授業の開始時間を取得し、時間差をチェック
                start_time_str = course['schedule']['time'].split('-')[0]
                start_time = datetime.datetime.strptime(start_time_str, "%H:%M")
                start_minutes = start_time.hour * 60 + start_time.minute

                # 入室時間と授業開始時間が5分以内であれば「○」を記録
                if abs(entry_minutes - start_minutes) <= 5:
                    sheet.update_cell(i, date_col, "○")
                    print(f"Marked: ○ for student {student_number} in class {course['class_name']}")
                else:
                    sheet.update_cell(i, date_col, "×")

# Firebaseからデータを取得し、出席を記録
students_data = get_data_from_firebase('Students')
courses_data = get_data_from_firebase('Courses')
record_attendance(students_data, courses_data)
