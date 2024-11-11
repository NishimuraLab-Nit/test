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
        
        # 入室データがない場合、エラーとして処理を停止
        if not entry_time_str:
            raise ValueError(f"学生 {student_id} の入室データが見つかりません。")

        # 入室時間の解析
        entry_time = datetime.datetime.strptime(entry_time_str, "%Y-%m-%d %H:%M:%S")
        entry_day = entry_time.strftime("%A")
        entry_minutes = entry_time.hour * 60 + entry_time.minute

        # 学生の登録クラスを確認
        if student_id not in enrollment_data:
            raise ValueError(f"学生 {student_id} の登録クラスが見つかりません。")

        for student_number, class_ids in enrollment_data.items():
            if student_id not in student_number:
                raise ValueError(f"学生番号 {student_id} に対応する {student_number} が見つかりません。")

            # スプレッドシートのIDが存在するか確認
            sheet_id = item_data.get(student_number, {}).get('sheet_id')
            if not sheet_id:
                raise ValueError(f"学生 {student_number} に対応するスプレッドシートIDが見つかりません。")
            
            # スプレッドシートを取得
            sheet = client.open_by_key(sheet_id).sheet1
            date_str = entry_time.strftime("%Y-%m-%d")

            # 日付に対応する列の取得
            try:
                date_col = sheet.row_values(1).index(date_str) + 1
            except ValueError:
                raise ValueError(f"スプレッドシートに日付 {date_str} が見つかりません。")

            # クラスごとの出席記録を更新
            for i, class_id in enumerate(class_ids.get('class_id', []), start=2):
                # クラスIDと入室日が一致する授業を検索
                course = next((c for c in courses_list if c and c.get('schedule', {}).get('class_room_id') == class_id), None)
                if not course:
                    raise ValueError(f"クラスID {class_id} に対応する授業が見つかりません。")
                if course['schedule']['day'] != entry_day:
                    raise ValueError(f"学生 {student_number} の授業 {course['class_name']} は、入室日 {entry_day} には実施されません。")

                # 授業の開始時間を取得し、時間差をチェック
                start_time_str = course['schedule']['time'].split('-')[0]
                start_time = datetime.datetime.strptime(start_time_str, "%H:%M")
                start_minutes = start_time.hour * 60 + start_time.minute

                # ○の条件チェック
                if abs(entry_minutes - start_minutes) <= 5:
                    sheet.update_cell(i, date_col, "○")
                    print(f"Marked: ○ for student {student_number} in class {course['class_name']}")
                else:
                    # 条件に当てはまらない場合、例外を発生させて実行を停止
                    raise ValueError(f"学生 {student_number} は授業 {course['class_name']} の出席条件を満たしていません。実行を停止します。")

# Firebaseからデータを取得し、出席を記録
students_data = get_data_from_firebase('Students')
courses_data = get_data_from_firebase('Courses')
record_attendance(students_data, courses_data)

# Firebaseからデータを取得し、出席を記録
students_data = get_data_from_firebase('Students')
courses_data = get_data_from_firebase('Courses')
record_attendance(students_data, courses_data)
