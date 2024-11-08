import datetime
import firebase_admin
from firebase_admin import credentials, db
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Firebaseアプリの初期化（未初期化の場合のみ実行）
if not firebase_admin._apps:
    # サービスアカウントキーのJSONファイルを指定
    cred = credentials.Certificate('/tmp/firebase_service_account.json')
    # Firebaseアプリを初期化し、Realtime DatabaseのURLを指定
    firebase_admin.initialize_app(cred, {
        'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
    })

# Google Sheets API用のスコープを設定
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
# サービスアカウントキーのJSONファイルを使って資格情報を取得
creds = ServiceAccountCredentials.from_json_keyfile_name('/tmp/gcp_service_account.json', scope)
# Google Sheetsクライアントを認証
client = gspread.authorize(creds)

# Firebaseからデータを取得する関数
def get_data_from_firebase(path):
    # 指定されたパスからデータベースの参照を取得
    ref = db.reference(path)
    # データを取得して返す
    return ref.get()

# 出席データを記録する関数
def record_attendance(students_data, courses_data):
    # 学生の出席データを取得
    attendance_data = students_data.get('attendance', {}).get('students_id', {})
    # 学生の登録情報を取得
    enrollment_data = students_data.get('enrollment', {}).get('student_number', {})
    # シートIDを含む学生のアイテムデータを取得
    item_data = students_data.get('item', {}).get('student_number', {})
    # コースデータのリストを取得
    courses_list = courses_data.get('course_id', [])

    # 各学生の出席情報をループ
    for student_id, attendance in attendance_data.items():
        # 入室時刻の文字列を取得
        entry_time_str = attendance.get('entry1', {}).get('read_datetime')
        if not entry_time_str:
            # 入室時刻がない場合は次へ
            continue

        # 入室時刻をdatetimeオブジェクトに変換
        entry_time = datetime.datetime.strptime(entry_time_str, "%Y-%m-%d %H:%M:%S")
        # 入室日を曜日形式の文字列に変換
        entry_day = entry_time.strftime("%A")

        # 学生番号と登録されたクラスIDをループ
        for student_number, class_ids in enrollment_data.items():
            for class_id in class_ids.get('class_id', []):
                # クラスIDに一致するコースを取得
                course = next((c for c in courses_list if c and c.get('schedule', {}).get('class_room_id') == class_id), None)
                if not course:
                    # コースが見つからない場合は次へ
                    continue

                # シリアル番号を取得して一致を確認
                serial_number = attendance.get('entry1', {}).get('serial_number')
                if serial_number == courses_data.get('class_room_id', [])[1].get('serial_number'):
                    # シートIDを取得
                    sheet_id = item_data.get(student_number, {}).get('sheet_id')
                    if sheet_id:
                        # Google Sheetsに接続し、行を追加
                        sheet = client.open_by_key(sheet_id).sheet1
                        sheet.append_row([student_number, course['class_name'], "○"])
                    continue

                # コースの曜日が入室日と一致するか確認
                if course['schedule']['day'] != entry_day:
                    continue

                # コースの開始時刻を取得し、datetimeオブジェクトに変換
                start_time_str = course['schedule']['time'].split('-')[0]
                start_time = datetime.datetime.strptime(start_time_str, "%H:%M")
                # 入室時間とコース開始時間を分に変換
                entry_minutes = entry_time.hour * 60 + entry_time.minute
                start_minutes = start_time.hour * 60 + start_time.minute

                # 入室時間が開始時間の5分以内か確認
                if abs(entry_minutes - start_minutes) <= 5:
                    # シートIDを取得
                    sheet_id = item_data.get(student_number, {}).get('sheet_id')
                    if sheet_id:
                        # Google Sheetsに接続し、行を追加
                        sheet = client.open_by_key(sheet_id).sheet1
                        sheet.append_row([student_number, course['class_name'], "○"])

# Firebaseから学生とコースのデータを取得
students_data = get_data_from_firebase('Students')
courses_data = get_data_from_firebase('Courses')

# 出席データを記録
record_attendance(students_data, courses_data)
