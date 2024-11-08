import datetime
import json
import firebase_admin
from firebase_admin import credentials, db
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# Firebase アプリを初期化（未初期化の場合）
if not firebase_admin._apps:
    # Firebaseのサービスアカウントキーを使って認証
    cred = credentials.Certificate('/tmp/firebase_service_account.json')
    # Realtime DatabaseのURLを指定してアプリを初期化
    firebase_admin.initialize_app(cred, {
        'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
    })

# Google APIの認証設定
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']  # スプレッドシートAPIへのアクセス権
SERVICE_ACCOUNT_FILE = '/tmp/gcp_service_account.json'  # サービスアカウントキーのファイルパス

# サービスアカウントキーを使ってGoogle APIに認証
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
# Sheets APIのサービスオブジェクトを作成
sheets_service = build('sheets', 'v4', credentials=creds)

# Firebaseからデータを取得する関数
def get_data_from_firebase(path):
    ref = db.reference(path)
    return ref.get()

# 出席を記録する関数
def record_attendance(students_data, courses_data):
    # 各種データを取得
    attendance_data = students_data.get('attendance', {}).get('students_id', {})
    enrollment_data = students_data.get('enrollment', {}).get('student_number', {})
    item_data = students_data.get('item', {}).get('student_number', {})

    # 出席データを順に確認
    for student_id, attendance in attendance_data.items():
        entry_time_str = attendance.get('entry1', {}).get('read_datetime')
        if not entry_time_str:
            continue  # エントリー時間がない場合はスキップ

        # 入室時間を日時形式に変換
        entry_time = datetime.datetime.strptime(entry_time_str, "%Y-%m-%d %H:%M:%S")
        entry_day = entry_time.strftime("%A")  # 曜日を取得

        # 学生の履修データを確認
        for student_number, class_ids in enrollment_data.items():
            for class_id in class_ids:
                course = courses_data.get(class_id)
                if not course or course['schedule']['day'] != entry_day:
                    continue  # コースが存在しないか、エントリーの日とコースの日が一致しない場合はスキップ

                # 授業開始時間を取得
                start_time_str = course['schedule']['time'].split('-')[0]
                start_time = datetime.datetime.strptime(start_time_str, "%H:%M")
                entry_minutes = entry_time.hour * 60 + entry_time.minute
                start_minutes = start_time.hour * 60 + start_time.minute

                # 入室時間が授業開始時間の5分前後かをチェック
                if abs(entry_minutes - start_minutes) <= 5:
                    # スプレッドシートIDを取得
                    sheet_id = item_data.get(student_number, {}).get('sheet_id')
                    if sheet_id:
                        # スプレッドシートに接続
                        sheet = client.open_by_key(sheet_id).sheet1
                        # スプレッドシートに出席情報を書き込み
                        sheet.append_row([student_number, course['class_name'], "○"])

# Firebaseから必要なデータを取得
students_data = get_data_from_firebase('Students')
courses_data = get_data_from_firebase('Courses/course_id')

# 出席を記録
record_attendance(students_data, courses_data)
