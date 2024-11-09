import datetime
from datetime import timedelta
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
def get_sheet_id(path):
    ref = db.reference(path)
    return ref.get()

# 出席データを記録する関数
def record_attendance(sheet_id):
    spreadsheet = client.open_by_key(sheet_id)
    worksheet = spreadsheet.sheet1

    weekday_map = {
        'Sun': '日', 'Mon': '月', 'Tue': '火', 'Wed': '水',
        'Thu': '木', 'Fri': '金', 'Sat': '土'
    }

    start_date = datetime.datetime(2024, 11, 1)
    dates = [(start_date + timedelta(days=i)).strftime('%m月%d日') + f"({weekday_map[(start_date + timedelta(days=i)).strftime('%a')]})" for i in range(30)]

    worksheet.update('B1', [dates])
    print("Dates written to sheet.")

    schedule = {
        '数学': [0],
        '英語': [1],
        '社会': [2],
        '理科': [3],
    }

    row_start = 2
    for subject, days in schedule.items():
        worksheet.update_cell(row_start, 1, subject)
        for col_num, date in enumerate(dates, 2):
            day_of_week = (start_date + timedelta(days=col_num-2)).weekday()
            if day_of_week in days:
                worksheet.update_cell(row_start, col_num, '○')
        total_days = len(dates)
        formula = f"=COUNTIF(B{row_start}:AE{row_start}, \"○\")/{total_days}*100"
        worksheet.update_cell(row_start, len(dates) + 2, formula)
        row_start += 1

    print("Googleスプレッドシートにデータを書き込みました。")

# 実行例
path_to_sheet_id = 'Students/item/student_number/e19139/sheet_id'
sheet_id = get_sheet_id(path_to_sheet_id)
record_attendance(sheet_id)
