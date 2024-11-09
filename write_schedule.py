import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import firebase_admin
from firebase_admin import credentials, db

# Firebaseの初期化
cred = credentials.Certificate('firebase-adminsdk.json')
firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://your-actual-database-name.firebaseio.com/'  # ここを確認
})

# Firebaseからsheet_idを取得
ref_path = 'Students/item/student_number/e19139/sheet_id'

ref = db.reference(ref_path)
sheet_id = ref.get()

# Google Sheets APIの認証設定
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('google-credentials.json', scope)
client = gspread.authorize(creds)

# スプレッドシートを開く
spreadsheet = client.open_by_key(sheet_id)
worksheet = spreadsheet.sheet1

# 曜日を日本語に対応
weekday_map = {
    'Sun': '日', 'Mon': '月', 'Tue': '火', 'Wed': '水',
    'Thu': '木', 'Fri': '金', 'Sat': '土'
}

# 11月の日付と曜日を列に設定
start_date = datetime(2024, 11, 1)
dates = [(start_date + timedelta(days=i)).strftime('%m月%d日') + f"({weekday_map[(start_date + timedelta(days=i)).strftime('%a')]})" for i in range(30)]

# 現在の列数を確認し、必要に応じて列を追加
max_columns = len(worksheet.row_values(1))
if max_columns < len(dates) + 2:
    worksheet.add_cols(len(dates) + 2 - max_columns)

# 日付をヘッダーに書き込む
worksheet.update(values=[dates], range_name='B1')
print("Dates written to sheet.")

# 曜日ごとの科目を行に設定
schedule = {
    '数学': [1],  # 月曜日
    '英語': [2],  # 火曜日
    '社会': [3],  # 水曜日
    '理科': [4],  # 木曜日
}

# 科目と出席情報を設定
row_start = 2
cell_updates = []

for subject, days in schedule.items():
    # 科目名を1列目に書き込む
    cell_updates.append(('A' + str(row_start), subject))
    
    # 各日付に対して出席情報を更新
    for col_num, date in enumerate(dates, 2):
        day_of_week = (start_date + timedelta(days=col_num - 2)).weekday() + 1
        if day_of_week in days:
            cell_updates.append((chr(64 + col_num) + str(row_start), '○'))
    
    # 割合計算式を追加
    formula = f"=COUNTIF(B{row_start}:AE{row_start}, \"○\")/{len(dates)}*100"
    cell_updates.append((chr(64 + len(dates) + 2) + str(row_start), formula))
    
    row_start += 1

# 一度にすべてのセルを更新
worksheet.update(cell_updates)


