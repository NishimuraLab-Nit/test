import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import firebase_admin
from firebase_admin import credentials, db

# Firebaseの初期化
cred = credentials.Certificate('/tmp/firebase_service_account.json')
firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'  # ここを確認
})

# Firebaseからsheet_idを取得
ref_path = 'Students/item/student_number/e19139/sheet_id'
ref = db.reference(ref_path)
sheet_id = ref.get()

# Google Sheets APIの認証設定
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('/tmp/gcp_service_account.json', scope)
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

# ヘッダー行に日付を書き込み
worksheet.update(range_name='B1', values=[dates])
print("Dates written to sheet.")

# 曜日ごとの科目を行に設定
schedule = {
    '数学': [0],  # 月曜日 (Pythonでは月曜日が0)
    '英語': [1],  # 火曜日
    '社会': [2],  # 水曜日
    '理科': [3],  # 木曜日
}

# 科目と出席情報を設定
row_start = 2
for subject, days in schedule.items():
    worksheet.update_cell(row_start, 1, subject)
    for col_num, date in enumerate(dates, 2):
        day_of_week = (start_date + timedelta(days=col_num-2)).weekday()
        if day_of_week in days:
            worksheet.update_cell(row_start, col_num, '○')
    # 各科目の割合計算式を追加
    total_days = len(dates)
    formula = f"=COUNTIF(B{row_start}:AE{row_start}, \"○\")/{total_days}*100"
    worksheet.update_cell(row_start, len(dates) + 2, formula)
    row_start += 1

print("Googleスプレッドシートにデータを書き込みました。")
