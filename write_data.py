import firebase_admin
from firebase_admin import credentials, db
from google.oauth2 import service_account
from googleapiclient.discovery import build

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

def update_existing_spreadsheet():
    # FirebaseのデータベースからStudentsノードを参照
    students_ref = db.reference('Students')
    # Studentsノードの全データを取得
    students_data = students_ref.get()

    # 各student_idごとにデータを処理
    for student_id, student_info in students_data.items():
        sheet_id = student_info.get('sheet_id')  # 各生徒のシートIDを取得

        if sheet_id:
            # スプレッドシートに追加するヘッダー
            headers = ['Field', 'Value']
            values = [headers]

            # 各フィールドとその値をスプレッドシートに追加
            for key, value in student_info.items():
                if key != 'sheet_id':  # sheet_idはスキップ
                    values.append([key, value])

            # スプレッドシートに書き込むデータを作成
            body = {'values': values}
            # Sheets APIを使ってスプレッドシートを更新
            sheets_service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range='Sheet1!A1',  # 書き込み開始位置
                valueInputOption='RAW',  # 値をそのまま書き込む
                body=body
            ).execute()
            print(f'Spreadsheet {sheet_id} updated for student {student_id}.')  # 更新完了メッセージ
        else:
            print(f"Student {student_id} does not have a sheet_id.")  # シートIDがない場合のメッセージ

# スプレッドシートの更新を実行
update_existing_spreadsheet()

