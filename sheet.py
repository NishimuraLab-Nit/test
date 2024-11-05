import firebase_admin
from firebase_admin import credentials, db
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Firebase アプリを初期化（未初期化の場合）
if not firebase_admin._apps:
    # サービスアカウントの認証情報を指定
    cred = credentials.Certificate('/tmp/firebase_service_account.json')
    # Firebase アプリを初期化
    firebase_admin.initialize_app(cred, {
        'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
    })

# Google Sheets と Drive API のスコープを定義
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
# サービスアカウントファイルのパス
SERVICE_ACCOUNT_FILE = '/tmp/gcp_service_account.json'

# サービスアカウントファイルから資格情報を取得
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

# Google Sheets と Drive のサービスクライアントを作成
sheets_service = build('sheets', 'v4', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)

def create_spreadsheet():
    try:
        # Firebase から student_number を取得
        student_id = '240f8b85'
        student_ref = db.reference(f'Students/{student_id}')
        student_data = student_ref.get()
        student_number = student_data['student_number']

        # 新しいスプレッドシートを作成
        spreadsheet = {
            'properties': {'title': student_number} # student_number をタイトルとして設定
        }
        # スプレッドシートを作成し、スプレッドシートIDを取得
        spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
        sheet_id = spreadsheet.get('spreadsheetId')
        print(f'Spreadsheet ID: {sheet_id}') # 作成されたスプレッドシートのIDを出力

        # サービスアカウント自身にも編集権限を与える
        permissions = [
            {'type': 'user', 'role': 'reader', 'emailAddress': f'{student_number}@denki.numazu-ct.ac.jp'},
            {'type': 'user', 'role': 'writer', 'emailAddress': f'{student_number}@gmail.com'}
        ]
        # 各メールアドレスに権限を設定
        for permission in permissions:
            drive_service.permissions().create(
                fileId=sheet_id,
                body=permission
            ).execute()

        # Firebase Realtime Database に sheet_id を保存
        student_ref.update({'sheet_id': sheet_id}) # sheet_id を更新

    except HttpError as error:
        # エラーが発生した場合に出力
        print(f'このようなエラーが発生しました {error}')

create_spreadsheet()
