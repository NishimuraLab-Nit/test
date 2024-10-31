import firebase_admin
from firebase_admin import credentials, db
from google.oauth2 import service_account
from googleapiclient.discovery import build
import time
from googleapiclient.errors import HttpError

# Firebase アプリ初期化
if not firebase_admin._apps:
    cred = credentials.Certificate('/tmp/firebase_service_account.json')
    firebase_admin.initialize_app(cred, {
        'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
    })

# データ取得
data = db.reference().get()
print(data)

# Google Sheets APIのサービス初期化
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = '/tmp/gcp_service_account.json'

# サービスアカウントから資格情報取得
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

# Google SheetsとDriveのサービスを作成
sheets_service = build('sheets', 'v4', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)

def create_spreadsheet():
    # スプレッドシートを作成
    spreadsheet = {
        'properties': {'title': 'New Spreadsheet'}
    }
    spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId').execute()
    sheet_id = spreadsheet.get('spreadsheetId')
    print(f'Spreadsheet ID: {sheet_id}')

    # データをスプレッドシートに追加
    # Firebaseのデータ構造に合わせて値を調整
    values = []
    # ヘッダー行を追加
    headers = ['student_id', 'start', 'finish']
    values.append(headers)

    # データ行を追加
    for student_id, student_data in data['Students'].items():  # Assuming 'Students' is the key
        row = [student_id, student_data['start1'], student_data['finish1']]
        values.append(row)

    body = {'values': values}
    sheets_service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range='Sheet1!A1',
        valueInputOption='RAW',
        body=body
    ).execute()

    # Google Driveでの共有設定 with exponential backoff
    permission = {
        'type': 'user',
        'role': 'reader',
        'emailAddress': 'e19139@denki.numazu-ct.ac.jp'
    }

    retries = 0
    max_retries = 3
    wait_time = 1

    while retries < max_retries:
        try:
            drive_service.permissions().create(fileId=sheet_id, body=permission).execute()
            print(f'Spreadsheet shared with e19139@denki.numazu-ct.ac.jp with ID: {sheet_id}')
            break
        except HttpError as error:
            if error.resp.status == 403 and "sharingRateLimitExceeded" in str(error):
                print(f"Rate limit exceeded. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
                wait_time *= 2
                retries += 1
            else:
                raise
    else:
        print("Failed to share spreadsheet after multiple retries.")

# スプレッドシートを作成して共有
create_spreadsheet()

