from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime, timedelta
from firebase_admin import credentials, db, initialize_app

def initialize_firebase():
    # Firebaseの初期化
    cred = credentials.Certificate("firebase-adminsdk.json")
    initialize_app(cred, {
        'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
    })

def get_google_sheets_service():
    # Google Sheets APIのサービスを取得
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file("google-credentials.json", scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)

def get_firebase_data(ref_path):
    # Firebaseからデータを取得
    return db.reference(ref_path).get()

def create_cell_update_request(sheet_id, row_index, column_index, value):
    # セルの値を更新するリクエストを作成
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": value}}]}],
            "start": {"sheetId": sheet_id, "rowIndex": row_index, "columnIndex": column_index},
            "fields": "userEnteredValue"
        }
    }

def create_dimension_request(sheet_id, dimension, start_index, end_index, pixel_size):
    # シートの次元（列や行）のプロパティを更新するリクエストを作成
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": dimension, "startIndex": start_index, "endIndex": end_index},
            "properties": {"pixelSize": pixel_size},
            "fields": "pixelSize"
        }
    }

def create_conditional_formatting_request(sheet_id, start_row, end_row, start_col, end_col, color, formula):
    # 条件付き書式のリクエストを作成
    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row, "startColumnIndex": start_col, "endColumnIndex": end_col}],
                "booleanRule": {
                    "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
                    "format": {"backgroundColor": color}
                }
            },
            "index": 0
        }
    }

def create_black_background_request(sheet_id, start_row, end_row, start_col, end_col):
    # シートの範囲外のセルを黒にするリクエストを作成
    black_color = {"red": 0.0, "green": 0.0, "blue": 0.0}
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row, "startColumnIndex": start_col, "endColumnIndex": end_col},
            "cell": {"userEnteredFormat": {"backgroundColor": black_color}},
            "fields": "userEnteredFormat.backgroundColor"
        }
    }


from googleapiclient.discovery import build
from datetime import datetime, timedelta
from firebase_admin import credentials, db, initialize_app

def initialize_firebase():
    # Firebaseの初期化
    cred = credentials.Certificate("firebase-adminsdk.json")
    initialize_app(cred, {
        'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
    })

def get_google_sheets_service():
    # Google Sheets APIのサービスを取得
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file("google-credentials.json", scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)

def get_firebase_data(ref_path):
    # Firebaseからデータを取得
    return db.reference(ref_path).get()

def create_cell_update_request(sheet_id, row_index, column_index, value):
    # セルの値を更新するリクエストを作成
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": value}}]}],
            "start": {"sheetId": sheet_id, "rowIndex": row_index, "columnIndex": column_index},
            "fields": "userEnteredValue"
        }
    }

def create_dimension_request(sheet_id, dimension, start_index, end_index, pixel_size):
    # シートの次元（列や行）のプロパティを更新するリクエストを作成
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": dimension, "startIndex": start_index, "endIndex": end_index},
            "properties": {"pixelSize": pixel_size},
            "fields": "pixelSize"
        }
    }

def create_conditional_formatting_request(sheet_id, start_row, end_row, start_col, end_col, color, formula):
    # 条件付き書式のリクエストを作成
    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row, "startColumnIndex": start_col, "endColumnIndex": end_col}],
                "booleanRule": {
                    "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
                    "format": {"backgroundColor": color}
                }
            },
            "index": 0
        }
    }

def create_black_background_request(sheet_id, start_row, end_row, start_col, end_col):
    # シートの範囲外のセルを黒にするリクエストを作成
    black_color = {"red": 0.0, "green": 0.0, "blue": 0.0}
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row, "startColumnIndex": start_col, "endColumnIndex": end_col},
            "cell": {"userEnteredFormat": {"backgroundColor": black_color}},
            "fields": "userEnteredFormat.backgroundColor"
        }
    }

def main():
  
