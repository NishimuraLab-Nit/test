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
    ref = db.reference(ref_path)
    return ref.get()

def create_cell_update_request(sheet_id, row_index, column_index, value):
    # セルの値を更新するリクエストを作成
    return {
        "updateCells": {
            "rows": [
                {
                    "values": [
                        {
                            "userEnteredValue": {
                                "stringValue": value
                            }
                        }
                    ]
                }
            ],
            "start": {
                "sheetId": sheet_id,
                "rowIndex": row_index,
                "columnIndex": column_index
            },
            "fields": "userEnteredValue"
        }
    }

def create_dimension_request(sheet_id, dimension, start_index, end_index, pixel_size):
    # シートの次元（列や行）のプロパティを更新するリクエストを作成
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": dimension,
                "startIndex": start_index,
                "endIndex": end_index
            },
            "properties": {
                "pixelSize": pixel_size
            },
            "fields": "pixelSize"
        }
    }

def create_conditional_formatting_request(sheet_id, start_col, end_col, color):
    # 条件付き書式のリクエストを作成
    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [
                    {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": 1000,  # 適用する行の範囲を指定
                        "startColumnIndex": start_col,
                        "endColumnIndex": end_col
                    }
                ],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [
                            {
                                "userEnteredValue": "TRUE"  # 全てのセルに適用
                            }
                        ]
                    },
                    "format": {
                        "backgroundColor": color
                    }
                }
            },
            "index": 0
        }
    }

def main():
    # Firebaseを初期化
    initialize_firebase()
    # Google Sheetsサービスを取得
    service_sheets = get_google_sheets_service()

    # Firebaseから必要なデータを取得
    sheet_id = get_firebase_data('Students/item/student_number/e19139/sheet_id')
    student_class_ids = get_firebase_data('Students/enrollment/student_number/e19139/class_id')
    courses = get_firebase_data('Courses/course_id')

    # 履修している教科名のリストを作成
    class_names = [
        courses[i]['class_name']
        for i in student_class_ids if i and i < len(courses) and courses[i]
    ]

    # 変更リクエストのリストを作成
    requests = [
        # 列を追加
        {"appendDimension": {"sheetId": 0, "dimension": "COLUMNS", "length": 30}},
        # 列幅と行の高さを設定
        create_dimension_request(0, "COLUMNS", 0, 1, 100),
        create_dimension_request(0, "COLUMNS", 1, 32, 35),
        create_dimension_request(0, "ROWS", 0, 1, 120),
        # セルの中央揃え
        {"repeatCell": {"range": {"sheetId": 0}, "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}}, "fields": "userEnteredFormat.horizontalAlignment"}},
        # シートの外枠を設定
        {"updateBorders": {"range": {"sheetId": 0}, "top": {"style": "SOLID", "width": 1}, "bottom": {"style": "SOLID", "width": 1}, "left": {"style": "SOLID", "width": 1}, "right": {"style": "SOLID", "width": 1}}},
        # フィルターの設定
        {"setBasicFilter": {"filter": {"range": {"sheetId": 0, "startRowIndex": 0, "startColumnIndex": 0, "endRowIndex": 10, "endColumnIndex": 31}}}}
    ]

    # A1に「教科」を入力
    requests.append(create_cell_update_request(0, 0, 0, "教科"))

    # A2から教科名を入力
    for i, class_name in enumerate(class_names):
        requests.append(create_cell_update_request(0, i + 1, 0, class_name))

    # 日付を入力と条件付き書式の設定
    japanese_weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    start_date = datetime(2023, 11, 1)

    date_requests = []
    for i in range(30):
        date = start_date + timedelta(days=i)
        weekday = date.weekday()
        date_string = f"{date.strftime('%m')}\n月\n{date.strftime('%d')}\n日\n⌢\n{japanese_weekdays[weekday]}\n⌣"
        date_requests.append(create_cell_update_request(0, 0, i + 1, date_string))

        # 土曜日の列を薄い青色に
        if weekday == 5:
            requests.append(create_conditional_formatting_request(0, i + 1, i + 2, {"red": 0.8, "green": 0.9, "blue": 1.0}))

        # 日曜日の列を薄い赤色に
        if weekday == 6:
            requests.append(create_conditional_formatting_request(0, i + 1, i + 2, {"red": 1.0, "green": 0.8, "blue": 0.8}))

    # リクエストを追加
    requests.extend(date_requests)

    # Google Sheets APIにバッチリクエストを送信
    service_sheets.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={'requests': requests}
    ).execute()

if __name__ == "__main__":
    main()
