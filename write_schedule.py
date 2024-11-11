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

def create_gray_background_request(sheet_id, start_row, end_row, start_col, end_col):
    # シートの範囲外のセルをグレーにするリクエストを作成
    gray_color = {"red": 0.9, "green": 0.9, "blue": 0.9}
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row, "startColumnIndex": start_col, "endColumnIndex": end_col},
            "cell": {"userEnteredFormat": {"backgroundColor": gray_color}},
            "fields": "userEnteredFormat.backgroundColor"
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
    class_names = [courses[i]['class_name'] for i in student_class_ids if i and i < len(courses) and courses[i]]

    # 変更リクエストのリストを作成
    requests = [
        {"appendDimension": {"sheetId": 0, "dimension": "COLUMNS", "length": 32}},
        create_dimension_request(0, "COLUMNS", 0, 1, 100),
        create_dimension_request(0, "COLUMNS", 1, 32, 35),
        create_dimension_request(0, "ROWS", 0, 1, 120),
        {"repeatCell": {"range": {"sheetId": 0}, "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}}, "fields": "userEnteredFormat.horizontalAlignment"}},
        {"updateBorders": {"range": {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 25, "startColumnIndex": 0, "endColumnIndex": 32},
                           "top": {"style": "SOLID", "width": 1},
                           "bottom": {"style": "SOLID", "width": 1},
                           "left": {"style": "SOLID", "width": 1},
                           "right": {"style": "SOLID", "width": 1}}},
        {"setBasicFilter": {"filter": {"range": {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 25, "startColumnIndex": 0, "endColumnIndex": 32}}}}
    ]

    # A1に「教科」を入力
    requests.append(create_cell_update_request(0, 0, 0, "教科"))

    # A2から教科名を入力
    requests.extend(create_cell_update_request(0, i + 1, 0, class_name) for i, class_name in enumerate(class_names))

    # 日付を入力と条件付き書式の設定
    japanese_weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    start_date = datetime(2023, 11, 1)
    end_row = 25
    end_col = 32

    for i in range(31):
        date = start_date + timedelta(days=i)
        if date.month != 11:
            break
        weekday = date.weekday()
        date_string = f"{date.strftime('%m')}\n月\n{date.strftime('%d')}\n日\n⌢\n{japanese_weekdays[weekday]}\n⌣"
        requests.append(create_cell_update_request(0, 0, i + 1, date_string))

        # 土曜日と日曜日の列を条件付き書式で全体に適用
        if weekday == 5:
            requests.append(create_conditional_formatting_request(0, 0, end_row, i + 1, i + 2, {"red": 0.8, "green": 0.9, "blue": 1.0}, f'=ISNUMBER(SEARCH("土", INDIRECT(ADDRESS(1, COLUMN()))))'))
        elif weekday == 6:
            requests.append(create_conditional_formatting_request(0, 0, end_row, i + 1, i + 2, {"red": 1.0, "green": 0.8, "blue": 0.8}, f'=ISNUMBER(SEARCH("日", INDIRECT(ADDRESS(1, COLUMN()))))'))

    # シートの範囲外のセルをグレーにする
    requests.append(create_gray_background_request(0, 25, 1000, 0, 32))
    requests.append(create_gray_background_request(0, 0, 25, 32, 1000))

    # Google Sheets APIにバッチリクエストを送信
    service_sheets.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={'requests': requests}).execute()

if __name__ == "__main__":
    main()
