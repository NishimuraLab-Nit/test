from firebase_admin import credentials, initialize_app, db
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime, timedelta


def initialize_firebase():
    """Firebaseを初期化してデータベースに接続します。"""
    try:
        firebase_cred = credentials.Certificate("firebase-adminsdk.json")
        initialize_app(firebase_cred, {
            'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
        })
    except Exception as e:
        print(f"Failed to initialize Firebase: {e}")


def get_google_sheets_service():
    """Google Sheets APIサービスを初期化して返します。"""
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets']
        google_creds = Credentials.from_service_account_file("google-credentials.json", scopes=scopes)
        return build('sheets', 'v4', credentials=google_creds)
    except Exception as e:
        print(f"Failed to initialize Google Sheets service: {e}")
        return None


def get_firebase_data(ref_path):
    """Firebaseから指定されたリファレンスパスのデータを取得します。"""
    try:
        return db.reference(ref_path).get()
    except Exception as e:
        print(f"Failed to retrieve data from Firebase for path '{ref_path}': {e}")
        return None


def create_cell_update_request(sheet_id, row_index, column_index, value):
    """指定セルを更新するリクエストを作成します。"""
    return {
        "updateCells": {
            "rows": [{"values": [{"userEnteredValue": {"stringValue": value}}]}],
            "start": {"sheetId": sheet_id, "rowIndex": row_index, "columnIndex": column_index},
            "fields": "userEnteredValue"
        }
    }


def create_dimension_request(sheet_id, dimension, start_index, end_index, pixel_size):
    """指定されたシートの次元（行または列）プロパティを更新するリクエストを作成します。"""
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": dimension, "startIndex": start_index, "endIndex": end_index},
            "properties": {"pixelSize": pixel_size},
            "fields": "pixelSize"
        }
    }


def create_conditional_formatting_request(sheet_id, start_row, end_row, start_col, end_col, color, formula):
    """条件付き書式を設定するリクエストを作成します。"""
    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row,
                            "startColumnIndex": start_col, "endColumnIndex": end_col}],
                "booleanRule": {
                    "condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue": formula}]},
                    "format": {"backgroundColor": color}
                }
            },
            "index": 0
        }
    }


def create_black_background_request(sheet_id, start_row, end_row, start_col, end_col):
    """黒の背景色を指定したセル範囲に設定するリクエストを作成します。"""
    black_color = {"red": 0.0, "green": 0.0, "blue": 0.0}
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row,
                      "startColumnIndex": start_col, "endColumnIndex": end_col},
            "cell": {"userEnteredFormat": {"backgroundColor": black_color}},
            "fields": "userEnteredFormat.backgroundColor"
        }
    }


def prepare_update_requests(sheet_id, class_names):
    """シートの更新リクエストリストを作成します。"""
    if not class_names:
        print("Class names list is empty. Check data retrieved from Firebase.")
        return []

    requests = [
        {"appendDimension": {"sheetId": 0, "dimension": "COLUMNS", "length": 32}},
        create_dimension_request(0, "COLUMNS", 0, 1, 100),
        create_dimension_request(0, "COLUMNS", 1, 32, 35),
        create_dimension_request(0, "ROWS", 0, 1, 120),
        {"repeatCell": {"range": {"sheetId": 0},
                        "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}},
                        "fields": "userEnteredFormat.horizontalAlignment"}},
        {"updateBorders": {"range": {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 25, "startColumnIndex": 0,
                                     "endColumnIndex": 32},
                           "top": {"style": "SOLID", "width": 1},
                           "bottom": {"style": "SOLID", "width": 1},
                           "left": {"style": "SOLID", "width": 1},
                           "right": {"style": "SOLID", "width": 1}}},
        {"setBasicFilter": {"filter": {"range": {"sheetId": 0, "startRowIndex": 0, "endRowIndex": 25,
                                                 "startColumnIndex": 0, "endColumnIndex": 32}}}}
    ]

    # 教科データを入力
    requests.append(create_cell_update_request(0, 0, 0, "教科"))
    requests.extend(create_cell_update_request(0, i + 1, 0, class_name) for i, class_name in enumerate(class_names))

    # 12月の日付と曜日を設定
    japanese_weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    start_date = datetime(2023, 12, 1)  # 12月1日を開始日に設定
    end_row = 25

    for i in range(31):  # 最大31日分を表示
        date = start_date + timedelta(days=i)
        if date.month != 12:
            break
        weekday = date.weekday()
        date_string = f"{date.strftime('%m')}\n月\n{date.strftime('%d')}\n日\n⌢\n{japanese_weekdays[weekday]}\n⌣"
        requests.append(create_cell_update_request(0, 0, i + 1, date_string))

        # 土日の条件付き書式を追加
        if weekday == 5:  # 土曜日
            requests.append(create_conditional_formatting_request(
                0, 0, end_row, i + 1, i + 2,
                {"red": 0.8, "green": 0.9, "blue": 1.0},
                f'=ISNUMBER(SEARCH("土", INDIRECT(ADDRESS(1, COLUMN()))))'
            ))
        elif weekday == 6:  # 日曜日
            requests.append(create_conditional_formatting_request(
                0, 0, end_row, i + 1, i + 2,
                {"red": 1.0, "green": 0.8, "blue": 0.8},
                f'=ISNUMBER(SEARCH("日", INDIRECT(ADDRESS(1, COLUMN()))))'
            ))

    # 黒の背景を追加
    requests.append(create_black_background_request(0, 25, 1000, 0, 1000))
    requests.append(create_black_background_request(0, 0, 1000, 32, 1000))

    return requests


def main():
    initialize_firebase()
    sheets_service = get_google_sheets_service()
    if sheets_service is None:
        print("Google Sheets service could not be initialized.")
        return

    # Firebaseからデータを取得
    sheet_id = get_firebase_data('Students/item/student_number/e19139/sheet_id')
    student_course_ids = get_firebase_data('Students/enrollment/student_number/e19139/course_id')
    courses = get_firebase_data('Courses/course_id')

    # デバッグ用に取得したデータの構造を出力
    print("Sheet ID:", sheet_id)
    print("Student Course IDs:", student_course_ids)
    print("Courses:", courses)

    # データの存在チェック
    if not sheet_id:
        print("Sheet ID not found.")
        return
    if not student_course_ids or not isinstance(student_course_ids, list):
        print("No valid course IDs found for the student or the data is not in expected list format.")
        return
    if not courses or not isinstance(courses, list):
        print("No valid courses found in the database or the data is not in expected list format.")
        return

    # coursesリストを辞書に変換（インデックスをキーにする）
    courses_dict = {i: course for i, course in enumerate(courses) if course}

    # クラス名のリストを作成
    class_names = [
        courses_dict[course_id]['class_name'] for course_id in student_course_ids
        if course_id in courses_dict and 'class_name' in courses_dict[course_id]
    ]

    # シートの更新リクエストを準備
    requests = prepare_update_requests(sheet_id, class_names)
    if not requests:
        print("No requests to update the sheet.")
        return

    # Google Sheets APIでリクエストを送信
    try:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={'requests': requests}
        ).execute()
        print("Sheet updated successfully.")
    except Exception as e:
        print(f"Failed to update the sheet: {e}")


if __name__ == "__main__":
    main()

