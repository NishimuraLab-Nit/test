from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime, timedelta
from firebase_admin import credentials, db, initialize_app

def initialize_firebase():
    cred = credentials.Certificate("firebase-adminsdk.json")
    initialize_app(cred, {
        'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
    })

def get_google_sheets_service():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file("google-credentials.json", scopes=SCOPES)
    return build('sheets', 'v4', credentials=creds)

def get_firebase_data(ref_path):
    ref = db.reference(ref_path)
    return ref.get()

def create_cell_update_request(sheet_id, row_index, column_index, value):
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

def main():
    initialize_firebase()
    service_sheets = get_google_sheets_service()

    sheet_id = get_firebase_data('Students/item/student_number/e19139/sheet_id')
    student_class_ids = get_firebase_data('Students/enrollment/student_number/e19139/class_id')
    courses = get_firebase_data('Courses/course_id')

    class_names = [
        courses[class_id]['class_name']
        for class_id in student_class_ids if class_id in courses
    ]

    requests = [
        {"appendDimension": {"sheetId": 0, "dimension": "COLUMNS", "length": 30}},
        create_dimension_request(0, "COLUMNS", 0, 1, 70),
        create_dimension_request(0, "COLUMNS", 1, 32, 35),
        create_dimension_request(0, "ROWS", 0, 1, 120),
        {"repeatCell": {"range": {"sheetId": 0}, "cell": {"userEnteredFormat": {"horizontalAlignment": "CENTER"}}, "fields": "userEnteredFormat.horizontalAlignment"}},
        {"updateBorders": {"range": {"sheetId": 0}, "top": {"style": "SOLID", "width": 1}, "bottom": {"style": "SOLID", "width": 1}, "left": {"style": "SOLID", "width": 1}, "right": {"style": "SOLID", "width": 1}}},
        {"setBasicFilter": {"filter": {"range": {"sheetId": 0, "startRowIndex": 0, "startColumnIndex": 0, "endRowIndex": 10, "endColumnIndex": 31}}}}
    ]

    requests.append(create_cell_update_request(0, 0, 0, "教科"))

    # Insert class names starting from A2
    for i, class_name in enumerate(class_names):
        requests.append(create_cell_update_request(0, i + 1, 0, class_name))

    japanese_weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    start_date = datetime(2023, 11, 1)

    date_requests = [
        create_cell_update_request(
            0, 0, i + 1,
            f"{(start_date + timedelta(days=i)).strftime('%m')}\n月\n{(start_date + timedelta(days=i)).strftime('%d')}\n日\n⌢\n{japanese_weekdays[(start_date + timedelta(days=i)).weekday()]}\n⌣"
        )
        for i in range(30)
    ]

    requests.extend(date_requests)

    service_sheets.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={'requests': requests}
    ).execute()

if __name__ == "__main__":
    main()
