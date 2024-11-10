from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime, timedelta
from firebase_admin import credentials, db, initialize_app

# Initialize Firebase
cred = credentials.Certificate("/Users/nishimura_lab/Desktop/Noname/test-51ebc-firebase-adminsdk-t5g9u-8c04279c5c.json")
initialize_app(cred, {
    'databaseURL': 'https://test-51ebc-default-rtdb.firebaseio.com/'
})

# Set up Google Sheets API credentials
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file("/Users/nishimura_lab/Desktop/Noname/test-51ebc-56a458435c36.json", scopes=SCOPES)
service_sheets = build('sheets', 'v4', credentials=creds)

# Firebaseからsheet_idを取得
ref_path = 'Students/item/student_number/e19139/sheet_id'
ref = db.reference(ref_path)
sheet_id = ref.get()

# Retrieve class_id for the student
student_ref_path = 'Students/enrollment/student_number/e19139/class_id'
student_ref = db.reference(student_ref_path)
student_class_ids = student_ref.get()

# Retrieve all class names
course_ref_path = 'Courses/course_id'
course_ref = db.reference(course_ref_path)
courses = course_ref.get()

# Find matching class names
class_names = [courses[class_id]['class_name'] for class_id in student_class_ids if class_id in courses]

# Prepare requests to update the sheet
requests = [
    # Insert "教科" in A1
    {
        "updateCells": {
            "rows": [
                {
                    "values": [
                        {
                            "userEnteredValue": {
                                "stringValue": "教科"
                            }
                        }
                    ]
                }
            ],
            "start": {
                "sheetId": 0,
                "rowIndex": 0,
                "columnIndex": 0
            },
            "fields": "userEnteredValue"
        }
    }
]

# Insert class names starting from A2
for i, class_name in enumerate(class_names):
    requests.append(
        {
            "updateCells": {
                "rows": [
                    {
                        "values": [
                            {
                                "userEnteredValue": {
                                    "stringValue": class_name
                                }
                            }
                        ]
                    }
                ],
                "start": {
                    "sheetId": 0,
                    "rowIndex": i + 1,
                    "columnIndex": 0
                },
                "fields": "userEnteredValue"
            }
        }
    )

# 共通のupdateDimensionPropertiesリクエストを作成する関数
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

# 曜日を日本語で表示するためのリスト
japanese_weekdays = ["月", "火", "水", "木", "金", "土", "日"]

# 日付をA2から右方向に入力するためのリクエストを作成
start_date = datetime(2023, 11, 1)  # 11月1日を開始日とする

date_requests = [
    {
        "updateCells": {
            "rows": [
                {
                    "values": [
                        {
                            "userEnteredValue": {
                                "stringValue": f"{(start_date + timedelta(days=i)).strftime('%m')}\n月\n{(start_date + timedelta(days=i)).strftime('%d')}\n日\n⌢\n{japanese_weekdays[(start_date + timedelta(days=i)).weekday()]}\n⌣"
                            }
                        }
                    ]
                }
            ],
            "start": {
                "sheetId": 0,
                "rowIndex": 0,  # 1行目を指定（B1から右に展開）
                "columnIndex": i + 1  # i+1列目（0がA列なのでB列は1）
            },
            "fields": "userEnteredValue"
        }
    }
    for i in range(30)
]

# 全体のリクエスト
requests = [
    # 列数を増やす（30列以上を確保する）
    {
        "appendDimension": {
            "sheetId": 0,
            "dimension": "COLUMNS",
            "length": 30  # 必要な列数
        }
    },
    
    create_dimension_request(sheet_id=0, dimension="COLUMNS", start_index=0, end_index=1, pixel_size=70),  # A列の幅
    create_dimension_request(sheet_id=0, dimension="COLUMNS", start_index=1, end_index=32, pixel_size=35), # B列からZ列の幅
    create_dimension_request(sheet_id=0, dimension="ROWS", start_index=0, end_index=1, pixel_size=120),    # 1行目の高さ
    
    # シート全体を中央揃えにする
    {
        "repeatCell": {
            "range": {
                "sheetId": 0,
            },
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "CENTER"
                }
            },
            "fields": "userEnteredFormat.horizontalAlignment"
        }
    },
    
    # シート全体を囲むための外枠設定
    {
        "updateBorders": {
            "range": {
                "sheetId": 0,
            },
            "top": {"style": "SOLID", "width": 1},
            "bottom": {"style": "SOLID", "width": 1},
            "left": {"style": "SOLID", "width": 1},
            "right": {"style": "SOLID", "width": 1}
        }
    },
    
    # フィルター設定
    {
        "setBasicFilter": {
            "filter": {
                "range": {
                    "sheetId": 0,
                    "startRowIndex": 0,       # フィルターの開始行
                    "startColumnIndex": 0,    # フィルターの開始列
                    "endRowIndex": 10,       # フィルター範囲の終了行（例: 100行目まで）
                    "endColumnIndex": 31      # フィルター範囲の終了列（例: A-Z列）
                }
            }
        }
    }
]

# 日付入力のリクエストを追加
requests.extend(date_requests)

# リクエストをSheets APIに送信して変更を適用
service_sheets.spreadsheets().batchUpdate(
    spreadsheetId=sheet_id,
    body={'requests': requests}
).execute()
