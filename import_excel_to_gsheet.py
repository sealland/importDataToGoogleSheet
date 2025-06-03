import os
import pandas as pd
import gspread
# from google.oauth2.credentials import Credentials # ไม่จำเป็นต้อง import โดยตรงแล้วถ้าใช้ gspread.oauth
# from google_auth_oauthlib.flow import InstalledAppFlow # gspread.oauth จะจัดการส่วนนี้
# from google.auth.transport.requests import Request # gspread.oauth จะจัดการส่วนนี้

# --- การตั้งค่า ---
EXCEL_FILE_DIRECTORY = r"D:\new-prj\pythonProject\importDataToGoogleSheet\excel"
# ไฟล์ JSON ที่ดาวน์โหลดจาก OAuth 2.0 Client ID (client_secret.json)
OAUTH_CREDENTIALS_FILE = r"D:\new-prj\pythonProject\importDataToGoogleSheet\credentials.json" # <<--- ตรวจสอบ Path และชื่อไฟล์
# ไฟล์สำหรับเก็บ token ที่ได้จากการ authorize (gspread จะสร้างและจัดการให้)
AUTHORIZED_USER_FILE = r"D:\new-prj\pythonProject\importDataToGoogleSheet\authorized_user.json" # <<--- เปลี่ยนชื่อไฟล์ token

TARGET_SHEET_NAME = "PO"
EXCEL_START_ROW = 13

# SCOPES ไม่จำเป็นต้องกำหนดแยกแล้ว เพราะ gspread.oauth จะใช้ default scopes ที่ครอบคลุม sheets และ drive
# SCOPES = [
#     'https://www.googleapis.com/auth/spreadsheets',
#     'https://www.googleapis.com/auth/drive.file'
# ]

def authenticate_google_sheets_gspread_oauth():
    """
    จัดการการยืนยันตัวตนกับ Google Sheets API โดยใช้ gspread.oauth().
    จะเปิดเบราว์เซอร์ให้ผู้ใช้ล็อกอินในครั้งแรก หรือถ้า token หมดอายุ/ไม่มี
    """
    try:
        # gspread.oauth จะจัดการการโหลด token จาก AUTHORIZED_USER_FILE
        # และจะเริ่ม OAuth flow ถ้าจำเป็น (เช่น ครั้งแรก หรือ token หมดอายุ)
        gc = gspread.oauth(
            credentials_filename=OAUTH_CREDENTIALS_FILE,
            authorized_user_filename=AUTHORIZED_USER_FILE
        )
        return gc
    except Exception as e:
        print(f"เกิดข้อผิดพลาดระหว่างการยืนยันตัวตน gspread.oauth: {e}")
        print("กรุณาตรวจสอบว่า:")
        print(f"1. ไฟล์ '{OAUTH_CREDENTIALS_FILE}' (client_secret.json) อยู่ในตำแหน่งที่ถูกต้อง")
        print("2. คุณได้ให้สิทธิ์ที่จำเป็นเมื่อเบราว์เซอร์เปิดขึ้นมา")
        raise # ส่งต่อ exception เพื่อให้โปรแกรมหยุดทำงาน

def list_excel_files(directory):
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    if not files:
        print(f"ไม่พบไฟล์ .xlsx ในโฟลเดอร์: {directory}")
        return None
    print("ไฟล์ Excel ที่พบ:")
    for i, f_name in enumerate(files):
        print(f"{i + 1}. {f_name}")
    return files

def select_excel_file(files):
    while True:
        try:
            choice = int(input("เลือกหมายเลขไฟล์ Excel ที่ต้องการ import: "))
            if 1 <= choice <= len(files):
                return files[choice - 1]
            else:
                print("หมายเลขไม่ถูกต้อง กรุณาเลือกใหม่")
        except ValueError:
            print("กรุณาป้อนเป็นตัวเลข")

def main():
    print("--- โปรแกรม Import ข้อมูล Excel ไปยัง Google Sheet (gspread.oauth) ---")

    excel_files = list_excel_files(EXCEL_FILE_DIRECTORY)
    if not excel_files:
        return

    selected_excel_file_name = select_excel_file(excel_files)
    excel_file_path = os.path.join(EXCEL_FILE_DIRECTORY, selected_excel_file_name)
    print(f"\nกำลังประมวลผลไฟล์: {excel_file_path}")

    google_sheet_id = input("กรุณาใส่ ID ของ Google Sheet: ")
    if not google_sheet_id:
        print("ไม่ได้ใส่ Google Sheet ID, ยกเลิกการทำงาน")
        return

    try:
        df = pd.read_excel(excel_file_path, sheet_name=0, skiprows=EXCEL_START_ROW - 1)
        print(f"อ่านข้อมูลจาก Excel สำเร็จ พบ {len(df)} แถว.")
    except FileNotFoundError:
        print(f"ไม่พบไฟล์ Excel: {excel_file_path}")
        return
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการอ่านไฟล์ Excel: {e}")
        return

    df = df.fillna('')
    data_to_upload = [df.columns.values.tolist()] + df.values.tolist()

    try:
        print("กำลังเชื่อมต่อกับ Google Sheets...")
        # ใช้ฟังก์ชัน authenticate แบบใหม่
        gc = authenticate_google_sheets_gspread_oauth() # gc คือ gspread.Client instance

        # ใช้ client.open_by_key() หรือ client.open_by_id() ก็ได้ใน gspread 6.x
        # หรือ client.open_by_url()
        spreadsheet = gc.open_by_key(google_sheet_id) # หรือ gc.open_by_id(google_sheet_id)
        # spreadsheet = gc.open_by_id(google_sheet_id) # ลองอันนี้ถ้า open_by_key ไม่ได้
        print(f"เปิด Google Sheet '{spreadsheet.title}' สำเร็จ")

        try:
            worksheet = spreadsheet.worksheet(TARGET_SHEET_NAME)
            print(f"พบ Sheet '{TARGET_SHEET_NAME}'")
        except gspread.exceptions.WorksheetNotFound:
            print(f"ไม่พบ Sheet '{TARGET_SHEET_NAME}', กำลังสร้าง Sheet ใหม่...")
            worksheet = spreadsheet.add_worksheet(title=TARGET_SHEET_NAME, rows="1000", cols="26")
            print(f"สร้าง Sheet '{TARGET_SHEET_NAME}' สำเร็จ")

        print(f"กำลังล้างข้อมูลเก่าใน Sheet '{TARGET_SHEET_NAME}'...")
        worksheet.clear()

        print(f"กำลังอัปโหลดข้อมูลไปยัง Sheet '{TARGET_SHEET_NAME}'...")
        worksheet.update(data_to_upload, 'A1')

        print("\n--- อัปโหลดข้อมูลสำเร็จ! ---")
        print(f"ดูผลลัพธ์ได้ที่: https://docs.google.com/spreadsheets/d/{google_sheet_id}/edit#gid={worksheet.id}")

    except FileNotFoundError as e:
        # ตรวจสอบว่า FileNotFoundError มาจาก credentials หรือไม่
        if OAUTH_CREDENTIALS_FILE in str(e):
            print(f"ไม่พบไฟล์ Credentials (OAuth 2.0): {OAUTH_CREDENTIALS_FILE}")
            print("กรุณาตรวจสอบว่าไฟล์ JSON (client_secret) อยู่ในตำแหน่งที่ถูกต้องและชื่อไฟล์ถูกต้อง")
        else:
            print(f"เกิด FileNotFoundError อื่น: {e}")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ไม่พบ Google Sheet ด้วย ID: {google_sheet_id}")
        print("กรุณาตรวจสอบ Google Sheet ID หรือ URL และสิทธิ์การเข้าถึง")
    except gspread.exceptions.APIError as e:
        print(f"เกิดข้อผิดพลาดจาก Google Sheets API: {e}")
    except Exception as e:
        print(f"เกิดข้อผิดพลาดที่ไม่คาดคิด: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()