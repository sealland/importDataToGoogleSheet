import os
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import numpy as np

# --- การตั้งค่า ---
EXCEL_FILE_DIRECTORY = r"D:\new-prj\pythonProject\importDataToGoogleSheet\excel"
CREDENTIALS_FILE = r"D:\new-prj\pythonProject\importDataToGoogleSheet\credentials.json" # <<--- !!! ชื่อไฟล์ OAuth 2.0 Client ID JSON !!!
TOKEN_FILE = r"D:\new-prj\pythonProject\importDataToGoogleSheet\token.json"

TARGET_SHEET_NAME = "PO"
# ลำดับของหัวคอลัมน์ (Header) อยู่ที่แถวที่ 12 (pandas ใช้ 0-indexed ดังนั้น header คือแถวที่ 11)
EXCEL_HEADER_ROW_NUMBER = 12 # แถวที่ 12 ใน Excel
EXCEL_DATA_START_ROW_NUMBER = 13 # ข้อมูลเริ่มที่แถวที่ 13 ใน Excel

PO_COLUMN_IN_SHEET = 'B' # คอลัมน์ใน Google Sheet ที่เก็บเลขที่ PO (เช่น B)
PO_COLUMN_NAME_IN_EXCEL = "เลขที่เอกสาร" # <<--- !!! ปรับชื่อคอลัมน์เลขที่ PO ใน Excel ให้ตรงกับไฟล์ของคุณ !!!
                                       # ถ้าคุณรู้ index ของคอลัมน์เลขที่ PO ใน Excel (0-indexed) ก็ใช้เลขนั้นได้
SUMMARY_ROW_KEYWORD = "รวม"
SUMMARY_ROW_COLUMN_INDEX_EXCEL = 11 # คอลัมน์ L คือ index 11 (A=0, B=1, ..., L=11)

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

def authenticate_google_sheets():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return gspread.authorize(creds)

def list_excel_files(directory):
    """แสดงรายการไฟล์ Excel ในไดเรกทอรีที่กำหนด (กรองไฟล์ temp ออก)"""
    files = [
        f for f in os.listdir(directory)
        if f.endswith('.xlsx') and not f.startswith('~$') # <--- เพิ่มเงื่อนไขนี้
    ]
    if not files:
        print(f"ไม่พบไฟล์ .xlsx ที่ถูกต้องในโฟลเดอร์: {directory}")
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

def get_last_po_number(worksheet, po_column_letter):
    """ดึงเลขที่ PO ล่าสุดจาก Google Sheet"""
    try:
        print(f"กำลังดึงข้อมูลเลขที่ PO จากคอลัมน์ {po_column_letter} ใน Google Sheet...")
        po_values = worksheet.col_values(gspread.utils.a1_to_rowcol(f"{po_column_letter}1")[1]) # เอาเฉพาะ column index
        # กรองค่าว่างและ header ออก, สมมติ header ไม่ได้อยู่ในรูปแบบ PO-xxxx
        # และเรียงลำดับเพื่อหาค่าล่าสุด
        valid_po_numbers = sorted([po for po in po_values if po and str(po).startswith("PO-")])
        if valid_po_numbers:
            last_po = valid_po_numbers[-1]
            print(f"เลขที่ PO ล่าสุดใน Google Sheet: {last_po}")
            return last_po
        else:
            print("ไม่พบเลขที่ PO ใน Google Sheet หรือยังไม่มีข้อมูล PO")
            return None
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการดึงเลขที่ PO ล่าสุด: {e}")
        return None

def main():
    print("--- โปรแกรม Import ข้อมูล Excel ไปยัง Google Sheet (ปรับปรุง) ---")

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

    # --- อ่านข้อมูลจาก Excel ---
    try:

        #df = pd.read_excel(excel_file_path,
        #                   sheet_name=0, # สมมติว่าใช้ชีทแรก
        #                   header=EXCEL_HEADER_ROW_NUMBER - 1, # pandas 0-indexed
        #                   skiprows=range(1, EXCEL_HEADER_ROW_NUMBER -1) # ข้ามแถวก่อน header แต่ไม่ข้าม header
        #                  )
        header_row_index_for_pandas = EXCEL_HEADER_ROW_NUMBER - 1
        df = pd.read_excel(excel_file_path,
                           sheet_name=0,  # สมมติว่าใช้ชีทแรก
                           header=header_row_index_for_pandas  # บอก pandas ว่าแถวไหนคือ header (0-indexed)
                           # ไม่ต้องใส่ skiprows ถ้า header parameter ถูกต้อง pandas จะเริ่มข้อมูลจากแถวถัดจาก header เอง
                           )
        print(f"อ่านข้อมูลจาก Excel สำเร็จ พบ {len(df)} แถว (ก่อนกรอง).")
        print(f"ชื่อคอลัมน์ทั้งหมดใน Excel (หลัง read_excel): {df.columns.tolist()}")  # <--- เพิ่มบรรทัดนี้
        # ตรวจสอบว่าชื่อคอลัมน์ PO_COLUMN_NAME_IN_EXCEL มีอยู่ใน DataFrame หรือไม่
        if PO_COLUMN_NAME_IN_EXCEL not in df.columns:
            print(f"!!! คำเตือน: ไม่พบคอลัมน์ '{PO_COLUMN_NAME_IN_EXCEL}' ในไฟล์ Excel !!!")
            print(f"คอลัมน์ที่มีใน Excel: {df.columns.tolist()}")
            print("โปรดตรวจสอบการตั้งค่า PO_COLUMN_NAME_IN_EXCEL ให้ถูกต้อง")
            # อาจจะให้ผู้ใช้เลือกคอลัมน์ หรือใช้ index แทน
            # return # หรือยกเลิกการทำงานไปเลย
            # ตัวอย่างการใช้ index (ถ้าคอลัมน์ B คือ index 1)
            # po_column_excel_actual = df.columns[1] # สมมติว่ารู้ว่าเป็นคอลัมน์ที่ 2 (index 1)
            # print(f"จะใช้คอลัมน์ '{po_column_excel_actual}' แทน '{PO_COLUMN_NAME_IN_EXCEL}' สำหรับเลขที่ PO จาก Excel")
            # หรือจะหยุดโปรแกรมไปเลยก็ได้ถ้าสำคัญมาก
            # return

    except FileNotFoundError:
        print(f"ไม่พบไฟล์ Excel: {excel_file_path}")
        return
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในการอ่านไฟล์ Excel: {e}")
        return

    # --- กรองแถว "รวม" ออกจาก DataFrame ---
    # ตรวจสอบว่าคอลัมน์ที่ใช้เช็ค "รวม" (คอลัมน์ L) มีอยู่จริงหรือไม่ และ index ไม่เกินขนาด df
    if SUMMARY_ROW_COLUMN_INDEX_EXCEL < len(df.columns):
        summary_column_name = df.columns[SUMMARY_ROW_COLUMN_INDEX_EXCEL]
        # กรองแถวที่คอลัมน์ L (ตาม index) ไม่มีคำว่า "รวม"
        # แปลงเป็น string เพื่อป้องกัน error ถ้ามีค่าตัวเลขหรือ NaN และใช้ .str.contains
        df_filtered = df[~df[summary_column_name].astype(str).str.contains(SUMMARY_ROW_KEYWORD, na=False)]
        print(f"กรองแถวที่มี '{SUMMARY_ROW_KEYWORD}' ในคอลัมน์ '{summary_column_name}' ออกแล้ว เหลือ {len(df_filtered)} แถว.")
        print(f"ชื่อคอลัมน์ทั้งหมดใน Excel (หลังกรอง 'รวม'): {df_filtered.columns.tolist()}")  # <--- เพิ่มบรรทัดนี้
    else:
        print(f"คำเตือน: ไม่สามารถกรองแถว '{SUMMARY_ROW_KEYWORD}' ได้ เนื่องจาก index คอลัมน์ ({SUMMARY_ROW_COLUMN_INDEX_EXCEL}) อยู่นอกช่วงของคอลัมน์ใน Excel ({len(df.columns)} คอลัมน์).")
        df_filtered = df.copy() # ใช้ DataFrame เดิมถ้ากรองไม่ได้

    # --- เชื่อมต่อ Google Sheet และดึงเลขที่ PO ล่าสุด ---
    try:
        print("กำลังเชื่อมต่อกับ Google Sheets...")
        client = authenticate_google_sheets()
        spreadsheet = client.open_by_key(google_sheet_id)
        print(f"เปิด Google Sheet '{spreadsheet.title}' สำเร็จ")

        try:
            worksheet = spreadsheet.worksheet(TARGET_SHEET_NAME)
            print(f"พบ Sheet '{TARGET_SHEET_NAME}'")
        except gspread.exceptions.WorksheetNotFound:
            print(f"ไม่พบ Sheet '{TARGET_SHEET_NAME}', กำลังสร้าง Sheet ใหม่...")
            worksheet = spreadsheet.add_worksheet(title=TARGET_SHEET_NAME, rows="1000", cols=len(df_filtered.columns) if not df_filtered.empty else 26)
            print(f"สร้าง Sheet '{TARGET_SHEET_NAME}' สำเร็จ")
            # ถ้าสร้างชีทใหม่ จะไม่มี PO ล่าสุด
            last_po_in_sheet = None
        else:
            # ดึงเลขที่ PO ล่าสุดจาก Sheet ที่มีอยู่
            last_po_in_sheet = get_last_po_number(worksheet, PO_COLUMN_IN_SHEET)

        # --- กรองข้อมูล Excel ตามเลขที่ PO ล่าสุด ---
        if last_po_in_sheet and PO_COLUMN_NAME_IN_EXCEL in df_filtered.columns:
            # กรอง DataFrame ให้มีเฉพาะ PO ที่ใหม่กว่าหรือเท่ากับ PO ล่าสุดใน Sheet
            # แปลงคอลัมน์ PO ใน Excel เป็น string เพื่อให้เปรียบเทียบได้ถูกต้อง
            df_to_upload = df_filtered[df_filtered[PO_COLUMN_NAME_IN_EXCEL].astype(str) > last_po_in_sheet].copy()
            print(f"กรองข้อมูล Excel: จะ import เฉพาะ PO ที่ใหม่กว่า '{last_po_in_sheet}'. พบ {len(df_to_upload)} รายการใหม่.")
            if df_to_upload.empty and not df_filtered.empty:
                print("ไม่พบรายการ PO ใหม่ที่จะ import จากไฟล์ Excel นี้")
                # ถ้าไม่มีข้อมูลใหม่ อาจจะถามผู้ใช้ว่าต้องการ import ทับทั้งหมดหรือไม่ หรือจบการทำงาน
                # ในที่นี้จะจบการทำงานถ้าไม่มีข้อมูลใหม่
                # return

        elif PO_COLUMN_NAME_IN_EXCEL not in df_filtered.columns:
            print(f"ไม่สามารถกรองตามเลขที่ PO ได้ เนื่องจากไม่พบคอลัมน์ '{PO_COLUMN_NAME_IN_EXCEL}' ใน Excel หลังจากกรองแถว 'รวม'")
            df_to_upload = df_filtered.copy() # Import ทั้งหมดที่กรองแถว "รวม" แล้ว
        else:
            print("ไม่พบเลขที่ PO ใน Google Sheet หรือเป็นการ Import ครั้งแรก จะ Import ข้อมูลทั้งหมด (หลังจากกรองแถว 'รวม')")
            df_to_upload = df_filtered.copy()


        if df_to_upload.empty:
            print("ไม่มีข้อมูลใหม่ที่จะอัปโหลดไปยัง Google Sheet")
            return

        import numpy as np  # ต้อง import numpy
        df_to_upload_cleaned = df_to_upload.replace([np.inf, -np.inf], np.nan).fillna('')
        print("ตรวจสอบข้อมูลตัวอย่างที่จะอัปโหลด (5 แถวแรก):")
        print(df_to_upload_cleaned.head().to_string())

        # --- เตรียมข้อมูลสำหรับอัปโหลด ---
        # ถ้าเป็นการ import ครั้งแรก หรือ sheet ว่างเปล่า และต้องการ header จาก excel
        # หรือถ้าเป็นการ import ต่อท้าย และ sheet มีข้อมูลอยู่แล้ว อาจจะไม่ต้องใส่ header อีก
        # ในที่นี้ เราจะ clear แล้วใส่ใหม่เสมอ (ตามโค้ดเดิม) หรือจะ append ก็ได้

        # ตรวจสอบว่า worksheet มีข้อมูลหรือไม่ ถ้าไม่มี ให้ใส่ header จาก excel
        # ถ้ามีข้อมูลแล้ว จะ append เฉพาะข้อมูล ไม่รวม header
        existing_data = worksheet.get_all_records(empty2zero=False, head=1) # ลอง get record แรกดูว่ามีไหม

        if not existing_data: # ถ้า sheet ว่าง หรือไม่มี header
            print(f"Sheet '{TARGET_SHEET_NAME}' ว่างเปล่า หรือไม่มี header. กำลังอัปโหลดข้อมูลพร้อม header จาก Excel...")
            worksheet.clear()
            data_to_gsheet = [df_to_upload.columns.values.tolist()] + df_to_upload.values.tolist()
            if data_to_gsheet and data_to_gsheet[0]:  # ตรวจสอบว่ามี header และข้อมูล
                worksheet.update(data_to_gsheet, 'A1')
            else:
                print("ไม่มี header หรือข้อมูลที่จะอัปโหลด")
        else:
            print(f"Sheet '{TARGET_SHEET_NAME}' มีข้อมูลอยู่แล้ว. กำลังเพิ่มข้อมูลใหม่ต่อท้าย...")
            data_to_gsheet = df_to_upload_cleaned.values.tolist()
            if data_to_gsheet:
                worksheet.append_rows(data_to_gsheet, value_input_option='USER_ENTERED',
                                      table_range='A1')  # เพิ่ม table_range
            else:
                print("ไม่มีข้อมูลใหม่ที่จะเพิ่มต่อท้าย")


        print("\n--- อัปโหลดข้อมูลสำเร็จ! ---")
        print(f"ดูผลลัพธ์ได้ที่: https://docs.google.com/spreadsheets/d/{google_sheet_id}/edit#gid={worksheet.id}")

    except FileNotFoundError as e:
        print(f"ไม่พบไฟล์: {e}")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"ไม่พบ Google Sheet ด้วย ID: {google_sheet_id}")
    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()