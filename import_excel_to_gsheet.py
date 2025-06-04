import os
import pandas as pd
import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import numpy as np
import pickle # สำหรับบันทึก/โหลดค่า
import customtkinter as ctk # GUI Library
from tkinter import filedialog, messagebox
import threading # สำหรับรัน process import ใน background
import time # สำหรับสาธิต progress bar (ถ้าต้องการ)

# --- ชื่อไฟล์สำหรับบันทึกการตั้งค่าทั้งหมดของแอป ---
APP_SETTINGS_FILE = "importer_app_settings.v3.pkl" # ตั้งชื่อใหม่เผื่อมีเวอร์ชันเก่า

# --- การตั้งค่าเริ่มต้น (อาจจะถูก override ด้วยค่าที่จำไว้) ---
DEFAULT_EXCEL_DIR = os.path.expanduser("~") # เริ่มที่ Home directory ของ User
DEFAULT_GOOGLE_SHEET_ID = ""
SETTINGS_FILE = "import_settings.pkl" # ไฟล์สำหรับบันทึกค่า


# --- การตั้งค่า ---
EXCEL_FILE_DIRECTORY = r"D:\new-prj\pythonProject\importDataToGoogleSheet\excel"
CREDENTIALS_FILE = r"D:\new-prj\pythonProject\importDataToGoogleSheet\credentials.json" # <<--- !!! ชื่อไฟล์ OAuth 2.0 Client ID JSON !!!
TOKEN_FILE = r"D:\new-prj\pythonProject\importDataToGoogleSheet\token.json"

DEFAULT_DOCUMENT_CONFIGS = {
    "PO_DETAIL": {
        "display_name": "รายละเอียดใบสั่งซื้อ (PO - Detail)",
        "document_type_code": "PO",
        "target_sheet_name": "PO",
        "header_row_excel": 12, # <--- แก้ไขตรงนี้ จาก 15 เป็น 12
        "parent_id_column_name_excel": "เลขที่เอกสาร", # คอลัมน์ใน Excel Detail ที่อ้างอิง PO หลัก
        # ไม่มี id_column_name_excel และ id_column_letter_gsheet สำหรับเช็ค ID detail โดยตรงในขั้นนี้
        "summary_keyword_excel": "รวม",
        "summary_column_index_excel": 26, # ตรวจสอบว่า index นี้ถูกต้องสำหรับไฟล์ PO Detail ของคุณ
        "date_columns_in_excel": ["วันที่อนุมัติ", "วันที่สร้าง", "วันที่เอกสาร"], # ตรวจสอบว่า PO Detail มีคอลัมน์วันที่เหล่านี้หรือไม่ ถ้าไม่มีให้เป็น []
        "last_used_excel_dir": os.path.expanduser("~"),
        "last_used_gsheet_id_input": ""
    },
    "QO_DETAIL": {
        "display_name": "รายละเอียดใบเสนอราคา (QO - Detail)",
        "document_type_code": "QO",
        "target_sheet_name": "QO",
        "header_row_excel": 12, # <<--- ตรวจสอบแถว Header ของ QO Detail Excel อีกครั้งว่าถูกต้องหรือไม่
        "parent_id_column_name_excel": "เลขที่เอกสาร", # <<--- !!! แก้ไขเป็นชื่อนี้ !!!
        "line_item_id_column_excel": "รายการที่", # <<--- ตรวจสอบว่าชื่อคอลัมน์ "รายการที่" ถูกต้องสำหรับ QO Detail
        "summary_keyword_excel": "รวม",
        "summary_column_index_excel": 25, # ตรวจสอบว่า index นี้ถูกต้องสำหรับไฟล์ PO Detail ของคุณ
        "date_columns_in_excel": ["วันที่อนุมัติ", "วันที่สร้าง", "วันที่เอกสาร"], # ตรวจสอบว่า PO Detail มีคอลัมน์วันที่เหล่านี้หรือไม่ ถ้าไม่มีให้เป็น []
        "last_used_excel_dir": os.path.expanduser("~"),
        "last_used_gsheet_id_input": ""
    },
    # SO_HEADER, SO_DETAIL, DO_HEADER, DO_DETAIL สามารถเพิ่มตามรูปแบบนี้ได้
}
current_app_configs = {} # จะถูก init โดย load_app_settings

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

current_app_configs = {}

app = None
log_textbox = None
progressbar = None # Global reference for progress bar
log_frame_visible = True # สถานะการมองเห็นของ Log Frame

def authenticate_google_sheets():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                # อาจจะต้องลบ token.json แล้วให้ user auth ใหม่
                log_message_ui(f"Error refreshing token: {e}. Please re-authenticate.")
                if os.path.exists(TOKEN_FILE):
                    os.remove(TOKEN_FILE)
                # เรียก flow ใหม่
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json",
                                                                 SCOPES)  # สมมติมี client_secret.json
                creds = flow.run_local_server(port=0)
        else:
            try:
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json",
                                                                 SCOPES)  # สมมติมี client_secret.json
                creds = flow.run_local_server(port=0)
            except FileNotFoundError:
                log_message_ui("Error: 'credentials.json' (OAuth client secret) not found.")
                messagebox.showerror("Authentication Error",
                                     "File 'credentials.json' not found. Please set up OAuth 2.0 client.")
                return None
            except Exception as e:
                log_message_ui(f"Authentication error: {e}")
                return None

        if creds:  # บันทึกเฉพาะเมื่อ creds ถูกสร้างหรือ refresh สำเร็จ
            with open(TOKEN_FILE, 'w') as token:
                token.write(creds.to_json())
    return gspread.authorize(creds) if creds else None

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

def save_app_settings(configs_to_save):
    global current_app_configs
    try:
        with open(APP_SETTINGS_FILE, "wb") as f:
            pickle.dump(configs_to_save, f)
        current_app_configs = configs_to_save
        if app and hasattr(app, 'log_message_ui'): # ตรวจสอบว่า UI พร้อม
             app.log_message_ui("บันทึกการตั้งค่าแอปพลิเคชันแล้ว")
        else:
            print(f"บันทึกการตั้งค่าแอปพลิเคชันไปยัง {APP_SETTINGS_FILE}")
    except Exception as e:
        if app and hasattr(app, 'log_message_ui'):
            app.log_message_ui(f"Error บันทึกการตั้งค่าแอปพลิเคชัน: {e}")
        else:
            print(f"Error บันทึกการตั้งค่าแอปพลิเคชัน: {e}")


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

def process_excel_and_gsheet(excel_file_path, google_sheet_id_main, doc_type_key):
    global progressbar, current_app_configs, app
    config = current_app_configs.get(doc_type_key)

    if not config:
        log_message_ui(f"Error: ไม่พบการตั้งค่าสำหรับประเภทเอกสาร '{doc_type_key}'")
        messagebox.showerror("Config Error", f"ไม่พบการตั้งค่าสำหรับประเภทเอกสาร '{doc_type_key}'")
        if progressbar: progressbar.stop(); progressbar.set(0)
        return

    doc_display_name = config.get('display_name', doc_type_key)
    target_sheet_name = config.get('target_sheet_name')
    excel_header_row = config.get('header_row_excel')
    summary_keyword = config.get('summary_keyword_excel', None)
    summary_column_idx = config.get('summary_column_index_excel', -1)
    date_columns_to_format = config.get('date_columns_in_excel', [])

    is_header_type = "HEADER" in doc_type_key.upper()

    id_column_excel_header = None
    id_column_gsheet_letter_header = None
    id_prefix_for_gsheet_check = ""

    if is_header_type:
        id_column_excel_header = config.get('id_column_name_excel')
        id_column_gsheet_letter_header = config.get('id_column_letter_gsheet')
        id_prefix_for_gsheet_check = config.get('document_type_code', "") + "-"

    if not target_sheet_name or not excel_header_row:
        log_message_ui(
            f"Error: การตั้งค่าสำหรับ '{doc_display_name}' ไม่สมบูรณ์ (target_sheet_name หรือ header_row_excel ขาดหายไป)")
        messagebox.showerror("Config Error", f"การตั้งค่าสำหรับ '{doc_display_name}' ไม่สมบูรณ์")
        if progressbar: progressbar.stop(); progressbar.set(0)
        return

    spreadsheet = None
    worksheet = None
    client = None
    df_to_upload = pd.DataFrame() # Initialize df_to_upload as an empty DataFrame

    try:
        # 1. อ่าน Excel
        log_message_ui(f"1. กำลังอ่านไฟล์ Excel '{os.path.basename(excel_file_path)}' สำหรับ '{doc_display_name}'...")
        header_row_idx_pandas = excel_header_row - 1
        df = pd.read_excel(excel_file_path, sheet_name=0, header=header_row_idx_pandas)
        log_message_ui(f"อ่าน Excel สำเร็จ: พบ {len(df)} แถว.")

        # ตรวจสอบคอลัมน์ ID หรือ Parent ID (ทำกับ df ดั้งเดิม)
        if is_header_type:
            if id_column_excel_header and (id_column_excel_header not in df.columns):
                 log_message_ui(f"!!! คำเตือน: ไม่พบคอลัมน์ ID หลัก '{id_column_excel_header}' ในไฟล์ Excel (ก่อนกรองสรุป) !!!")
        else: # Detail type
            parent_id_col_excel = config.get('parent_id_column_name_excel')
            if parent_id_col_excel and (parent_id_col_excel not in df.columns):
                log_message_ui(f"!!! คำเตือน: ไม่พบคอลัมน์ Parent ID '{parent_id_col_excel}' ในไฟล์ Excel Detail (ก่อนกรองสรุป) !!!")


        # 2. กรองแถวสรุปออกจาก DataFrame หลัก (df) เพื่อได้ df_filtered
        df_filtered = df.copy() # เริ่มจาก copy df ทั้งหมด
        if summary_keyword and summary_column_idx >= 0:
            if summary_column_idx < len(df.columns): # ตรวจสอบว่า index อยู่ในขอบเขต
                summary_col_name = df.columns[summary_column_idx]
                # กรองแถวที่คอลัมน์ summary_col_name *ไม่มี* คำว่า summary_keyword (case-insensitive)
                df_filtered = df[~df[summary_col_name].astype(str).str.contains(summary_keyword, na=False, case=False)].copy()
                log_message_ui(f"กรองแถว '{summary_keyword}' ออกแล้ว เหลือ {len(df_filtered)} แถว.")
            else:
                log_message_ui(f"คำเตือน: Index คอลัมน์สรุป ({summary_column_idx}) อยู่นอกช่วงของ DataFrame ({len(df.columns)} คอลัมน์). ไม่สามารถกรองแถวสรุป.")
                # df_filtered จะยังคงเป็น df ทั้งหมดถ้า index ผิด
        elif summary_keyword: # มี keyword แต่ index อาจจะไม่ได้ตั้งค่า (เช่น -1)
             log_message_ui(f"คำเตือน: ไม่ได้ระบุ index คอลัมน์สรุปที่ถูกต้อง ({summary_column_idx}) สำหรับ '{summary_keyword}'. ไม่สามารถกรองแถวสรุป.")
             # df_filtered จะยังคงเป็น df ทั้งหมด

        # --- สร้างและ clean df_cleaned_for_signature จาก df_filtered ---
        df_cleaned_for_signature = pd.DataFrame() # Initialize
        if not df_filtered.empty:
            excel_columns_list_for_cleaning = df_filtered.columns.tolist()
            df_cleaned_for_signature = df_filtered.copy()
            for col in excel_columns_list_for_cleaning:
                # 1. Fill NaN with empty string FIRST
                df_cleaned_for_signature[col] = df_cleaned_for_signature[col].fillna('')
                # 2. Convert to string type
                df_cleaned_for_signature[col] = df_cleaned_for_signature[col].astype(str)
                # 3. Strip whitespace from string
                df_cleaned_for_signature[col] = df_cleaned_for_signature[col].str.strip()

                # 4. Clean numeric-like strings (remove .0 for integers, remove commas)
                def clean_numeric_like_string(val_str):
                    cleaned_val = val_str
                    # Remove comma first
                    cleaned_val = cleaned_val.replace(',', '')
                    # Remove .0 if it looks like an integer float
                    if cleaned_val.endswith(".0") and cleaned_val[:-2].isdigit(): # Check if part before .0 is digit
                        cleaned_val = cleaned_val[:-2]
                    elif cleaned_val.endswith(".00") and cleaned_val[:-3].isdigit(): # For cases like 1,234.00
                        cleaned_val = cleaned_val[:-3]
                    return cleaned_val

                # Apply cleaning for numeric-like strings
                # This is a heuristic. A more robust way would be to identify actual numeric columns
                # or apply this cleaning more selectively based on expected column types.
                # For now, try to apply if it doesn't break non-numeric strings.
                # A simple check: if the original (stripped) string can be converted to a number without error
                # then apply the cleaning.
                temp_series_for_check = df_cleaned_for_signature[col].copy()
                # Attempt to apply cleaning, if it causes issues for a column, it might need specific handling
                try:
                    # Only apply if the column *could* contain numbers that pandas might format with .0
                    # This is tricky because a column can have mixed types.
                    # A safer bet is to apply and see if it breaks things, or be more specific.
                    # For now, let's assume we want to apply this cleaning to all string columns.
                    # If a column is purely text, replace and endswith won't do much.
                    df_cleaned_for_signature[col] = df_cleaned_for_signature[col].apply(clean_numeric_like_string)
                except Exception as e_clean_num_str:
                    log_message_ui(f"Notice: Could not apply numeric string cleaning to column '{col}'. Error: {e_clean_num_str}")
                    # Keep the stripped string if cleaning fails

            log_message_ui(f"DEBUG: df_cleaned_for_signature head (first 2 rows):\n{df_cleaned_for_signature.head(2).to_string(index=False)}") # Added index=False
        else:
            # ถ้า df_filtered ว่าง, df_cleaned_for_signature ก็ควรจะว่าง
            # และควรจะมี header ที่ถูกต้องถ้าเป็นไปได้ (จาก df_filtered เดิม ก่อนจะรู้ว่ามัน empty)
            if hasattr(df_filtered, 'columns') and len(df_filtered.columns) > 0 :
                 df_cleaned_for_signature = pd.DataFrame(columns=df_filtered.columns)
            else:
                 df_cleaned_for_signature = pd.DataFrame() # Fallback
            log_message_ui("DEBUG: df_filtered was empty, df_cleaned_for_signature is an empty DataFrame.")

        # ตรวจสอบ df_cleaned_for_signature อีกครั้ง
        if df_cleaned_for_signature.empty and not df_filtered.empty:
            log_message_ui(f"Warning: df_cleaned_for_signature became empty after initial setup/cleaning, but df_filtered was not. This might indicate an issue or an empty Excel file after filtering summary rows.")
            # If df_filtered had content but df_cleaned_for_signature is empty, it means the Excel file itself (after summary filter) was empty.
            # The logic below will handle an empty df_cleaned_for_signature.


        # 3. เชื่อมต่อ Google Sheet
        log_message_ui(f"3. กำลังเชื่อมต่อ Google Spreadsheet ID: {google_sheet_id_main}...")
        client = authenticate_google_sheets()
        if not client:
            raise ConnectionError("Authentication failed with Google Sheets.")

        spreadsheet = client.open_by_key(google_sheet_id_main)
        log_message_ui(f"เปิด Spreadsheet '{spreadsheet.title}' สำเร็จ")

        try:
            worksheet = spreadsheet.worksheet(target_sheet_name)
            log_message_ui(f"พบ Sheet '{target_sheet_name}'")
        except gspread.exceptions.WorksheetNotFound:
            log_message_ui(f"ไม่พบ Sheet '{target_sheet_name}', กำลังสร้างใหม่...")
            # ใช้ df_cleaned_for_signature (ซึ่งอาจจะว่างเปล่า) ในการกำหนดจำนวนคอลัมน์
            # ถ้า df_cleaned_for_signature ว่างแต่มี columns, ให้ใช้จำนวน columns นั้น
            # ถ้า df_cleaned_for_signature ว่างและไม่มี columns, ใช้ default
            num_cols_for_new_sheet = len(df_cleaned_for_signature.columns) if not df_cleaned_for_signature.empty or (hasattr(df_cleaned_for_signature, 'columns') and len(df_cleaned_for_signature.columns) > 0) else 26
            worksheet = spreadsheet.add_worksheet(title=target_sheet_name, rows="100", cols=num_cols_for_new_sheet) # Adjusted rows for new sheet
            log_message_ui(f"สร้าง Sheet '{target_sheet_name}' สำเร็จ")


        # 4. กรองข้อมูล Excel (df_to_upload จะถูกกำหนดค่าในส่วนนี้)
        if is_header_type:
            log_message_ui(f"4. กำลังประมวลผล '{doc_display_name}' (Header)...")
            # สำหรับ Header type, เราจะใช้ df_cleaned_for_signature ในการกรอง
            pass
        if df_cleaned_for_signature.empty:
            log_message_ui(
                f"ไม่มีข้อมูลใน Excel (df_cleaned_for_signature) ที่จะประมวลผลสำหรับ '{doc_display_name}' (Detail).")
            df_to_upload = pd.DataFrame(
                columns=df_cleaned_for_signature.columns if hasattr(df_cleaned_for_signature, 'columns') else [])
        else:
            parent_id_col_name_excel = config.get('parent_id_column_name_excel')
            line_item_id_col_name_excel = config.get('line_item_id_column_excel',
                                                     'รายการที่')  # Default เป็น 'รายการที่'

            if not parent_id_col_name_excel or parent_id_col_name_excel not in df_cleaned_for_signature.columns:
                log_message_ui(
                    f"!!! Error: ไม่พบคอลัมน์ Parent ID '{parent_id_col_name_excel}' ใน Excel สำหรับ Detail type. ไม่สามารถดำเนินการต่อ.")
                messagebox.showerror("Config Error",
                                     f"ไม่พบคอลัมน์ Parent ID '{parent_id_col_name_excel}' สำหรับ Detail.")
                if progressbar: progressbar.stop(); progressbar.set(0)
                return
            if not line_item_id_col_name_excel or line_item_id_col_name_excel not in df_cleaned_for_signature.columns:
                log_message_ui(
                    f"!!! Error: ไม่พบคอลัมน์ Line Item ID '{line_item_id_col_name_excel}' ใน Excel สำหรับ Detail type. ไม่สามารถดำเนินการต่อ.")
                messagebox.showerror("Config Error",
                                     f"ไม่พบคอลัมน์ Line Item ID '{line_item_id_col_name_excel}' สำหรับ Detail.")
                if progressbar: progressbar.stop(); progressbar.set(0)
                return

            log_message_ui(
                f"DEBUG: Using Parent ID column: '{parent_id_col_name_excel}', Line Item ID column: '{line_item_id_col_name_excel}'")

            existing_gsheet_data_records = []
            try:
                header_values_gsheet = worksheet.row_values(1)
                if header_values_gsheet and any(h.strip() for h in header_values_gsheet):
                    existing_gsheet_data_records = worksheet.get_all_records(empty2zero=False, head=1, default_blank='')
                    log_message_ui(
                        f"พบข้อมูล {len(existing_gsheet_data_records)} แถวใน Google Sheet '{target_sheet_name}'.")
                else:
                    log_message_ui(f"Google Sheet '{target_sheet_name}' ว่างเปล่าหรือไม่พบ Header.")
            except Exception as e_get_gsheet:
                log_message_ui(
                    f"Error ขณะดึงข้อมูลจาก Google Sheet '{target_sheet_name}': {e_get_gsheet}. จะถือว่าชีตว่าง.")
                existing_gsheet_data_records = []

            # สร้าง Set ของ (ParentID, LineItemID) และ Set ของ ParentID จาก Google Sheet
            existing_parent_ids_gsheet = set()
            existing_parent_line_item_keys_gsheet = set()

            if existing_gsheet_data_records:
                # พยายาม clean ค่าจาก GSheet ก่อนสร้าง key set
                for i, record_dict in enumerate(existing_gsheet_data_records):
                    try:
                        parent_id_gsheet_val = str(record_dict.get(parent_id_col_name_excel, '')).strip()
                        line_item_id_gsheet_val = str(record_dict.get(line_item_id_col_name_excel, '')).strip()

                        # (Optional but recommended) Clean numeric-like strings for GSheet keys
                        if parent_id_gsheet_val.endswith(".0") and parent_id_gsheet_val[
                                                                   :-2].isdigit(): parent_id_gsheet_val = parent_id_gsheet_val[
                                                                                                          :-2]
                        if line_item_id_gsheet_val.endswith(".0") and line_item_id_gsheet_val[
                                                                      :-2].isdigit(): line_item_id_gsheet_val = line_item_id_gsheet_val[
                                                                                                                :-2]
                        parent_id_gsheet_val = parent_id_gsheet_val.replace(',', '')
                        line_item_id_gsheet_val = line_item_id_gsheet_val.replace(',', '')

                        if parent_id_gsheet_val:  # Ensure Parent ID is not empty
                            existing_parent_ids_gsheet.add(parent_id_gsheet_val)
                            if line_item_id_gsheet_val:  # Ensure Line Item ID is not empty for the combined key
                                existing_parent_line_item_keys_gsheet.add(
                                    (parent_id_gsheet_val, line_item_id_gsheet_val))
                            elif parent_id_gsheet_val and not line_item_id_gsheet_val:
                                # Handle cases where a parent ID might exist but line item is missing/blank in GSheet for some reason
                                # This scenario might need specific business logic if it occurs.
                                # For now, if line item is blank, it won't be in existing_parent_line_item_keys_gsheet
                                pass
                        if i < 5:
                            log_message_ui(
                                f"DEBUG: GSheet Key {i}: Parent='{parent_id_gsheet_val}', LineItem='{line_item_id_gsheet_val}'")

                    except Exception as e_gsheet_key_creation:
                        log_message_ui(
                            f"Warning: ไม่สามารถสร้าง key จาก GSheet Record {i}: {record_dict}. Error: {e_gsheet_key_creation}")

            log_message_ui(f"DEBUG: GSheet Parent IDs count: {len(existing_parent_ids_gsheet)}")
            log_message_ui(f"DEBUG: GSheet Parent-LineItem Keys count: {len(existing_parent_line_item_keys_gsheet)}")

            new_rows_data = []
            excel_cols = df_cleaned_for_signature.columns.tolist()

            for index, excel_row_series in df_cleaned_for_signature.iterrows():
                try:
                    # ค่าใน excel_row_series เป็น string ที่ clean แล้ว
                    parent_id_excel_val = excel_row_series[parent_id_col_name_excel]
                    line_item_id_excel_val = excel_row_series[line_item_id_col_name_excel]

                    if not parent_id_excel_val:  # ข้ามแถว Excel ถ้า Parent ID ว่าง
                        log_message_ui(
                            f"DEBUG: ข้ามแถว Excel Index {index} เนื่องจาก Parent ID ('{parent_id_col_name_excel}') ว่างเปล่า.")
                        continue

                    # กรณีที่ 1: Parent ID จาก Excel ไม่มีใน GSheet เลย
                    if parent_id_excel_val not in existing_parent_ids_gsheet:
                        new_rows_data.append(excel_row_series.tolist())
                        if index < 10:  # Log for first few new parent ID entries
                            log_message_ui(
                                f"DEBUG: Excel Row {index} - ADDING (New Parent ID): ('{parent_id_excel_val}', '{line_item_id_excel_val}')")
                    else:
                        # กรณีที่ 2: Parent ID มีใน GSheet, เช็ค (Parent ID, Line Item ID)
                        current_excel_key = (parent_id_excel_val, line_item_id_excel_val)
                        if not line_item_id_excel_val:  # ถ้า Line Item ID ใน Excel ว่าง
                            log_message_ui(
                                f"DEBUG: Excel Row {index} - SKIPPING (Parent ID '{parent_id_excel_val}' exists, but Line Item ID in Excel is blank).")
                            # หรือคุณอาจจะต้องการ insert ถ้า line_item_id_excel_val ว่าง:
                            # new_rows_data.append(excel_row_series.tolist())
                            # log_message_ui(f"DEBUG: Excel Row {index} - ADDING (Parent ID exists, Line Item ID in Excel is blank).")
                            continue  # ปัจจุบันคือข้าม ถ้า line item ใน excel ว่างและ parent มีแล้ว

                        if current_excel_key not in existing_parent_line_item_keys_gsheet:
                            new_rows_data.append(excel_row_series.tolist())
                            if index < 10 or len(new_rows_data) < 5:  # Log for first few new line item entries
                                log_message_ui(
                                    f"DEBUG: Excel Row {index} - ADDING (Parent Exists, New Line Item): {current_excel_key}")
                        else:
                            if index < 10:  # Log for first few skipped entries
                                log_message_ui(f"DEBUG: Excel Row {index} - SKIPPING (Key Exists): {current_excel_key}")
                except KeyError as ke:
                    log_message_ui(
                        f"!!! KeyError ขณะประมวลผลแถว Excel Index {index}: ไม่พบคอลัมน์ {ke}. ตรวจสอบการตั้งค่าชื่อคอลัมน์.")
                    continue  # ข้ามแถวที่มีปัญหา
                except Exception as e_excel_row_proc:
                    log_message_ui(f"Warning: เกิดปัญหาขณะประมวลผลแถว Excel Index {index}. Error: {e_excel_row_proc}")

            if new_rows_data:
                df_to_upload = pd.DataFrame(new_rows_data, columns=excel_cols)
                log_message_ui(
                    f"พบ {len(df_to_upload)} รายการใหม่ (Detail) ที่จะ Import ตามเงื่อนไข 'เลขที่เอกสาร' และ 'รายการที่'.")
            else:
                df_to_upload = pd.DataFrame(columns=excel_cols)
                log_message_ui("ไม่พบรายการใหม่ (Detail) ที่จะ Import ตามเงื่อนไข 'เลขที่เอกสาร' และ 'รายการที่'.")

        # --- จบส่วน if is_header_type / else (Detail type) ---

        # ตรวจสอบ df_to_upload อีกครั้งก่อนดำเนินการต่อ
        if df_to_upload.empty:
            log_message_ui(f"ไม่มีข้อมูลใหม่ที่จะอัปโหลดสำหรับ '{doc_display_name}'.")
            # ไม่ raise error, แสดง info message และจบ process อย่างสงบ
            app.after(0, lambda: messagebox.showinfo("Import Info", f"ไม่มีข้อมูลใหม่ที่จะอัปโหลดสำหรับ '{doc_display_name}'"))
            if progressbar: progressbar.stop(); progressbar.set(0)
            return # จบการทำงานของฟังก์ชันนี้ถ้าไม่มีอะไรให้อัปโหลด


        # 5. เตรียมข้อมูลและอัปโหลด (df_to_upload ตอนนี้คือข้อมูลที่ผ่านการกรองแล้ว)
        df_to_upload_formatted = df_to_upload.copy()
        if date_columns_to_format:
            for col_name_date in date_columns_to_format:
                if col_name_date in df_to_upload_formatted.columns:
                    df_to_upload_formatted[col_name_date] = df_to_upload_formatted[col_name_date].astype(str).apply(
                        lambda x: f"'{x.strip()}" if x and x.lower() != 'nan' and x.strip() != '' else x
                    )
                    log_message_ui(f"เพิ่ม single quote ให้คอลัมน์วันที่: '{col_name_date}'")

        df_to_upload_cleaned = df_to_upload_formatted.replace([np.inf, -np.inf], np.nan).fillna('')

        if not worksheet: raise RuntimeError(f"Worksheet '{target_sheet_name}' is not available.")

        # ตรวจสอบว่า worksheet มี header หรือยัง (อาจจะสร้างใหม่และยังไม่มี header)
        # หรืออาจจะมีข้อมูลเดิมอยู่แล้ว
        existing_data_headers = worksheet.row_values(1) # ดึงแถวแรกมาดู

        if not existing_data_headers or not any(h.strip() for h in existing_data_headers):  # Sheet ว่าง หรือ แถวแรกว่างเปล่า
            log_message_ui(f"Sheet '{target_sheet_name}' ว่างเปล่า หรือไม่มี header. กำลังอัปโหลดพร้อม header...")
            if not df_to_upload_cleaned.empty: # ตรวจสอบอีกครั้งว่ามีข้อมูลจะอัปโหลดจริง
                worksheet.clear() # ล้างชีตก่อนถ้าจะใส่ header + data ใหม่
                data_to_gsheet = [df_to_upload_cleaned.columns.values.tolist()] + df_to_upload_cleaned.values.tolist()
                if data_to_gsheet and data_to_gsheet[0]: # ตรวจสอบว่ามี header จริงๆ
                    worksheet.update(data_to_gsheet, 'A1', value_input_option='USER_ENTERED')
            else:
                log_message_ui("ไม่มีข้อมูลที่จะอัปโหลด (ชีตว่าง)")

        else:  # Sheet มีข้อมูล (และมี header แล้ว)
            log_message_ui(f"Sheet '{target_sheet_name}' มีข้อมูล. กำลังเพิ่มข้อมูลใหม่ต่อท้าย (ถ้ามี)...")
            if not df_to_upload_cleaned.empty: # ตรวจสอบอีกครั้ง
                data_to_gsheet = df_to_upload_cleaned.values.tolist()
                if data_to_gsheet: # ตรวจสอบว่า list ไม่ว่าง
                    worksheet.append_rows(data_to_gsheet, value_input_option='USER_ENTERED', table_range='A1')
            else:
                log_message_ui("ไม่มีข้อมูลใหม่ที่จะเพิ่มต่อท้าย (ข้อมูลอาจซ้ำทั้งหมด)")


        log_message_ui(f"\n--- อัปโหลดข้อมูลสำหรับ '{doc_display_name}' สำเร็จ! ---")
        log_message_ui(f"ดูผลลัพธ์: https://docs.google.com/spreadsheets/d/{google_sheet_id_main}/edit#gid={worksheet.id}")
        app.after(0, lambda: messagebox.showinfo("Success", f"อัปโหลดข้อมูลสำหรับ '{doc_display_name}' สำเร็จ!"))

        current_app_configs[doc_type_key]['last_used_excel_dir'] = os.path.dirname(excel_file_path) if os.path.isfile(excel_file_path) else excel_file_path
        current_app_configs[doc_type_key]['last_used_gsheet_id_input'] = google_sheet_id_main
        save_app_settings(current_app_configs)

    except FileNotFoundError as e_fnf:
        log_message_ui(f"Error: ไม่พบไฟล์: {e_fnf}")
        app.after(0, lambda: messagebox.showerror("File Error", f"ไม่พบไฟล์: {e_fnf}"))
    except gspread.exceptions.SpreadsheetNotFound:
        log_message_ui(f"Error: ไม่พบ Google Spreadsheet หลักด้วย ID: {google_sheet_id_main}")
        app.after(0, lambda: messagebox.showerror("Google Sheet Error", f"ไม่พบ Google Spreadsheet หลักด้วย ID: {google_sheet_id_main}"))
    except ConnectionError as e_conn:
        log_message_ui(f"Connection Error: {e_conn}")
    except ValueError as e_val: # เช่น "No data in Excel file after filtering summary rows."
        log_message_ui(f"Info/Error: {e_val}")
        app.after(0, lambda: messagebox.showerror("Data Error", f"{e_val} (สำหรับ {doc_display_name})"))
    except RuntimeError as e_rt:
        log_message_ui(f"Runtime Error: {e_rt}")
        app.after(0, lambda: messagebox.showerror("Runtime Error", str(e_rt)))
    except Exception as e_proc:
        log_message_ui(f"Error ในกระบวนการ Import '{doc_display_name}': {e_proc}")
        import traceback
        tb_str = traceback.format_exc()
        log_message_ui(tb_str)
        app.after(0, lambda: messagebox.showerror("Import Error", f"เกิดข้อผิดพลาดในกระบวนการ Import '{doc_display_name}':\n{e_proc}"))
    finally:
        if progressbar:
            progressbar.stop()
            progressbar.set(0)
        log_message_ui(f"--- สิ้นสุดกระบวนการสำหรับ: {doc_display_name} ---")

def get_last_id_from_gsheet(worksheet, id_column_letter, doc_type_display_name, id_prefix=""):
    """ดึง ID ล่าสุดจาก Google Sheet, ปรับให้รับ prefix ได้"""
    try:
        log_message_ui(f"กำลังดึงข้อมูล ID ล่าสุดของ '{doc_type_display_name}' จากคอลัมน์ {id_column_letter}...")
        # แปลงตัวอักษรคอลัมน์เป็น index (1-based for gspread)
        col_index = gspread.utils.a1_to_rowcol(f"{id_column_letter}1")[1]
        id_values = worksheet.col_values(col_index)

        # กรอง ID ที่ถูกต้อง (อาจจะต้องปรับปรุงตามรูปแบบ ID ของแต่ละประเภท)
        if id_prefix:
            valid_ids = sorted([val for val in id_values if val and str(val).strip().startswith(id_prefix)])
        else:
            # ถ้าไม่มี prefix, กรองค่าว่างและอาจจะ header (สมมติ header ไม่ใช่ตัวเลขหรือรูปแบบ ID ที่ยาวๆ)
            # และ .strip() เพื่อตัดช่องว่างหัวท้ายก่อนตรวจสอบ
            valid_ids = sorted([val for val in id_values if val and str(val).strip() and not str(val).strip().isspace() and len(str(val).strip()) > 2]) # ปรับปรุงการกรอง

        if valid_ids:
            last_id = valid_ids[-1].strip() # .strip() อีกครั้งเพื่อให้แน่ใจ
            log_message_ui(f"ID ล่าสุดของ '{doc_type_display_name}' ใน Google Sheet: {last_id}")
            return last_id
        else:
            log_message_ui(f"ไม่พบ ID ของ '{doc_type_display_name}' ใน Google Sheet หรือยังไม่มีข้อมูล")
            return None
    except Exception as e:
        log_message_ui(f"เกิดข้อผิดพลาดในการดึง ID ล่าสุดของ '{doc_type_display_name}': {e}")
        import traceback
        log_message_ui(traceback.format_exc()) # เพิ่ม traceback สำหรับ debug
        return None

def load_app_settings():
    global current_app_configs
    try:
        if os.path.exists(APP_SETTINGS_FILE):
            with open(APP_SETTINGS_FILE, "rb") as f:
                loaded_configs = pickle.load(f)
                for doc_type, default_conf in DEFAULT_DOCUMENT_CONFIGS.items():
                    if doc_type not in loaded_configs:
                        loaded_configs[doc_type] = default_conf
                    else:
                        for key, value in default_conf.items():
                            if key not in loaded_configs[doc_type]:
                                loaded_configs[doc_type][key] = value
                current_app_configs = loaded_configs
                if app and hasattr(app, 'log_message_ui'):
                    app.log_message_ui("โหลดการตั้งค่าแอปพลิเคชันแล้ว")
                else:
                    print(f"โหลดการตั้งค่าแอปพลิเคชันจาก {APP_SETTINGS_FILE}")
                return current_app_configs
        else:
            if app and hasattr(app, 'log_message_ui'):
                app.log_message_ui("ไม่พบไฟล์ตั้งค่า ใช้ค่าเริ่มต้น")
            else:
                print("ไม่พบไฟล์ตั้งค่า ใช้ค่าเริ่มต้น")
            current_app_configs = DEFAULT_DOCUMENT_CONFIGS.copy()
            return current_app_configs
    except Exception as e:
        if app and hasattr(app, 'log_message_ui'):
            app.log_message_ui(f"Error โหลดการตั้งค่าแอปพลิเคชัน: {e}. ใช้ค่าเริ่มต้น")
        else:
            print(f"Error โหลดการตั้งค่าแอปพลิเคชัน: {e}. ใช้ค่าเริ่มต้น")
        current_app_configs = DEFAULT_DOCUMENT_CONFIGS.copy()
        return current_app_configs

def log_message_ui(message):
    if log_textbox and app: # ตรวจสอบ app ด้วย
        log_textbox.configure(state="normal")
        log_textbox.insert(ctk.END, str(message) + "\n")
        log_textbox.configure(state="disabled")
        log_textbox.see(ctk.END)
        app.update_idletasks()
    else:
        print(message)

def save_settings(excel_path, sheet_id):
    settings = {"excel_path": excel_path, "sheet_id": sheet_id}
    try:
        with open(SETTINGS_FILE, "wb") as f:
            pickle.dump(settings, f)
        log_message_ui("บันทึกการตั้งค่าล่าสุดแล้ว")
    except Exception as e:
        log_message_ui(f"Error บันทึกการตั้งค่า: {e}")

def load_settings():
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "rb") as f:
                settings = pickle.load(f)
                log_message_ui("โหลดการตั้งค่าล่าสุดแล้ว")
                return settings.get("excel_path", DEFAULT_EXCEL_DIR), settings.get("sheet_id", DEFAULT_GOOGLE_SHEET_ID)
    except Exception as e:
        log_message_ui(f"Error โหลดการตั้งค่า: {e}")
    return DEFAULT_EXCEL_DIR, DEFAULT_GOOGLE_SHEET_ID


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        global app, log_textbox, progressbar, log_frame_visible, current_app_configs
        app = self

        # --- โหลดการตั้งค่าทั้งหมดของแอปพลิเคชัน ---
        current_app_configs = load_app_settings()

        self.title("Excel to Google Sheet Importer v1.2")
        self.geometry("700x650") # เพิ่มความสูงอีกเล็กน้อยสำหรับ OptionMenu
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # --- ตัวแปรสำหรับ UI state ---
        self.selected_doc_type_key = None # Key ของประเภทเอกสารที่เลือก (เช่น "PO", "QO")
        self.current_excel_file_path_var = ctk.StringVar(value="ยังไม่ได้เลือกไฟล์ Excel") # สำหรับ label ไฟล์ excel

        main_frame = ctk.CTkFrame(self)
        main_frame.pack(pady=10, padx=20, fill="both", expand=True)

        input_frame = ctk.CTkFrame(main_frame)
        input_frame.pack(pady=(0,10), padx=10, fill="x")

        # --- 1. OptionMenu สำหรับเลือกประเภทเอกสาร ---
        doc_type_frame = ctk.CTkFrame(input_frame)
        doc_type_frame.pack(pady=5, fill="x")

        doc_type_label = ctk.CTkLabel(doc_type_frame, text="ประเภทเอกสาร:")
        doc_type_label.pack(side=ctk.LEFT, padx=(0, 10))

        # สร้าง list ของ display names สำหรับ OptionMenu
        self.doc_type_display_names = [conf['display_name'] for conf in current_app_configs.values()]
        # สร้าง mapping จาก display name กลับไปเป็น key
        self.display_name_to_key_map = {conf['display_name']: key for key, conf in current_app_configs.items()}

        self.doc_type_var = ctk.StringVar(value=self.doc_type_display_names[0] if self.doc_type_display_names else "N/A")
        self.doc_type_optionmenu = ctk.CTkOptionMenu(doc_type_frame,
                                                     values=self.doc_type_display_names,
                                                     variable=self.doc_type_var,
                                                     command=self.on_doc_type_selected)
        self.doc_type_optionmenu.pack(side=ctk.LEFT, expand=True, fill="x")


        # --- 2. ส่วนเลือกไฟล์ Excel ---
        excel_frame = ctk.CTkFrame(input_frame)
        excel_frame.pack(pady=5, fill="x")
        self.excel_path_label = ctk.CTkLabel(excel_frame, textvariable=self.current_excel_file_path_var, width=350, anchor="w") # ใช้ textvariable
        self.excel_path_label.pack(side=ctk.LEFT, padx=(0,10), expand=True, fill="x")
        self.select_excel_button = ctk.CTkButton(excel_frame, text="เลือกไฟล์ Excel", command=self.select_excel_file)
        self.select_excel_button.pack(side=ctk.LEFT)

        # --- 3. ส่วน Google Sheet ID (Spreadsheet หลัก) ---
        gsheet_frame = ctk.CTkFrame(input_frame)
        gsheet_frame.pack(pady=5, fill="x")
        gsheet_id_label = ctk.CTkLabel(gsheet_frame, text="Google Sheet ID (Spreadsheet หลัก):")
        gsheet_id_label.pack(side=ctk.LEFT, padx=(0,10))
        self.gsheet_id_entry = ctk.CTkEntry(gsheet_frame, placeholder_text="ใส่ Google Sheet ID ของ Spreadsheet หลัก", width=250)
        self.gsheet_id_entry.pack(side=ctk.LEFT, expand=True, fill="x")

        # --- 4. Progress Bar, ปุ่ม Import, Log Area (เหมือนเดิม) ---
        progressbar = ctk.CTkProgressBar(main_frame, orientation="horizontal", mode="determinate")
        progressbar.pack(pady=(5, 10), padx=10, fill="x")
        progressbar.set(0)

        self.import_button = ctk.CTkButton(main_frame, text="เริ่ม Import", command=self.start_import_thread, height=40, font=("Arial", 14, "bold"))
        self.import_button.pack(pady=10)

        self.log_outer_frame = ctk.CTkFrame(main_frame)
        self.log_outer_frame.pack(pady=10, padx=10, fill="both", expand=True)
        log_header_frame = ctk.CTkFrame(self.log_outer_frame, fg_color="transparent")
        log_header_frame.pack(fill="x", pady=(0,5))
        log_label_title = ctk.CTkLabel(log_header_frame, text="สถานะการทำงาน:")
        log_label_title.pack(side=ctk.LEFT, anchor="w")
        self.toggle_log_button = ctk.CTkButton(log_header_frame, text="ซ่อน Log", width=80, command=self.toggle_log_visibility)
        self.toggle_log_button.pack(side=ctk.RIGHT)
        log_textbox = ctk.CTkTextbox(self.log_outer_frame, height=180, state="disabled", wrap="word")
        if log_frame_visible:
            log_textbox.pack(fill="both", expand=True)
        else:
            self.toggle_log_button.configure(text="แสดง Log")

        # --- เรียก on_doc_type_selected ครั้งแรกเพื่อตั้งค่า UI ตาม default ---
        if self.doc_type_display_names:
            self.on_doc_type_selected(self.doc_type_var.get()) # ใช้ค่า default ของ OptionMenu
        else:
            log_message_ui("คำเตือน: ไม่มีการตั้งค่าประเภทเอกสาร")

        log_message_ui("โปรแกรมพร้อมทำงาน กรุณาเลือกประเภทเอกสาร ไฟล์ และใส่ Google Sheet ID")

    def toggle_log_visibility(self):
        global log_frame_visible, log_textbox
        log_frame_visible = not log_frame_visible
        if log_frame_visible:
            log_textbox.pack(fill="both", expand=True, before=None)  # pack it back
            self.toggle_log_button.configure(text="ซ่อน Log")
        else:
            log_textbox.pack_forget()  # hide it
            self.toggle_log_button.configure(text="แสดง Log")

    def select_excel_file(self):
        if not self.selected_doc_type_key:
            messagebox.showwarning("คำเตือน", "กรุณาเลือกประเภทเอกสารก่อนเลือกไฟล์")
            return

        # ใช้ self.current_excel_dir ที่ถูกตั้งค่าโดย on_doc_type_selected
        file_path = filedialog.askopenfilename(
            initialdir=self.current_excel_dir,
            title=f"เลือกไฟล์สำหรับ {current_app_configs[self.selected_doc_type_key]['display_name']}",
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )
        if file_path:
            self.current_excel_file_path_var.set(file_path) # อัปเดต StringVar
            # อัปเดต current_excel_dir สำหรับประเภทเอกสารที่เลือกอยู่
            # เพื่อให้ครั้งถัดไปที่เลือกไฟล์สำหรับประเภทเดิม จะเปิดที่โฟลเดอร์นี้
            current_app_configs[self.selected_doc_type_key]['last_used_excel_dir'] = os.path.dirname(file_path)
            # ไม่ต้อง save_app_settings ทันที รอจนกว่าจะ import สำเร็จ
            log_message_ui(f"เลือกไฟล์: {file_path}")

    def on_doc_type_selected(self, selected_display_name):
        """Event handler เมื่อผู้ใช้เลือกประเภทเอกสาร"""
        self.selected_doc_type_key = self.display_name_to_key_map.get(selected_display_name)
        if not self.selected_doc_type_key:
            log_message_ui(f"Error: ไม่พบ key สำหรับ display name '{selected_display_name}'")
            return

        config = current_app_configs.get(self.selected_doc_type_key)
        if not config:
            log_message_ui(f"Error: ไม่พบ config สำหรับ key '{self.selected_doc_type_key}'")
            return

        log_message_ui(f"เลือกประเภทเอกสาร: {config['display_name']}")

        # อัปเดต Google Sheet ID entry
        self.gsheet_id_entry.delete(0, ctk.END)
        self.gsheet_id_entry.insert(0, config.get('last_used_gsheet_id_input', ""))

        # อัปเดต current_excel_dir (สำหรับ File Dialog)
        self.current_excel_dir = config.get('last_used_excel_dir', os.path.expanduser("~"))

        # อัปเดต Label แสดงไฟล์ Excel (ถ้ามีไฟล์ที่จำไว้สำหรับประเภทนี้)
        # เนื่องจากเราไม่ได้เก็บ path ไฟล์ล่าสุดโดยตรงใน config แต่เป็น dir
        # เราจะรีเซ็ต label ไฟล์เมื่อเปลี่ยนประเภทเอกสาร
        self.current_excel_file_path_var.set("ยังไม่ได้เลือกไฟล์ Excel")
        # (ถ้าต้องการให้จำไฟล์ล่าสุดของแต่ละประเภท จะต้องเพิ่ม key 'last_used_excel_file' ใน config)

        # (Optional) แสดงชื่อ Sheet ปลายทางใน UI ถ้าต้องการ
        # self.target_sheet_label.configure(text=f"Sheet ปลายทาง: {config['target_sheet_name_or_id']}")

    def get_last_id_from_gsheet(worksheet, id_column_letter, doc_type_display_name, id_prefix=""):
        """ดึง ID ล่าสุดจาก Google Sheet, ปรับให้รับ prefix ได้"""
        try:
            log_message_ui(f"กำลังดึงข้อมูล ID ล่าสุดของ '{doc_type_display_name}' จากคอลัมน์ {id_column_letter}...")
            id_values = worksheet.col_values(gspread.utils.a1_to_rowcol(f"{id_column_letter}1")[1])

            # กรอง ID ที่ถูกต้อง (อาจจะต้องปรับปรุงตามรูปแบบ ID ของแต่ละประเภท)
            # ตัวอย่าง: ถ้ามี prefix (เช่น PO-, QO-)
            if id_prefix:
                valid_ids = sorted([val for val in id_values if val and str(val).startswith(id_prefix)])
            else:  # ถ้าไม่มี prefix หรือรูปแบบไม่แน่นอน อาจจะต้องกรองแบบอื่น หรือแค่เอาตัวที่ไม่ใช่ header
                # สมมติ header ไม่ใช่ตัวเลขหรือรูปแบบ ID
                valid_ids = sorted(
                    [val for val in id_values if val and not str(val).isspace() and len(str(val)) > 3])  # กรองแบบง่ายๆ

            if valid_ids:
                last_id = valid_ids[-1]
                log_message_ui(f"ID ล่าสุดของ '{doc_type_display_name}' ใน Google Sheet: {last_id}")
                return last_id
            else:
                log_message_ui(f"ไม่พบ ID ของ '{doc_type_display_name}' ใน Google Sheet หรือยังไม่มีข้อมูล")
                return None
        except Exception as e:
            log_message_ui(f"เกิดข้อผิดพลาดในการดึง ID ล่าสุดของ '{doc_type_display_name}': {e}")
            return None

    def start_import_thread(self):
        if not self.selected_doc_type_key:
            messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกประเภทเอกสาร")
            return

        excel_file = self.current_excel_file_path_var.get()
        if excel_file == "ยังไม่ได้เลือกไฟล์ Excel" or not os.path.isfile(excel_file):
            messagebox.showerror("ข้อผิดพลาด", "กรุณาเลือกไฟล์ Excel ที่ถูกต้อง")
            return

        google_sheet_id_main = self.gsheet_id_entry.get() # ID ของ Spreadsheet หลัก
        if not google_sheet_id_main:
            messagebox.showerror("ข้อผิดพลาด", "กรุณาใส่ Google Sheet ID (Spreadsheet หลัก)")
            return

        self.import_button.configure(state="disabled", text="กำลัง Import...")
        if progressbar:
            progressbar.configure(mode="indeterminate")
            progressbar.start()

        # ส่ง self.selected_doc_type_key ไปยัง thread
        import_thread = threading.Thread(target=self.run_import_process,
                                         args=(excel_file, google_sheet_id_main, self.selected_doc_type_key))
        import_thread.daemon = True
        import_thread.start()



    def run_import_process(self, excel_file, google_sheet_id_main, doc_type_key): # รับ doc_type_key
        try:
            process_excel_and_gsheet(excel_file, google_sheet_id_main, doc_type_key) # ส่งต่อ
        finally:
            self.import_button.configure(state="normal", text="เริ่ม Import")
            if progressbar:
                progressbar.stop()
                progressbar.set(0)

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
            # ล้างข้อมูลเก่า (ถ้ามีอะไรค้างอยู่แบบไม่มี header)
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
    if not os.path.exists("credentials.json"):
        messagebox.showwarning("Setup Required",
                               "ไม่พบไฟล์ 'credentials.json' สำหรับ OAuth 2.0.\n"
                               "โปรแกรมอาจจะไม่สามารถยืนยันตัวตนกับ Google Sheets ได้.\n"
                               "กรุณาสร้าง OAuth 2.0 Client ID (Desktop app) จาก Google Cloud Console "
                               "และดาวน์โหลดไฟล์ JSON มาบันทึกเป็น 'credentials.json' ในโฟลเดอร์เดียวกับโปรแกรมนี้")
    app_instance = App()
    app_instance.mainloop()