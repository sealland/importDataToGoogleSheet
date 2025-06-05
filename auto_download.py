# auto_downloader.py
import os
import time # เพิ่ม import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager # ตรวจสอบว่า import ถูกต้อง
from selenium.common.exceptions import TimeoutException, NoSuchElementException # เพิ่ม import ที่อาจจำเป็น
from selenium.webdriver.common.action_chains import ActionChains
# auto_downloader.py
# ... (import และ setup อื่นๆ เหมือนเดิม) ...

def download_peak_purchase_order_report(username, password, target_business_name_to_select, # เพิ่ม parameter นี้
                                     save_directory, desired_file_name="peak_po_report.xlsx", log_callback=None):
    """
    Automates downloading the Purchase Order report from PEAK Account,
    including business selection.
    """
    print("--- [DEBUG] INSIDE download_peak_purchase_order_report FUNCTION ---")

    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(f"[PEAK_PO_Downloader_LOG] {msg}")

    _log("Function started.")
    download_path = os.path.abspath(save_directory)
    # ... (สร้าง download_path, chrome_options เหมือนเดิม) ...
    if not os.path.exists(download_path):
        try:
            os.makedirs(download_path)
            _log(f"สร้างโฟลเดอร์ดาวน์โหลด: {download_path}")
        except Exception as e_mkdir:
            _log(f"!!! Error สร้างโฟลเดอร์ดาวน์โหลด {download_path}: {e_mkdir} !!!")
            return None

    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    _log("Chrome options configured.")


    driver = None
    try:
        _log("กำลังเริ่ม WebDriver สำหรับ PEAK...")
        # ... (โค้ดเริ่ม WebDriver เหมือนเดิม) ...
        try:
            chrome_driver_path = ChromeDriverManager().install()
            _log(f"ChromeDriver path: {chrome_driver_path}")
            driver_service = ChromeService(executable_path=chrome_driver_path)
            driver = webdriver.Chrome(service=driver_service, options=chrome_options)
        except Exception as e_webdriver_init:
            _log(f"!!! Error initializing WebDriver: {e_webdriver_init} !!!")
            import traceback
            _log(traceback.format_exc())
            return None

        wait = WebDriverWait(driver, 30) # เพิ่มเวลารอโดยรวมเล็กน้อย เผื่อหน้าโหลดช้า
        short_wait = WebDriverWait(driver, 10) # สำหรับ element ที่ควรปรากฏเร็ว

        # 1. Login
        login_url = "https://secure.peakaccount.com/login"
        _log(f"กำลังไปที่หน้า Login: {login_url}")
        driver.get(login_url)
        # ... (โค้ดกรอก email, password, คลิกปุ่ม Login เหมือนเดิมที่คุณทดสอบแล้วว่าได้ผล) ...
        wait.until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(username)
        wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(password)
        login_button_xpath = "//button[normalize-space(.)='เข้าสู่ระบบ PEAK']"
        wait.until(EC.element_to_be_clickable((By.XPATH, login_button_xpath))).click()
        _log("คลิกปุ่ม Login แล้ว.")

        # --- รอและตรวจสอบการ Login (หรือตรวจสอบว่ามาถึงหน้าเลือกกิจการ) ---
        _log("รอสักครู่หลังกด Login (ประมาณ 5-10 วินาที)...")
        time.sleep(7) # รอให้ redirect หรือโหลดหน้าเลือกกิจการ
        current_url_after_login = driver.current_url
        _log(f"URL ปัจจุบันหลังจากพยายาม Login: {current_url_after_login}")

        # ตรวจสอบว่ายังอยู่หน้า login หรือไม่ หรือว่ามาถึงหน้าที่คาดหวัง (เช่น หน้าเลือกกิจการ)
        if "login" in current_url_after_login.lower():
            _log("!!! Login ไม่สำเร็จ หรือยังอยู่ที่หน้า Login !!!")
            # ... (save screenshot, quit driver, return None) ...
            try: driver.save_screenshot(os.path.join(download_path, "po_login_failed.png"))
            except: pass
            if driver: driver.quit()
            return None
        else:
            _log("Login ดูเหมือนจะสำเร็จแล้ว (ไม่ได้อยู่ที่หน้า Login).")


        # --------------------------------------------------------------------
        # ขั้นตอนที่ 1.5: เลือกกิจการ (ถ้าหน้าเว็บแสดงรายการให้เลือก)
        # --------------------------------------------------------------------
        # ตรวจสอบว่าตอนนี้อยู่หน้า "เลือกกิจการ" หรือไม่
        # คุณอาจจะต้องหา element ที่บ่งบอกว่าเป็นหน้าเลือกกิจการ เช่น title ของหน้า หรือ header บางอย่าง
        # ตัวอย่าง: ถ้าหน้าเลือกกิจการมี <H1> ที่มี text "เลือกกิจการที่ต้องการใช้งาน"
        try:
            # รอให้ element ที่บ่งบอกว่าเป็นหน้าเลือกกิจการปรากฏ (ถ้ามี)
            # หรือรอให้รายการกิจการแรกปรากฏ
            # ตัวอย่าง: รอ div ที่มี class 'list' และภายในมี p.crop.textBold
            first_business_in_list_xpath = "//div[contains(@class, 'list')]//p[contains(@class, 'crop') and contains(@class, 'textBold')]"
            _log(f"กำลังรอรายการกิจการปรากฏ (ใช้ XPATH: {first_business_in_list_xpath} สำหรับเช็ค)...")
            short_wait.until(EC.visibility_of_element_located((By.XPATH, first_business_in_list_xpath)))
            _log("หน้ารายการกิจการปรากฏแล้ว (หรือมีรายการกิจการแสดง).")

            # ทำการเลือกกิจการ
            _log(f"กำลังค้นหากิจการ '{target_business_name_to_select}' ในรายการ...")
            business_item_xpath = f"//div[contains(@class, 'list') and .//p[normalize-space(text())='{target_business_name_to_select}']]"
            # หรือถ้า div ที่คลิกได้โดยตรงคือ div.list ที่มี p นั้นอยู่ภายใน:
            # business_item_xpath = f"//div[contains(@class, 'list')][.//p[normalize-space(text())='{target_business_name_to_select}']]"

            business_element = wait.until(EC.element_to_be_clickable((By.XPATH, business_item_xpath)))
            _log(f"พบกิจการ '{target_business_name_to_select}'. กำลังคลิก...")
            business_element.click()
            _log("คลิกเลือกกิจการแล้ว. รอหน้า Dashboard/หน้าหลักของกิจการโหลด...")
            time.sleep(7) # รอให้หน้าเว็บเปลี่ยนหรือโหลดข้อมูลของกิจการนั้น

            current_url_after_select_business = driver.current_url
            _log(f"URL หลังจากเลือกกิจการ: {current_url_after_select_business}")
            if target_business_name_to_select.split(' ')[0].lower().replace('.','').replace(' ','') not in driver.title.lower().replace('.','').replace(' ',''): # เช็คแบบคร่าวๆ
                 _log(f"!!! คำเตือน: Title ของหน้าเว็บ ({driver.title}) อาจจะไม่ตรงกับกิจการที่เลือก ({target_business_name_to_select}). กรุณาตรวจสอบ URL และเนื้อหาหน้าเว็บ.")
                 # driver.save_screenshot(os.path.join(download_path, "business_selection_title_mismatch.png"))
            else:
                 _log("ดูเหมือนว่าเข้าสู่กิจการที่เลือกสำเร็จแล้ว.")

        except TimeoutException:
            _log(f"!!! ไม่พบรายการกิจการให้เลือก หรือไม่พบกิจการ '{target_business_name_to_select}' ภายในเวลาที่กำหนด !!!")
            _log("อาจจะเป็นไปได้ว่า Login แล้วเข้าสู่กิจการล่าสุดโดยอัตโนมัติ หรือมีปัญหาในการโหลดหน้ารายการกิจการ")
            # driver.save_screenshot(os.path.join(download_path, "business_list_or_item_not_found.png"))
            # ถ้าไม่เจอหน้ารายการกิจการ อาจจะลองข้ามไปขั้นตอนคลิกเมนูเลยก็ได้
            # แต่ถ้าการเลือกกิจการเป็นสิ่งจำเป็น และหาไม่เจอ ก็ควรจะ return None
            # สำหรับตอนนี้จะให้โปรแกรมพยายามทำขั้นตอนต่อไป เผื่อว่ามันเข้ากิจการ default ไปแล้ว
            _log("จะพยายามดำเนินการขั้นตอนต่อไป (คลิกเมนู) แม้ว่าจะไม่สามารถยืนยันการเลือกกิจการได้")
            pass # ให้โปรแกรมลองทำขั้นตอนต่อไป


        # --------------------------------------------------------------------
        # ขั้นตอนที่ 2: การนำทางไปยังหน้า Purchase Order ผ่านเมนู
        # (โค้ดส่วนนี้จะทำงานหลังจากเลือกกิจการแล้ว หรือถ้าไม่มีหน้าเลือกกิจการ)
        # --------------------------------------------------------------------
        _log("ขั้นตอน: การนำทางไปยังหน้า Purchase Order ผ่านเมนู...")
        # 1. คลิกไอคอน/ปุ่มเพื่อเปิดเมนูหลัก
        main_menu_trigger_xpath = "//i[@class='icon-arrow-bropdown']" # <<--- !!! ตรวจสอบ Selector นี้ให้ถูกต้อง !!!
        try:
            _log(f"กำลังค้นหาตัวเปิดเมนูหลักด้วย XPATH: {main_menu_trigger_xpath}")
            menu_trigger = wait.until(EC.element_to_be_clickable((By.XPATH, main_menu_trigger_xpath)))
            menu_trigger.click()
            _log("คลิกตัวเปิดเมนูหลักแล้ว. รอเมนูย่อยปรากฏ...")
            time.sleep(2) # เพิ่มเวลารอเมนูเปิด
        except TimeoutException:
            _log(f"!!! ไม่พบตัวเปิดเมนูหลัก หรือไม่สามารถคลิกได้ด้วย XPATH: {main_menu_trigger_xpath} !!!")
            # ... (save screenshot, quit driver, return None) ...
            try: driver.save_screenshot(os.path.join(download_path, "main_menu_trigger_not_found.png"))
            except: pass
            if driver: driver.quit()
            return None

            # --------------------------------------------------------------------
            # ขั้นตอนที่ 2: การนำทางไปยังหน้า Purchase Order ผ่านเมนู (แบบ Hover)
            # --------------------------------------------------------------------
            _log("ขั้นตอน: การนำทางไปยังหน้า Purchase Order ผ่านเมนู (Hover)...")

            # สร้าง ActionChains object
            actions = ActionChains(driver)

            # 1. Hover เหนือเมนู "รายจ่าย"
            # !!! ตรวจสอบและแก้ไข Selector ของเมนู "รายจ่าย" !!!
            # อาจจะเป็น <li>, <a>, <span> ที่มี text "รายจ่าย" หรือ class/id ที่เฉพาะเจาะจง
            expense_menu_xpath = "//a[normalize-space(.)='รายจ่าย']"  # ตัวอย่างสมมติ
            # หรือถ้าเป็นส่วนหนึ่งของ nav bar: "//ul[@id='main-nav']//a[normalize-space(.)='รายจ่าย']"
            try:
                _log(f"กำลังค้นหาเมนู 'รายจ่าย' ด้วย XPATH: {expense_menu_xpath}")
                expense_menu_element = wait.until(EC.visibility_of_element_located((By.XPATH, expense_menu_xpath)))
                _log("พบเมนู 'รายจ่าย'. กำลัง Hover...")
                actions.move_to_element(expense_menu_element).perform()  # ทำการ Hover
                _log("Hover เหนือเมนู 'รายจ่าย' แล้ว. รอเมนูย่อยปรากฏ...")
                time.sleep(1)  # รอให้เมนูย่อยมีเวลาปรากฏขึ้น
            except TimeoutException:
                _log(f"!!! ไม่พบเมนู 'รายจ่าย' หรือไม่สามารถ Hover ได้ด้วย XPATH: {expense_menu_xpath} !!!")
                # ... (save screenshot, quit driver, return None) ...
                try:
                    driver.save_screenshot(os.path.join(download_path, "expense_menu_not_found.png"))
                except:
                    pass
                if driver: driver.quit()
                return None

            # 2. Hover (ถ้าจำเป็น) หรือ คลิกเมนูย่อย "ใบสั่งซื้อ"
            # !!! ตรวจสอบและแก้ไข Selector ของเมนูย่อย "ใบสั่งซื้อ" !!!
            # เมนูย่อยนี้ควรจะ "มองเห็นได้" (visible) หลังจาก hover ที่ "รายจ่าย" แล้ว
            # po_submenu_xpath = "//a[normalize-space(.)='ใบสั่งซื้อ' and contains(@href,'/expense/PO')]" # ตัวอย่าง
            # หรือถ้ามันเป็นส่วนหนึ่งของ dropdown ที่ปรากฏ:
            po_submenu_xpath = "//ul[contains(@class,'dropdown-menu-visible')]//a[normalize-space(.)='ใบสั่งซื้อ' and contains(@href,'/expense/PO')]"  # ตัวอย่างที่อาจจะแม่นยำขึ้น

            try:
                _log(f"กำลังค้นหาเมนูย่อย 'ใบสั่งซื้อ' ด้วย XPATH: {po_submenu_xpath}")
                po_submenu_element = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, po_submenu_xpath)))  # ใช้ visibility_of_element_located เพราะมันควรจะเห็นแล้ว
                _log("พบเมนูย่อย 'ใบสั่งซื้อ'.")

                # ถ้า "ใบสั่งซื้อ" เองเป็นตัวที่ต้อง hover เพื่อให้ "ดูทั้งหมด" ปรากฏ:
                # _log("กำลัง Hover เหนือเมนูย่อย 'ใบสั่งซื้อ'...")
                # actions.move_to_element(po_submenu_element).perform()
                # _log("Hover เหนือ 'ใบสั่งซื้อ' แล้ว. รอ 'ดูทั้งหมด' ปรากฏ...")
                # time.sleep(1)
                #
                # # 3. คลิกเมนู "ดูทั้งหมด" ที่อยู่ภายใต้ "ใบสั่งซื้อ"
                # # !!! ตรวจสอบและแก้ไข Selector ของ "ดูทั้งหมด" !!!
                # view_all_po_xpath = "//a[normalize-space(.)='ดูทั้งหมด' and contains(@href,'/expense/PO?stid=0')]" # ตัวอย่าง
                # _log(f"กำลังค้นหา 'ดูทั้งหมด' สำหรับ PO ด้วย XPATH: {view_all_po_xpath}")
                # view_all_po_link = wait.until(EC.element_to_be_clickable((By.XPATH, view_all_po_xpath)))
                # view_all_po_link.click()
                # _log("คลิก 'ดูทั้งหมด' สำหรับ PO แล้ว.")

                # หรือถ้าคลิกที่ "ใบสั่งซื้อ" แล้วจะไปหน้าที่มี "ดูทั้งหมด" หรือไปหน้ารายการ PO เลย:
                _log("กำลังคลิกเมนูย่อย 'ใบสั่งซื้อ'...")
                po_submenu_element.click()  # ถ้าคลิกที่ "ใบสั่งซื้อ" แล้วไปหน้ารายการ PO เลย
                _log("คลิกเมนูย่อย 'ใบสั่งซื้อ' แล้ว.")

                _log("รอหน้า PO โหลด...")
                time.sleep(5)  # รอให้หน้า PO โหลดอย่างสมบูรณ์
            except TimeoutException:
                _log(f"!!! ไม่พบเมนูย่อย 'ใบสั่งซื้อ' หรือ 'ดูทั้งหมด' หรือไม่สามารถโต้ตอบได้ด้วย XPATH ที่กำหนด !!!")
                # ... (save screenshot, quit driver, return None) ...
                try:
                    driver.save_screenshot(os.path.join(download_path, "po_submenu_not_found.png"))
                except:
                    pass
                if driver: driver.quit()
                return None

            # ตรวจสอบว่ามาถึงหน้า PO ถูกต้องหรือไม่ (เหมือนเดิม)
            # ... (โค้ดตรวจสอบ URL และ Title ของหน้า PO) ...
            current_url_after_po_nav = driver.current_url
            _log(f"URL ปัจจุบันหลังจากคลิกเมนู PO: {current_url_after_po_nav}")
            # (อาจจะต้องตรวจสอบว่า URL มี ?stid=0 หรือพารามิเตอร์ที่บอกว่าเป็นหน้า "ดูทั้งหมด")
            if "/expense/PO" not in current_url_after_po_nav:  # ตรวจสอบ URL พื้นฐาน
                _log(
                    f"!!! ดูเหมือนจะไม่ได้อยู่ที่หน้า PO ที่ถูกต้องหลังคลิกเมนู. URL คือ: {current_url_after_po_nav} !!!")
            else:
                _log("น่าจะอยู่ที่หน้า PO ถูกต้องแล้ว (หลังคลิกเมนู).")

            # --- จุดทดสอบถัดไป: กดปุ่ม "พิมพ์รายงาน" บนหน้า PO ---
            # ... (โค้ดส่วนที่เหลือของการดาวน์โหลดจะตามมาที่นี่) ...

            _log("โปรแกรมจะหยุดที่นี่สำหรับการทดสอบการนำทางผ่านเมนู Hover. ปิดเบราว์เซอร์ใน 10 วินาที...")
            time.sleep(10)

            if driver:
                _log("กำลังปิด WebDriver...")
                driver.quit()
                _log("WebDriver ปิดแล้ว.")
            return None  # สำหรับการทดสอบนี้

        # 2. คลิกรายการเมนู "ใบสั่งซื้อ"
        po_menu_item_xpath = "//a[contains(@href,'/expense/PO') and (normalize-space(.)='ใบสั่งซื้อ' or normalize-space(.)='Purchase Order')]" # <<--- !!! ตรวจสอบ Selector นี้ให้ถูกต้อง !!!
        try:
            _log(f"กำลังค้นหารายการเมนู 'ใบสั่งซื้อ' ด้วย XPATH: {po_menu_item_xpath}")
            po_menu_link = wait.until(EC.element_to_be_clickable((By.XPATH, po_menu_item_xpath)))
            po_menu_link.click()
            _log("คลิกรายการเมนู 'ใบสั่งซื้อ' แล้ว. รอหน้า PO โหลด...")
            time.sleep(5)
        except TimeoutException:
            _log(f"!!! ไม่พบรายการเมนู 'ใบสั่งซื้อ' หรือไม่สามารถคลิกได้ด้วย XPATH: {po_menu_item_xpath} !!!")
            # ... (save screenshot, quit driver, return None) ...
            try: driver.save_screenshot(os.path.join(download_path, "po_menu_item_not_found.png"))
            except: pass
            if driver: driver.quit()
            return None

        # ตรวจสอบว่ามาถึงหน้า PO ถูกต้องหรือไม่
        # ... (โค้ดตรวจสอบ URL และ Title ของหน้า PO เหมือนเดิม) ...
        current_url_after_po_nav = driver.current_url
        _log(f"URL ปัจจุบันหลังจากคลิกเมนู PO: {current_url_after_po_nav}")
        page_title_after_po_nav = driver.title
        _log(f"Page Title หลังจากคลิกเมนู PO: {page_title_after_po_nav}")
        if "PO" not in page_title_after_po_nav.upper() and "ใบสั่งซื้อ" not in page_title_after_po_nav:
             _log(f"!!! ดูเหมือนจะไม่ได้อยู่ที่หน้า PO ที่ถูกต้องหลังคลิกเมนู. Title คือ: {page_title_after_po_nav} !!!")
        else:
            _log("น่าจะอยู่ที่หน้า PO ถูกต้องแล้ว (หลังคลิกเมนู).")


        # --- จุดทดสอบถัดไป: กดปุ่ม "พิมพ์รายงาน" บนหน้า PO ---
        _log("ขั้นตอนต่อไปคือการกดปุ่ม 'พิมพ์รายงาน' บนหน้า PO.")
        _log("โปรแกรมจะหยุดที่นี่สำหรับการทดสอบนี้. ปิดเบราว์เซอร์ใน 10 วินาที...")
        time.sleep(10)
        # --- จบจุดทดสอบ ---


        # (คอมเมนต์โค้ดส่วนที่เหลือของการดาวน์โหลด (ขั้นตอน 3-8 ของคุณ) ออกไปก่อน)
        # # 3. กดปุ่ม "พิมพ์รายงาน" บนหน้า PO
        # # ...

        if driver:
            _log("กำลังปิด WebDriver...")
            driver.quit()
            _log("WebDriver ปิดแล้ว.")
        return None # สำหรับการทดสอบนี้

    # ... (ส่วน except และ finally เหมือนเดิม) ...
    except TimeoutException as te:
        # ... (เหมือนเดิม) ...
        _log(f"!!! TimeoutException: {te} !!!")
        if driver:
            try: driver.save_screenshot(os.path.join(download_path, "peak_po_timeout_error.png"))
            except: pass
            driver.quit()
        return None
    except Exception as e:
        # ... (เหมือนเดิม) ...
        _log(f"!!! Fatal Error: {e} !!!")
        if driver:
            try: driver.save_screenshot(os.path.join(download_path, "peak_po_fatal_error.png"))
            except: pass
            driver.quit()
        return None


# --- ส่วน if __name__ == '__main__': สำหรับทดสอบ ---
if __name__ == '__main__':
    print("="*30)
    print("  Testing auto_downloader.py (PEAK PO - Login & Business Select)  ")
    print("="*30)

    test_peak_user = "sirichai.c@zubbsteel.com"  # <<--- !!! ใส่ข้อมูลจริง !!!
    test_peak_pass = "Zubb*2013" # <<--- !!! ใส่ข้อมูลจริง !!!
    # ชื่อกิจการที่คุณต้องการให้โปรแกรมคลิกเลือก (ต้องตรงกับที่แสดงบนเว็บ)
    test_target_business = "บจกก. บิช ฮีโร่ (สำนักงานใหญ่)" # <<--- !!! แก้ไขเป็นชื่อกิจการของคุณ !!!

    if test_peak_user == "YOUR_PEAK_EMAIL_HERE" or \
       test_peak_pass == "YOUR_PEAK_PASSWORD_HERE" : # เพิ่มเงื่อนไขเช็คชื่อกิจการ
        print("\n!!! กรุณาแก้ไข test_peak_user, test_peak_pass, และ test_target_business ใน auto_downloader.py ก่อนรันทดสอบ !!!\n")
    else:
        current_script_dir = os.getcwd()
        test_save_dir_for_artifacts = os.path.join(current_script_dir, "peak_po_test_artifacts")

        def standalone_logger(message):
            print(f"[StandaloneTestLogger] {message}")

        print(f"Artifacts (screenshots) will be saved to: {test_save_dir_for_artifacts}")

        download_peak_purchase_order_report(
            username=test_peak_user,
            password=test_peak_pass,
            target_business_name_to_select=test_target_business, # ส่งชื่อกิจการไปด้วย
            save_directory=test_save_dir_for_artifacts,
            log_callback=standalone_logger
        )
    print("="*30)
    print("  Finished Testing auto_downloader.py (PEAK PO - Login & Business Select)  ")
    print("="*30)