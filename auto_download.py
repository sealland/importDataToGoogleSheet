import os
import time
import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
import traceback


# <<< START OF THE FINAL, COMPLETE FUNCTION FOR PURCHASE ORDER >>>
def download_peak_purchase_order_report(username, password, target_business_name_to_select,
                                        save_directory, desired_file_name="peak_po_report.xlsx", log_callback=None):
    print("--- [DEBUG] INSIDE download_peak_purchase_order_report FUNCTION (FINAL VERSION) ---")

    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(f"[PEAK_PO_Downloader_LOG] {msg}")

    _log("Function started.")
    download_path = os.path.abspath(save_directory)

    # --- Setup: Create folder and clean old files ---
    if not os.path.exists(download_path):
        os.makedirs(download_path)
        _log(f"สร้างโฟลเดอร์ดาวน์โหลด: {download_path}")

    for f in glob.glob(os.path.join(download_path, "*.xlsx")):
        if os.path.basename(f) == desired_file_name or "purchaseOrder_report_export_" in os.path.basename(f):
            try:
                os.remove(f)
                _log(f"ลบไฟล์ report เก่าที่อาจค้างอยู่: {f}")
            except Exception as e_rm_old:
                _log(f"!!! Warning: ไม่สามารถลบไฟล์เก่า '{f}': {e_rm_old} !!!")

    # --- Setup WebDriver ---
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_path, "download.prompt_for_download": False,
        "download.directory_upgrade": True, "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    driver = None
    try:
        _log("กำลังเริ่ม WebDriver...")
        driver_service = ChromeService(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(service=driver_service, options=chrome_options)

        wait = WebDriverWait(driver, 30)
        long_wait = WebDriverWait(driver, 45)
        short_wait = WebDriverWait(driver, 15)

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 1: Login และ เลือกกิจการ (เหมือนกัน)
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 1: Login และ เลือกกิจการ...")
        driver.get("https://secure.peakaccount.com/login")
        wait.until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(username)
        wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(password)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(.)='เข้าสู่ระบบ PEAK']"))).click()
        long_wait.until(
            EC.any_of(EC.url_contains("selectlist"), EC.presence_of_element_located((By.ID, "mainNavBarBottom"))))

        if "selectlist" in driver.current_url.lower():
            _log("อยู่ที่หน้าเลือกกิจการ...")
            business_item_xpath = f"//p[contains(@class, 'textBold') and normalize-space(.)='{target_business_name_to_select}']"
            long_wait.until(EC.element_to_be_clickable((By.XPATH, business_item_xpath))).click()
        _log("เข้าสู่กิจการสำเร็จ.")
        long_wait.until(EC.presence_of_element_located((By.ID, "mainNavBarBottom")))
        _log("ขั้นตอนที่ 1 เสร็จสิ้น.")

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 2: นำทางไปยังหน้า Purchase Order
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 2: การนำทางไปยังหน้า Purchase Order...")
        actions = ActionChains(driver)
        expense_menu_xpath = "//li[@id='Menu_expense']/descendant::a[contains(normalize-space(.), 'รายจ่าย')][1]"
        po_submenu_to_hover_xpath = "//li[@id='Menu_expense']//div[contains(@class, 'dropdown menu-margin')]//div[@name='selectDropdown'][.//a[normalize-space(.)='ใบสั่งซื้อ']]//a[@class='nameSelect' and normalize-space(.)='ใบสั่งซื้อ']"
        view_all_po_actual_link_xpath = "//li[@id='Menu_expense']//div[@name='selectDropdown'][.//a[normalize-space(.)='ใบสั่งซื้อ']]//div[contains(@class, 'optionDropdown')]//a[@class='nemeOption' and normalize-space(.)='ดูทั้งหมด']"

        actions.move_to_element(wait.until(EC.visibility_of_element_located((By.XPATH, expense_menu_xpath)))).perform()
        time.sleep(1.5)
        actions.move_to_element(
            wait.until(EC.visibility_of_element_located((By.XPATH, po_submenu_to_hover_xpath)))).perform()
        time.sleep(1.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, view_all_po_actual_link_xpath))).click()

        long_wait.until(lambda d: "/expense/po" in d.current_url.lower())
        _log("อยู่ที่หน้า 'ใบสั่งซื้อ' ถูกต้องแล้ว.")
        driver.find_element(By.TAG_NAME, "body").click()  # ปิดเมนูที่ค้าง
        time.sleep(1)
        _log("ขั้นตอนที่ 2 เสร็จสิ้น.")

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 3: สั่งพิมพ์รายงาน
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 3: คลิกปุ่ม 'พิมพ์รายงาน' และจัดการ Pop-up...")
        wait.until(EC.element_to_be_clickable((By.XPATH,
                                               "//div[contains(@class, 'header-section')]//button[contains(normalize-space(.), 'พิมพ์รายงาน')]"))).click()
        modal_xpath = "//div[@id='modalBox' and @showmodal='true']"
        wait.until(EC.visibility_of_element_located((By.XPATH, modal_xpath)))
        _log("Pop-up ปรากฏแล้ว.")
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"{modal_xpath}//label[.//p[normalize-space(.)='แสดงรายละเอียด']]"))).click()
        checkbox_names = ["ใบสั่งซื้อสินทรัพย์", "ข้อมูลราคาและภาษี", "กลุ่มจัดประเภท", "ข้อมูลอื่น", "ประวัติเอกสาร",
                          "เอกสารที่ถูกยกเลิก"]
        for name in checkbox_names:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"{modal_xpath}//label[.//span[normalize-space(.)='{name}']]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH,
                                               f"{modal_xpath}//button[contains(normalize-space(.), 'พิมพ์รายงาน') and not(ancestor::div[contains(@class,'secondary')])]"))).click()
        _log("สั่งพิมพ์รายงานใน Pop-up แล้ว.")

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 4: Polling & Download (Fire and Forget + File Check)
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 4: เริ่มกระบวนการตรวจสอบ Notification...")
        NOTIFICATION_TIMEOUT_SECONDS = 300
        POLLING_INTERVAL_SECONDS = 15
        bell_icon_to_click_xpath = "//div[@id='notification']//a[contains(@class, 'fa-bell')]"
        notification_panel_xpath = "//div[contains(@class, 'dropdownNotification') and contains(@class, 'showNotification')]"
        body_element_xpath = "//body"

        start_time = time.time()
        download_triggered = False

        while time.time() - start_time < NOTIFICATION_TIMEOUT_SECONDS:
            try:
                _log("...กำลังตรวจสอบ Notification...")

                # 1. คลิกกระดิ่ง
                try:
                    bell_icon = short_wait.until(EC.element_to_be_clickable((By.XPATH, bell_icon_to_click_xpath)))
                    driver.execute_script("arguments[0].click();", bell_icon)
                except Exception:
                    driver.find_element(By.XPATH, body_element_xpath).click();
                    time.sleep(1)
                    bell_icon = short_wait.until(EC.element_to_be_clickable((By.XPATH, bell_icon_to_click_xpath)))
                    driver.execute_script("arguments[0].click();", bell_icon)

                # 2. รอ Panel และเนื้อหา
                wait.until(EC.visibility_of_element_located((By.XPATH, notification_panel_xpath)))
                time.sleep(2)

                # 3. ยิง JavaScript (Fire and Forget)
                _log("   กำลังใช้ JavaScript เพื่อ 'พยายาม' คลิกปุ่มดาวน์โหลด...")
                js_script = """
                const items = document.querySelectorAll('.notificationItem');
                for (let i = items.length - 1; i >= 0; i--) {
                    const item = items[i];
                    if (item.querySelector('h3')?.textContent.includes('รายงานใบสั่งซื้อ')) {
                        const btn = item.querySelector('.hyperLinkText');
                        if (btn?.textContent.trim() === 'ดาวน์โหลด') {
                            btn.click();
                            break; 
                        }
                    }
                }
                """
                driver.execute_script(js_script)

                # 4. ตรวจสอบผลลัพธ์ด้วยไฟล์
                _log("   ยิง Script เสร็จสิ้น. จะเริ่มตรวจสอบไฟล์ในโฟลเดอร์...")
                FILE_CHECK_TIMEOUT = 10
                file_check_start_time = time.time()
                download_started = False
                while time.time() - file_check_start_time < FILE_CHECK_TIMEOUT:
                    if glob.glob(os.path.join(download_path, "*.crdownload")) or glob.glob(
                            os.path.join(download_path, "purchaseOrder_report_export_*.xlsx")):
                        _log("   ตรวจพบไฟล์ใหม่! การดาวน์โหลดเริ่มต้นแล้ว!")
                        download_started = True
                        break
                    time.sleep(1)

                if download_started:
                    download_triggered = True
                    break

                _log("   ไม่พบไฟล์ใหม่ในรอบนี้. จะลองในรอบถัดไป")

            except Exception as e_poll:
                _log(f"!!! Error ระหว่างการ Polling: {e_poll} !!!")

            if not download_triggered:
                try:
                    driver.find_element(By.XPATH, body_element_xpath).click()
                except:
                    pass
                time.sleep(POLLING_INTERVAL_SECONDS - (FILE_CHECK_TIMEOUT + 4))

        if not download_triggered:
            _log(f"!!! ล้มเหลว: หมดเวลา {NOTIFICATION_TIMEOUT_SECONDS} วินาทีแล้ว แต่ยังไม่พบรายงานให้ดาวน์โหลด !!!")
            return None

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 5: รอไฟล์ดาวน์โหลดให้เสร็จสมบูรณ์
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 5: รอไฟล์ดาวน์โหลดให้เสร็จสมบูรณ์...")
        DOWNLOAD_WAIT_TIMEOUT = 120
        wait_start_time = time.time()
        final_filepath = None
        while time.time() - wait_start_time < DOWNLOAD_WAIT_TIMEOUT:
            if not glob.glob(os.path.join(download_path, "*.crdownload")):
                # ** แก้ไขชื่อไฟล์ที่ค้นหาให้ถูกต้อง **
                xlsx_files = glob.glob(os.path.join(download_path, "purchaseOrder_report_export_*.xlsx"))
                if xlsx_files:
                    downloaded_file = xlsx_files[0]
                    _log(f"ตรวจพบไฟล์ที่ดาวน์โหลดเสร็จแล้ว: {downloaded_file}")
                    final_filepath_target = os.path.join(download_path, desired_file_name)
                    os.rename(downloaded_file, final_filepath_target)
                    _log(f"เปลี่ยนชื่อไฟล์เป็น: {final_filepath_target}")
                    final_filepath = final_filepath_target
                    break
            time.sleep(1)

        if not final_filepath:
            _log(f"!!! ล้มเหลว: หมดเวลารอไฟล์ดาวน์โหลด ({DOWNLOAD_WAIT_TIMEOUT} วินาที) !!!")
            return None

        _log("🎉🎉🎉 ดาวน์โหลดและเปลี่ยนชื่อไฟล์รายงานใบสั่งซื้อสำเร็จ! 🎉🎉🎉")
        return final_filepath

    except Exception as e:
        _log(f"!!! Fatal Error (Overall Function Level): {e} !!!")
        _log(traceback.format_exc())
        return None
    finally:
        if driver:
            driver.quit()
        _log("Function finished.")


# <<< END OF THE FINAL, COMPLETE FUNCTION FOR PURCHASE ORDER >>>

# <<< START OF THE FINAL, COMPLETE FUNCTION FOR QUOTATION >>>
def download_peak_quotation_report(username, password, target_business_name_to_select,
                                   save_directory, desired_file_name="peak_quotation_report.xlsx", log_callback=None):
    print("--- [DEBUG] INSIDE download_peak_quotation_report FUNCTION (FINAL VERSION) ---")

    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(f"[PEAK_QT_Downloader_LOG] {msg}")

    _log("Function started.")
    download_path = os.path.abspath(save_directory)

    # --- Setup: Create folder and clean old files ---
    if not os.path.exists(download_path):
        os.makedirs(download_path)
        _log(f"สร้างโฟลเดอร์ดาวน์โหลด: {download_path}")

    for f in glob.glob(os.path.join(download_path, "*.xlsx")):
        if os.path.basename(f) == desired_file_name or "quotation_report_export_" in os.path.basename(f):
            try:
                os.remove(f)
                _log(f"ลบไฟล์ report เก่าที่อาจค้างอยู่: {f}")
            except Exception as e_rm_old:
                _log(f"!!! Warning: ไม่สามารถลบไฟล์เก่า '{f}': {e_rm_old} !!!")

    # --- Setup WebDriver ---
    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_path, "download.prompt_for_download": False,
        "download.directory_upgrade": True, "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    driver = None
    try:
        _log("กำลังเริ่ม WebDriver...")
        driver_service = ChromeService(executable_path=ChromeDriverManager().install())
        driver = webdriver.Chrome(service=driver_service, options=chrome_options)

        wait = WebDriverWait(driver, 30)
        long_wait = WebDriverWait(driver, 45)
        short_wait = WebDriverWait(driver, 15)

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 1: Login และ เลือกกิจการ
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 1: Login และ เลือกกิจการ...")
        driver.get("https://secure.peakaccount.com/login")
        wait.until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(username)
        wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(password)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(.)='เข้าสู่ระบบ PEAK']"))).click()
        long_wait.until(
            EC.any_of(EC.url_contains("selectlist"), EC.presence_of_element_located((By.ID, "mainNavBarBottom"))))

        if "selectlist" in driver.current_url.lower():
            _log("อยู่ที่หน้าเลือกกิจการ...")
            business_item_xpath = f"//p[contains(@class, 'textBold') and normalize-space(.)='{target_business_name_to_select}']"
            long_wait.until(EC.element_to_be_clickable((By.XPATH, business_item_xpath))).click()
        _log("เข้าสู่กิจการสำเร็จ.")
        long_wait.until(EC.presence_of_element_located((By.ID, "mainNavBarBottom")))
        _log("ขั้นตอนที่ 1 เสร็จสิ้น.")

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 2: นำทางไปยังหน้า Quotation
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 2: การนำทางไปยังหน้า Quotation...")
        actions = ActionChains(driver)
        income_menu_xpath = "//li[@id='Menu_income']/descendant::a[contains(normalize-space(.), 'รายรับ')][1]"
        quotation_submenu_to_hover_xpath = "//li[@id='Menu_income']//div[contains(@class, 'dropdown menu-margin')]//a[normalize-space(.)='ใบเสนอราคา']"
        view_all_quotation_link_xpath = "//li[@id='Menu_income']//div[.//a[normalize-space(.)='ใบเสนอราคา']]//a[normalize-space(.)='ดูทั้งหมด']"

        actions.move_to_element(wait.until(EC.visibility_of_element_located((By.XPATH, income_menu_xpath)))).perform()
        time.sleep(1.5)
        actions.move_to_element(
            wait.until(EC.visibility_of_element_located((By.XPATH, quotation_submenu_to_hover_xpath)))).perform()
        time.sleep(1.5)
        wait.until(EC.element_to_be_clickable((By.XPATH, view_all_quotation_link_xpath))).click()

        long_wait.until(lambda d: "income/quotation" in d.current_url.lower())
        _log("อยู่ที่หน้า 'ใบเสนอราคา' ถูกต้องแล้ว.")
        driver.find_element(By.TAG_NAME, "body").click()  # ปิดเมนูที่ค้าง
        time.sleep(1)
        _log("ขั้นตอนที่ 2 เสร็จสิ้น.")

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 3: สั่งพิมพ์รายงาน
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 3: คลิกปุ่ม 'พิมพ์รายงาน' และจัดการ Pop-up...")
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(normalize-space(.), 'พิมพ์รายงาน')]"))).click()
        modal_xpath = "//div[@id='modalBox' and @showmodal='true']"
        wait.until(EC.visibility_of_element_located((By.XPATH, modal_xpath)))
        _log("Pop-up ปรากฏแล้ว.")
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"{modal_xpath}//label[.//p[normalize-space(.)='แสดงรายละเอียด']]"))).click()
        checkbox_names = ["ข้อมูลราคาและภาษี", "กลุ่มจัดประเภท", "ข้อมูลอื่น", "ประวัติเอกสาร", "เอกสารที่ถูกยกเลิก"]
        for name in checkbox_names:
            wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"{modal_xpath}//label[.//span[normalize-space(.)='{name}']]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH,
                                               f"{modal_xpath}//button[contains(normalize-space(.), 'พิมพ์รายงาน') and not(ancestor::div[contains(@class,'secondary')])]"))).click()
        _log("สั่งพิมพ์รายงานใน Pop-up แล้ว.")

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 4: Polling & Download (Fire and Forget + File Check)
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 4: เริ่มกระบวนการตรวจสอบ Notification...")
        NOTIFICATION_TIMEOUT_SECONDS = 300
        POLLING_INTERVAL_SECONDS = 15
        bell_icon_to_click_xpath = "//div[@id='notification']//a[contains(@class, 'fa-bell')]"
        notification_panel_xpath = "//div[contains(@class, 'dropdownNotification') and contains(@class, 'showNotification')]"
        body_element_xpath = "//body"

        start_time = time.time()
        download_triggered = False

        while time.time() - start_time < NOTIFICATION_TIMEOUT_SECONDS:
            try:
                _log("...กำลังตรวจสอบ Notification...")

                # 1. คลิกกระดิ่ง
                try:
                    bell_icon = short_wait.until(EC.element_to_be_clickable((By.XPATH, bell_icon_to_click_xpath)))
                    driver.execute_script("arguments[0].click();", bell_icon)
                except Exception:
                    driver.find_element(By.XPATH, body_element_xpath).click();
                    time.sleep(1)
                    bell_icon = short_wait.until(EC.element_to_be_clickable((By.XPATH, bell_icon_to_click_xpath)))
                    driver.execute_script("arguments[0].click();", bell_icon)

                # 2. รอ Panel และเนื้อหา
                wait.until(EC.visibility_of_element_located((By.XPATH, notification_panel_xpath)))
                time.sleep(2)

                # 3. ยิง JavaScript (Fire and Forget)
                _log("   กำลังใช้ JavaScript เพื่อ 'พยายาม' คลิกปุ่มดาวน์โหลด...")
                js_script = """
                const items = document.querySelectorAll('.notificationItem');
                for (let i = items.length - 1; i >= 0; i--) {
                    const item = items[i];
                    if (item.querySelector('h3')?.textContent.includes('รายงานใบเสนอราคา')) {
                        const btn = item.querySelector('.hyperLinkText');
                        if (btn?.textContent.trim() === 'ดาวน์โหลด') {
                            btn.click();
                            break; 
                        }
                    }
                }
                """
                driver.execute_script(js_script)

                # 4. ตรวจสอบผลลัพธ์ด้วยไฟล์
                _log("   ยิง Script เสร็จสิ้น. จะเริ่มตรวจสอบไฟล์ในโฟลเดอร์...")
                FILE_CHECK_TIMEOUT = 10
                file_check_start_time = time.time()
                download_started = False
                while time.time() - file_check_start_time < FILE_CHECK_TIMEOUT:
                    if glob.glob(os.path.join(download_path, "*.crdownload")) or glob.glob(
                            os.path.join(download_path, "quotation_report_export_*.xlsx")):
                        _log("   ตรวจพบไฟล์ใหม่! การดาวน์โหลดเริ่มต้นแล้ว!")
                        download_started = True
                        break
                    time.sleep(1)

                if download_started:
                    download_triggered = True
                    break

                _log("   ไม่พบไฟล์ใหม่ในรอบนี้. จะลองในรอบถัดไป")

            except Exception as e_poll:
                _log(f"!!! Error ระหว่างการ Polling: {e_poll} !!!")

            if not download_triggered:
                try:
                    driver.find_element(By.XPATH, body_element_xpath).click()
                except:
                    pass
                time.sleep(POLLING_INTERVAL_SECONDS - (FILE_CHECK_TIMEOUT + 4))

        if not download_triggered:
            _log(f"!!! ล้มเหลว: หมดเวลา {NOTIFICATION_TIMEOUT_SECONDS} วินาทีแล้ว แต่ยังไม่พบรายงานให้ดาวน์โหลด !!!")
            return None

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 5: รอไฟล์ดาวน์โหลดให้เสร็จสมบูรณ์
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 5: รอไฟล์ดาวน์โหลดให้เสร็จสมบูรณ์...")
        DOWNLOAD_WAIT_TIMEOUT = 120
        wait_start_time = time.time()
        final_filepath = None
        while time.time() - wait_start_time < DOWNLOAD_WAIT_TIMEOUT:
            if not glob.glob(os.path.join(download_path, "*.crdownload")):
                xlsx_files = glob.glob(os.path.join(download_path, "quotation_report_export_*.xlsx"))
                if xlsx_files:
                    downloaded_file = xlsx_files[0]
                    _log(f"ตรวจพบไฟล์ที่ดาวน์โหลดเสร็จแล้ว: {downloaded_file}")
                    final_filepath_target = os.path.join(download_path, desired_file_name)
                    os.rename(downloaded_file, final_filepath_target)
                    _log(f"เปลี่ยนชื่อไฟล์เป็น: {final_filepath_target}")
                    final_filepath = final_filepath_target
                    break
            time.sleep(1)

        if not final_filepath:
            _log(f"!!! ล้มเหลว: หมดเวลารอไฟล์ดาวน์โหลด ({DOWNLOAD_WAIT_TIMEOUT} วินาที) !!!")
            return None

        _log("🎉🎉🎉 ดาวน์โหลดและเปลี่ยนชื่อไฟล์รายงานใบเสนอราคาสำเร็จ! 🎉🎉🎉")
        return final_filepath

    except Exception as e:
        _log(f"!!! Fatal Error (Overall Function Level): {e} !!!")
        _log(traceback.format_exc())
        return None
    finally:
        if driver:
            driver.quit()
        _log("Function finished.")


# <<< END OF THE FINAL, COMPLETE FUNCTION FOR QUOTATION >>>

if __name__ == '__main__':
    # --- ส่วนกำหนดค่า (เหมือนเดิม) ---
    print("=" * 40)
    print("  STARTING PEAK AUTOMATION TEST SUITE  ")
    print("=" * 40)

    test_peak_user = "sirichai.c@zubbsteel.com"
    test_peak_pass = "Zubb*2013"  # หรือรหัสผ่านของคุณ
    test_target_business = "บจ. บิซ ฮีโร่ (สำนักงานใหญ่)"
    current_script_dir = os.getcwd()


    def standalone_logger(message):
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}][StandaloneTestLogger] {message}")


    # --- 1. ทดสอบฟังก์ชันดาวน์โหลดใบเสนอราคา (Quotation) ---
    print("\n" + "=" * 10 + " TESTING QUOTATION REPORT " + "=" * 10)
    test_save_dir_for_qt = os.path.join(current_script_dir, "peak_qt_test_artifacts")

    qt_result_path = download_peak_quotation_report(
        username=test_peak_user,
        password=test_peak_pass,
        target_business_name_to_select=test_target_business,
        save_directory=test_save_dir_for_qt,
        log_callback=standalone_logger
    )

    if qt_result_path:
        print(f"\n[SUCCESS] Quotation download complete! File at: {qt_result_path}")
    else:
        print(f"\n[FAILURE] Quotation download failed. Please check logs.")
    print("=" * 40)

    # --- 2. ทดสอบฟังก์ชันดาวน์โหลดใบสั่งซื้อ (Purchase Order) ---
    print("\n" + "=" * 10 + " TESTING PURCHASE ORDER REPORT " + "=" * 10)
    test_save_dir_for_po = os.path.join(current_script_dir, "peak_po_test_artifacts")

    po_result_path = download_peak_purchase_order_report(
        username=test_peak_user,
        password=test_peak_pass,
        target_business_name_to_select=test_target_business,
        save_directory=test_save_dir_for_po,
        log_callback=standalone_logger
    )

    if po_result_path:
        print(f"\n[SUCCESS] Purchase Order download complete! File at: {po_result_path}")
    else:
        print(f"\n[FAILURE] Purchase Order download failed. Please check logs.")
    print("=" * 40)

    print("\n" + "=" * 40)
    print("  PEAK AUTOMATION TEST SUITE FINISHED  ")
    print("=" * 40)