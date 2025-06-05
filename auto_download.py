# auto_downloader.py
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains


def download_peak_purchase_order_report(username, password, target_business_name_to_select,
                                        save_directory, desired_file_name="peak_po_report.xlsx", log_callback=None):
    print("--- [DEBUG] INSIDE download_peak_purchase_order_report FUNCTION ---")

    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(f"[PEAK_PO_Downloader_LOG] {msg}")

    _log("Function started.")
    download_path = os.path.abspath(save_directory)
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
    # chrome_options.add_argument("--headless") # ลองเปิดถ้าต้องการให้ทำงานเบื้องหลัง (อาจต้องปรับ XPath บางจุด)
    # chrome_options.add_argument("--disable-gpu") # Often used with headless
    # chrome_options.add_argument("--window-size=1920,1080") # Specify window size for headless
    _log("Chrome options configured.")

    driver = None
    try:
        _log("กำลังเริ่ม WebDriver สำหรับ PEAK...")
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

        wait = WebDriverWait(driver, 30) # Default wait
        long_wait = WebDriverWait(driver, 45) # For slower loading pages
        short_wait = WebDriverWait(driver, 15) # For faster elements

        # 1. Login
        login_url = "https://secure.peakaccount.com/login"
        _log(f"กำลังไปที่หน้า Login: {login_url}")
        driver.get(login_url)
        wait.until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(username)
        wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(password)
        login_button_xpath = "//button[normalize-space(.)='เข้าสู่ระบบ PEAK']"
        wait.until(EC.element_to_be_clickable((By.XPATH, login_button_xpath))).click()
        _log("คลิกปุ่ม Login แล้ว.")

        _log("รอ URL เปลี่ยนหลังจาก Login และหน้าพร้อม...")
        try:
            # รอจน URL ไม่มีคำว่า "login" และ element ที่บ่งบอกว่า login เสร็จแล้วปรากฏ (เช่น ปุ่ม logout หรือ ชื่อ user)
            long_wait.until(
                EC.all_of(  # รอเงื่อนไขทั้งหมดเป็นจริง
                    EC.url_contains("peakaccount.com/"),  # URL หลักของ peak
                    EC.none_of(EC.url_contains("login"))  # ไม่มีคำว่า login ใน URL
                )
            )
            # เพิ่มการรอ element ที่เฉพาะเจาะจงของหน้าหลัง login (ถ้ามี)
            # เช่น รอให้ div ที่มี id 'selectcompany' (ถ้ายังอยู่หน้า selectlist) หรือ ชื่อ user ปรากฏ
            long_wait.until(
                EC.any_of(  # รอเงื่อนไขใดเงื่อนไขหนึ่งเป็นจริง
                    EC.presence_of_element_located(
                        (By.XPATH, "//div[contains(@class, 'list')]//p[contains(@class, 'crop')]")),
                    # Indicator ของหน้า selectlist
                    EC.presence_of_element_located((By.ID, "mainNavBarBottom"))  # Indicator ของหน้า dashboard
                )
            )

            current_url_after_login = driver.current_url
            _log(f"URL ปัจจุบันหลังจากพยายาม Login: {current_url_after_login}")

        except TimeoutException:
            _log("!!! Timeout: ดูเหมือน Login ไม่สำเร็จ หรือหน้าหลัง Login ไม่โหลดสมบูรณ์ !!!")
            try:
                driver.save_screenshot(os.path.join(download_path, "po_login_failed_or_page_not_ready.png"))
            except:
                pass
            if driver: driver.quit()
            return None
        else:
            _log("Login ดูเหมือนจะสำเร็จแล้ว และหน้าหลัง Login เริ่มโหลด.")

        # 1.5: เลือกกิจการ
        # ตรวจสอบเสมอว่าตอนนี้อยู่ที่หน้า selectlist หรือไม่
        if "selectlist" in driver.current_url.lower():
            _log("อยู่ที่หน้าเลือกกิจการ กำลังดำเนินการเลือก...")
            try:
                first_business_in_list_xpath = "//div[contains(@class, 'list')]//p[contains(@class, 'crop') and contains(@class, 'textBold')]"
                _log(f"กำลังรอรายการกิจการแรกปรากฏ (ใช้ XPATH: {first_business_in_list_xpath} สำหรับเช็ค)...")
                # ใช้ long_wait เพราะบางทีกิจการโหลดช้า
                long_wait.until(EC.visibility_of_element_located((By.XPATH, first_business_in_list_xpath)))
                _log("หน้ารายการกิจการปรากฏแล้ว.")

                _log(f"กำลังค้นหากิจการ '{target_business_name_to_select}' ในรายการ...")
                business_item_xpath = f"//div[contains(@class, 'list')]//div[contains(@class, 'col2')]/p[contains(@class, 'textBold') and normalize-space(.)='{target_business_name_to_select}']"
                _log(f"XPath ที่ใช้ค้นหากิจการ (p element): {business_item_xpath}")

                business_element = long_wait.until(EC.element_to_be_clickable((By.XPATH, business_item_xpath)))
                _log(f"พบกิจการ '{target_business_name_to_select}' (p element). กำลังคลิก...")
                business_element.click()
                _log("คลิกเลือกกิจการแล้ว. รอหน้า Dashboard/หน้าหลักของกิจการโหลด...")

                # **** จุดตรวจสอบสำคัญหลังคลิกเลือกกิจการ ****
                _log("รอ URL เปลี่ยนหลังจากเลือกกิจการ และรอ main navigation bar ปรากฏ...")
                try:
                    long_wait.until(
                        EC.all_of(
                            EC.none_of(EC.url_contains("selectlist")),  # ต้องไม่อยู่ใน selectlist แล้ว
                            EC.presence_of_element_located((By.ID, "mainNavBarBottom"))  # แถบเมนูหลักต้องปรากฏ
                        )
                    )
                    current_url_after_select = driver.current_url
                    page_title_after_select = driver.title
                    _log(f"URL หลังจากเลือกกิจการสำเร็จ: {current_url_after_select}")
                    _log(f"Page Title หลังจากเลือกกิจการสำเร็จ: {page_title_after_select}")
                    if "selectlist" in current_url_after_select.lower():  # เช็คซ้ำเผื่อกรณีแปลกๆ
                        _log(f"!!! วิกฤต: ยังคงอยู่ที่หน้า selectlist แม้ควรจะออกจากหน้านั้นแล้ว !!!")
                        driver.save_screenshot(
                            os.path.join(download_path, "business_selection_STUCK_ON_SELECTLIST.png"))
                        if driver: driver.quit(); return None
                    _log(f"เข้าสู่กิจการ '{target_business_name_to_select}' สำเร็จแล้ว.")

                except TimeoutException:
                    current_url_at_fail = driver.current_url
                    _log(
                        f"!!! Timeout: ไม่สามารถยืนยันการออกจากหน้า selectlist หรือหน้า Dashboard ไม่โหลดสมบูรณ์หลังเลือกกิจการ. URL ปัจจุบัน: {current_url_at_fail} !!!")
                    driver.save_screenshot(os.path.join(download_path, "business_selection_timeout_critical.png"))
                    if driver: driver.quit()
                    return None

            except TimeoutException:
                _log(
                    f"!!! TimeoutException: ไม่พบกิจการ '{target_business_name_to_select}' ในหน้า selectlist หรือหน้า selectlist มีปัญหา !!!")
                driver.save_screenshot(os.path.join(download_path, "business_selection_not_found_or_page_issue.png"))
                if driver: driver.quit()
                return None
        elif EC.presence_of_element_located((By.ID, "mainNavBarBottom"))(
                driver):  # ตรวจสอบว่าถ้าไม่ได้อยู่ selectlist แล้วอยู่หน้า dashboard เลยไหม
            _log(
                "ไม่ได้อยู่ที่หน้าเลือกกิจการ และดูเหมือนจะอยู่ที่หน้า Dashboard (พบ mainNavBarBottom) อาจจะเข้ากิจการ default โดยอัตโนมัติแล้ว.")
            # อาจจะตรวจสอบเพิ่มเติมว่ากิจการที่เข้า default คือกิจการที่ต้องการหรือไม่ (ถ้าจำเป็น)
            # current_company_name_on_header = driver.find_element(By.XPATH, "//p[contains(@class, 'merchantName')]").text
            # if target_business_name_to_select not in current_company_name_on_header:
            # _log(f"!!! คำเตือน: เข้ากิจการ default '{current_company_name_on_header}' ซึ่งไม่ใช่ '{target_business_name_to_select}' !!!")
            # # อาจจะเลือก fail หรือดำเนินการต่อด้วยความระมัดระวัง
        else:
            _log("!!! สถานะไม่คาดคิด: ไม่ได้อยู่ที่หน้า selectlist และไม่สามารถยืนยันได้ว่าอยู่ที่หน้า Dashboard. !!!")
            _log(f"URL ปัจจุบัน: {driver.current_url}")
            driver.save_screenshot(os.path.join(download_path, "unknown_state_after_login.png"))
            if driver: driver.quit()
            return None

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 2: การนำทางไปยังหน้า Purchase Order (ดูทั้งหมด) ผ่านเมนู Hover
        # --------------------------------------------------------------------
        _log("ขั้นตอน: การนำทางไปยังหน้า Purchase Order (ดูทั้งหมด) ผ่านเมนู Hover...")
        actions = ActionChains(driver)

        # XPath ที่ปรับปรุงแล้ว
        expense_menu_xpath = "//li[@id='Menu_expense']/descendant::a[contains(normalize-space(.), 'รายจ่าย')][1]"
        # XPath ที่ปรับปรุงสำหรับ "ใบสั่งซื้อ" (สำหรับ Hover) - เลือกอันที่คิดว่าเสถียรที่สุด
        po_submenu_to_hover_xpath = "//li[@id='Menu_expense']//div[contains(@class, 'dropdown menu-margin')]//div[@name='selectDropdown'][.//a[normalize-space(.)='ใบสั่งซื้อ']]//a[@class='nameSelect' and normalize-space(.)='ใบสั่งซื้อ']"
        # XPath ที่ปรับปรุงสำหรับ "ดูทั้งหมด" (สำหรับ Click)
        view_all_po_actual_link_xpath = "//li[@id='Menu_expense']//div[@name='selectDropdown'][.//a[normalize-space(.)='ใบสั่งซื้อ']]//div[contains(@class, 'optionDropdown')]//a[@class='nemeOption' and normalize-space(.)='ดูทั้งหมด']"

        try:
            _log(f"กำลังรอเมนู 'รายจ่าย' (XPATH: {expense_menu_xpath})")
            expense_menu_element = wait.until(EC.visibility_of_element_located((By.XPATH, expense_menu_xpath)))
            _log("พบเมนู 'รายจ่าย'. กำลัง Hover...")
            actions.move_to_element(expense_menu_element).perform()
            # time.sleep(0.5) # ให้เวลา dropdown แสดงเล็กน้อย (อาจจะไม่จำเป็นถ้า wait ต่อไปดีพอ)

            _log(f"กำลังรอเมนูย่อย 'ใบสั่งซื้อ' (XPATH: {po_submenu_to_hover_xpath})")
            # *** สำคัญ: รอให้ element ของ "ใบสั่งซื้อ" ปรากฏและ clickable ก่อน hover ***
            po_submenu_element_to_hover = wait.until(
                EC.element_to_be_clickable((By.XPATH, po_submenu_to_hover_xpath))
            )
            _log("พบเมนูย่อย 'ใบสั่งซื้อ' (สำหรับ Hover). กำลัง Hover...")
            actions.move_to_element(po_submenu_element_to_hover).perform()
            # time.sleep(0.5) # ให้เวลา sub-dropdown แสดงเล็กน้อย (อาจจะไม่จำเป็นถ้า wait ต่อไปดีพอ)

            _log(f"กำลังรอลิงก์ 'ดูทั้งหมด' (XPATH: {view_all_po_actual_link_xpath})")
            view_all_link_element = wait.until(EC.element_to_be_clickable((By.XPATH, view_all_po_actual_link_xpath)))
            _log("พบลิงก์ 'ดูทั้งหมด'. กำลังคลิก...")
            # view_all_link_element.click() # ลอง click ปกติก่อน
            # ถ้า click ปกติไม่ทำงาน ลอง JavaScript click (แต่ควรเป็นทางเลือกสุดท้าย)
            driver.execute_script("arguments[0].click();", view_all_link_element)
            _log("คลิก 'ดูทั้งหมด' แล้ว (ด้วย JavaScript). รอหน้า PO โหลด...")

            # รอให้ URL หรือ Title เปลี่ยนไปเป็นของหน้า PO
            _log("รอการนำทางไปยังหน้า PO...")
            try:
                long_wait.until(
                    lambda d: ("/expense/PO" in d.current_url.lower() and \
                              ("PO" in d.title.upper() or "ใบสั่งซื้อ" in d.title)) or \
                              ("ใบสั่งซื้อ" in d.title and not "/income" in d.current_url.lower()) # เพิ่มเงื่อนไขเผื่อ title เปลี่ยนก่อน url
                )
                current_url_after_po_nav = driver.current_url
                page_title_after_po_nav = driver.title
                _log(f"URL ปัจจุบันหลังจากนำทางไปหน้า PO: {current_url_after_po_nav}")
                _log(f"Page Title หลังจากนำทางไปหน้า PO: {page_title_after_po_nav}")

                if "/expense/PO" in current_url_after_po_nav and \
                   ("PO" in page_title_after_po_nav.upper() or "ใบสั่งซื้อ" in page_title_after_po_nav):
                    _log("น่าจะอยู่ที่หน้า PO 'ดูทั้งหมด' ถูกต้องแล้ว.")
                    _log("--- นี่คือจุดที่เราจะเริ่มหาปุ่ม 'พิมพ์รายงาน' ---")
                    _log("โปรแกรมจะหยุดที่นี่เพื่อให้คุณตรวจสอบว่ามาถึงหน้า 'ใบสั่งซื้อ - ดูทั้งหมด' ถูกต้องหรือไม่.")
                    _log("ถ้าถูกต้องแล้ว, ขั้นตอนต่อไปคือการหา XPath ของปุ่ม 'พิมพ์รายงาน' หรือ 'Export'.")
                    _log("ปิดเบราว์เซอร์ใน 20 วินาที...")
                    time.sleep(20) # Placeholder for next steps
                    # ถ้ามาถึงตรงนี้ได้จริง ก็ถือว่าสำเร็จในการนำทาง
                    if driver: driver.quit()
                    return "NAV_SUCCESSFUL_TO_PO_LIST" # เปลี่ยน result string
                else:
                    _log(f"!!! ดูเหมือนจะไม่ได้อยู่ที่หน้า PO ที่ถูกต้องหลังคลิก 'ดูทั้งหมด'. URL: {current_url_after_po_nav}, Title: {page_title_after_po_nav} !!!")
                    driver.save_screenshot(os.path.join(download_path, "po_nav_menu_failed_target_page.png"))
                    if driver: driver.quit(); return None

            except TimeoutException:
                _log(f"!!! Timeout: ไม่สามารถยืนยันการไปถึงหน้า PO ที่ถูกต้องหลังคลิก 'ดูทั้งหมด'. URL ปัจจุบัน: {driver.current_url}, Title: {driver.title} !!!")
                driver.save_screenshot(os.path.join(download_path, "po_nav_menu_timeout_target_page.png"))
                if driver: driver.quit(); return None

        except TimeoutException as e_nav:
            _log(f"!!! TimeoutException ระหว่างการนำทางผ่านเมนู (ขั้นตอนค้นหา element): {e_nav} !!!")
            _log(f"ตรวจสอบ XPath: expense='{expense_menu_xpath}', po_hover='{po_submenu_to_hover_xpath}', view_all='{view_all_po_actual_link_xpath}'")
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_menu_nav_timeout_element.png"))
                body_html = driver.find_element(By.TAG_NAME, "body").get_attribute("outerHTML")
                with open(os.path.join(download_path, "page_source_at_menu_timeout.html"), "w", encoding="utf-8") as f:
                    f.write(body_html)
                _log("ภาพหน้าจอและ HTML source ถูกบันทึก (ถ้าสำเร็จ).")
            except Exception as e_diag:
                _log(f"ไม่สามารถบันทึก diagnostics ได้: {e_diag}")
            if driver: driver.quit()
            return None

        # ส่วนนี้จะยังไม่ถึง ถ้าข้างบนสำเร็จแล้วมีการ return หรือ quit
        if driver:
            _log("กำลังปิด WebDriver (สิ้นสุดการทดสอบนำทาง)...")
            driver.quit()
        return "UNEXPECTED_END_OF_NAVIGATION_LOGIC"

    except TimeoutException as te:
        _log(f"!!! TimeoutException (Overall): {te} !!!")
        if driver:
            try: driver.save_screenshot(os.path.join(download_path, "peak_po_timeout_error.png"))
            except: pass
            driver.quit()
        return None
    except Exception as e:
        _log(f"!!! Fatal Error (Overall): {e} !!!")
        import traceback
        _log(traceback.format_exc())
        if driver:
            try: driver.save_screenshot(os.path.join(download_path, "peak_po_fatal_error.png"))
            except: pass
            driver.quit()
        return None
    finally:
        if driver:
            try:
                _log("Ensuring WebDriver is quit in finally block.")
                driver.quit()
            except:
                _log("WebDriver already quit or error during quit in finally.")


# --- ส่วน if __name__ == '__main__': สำหรับทดสอบ ---
if __name__ == '__main__':
    print("=" * 30)
    print("  Testing auto_downloader.py (PEAK PO - Nav to PO All - Attempt 3)  ") # Update test name
    print("=" * 30)

    test_peak_user = "sirichai.c@zubbsteel.com" # คงเดิม
    test_peak_pass = "Zubb*2013" # คงเดิม
    test_target_business = "บจ. บิซ ฮีโร่ (สำนักงานใหญ่)" # คงเดิม

    if "YOUR_PEAK_EMAIL_HERE" in test_peak_user or \
            "YOUR_PEAK_PASSWORD_HERE" in test_peak_pass or \
            "ชื่อกิจการของคุณที่นี่" in test_target_business:
        print("\n!!! กรุณาแก้ไข test_peak_user, test_peak_pass, และ test_target_business ก่อนรันทดสอบ !!!\n")
    else:
        current_script_dir = os.getcwd()
        test_save_dir_for_artifacts = os.path.join(current_script_dir, "peak_po_nav_test_artifacts_v3") # Update artifact dir

        def standalone_logger(message):
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{timestamp}][StandaloneTestLogger] {message}")

        print(f"Username: {test_peak_user}")
        print(f"Target Business: {test_target_business}")
        print(f"Artifacts will be saved to: {test_save_dir_for_artifacts}")

        result = download_peak_purchase_order_report(
            username=test_peak_user,
            password=test_peak_pass,
            target_business_name_to_select=test_target_business,
            save_directory=test_save_dir_for_artifacts,
            log_callback=standalone_logger
        )
        if result == "NAV_SUCCESSFUL_TO_PO_LIST": # Update expected result
            print("\n[StandaloneTestLogger] การนำทางไปยังหน้า PO 'ดูทั้งหมด' ดูเหมือนจะสำเร็จแล้ว!")
        else:
            print(f"\n[StandaloneTestLogger] การนำทางไปยังหน้า PO 'ดูทั้งหมด' ไม่สำเร็จ หรือมีข้อผิดพลาด. Result: {result}")

    print("=" * 30)
    print("  Finished Testing auto_downloader.py (PEAK PO - Nav to PO All - Attempt 3)  ") # Update test name
    print("=" * 30)