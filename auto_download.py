# auto_downloader.py
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
import glob  # สำหรับค้นหาไฟล์ที่ดาวน์โหลด


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

    # ลบไฟล์เก่าที่อาจจะชื่อซ้ำกันในโฟลเดอร์ดาวน์โหลด เพื่อให้แน่ใจว่าไฟล์ที่ได้เป็นไฟล์ใหม่จริงๆ
    # (สำคัญมากถ้าสคริปต์เคยรันแล้วล้มเหลวก่อนจะเปลี่ยนชื่อไฟล์)
    potential_old_file = os.path.join(download_path, desired_file_name)
    if os.path.exists(potential_old_file):
        try:
            os.remove(potential_old_file)
            _log(f"ลบไฟล์เก่า '{potential_old_file}' ที่อาจค้างอยู่แล้ว")
        except Exception as e_rm_old:
            _log(f"!!! Warning: ไม่สามารถลบไฟล์เก่า '{potential_old_file}': {e_rm_old} !!!")
            # อาจจะตัดสินใจ return None ถ้าการมีไฟล์เก่าเป็นปัญหาใหญ่

    # เราจะค้นหาไฟล์ที่ดาวน์โหลดมาโดยอิงจากชื่อที่ PEAK ตั้งให้ก่อน แล้วค่อยเปลี่ยนชื่อ
    # ดังนั้น เราต้องลบไฟล์ที่มีชื่อคล้ายๆ กันที่ PEAK อาจจะเคยดาวน์โหลดไว้ก่อนหน้านี้ด้วย (เช่น PEAK_PO_Export_*.xlsx)
    # เพื่อให้แน่ใจว่าตอนตรวจสอบไฟล์หลังดาวน์โหลด เราจะเจอไฟล์ที่เพิ่งดาวน์โหลดมาจริงๆ
    # (อาจจะปรับ pattern ให้ตรงกับชื่อไฟล์ที่ PEAK ตั้งให้มากที่สุด)
    existing_peak_files = glob.glob(
        os.path.join(download_path, "PEAK_PO_Export_*.xlsx"))  # หรือชื่อ pattern อื่นๆ ที่ PEAK ใช้
    existing_peak_files += glob.glob(os.path.join(download_path, "รายงานใบสั่งซื้อ*.xlsx"))
    for f_path in existing_peak_files:
        if os.path.basename(f_path) != desired_file_name:  # ไม่ลบ desired_file_name อีกรอบ (เผื่อไว้)
            try:
                os.remove(f_path)
                _log(f"ลบไฟล์ PEAK export เก่าที่อาจค้างอยู่: {f_path}")
            except Exception as e_rm_peak_old:
                _log(f"!!! Warning: ไม่สามารถลบไฟล์ PEAK export เก่า '{f_path}': {e_rm_peak_old} !!!")

    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    # chrome_options.add_argument("--headless")
    # chrome_options.add_argument("--disable-gpu")
    # chrome_options.add_argument("--window-size=1920,1080")
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

        wait = WebDriverWait(driver, 30)
        long_wait = WebDriverWait(driver, 45)
        short_wait = WebDriverWait(driver, 15)
        very_short_wait = WebDriverWait(driver, 5)  # สำหรับรอ element เล็กๆ น้อยๆ

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
            long_wait.until(
                EC.all_of(
                    EC.url_contains("peakaccount.com/"),
                    EC.none_of(EC.url_contains("login"))
                )
            )
            long_wait.until(
                EC.any_of(
                    EC.presence_of_element_located(
                        (By.XPATH, "//div[contains(@class, 'list')]//p[contains(@class, 'crop')]")),
                    EC.presence_of_element_located((By.ID, "mainNavBarBottom"))
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
            if driver: driver.quit(); return None
        else:
            _log("Login ดูเหมือนจะสำเร็จแล้ว และหน้าหลัง Login เริ่มโหลด.")

        # 1.5: เลือกกิจการ
        if "selectlist" in driver.current_url.lower():
            _log("อยู่ที่หน้าเลือกกิจการ กำลังดำเนินการเลือก...")
            try:
                first_business_in_list_xpath = "//div[contains(@class, 'list')]//p[contains(@class, 'crop') and contains(@class, 'textBold')]"
                long_wait.until(EC.visibility_of_element_located((By.XPATH, first_business_in_list_xpath)))
                _log("หน้ารายการกิจการปรากฏแล้ว.")
                _log(f"กำลังค้นหากิจการ '{target_business_name_to_select}' ในรายการ...")
                business_item_xpath = f"//div[contains(@class, 'list')]//div[contains(@class, 'col2')]/p[contains(@class, 'textBold') and normalize-space(.)='{target_business_name_to_select}']"
                business_element = long_wait.until(EC.element_to_be_clickable((By.XPATH, business_item_xpath)))
                business_element.click()
                _log("คลิกเลือกกิจการแล้ว. รอหน้า Dashboard/หน้าหลักของกิจการโหลด...")
                try:
                    long_wait.until(
                        EC.all_of(
                            EC.none_of(EC.url_contains("selectlist")),
                            EC.presence_of_element_located((By.ID, "mainNavBarBottom"))
                        )
                    )
                    current_url_after_select = driver.current_url
                    _log(f"URL หลังจากเลือกกิจการสำเร็จ: {current_url_after_select}")
                    if "selectlist" in current_url_after_select.lower():
                        _log(f"!!! วิกฤต: ยังคงอยู่ที่หน้า selectlist !!!")
                        driver.save_screenshot(
                            os.path.join(download_path, "business_selection_STUCK_ON_SELECTLIST.png"))
                        if driver: driver.quit(); return None
                    _log(f"เข้าสู่กิจการ '{target_business_name_to_select}' สำเร็จแล้ว.")
                except TimeoutException:
                    current_url_at_fail = driver.current_url
                    _log(f"!!! Timeout: ไม่สามารถยืนยันการออกจากหน้า selectlist. URL: {current_url_at_fail} !!!")
                    driver.save_screenshot(os.path.join(download_path, "business_selection_timeout_critical.png"))
                    if driver: driver.quit(); return None
            except TimeoutException:
                _log(f"!!! TimeoutException: ไม่พบกิจการ '{target_business_name_to_select}' !!!")
                driver.save_screenshot(os.path.join(download_path, "business_selection_not_found.png"))
                if driver: driver.quit(); return None
        elif EC.presence_of_element_located((By.ID, "mainNavBarBottom"))(driver):
            _log("ไม่ได้อยู่ที่หน้าเลือกกิจการ และดูเหมือนจะอยู่ที่หน้า Dashboard แล้ว.")
        else:
            _log("!!! สถานะไม่คาดคิด: ไม่ได้อยู่ที่หน้า selectlist และไม่สามารถยืนยันได้ว่าอยู่ที่หน้า Dashboard. !!!")
            _log(f"URL ปัจจุบัน: {driver.current_url}")
            driver.save_screenshot(os.path.join(download_path, "unknown_state_after_login.png"))
            if driver: driver.quit(); return None

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 2: นำทางไปยังหน้า Purchase Order (ดูทั้งหมด)
        # --------------------------------------------------------------------
        _log("ขั้นตอน: การนำทางไปยังหน้า Purchase Order (ดูทั้งหมด) ผ่านเมนู Hover...")
        actions = ActionChains(driver)
        expense_menu_xpath = "//li[@id='Menu_expense']/descendant::a[contains(normalize-space(.), 'รายจ่าย')][1]"
        # XPath ที่ปรับปรุงสำหรับ "ใบสั่งซื้อ" (สำหรับ Hover)
        po_submenu_to_hover_xpath = "//li[@id='Menu_expense']//div[contains(@class, 'dropdown menu-margin')]//div[@name='selectDropdown'][.//a[normalize-space(.)='ใบสั่งซื้อ']]//a[@class='nameSelect' and normalize-space(.)='ใบสั่งซื้อ']"
        # XPath ที่ปรับปรุงสำหรับ "ดูทั้งหมด" (สำหรับ Click)
        view_all_po_actual_link_xpath = "//li[@id='Menu_expense']//div[@name='selectDropdown'][.//a[normalize-space(.)='ใบสั่งซื้อ']]//div[contains(@class, 'optionDropdown')]//a[@class='nemeOption' and normalize-space(.)='ดูทั้งหมด']"

        try:  # Try block สำหรับการนำทางเมนูทั้งหมด
            _log(f"กำลังรอเมนู 'รายจ่าย' (XPATH: {expense_menu_xpath})")
            expense_menu_element = wait.until(EC.visibility_of_element_located((By.XPATH, expense_menu_xpath)))
            _log("พบเมนู 'รายจ่าย'. กำลัง Hover...")
            actions.move_to_element(expense_menu_element).perform()
            _log("Hover บน 'รายจ่าย' แล้ว. รอ 1.5 วินาที...")
            time.sleep(1.5)

            _log(f"กำลังรอเมนูย่อย 'ใบสั่งซื้อ' (XPATH: {po_submenu_to_hover_xpath}) ให้ 'ปรากฏ' (visibility)")
            try:
                short_wait.until(EC.visibility_of_element_located((By.XPATH, po_submenu_to_hover_xpath)))
                _log("เมนูย่อย 'ใบสั่งซื้อ' ปรากฏแล้ว (visible).")
            except TimeoutException:
                _log(f"!!! Timeout: เมนูย่อย 'ใบสั่งซื้อ' (XPATH: {po_submenu_to_hover_xpath}) ไม่ปรากฏ (visible)!!!")
                driver.save_screenshot(os.path.join(download_path, "peak_po_submenu_not_visible.png"))
                # ถ้าเมนูย่อยไม่ปรากฏ ถือว่าการนำทางล้มเหลว
                if driver: driver.quit()
                return None  # ออกจากฟังก์ชันหลัก

            _log(f"กำลังรอเมนูย่อย 'ใบสั่งซื้อ' (XPATH: {po_submenu_to_hover_xpath}) ให้ 'clickable'")
            po_submenu_element_to_hover = wait.until(EC.element_to_be_clickable((By.XPATH, po_submenu_to_hover_xpath)))
            _log("พบเมนูย่อย 'ใบสั่งซื้อ' (สำหรับ Hover). กำลัง Hover...")
            actions.move_to_element(po_submenu_element_to_hover).perform()
            _log("Hover บน 'ใบสั่งซื้อ' แล้ว. รอ 1.5 วินาที...")
            time.sleep(1.5)

            _log(f"กำลังรอลิงก์ 'ดูทั้งหมด' (XPATH: {view_all_po_actual_link_xpath}) ให้ 'ปรากฏ' (visibility)")
            try:
                short_wait.until(EC.visibility_of_element_located((By.XPATH, view_all_po_actual_link_xpath)))
                _log("ลิงก์ 'ดูทั้งหมด' ปรากฏแล้ว (visible).")
            except TimeoutException:
                _log(f"!!! Timeout: ลิงก์ 'ดูทั้งหมด' (XPATH: {view_all_po_actual_link_xpath}) ไม่ปรากฏ (visible)!!!")
                driver.save_screenshot(os.path.join(download_path, "peak_po_viewall_link_not_visible.png"))
                # ถ้าลิงก์ "ดูทั้งหมด" ไม่ปรากฏ ถือว่าการนำทางล้มเหลว
                if driver: driver.quit()
                return None  # ออกจากฟังก์ชันหลัก

            _log(f"กำลังรอลิงก์ 'ดูทั้งหมด' (XPATH: {view_all_po_actual_link_xpath}) ให้ 'clickable'")
            view_all_link_element = wait.until(EC.element_to_be_clickable((By.XPATH, view_all_po_actual_link_xpath)))
            _log("พบลิงก์ 'ดูทั้งหมด'. กำลังคลิก...")
            driver.execute_script("arguments[0].click();", view_all_link_element)
            _log("คลิก 'ดูทั้งหมด' แล้ว (ด้วย JavaScript). รอหน้า PO โหลด...")

            # --- ตรวจสอบการมาถึงหน้า PO ---
            _log("รอการนำทางไปยังหน้า PO...")
            try:
                long_wait.until(lambda d: (("/expense/po" in d.current_url.lower() and (
                            "po" in d.title.lower() or "ใบสั่งซื้อ" in d.title.strip())) or (
                                                       "ใบสั่งซื้อ" in d.title.strip() and not "/income" in d.current_url.lower())))

                current_url_after_po_nav = driver.current_url
                page_title_after_po_nav = driver.title.strip()

                _log(f"URL ปัจจุบันหลังจากนำทางไปหน้า PO (raw): '{driver.current_url}'")
                _log(f"Page Title หลังจากนำทางไปหน้า PO (stripped): '{page_title_after_po_nav}'")

                is_on_po_page = False
                if "/expense/po" in current_url_after_po_nav.lower():
                    if "ใบสั่งซื้อ" in page_title_after_po_nav or "po" in page_title_after_po_nav.lower():
                        is_on_po_page = True

                if not is_on_po_page and (
                        "ใบสั่งซื้อ" in page_title_after_po_nav or "po" in page_title_after_po_nav.lower()):
                    try:
                        # **** สำคัญ: ปรับ XPath นี้ให้ตรงกับ element ที่ยืนยันว่าเป็นหน้ารายการ PO ****
                        # เช่น ปุ่ม "สร้างใบสั่งซื้อ" หรือ header ของตาราง หรือตารางเอง
                        confirm_po_page_element_xpath = "//div[contains(@class, 'header-section')]//h1[contains(normalize-space(), 'ใบสั่งซื้อ')]"
                        short_wait.until(EC.presence_of_element_located((By.XPATH, confirm_po_page_element_xpath)))
                        _log(
                            f"พบ element '{confirm_po_page_element_xpath}' บ่งชี้ของหน้ารายการ PO. ถือว่ามาถึงหน้า PO แล้ว.")
                        is_on_po_page = True
                    except TimeoutException:
                        _log("ไม่พบ element บ่งชี้ของหน้ารายการ PO เพิ่มเติม.")

                if is_on_po_page:
                    _log("น่าจะอยู่ที่หน้า PO 'ดูทั้งหมด' ถูกต้องแล้ว. ดำเนินการต่อ...")
                    # <<< ไม่มี return หรือ quit ที่นี่ เพื่อให้โค้ดไหลไปส่วนที่ 3 >>>
                else:
                    _log(
                        f"!!! ล้มเหลว: ดูเหมือนจะไม่ได้อยู่ที่หน้า PO ที่ถูกต้องหลังคลิก 'ดูทั้งหมด'. URL: {current_url_after_po_nav}, Title: '{page_title_after_po_nav}' !!!")
                    driver.save_screenshot(os.path.join(download_path, "po_nav_failed_target_page_critical.png"))
                    if driver: driver.quit()
                    return None  # ออกจากฟังก์ชันหลักถ้ามาหน้า PO ไม่ถูก

            except TimeoutException:  # Timeout ในการรอ URL/Title ของหน้า PO เปลี่ยน
                _log(
                    f"!!! ล้มเหลว: Timeout ขณะรอการยืนยันหน้า PO. URL ปัจจุบัน: {driver.current_url}, Title: {driver.title} !!!")
                driver.save_screenshot(os.path.join(download_path, "po_nav_timeout_target_page_critical.png"))
                if driver: driver.quit()
                return None  # ออกจากฟังก์ชันหลัก

        except TimeoutException as e_nav_elements:  # Timeout ในการหา element ของเมนูต่างๆ (รายจ่าย, ใบสั่งซื้อ, ดูทั้งหมด)
            _log(f"!!! ล้มเหลว: TimeoutException ระหว่างการค้นหา element ของเมนู: {e_nav_elements} !!!")
            _log(
                f"ตรวจสอบ XPath: expense='{expense_menu_xpath}', po_hover='{po_submenu_to_hover_xpath}', view_all='{view_all_po_actual_link_xpath}'")
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_menu_nav_timeout_element_critical.png"))
                with open(os.path.join(download_path, "page_source_at_menu_timeout_critical.html"), "w",
                          encoding="utf-8") as f:
                    f.write(driver.page_source)
            except Exception as e_diag:
                _log(f"ไม่สามารถบันทึก diagnostics เพิ่มเติมได้: {e_diag}")
            if driver: driver.quit()
            return None  # ออกจากฟังก์ชันหลัก
        except Exception as e_general_nav:  # Error อื่นๆ ที่อาจเกิดขึ้นระหว่างการนำทางเมนู
            _log(f"!!! ล้มเหลว: Error ทั่วไประหว่างการนำทางผ่านเมนู: {e_general_nav} !!!")
            import traceback
            _log(traceback.format_exc())
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_menu_nav_general_error_critical.png"))
            except:
                pass
            if driver: driver.quit()
            return None  # ออกจากฟังก์ชันหลัก

            # ถ้าโค้ดมาถึงตรงนี้ หมายความว่าการนำทางไปยังหน้า PO "ดูทั้งหมด" สำเร็จสมบูรณ์
            # และไม่ได้มีการ return None จาก try-except block ข้างบน
            # โค้ดจะไหลต่อไปยัง "ขั้นตอนที่ 3" ที่วางอยู่ถัดจากส่วนนี้ในฟังก์ชันหลัก

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 3: คลิกปุ่ม "พิมพ์รายงาน" และจัดการ Pop-up
        # --------------------------------------------------------------------
        _log("ขั้นตอน: คลิกปุ่ม 'พิมพ์รายงาน' และจัดการ Pop-up...")

        # เพิ่มการรอให้ส่วนสำคัญของหน้า PO โหลดเสร็จก่อน (เช่น ตารางข้อมูล)
        try:
            _log("รอให้ตารางรายการ PO โหลด (อย่างน้อยแถวแรก)...")
            # **** สำคัญ: ปรับ XPath นี้ให้ตรงกับโครงสร้างตาราง PO ของคุณ ****
            # ตัวอย่าง: "//div[@id='tableBox']//table//tbody/tr[1]" หรือ "//table[contains(@class,'po-list-table')]//tbody/tr[1]"
            # หรือ "//div[contains(@class,'contentTable')]//table//tbody/tr[1]" ที่เคยลอง
            # หรือ XPath ที่ชี้ไปที่ container ของตารางที่มั่นใจว่ามีข้อมูล
            po_table_first_row_xpath = "//div[contains(@class, 'tableBox')]//div[contains(@class, 'table-responsive')]//table//tbody//tr[1]"  # ลอง XPath ที่เจาะจงขึ้นสำหรับ PEAK
            short_wait.until(EC.presence_of_element_located((By.XPATH, po_table_first_row_xpath)))
            _log("ตารางรายการ PO ดูเหมือนจะโหลดแล้ว.")
        except TimeoutException:
            _log(
                "!!! Timeout: ไม่สามารถยืนยันการโหลดตารางรายการ PO ได้. อาจจะยังพยายามคลิกปุ่ม 'พิมพ์รายงาน' ต่อ...")
            driver.save_screenshot(os.path.join(download_path, "peak_po_table_not_loaded.png"))
            # ไม่ return None ทันที เพราะปุ่มหลักอาจจะยังคลิกได้

        # XPath สำหรับปุ่ม "พิมพ์รายงาน" หลัก (ปรับปรุงแล้ว)
        print_report_main_button_xpath = "//div[contains(@class, 'header-section')]//button[contains(normalize-space(.), 'พิมพ์รายงาน') and .//i[contains(@class, 'icon-printer_device')]]"

        # XPath อื่นๆ สำหรับ Pop-up
        modal_xpath = "//div[contains(@class, 'modalBox') and @style[not(contains(., 'display: none'))] and contains(@class,'showModal')]"  # เพิ่ม .showModal เพื่อความแม่นยำ
        show_details_radio_label_xpath = "//div[contains(@class, 'modalBox')]//p[contains(@class, 'nameRadio') and normalize-space(.)='แสดงรายละเอียด']/preceding-sibling::div[contains(@class,'radio')]/label"
        checkbox_xpaths = {
            "ใบสั่งซื้อสินทรัพย์": "//div[contains(@class, 'modalBox')]//label[.//span[contains(@class, 'label') and normalize-space(.)='ใบสั่งซื้อสินทรัพย์']]",
            "ข้อมูลราคาและภาษี": "//div[contains(@class, 'modalBox')]//label[.//span[contains(@class, 'label') and normalize-space(.)='ข้อมูลราคาและภาษี']]",
            "กลุ่มจัดประเภท": "//div[contains(@class, 'modalBox')]//label[.//span[contains(@class, 'label') and normalize-space(.)='กลุ่มจัดประเภท']]",
            "ข้อมูลอื่น": "//div[contains(@class, 'modalBox')]//label[.//span[contains(@class, 'label') and normalize-space(.)='ข้อมูลอื่น']]",
            "ประวัติเอกสาร": "//div[contains(@class, 'modalBox')]//label[.//span[contains(@class, 'label') and normalize-space(.)='ประวัติเอกสาร']]",
            "เอกสารที่ถูกยกเลิก": "//div[contains(@class, 'modalBox')]//label[.//span[contains(@class, 'label') and normalize-space(.)='เอกสารที่ถูกยกเลิก']]"
        }
        print_report_in_modal_button_xpath = "//div[contains(@class, 'modalBox')]//div[contains(@class, 'Footer')]//button[normalize-space(.)='พิมพ์รายงาน' and not(contains(@class, 'btn-secondary')) and not(contains(@class, 'cancel'))]"

        # --- เริ่มการคลิกปุ่มหลัก และจัดการ Pop-up ---
        try:
            _log(f"กำลังค้นหาปุ่ม 'พิมพ์รายงาน' (หลัก) ด้วย XPath: {print_report_main_button_xpath}")
            _log(" - รอให้ปุ่ม 'พิมพ์รายงาน' (หลัก) ปรากฏ (visible)...")
            main_print_button_visible = wait.until(
                EC.visibility_of_element_located((By.XPATH, print_report_main_button_xpath)))
            _log(" - ปุ่ม 'พิมพ์รายงาน' (หลัก) ปรากฏแล้ว. รอให้ clickable...")
            main_print_button = wait.until(EC.element_to_be_clickable((By.XPATH, print_report_main_button_xpath)))

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", main_print_button)
            time.sleep(0.5)

            main_print_button.click()
            _log(
                "คลิกปุ่ม 'พิมพ์รายงาน' (หลัก) สำเร็จแล้ว. รอ Pop-up ปรากฏ...")  # Log นี้ควรจะแสดงก่อนเข้า try block ของ Pop-up

            # -----[ POPUP HANDLING BLOCK ]-----
            try:
                _log("--- ENTERING POPUP HANDLING TRY BLOCK ---")  # LOG สำหรับ DEBUG

                # XPath สำหรับ Modal - อัปเดตตาม HTML ที่ให้มา
                modal_content_to_wait_for_xpath = "//div[@id='modalBox' and @showmodal='true']"

                _log(f"กำลังรอ Pop-up ปรากฏ (XPath: {modal_content_to_wait_for_xpath})")
                # รอให้ element ปรากฏใน DOM และ visible (attribute showmodal='true' ควรจะทำให้มัน visible)
                # visibility_of_element_located จะเช็คทั้ง presence และว่า element นั้น displayed (height/width > 0)
                modal_element = wait.until(
                    EC.visibility_of_element_located((By.XPATH, modal_content_to_wait_for_xpath)))
                _log("Pop-up ปรากฏแล้ว (visible).")
                time.sleep(1.5)  # ให้เวลา UI update และ animation (ถ้ามี)

                # 1. คลิก Radio button "แสดงรายละเอียด"
                # **** อัปเดต XPath ที่นี่ ****
                show_details_radio_label_xpath_updated = "//div[@id='modalBox' and @showmodal='true']//label[.//p[normalize-space(.)='แสดงรายละเอียด']]"
                # หรือถ้า id 'tjd7eo1' คงที่:
                # show_details_radio_label_xpath_updated = "//div[@id='modalBox' and @showmodal='true']//label[@for='tjd7eo1']"

                _log(f"กำลังคลิก Radio button 'แสดงรายละเอียด' (XPath: {show_details_radio_label_xpath_updated})")

                # รอให้ label clickable
                show_details_radio_label_element = wait.until(
                    EC.element_to_be_clickable((By.XPATH, show_details_radio_label_xpath_updated)))

                # ลอง scroll element ให้อยู่ใน view ก่อนคลิก
                driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
                                      show_details_radio_label_element)
                time.sleep(0.3)  # ให้เวลานิดหน่อยหลัง scroll

                # ใช้ JavaScript click เพื่อความแน่นอน
                driver.execute_script("arguments[0].click();", show_details_radio_label_element)
                # หรือลอง click ปกติ ถ้า JavaScript click ไม่ได้ผลด้วยเหตุผลบางอย่าง
                # show_details_radio_label_element.click()
                _log("คลิก Radio 'แสดงรายละเอียด' แล้ว.")
                time.sleep(0.5)  # ให้เวลา UI update

                # 2. คลิก Checkbox ทั้งหมดในส่วน "แสดงข้อมูลเพิ่มเติม"
                _log("กำลังคลิก Checkbox ใน 'แสดงข้อมูลเพิ่มเติม'...")
                for name, cb_xpath in checkbox_xpaths.items():
                    try:
                        _log(f" - กำลังค้นหา Checkbox '{name}' (XPath: {cb_xpath})")
                        checkbox_label = wait.until(EC.presence_of_element_located((By.XPATH, cb_xpath)))

                        # ตรวจสอบว่า checkbox element (input) ที่สัมพันธ์กับ label นี้ถูกติ๊กหรือยัง
                        # โดยทั่วไป label จะมี attribute 'for' ที่ตรงกับ 'id' ของ input
                        input_id = checkbox_label.get_attribute("for")
                        if input_id:
                            actual_checkbox_input = driver.find_element(By.ID, input_id)
                            if not actual_checkbox_input.is_selected():
                                _log(f"   Checkbox '{name}' ยังไม่ได้ถูกติ๊ก. กำลังคลิกที่ Label...")
                                driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();",
                                                      checkbox_label)
                                _log(f"   คลิก Label ของ Checkbox '{name}' แล้ว.")
                                time.sleep(0.2)
                            else:
                                _log(f"   Checkbox '{name}' ถูกติ๊กอยู่แล้ว.")
                        else:
                            _log(
                                f"   ไม่พบ attribute 'for' บน label ของ '{name}'. ลองคลิก label โดยตรงเผื่อได้ผล...")
                            # ถ้าไม่มี 'for', อาจจะต้องคลิกที่ input โดยตรง หรือ label อาจจะครอบ input
                            # ลองคลิก label ดูก่อน
                            # ตรวจสอบสถานะการเลือกที่ซับซ้อนขึ้นถ้าจำเป็น (เช่น ตรวจสอบ class ของ parent)
                            # ในที่นี้จะลองคลิก label ไปเลย
                            is_selected_somehow = False  # ต้องหาวิธีเช็คถ้าไม่มี for/id
                            # ตัวอย่าง: is_selected_somehow = "active" in checkbox_label.find_element(By.XPATH, "./parent::div[contains(@class,'checkbox')]").get_attribute("class")
                            # ถ้ายังไม่ซับซ้อนขนาดนั้น ลองคลิกไปก่อน
                            _log(f"   Checkbox '{name}' (ไม่มี for/id ชัดเจน) - กำลังคลิกที่ Label...")
                            driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();",
                                                  checkbox_label)
                            _log(f"   คลิก Label ของ Checkbox '{name}' แล้ว.")
                            time.sleep(0.2)

                    except TimeoutException:
                        _log(f"!!! Timeout: ไม่พบ Checkbox '{name}' หรือไม่พร้อมคลิก. ข้ามไป...")
                    except NoSuchElementException:
                        _log(f"!!! Error: ไม่พบ Input ของ Checkbox '{name}' จาก ID (ถ้ามี). ข้ามไป...")
                    except Exception as e_cb:
                        _log(f"!!! Error ขณะพยายามคลิก Checkbox '{name}': {e_cb}. ข้ามไป...")

                _log("คลิก Checkbox ทั้งหมด (ที่พบ) เสร็จสิ้น.")
                time.sleep(0.5)

                # 3. คลิกปุ่ม "พิมพ์รายงาน" ใน Pop-up
                _log(f"กำลังคลิกปุ่ม 'พิมพ์รายงาน' ใน Pop-up (XPath: {print_report_in_modal_button_xpath})")
                final_print_button = wait.until(
                    EC.element_to_be_clickable((By.XPATH, print_report_in_modal_button_xpath)))
                driver.execute_script("arguments[0].click();", final_print_button)

                _log("คลิกปุ่ม 'พิมพ์รายงาน' ใน Pop-up แล้ว. การดาวน์โหลดควรจะเริ่มขึ้น...")
                _log("--- EXITING POPUP HANDLING TRY BLOCK SUCCESSFULLY ---")  # LOG สำหรับ DEBUG

            except TimeoutException as e_report_section:
                _log(f"!!! POPUP HANDLING: TimeoutException: {e_report_section} !!!")
                try:
                    driver.save_screenshot(os.path.join(download_path, "peak_po_report_popup_timeout.png"))
                    with open(os.path.join(download_path, "page_source_at_report_popup_timeout.html"), "w",
                              encoding="utf-8") as f:
                        f.write(driver.page_source)
                except Exception as e_diag:
                    _log(f"ไม่สามารถบันทึก diagnostics ได้: {e_diag}")
                if driver: driver.quit()
                return None
            except ElementClickInterceptedException as e_click_intercept:
                _log(f"!!! POPUP HANDLING: ElementClickInterceptedException: {e_click_intercept} !!!")
                try:
                    driver.save_screenshot(os.path.join(download_path, "peak_po_click_intercepted.png"))
                except:
                    pass
                if driver: driver.quit()
                return None
            except Exception as e_popup_general:
                _log(f"!!! POPUP HANDLING: Error ทั่วไป: {e_popup_general} !!!")
                import traceback
                _log(traceback.format_exc())
                try:
                    driver.save_screenshot(os.path.join(download_path, "peak_po_popup_general_error.png"))
                except:
                    pass
                if driver: driver.quit()
                return None
            # -----[ END OF POPUP HANDLING BLOCK ]-----

            # ถ้าโค้ดมาถึงตรงนี้ได้ แสดงว่าการจัดการ Pop-up (ถ้ามี) สำเร็จ หรือไม่มี Pop-up ให้จัดการ
            # และควรจะไปต่อที่ขั้นตอนการรอไฟล์ดาวน์โหลด
            _log("--- PROCEEDING TO DOWNLOAD WAIT ---")

        except TimeoutException as e_main_print_button:  # Timeout จากการรอปุ่ม "พิมพ์รายงาน" (หลัก)
            _log(f"!!! TimeoutException ขณะรอหรือคลิกปุ่ม 'พิมพ์รายงาน' (หลัก): {e_main_print_button} !!!")
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_main_print_button_timeout.png"))
                with open(os.path.join(download_path, "page_source_at_main_print_button_timeout.html"), "w",
                          encoding="utf-8") as f:
                    f.write(driver.page_source)
            except Exception as e_diag:
                _log(f"ไม่สามารถบันทึก diagnostics ได้: {e_diag}")
            if driver: driver.quit()
            return None
        except Exception as e_outer_level3:  # Error อื่นๆ ในการคลิกปุ่มหลัก ก่อนเข้า popup handling
            _log(f"!!! Error ทั่วไปในขั้นตอนที่ 3 (ก่อนเข้า Pop-up handling): {e_outer_level3} !!!")
            import traceback
            _log(traceback.format_exc())
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_step3_outer_error.png"))
            except:
                pass
            if driver: driver.quit()
            return None

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 4: รอไฟล์ดาวน์โหลด และเปลี่ยนชื่อ
        # --------------------------------------------------------------------
        _log("ขั้นตอน: รอไฟล์ดาวน์โหลด และเปลี่ยนชื่อ...")

        # คาดเดาชื่อไฟล์ที่ PEAK จะดาวน์โหลด (อาจจะต้องปรับปรุง)
        # ปกติ PEAK จะใช้ชื่อประมาณ "PEAK_PO_Export_YYYYMMDDHHMMSS.xlsx" หรือ "รายงานใบสั่งซื้อ_ถึง_วันที่_DD_MM_YYYY.xlsx"
        # เราต้องหา pattern ที่ยืดหยุ่นพอ
        # หรือใช้วิธีเช็คไฟล์ใหม่ที่เกิดขึ้นใน download_path

        download_wait_timeout = 120  # รอสูงสุด 2 นาทีสำหรับการดาวน์โหลด
        poll_interval = 2  # ตรวจสอบทุกๆ 2 วินาที
        time_elapsed = 0
        downloaded_file_path = None

        _log(f"กำลังรอไฟล์ .xlsx ใหม่ในโฟลเดอร์: {download_path} (รอสูงสุด {download_wait_timeout} วินาที)")

        # เก็บรายชื่อไฟล์ .xlsx ที่มีอยู่ก่อนเริ่มดาวน์โหลด (หรือก่อนคลิกปุ่มสุดท้าย)
        # เราได้ลบไฟล์เก่าไปแล้ว แต่เผื่อมีกรณีอื่น
        files_before_download = set(glob.glob(os.path.join(download_path, "*.xlsx")))

        while time_elapsed < download_wait_timeout:
            time.sleep(poll_interval)
            time_elapsed += poll_interval

            current_files = set(glob.glob(os.path.join(download_path, "*.xlsx")))
            new_files = current_files - files_before_download

            if new_files:
                # ตรวจสอบว่ามีไฟล์ .crdownload หรือ .tmp หรือไม่ (แสดงว่ายังดาวน์โหลดไม่เสร็จ)
                # (การตรวจสอบ .crdownload ใช้ได้กับ Chrome)
                temp_files_exist = any(
                    f.endswith((".crdownload", ".tmp")) for f in glob.glob(os.path.join(download_path, "*.*")))

                if not temp_files_exist:
                    # สมมติว่าไฟล์แรกที่พบคือไฟล์ที่ต้องการ (ถ้ามีหลายไฟล์ใหม่ อาจต้องมี logic เพิ่ม)
                    downloaded_file_path = list(new_files)[0]
                    _log(f"พบไฟล์ใหม่ที่ดาวน์โหลดเสร็จแล้ว: {downloaded_file_path}")
                    break
                else:
                    _log(
                        f"พบไฟล์ใหม่ ({len(new_files)}) แต่ยังอยู่ในสถานะดาวน์โหลด (.crdownload/.tmp)... รอต่อ ({time_elapsed}/{download_wait_timeout}s)")
            else:
                _log(f"ยังไม่พบไฟล์ .xlsx ใหม่... รอต่อ ({time_elapsed}/{download_wait_timeout}s)")

        if downloaded_file_path and os.path.exists(downloaded_file_path):
            _log(f"ไฟล์ '{os.path.basename(downloaded_file_path)}' ดาวน์โหลดสำเร็จแล้ว.")

            final_file_path = os.path.join(download_path, desired_file_name)

            # ตรวจสอบอีกครั้งว่า desired_file_name ซ้ำกับไฟล์ที่เพิ่งดาวน์โหลดมาหรือไม่ (ไม่ควรเกิดถ้า logic ลบไฟล์เก่าดี)
            if os.path.abspath(downloaded_file_path) == os.path.abspath(final_file_path):
                _log(
                    f"ไฟล์ที่ดาวน์โหลดมามีชื่อตรงกับ desired_file_name อยู่แล้ว: '{final_file_path}' ไม่ต้องเปลี่ยนชื่อ.")
                if driver: driver.quit(); return final_file_path  # สำเร็จ
            else:
                # ถ้า desired_file_name มีอยู่แล้ว (กรณีแปลกๆ) ให้ลบก่อน
                if os.path.exists(final_file_path):
                    try:
                        os.remove(final_file_path)
                        _log(f"ลบไฟล์ '{final_file_path}' ที่มีชื่อซ้ำกับ desired_file_name ก่อนทำการเปลี่ยนชื่อ.")
                    except Exception as e_rm_冲突:
                        _log(f"!!! Error: ไม่สามารถลบไฟล์ '{final_file_path}' ที่มีชื่อซ้ำ: {e_rm_冲突} !!!")
                        # อาจจะ return None หรือลองเปลี่ยนชื่อเป็นชื่ออื่น
                        if driver: driver.quit(); return None

                try:
                    os.rename(downloaded_file_path, final_file_path)
                    _log(
                        f"เปลี่ยนชื่อไฟล์จาก '{os.path.basename(downloaded_file_path)}' เป็น '{desired_file_name}' เรียบร้อยแล้ว.")
                    _log(f"ตำแหน่งไฟล์สุดท้าย: {final_file_path}")
                    if driver: driver.quit(); return final_file_path  # สำเร็จ
                except Exception as e_rename:
                    _log(
                        f"!!! Error เปลี่ยนชื่อไฟล์จาก '{downloaded_file_path}' เป็น '{final_file_path}': {e_rename} !!!")
                    _log(f"ไฟล์ยังคงอยู่ที่: {downloaded_file_path}")
                    if driver: driver.quit(); return downloaded_file_path  # คืนไฟล์เดิมถ้าเปลี่ยนชื่อไม่ได้
        else:
            _log(
                f"!!! Timeout หรือ Error: ไม่สามารถยืนยันการดาวน์โหลดไฟล์ .xlsx ได้ภายใน {download_wait_timeout} วินาที !!!")
            driver.save_screenshot(os.path.join(download_path, "peak_po_download_timeout_or_error.png"))
            if driver: driver.quit(); return None

    except TimeoutException as e_report_section:
        _log(f"!!! TimeoutException ระหว่างการจัดการ Pop-up พิมพ์รายงาน: {e_report_section} !!!")
        try:
            driver.save_screenshot(os.path.join(download_path, "peak_po_report_popup_timeout.png"))
            with open(os.path.join(download_path, "page_source_at_report_popup_timeout.html"), "w",
                      encoding="utf-8") as f:
                f.write(driver.page_source)
        except Exception as e_diag:
            _log(f"ไม่สามารถบันทึก diagnostics ได้: {e_diag}")
        if driver: driver.quit(); return None
    except ElementClickInterceptedException as e_click_intercept:
        _log(f"!!! ElementClickInterceptedException: Element ถูกบัง ทำให้คลิกไม่ได้: {e_click_intercept} !!!")
        _log("อาจจะต้องลองใช้ JavaScript click หรือเพิ่มการรอให้ element ที่บังหายไป")
        try:
            driver.save_screenshot(os.path.join(download_path, "peak_po_click_intercepted.png"))
        except:
            pass
        if driver: driver.quit(); return None


    except TimeoutException as te:
        _log(f"!!! TimeoutException (Overall): {te} !!!")
        if driver:
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_timeout_error.png"))
            except:
                pass
            driver.quit()
        return None
    except Exception as e:
        _log(f"!!! Fatal Error (Overall): {e} !!!")
        import traceback
        _log(traceback.format_exc())
        if driver:
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_fatal_error.png"))
            except:
                pass
            driver.quit()
        return None
    finally:
        if driver:
            try:
                _log("Ensuring WebDriver is quit in finally block.")
                driver.quit()
            except:  # WebDriver อาจจะถูก quit ไปแล้ว หรือมี error ตอน quit
                _log("WebDriver already quit or error during quit in finally.")
        _log("Function finished.")  # เพิ่ม log ตอนจบฟังก์ชัน


# --- ส่วน if __name__ == '__main__': สำหรับทดสอบ ---
if __name__ == '__main__':
    print("=" * 30)
    print("  Testing auto_downloader.py (PEAK PO - Full Download Attempt)  ")
    print("=" * 30)

    test_peak_user = "sirichai.c@zubbsteel.com"
    test_peak_pass = "Zubb*2013"
    test_target_business = "บจ. บิซ ฮีโร่ (สำนักงานใหญ่)"
    test_desired_filename = "BizHero_PO_Report_Latest.xlsx"  # ลองเปลี่ยนชื่อไฟล์ที่ต้องการ

    if "YOUR_PEAK_EMAIL_HERE" in test_peak_user or \
            "YOUR_PEAK_PASSWORD_HERE" in test_peak_pass or \
            "ชื่อกิจการของคุณที่นี่" in test_target_business:
        print("\n!!! กรุณาแก้ไข test_peak_user, test_peak_pass, และ test_target_business ก่อนรันทดสอบ !!!\n")
    else:
        current_script_dir = os.getcwd()
        # สร้างโฟลเดอร์แยกสำหรับแต่ละครั้งที่รันเทส หรือใช้ชื่อเดิมก็ได้
        # test_save_dir_for_artifacts = os.path.join(current_script_dir, f"peak_po_full_test_{time.strftime('%Y%m%d_%H%M%S')}")
        test_save_dir_for_artifacts = os.path.join(current_script_dir, "peak_po_download_results")


        def standalone_logger(message):
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{timestamp}][StandaloneTestLogger] {message}")


        print(f"Username: {test_peak_user}")
        print(f"Target Business: {test_target_business}")
        print(f"Save Directory: {test_save_dir_for_artifacts}")
        print(f"Desired Filename: {test_desired_filename}")

        # ตรวจสอบและสร้างโฟลเดอร์ save_directory ก่อนเรียกฟังก์ชันหลัก
        # (ฟังก์ชันหลักก็มีสร้าง แต่ทำตรงนี้ด้วยเพื่อความชัดเจน)
        if not os.path.exists(test_save_dir_for_artifacts):
            os.makedirs(test_save_dir_for_artifacts)
            print(f"สร้างโฟลเดอร์สำหรับผลลัพธ์: {test_save_dir_for_artifacts}")

        result_file_path = download_peak_purchase_order_report(
            username=test_peak_user,
            password=test_peak_pass,
            target_business_name_to_select=test_target_business,
            save_directory=test_save_dir_for_artifacts,
            desired_file_name=test_desired_filename,  # ส่งชื่อไฟล์ที่ต้องการ
            log_callback=standalone_logger
        )
        if result_file_path and os.path.exists(result_file_path):
            print(f"\n[StandaloneTestLogger] ดาวน์โหลดไฟล์สำเร็จ! ไฟล์อยู่ที่: {result_file_path}")
        elif result_file_path:  # กรณีคืน path มาแต่ไฟล์อาจจะไม่มี (เช่น เปลี่ยนชื่อไม่ได้)
            print(f"\n[StandaloneTestLogger] การดาวน์โหลดอาจจะมีปัญหาบางส่วน. Result path: {result_file_path}")
            print(f"กรุณาตรวจสอบไฟล์ในโฟลเดอร์ '{test_save_dir_for_artifacts}'")
        else:
            print(f"\n[StandaloneTestLogger] การดาวน์โหลดไฟล์ไม่สำเร็จ หรือมีข้อผิดพลาด.")
            print(f"กรุณาตรวจสอบ log และไฟล์ screenshot (ถ้ามี) ในโฟลเดอร์ '{test_save_dir_for_artifacts}'")

    print("=" * 30)
    print("  Finished Testing auto_downloader.py (PEAK PO - Full Download Attempt)  ")
    print("=" * 30)