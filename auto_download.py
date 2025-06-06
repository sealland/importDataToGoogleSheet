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
    # ... (ส่วนสร้างโฟลเดอร์ download_path และลบไฟล์เก่า เหมือนเดิม) ...
    if not os.path.exists(download_path):
        try:
            os.makedirs(download_path)
            _log(f"สร้างโฟลเดอร์ดาวน์โหลด: {download_path}")
        except Exception as e_mkdir:
            _log(f"!!! Error สร้างโฟลเดอร์ดาวน์โหลด {download_path}: {e_mkdir} !!!")
            return None

    potential_old_file = os.path.join(download_path, desired_file_name)
    if os.path.exists(potential_old_file):
        try:
            os.remove(potential_old_file)
            _log(f"ลบไฟล์เก่า '{potential_old_file}' ที่อาจค้างอยู่แล้ว")
        except Exception as e_rm_old:
            _log(f"!!! Warning: ไม่สามารถลบไฟล์เก่า '{potential_old_file}': {e_rm_old} !!!")

    existing_peak_files = glob.glob(os.path.join(download_path, "PEAK_PO_Export_*.xlsx"))
    existing_peak_files += glob.glob(os.path.join(download_path, "รายงานใบสั่งซื้อ*.xlsx"))
    for f_path in existing_peak_files:
        if os.path.basename(f_path) != desired_file_name:
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
    _log("Chrome options configured.")

    driver = None
    try:  # TRY BLOCK หลักของฟังก์ชัน
        _log("กำลังเริ่ม WebDriver สำหรับ PEAK...")
        try:
            chrome_driver_path = ChromeDriverManager().install()
            _log(f"ChromeDriver path: {chrome_driver_path}")
            driver_service = ChromeService(executable_path=chrome_driver_path)
            driver = webdriver.Chrome(service=driver_service, options=chrome_options)
        except Exception as e_webdriver_init:
            _log(f"!!! Error initializing WebDriver: {e_webdriver_init} !!!")
            _log(traceback.format_exc())
            return None

        wait = WebDriverWait(driver, 30)
        long_wait = WebDriverWait(driver, 45)
        short_wait = WebDriverWait(driver, 15)

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 1: Login และ เลือกกิจการ
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 1: Login และ เลือกกิจการ...")
        # ... (โค้ด Login และ Select Business ที่ทำงานได้ดีแล้ว) ...
        login_url = "https://secure.peakaccount.com/login"
        driver.get(login_url)
        wait.until(EC.presence_of_element_located((By.NAME, "email"))).send_keys(username)
        wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(password)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space(.)='เข้าสู่ระบบ PEAK']"))).click()
        _log("คลิกปุ่ม Login แล้ว.")
        try:
            long_wait.until(EC.all_of(EC.url_contains("peakaccount.com/"), EC.none_of(EC.url_contains("login"))))
            long_wait.until(EC.any_of(
                EC.presence_of_element_located(
                    (By.XPATH, "//div[contains(@class, 'list')]//p[contains(@class, 'crop')]")),
                EC.presence_of_element_located((By.ID, "mainNavBarBottom"))
            ))
            _log(f"Login สำเร็จ. URL ปัจจุบัน: {driver.current_url}")
        except TimeoutException:
            _log("!!! Timeout: Login ไม่สำเร็จ หรือหน้าหลัง Login ไม่โหลด !!!")
            try:
                driver.save_screenshot(os.path.join(download_path, "po_login_failed.png"))
            except:
                pass
            if driver: driver.quit()
            return None

        if "selectlist" in driver.current_url.lower():
            _log("อยู่ที่หน้าเลือกกิจการ...")
            try:
                long_wait.until(EC.visibility_of_element_located((By.XPATH,
                                                                  "//div[contains(@class, 'list')]//p[contains(@class, 'crop') and contains(@class, 'textBold')]")))
                business_item_xpath = f"//div[contains(@class, 'list')]//div[contains(@class, 'col2')]/p[contains(@class, 'textBold') and normalize-space(.)='{target_business_name_to_select}']"
                long_wait.until(EC.element_to_be_clickable((By.XPATH, business_item_xpath))).click()
                _log("คลิกเลือกกิจการแล้ว.")
                long_wait.until(EC.all_of(
                    EC.none_of(EC.url_contains("selectlist")),
                    EC.presence_of_element_located((By.ID, "mainNavBarBottom"))
                ))
                _log(f"เข้าสู่กิจการ '{target_business_name_to_select}' สำเร็จ. URL: {driver.current_url}")
            except TimeoutException:
                _log(f"!!! Timeout: เลือกกิจการ '{target_business_name_to_select}' ไม่สำเร็จ !!!")
                try:
                    driver.save_screenshot(os.path.join(download_path, "po_business_select_failed.png"))
                except:
                    pass
                if driver: driver.quit()
                return None
        elif EC.presence_of_element_located((By.ID, "mainNavBarBottom"))(driver):
            _log("เข้าสู่ Dashboard โดยตรงแล้ว.")
        else:
            _log("!!! สถานะไม่คาดคิดหลัง Login !!!")
            if driver: driver.quit()
            return None
        _log("ขั้นตอนที่ 1 เสร็จสิ้น.")

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 2: นำทางไปยังหน้า Purchase Order (ดูทั้งหมด)
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 2: การนำทางไปยังหน้า Purchase Order (ดูทั้งหมด)...")
        actions = ActionChains(driver)
        expense_menu_xpath = "//li[@id='Menu_expense']/descendant::a[contains(normalize-space(.), 'รายจ่าย')][1]"
        po_submenu_to_hover_xpath = "//li[@id='Menu_expense']//div[contains(@class, 'dropdown menu-margin')]//div[@name='selectDropdown'][.//a[normalize-space(.)='ใบสั่งซื้อ']]//a[@class='nameSelect' and normalize-space(.)='ใบสั่งซื้อ']"
        view_all_po_actual_link_xpath = "//li[@id='Menu_expense']//div[@name='selectDropdown'][.//a[normalize-space(.)='ใบสั่งซื้อ']]//div[contains(@class, 'optionDropdown')]//a[@class='nemeOption' and normalize-space(.)='ดูทั้งหมด']"

        try:
            expense_menu_element = wait.until(EC.visibility_of_element_located((By.XPATH, expense_menu_xpath)))
            actions.move_to_element(expense_menu_element).perform();
            _log("Hover บน 'รายจ่าย' แล้ว. รอ 1.5 วินาที...");
            time.sleep(1.5)
            wait.until(EC.visibility_of_element_located((By.XPATH, po_submenu_to_hover_xpath)))
            po_submenu_element_to_hover = wait.until(EC.element_to_be_clickable((By.XPATH, po_submenu_to_hover_xpath)))
            actions.move_to_element(po_submenu_element_to_hover).perform();
            _log("Hover บน 'ใบสั่งซื้อ' แล้ว. รอ 1.5 วินาที...");
            time.sleep(1.5)
            wait.until(EC.visibility_of_element_located((By.XPATH, view_all_po_actual_link_xpath)))
            view_all_link_element = wait.until(EC.element_to_be_clickable((By.XPATH, view_all_po_actual_link_xpath)))
            driver.execute_script("arguments[0].click();", view_all_link_element);
            _log("คลิก 'ดูทั้งหมด' แล้ว.")
            _log("รอการนำทางไปยังหน้า PO...")
            try:
                long_wait.until(lambda d: (("/expense/po" in d.current_url.lower() and (
                            "po" in d.title.lower() or "ใบสั่งซื้อ" in d.title.strip())) or (
                                                       "ใบสั่งซื้อ" in d.title.strip() and not "/income" in d.current_url.lower())))
                is_on_po_page = True
                if not ("/expense/po" in driver.current_url.lower() and (
                        "ใบสั่งซื้อ" in driver.title.strip() or "po" in driver.title.lower())):
                    try:
                        confirm_po_page_element_xpath = "//div[contains(@class, 'header-section')]//h1[contains(normalize-space(), 'ใบสั่งซื้อ')]"
                        short_wait.until(EC.presence_of_element_located((By.XPATH, confirm_po_page_element_xpath)))
                    except TimeoutException:
                        is_on_po_page = False
                if is_on_po_page:
                    _log(
                        f"น่าจะอยู่ที่หน้า PO 'ดูทั้งหมด' ถูกต้องแล้ว. URL: {driver.current_url}, Title: {driver.title.strip()}")
                else:
                    _log(
                        f"!!! ล้มเหลว: ไม่ได้อยู่ที่หน้า PO ที่ถูกต้อง. URL: {driver.current_url}, Title: {driver.title.strip()} !!!");
                    if driver: driver.quit(); return None
            except TimeoutException:
                _log(
                    f"!!! ล้มเหลว: Timeout ขณะรอการยืนยันหน้า PO. URL: {driver.current_url}, Title: {driver.title} !!!");
                if driver: driver.quit(); return None
        except Exception as e_nav:
            _log(f"!!! ล้มเหลว: Error ระหว่างการนำทางผ่านเมนู: {e_nav} !!!");
            _log(traceback.format_exc())
            try:
                driver.save_screenshot(os.path.join(download_path, "po_menu_nav_error.png"))
            except:
                pass
            if driver: driver.quit(); return None
        _log("ขั้นตอนที่ 2 เสร็จสิ้น.")

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 3: คลิกปุ่ม "พิมพ์รายงาน" และจัดการ Pop-up
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 3: คลิกปุ่ม 'พิมพ์รายงาน' และจัดการ Pop-up...")
        print_report_main_button_xpath = "//div[contains(@class, 'header-section')]//button[contains(normalize-space(.), 'พิมพ์รายงาน') and .//i[contains(@class, 'icon-printer_device')]]"
        modal_xpath = "//div[@id='modalBox' and @showmodal='true']"
        show_details_radio_label_xpath = "//div[@id='modalBox' and @showmodal='true']//label[.//p[normalize-space(.)='แสดงรายละเอียด']]"
        checkbox_xpaths = {
            "ใบสั่งซื้อสินทรัพย์": "//div[@id='modalBox' and @showmodal='true']//label[.//span[contains(@class, 'label') and normalize-space(.)='ใบสั่งซื้อสินทรัพย์']]",
            "ข้อมูลราคาและภาษี": "//div[@id='modalBox' and @showmodal='true']//label[.//span[contains(@class, 'label') and normalize-space(.)='ข้อมูลราคาและภาษี']]",
            "กลุ่มจัดประเภท": "//div[@id='modalBox' and @showmodal='true']//label[.//span[contains(@class, 'label') and normalize-space(.)='กลุ่มจัดประเภท']]",
            "ข้อมูลอื่น": "//div[@id='modalBox' and @showmodal='true']//label[.//span[contains(@class, 'label') and normalize-space(.)='ข้อมูลอื่น']]",
            "ประวัติเอกสาร": "//div[@id='modalBox' and @showmodal='true']//label[.//span[contains(@class, 'label') and normalize-space(.)='ประวัติเอกสาร']]",
            "เอกสารที่ถูกยกเลิก": "//div[@id='modalBox' and @showmodal='true']//label[.//span[contains(@class, 'label') and normalize-space(.)='เอกสารที่ถูกยกเลิก']]"
        }
        print_report_in_modal_button_xpath = "//div[@id='modalBox' and @showmodal='true']//button[contains(normalize-space(.), 'พิมพ์รายงาน') and not(ancestor::div[contains(@class,'secondary')])]"

        try:
            main_print_button = wait.until(EC.element_to_be_clickable((By.XPATH, print_report_main_button_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", main_print_button);
            time.sleep(0.5)
            main_print_button.click();
            _log("คลิกปุ่ม 'พิมพ์รายงาน' (หลัก) สำเร็จแล้ว.")

            try:  # POPUP HANDLING BLOCK
                _log("--- ENTERING POPUP HANDLING TRY BLOCK ---")
                wait.until(EC.visibility_of_element_located((By.XPATH, modal_xpath)));
                _log(f"Pop-up ปรากฏแล้ว (XPath: {modal_xpath}).");
                time.sleep(1.5)

                el = wait.until(EC.element_to_be_clickable((By.XPATH, show_details_radio_label_xpath)))
                driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", el);
                _log("คลิก Radio 'แสดงรายละเอียด' แล้ว.");
                time.sleep(0.5)

                _log("กำลังคลิก Checkbox ใน 'แสดงข้อมูลเพิ่มเติม'...")
                for name, cb_xpath in checkbox_xpaths.items():
                    try:
                        el_cb = wait.until(EC.element_to_be_clickable((By.XPATH, cb_xpath)))
                        # (ส่วนตรวจสอบ is_selected ถ้าจำเป็น)
                        driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", el_cb);
                        _log(f"   คลิก Label ของ Checkbox '{name}'.");
                        time.sleep(0.2)
                    except Exception as e_cb:
                        _log(f"!!! Error/Timeout คลิก Checkbox '{name}': {e_cb}. ข้าม...")
                _log("คลิก Checkbox ทั้งหมด (ที่พบ) เสร็จสิ้น.");
                time.sleep(2)

                _log(f"กำลังคลิกปุ่ม 'พิมพ์รายงาน' ใน Pop-up (XPath: {print_report_in_modal_button_xpath})")
                final_print_button_element = wait.until(
                    EC.visibility_of_element_located((By.XPATH, print_report_in_modal_button_xpath)))
                _log(
                    f"   ปุ่ม visible: {final_print_button_element.is_displayed()}, HTML: {final_print_button_element.get_attribute('outerHTML')}")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
                                      final_print_button_element);
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", final_print_button_element)
                _log("คลิกปุ่ม 'พิมพ์รายงาน' ใน Pop-up แล้ว (ด้วย JavaScript click โดยตรง).")
            except Exception as e_popup:
                _log(f"!!! POPUP HANDLING: Error: {e_popup} !!!");
                _log(traceback.format_exc())
                try:
                    driver.save_screenshot(os.path.join(download_path, "po_popup_handling_error.png"))
                except:
                    pass
                if driver: driver.quit(); return None
        except Exception as e_step3_main:
            _log(f"!!! Error ในขั้นตอนที่ 3 (คลิกปุ่มหลัก หรือครอบ Pop-up): {e_step3_main} !!!");
            _log(traceback.format_exc())
            try:
                driver.save_screenshot(os.path.join(download_path, "po_step3_main_error.png"))
            except:
                pass
            if driver: driver.quit(); return None

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 4: จัดการ Notification "กระดิ่ง" และรอไฟล์ดาวน์โหลด
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 4: เริ่มกระบวนการตรวจสอบ Notification 'กระดิ่ง'...")

        # --- กำหนดค่าและ XPath สำหรับขั้นตอนที่ 4 ---
        # เวลาสูงสุดที่จะรอรายงาน (เป็นวินาที)
        # ให้เวลา 5 นาที (300 วินาที) เผื่อรายงานมีขนาดใหญ่
        NOTIFICATION_TIMEOUT_SECONDS = 300
        # ความถี่ในการตรวจสอบ (เป็นวินาที)
        POLLING_INTERVAL_SECONDS = 15

        # XPath ที่ได้จากการวิเคราะห์
        bell_icon_to_click_xpath = "//div[@id='notification']//a[contains(@class, 'fa-bell')]"
        notification_panel_xpath = "//div[contains(@class, 'dropdownNotification') and contains(@class, 'showNotification')]"
        download_trigger_item_xpath = (
            "//div[contains(@class, 'notificationItem')]"
            "[.//h3[contains(text(), 'รายงานใบสั่งซื้อ') and contains(text(), 'พร้อมดาวน์โหลดแล้ว')]]"
        )
        # XPath สำหรับใช้คลิกปิด หากจำเป็น (คลิกที่พื้นที่ว่างๆ)
        body_element_xpath = "//body"

        # --- เริ่ม Polling Loop ---
        _log(
            f"จะตรวจสอบ Notification ทุกๆ {POLLING_INTERVAL_SECONDS} วินาที เป็นเวลาสูงสุด {NOTIFICATION_TIMEOUT_SECONDS} วินาที")
        start_time = time.time()
        download_triggered = False

        while time.time() - start_time < NOTIFICATION_TIMEOUT_SECONDS:
            try:
                _log("...กำลังตรวจสอบ Notification...")
                # 1. คลิกที่ไอคอนกระดิ่งเพื่อเปิด Panel
                # ใช้ try-except เผื่อว่า Panel เปิดค้างอยู่แล้วคลิกซ้ำจะเกิด Error
                try:
                    bell_icon = short_wait.until(EC.element_to_be_clickable((By.XPATH, bell_icon_to_click_xpath)))
                    driver.execute_script("arguments[0].click();", bell_icon)
                    _log("   คลิกไอคอนกระดิ่งแล้ว")
                except ElementClickInterceptedException:
                    _log("   ไอคอนกระดิ่งถูกบัง, อาจมี Panel อื่นเปิดอยู่. จะลองคลิกที่ Body เพื่อปิด")
                    driver.find_element(By.XPATH, body_element_xpath).click()
                    time.sleep(1)
                    bell_icon = short_wait.until(EC.element_to_be_clickable((By.XPATH, bell_icon_to_click_xpath)))
                    driver.execute_script("arguments[0].click();", bell_icon)
                    _log("   คลิกไอคอนกระดิ่งอีกครั้งสำเร็จ")
                except Exception as e_click_bell:
                    _log(f"   Warning: ไม่สามารถคลิกกระดิ่งได้ในรอบนี้: {e_click_bell}")
                    # ข้ามไปรอบถัดไป
                    time.sleep(POLLING_INTERVAL_SECONDS)
                    continue

                # 2. รอให้ Panel แสดงผล และค้นหารายการที่ต้องการ
                try:
                    # รอ Panel เปิด
                    wait.until(EC.visibility_of_element_located((By.XPATH, notification_panel_xpath)))
                    _log("   Notification panel ปรากฏขึ้นแล้ว")

                    # ค้นหารายการดาวน์โหลด (ใช้ find_elements แบบไม่รอ เพราะอาจจะยังไม่มา)
                    report_items = driver.find_elements(By.XPATH, download_trigger_item_xpath)

                    if report_items:
                        _log(f"   เจอรายการ 'รายงานใบสั่งซื้อ พร้อมดาวน์โหลดแล้ว' จำนวน {len(report_items)} รายการ!")
                        # คลิกที่รายการแรกที่เจอ (ซึ่งควรจะเป็นอันล่าสุด)
                        report_to_click = report_items[0]
                        driver.execute_script("arguments[0].scrollIntoView(true);", report_to_click)
                        time.sleep(0.5)
                        report_to_click.click()
                        _log("   คลิกที่รายการแจ้งเตือนเพื่อเริ่มการดาวน์โหลดแล้ว!")
                        download_triggered = True
                        break  # ออกจาก while loop เพราะสั่งดาวน์โหลดแล้ว
                    else:
                        _log("   ยังไม่พบรายการรายงานที่พร้อมดาวน์โหลดในรอบนี้")
                        # คลิกที่ Body เพื่อปิด Panel เตรียมสำหรับรอบต่อไป
                        try:
                            driver.find_element(By.XPATH, body_element_xpath).click()
                            time.sleep(0.5)
                        except:
                            pass  # ถ้าคลิกไม่ได้ก็ไม่เป็นไร

                except TimeoutException:
                    _log("   Timeout: Notification panel ไม่ปรากฏขึ้นหลังคลิกกระดิ่ง. จะลองใหม่ในรอบถัดไป")
                    # อาจจะลองคลิก body เพื่อปิด element ที่อาจจะบังอยู่
                    try:
                        driver.find_element(By.XPATH, body_element_xpath).click()
                        time.sleep(0.5)
                    except:
                        pass

            except Exception as e_poll:
                _log(f"!!! Error ระหว่างการ Polling: {e_poll}. จะลองใหม่ในรอบถัดไป !!!")

            # หน่วงเวลาก่อนตรวจสอบรอบถัดไป
            if not download_triggered:
                _log(f"--- รออีก {POLLING_INTERVAL_SECONDS} วินาที ---")
                time.sleep(POLLING_INTERVAL_SECONDS)

        # --- ตรวจสอบผลลัพธ์หลังออกจาก Loop ---
        if not download_triggered:
            _log(f"!!! ล้มเหลว: หมดเวลา {NOTIFICATION_TIMEOUT_SECONDS} วินาทีแล้ว แต่ยังไม่พบรายงานให้ดาวน์โหลด !!!")
            try:
                driver.save_screenshot(os.path.join(download_path, "po_notification_timeout_error.png"))
            except:
                pass
            if driver: driver.quit()
            return None

        _log("ขั้นตอนที่ 4 สำเร็จ (สั่งดาวน์โหลดจาก Notification แล้ว).")

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 5: รอไฟล์ดาวน์โหลดให้เสร็จสมบูรณ์ และเปลี่ยนชื่อ
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 5: รอไฟล์ดาวน์โหลดให้เสร็จสมบูรณ์...")
        # ให้เวลาดาวน์โหลดสูงสุด 2 นาที
        DOWNLOAD_WAIT_TIMEOUT = 120
        wait_start_time = time.time()
        final_filepath = None

        while time.time() - wait_start_time < DOWNLOAD_WAIT_TIMEOUT:
            # ค้นหาไฟล์ Excel ที่ไม่ใช่ไฟล์ชั่วคราว (ใช้ Pattern ใหม่ที่ถูกต้อง)
            xlsx_files = glob.glob(os.path.join(download_path, "purchaseOrder_report_export_*.xlsx"))

            # ตรวจสอบว่าไม่มีไฟล์ .crdownload ที่ชื่อใกล้เคียงกัน (สำหรับ Chrome)
            crdownload_files = glob.glob(os.path.join(download_path, "*.crdownload"))

            if xlsx_files and not crdownload_files:
                downloaded_file = xlsx_files[0]
                # อาจจะรออีกนิดหน่อยเพื่อให้แน่ใจว่าไฟล์เขียนเสร็จสมบูรณ์
                time.sleep(2)
                _log(f"ตรวจพบไฟล์ที่ดาวน์โหลดเสร็จแล้ว: {downloaded_file}")

                # เปลี่ยนชื่อไฟล์
                final_filepath_target = os.path.join(download_path, desired_file_name)
                try:
                    os.rename(downloaded_file, final_filepath_target)
                    _log(f"เปลี่ยนชื่อไฟล์เป็น: {final_filepath_target}")
                    final_filepath = final_filepath_target
                    break  # สำเร็จ ออกจาก loop
                except Exception as e_rename:
                    _log(f"!!! Error ขณะเปลี่ยนชื่อไฟล์: {e_rename} !!!")
                    # ถ้าเปลี่ยนชื่อไม่ได้ ก็คืนชื่อเดิมไปก่อน
                    final_filepath = downloaded_file
                    break

            time.sleep(1)  # รอ 1 วินาทีก่อนเช็คใหม่

        if not final_filepath:
            _log(f"!!! ล้มเหลว: หมดเวลารอไฟล์ดาวน์โหลด ({DOWNLOAD_WAIT_TIMEOUT} วินาที) !!!")
            try:
                driver.save_screenshot(os.path.join(download_path, "po_file_download_timeout_error.png"))
            except:
                pass
            if driver: driver.quit()
            return None

        _log("🎉🎉🎉 ดาวน์โหลดและเปลี่ยนชื่อไฟล์รายงานใบสั่งซื้อสำเร็จ! 🎉🎉🎉")
        # เมื่อทุกอย่างสำเร็จ ให้ return path ของไฟล์
        return final_filepath

    except TimeoutException as te:
        _log(f"!!! TimeoutException (Overall Function Level): {te} !!!")
        if driver:
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_OVERALL_timeout_error.png"))
            except:
                pass
            if driver: driver.quit()
        return None
    except Exception as e:
        _log(f"!!! Fatal Error (Overall Function Level): {e} !!!")
        _log(traceback.format_exc())
        if driver:
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_OVERALL_fatal_error.png"))
            except:
                pass
            if driver: driver.quit()
        return None
    finally:
        # ทำให้แน่ใจว่า driver ถูก quit ถ้า session ยังอยู่ และยังไม่มีการ quit ก่อนหน้าในกรณี error
        # ถ้าต้องการให้ browser เปิดค้างหลัง 'NOTIFICATION_HANDLING_NEXT' อาจจะต้องปรับ logic นี้
        if driver and driver.session_id:
            _log("Ensuring WebDriver is quit in finally block (if not already).")
            try:
                driver.quit()
            except Exception as e_final_quit:
                _log(f"Error during final quit: {e_final_quit}")
        else:
            _log("WebDriver already quit or not initialized in finally.")
        _log("Function finished.")

    _log("!!! UNEXPECTED: Reached end of function. Should have returned earlier. !!!")
    return None


if __name__ == '__main__':
    print("=" * 30)
    print("  Testing auto_downloader.py (PEAK PO - Steps 1-3 & Prep for Notification)  ")
    print("=" * 30)

    test_peak_user = "sirichai.c@zubbsteel.com"
    test_peak_pass = "Zubb*2013"
    test_target_business = "บจ. บิซ ฮีโร่ (สำนักงานใหญ่)"

    if "YOUR_PEAK_EMAIL_HERE" in test_peak_user or \
            "YOUR_PEAK_PASSWORD_HERE" in test_peak_pass or \
            "ชื่อกิจการของคุณที่นี่" in test_target_business:
        print("\n!!! กรุณาแก้ไข test_peak_user, test_peak_pass, และ test_target_business ก่อนรันทดสอบ !!!\n")
    else:
        current_script_dir = os.getcwd()
        test_save_dir_for_artifacts = os.path.join(current_script_dir, "peak_po_notification_test_artifacts")
        if not os.path.exists(test_save_dir_for_artifacts):
            os.makedirs(test_save_dir_for_artifacts)


        def standalone_logger(message):
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            print(f"[{timestamp}][StandaloneTestLogger] {message}")


        print(f"Username: {test_peak_user}")
        print(f"Target Business: {test_target_business}")
        print(f"Artifacts will be saved to: {test_save_dir_for_artifacts}")

        result_status = download_peak_purchase_order_report(
            username=test_peak_user,
            password=test_peak_pass,
            target_business_name_to_select=test_target_business,
            save_directory=test_save_dir_for_artifacts,
            log_callback=standalone_logger
        )
        if result_status == "NOTIFICATION_HANDLING_NEXT":
            print(f"\n[StandaloneTestLogger] ขั้นตอนการสั่งประมวลผลรายงานสำเร็จ!")
            print(f"วันพรุ่งนี้: พัฒนาส่วนการตรวจสอบ Notification 'กระดิ่ง' และดาวน์โหลดไฟล์.")
        else:
            print(f"\n[StandaloneTestLogger] การดำเนินการไม่สำเร็จ หรือมีข้อผิดพลาด. Result: {result_status}")
            print(f"กรุณาตรวจสอบ log และไฟล์ screenshot (ถ้ามี) ในโฟลเดอร์ '{test_save_dir_for_artifacts}'")

    print("=" * 30)
    print("  Finished Testing auto_downloader.py (PEAK PO - Prep for Notification)  ")
    print("=" * 30)
