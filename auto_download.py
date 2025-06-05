import os
import time
import glob  # ยังคงเก็บไว้เผื่อใช้ในอนาคต หรือการลบไฟล์เก่า
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
import traceback  # สำหรับ traceback.format_exc()


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

    # --- ลบไฟล์เก่า (ส่วนนี้ยังคงมีประโยชน์) ---
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
    # --- สิ้นสุดการลบไฟล์เก่า ---

    chrome_options = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": download_path,  # ยังคงตั้งค่าเผื่อไว้
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    _log("Chrome options configured.")

    driver = None
    try:  # <--- TRY BLOCK หลักของฟังก์ชัน
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
        # very_short_wait = WebDriverWait(driver, 5) # อาจจะไม่ต้องใช้ถ้าไม่รอตาราง

        # --------------------------------------------------------------------
        # ขั้นตอนที่ 1: Login และ เลือกกิจการ
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 1: Login และ เลือกกิจการ...")
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
            actions.move_to_element(expense_menu_element).perform()
            _log("Hover บน 'รายจ่าย' แล้ว. รอ 1.5 วินาที...")
            time.sleep(1.5)

            wait.until(EC.visibility_of_element_located((By.XPATH, po_submenu_to_hover_xpath)))
            po_submenu_element_to_hover = wait.until(EC.element_to_be_clickable((By.XPATH, po_submenu_to_hover_xpath)))
            actions.move_to_element(po_submenu_element_to_hover).perform()
            _log("Hover บน 'ใบสั่งซื้อ' แล้ว. รอ 1.5 วินาที...")
            time.sleep(1.5)

            wait.until(EC.visibility_of_element_located((By.XPATH, view_all_po_actual_link_xpath)))
            view_all_link_element = wait.until(EC.element_to_be_clickable((By.XPATH, view_all_po_actual_link_xpath)))
            driver.execute_script("arguments[0].click();", view_all_link_element)
            _log("คลิก 'ดูทั้งหมด' แล้ว.")

            _log("รอการนำทางไปยังหน้า PO...")
            try:
                long_wait.until(lambda d: (("/expense/po" in d.current_url.lower() and (
                            "po" in d.title.lower() or "ใบสั่งซื้อ" in d.title.strip())) or (
                                                       "ใบสั่งซื้อ" in d.title.strip() and not "/income" in d.current_url.lower())))
                is_on_po_page = True  # สมมติว่าถ้า lambda ผ่าน คือมาถึงหน้า PO
                if not ("/expense/po" in driver.current_url.lower() and (
                        "ใบสั่งซื้อ" in driver.title.strip() or "po" in driver.title.lower())):  # ตรวจสอบซ้ำ
                    _log("ตรวจสอบซ้ำ: Lambda ผ่าน แต่ URL/Title ไม่ตรงเป๊ะ ลองเช็ค element ยืนยัน")
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
                        f"!!! ล้มเหลว: ไม่ได้อยู่ที่หน้า PO ที่ถูกต้อง. URL: {driver.current_url}, Title: {driver.title.strip()} !!!")
                    if driver: driver.quit()
                    return None
            except TimeoutException:
                _log(
                    f"!!! ล้มเหลว: Timeout ขณะรอการยืนยันหน้า PO. URL: {driver.current_url}, Title: {driver.title} !!!")
                if driver: driver.quit()
                return None
        except Exception as e_nav:  # จับ error ทั้งหมดในการนำทางเมนู
            _log(f"!!! ล้มเหลว: Error ระหว่างการนำทางผ่านเมนู: {e_nav} !!!")
            _log(traceback.format_exc())
            try:
                driver.save_screenshot(os.path.join(download_path, "po_menu_nav_error.png"))
            except:
                pass
            if driver: driver.quit()
            return None
        _log("ขั้นตอนที่ 2 เสร็จสิ้น.")

        # ขั้นตอนที่ 3: คลิกปุ่ม "พิมพ์รายงาน" และจัดการ Pop-up
        # --------------------------------------------------------------------
        _log("ขั้นตอนที่ 3: คลิกปุ่ม 'พิมพ์รายงาน' และจัดการ Pop-up...")

        # (Comment out การรอตาราง PO โหลด ชั่วคราว - ถ้ายังไม่ได้เอาออก)

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
        # เพิ่ม not(ancestor::div[contains(@class,'secondary')]) เพื่อให้แน่ใจว่าไม่ใช่ปุ่ม "ยกเลิก" ถ้ามันมีข้อความคล้ายกัน
        # (จาก HTML ที่เคยเห็น ปุ่มยกเลิกอยู่ใน <div class="secondary">)

        _log(f"กำลังคลิกปุ่ม 'พิมพ์รายงาน' ใน Pop-up (XPath: {print_report_in_modal_button_xpath})")

        try:  # Try block สำหรับการคลิกปุ่มหลัก และครอบ Pop-up handling ทั้งหมด
            _log(f"กำลังค้นหาปุ่ม 'พิมพ์รายงาน' (หลัก) ด้วย XPath: {print_report_main_button_xpath}")
            main_print_button = wait.until(EC.element_to_be_clickable((By.XPATH, print_report_main_button_xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", main_print_button)
            time.sleep(0.5)
            main_print_button.click()
            _log("คลิกปุ่ม 'พิมพ์รายงาน' (หลัก) สำเร็จแล้ว.")

            # -----[ POPUP HANDLING BLOCK ]-----
            try:
                # 1. รอให้ปุ่มปรากฏ (visible) เท่านั้น
                _log(" - รอให้ปุ่ม 'พิมพ์รายงาน' ใน Pop-up ปรากฏ (visible)...")
                final_print_button_element = wait.until(
                    EC.visibility_of_element_located((By.XPATH, print_report_in_modal_button_xpath)))
                _log(f"   ปุ่ม visible: {final_print_button_element.is_displayed()}")
                _log(f"   HTML ของปุ่ม: {final_print_button_element.get_attribute('outerHTML')}")

                # 2. Scroll ปุ่มให้อยู่ใน view
                driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});",
                                      final_print_button_element)
                time.sleep(0.5)  # ให้เวลานิดหน่อยหลัง scroll

                # 3. คลิกด้วย JavaScript click โดยตรง
                _log("   พยายามคลิกปุ่มด้วย JavaScript click โดยตรง...")
                driver.execute_script("arguments[0].click();", final_print_button_element)

                _log("คลิกปุ่ม 'พิมพ์รายงาน' ใน Pop-up แล้ว (ด้วย JavaScript click โดยตรง).")

                _log("รอ 10 วินาที เพื่อให้การดาวน์โหลดเริ่มขึ้น และให้คุณมีเวลาสังเกต...")
                time.sleep(10)

                _log("--- EXITING POPUP HANDLING TRY BLOCK SUCCESSFULLY ---")
                _log("ขั้นตอนที่ 1, 2, 3 (Pop-up actions) สำเร็จสมบูรณ์.")
                _log("เบราว์เซอร์จะยังคงเปิดอยู่สักครู่. ปิด Manual หรือรอ Timeout ของ Test.")
                return "POPUP_ACTIONS_COMPLETED_SUCCESSFULLY_AND_DOWNLOAD_SHOULD_START"

            except TimeoutException as e_timeout_final_button:
                _log(f"!!! Timeout ขณะรอปุ่ม 'พิมพ์รายงาน' ใน Pop-up ให้ visible: {e_timeout_final_button}")
                try:
                    driver.save_screenshot(
                        os.path.join(download_path, "peak_po_final_print_button_visible_timeout.png"))
                except:
                    pass
                raise
            except Exception as e_click_final_button:
                _log(f"!!! Error อื่นๆ ขณะพยายามคลิกปุ่ม 'พิมพ์รายงาน' ใน Pop-up: {e_click_final_button}")
                try:
                    driver.save_screenshot(os.path.join(download_path, "peak_po_final_print_button_click_error.png"))
                except:
                    pass
                raise

            except Exception as e_popup:  # จับ error ทั้งหมดใน POPUP HANDLING BLOCK
                _log(f"!!! POPUP HANDLING: Error: {e_popup} !!!")
                _log(traceback.format_exc())
                try:
                    driver.save_screenshot(os.path.join(download_path, "po_popup_handling_error.png"))
                except:
                    pass
                if driver: driver.quit()
                return None
            # -----[ END OF POPUP HANDLING BLOCK ]-----

        except Exception as e_step3_main:  # จับ error ทั้งหมดในการคลิกปุ่มหลัก หรือครอบ Pop-up handling
            _log(f"!!! Error ในขั้นตอนที่ 3 (คลิกปุ่มหลัก หรือครอบ Pop-up): {e_step3_main} !!!")
            _log(traceback.format_exc())
            try:
                driver.save_screenshot(os.path.join(download_path, "po_step3_main_error.png"))
            except:
                pass
            if driver: driver.quit()
            return None
        # ไม่ต้องมีส่วนที่ 4 ในตอนนี้

    # --- EXCEPT และ FINALLY ของ TRY BLOCK หลัก ---
    except TimeoutException as te:
        _log(f"!!! TimeoutException (Overall Function Level): {te} !!!")
        if driver:
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_OVERALL_timeout_error.png"))
            except:
                pass
            driver.quit()
        return None
    except Exception as e:
        _log(f"!!! Fatal Error (Overall Function Level): {e} !!!")
        _log(traceback.format_exc())
        if driver:
            try:
                driver.save_screenshot(os.path.join(download_path, "peak_po_OVERALL_fatal_error.png"))
            except:
                pass
            driver.quit()
        return None
    finally:
        if driver and driver.session_id: # ตรวจสอบว่า session ยัง active ก่อน quit
            try:
                _log("Ensuring WebDriver is quit in finally block (if not already).")
                driver.quit()
            except Exception as e_quit_finally:
                _log(f"Error during quit in finally: {e_quit_finally}")
        else:
            _log("WebDriver already quit or not initialized.")
        _log("Function finished.")

    _log("!!! UNEXPECTED: Reached end of function outside try/finally blocks. Should have returned earlier. !!!")
    return None  # Fallback return


# --- ส่วน if __name__ == '__main__': สำหรับทดสอบ ---
if __name__ == '__main__':
    print("=" * 30)
    print("  Testing auto_downloader.py (PEAK PO - Steps 1-3 Test)  ")  # อัปเดตชื่อ Test
    print("=" * 30)

    test_peak_user = "sirichai.c@zubbsteel.com"
    test_peak_pass = "Zubb*2013"
    test_target_business = "บจ. บิซ ฮีโร่ (สำนักงานใหญ่)"
    # desired_filename ไม่ได้ใช้ใน test นี้ เพราะเรายังไม่ดาวน์โหลด
    # test_desired_filename = "BizHero_PO_Report_Latest.xlsx"

    if "YOUR_PEAK_EMAIL_HERE" in test_peak_user or \
            "YOUR_PEAK_PASSWORD_HERE" in test_peak_pass or \
            "ชื่อกิจการของคุณที่นี่" in test_target_business:
        print("\n!!! กรุณาแก้ไข test_peak_user, test_peak_pass, และ test_target_business ก่อนรันทดสอบ !!!\n")
    else:
        current_script_dir = os.getcwd()
        test_save_dir_for_artifacts = os.path.join(current_script_dir, "peak_po_s123_test_artifacts")
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
            # desired_file_name=test_desired_filename, # ไม่ต้องส่งถ้ายังไม่ใช้
            log_callback=standalone_logger
        )
        if result_status == "POPUP_ACTIONS_COMPLETED_SUCCESSFULLY_AND_DOWNLOAD_SHOULD_START":
            print(f"\n[StandaloneTestLogger] การดำเนินการใน Pop-up (ขั้นตอนที่ 1-3) และการคลิกเพื่อเริ่มดาวน์โหลด สำเร็จ!")
            print(f"เบราว์เซอร์ควรจะยังเปิดอยู่ และไฟล์ควรจะเริ่มดาวน์โหลด.")
            input("ตรวจสอบเบราว์เซอร์สำหรับการดาวน์โหลด แล้วกด Enter เพื่อปิดสคริปต์และเบราว์เซอร์...") # <<<< เพิ่ม input()
        else:
            print(f"\n[StandaloneTestLogger] การดำเนินการ (ขั้นตอนที่ 1-3) ไม่สำเร็จ หรือมีข้อผิดพลาด. Result: {result_status}")
            print(f"กรุณาตรวจสอบ log และไฟล์ screenshot (ถ้ามี) ในโฟลเดอร์ '{test_save_dir_for_artifacts}'")
            # อาจจะ input ที่นี่ด้วยเพื่อให้มีเวลาดูเบราว์เซอร์ถ้ามี error
            # input("พบข้อผิดพลาด กด Enter เพื่อปิดสคริปต์และเบราว์เซอร์...")

    print("=" * 30)
    print("  Finished Testing auto_downloader.py (PEAK PO - Steps 1-3 Test)  ")
    print("=" * 30)