import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def setup_driver(download_path):
    print("Setting up the Chrome driver...")
    options = webdriver.ChromeOptions()

    # options.add_argument("--headles") 
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--start-maximized")
    options.add_argument('--log-level=3')
    prefs = {"download.default_directory": download_path}
    options.add_experimental_option("prefs", prefs)
    service = ChromeService(executable_path=ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    print("✅ Driver setup complete.")
    return driver

def search_and_navigate(driver, wait, symbol):
    try:
        print(f"\n🔎 Searching for symbol: {symbol}")
        search_input = wait.until(EC.element_to_be_clickable((By.ID, "autocomplete")))
        search_input.clear()
        search_input.send_keys(symbol)
        results = wait.until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "li.ui-menu-item")))
        original_tab = driver.current_window_handle
        results[0].click()
        print("✅ Clicked on the first result.")
        
        wait.until(EC.number_of_windows_to_be(2))
        new_tab = [window for window in driver.window_handles if window != original_tab][0]
        driver.switch_to.window(new_tab)
        print("🔄 Switched focus to the new tab.")
        
     
        wait.until(EC.presence_of_element_located((By.ID, "form1")))
        print("📄 Symbol details page is loading.")
        return True

    except (TimeoutException, IndexError):
        print(f"❌ Timed out or failed while processing {symbol}.")
        if len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
        return False

def click_excel_export(driver, wait):
    try:
        print("🖱️ Looking for the Excel Export button...")
        export_button_locator = (By.ID, "ContentPlaceHolder1_btnExport")
        
        export_button = wait.until(EC.presence_of_element_located(export_button_locator))

        print("📜 Scrolling the button into view...")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", export_button)
        
        wait.until(EC.element_to_be_clickable(export_button_locator)).click()
        
        print("✅ Excel Export button clicked successfully.")
        return True
    except TimeoutException:
        print("❌ Could not find or click the Excel Export button.")
        return False

def wait_for_download_complete(folder_path, timeout=30):
    print(f"⏳ Waiting up to {timeout} seconds for download to complete...")
    seconds_waited = 0
    while seconds_waited < timeout:
        if any(file.endswith('.crdownload') for file in os.listdir(folder_path)):
            time.sleep(1)
            seconds_waited += 1
        else:
            time.sleep(1)
            print("✅ Download finished.")
            return True
    print("❌ Download did not complete in time.")
    return False

def main():
    download_folder = os.path.join(os.getcwd(), "Faraborse_Downloads")
    os.makedirs(download_folder, exist_ok=True)
    symbols = ["اراد1101", "اراد1102"]
    driver = setup_driver(download_folder)
    wait = WebDriverWait(driver, 20)
    try:
        base_url = "https://www.ifb.ir/"
        driver.get(base_url)
        print(f"🌐 Opened main page: {base_url}")
        original_tab = driver.current_window_handle
        for symbol in symbols:
            if not search_and_navigate(driver, wait, symbol):
                driver.get(base_url)
                continue
            if not click_excel_export(driver, wait):
                driver.close()
                driver.switch_to.window(original_tab)
                continue
            wait_for_download_complete(download_folder)
            driver.close()
            driver.switch_to.window(original_tab)
            print(f"--- Finished processing {symbol} ---")
    finally:
        if driver:
            driver.quit()
            print("\n🌙 Process finished.")

if __name__ == "__main__":
    main()
