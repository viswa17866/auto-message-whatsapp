import os
import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

GROUP_NAME = "wishes grp" 
PROJECT_FOLDER = os.path.dirname(__file__)
EXCEL_FILE = os.path.join(PROJECT_FOLDER, "Students.xlsx")
PROFILE_DIR = r"C:\Users\VISWA\ChromeWhatsAppProfile"

try:
    events = pd.read_excel(EXCEL_FILE)
except PermissionError:
    print(f"ERROR: Close '{EXCEL_FILE}' if open in Excel.")
    exit()
except FileNotFoundError:
    print(f"ERROR: '{EXCEL_FILE}' not found in {PROJECT_FOLDER}")
    exit()


def normalize_date(v):
    try:
        dt = pd.to_datetime(v, dayfirst=True)
        return dt.strftime("%d-%m")
    except Exception:
        s = str(v).strip()
        if len(s) == 5 and s[2] == '-':
            return s
        return s

events['Date_norm'] = events['Date'].apply(normalize_date)
today = datetime.now().strftime("%d-%m")
today_events = events[events['Date_norm'] == today]

if today_events.empty:
    print("No events today. Exiting...")
    exit()
final_message = "\n".join(today_events['Message'].astype(str).tolist())
chrome_options = Options()
chrome_options.add_argument(f'--user-data-dir={PROFILE_DIR}')
chrome_options.add_argument(r'--profile-directory=Default')
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-extensions")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 30)


driver.get("https://web.whatsapp.com/")
wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true" and @role="textbox"]')))


search_box = driver.find_element(By.XPATH, '//div[@contenteditable="true" and @role="textbox"]')
search_box.clear()
search_box.send_keys(GROUP_NAME)
time.sleep(2)
search_box.send_keys(Keys.ENTER)

message_box = wait.until(EC.presence_of_element_located((
    By.XPATH, '//footer//div[@contenteditable="true"][@data-tab="10"]'
)))

message_box.click()
message_box.send_keys(final_message)
time.sleep(0.5)


send_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Send"]')))
send_button.click()

print("✅ Message sent successfully!")
time.sleep(2)  
driver.quit()
