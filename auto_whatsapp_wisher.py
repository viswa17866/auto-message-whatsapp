import os
import time
import random
import json
import pandas as pd
from pathlib import Path
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openai import OpenAI
from dotenv import load_dotenv

# ===== LOAD ENVIRONMENT VARIABLES =====
load_dotenv()

# ===== OPENAI API SETUP =====
api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key)

# ===== HISTORY FILE =====
HISTORY_FILE = Path("sent_wishes.json")

def load_history():
    """Load sent wishes history from file."""
    if HISTORY_FILE.exists():
        return json.loads(HISTORY_FILE.read_text())
    return {}

def save_history(history):
    """Save sent wishes history to file."""
    HISTORY_FILE.write_text(json.dumps(history, indent=2))

def generate_ai_message(name, event_type, gender=None, role=None):
    """Generate a unique AI wish for the given name, event type, gender, and role."""
    
    # Define style preferences based on gender/role
    if role and role.lower() == "staff":
        tone = "make it formal, respectful, and appreciative of their contribution"
    elif gender and gender.lower() in ["female", "girl", "woman"]:
        tone = "make it warm, kind, and cheerful with a gentle tone"
    elif gender and gender.lower() in ["male", "boy", "man"]:
        tone = "make it energetic, friendly, and motivating"
    else:
        tone = "make it friendly and heartfelt"

    styles = [
        "mention a positive quality about them",
        "include a short blessing or wish for their future",
        "add a line appreciating their presence in our lives",
        "make it inspiring and uplifting",
        "include a kind compliment"
    ]
    chosen_style = random.choice(styles)

    prompt = (
        f"Write a WhatsApp {event_type} wish for {name}. "
        f"The person is a {role if role else 'member'} and is {gender if gender else 'unspecified gender'}. "
        f"{tone}. Also, {chosen_style}. "
        f"Mention something about {datetime.now().strftime('%A')} "
        f"or the month of {datetime.now().strftime('%B')}. "
        f"No emojis. Two to three sentences."
    )

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are an assistant that writes unique and heartfelt WhatsApp wishes."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.9
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"⚠ OpenAI Error: {e}")
        fallback_messages = [
            f"Happy {event_type} {name}! Wishing you joy, good health, and wonderful moments.",
            f"Wishing you a wonderful {event_type}, {name}! May your day be filled with laughter and love.",
            f"Cheers to you, {name}, on your {event_type}! Hope it’s as amazing as you are.",
            f"Happy {event_type}, {name}! Here’s to more success, happiness, and cherished memories ahead.",
            f"Many happy returns on your {event_type}, {name}! Wishing you endless smiles and blessings."
        ]
        return random.choice(fallback_messages)


def generate_unique_message(name, event_type):
    """Generate a wish that hasn't been used before for this person."""
    history = load_history()

    # Ensure person has a history entry
    if name not in history:
        history[name] = []

    while True:
        msg = generate_ai_message(name, event_type)
        if msg not in history[name]:
            history[name].append(msg)
            save_history(history)
            return msg

# ===== CONFIG =====
GROUP_NAME = "Wishing group"
PROJECT_FOLDER = os.path.dirname(__file__)
EXCEL_FILE = os.path.join(PROJECT_FOLDER, "Sample.xlsx")
PROFILE_DIR = r"C:\Users\VISWA\ChromeWhatsAppProfile"

# ===== LOAD EXCEL =====
try:
    events = pd.read_excel(EXCEL_FILE)
except PermissionError:
    print(f"ERROR: Close '{EXCEL_FILE}' if open in Excel.")
    exit()
except FileNotFoundError:
    print(f"ERROR: '{EXCEL_FILE}' not found in {PROJECT_FOLDER}")
    exit()

# ===== DATE NORMALIZATION =====
# ===== DATE NORMALIZATION =====
def normalize_date(v):
    try:
        dt = pd.to_datetime(v, dayfirst=True)
        return dt.strftime("%d-%m")
    except Exception:
        s = str(v).strip()
        if len(s) == 5 and s[2] == '-':
            return s
        return s

events['Date_norm'] = events['DOB'].apply(normalize_date)
today = datetime.now().strftime("%d-%m")
today_events = events[events['Date_norm'] == today]

if today_events.empty:
    print("No events today. Exiting...")
    exit()

# ===== GENERATE ALL MESSAGES =====
all_messages = []
for _, row in today_events.iterrows():
    name = row['Name']
    role = row.get('Role', None)
    gender = row.get('Gender', None)
    event_type = row.get('Type', 'birthday')  # Optional column
    ai_message = generate_ai_message(name, event_type, gender, role)
    all_messages.append(ai_message)


# ===== OPEN WHATSAPP WEB =====
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

# ===== SEARCH GROUP =====
search_box = driver.find_element(By.XPATH, '//div[@contenteditable="true" and @role="textbox"]')
search_box.clear()
search_box.send_keys(GROUP_NAME)
time.sleep(2)
search_box.send_keys(Keys.ENTER)

# ===== GENERATE ALL MESSAGES =====
all_messages = []
for _, row in today_events.iterrows():
    name = row['Name']
    event_type = row.get('Type', 'birthday')
    ai_message = generate_unique_message(name, event_type)
    all_messages.append(ai_message)

final_message = "\n".join(all_messages)

# ===== SEND MESSAGE =====
message_box = wait.until(EC.presence_of_element_located((By.XPATH, '//footer//div[@contenteditable="true"][@data-tab="10"]')))
message_box.click()
message_box.send_keys(final_message)
time.sleep(0.5)

send_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Send"]')))
send_button.click()
time.sleep(3)

print(f"✅ Sent to group '{GROUP_NAME}':\n{final_message}")
