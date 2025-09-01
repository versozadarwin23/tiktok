import string
from openpyxl import Workbook, load_workbook
import os
import requests
from bs4 import BeautifulSoup
import time
import random
import json
import concurrent.futures
import threading
import hashlib  # Import the hashlib library

xlsx_lock = threading.Lock()
console_lock = threading.Lock()
os.system("clear")

def random_device_model():
    # Expanded list of realistic device models
    models = [
        "SM-G998B", "iPhone13,4", "Pixel 5", "Mi 11", "OnePlus 9 Pro",
        "Samsung SM-A525F", "Google Pixel 6", "Xiaomi Redmi Note 10 Pro",
        "Huawei P40 Pro", "LG Velvet", "Sony Xperia 1 III", "Oppo Find X3 Pro",
        "vivo X60 Pro+", "Realme GT", "ASUS ROG Phone 5", "Lenovo Legion Phone Duel",
        "Nokia XR20", "Motorola Edge 20 Pro", "iPhone14,2", "iPhone14,3",
        "Samsung SM-S908U", "Google Pixel 7 Pro", "Xiaomi 12 Pro", "OnePlus 10 Pro"
    ]
    return random.choice(models)

def random_device_id():
    # Example placeholder:
    return hashlib.md5(os.urandom(16)).hexdigest()

def random_fingerprint():
    # Example placeholder:
    return hashlib.md5(os.urandom(16)).hexdigest()

# Assuming 'ua' is a list of user agents defined elsewhere in your project
# If not, you'll need to define it, e.g.:
ua = [
    'Mozilla/5.0 (Linux; Android 10) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.5304.105 Mobile Safari/537.36',
    'Mozilla/5.0 (iPhone; CPU iPhone OS 15_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.0 Mobile/15E148 Safari/604.1',
    # Add more user agents here
]

def generate_email():
    """Gumawa ng random username para sa harakirimail at bumalik ang email address at url."""
    rchjtrchjb = ''.join(random.choices(string.ascii_lowercase + string.digits, k=12))
    email_address = f"{rchjtrchjb}@harakirimail.com"
    drtyghbj5hgcbv = f"https://harakirimail.com/inbox/{rchjtrchjb}"
    return email_address, drtyghbj5hgcbv


def save_to_txt(filename, data):
    try:
        with open(filename, "a", encoding="utf-8") as f:
            f.write("|".join(data) + "\n")
    except Exception as e:
        print(f"\033[1;91m❗ Error saving to {filename}: {e}\033[0m")


def save_to_xlsx(filename, data):
    while True:
        with xlsx_lock:
            try:
                if os.path.exists(filename):
                    wb = load_workbook(filename)
                    ws = wb.active
                else:
                    wb = Workbook()
                    ws = wb.active
                    ws.append(
                        ["NAME", "USERNAME", "PASSWORD", "ACCOUNT LINK", "ACCESS TOKEN"])  # Added Access Token column
                ws.append(data)
                wb.save(filename)
                break
            except Exception as e:
                print(f"\033[1;91m❗ Error saving to {filename}: {e}. Retrying...\033[0m")
                time.sleep(RETRY_DELAY)


MAX_RETRIES = 3
RETRY_DELAY = 2
SUCCESS = "✅"
FAILURE = "❌"
INFO = "ℹ️"
WARNING = "⚠️"
LOADING = "⏳"


def load_names_from_file(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return [line.strip() for line in file.readlines() if line.strip()]
    except FileNotFoundError:
        print(f"\033[1;91m❗ Error: Name file not found at {file_path}. Please create it.\033[0m")
        return ["John", "Jane", "Doe", "Smith"]


def get_names(account_type, gender):
    male_first_names_file = "first_name.txt"
    last_names_file = "last_name.txt"

    male_first_names = load_names_from_file(male_first_names_file)
    last_names = load_names_from_file(last_names_file)

    if not male_first_names:
        male_first_names = ["Juan", "Pedro"]
    if not last_names:
        last_names = ["Dela Cruz", "Reyes"]

    firstname = random.choice(male_first_names)
    lastname = random.choice(last_names)
    return firstname, lastname


def generate_random_phone_number():
    random_number = str(random.randint(1000000, 9999999))
    third = random.randint(0, 4)
    forth = random.randint(1, 7)
    return f"9{third}{forth}{random_number}"


def generate_random_password():
    base = "Promises@"
    six_digit = str(random.randint(100000, 999999))
    return base + six_digit


def generate_user_details(account_type, gender, password=None):
    firstname, lastname = get_names(account_type, gender)
    year = random.randint(1978, 2001)
    date = random.randint(1, 28)
    month = random.randint(1, 12)
    if password is None:
        password = generate_random_password()
    phone_number = generate_random_phone_number()
    return firstname, lastname, date, year, month, phone_number, password


custom_password_base = None


def create_fbunconfirmed(account_num, account_type, gender, password=None, session=None):
    agent = random.choice(ua)
    global custom_password_base
    os.system("clear")
    email_address, drtyghbj5hgcbv = generate_email()
    if password is None:
        if custom_password_base:
            six_digit = str(random.randint(100000, 999999))
            password = custom_password_base + six_digit
        else:
            password = generate_random_password()

    firstname, lastname, date, year, month, phone_number, used_password = generate_user_details(account_type, gender,
                                                                                                password)

    def check_page_loaded(url, headers, current_session):
        retries = 0
        while retries < MAX_RETRIES:
            try:
                response = current_session.get(url, timeout=30, headers=headers)
                soup = BeautifulSoup(response.text, 'html.parser')
                form = soup.find("form")
                if form:
                    return form
                else:
                    print(f"{WARNING} Form not found on {url}, retrying... (Account #{account_num})")
            except requests.exceptions.RequestException as e:
                print(
                    f"{FAILURE} No internet or connection issue: {e}. Retrying in {RETRY_DELAY} seconds... (Account #{account_num})")
            except Exception as e:
                print(
                    f"{FAILURE} An unexpected error occurred: {e}. Retrying in {RETRY_DELAY} seconds... (Account #{account_num})")

            time.sleep(RETRY_DELAY)
            retries += 1
        print(f"{FAILURE} Failed to load page and find form after {MAX_RETRIES} retries. (Account #{account_num})")
        return None

    url = "https://m.facebook.com/reg?soft=hjk"
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",  #
        "Referer": "https://m.facebook.com/reg",  #
        "Connection": "keep-alive",  #
        "Accept-Language": "en-US,en;q=0.9",  #
        "X-FB-Connection-Type": "mobile.LTE",  #
        "X-FB-Device": random_device_model(),  #
        "X-FB-Device-ID": random_device_id(),  #
        "X-FB-Fingerprint": random_fingerprint(),  #
        "X-FB-Connection-Quality": "EXCELLENT",  #
        "X-FB-Net-HNI": "51502",  #
        "X-FB-SIM-HNI": "51502",  #
        "X-FB-HTTP-Engine": "Liger",  #
        'x-fb-connection-type': 'Unknown',  #
        'accept-encoding': 'gzip, deflate',  #
        'content-type': 'application/x-www-form-urlencoded',  #
        'x-fb-http-engine': 'Liger',  #
        'User-Agent': 'Mozilla/5.0 (Linux; Android 10; SM-G975F Build/QP1A.190711.020; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/80.0.3987.149 Mobile Safari/537.36 [FB_IAB/Orca-Android;FBAV/256.2.0.23.117;]',
        #
    }
    if session is None:
        session = requests.Session()

    form = check_page_loaded(url, headers, session)
    if not form:
        print(
            f"\033[1;91m{FAILURE} Could not load registration page or find form. Aborting attempt for account #{account_num}.\033[0m")
        return "FAILED_PAGE_LOAD"

    retries = 0
    while retries < MAX_RETRIES:
        try:
            response = session.get(url, headers=headers, timeout=10)
            soup = BeautifulSoup(response.text, "html.parser")
            form = soup.find("form")
            if form:
                break
        except requests.exceptions.RequestException as e:
            print(f"\033[1;91m{FAILURE} Error fetching form: {e}. Retrying... (Account #{account_num})\033[0m")
        except Exception as e:
            print(
                f"\033[1;91m{FAILURE} An unexpected error occurred while fetching form: {e}. Retrying... (Account #{account_num})\033[0m")
        time.sleep(RETRY_DELAY)
        retries += 1

    if not form:
        print(
            f"\033[1;91m{FAILURE} Failed to get registration form after retries. Aborting attempt for account #{account_num}.\033[0m")
        return "FAILED_FORM_FETCH"

    action_url = requests.compat.urljoin(url, form["action"]) if form.has_attr("action") else url
    inputs = form.find_all("input")
    data = {
        "firstname": firstname,
        "lastname": lastname,
        "birthday_day": str(date),
        "birthday_month": str(month),
        "birthday_year": str(year),
        "reg_email__": email_address,
        "sex": str(gender),
        "encpass": used_password,
        "submit": "Sign Up"
    }

    for inp in inputs:
        if inp.has_attr("name") and inp["name"] not in data:
            data[inp["name"]] = inp.get("value", "")

    try:
        response = session.post(action_url, headers=headers, data=data, timeout=30)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(
            f"\033[1;91m{FAILURE} Network error during submission: {e}. Cannot complete account creation for account #{account_num}.\033[0m")
        return "FAILED_SUBMISSION_NETWORK"
    except Exception as e:
        print(
            f"\033[1;91m{FAILURE} An unexpected error occurred during submission: {e}. Cannot complete account creation for account #{account_num}.\033[0m")
        return "FAILED_SUBMISSION_UNEXPECTED"

    with open(f"status.html", "w", encoding="utf-8") as file:
        file.write(response.text)

    if "c_user" not in session.cookies:
        print(
            f"\033[1;91m⚠️ Create Account Failed for account #{account_num}. No c_user cookie found. Try toggling airplane mode or use another email.\033[0m")
        time.sleep(3)
        return "FAILED_NO_C_USER"

    uid = session.cookies.get("c_user")
    profile_id = f'https://www.facebook.com/profile.php?id={uid}'

    cookie_dir = "/storage/emulated/0/cookie"
    os.makedirs(cookie_dir, exist_ok=True)
    cookie_file = os.path.join(cookie_dir, f"{uid}.json")
    cookie_names = ["c_user", "datr", "fr", "noscript", "sb", "xs"]
    cookies_data = {name: session.cookies.get(name, "") for name in cookie_names}
    try:
        with open(cookie_file, "w") as f:
            json.dump(cookies_data, f, indent=4)
    except IOError as e:
        print(f"\033[1;91m{FAILURE} Error saving cookies for account #{account_num}: {e}\033[0m")

    soup = BeautifulSoup(response.text, "html.parser")
    form_checkpoint = soup.find('form', action=lambda x: x and 'checkpoint' in x)
    if form_checkpoint:
        print(
            f"\033[1;91m⚠️ Created account #{account_num} blocked. Try toggling airplane mode or clearing Facebook Lite data.\033[0m")
        time.sleep(3)
        return "BLOCKED"

    jbkj = None
    retries = 0
    while retries < MAX_RETRIES * 5:
        try:
            dtryvghjuijhn = requests.get(drtyghbj5hgcbv, timeout=30)
            dtryvghjuijhn.raise_for_status()
            soup_mail = BeautifulSoup(dtryvghjuijhn.text, "html.parser")
            table = soup_mail.find("table", class_="table table-hover table-striped")
            if table:
                subject_link = table.find("tbody", id="mail_list_body").find("a")
                if subject_link:
                    subject_div = subject_link.find("div")
                    if subject_div:
                        subject = subject_div.get_text(strip=True)
                        if "is your confirmation code" in subject:
                            jbkj = subject.replace(" is your confirmation code", "")
                            if jbkj:
                                break
        except requests.exceptions.RequestException as e:
            print(f"\033[1;91m{FAILURE} Error fetching email for account #{account_num}: {e}. Retrying...\033[0m")
        except Exception as e:
            print(
                f"\033[1;91m{FAILURE} An unexpected error occurred while processing email for account #{account_num}: {e}. Retrying...\033[0m")

        time.sleep(5)
        retries += 1

    if not jbkj:
        print(f"\033[1;91m{FAILURE} Failed to get confirmation code for account #{account_num} after multiple attempts. Account might be unconfirmed.\033[0m")

    full_name = f"{firstname} {lastname}"
    with console_lock:
        print("\n\033[1;96m======================================\033[0m")
        print(f"\033[1;92m✅     Full Name: | {full_name} |\033[0m")
        print(f"\033[1;92m✅     Email: | {email_address} |\033[0m")
        print(f"\033[1;92m✅     Pass:  | {password} |\033[0m")
        print(f"\033[1;92m✅     Profile ID:  | {uid} |\033[0m")
        print(f"\033[1;92m✅     Code:  | {jbkj if jbkj else 'N/A (Code not found)'}\033[0m")
        print("\033[1;96m======================================\033[0m\n")

        filename_xlsx = "/storage/emulated/0/Acc_Created.xlsx"
        filename_txt = "/storage/emulated/0/Acc_created.txt"

        choice = input("💾 Do you want to save this account? (y/n): ").strip().lower()
        if choice == "y":
            data_to_save = [full_name, email_address, password, profile_id]
            save_to_xlsx(filename_xlsx, data_to_save)
            save_to_txt(filename_txt, data_to_save)
            print("✅ Account saved.")
        elif choice == "n":
            print("Account not saved.")
        else:
            print("Invalid input. Please enter 'y' or 'n'.")


def main():
    try:
        max_create_input = input("\033[93mEnter the maximum number of accounts to create: \033[0m").strip()
        max_create = int(max_create_input)
        if max_create <= 0:
            print("\033[1;91m❗ Please enter a positive number.\033[0m")
            return
    except ValueError:
        print("\033[1;91m❗ Invalid input. Please enter a number.\033[0m")
        return

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_create) as executor:
        futures = [
            executor.submit(create_fbunconfirmed, i, "personal", random.choice([1, 2]))
            for i in range(1, max_create + 1)
        ]
        # Just wait for them all to finish (no printing)
        concurrent.futures.wait(futures)

if __name__ == "__main__":
    while True:
        main()
