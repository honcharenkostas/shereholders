import os
import time
import os.path
import pickle
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from dotenv import load_dotenv


load_dotenv()


class Bot:
    GOOGLE_API_SCOPES = ['https://www.googleapis.com/auth/drive.file']
    GOOGLE_API_CLIENT_SECRET_FILE = 'client_secret.json'
    DOWNLOAD_DIR = f"{os.path.abspath(os.getcwd())}/data"
    GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")
    driver = None
    google_service = None

    def __init__(self):
        self.google_service = self.google_authenticate()

        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--no-sandbox')

        os.makedirs(self.DOWNLOAD_DIR, exist_ok=True)
        options.add_experimental_option("prefs", {
            "download.default_directory": self.DOWNLOAD_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        service = Service(executable_path="./chromedriver")
        self.driver = webdriver.Chrome(
            service=service,
            options=options
        )
        self.driver.maximize_window()

    def run(self):
        company_name = "inoage GmbH"
        register_number = "29795"
        court = "Dresden"
        self.driver.get("https://www.handelsregister.de/rp_web/normalesuche.xhtml")
        time.sleep(3)

        self.driver.execute_script(f"document.getElementById('form:schlagwoerter').innerText = '{company_name}'")
        self.driver.execute_script(f"document.getElementById('form:registerNummer').value = '{register_number}'")
        # self.driver.execute_script(f"document.getElementById('form:registergericht_input').value = '{court}'")

        # self.driver.find_element(By.ID, "form").submit()

        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        btn = self.driver.find_element(By.ID, "form:btnSuche")
        btn.click()
        time.sleep(5)
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        for a in self.driver.find_elements(By.CSS_SELECTOR, "a.dokumentList"):
            if a.find_element(By.CSS_SELECTOR, "span").text.strip() == "DK":
                a.click()
                break
        time.sleep(5)
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)

        self.driver.execute_script('''
        elements = document.querySelectorAll('li span');
        targetElement = Array.from(elements).find(element => element.innerText === "Documents on register number");
        targetElement.closest('li').querySelector('.ui-tree-toggler').click()
        ''')
        time.sleep(1)
        self.driver.execute_script('''
        elements = document.querySelectorAll('li span');
        targetElement = Array.from(elements).find(element => element.innerText === "List of shareholders");
        targetElement.closest('li').querySelector('.ui-tree-toggler').click()
        ''')
        time.sleep(1)

        for el in self.driver.find_elements(By.CSS_SELECTOR, 'li span'):
            if el.text.strip().startswith("List of shareholders – "):
                el.click()
                break
        time.sleep(1)
        self.driver.execute_script("document.querySelectorAll(\"input[name='dk_form:radio_dkbuttons']\")[1].click()")
        time.sleep(1)
        self.driver.execute_script("document.querySelector(\"button[type=submit]\").click()")
        time.sleep(10)

        for file in self.get_downloded_files():
            self.upload_file_to_google_drive(file)
            os.remove(f"{self.DOWNLOAD_DIR}/{file}")

        self.driver.close()

    def get_downloded_files(self):
        all_entries = os.listdir(self.DOWNLOAD_DIR)
        return [entry for entry in all_entries if os.path.isfile(os.path.join(self.DOWNLOAD_DIR, entry))]

    def google_authenticate(self):
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.GOOGLE_API_CLIENT_SECRET_FILE, self.GOOGLE_API_SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('drive', 'v3', credentials=creds)
        return service

    def upload_file_to_google_drive(self, file_path):
        file_path = f"{self.DOWNLOAD_DIR}/{file_path}"
        file_metadata = {
            'name': os.path.basename(file_path),
            'parents': [self.GOOGLE_DRIVE_FOLDER_ID] if self.GOOGLE_DRIVE_FOLDER_ID else []
        }
        media = MediaFileUpload(file_path, resumable=True)
        file = self.google_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()

        # https://drive.google.com/file/d/15sh561qMKauH1rx6v1iPN1sR2MC16rVO
        print(f'File ID: {file.get("id")}')


bot = Bot()
bot.run()
