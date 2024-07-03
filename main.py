import os
import os.path
import time
import json
from datetime import datetime
import pickle
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from dotenv import load_dotenv
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
import google.generativeai as genai


load_dotenv()


class Bot:
    GOOGLE_API_SCOPES = ['https://www.googleapis.com/auth/drive.file']
    GOOGLE_API_CLIENT_SECRET_FILE = 'client_secret.json'
    DOWNLOAD_DIR = f"{os.path.abspath(os.getcwd())}/data"
    GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")
    driver = None
    google_service = None
    ai_client = None

    def __init__(self):
        self.google_service = self.google_authenticate()

        genai.configure(api_key=os.environ['GEMINI_API_KEY'])
        self.ai_client = genai.GenerativeModel('gemini-1.5-flash')

        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--no-sandbox')
        # options.add_argument('--headless')
        options.add_argument("--window-size=1440,900")
        options.add_argument("--lang=en")

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

        for file in self.get_downloded_files():
            os.remove(f"{self.DOWNLOAD_DIR}/{file}")

    def run(self):
        company_name = "Minessa Medical Deutschland GmbH"  # "inoage GmbH"
        register_number = "774122"  # "29795"
        court = "Winnenden"  # "Dresden"
        self.driver.get("https://www.handelsregister.de/rp_web/normalesuche.xhtml")
        time.sleep(3)
        self.driver.save_screenshot("1.png")

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

        self.driver.save_screenshot("2.png")

        for phrase in ["Dokumente zur Registernummer", "Documents on register number"]:
            try:
                self.driver.execute_script(f'''
                elements = document.querySelectorAll('li span');
                targetElement = Array.from(elements).find(element => element.innerText === "{phrase}");
                targetElement.closest('li').querySelector('.ui-tree-toggler').click()
                ''')
                break
            except:
                pass
        time.sleep(1)

        for phrase in ["Liste der Gesellschafter", "List of shareholders"]:
            try:
                self.driver.execute_script(f'''
                elements = document.querySelectorAll('li span');
                targetElement = Array.from(elements).find(element => element.innerText === "{phrase}");
                targetElement.closest('li').querySelector('.ui-tree-toggler').click()
                ''')
                break
            except:
                pass
        time.sleep(1)

        element_found = False
        for el in self.driver.find_elements(By.CSS_SELECTOR, 'li span'):
            for phrase in ["Liste der Gesellschafter -", "List of shareholders –"]:
                try:
                    if el.text.strip().startswith(phrase):
                        element_found = True
                        el.click()
                        break
                except:
                    pass
            if element_found:
                break

        time.sleep(1)
        self.driver.execute_script("document.querySelectorAll(\"input[name='dk_form:radio_dkbuttons']\")[1].click()")
        time.sleep(1)
        self.driver.execute_script("document.querySelector(\"button[type=submit]\").click()")
        time.sleep(10)

        for file in self.get_downloded_files():
            if not file.endswith(".tiff") and not file.endswith(".pdf"):
                continue
            google_drive_file_url = self.upload_file_to_google_drive(file)
            print("google_drive_file_url", google_drive_file_url)
            shareholders = self.extract_shareholders_from_file(f"{self.DOWNLOAD_DIR}/{file}")
            for row in shareholders:
                print(row)

            os.remove(f"{self.DOWNLOAD_DIR}/{file}")

        self.driver.save_screenshot("3.png")
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

        return f"https://drive.google.com/file/d/{file.get('id')}"

    def extract_shareholders_from_file(self, file_path):
        # tiff/pdf to jpg
        images = self.pdf_to_images(file_path)
        merged_image = self.merge_images(images)
        if file_path.endswith(".tiff"):
            merged_image_path = f"{file_path[:-5]}.jpg"
        elif file_path.endswith(".pdf"):
            merged_image_path = f"{file_path[:-4]}.jpg"
        else:
            raise Exception(f"Invalid file extension. File path: {file_path}")
        self.save_image(merged_image, merged_image_path)

        try:
            # upload file to Gemini
            sample_file = genai.upload_file(path=merged_image_path, display_name="Image of shareholders pdf file")

            # extract data from file
            prompt_text = '''
            Please analyse the image
            and return a json of shareholders as list of objects with next fields per object:
            1. name - (str) shareholder name;
            2. percentage - (int) percentage of shareholdings;
            3. date_of_birth - (str) date of birth of shareholder in format 'day.month.year';
            4. age - (int) age of shareholder in years;
    
            Also please note:
            1. This image made from pdf documents of shareholders of the company.
            2. The image can contain data in a table, make sure you extract data correct from the table.
            3. If you don't see percentage - calculate it.
            4. No need to explain, no need to wrap in "```json". Return plain json only.
            '''.strip()

            response = self.ai_client.generate_content([sample_file, prompt_text])
            result = json.loads(response.text)
        except Exception as e:
            print(f"{e}")
            result = []

        os.remove(merged_image_path)

        return result

    @staticmethod
    def pdf_to_images(file_path):
        pdf_document = fitz.open(file_path)
        images = []

        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            # Use a higher resolution for better quality
            zoom = 2.0  # You can adjust this value to get the desired resolution
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(img)

        return images

    @staticmethod
    def merge_images(images):
        widths, heights = zip(*(img.size for img in images))

        total_height = sum(heights)
        max_width = max(widths)

        merged_image = Image.new('RGB', (max_width, total_height))

        y_offset = 0
        for img in images:
            merged_image.paste(img, (0, y_offset))
            y_offset += img.height

        return merged_image

    @staticmethod
    def save_image(image, file_path):
        image.save(file_path, 'JPEG', quality=100)  # Save with maximum quality


bot = Bot()
bot.run()
