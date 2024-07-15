import os
import os.path
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from dotenv import load_dotenv
from PIL import Image
import fitz  # PyMuPDF
import google.generativeai as genai
import pandas as pd
import shutil
import traceback


load_dotenv()


class Bot:
    TMP_DIR = f"{os.path.abspath(os.getcwd())}/tmp"
    DOWNLOAD_DIR = f"{os.path.abspath(os.getcwd())}/data"
    driver = None
    ai_client = None
    csv = None

    def __init__(self):
        os.makedirs(self.DOWNLOAD_DIR, exist_ok=True)
        os.makedirs(self.TMP_DIR, exist_ok=True)
        for file_name in os.listdir(self.TMP_DIR):
            file_path = os.path.join(self.TMP_DIR, file_name)
            if os.path.isfile(file_path):
                os.remove(file_path)

        # init Gemini
        genai.configure(api_key=os.environ['GEMINI_API_KEY'])
        self.ai_client = genai.GenerativeModel('gemini-1.5-flash')

        # init browser
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--no-sandbox')
        options.add_argument('--headless')
        options.add_argument("--window-size=1440,900")
        options.add_argument("--lang=en")
        options.add_experimental_option("prefs", {
            "download.default_directory": self.TMP_DIR,
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

        # load csv
        self.csv = self.xls_to_list_of_dicts("main.xls")

    def run(self):
        row_number = -1
        for row in self.csv:
            row_number += 1
            if row.get("Processed") and row.get("Processed").strip() == "yes":
                print(f"Skipping row #{row_number + 1}")
                continue

            print(f"Processing row #{row_number + 1} ...")
            try:
                company_name = row.get("Name")
                register_number = row.get("HRB")
                if not company_name or not register_number:
                    continue

                self.driver.get("https://www.handelsregister.de/rp_web/normalesuche.xhtml")
                time.sleep(3)

                self.driver.execute_script(f"document.getElementById('form:schlagwoerter').innerText = '{company_name}'")
                self.driver.execute_script(f"document.getElementById('form:registerNummer').value = '{register_number}'")

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
            except Exception as e:
                traceback.print_exc()

            new_row = self.csv[row_number].copy()
            try:
                for file in self.get_downloded_files():
                    if not file.endswith(".tiff") and not file.endswith(".pdf"):
                        os.remove(file)
                        continue

                    shutil.copy(f"{self.TMP_DIR}/{file}", f"{self.DOWNLOAD_DIR}/{file}")
                    new_row["Document Link"] = f"file://{self.DOWNLOAD_DIR}/{file}"

                    # enrich data
                    shareholders = self.extract_shareholders_from_file(f"{self.TMP_DIR}/{file}")
                    os.remove(f"{self.TMP_DIR}/{file}")
                    i = 1
                    for s in shareholders:
                        new_row[f"Shareholder-{i}"] = s["name"]
                        new_row[f"Shareholder-{i} %"] = s["percentage"]
                        new_row[f"Shareholder-{i} DB"] = s["date_of_birth"]
                        new_row[f"Shareholder-{i} age"] = s["age"]
                        i += 1
                processed = "yes"
            except Exception as e:
                processed = "error"
                traceback.print_exc()

            # rewrite csv
            new_row["Processed"] = processed
            self.update_xls_by_index("main.xls", row_number, new_row)

        self.driver.close()

    def get_downloded_files(self):
        all_entries = os.listdir(self.TMP_DIR)
        return [entry for entry in all_entries if os.path.isfile(os.path.join(self.TMP_DIR, entry))]

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
            1. name - (str) shareholder name (extract only name and ignore any additional details);
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
            traceback.print_exc()
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

    @staticmethod
    def xls_to_list_of_dicts(file_path):
        df = pd.read_excel(file_path)
        return df.to_dict(orient='records')

    def update_xls_by_index(self, file_path, line_index, update_dict):
        df = pd.read_excel(file_path)
        for key, value in update_dict.items():
            if key not in df.columns:
                df[key] = pd.NA
            df.at[line_index, key] = value
        df.to_excel(file_path, index=False, engine="openpyxl")


bot = Bot()
bot.run()
