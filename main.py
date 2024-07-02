import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from dotenv import load_dotenv


load_dotenv()


class Bot:
    driver = None

    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        options.add_argument('--no-sandbox')

        download_dir = f"{os.path.abspath(os.getcwd())}/data"
        os.makedirs(download_dir, exist_ok=True)
        options.add_experimental_option("prefs", {
            "download.default_directory": download_dir,
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
        self.driver.close()


bot = Bot()
bot.run()
