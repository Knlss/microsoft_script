from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import openpyxl


class Chrome():
    def __init__(self):
        self.actions = ActionChains()
        chrome_options = Options()

        self.spreadsheet = openpyxl.load_workbook('caminho/para/o/arquivo/exemplo.xlsx').active
        self.rMax = self.spreadsheet["E6"]

        self.textToSearch = self.spreadsheet["E3"]

        chrome_options.add_experimental_option("detach", True)
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.maximize_window()

        self.wait = WebDriverWait(self.driver, 10)

    def obtainEmails(self, spreadsheet, rMax):
        for linha in spreadsheet.iter_rows(min_row=2, max_row=rMax, min_col=2, max_col=3):

            email = linha[0].value
            senha = linha[1].value
            
            return email, senha


    def openBing(self, account, password):
        self.actions.key_down(Keys.CONTROL).key_down(Keys.SHIFT).send_keys("n").key_up(Keys.CONTROL).key_up(Keys.SHIFT).perform()

        self.driver.get("https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=151&id=264960&wreply=https%3a%2f%2fwww.bing.com%2fsecure%2fPassport.aspx%3fedge_suppress_profile_switch%3d1%26requrl%3dhttps%253a%252f%252fwww.bing.com%252f%253fcc%253dbr%2526wlexpsignin%253d1%26sig%3d1C60459059DB624322E751E558B263A0%26nopa%3d2&wp=MBI_SSL&lc=1046&CSRFToken=f32eb264-e4cc-4043-a1c6-6a14b52819e6&cobrandid=c333cba8-c15c-4458-b082-7c8ce81bee85&aadredir=1&nopa=2")
        email_field = self.wait.until(EC.visibility_of_element_located((By.ID, "i0116")))
        email_field.send_keys(account)
        self.driver.find_element(By.ID, "idSIButton9").click()

        password_field = self.wait.until(EC.visibility_of_element_located((By.ID, "i0118")))
        password_field.send_keys(password)
        self.driver.find_element(By.ID, "idSIButton9").click()

        try:
            no_button = self.wait.until(EC.element_to_be_clickable((By.ID, "declineButton")))
            no_button.click()
        except TimeoutException:
            pass

        try:
            yes_button = self.wait.until(EC.element_to_be_clickable((By.ID, "bnp_btn_accept")))
            yes_button.click()
        except TimeoutException:
            pass

        self.wait.until(EC.url_contains("https://www.bing.com"))

    def startSearch(self, textToSearch):
        search_field = self.wait.until(EC.visibility_of_element_located((By.ID, "sb_form_q")))
        search_field.send_keys(textToSearch)
        search_field.submit()

        for _ in range(5):
            result = self.wait.until(EC.element_to_be_clickable((By.ID, "sb_form_q")))
            result.click()
            for _ in range(2):
                self.driver.find_element(By.ID, "sb_form_q").send_keys(Keys.DELETE)
            self.driver.find_element(By.ID, "sb_form_q").send_keys(Keys.ENTER)
            time.sleep(8)

    def dayStreak(self):
        self.driver.get("https://rewards.bing.com/?ref=rewardspanel")

        elements = self.driver.find_elements(By.CLASS_NAME, 'ds-card-sec')
        for link in elements[:3]:
            if link.is_enabled():
                link.click()
                time.sleep(8)
                self.actions.key_down(Keys.CONTROL).send_keys("w").key_up(Keys.CONTROL).perform()
                time.sleep(2)

    def closeBing(self):
        self.actions.key_down(Keys.ALT).send_keys(Keys.F4).key_up(Keys.ALT).perform()

    def bingLoop(self):
        self.obtainEmails(self.spreadsheet, self.rMax)

        for account, password in range(5):
            self.openBing(account, password)
            self.startSearch(self.textToSearch)
            self.dayStreak()
            self.closeBing()
            time.sleep(3)