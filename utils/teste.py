from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time

# Configurar as opções do Chrome para abrir em modo anônimo
chrome_options = Options()
chrome_options.add_argument("--incognito")
chrome_options.add_experimental_option("detach", True)

# Inicializar o driver do Chrome
driver = webdriver.Chrome(options=chrome_options)
driver.maximize_window()
wait = WebDriverWait(driver, 10)

# Abrir a página de login do Microsoft
driver.get("https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=151&id=264960&wreply=https%3a%2f%2fwww.bing.com%2fsecure%2fPassport.aspx%3fedge_suppress_profile_switch%3d1%26requrl%3dhttps%253a%252f%252fwww.bing.com%252f%253fcc%253dbr%2526wlexpsignin%253d1%26sig%3d1C60459059DB624322E751E558B263A0%26nopa%3d2&wp=MBI_SSL&lc=1046&CSRFToken=f32eb264-e4cc-4043-a1c6-6a14b52819e6&cobrandid=c333cba8-c15c-4458-b082-7c8ce81bee85&aadredir=1&nopa=2")

# Preencher o campo de e-mail e fazer login
email_field = wait.until(EC.visibility_of_element_located((By.ID, "i0116")))
email_field.send_keys("kotakafly@gmail.com")
driver.find_element(By.ID, "idSIButton9").click()

# Preencher o campo de senha e fazer login
password_field = wait.until(EC.visibility_of_element_located((By.ID, "i0118")))
password_field.send_keys("n1a2u3a4k5")
driver.find_element(By.ID, "idSIButton9").click()

try:
    no_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "declineButton"))
    )
    no_button.click()
except TimeoutException:
    pass

try:
    yes_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "bnp_btn_accept"))
    )
    yes_button.click()
except TimeoutException:
    pass

# Esperar pelo login
wait.until(EC.url_contains("https://www.bing.com"))

# Pesquisar no Bing
search_field = wait.until(EC.visibility_of_element_located((By.ID, "sb_form_q")))
search_field.send_keys("Quando eu digo que deixei de te amar É porque eu te amo Quando eu digo que não quero mais você É porque eu te quero Eu tenho medo de te dar meu coração E confessar que eu estou em tuas mãos Mas não posso imaginar O que vai ser de mim Se eu te perder um dia")
search_field.submit()

# Deletar resultados de pesquisa
for _ in range(5):
    result = wait.until(EC.element_to_be_clickable((By.ID, "sb_form_q")))
    result.click()
    for _ in range(2):
        driver.find_element(By.ID, "sb_form_q").send_keys(Keys.DELETE)
    driver.find_element(By.ID, "sb_form_q").send_keys(Keys.ENTER)
    time.sleep(8)
