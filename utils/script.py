import pyautogui, time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Configurar as opções do Chrome para abrir em modo anônimo
chrome_options = Options()
chrome_options.add_argument("--incognito")
chrome_options.add_experimental_option("detach", True)

for i in range(1):
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    driver.get("https://login.live.com/login.srf?wa=wsignin1.0&rpsnv=151&id=264960&wreply=https%3a%2f%2fwww.bing.com%2fsecure%2fPassport.aspx%3fedge_suppress_profile_switch%3d1%26requrl%3dhttps%253a%252f%252fwww.bing.com%252f%253fcc%253dbr%2526wlexpsignin%253d1%26sig%3d1C60459059DB624322E751E558B263A0%26nopa%3d2&wp=MBI_SSL&lc=1046&CSRFToken=f32eb264-e4cc-4043-a1c6-6a14b52819e6&cobrandid=c333cba8-c15c-4458-b082-7c8ce81bee85&aadredir=1&nopa=2")
    time.sleep(1)
    pyautogui.moveTo(x=688, y=555)
    pyautogui.typewrite("henrycasterfly@gmail.com")
    pyautogui.press("Enter")
    time.sleep(1)
    pyautogui.typewrite("n1a2u3a4k5")
    pyautogui.press("Enter")
    time.sleep(2)
    pyautogui.leftClick()
    pyautogui.moveTo(x=598, y=559)
    time.sleep(4)
    pyautogui.leftClick()
    driver.get("https://bing.com/search")
    pyautogui.typewrite("Deixa acontecer naturalmente Eu não quero ver você chorar Deixa que o amor encontre a gente Nosso caso vai eternizar")
    pyautogui.press("Enter")
    time.sleep(8)
    for i in range(15):
        pyautogui.moveTo(x=221, y=176)
        pyautogui.leftClick()
        pyautogui.press("Delete")
        pyautogui.press("Delete")
        pyautogui.press("Enter")
        time.sleep(8)












