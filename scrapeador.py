from seleniumwire import webdriver
import selenium
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
chrome_options = webdriver.ChromeOptions()
import time
#chrome_options.add_argument('--headless')

path = './chromedriver'
driver = webdriver.Chrome(path)
delay = 10

driver2 = driver.get('https://consopt.www8.receita.fazenda.gov.br/consultaoptantes')
time.sleep(10)
input_text = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'form-control')))
time.sleep(1)
input_text.send_keys('22135112000192')
button_text = driver.find_element(by=By.CLASS_NAME, value='btn-verde')
time.sleep(1)
button_text.click()
optante = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'spanValorVerde')))
print(optante.text)

