from seleniumwire import webdriver
import selenium
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver.v2 as uc
import time
#chrome_options.add_argument('--headless')

if __name__ == '__main__':
    
    driver = uc.Chrome()
    delay = 10
    driver2 = driver.get('https://consopt.www8.receita.fazenda.gov.br/consultaoptantes')
    time.sleep(3)
    input_text = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'form-control')))

    button_text = driver.find_element(by=By.CLASS_NAME, value='btn-verde')


    button_text.click()
    time.sleep(3)
    input_text = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'form-control')))
    input_text.send_keys('22359304000182')
    time.sleep(2)
    button_text = driver.find_element(by=By.CLASS_NAME, value='btn-verde')
    button_text.click()
    optante = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CLASS_NAME, 'spanValorVerde')))
    print(optante.text)

