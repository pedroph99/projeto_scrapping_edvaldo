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
    p