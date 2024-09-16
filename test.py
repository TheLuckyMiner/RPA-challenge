import pandas as pd
from time import sleep

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

chrome_options = webdriver.ChromeOptions()
service = Service(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

try:
    # Вход на сайт
    driver.get("https://www.rpachallenge.com/")
    start_btn = driver.find_element('xpath', '//button[@class="waves-effect col s12 m12 l12 btn-large uiColorButton"]')
    start_btn.click()

    # Открытие файла Excel
    excel_file = 'challenge.xlsx'
    df = pd.read_excel(excel_file)
    row_count = len(df)

    # Заполнение полей сайта путем ввода данных из таблицы
    for i in range(row_count):
        first_name_label = driver.find_element("xpath",'//input[@ng-reflect-name="labelFirstName"]')
        first_name_label.send_keys(df.loc[i, "First Name"])

        last_name_label = driver.find_element("xpath",'//input[@ng-reflect-name="labelLastName"]')   
        last_name_label.send_keys(df.loc[i, "Last Name "])

        company_name_label = driver.find_element("xpath",'//input[@ng-reflect-name="labelCompanyName"]')
        company_name_label.send_keys(df.loc[i, "Company Name"])

        role_label = driver.find_element("xpath",'//input[@ng-reflect-name="labelRole"]')
        role_label.send_keys(df.loc[i, "Role in Company"])

        address_label = driver.find_element("xpath",'//input[@ng-reflect-name="labelAddress"]')
        address_label.send_keys(df.loc[i, "Address"])

        email_label = driver.find_element("xpath",'//input[@ng-reflect-name="labelEmail"]')
        email_label.send_keys(df.loc[i, "Email"])

        phone_label = driver.find_element("xpath", '//input[@ng-reflect-name="labelPhone"]')    
        phone_label.send_keys(str(df.loc[i, "Phone Number"]))

        # Нажатие кнопки "submit" после ввода всех полей
        driver.find_element("xpath", '//input[@type="submit"]').click()

    # Создание скриншота выполненного задания
    driver.get_screenshot_as_file('Result.png') 
    sleep(3)  

except:
    print("Что-то пошло не так, попробуйте снова")
