import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
import time

import dotenv
import os

dotenv.load_dotenv()

browser = webdriver.Chrome()
browser.implicitly_wait(1)

def checkelem(selector):
    if len(browser.find_elements(By.XPATH, selector)) == 1:
        return True
    return False

def get_excel_tab(excel_name: str, excel_sheet: str):
    workbook = openpyxl.load_workbook(filename=f"{excel_name}.xlsx")
    worksheet = workbook[excel_sheet]
    return worksheet.iter_rows()


listexcel = get_excel_tab('otch', 'стейдж')
browser.get(os.getenv('site'))
browser.find_element(By.XPATH, "//input [@type='email']").send_keys(os.getenv('login'))
browser.find_element(By.XPATH, "//input [@id='login-password']").send_keys(os.getenv('pass'))
browser.find_element(By.XPATH, "//button [text() = 'Войти']").click()
#"//button [text() = 'Войти']" только для кнопки текст = наше значение
#"//span[contains(text(),'В очереди')]" нажмется, когда в кнопке есть наше значение

for num, stroka in enumerate(listexcel, start=1):
    browser.get(stroka[0].value)
    browser.find_element(By.XPATH, "//span[contains(text(),'новый')]").click() #кнопка генерировать новый
    while True:
        time.sleep(1)
        browser.refresh()
        if not checkelem("//span[contains(text(),'В очереди')]") and not checkelem("//span[contains(text(),'Запущен')]"):
            browser.find_element(By.XPATH, "//tbody/tr[1]/td[2]/div[1]/span[1]/a[1]/span[1]").click() #последний отчет
            print(num, ' ', stroka[0].value)
            break 
    time.sleep(1)
browser.quit()
