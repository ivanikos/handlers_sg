# -*- coding: utf-8 -*-
import os
import selenium, requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
import time
from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time

with open('account_data.txt', 'r') as ad:
    ac_data = ad.readlines()

options = webdriver.ChromeOptions()

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.set_window_size(1920, 1200)
driver.get(f'{ac_data[0].strip()}')


id_box = driver.find_element(By.ID, 'loginInput')
id_box.send_keys(f'{ac_data[1].strip()}')

passw_box = driver.find_element(By.ID, 'passInput')
passw_box.send_keys(f'{ac_data[2].strip()}')


#  -------------2 phase

login_button = driver.find_element(By.ID, 'enter_btn')
login_button.click()
driver.maximize_window()
# Выбор строительство
time.sleep(3)
driver.find_element(By.XPATH, '/html/body/div[2]/form/div[2]/div[2]/div/div[2]/div/div[2]'
                              '/table/tbody/tr/td/table/tbody/tr/td[1]/div/div[2]/table/'
                              'tbody/tr/td[1]/div/a[2]/i').click()
time.sleep(5)
# Предписания
driver.find_element(By.XPATH, '//*[@id="C30efmo"]/a[5]').click()
time.sleep(5)
driver.find_element(By.XPATH, '//*[@id="C9mgv59_2::content"]').click()
driver.find_element(By.XPATH, "//option[text()='предписания']").click()
time.sleep(5)
driver.find_element(By.XPATH, '/html/body/div[1]/form/div[2]/div[2]/div/div[3]/div/div/div[2]/div/div/'
                              'div/div[2]/div/div/div/div[1]/div/div/div/div/div[1]/div[1]/table/tbody/'
                              'tr/td[5]/div/a/div/i').click()
print('Выгрузил предписания')

# Инспекции
time.sleep(3)
driver.find_element(By.XPATH, '/html/body/div[1]/form/div[2]/div[2]/div/div[3]/div/div/div[1]/div/a[4]').click()
time.sleep(3)

# driver.find_element_by_xpath('//div[@id="vb8d4u:2:cub9is"]').click()
# time.sleep(13)

driver.find_element(By.XPATH, '//button[@class="select_many_choice-list-button"]').click()
time.sleep(3)
driver.find_element(By.XPATH, '//div[@title="RFI"]').click()

time.sleep(3)
driver.find_element(By.XPATH, '//div[@id="hyj3sh_1"]').click()

element = driver.find_element(By.XPATH, '//div[@data-value="35"]')

element.location_once_scrolled_into_view

# Выбор статус инспекции СК

time.sleep(1)
driver.find_element(By.XPATH, f"/html/body/*/*/span[contains(text(), 'СК: Не принято')]").click()
# driver.find_element_by_xpath('//div[@data-value="12"]').click() /html/body/div[5]/div[21]/span
time.sleep(1)
driver.find_element(By.XPATH, f"/html/body/*/*/span[contains(text(), 'СК: Принято с замечаниями')]").click()
time.sleep(1)
driver.find_element(By.XPATH, '/html/body/div[5]/div[46]/span').click()
time.sleep(1)

date_box = driver.find_element(By.ID, 'C8zl006::content').clear()
driver.find_element(By.ID, 'C8zl006::content').click()

time.sleep(5)
driver.find_element(By.XPATH, '//*[@id="k6gtgh_11"]/a/i').click()
time.sleep(4)
driver.find_element(By.XPATH, '/html/body/div[1]/form/div[2]/div[2]/div/div[3]/div/div/div[2]/div/div/div/div[2]/div/'
                              'div[1]/div/div/div/div/div/div/table/tbody/tr/td/div/div[1]/div/div/div[2]/table/tbody/'
                              'tr/td/table/tbody/tr/td/div[2]/div/div[1]').click()
time.sleep(1)
driver.find_element(By.XPATH, '//div[@title="Антикоррозийная защита"]').click()
time.sleep(1)
element = driver.find_element(By.XPATH, '//div[@title="Технологические трубопроводы"]')
element.location_once_scrolled_into_view

time.sleep(1)
driver.find_element(By.XPATH, '//div[@title="Технологические трубопроводы"]').click()

driver.find_element(By.XPATH, '/html/body/div[1]/form/div[2]/div[2]/div/div[3]/div/div/div[2]/div/div/div/div[1]/'
                              'div/div/div/div/div[1]/div[1]/table/tbody/tr/td[26]/div/a/i').click()
time.sleep(5)
# Выгрузка инспекций
driver.find_element(By.XPATH, '/html/body/div[1]/form/div[2]/div[2]/div/div[3]/div/div/div[2]/div/div/div/div[1]/'
                              'div/div/div/div/div[1]/div[1]/table/tbody/tr/td[12]/div/a/div/i').click()



# dir_files = r'C:\Users\ignatenkoia\Downloads\\'
# dir_destination = r'C:\Users\ignatenkoia\Desktop\work\GIT_PROJECTS\handlers_sg\\'


# home laptop
dir_files = r'C:\Users\vanik\Downloads\\'
dir_destination = r'C:\Users\vanik\PycharmProjects\handlers_sg\Сводки\\'


def get_new_file():
    get_files = os.listdir(dir_files)
    date_list = [[x, os.path.getctime(dir_files + x)] for x in get_files]
    sort_date_list = sorted(date_list, key=lambda x: x[1], reverse=True)
    itog_file = sort_date_list[0][0]
    date_last_file = os.path.getctime(dir_files + sort_date_list[0][0])
    tim = time.time() - date_last_file
    if sort_date_list[0][0][-1] == 'x':
        if tim > 10:
            print('wait')
            time.sleep(5)
            return get_new_file()
        else:
            print('Done')
            print(itog_file)
            return itog_file
    else:
        print(itog_file)
        time.sleep(2)
        return get_new_file()


get_new_file()
driver.close()

print('Выгрузил инспекции')

get_files = os.listdir(dir_files)

date_list = [[x, os.path.getctime(dir_files + x)] for x in get_files]
sort_date_list = sorted(date_list, key=lambda x: x[1], reverse=True)

itog_files = [sort_date_list[0][0], sort_date_list[1][0]]
for f in itog_files:
    if 'Инспекции на' in f:
        os.replace(dir_files + f, dir_destination + 'Журнал заявок общий.xlsx')
    else:
        os.replace(dir_files + f, dir_destination + 'Реестр уведомлений.xlsx')

print('Сохранил журналы куда нужно.')

print(itog_files)


