# -*- coding: utf-8 -*-
import os
import selenium, requests
from selenium import webdriver

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time

options = webdriver.ChromeOptions()
options.add_argument('headless')

driver = webdriver.Chrome(r"C:\Users\ignatenkoia\Documents\python\act_work\weldlog_summary\chromedriver.exe")
driver.get('https://agpz.sgaz.pro/faces/zeroLevelOOP')

print('Запустился в фоновом режиме.')

id_box = driver.find_element_by_id('loginInput')
id_box.send_keys('IgnatenkoIA')

passw_box = driver.find_element_by_id('passInput')
passw_box.send_keys('Buyfntyrj22')

login_button = driver.find_element_by_id('enter_btn')
login_button.click()
driver.maximize_window()
# Выбор строительство
time.sleep(3)
driver.find_element_by_xpath('//div[@id="p7zpnq:1:k6gtgh_1"]').click()
time.sleep(5)

# Инспекции
time.sleep(3)
driver.find_element_by_id('klqyd3_3').click()
time.sleep(3)

driver.find_element_by_xpath('//button[@class="select_many_choice-list-button"]').click()
time.sleep(3)
driver.find_element_by_xpath('//div[@title="RFI"]').click()

time.sleep(3)
driver.find_element_by_xpath('//div[@id="hyj3sh_1"]').click()

element = driver.find_element_by_xpath('//div[@data-value="35"]')

element.location_once_scrolled_into_view

time.sleep(1)
driver.find_element_by_xpath('//div[@data-value="35"]').click()
time.sleep(1)
driver.find_element_by_xpath('//div[@data-value="36"]').click()
time.sleep(1)
driver.find_element_by_xpath('//div[@data-value="30"]').click()

time.sleep(1)
date_box = driver.find_element_by_id('C8zl006::content').clear()
driver.find_element_by_id('C8zl006::content').click()

time.sleep(5)
driver.find_element_by_xpath('//div[@id="k6gtgh_10"]').click()
time.sleep(5)
driver.find_element_by_xpath('//div[@id="hyj3sh_5-dropdown-target"]').click()

time.sleep(1)
element = driver.find_element_by_xpath('//div[@title="Технологические трубопроводы"]')
element.location_once_scrolled_into_view
time.sleep(1)
driver.find_element_by_xpath('//div[@title="Технологические трубопроводы"]').click()

driver.find_element_by_id('k6gtgh_15').click()
time.sleep(5)
# Выгрузка инспекций
driver.find_element_by_xpath('//div[@id="wwyw8f"]').click()

dir_files = r'C:\Users\ignatenkoia\Downloads\\'
dir_destination = r'C:\Users\ignatenkoia\Documents\python\act_work\weldlog_summary\\'


def get_new_file():
    get_files = os.listdir(dir_files)
    date_list = [[x, os.path.getctime(r'C:\Users\ignatenkoia\Downloads\\' + x)] for x in get_files]
    sort_date_list = sorted(date_list, key=lambda x: x[1], reverse=True)
    itog_file = sort_date_list[0][0]
    date_last_file = os.path.getctime(r'C:\Users\ignatenkoia\Downloads\\' + sort_date_list[0][0])
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


file = get_new_file()

os.replace(dir_files + get_new_file(), dir_destination + 'Журнал заявок 1 фаза + спутники.xlsx')

driver.close()
print('Выгрузил инспекции 1 фаза + спутники')
print('Сохранил журнал куда нужно.')

#  -------------2 phase
driver = webdriver.Chrome(r"C:\Users\ignatenkoia\Documents\python\act_work\weldlog_summary\chromedriver.exe")
driver.get('https://agpz.sgaz.pro/faces/zeroLevelOOP')

id_box = driver.find_element_by_id('loginInput')
id_box.send_keys('IgnatenkoIA')

passw_box = driver.find_element_by_id('passInput')
passw_box.send_keys('Buyfntyrj22')

login_button = driver.find_element_by_id('enter_btn')
login_button.click()
driver.maximize_window()
# Выбор строительство
time.sleep(3)
driver.find_element_by_xpath('//div[@id="p7zpnq:1:k6gtgh_1"]').click()
time.sleep(5)
# Предписания
driver.find_element_by_id('klqyd3_4').click()
time.sleep(3)
driver.find_element_by_id('C9mgv59_1::content').click()
driver.find_element_by_xpath("//option[text()='предписания']").click()
time.sleep(3)
driver.find_element_by_xpath('//div[@id="wwyw8f_3"]').click()
print('Выгрузил предписания')

# Инспекции
time.sleep(3)
driver.find_element_by_id('klqyd3_3').click()
time.sleep(3)

driver.find_element_by_xpath('//div[@id="vb8d4u:2:cub9is"]').click()
time.sleep(13)

driver.find_element_by_xpath('//button[@class="select_many_choice-list-button"]').click()
time.sleep(3)
driver.find_element_by_xpath('//div[@title="RFI"]').click()

time.sleep(3)
driver.find_element_by_xpath('//div[@id="hyj3sh_1"]').click()

element = driver.find_element_by_xpath('//div[@data-value="35"]')

element.location_once_scrolled_into_view

# Выбор статус инспекции СК

time.sleep(1)
driver.find_element_by_xpath('//div[@data-value="35"]').click()
time.sleep(1)
driver.find_element_by_xpath('//div[@data-value="36"]').click()
time.sleep(1)
driver.find_element_by_xpath('//div[@data-value="30"]').click()

time.sleep(1)
date_box = driver.find_element_by_id('C8zl006::content').clear()
driver.find_element_by_id('C8zl006::content').click()

time.sleep(5)
driver.find_element_by_xpath('//div[@id="k6gtgh_10"]').click()
time.sleep(3)
driver.find_element_by_xpath('//div[@id="hyj3sh_5-dropdown-target"]').click()
time.sleep(1)
driver.find_element_by_xpath('//div[@title="Антикоррозийная защита"]').click()
time.sleep(1)
element = driver.find_element_by_xpath('//div[@title="Технологические трубопроводы"]')
element.location_once_scrolled_into_view
time.sleep(1)
driver.find_element_by_xpath('//div[@title="Технологические трубопроводы"]').click()

driver.find_element_by_id('k6gtgh_15').click()
time.sleep(5)
# Выгрузка инспекций
driver.find_element_by_xpath('//div[@id="wwyw8f"]').click()

print('Выгрузил инспекции')

dir_files = r'C:\Users\ignatenkoia\Downloads\\'
dir_destination = r'C:\Users\ignatenkoia\Documents\python\act_work\weldlog_summary\\'


def get_new_file():
    get_files = os.listdir(dir_files)
    date_list = [[x, os.path.getctime(r'C:\Users\ignatenkoia\Downloads\\' + x)] for x in get_files]
    sort_date_list = sorted(date_list, key=lambda x: x[1], reverse=True)
    itog_file = sort_date_list[0][0]
    date_last_file = os.path.getctime(r'C:\Users\ignatenkoia\Downloads\\' + sort_date_list[0][0])
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

get_files = os.listdir(dir_files)

date_list = [[x, os.path.getctime(r'C:\Users\ignatenkoia\Downloads\\' + x)] for x in get_files]
sort_date_list = sorted(date_list, key=lambda x: x[1], reverse=True)

itog_files = [sort_date_list[0][0], sort_date_list[1][0]]
for f in itog_files:
    if 'Журнал заявок' in f:
        os.replace(dir_files + f, dir_destination + 'Журнал заявок.xlsx')
    else:
        os.replace(dir_files + f, dir_destination + 'Реестр уведомлений.xlsx')

print('Сохранил журналы куда нужно.')

print(itog_files)
