# -*- coding: cp1251 -*-

import openpyxl as xl
import re
import xlsxwriter
from dateutil import parser
import datetime, time
import pandas as pd


file_isotp_db = r'C:\Users\vanik\PycharmProjects\handlers_sg'



# Создание общего словаря ТП-------------------------------------------
def update_phase2_dbs():
    wb_phase2 = xl.load_workbook('БД ТП ФАЗА 1, 2.xlsx')
    sheet_phase2_TP = wb_phase2['Сводная информация']
    sheet_phase2_ISO = wb_phase2['TP']

