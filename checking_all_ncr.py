import openpyxl as xl
import re, os
import xlsxwriter
from dateutil import parser
import datetime, time
import pandas as pd




wb_ncr = xl.load_workbook('Реестр уведомлений.xlsx')
sheet_ncr = wb_ncr['Предписания (Instructions)']
open_iso_ncr = []
open_ncr = [['Изометрия', 'стык', 'номер ncr', 'пункт', 'отметка о закрытии', 'сцеп изо+стык']]
for i in sheet_ncr['B4':'V5500']:
    number_ncr = str(i[0].value)
    mark_execution = str(i[16].value)
    notification_items = str(i[1].value)
    type_violation = str(i[5].value)
    content_remarks = str(i[6].value).replace(' ', '')
    content_remarks_iso = re.findall(r'\d-\d-\d-\d\d-\d\d\d-\s?\w*\+?-[0-9A-Z][0-9A-Z]-\d\d\d\d-\d\d\d', content_remarks)
    content_remarks_joints = re.findall(r'\s{1}[Ss]\s?\-?\d*.\d*|\s{1}F\s?\-?\d*.\d*', str(i[6].value))

    if 'ЗАО'in content_remarks:

        isometrics_1 = re.findall(
            r'\d-\d-\d-\d\d-\d\d\d-\w*-\d\w-\d\d\d\d-\d\d\d|\d-\d-\d-\d\d-\d\d\d-\w*-\d\d-\d\d\d\d-\d\d\d|\d-\d-\d-\d\d-\d\d\d-NHC3P\+-\d\d-\d\d\d\d-\d\d\d|'
            r'\d-\d-\d-\d\d-\d\d\d-NHC3\+-\d\d-\d\d\d\d-\d\d\d|\d-\d-\d-\d\d-\d\d\d-NHC4P\+-\d\d-\d\d\d\d-\d\d\d|\d-\d-\d-\d\d-\d\d\d-NHC5\+-\d\d-\d\d\d\d-\d\d\d|'
            r'\d-\d-\d-\d\d-\d\d\d-NHC4\+-\d\d-\d\d\d\d-\d\d\d',
            content_remarks.replace(' ', '').replace('\n', '').replace('Р', 'P').replace('С', 'C').strip())

        for i in isometrics_1:
            content_remarks = content_remarks.replace(i, '')

        wrong_joints_s = re.findall(r'S-\d*\.\d*RW|S-\d*\.\d*R1|S-\d*\.\d*R|S-\d*\.\d*|S-\d*,\d*RW|S-\d*,\d{1,}R1|S-\d*,\d{1,}R|S-\d*,\d{1,}|S-\d*RW|S-\d*R1|S-\d*R|S-\d{1,}|S-\d*', content_remarks.upper())
        wrong_joints_f = re.findall(r'F-\d*\.\d*RW|F-\d*\.\d*R1|F-\d*\.\d*R|F-\d*\.\d*|F-\d*,\d*RW|F-\d*,\d{1,}R1|F-\d*,\d{1,}R|F-\d*,\d{1,}|F-\d*RW|F-\d*R1|F-\d*R|F-\d{1,}|F-\d*', content_remarks.upper())

        find_joint_s = re.findall(r'S\d*\.\d*RW|S\d*\.\d*R1|S\d*\.\d*R|S\d*\.\d*|S\d*RW|S\d*R1|S\d*R|S\d{1,}', content_remarks.upper())
        find_joint_f = re.findall(r'F\d*\.\d*RW|F\d*\.\d*R1|F\d*\.\d*R|F\d*\.\d*|F\d*RW|F\d*R1|F\d*R|F\d{1,}', content_remarks.upper())

        joints_list = []
        for z in wrong_joints_s:
            if z[-1] == '.':
                z = z[:-1]
            joints_list.append(z.replace('-', '').replace(',', '.'))
        for z in wrong_joints_f:
            if z[-1] == '.':
                z = z[:-1]
            joints_list.append(z.replace('-', '').replace(',', '.'))
        for z in find_joint_s:
            if z[-1] == '.':
                z = z[:-1]
            joints_list.append(z)
        for z in find_joint_f:
            if z[-1] == '.':
                z = z[:-1]
            joints_list.append(z)
        if len(isometrics_1) == 0 or len(joints_list) == 0:
            print(joints_list, number_ncr, notification_items, isometrics_1)


        if isometrics_1:
            for i in isometrics_1:
                for l in joints_list:
                    open_ncr.append([i, l, number_ncr, notification_items, mark_execution, i + l])

for i in open_ncr:
    print(i)


workbook_ncr = xlsxwriter.Workbook(f'Сводка по уведомлениям на {datetime.datetime.now().strftime("%d.%m.%Y")}.xlsx')
ws5 = workbook_ncr.add_worksheet('уведомления')
ws5.set_column(0, 0, 33)
ws5.set_column(1, 4, 15)
ws5.set_column(5, 5, 37)
ws5.autofilter('A1:J5000')
for i, (testpack, joint_marks, ustan, status, exec, scep) in enumerate(open_ncr, start=1):
    ws5.write(f'A{i}', testpack)
    ws5.write(f'B{i}', joint_marks)
    ws5.write(f'C{i}', ustan)
    ws5.write(f'D{i}', status)
    ws5.write(f'E{i}', exec)
    ws5.write(f'F{i}', scep)
workbook_ncr.close()


print('OK')
