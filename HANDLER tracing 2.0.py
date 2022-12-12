import csv
import os

import openpyxl as xl
import re
import xlsxwriter
from dateutil import parser
import datetime
import pandas as pd


"""
Формирование словаря из БД
"""

def read_db():
    db_tracing = {}
    short_draw_temp_dic = {}

    with open(os.getcwd() + "\\dbs\\db_tracing.csv", "r") as read_file:
        readed_file = csv.reader(read_file, delimiter=";")
        for row in readed_file:
            short_draw_temp_dic[row[1]] = []

    with open(os.getcwd() + "\\dbs\\db_tracing.csv", "r") as read_file:
        readed_file = csv.reader(read_file, delimiter=";")
        for row in readed_file:
            db_tracing[row[0]] = [row[1], row[2], row[3], row[4], row[5], row[6], '', '']
            short_draw_temp_dic[row[1]].append(row[0])

    return [db_tracing, short_draw_temp_dic]


"""
Чтение журнала заявок
"""
def check_rfi_journal(db_tracing: dict, short_db: dict):
    df = pd.read_excel('Журнал заявок общий.xlsx')
    df = df.sort_values(by='Дата назначения инспекции / Date of scheduled inspection', ascending=True)
    df.to_excel('Журнал заявок общий.xlsx', index=0)

    wb_tracing = xl.load_workbook('Журнал заявок общий.xlsx')
    sheet_tracing = wb_tracing['Sheet1']
    for i in sheet_tracing['B2':'AO250000']:
        if i[0].value:
            rfi_number = str(i[1].value)
            description_rfi = str(i[16].value).replace('\n', '').strip().replace(' ', '').\
                replace('TM5-0', 'TM5-ID-A-0').replace('.ID', '-ID').\
                replace('TM4-0', 'TM4-ID-A-0').replace('TM4.ID-', 'TM4-ID-A-').replace('0055-4', '0055-CPC-GGC-4')

            name_insp = str(i[26].value)
            list_iso = str(i[8].value)
            volume_meter = re.sub(r'[^0-9.]', '', str(i[18].value))
            category_cancelled = str(i[31].value)
            comment = str(i[39].value)  # комментарий для сортировки Физ. объём подтверждён  на прочность и плотность
            violation = str(i[35].value)


            # 0055-CPC-GGC-4.3.3.30.436-TM5-0006

            drawings = []
            list_drawing_1 = re.findall(r'0055-CPC-GGC-4\.\d\.\d\.\d\d\.\d\d\d-\w\w\d-ID-\d\d\d\d', description_rfi)
            list_drawing_2 = re.findall(r'0055-CPC-GGC-4\.\d\.\d\.\d\d\.\d\d\d-\w\w\d-ID-A-\d\d\d\d', description_rfi)
            list_drawing_3 = re.findall(r'0055-4\.\d\.\d\.\d\d\.\d\d\d-\w\w\d-ID-A-\d\d\d\d', description_rfi)

            list_short_draw = re.findall(r'\d-\d\d-HWSM-\d\d\d-\d\d/\d-\d\d-HWSM-\d\d\d-\d\d|\d-\d\d-HWSM-\d\d\d-\d\d|'
                                         r'\d-\d\d-STSM-\d\d\d-\d\d/\d-\d\d-STSM-\d\d\d-\d\d|\d-\d\d-STSM-\d\d\d-\d\d',
                                         description_rfi.replace(' ', '').strip())

            if list_drawing_1:
                for draw in list_drawing_1:
                    drawings.append(draw.replace('-ID-', '-ID-A-'))
            if list_drawing_2:
                for draw in list_drawing_2:
                    drawings.append(draw)
            if list_drawing_3:
                for draw in list_drawing_3:
                    drawings.append(draw.replace('0055-', '0055-CPC-GGC-'))

            if rfi_number == 'CPECC-CC-97235':
                description_rfi = description_rfi + 'родувкаспутник'
            else:
                pass

            if 'онтаж' in description_rfi and 'спутник' in description_rfi:
                if drawings:
                    for draw in drawings:
                        if draw in db_tracing.keys():
                            if 'Принято' == category_cancelled:
                                db_tracing[draw][3] = rfi_number
                            if 'Принято с замечаниями' == category_cancelled:
                                db_tracing[draw][3] = rfi_number + " ПЗ"


                            if 'Не принято' == category_cancelled:
                                if 'документы, подтверждающие' in violation \
                                    or 'представлены не в полном объеме' in violation:
                                    db_tracing[draw][3] = rfi_number + " ФОП"
                                else:
                                    if 'ФОП' in comment or 'потвержд' in comment:
                                        db_tracing[draw][3] = rfi_number + " ФОП"

                if list_short_draw:
                    for draw in list_short_draw:
                        if draw in short_db.keys():
                            for gost_dr in short_db[draw]:
                                if 'Принято' == category_cancelled:
                                    db_tracing[gost_dr][3] = rfi_number
                                if 'Принято с замечаниями' == category_cancelled:
                                    db_tracing[gost_dr][3] = rfi_number + " ПЗ"

                                if 'Не принято' == category_cancelled:
                                    if 'документы, подтверждающие' in violation \
                                            or 'представлены не в полном объеме' in violation:
                                        db_tracing[gost_dr][3] = rfi_number + " ФОП"
                                    else:
                                        if 'ФОП' in comment or 'потвержд' in comment:
                                            db_tracing[gost_dr][3] = rfi_number + " ФОП"

            if 'спыта' in description_rfi and 'спутник' in description_rfi:
                if drawings:
                    for draw in drawings:
                        if draw in db_tracing.keys():
                            if 'Принято' == category_cancelled:
                                db_tracing[draw][4] = rfi_number
                            if 'Принято с замечаниями' == category_cancelled:
                                db_tracing[draw][4] = rfi_number + " ПЗ"


                            if 'Не принято' == category_cancelled:
                                if 'документы, подтверждающие' in violation \
                                    or 'представлены не в полном объеме' in violation:
                                    db_tracing[draw][4] = rfi_number + " ФОП"
                                else:
                                    if 'выдерж' in comment or 'потвержд' in comment or 'ФОП' in comment:
                                        db_tracing[draw][4] = rfi_number + " ФОП"

                if list_short_draw:
                    for draw in list_short_draw:
                        if draw in short_db.keys():
                            for gost_dr in short_db[draw]:
                                if 'Принято' == category_cancelled:
                                    db_tracing[gost_dr][4] = rfi_number
                                if 'Принято с замечаниями' == category_cancelled:
                                    db_tracing[gost_dr][4] = rfi_number + " ПЗ"

                                if 'Не принято' == category_cancelled:
                                    if 'документы, подтверждающие' in violation \
                                            or 'представлены не в полном объеме' in violation:
                                        db_tracing[gost_dr][4] = rfi_number + " ФОП"
                                    else:
                                        if 'выдерж' in comment or 'потвержд' in comment or 'ФОП' in comment:
                                            db_tracing[gost_dr][4] = rfi_number + " ФОП"

            if 'родувка' in description_rfi and 'спутник' in description_rfi:
                if drawings:
                    for draw in drawings:
                        if draw in db_tracing.keys():
                            if 'Принято' == category_cancelled:
                                db_tracing[draw][5] = rfi_number
                            if 'Принято с замечаниями' == category_cancelled:
                                db_tracing[draw][5] = rfi_number + " ПЗ"


                            if 'Не принято' == category_cancelled:
                                if 'документы, подтверждающие' in violation \
                                    or 'представлены не в полном объеме' in violation:
                                    db_tracing[draw][5] = rfi_number + " ФОП"
                                else:
                                    if 'зафиксирован' in comment or 'ФОП' in comment:
                                        db_tracing[draw][5] = rfi_number + " ФОП"
                if list_short_draw:
                    for draw in list_short_draw:
                        if draw in short_db.keys():
                            for gost_dr in short_db[draw]:
                                if 'Принято' == category_cancelled:
                                    db_tracing[gost_dr][5] = rfi_number
                                if 'Принято с замечаниями' == category_cancelled:
                                    db_tracing[gost_dr][5] = rfi_number + " ПЗ"

                                if 'Не принято' == category_cancelled:
                                    if 'документы, подтверждающие' in violation \
                                            or 'представлены не в полном объеме' in violation:
                                        db_tracing[gost_dr][5] = rfi_number + " ФОП"
                                    else:
                                        if 'выдерж' in comment or 'потвержд' in comment or 'ФОП' in comment:
                                            db_tracing[gost_dr][5] = rfi_number + " ФОП"



    #         ИЗОЛЯЦИЯ СПУТНИКОВ

            if 'теплоизоляционного' in description_rfi and 'спутник' in description_rfi:
                if drawings:
                    for draw in drawings:
                        if draw in db_tracing.keys():
                            if 'Принято' == category_cancelled:
                                db_tracing[draw][6] = rfi_number
                            if 'Принято с замечаниями' == category_cancelled:
                                db_tracing[draw][6] = rfi_number + " ПЗ"

                            if 'Не принято' == category_cancelled:
                                if 'документы, подтверждающие' in violation \
                                        or 'представлены не в полном объеме' in violation:
                                    db_tracing[draw][6] = rfi_number + " ФОП"
                                else:
                                    if 'потвержд' in comment:
                                        db_tracing[draw][6] = rfi_number + " ФОП"
                if list_short_draw:
                    for draw in list_short_draw:
                        if draw in short_db.keys():
                            for gost_dr in short_db[draw]:
                                if 'Принято' == category_cancelled:
                                    db_tracing[gost_dr][6] = rfi_number
                                if 'Принято с замечаниями' == category_cancelled:
                                    db_tracing[gost_dr][6] = rfi_number + " ПЗ"

                                if 'Не принято' == category_cancelled:
                                    if 'документы, подтверждающие' in violation \
                                            or 'представлены не в полном объеме' in violation:
                                        db_tracing[gost_dr][6] = rfi_number + " ФОП"
                                    else:
                                        if 'выдерж' in comment or 'потвержд' in comment:
                                            db_tracing[gost_dr][6] = rfi_number + " ФОП"


            if 'кожух' in description_rfi and 'спутник' in description_rfi:
                if drawings:
                    for draw in drawings:
                        if draw in db_tracing.keys():
                            if 'Принято' == category_cancelled:
                                db_tracing[draw][7] = rfi_number
                            if 'Принято с замечаниями' == category_cancelled:
                                db_tracing[draw][7] = rfi_number + " ПЗ"

                            if 'Не принято' == category_cancelled:
                                if 'документы, подтверждающие' in violation \
                                        or 'представлены не в полном объеме' in violation:
                                    db_tracing[draw][7] = rfi_number + " ФОП"
                                else:
                                    if 'выдерж' in comment or 'потвержд' in comment:
                                        db_tracing[draw][7] = rfi_number + " ФОП"
                if list_short_draw:
                    for draw in list_short_draw:
                        if draw in short_db.keys():
                            for gost_dr in short_db[draw]:
                                if 'Принято' == category_cancelled:
                                    db_tracing[gost_dr][7] = rfi_number
                                if 'Принято с замечаниями' == category_cancelled:
                                    db_tracing[gost_dr][7] = rfi_number + " ПЗ"

                                if 'Не принято' == category_cancelled:
                                    if 'документы, подтверждающие' in violation \
                                            or 'представлены не в полном объеме' in violation:
                                        db_tracing[gost_dr][7] = rfi_number + " ФОП"
                                    else:
                                        if 'выдерж' in comment or 'потвержд' in comment:
                                            db_tracing[gost_dr][7] = rfi_number + " ФОП"

    return db_tracing



db_tracing = read_db()

summary_tracing = check_rfi_journal(db_tracing[0], db_tracing[1])


# for key in summary_tracing.keys():
#     print(key, summary_tracing[key][0], summary_tracing[key][1], summary_tracing[key][2], summary_tracing[key][3],
#           summary_tracing[key][4], summary_tracing[key][5])



summary_sputnik = [['Чертеж по ГОСТ', 'Чертеж', 'Установка', 'Длина', 'RFI  ERECTION', 'RFI TEST',
                    'RFI BLOWING', 'RFI ВАТА', 'RFI Металл']]


for key in summary_tracing.keys():
    summary_sputnik.append([key, summary_tracing[key][0], summary_tracing[key][1], summary_tracing[key][2],
                            summary_tracing[key][3], summary_tracing[key][4], summary_tracing[key][5],
                            summary_tracing[key][6], summary_tracing[key][7]])

"""
Запись в файл сводки

"""
workbook_summary_sputnik = xlsxwriter.Workbook(f'Сводка по теплоспутникам по ФАЗАМ 1, 2, 3 на {datetime.datetime.now().strftime("%d.%m.%Y")}.xlsx')

cell_format_green = workbook_summary_sputnik.add_format()
cell_format_green.set_bg_color('#98FB98')
cell_format_blue = workbook_summary_sputnik.add_format()
cell_format_blue.set_bg_color('#B0E0E6')
cell_format_hat = workbook_summary_sputnik.add_format()
cell_format_hat.set_bg_color('#F0E68C')
cell_format_date = workbook_summary_sputnik.add_format()
cell_format_date.set_font_size(font_size=14)


ws11 = workbook_summary_sputnik.add_worksheet('Сводка по спутникам')
ws11.set_column(0, 1, 38)
ws11.set_column(2, 3, 12)
ws11.set_column(4, 9, 22)
ws11.autofilter('A1:J2000')


for i, (one, two, three, four, five, six, seven, eight, nine) in enumerate(summary_sputnik, start=1):
    if one == 'Чертеж по ГОСТ':
        color = cell_format_hat
        color.set_bold('bold')
    elif 'CC' in seven:
        color = cell_format_green
    else:
        color = cell_format_blue
    try:
        color.set_border(style=1)
        color.set_text_wrap(text_wrap=1)
    except:
        pass
    ws11.write(f'A{i}', one, color)
    ws11.write(f'B{i}', two, color)
    ws11.write(f'C{i}', three, color)
    ws11.write(f'D{i}', four, color)
    ws11.write(f'E{i}', five, color)
    ws11.write(f'F{i}', six, color)
    ws11.write(f'G{i}', seven, color)
    ws11.write(f'H{i}', eight, color)
    ws11.write(f'I{i}', nine, color)
    # ws11.write(f'J{i}', ten, color)



workbook_summary_sputnik.close()

print('Done')

print('Файл по спутникам создан.')


