import os
import csv
import math
import openpyxl as xl
import xlsxwriter
from datetime import datetime as dt
import pandas as pd


print('start')


journal_folder = os.getcwd() + "\\out_files_for_dbs\\"
db_folder = os.getcwd() + "\\dbs\\"

journal_rt_sg = journal_folder + "Журнал P2.xlsx"
journal_ut_sg = journal_folder + "Журнал УЗК.xlsx"

wl_phase_23 = journal_folder + "WeldLog_Yamata_Phase 2-3.xlsx"
wl_phase_1 = journal_folder + "WeldLog_Yamata_Phase-1.xlsx"

summary_WL_phase_1 = [['Линия', 'Процент контроля по проекту', 'Сварено стыков',
              'Необходим контроль ПО', 'Проконтролировано ПО',
              'Необходим НК СГ, ст.', 'Проведён НК СГ ст.']]

summary_WL_phase_2 = [['Линия', 'Тестпакет', 'Установка', 'Процент контроля по проекту', 'Сварено стыков',
              'Необходим контроль ПО', 'Проконтролировано ПО',
              'Необходим НК СГ, ст.', 'Проведён НК СГ ст.']]

"""
Сбор инфы на линии
"""
def create_ll_dic():
    dic_inf_ll = {}

    with open(db_folder + "ll_all.csv", 'r') as r_file:
        readed_file = csv.reader(r_file, delimiter=";")
        for row in readed_file:
            dic_inf_ll[row[0]] = [row[1], row[2], row[3], row[4], row[5], row[6], row[7]]

    return dic_inf_ll



# sorting_journal = pd.read_excel(journal_rt_sg)
# sorting_journal = sorting_journal.sort_values(by='Дата РК', ascending=True)
# sorting_journal.to_excel('Журнал P2.xlsx')
#
# sorting_journal = pd.read_excel(journal_ut_sg, sheet_name='P2')
# sorting_journal = sorting_journal.sort_values('Дата УЗК', ascending=True)
# sorting_journal.to_excel('Журнал УЗК.xlsx')
#
# print('Журналы НК СГ отсортированы.')


def check_sg_journals():
    # Журнал РК
    sg_control_dic = {}

    wb_rt_sg = xl.load_workbook(journal_rt_sg)
    sheet_sg = wb_rt_sg['лог']

    for k in sheet_sg['H2':'X240000']:
        if k[0].value:
            iso_number_sg = str(k[0].value).strip()
            joint_number_sg = str(k[5].value).strip().replace(',', '.')
            if 'R' in str(k[6].value):
                joint_number_sg = str(k[5].value).strip().replace(',', '.') + str(k[6].value).strip()
            date_rt_sg = dt.strptime(str(k[15].value), "%Y-%m-%d %H:%M:%S")

            sg_control_dic[iso_number_sg+joint_number_sg] = [iso_number_sg, joint_number_sg, date_rt_sg]

    print('сг RT закончил')
    wb_rt_sg.close()

    #Журнал УЗК------------------------------------------

    wb_ut_sg = xl.load_workbook(journal_ut_sg)
    sheet_ut_sg = wb_ut_sg['P2']

    for i in sheet_ut_sg['D3':'P300000']:
        if i[0].value:
            iso_number_ut = str(i[0].value).strip()
            joint_number_ut = str(i[4].value).strip().replace(',', '.')
            if 'R' in str(i[5].value):
                joint_number_ut = str(i[4].value).strip().replace(',', '.') + str(i[5].value).strip()
            date_ut_sg = dt.strptime(str(i[11].value), "%Y-%m-%d %H:%M:%S")

            sg_control_dic[iso_number_ut+joint_number_ut] = [iso_number_ut, joint_number_ut, date_ut_sg]

    print('сг UT закончил')
    wb_ut_sg.close()

    return sg_control_dic




# Скан WeldLog'а -----------------
def check_wl_p1():
    wb_WL = xl.load_workbook(wl_phase_1)
    sheet = wb_WL['WELDLOG']

    check_sg_control = check_sg_journals()

    line_check_p1 = {}

    for i in sheet['A3':'CJ100000']:
        if i[0].value:
            if not i[44].value:
                line_number = str(i[3].value).strip()
                line_check_p1[line_number] = [0, 0, 0]



    for i in sheet['A3':'CJ100000']:
        if i[0].value:
            if '4.1.4.60' in str(i[1].value) or '4.1.4.70' in str(i[1].value):
                if not i[44].value:
                    line_number = str(i[3].value).strip()
                    isometric = str(i[6].value).strip()
                    number_joint = str(i[12].value).strip() + \
                                   str(i[13].value).replace(',', '.').strip()

                    if i[14].value:
                        number_joint = str(i[12].value).strip() + \
                                       str(i[13].value).replace(',', '.').strip() + str(i[14].value).strip()

                    elem1 = str(i[55].value).strip()
                    elem2 = str(i[57].value).strip()

                    print(isometric + number_joint)

                    if 'BW' in str(i[15].value).strip() and 'FLANGE SW' not in elem1 and 'FLANGE SW' not in elem2:
                        line_check_p1[line_number][0] += 1

                        if i[30].value or i[33].value:
                            line_check_p1[line_number][1] += 1

                        if isometric + number_joint in check_sg_control.keys():
                            line_check_p1[line_number][2] += 1
    print('WL_1 закончил')
    return line_check_p1


lines_p1 = check_wl_p1()

for i in lines_p1.keys():
    ll_dic = create_ll_dic()
    line_control_percent = float(ll_dic[i][4])

    need_control_po = math.ceil(float(lines_p1[i][0]) * line_control_percent)

    print(i, lines_p1[i])






#Запись в файл выходных данных
workbook_wl = xlsxwriter.Workbook(f'WL_Phase2_DC_summary {datetime.datetime.now().strftime("%d.%m.%Y")}.xlsx')
ws = workbook_wl.add_worksheet()

ws.set_column(0, 0, 30)
ws.set_column(1, 1, 33)
ws.set_column(2, 2, 10)
ws.set_column(3, 3, 12)
ws.set_column(4, 4, 27)
ws.set_column(5, 5, 10)
ws.set_column(6, 6, 27)
ws.set_column(7, 7, 10)
ws.set_column(8, 8, 20)
ws.autofilter(f'A1:I{len(wl_dic.keys())}')

cell_form = workbook_wl.add_format()
cell_form.set_text_wrap(text_wrap=1)
for i, (short_tp, number_line, number_iso, number_joint, percent_control, number_ut, res_ut,
        number_xr, res_xr) in enumerate(wl_summary_list, start=1):
    color = cell_form

    ws.write(f'A{i}', short_tp, color)
    ws.write(f'B{i}', number_line, color)
    ws.write(f'C{i}', number_iso, color)
    ws.write(f'D{i}', number_joint, color)
    ws.write(f'E{i}', percent_control, color)
    ws.write(f'F{i}', number_ut, color)
    ws.write(f'G{i}', res_ut, color)
    ws.write(f'H{i}', number_xr, color)
    ws.write(f'I{i}', res_xr, color)



workbook_wl.close()
print('done')




